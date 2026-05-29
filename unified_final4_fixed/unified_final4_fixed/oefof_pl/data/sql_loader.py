"""SQL Server loader for the HiPort valuation snapshot.

Replaces the legacy ``load_snapshot`` (HP_VAL.xlsm) for the pipeline's
end- and start-snapshot inputs. The ``.xlsm`` path remains in use ONLY
by :func:`output.valuation.extract_valuation_a_via_com`, which reads the
``ValuationA`` pivot tab — a different artifact whose row layout is not
exposed by any single SQL view.

Source view
-----------

* ``ccl.dbo.vw_RPT_VAL`` — historical, ~4k distinct VDATEs for OEFOF,
  date range 2013-04-02 .. today. Same column superset as
  ``vw_RPT_VAL_CURRENT_V2``. We always read this view so the start
  snapshot, end snapshot, and any ad-hoc backfill go through one path.

The view is firm-wide (all PCODEs, ~982 rows/day across all funds). The
loader filters at the SQL level to the OEFOF PCODE family for speed and
to mirror what the legacy ``_filter_snapshot_to_pcodes`` does on Excel.

Contract
--------

``load_snapshot_sql(yyyymmdd, *, conn_str, fund_pcodes)`` returns
``(df, vdate)``. The DataFrame has been put through the same normalisation
pipeline as :func:`data.loader.load_snapshot` (``PORT_RENAMES``,
``PORT_DROP``, ISIN/CCY/VDATE normalisation, ``PORT_REQUIRED`` schema check).
``vdate`` is a single ``pandas.Timestamp`` resolved from the data — it
is NEVER derived from the requested date, so a server-side surprise
still raises ``VDateMismatchError`` via :func:`assert_vdate_matches`.
"""

from __future__ import annotations

import re
from typing import Iterable

import pandas as pd

from ..exceptions import DataLoadError, EmptyDataError, SchemaError, VDateMismatchError
from .normalise import (
    PORT_DROP,
    PORT_RENAMES,
    PORT_REQUIRED,
    TRADES_REQUIRED,
    assert_schema,
    normalise_isin,
)


SQL_VIEW = "dbo.vw_RPT_VAL"
_SAFE_SQL_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_\.]*$")

# Trade blotter source: ccl.dbo.tTRANS joined to ccl.dbo.tSecurity on SCODE.
# Replaces the legacy ``StockTrList.xlsb`` (TR sheet) feed. Mapping verified
# against bottler for OEFOF samples on 2026-04-28: UNITS, NUPRICEG, NUPRICEN,
# BROKERAGE, NSETTLE_AMOUNT all match exactly. EXPENSES_LOCAL in the bottler
# is ``TOT_EXP - BROKERAGE`` (verified across 5 trades), so we recompute it
# the same way here to keep parity with WAC / per-stock cost calculations.
SQL_TRANS = "ccl.dbo.tTRANS"
SQL_SECURITY = "ccl.dbo.tSecurity"

# Numeric columns that arrive from the SQL view as ``decimal.Decimal``.
# pandas keeps Decimals as ``object`` dtype, which then trips downstream
# numeric ops (``pd.to_numeric`` works but is per-call). We cast once
# here so the downstream looks identical to what ``read_excel`` produced.
_NUMERIC_COLS: tuple[str, ...] = (
    "UNITS", "PTCOST", "PTVALUE", "UNREAL", "LST_PRICE", "PFACTOR", "XRATE",
    "CURRENT_MONTH_PL", "PERC_CURRENT_MONTH_PL",
    "PERC_TOTAL_PTVALUE", "PERC_TOTAL_EXPOSURE", "PERC_EXP_PTVALUE",
    "STOCK_PERF", "CNTR_PERF", "REL_PERF",
    "EXPOSURE", "PERC_GAV", "NUCOST",
)


def load_snapshot_sql(
    yyyymmdd: str,
    *,
    conn_str: str,
    fund_pcodes: Iterable[str],
) -> tuple[pd.DataFrame, pd.Timestamp]:
    """Load the OEFOF-family valuation snapshot for ``yyyymmdd`` from SQL.

    Args:
        yyyymmdd: Requested valuation date in ``YYYYMMDD`` form.
        conn_str: pyodbc-compatible connection string (see
            :data:`config.SQL_CONN_STR`).
        fund_pcodes: PCODEs that make up the OEFOF fund family
            (see :data:`config.OEFOF_PCODES`).

    Returns:
        Tuple ``(df, vdate)`` where ``vdate`` is the single ``VDATE``
        present in the result, and ``df`` matches ``PORT_REQUIRED``.

    Raises:
        DataLoadError: pyodbc connect/query failure or driver missing.
        EmptyDataError: zero rows returned for the requested date.
        SchemaError: required column missing after rename.
        VDateMismatchError: rows have differing VDATEs (server bug).
    """
    if len(yyyymmdd) != 8 or not yyyymmdd.isdigit():
        raise ValueError(f"yyyymmdd must be 8 digits, got {yyyymmdd!r}")

    pcodes = tuple(p.strip().upper() for p in fund_pcodes if str(p).strip())
    if not pcodes:
        raise ValueError("fund_pcodes must contain at least one PCODE")

    where = f"{SQL_VIEW} VDATE={yyyymmdd[:4]}-{yyyymmdd[4:6]}-{yyyymmdd[6:8]}"

    try:
        import pyodbc  # type: ignore
    except ImportError as e:  # pragma: no cover — dependency required in prod
        raise DataLoadError(
            "pyodbc is required for SQL snapshot loading. "
            "Install with: pip install pyodbc"
        ) from e

    iso_date = f"{yyyymmdd[:4]}-{yyyymmdd[4:6]}-{yyyymmdd[6:8]}"
    placeholders = ",".join("?" for _ in pcodes)
    sql = (
        "SELECT * FROM " + SQL_VIEW
        + " WHERE CAST(VDATE AS date) = CAST(? AS date)"
        + f" AND UPPER(PCODE) IN ({placeholders})"
    )
    params = (iso_date, *pcodes)

    try:
        with pyodbc.connect(conn_str, timeout=30) as cn:
            cur = cn.cursor()
            cur.execute(sql, *params)
            cols = [c[0] for c in cur.description]
            rows = cur.fetchall()
    except pyodbc.Error as e:
        raise DataLoadError(
            f"{where}: SQL query failed ({e})"
        ) from e

    # Build the DataFrame ourselves rather than going through ``pd.read_sql``,
    # which emits a ``UserWarning`` for raw pyodbc connections (it wants
    # SQLAlchemy). The cursor path is just as fast for ~150 rows and keeps
    # SQLAlchemy out of the dependency footprint.
    df = pd.DataFrame.from_records(
        [tuple(r) for r in rows], columns=cols,
    )

    if df.empty:
        raise EmptyDataError(
            f"{where}: no rows returned for OEFOF PCODEs "
            f"{pcodes} on {iso_date}. Check the server has loaded "
            "this date's snapshot."
        )

    df = _apply_port_normalisation(df)
    assert_schema(df, PORT_REQUIRED, where=f"load_snapshot_sql({iso_date})")

    vdate = _resolve_single_vdate(df, where=where)
    return df, vdate


def load_snapshot_history_sql(
    start_yyyymmdd: str,
    end_yyyymmdd: str,
    *,
    conn_str: str,
    fund_pcodes: Iterable[str],
) -> pd.DataFrame:
    """Load all OEFOF snapshot rows for ``start < VDATE <= end``.

    The returned frame uses the same normalised schema as
    :func:`load_snapshot_sql`, but may contain multiple ``VDATE`` values.
    """
    if len(start_yyyymmdd) != 8 or not start_yyyymmdd.isdigit():
        raise ValueError(f"start_yyyymmdd must be 8 digits, got {start_yyyymmdd!r}")
    if len(end_yyyymmdd) != 8 or not end_yyyymmdd.isdigit():
        raise ValueError(f"end_yyyymmdd must be 8 digits, got {end_yyyymmdd!r}")

    pcodes = tuple(p.strip().upper() for p in fund_pcodes if str(p).strip())
    if not pcodes:
        raise ValueError("fund_pcodes must contain at least one PCODE")

    try:
        import pyodbc  # type: ignore
    except ImportError as e:  # pragma: no cover
        raise DataLoadError(
            "pyodbc is required for SQL snapshot loading. "
            "Install with: pip install pyodbc"
        ) from e

    start_iso = f"{start_yyyymmdd[:4]}-{start_yyyymmdd[4:6]}-{start_yyyymmdd[6:8]}"
    end_iso = f"{end_yyyymmdd[:4]}-{end_yyyymmdd[4:6]}-{end_yyyymmdd[6:8]}"
    placeholders = ",".join("?" for _ in pcodes)
    sql = (
        "SELECT * FROM " + SQL_VIEW
        + " WHERE CAST(VDATE AS date) > CAST(? AS date)"
        + " AND CAST(VDATE AS date) <= CAST(? AS date)"
        + f" AND UPPER(PCODE) IN ({placeholders})"
    )
    params = (start_iso, end_iso, *pcodes)
    where = f"{SQL_VIEW} {start_iso}<{end_iso}"

    try:
        with pyodbc.connect(conn_str, timeout=30) as cn:
            cur = cn.cursor()
            cur.execute(sql, *params)
            cols = [c[0] for c in cur.description]
            rows = cur.fetchall()
    except pyodbc.Error as e:
        raise DataLoadError(f"{where}: SQL query failed ({e})") from e

    df = pd.DataFrame.from_records([tuple(r) for r in rows], columns=cols)
    if df.empty:
        raise EmptyDataError(
            f"{where}: no rows returned for OEFOF PCODEs {pcodes}."
        )

    df = _apply_port_normalisation(df)
    assert_schema(df, PORT_REQUIRED, where=f"load_snapshot_history_sql({start_iso},{end_iso})")
    return df.reset_index(drop=True)


def _apply_port_normalisation(df: pd.DataFrame) -> pd.DataFrame:
    """Mirror of ``loader._apply_port_normalisation`` for the SQL frame."""
    # Coerce Decimal -> float so downstream numeric ops behave identically
    # to the read_excel path (which delivers float64 directly).
    for col in _NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.rename(columns=PORT_RENAMES)
    df = df.drop(columns=[c for c in PORT_DROP if c in df.columns],
                 errors="ignore")
    if "ISIN" in df.columns:
        df["ISIN"] = df["ISIN"].map(normalise_isin)
    if "VDATE" in df.columns:
        df["VDATE"] = pd.to_datetime(df["VDATE"], errors="coerce")
    if "CCY" in df.columns:
        df["CCY"] = df["CCY"].astype("string").str.strip().str.upper()
    return df


def _assert_safe_sql_identifier(value: str, label: str) -> None:
    if not _SAFE_SQL_IDENTIFIER.match(str(value)):
        raise ValueError(f"{label} contains unsafe SQL identifier characters: {value!r}")


def _resolve_single_vdate(df: pd.DataFrame, *, where: str) -> pd.Timestamp:
    vd = df["VDATE"].dropna()
    if vd.empty:
        raise SchemaError(f"{where}: VDATE column has no valid dates.")
    uniq = pd.Index(vd.unique())
    if len(uniq) > 1:
        raise VDateMismatchError(
            f"{where}: result contains {len(uniq)} distinct VDATE values "
            f"({sorted(str(x) for x in uniq)}). Snapshots must be a single date."
        )
    return pd.Timestamp(uniq[0])


# ─── Trade blotter (replaces load_bottler) ────────────────────────────────────

# Numeric columns returned from tTRANS as ``decimal.Decimal``; cast to float.
_TRADES_NUMERIC_COLS: tuple[str, ...] = (
    "UNITS", "GROSSPRICE_LOCAL", "NETPRICE_LOCAL",
    "BUY_NET_LOCAL", "SELL_NET_LOCAL", "INCOME_LOCAL",
    "BCOMM_LOCAL", "EXPENSES_LOCAL", "STAMP_LOCAL",
)


def load_trades_sql(
    *,
    conn_str: str,
    fund_pcodes: Iterable[str],
) -> pd.DataFrame:
    """Load all trades for the OEFOF PCODE family from SQL Server.

    Replacement for ``loader.load_bottler``. Pulls every row in
    ``ccl.dbo.tTRANS`` for the requested PCODEs, joining ``ccl.dbo.tSecurity``
    on ``SCODE`` for ISIN / SHORT_NAME / CCY. ``BUY_NET_LOCAL`` and
    ``SELL_NET_LOCAL`` are derived from ``NSETTLE_AMOUNT`` based on the
    trade type ``T`` ('P' = purchase / 'S' = sale), matching the
    legacy bottler columns 1:1. ``EXPENSES_LOCAL`` is recomputed as
    ``TOT_EXP - BROKERAGE - STAMP_DUTY`` so the three cost components
    sum cleanly without overlap. Stamp duty (HK/UK markets) is exposed
    separately as ``STAMP_LOCAL`` for line-by-line transparency.

    Returns:
        DataFrame with at least :data:`TRADES_REQUIRED` columns plus
        ``BCODE`` (needed by VAL-03 and ``compute.pl`` to detect ZZZZ /
        CA_ADJ rows). Date filtering is applied later by
        ``pipeline._filter_trade_window``.

    Raises:
        DataLoadError: pyodbc connect/query failure or driver missing.
        EmptyDataError: zero rows returned for the requested PCODEs.
        SchemaError: required column missing after rename.
    """
    pcodes = tuple(p.strip().upper() for p in fund_pcodes if str(p).strip())
    if not pcodes:
        raise ValueError("fund_pcodes must contain at least one PCODE")

    where = f"{SQL_TRANS} PCODE in {pcodes}"

    try:
        import pyodbc  # type: ignore
    except ImportError as e:  # pragma: no cover — dependency required in prod
        raise DataLoadError(
            "pyodbc is required for SQL trade loading. "
            "Install with: pip install pyodbc"
        ) from e

    placeholders = ",".join("?" for _ in pcodes)
    sql = (
        "SELECT t.ID            AS TRADE_ID,"
        "       UPPER(t.PCODE)  AS PCODE_ORIG,"
        "       s.ISIN          AS ISIN,"
        "       s.SHORT_NAME    AS SNAME,"
        "       s.CCY           AS CCY,"
        "       t.T             AS T,"
        "       t.CDATE         AS CDATE,"
        "       t.UNITS         AS UNITS,"
        "       t.NUPRICEG      AS GROSSPRICE_LOCAL,"
        "       t.NUPRICEN      AS NETPRICE_LOCAL,"
        "       CASE WHEN t.T='P' THEN t.NSETTLE_AMOUNT ELSE 0 END AS BUY_NET_LOCAL,"
        "       CASE WHEN t.T='S' THEN t.NSETTLE_AMOUNT ELSE 0 END AS SELL_NET_LOCAL,"
        "       t.NINCOME       AS INCOME_LOCAL,"
        "       t.BROKERAGE     AS BCOMM_LOCAL,"
        "       (ISNULL(t.TOT_EXP,0) - ISNULL(t.BROKERAGE,0) - ISNULL(t.STAMP_DUTY,0)) AS EXPENSES_LOCAL,"
        "       ISNULL(t.STAMP_DUTY,0) AS STAMP_LOCAL,"
        "       t.BCODE         AS BCODE,"
        "       t.SCODE         AS SCODE"
        f" FROM {SQL_TRANS} t"
        f" LEFT JOIN {SQL_SECURITY} s ON s.SCODE = t.SCODE"
        f" WHERE UPPER(t.PCODE) IN ({placeholders})"
        # Exclude soft-deleted trades. The legacy bottler dropped rows where
        # ``DELETED == 'Y'``; the equivalent flag in tTRANS is ``D``. Without
        # this filter ~4k cancelled trades inflate WAC, P&L and exit counts.
        " AND (t.D IS NULL OR t.D <> 'Y')"
    )

    try:
        with pyodbc.connect(conn_str, timeout=30) as cn:
            cur = cn.cursor()
            cur.execute(sql, *pcodes)
            cols = [c[0] for c in cur.description]
            rows = cur.fetchall()
    except pyodbc.Error as e:
        raise DataLoadError(f"{where}: SQL query failed ({e})") from e

    df = pd.DataFrame.from_records(
        [tuple(r) for r in rows], columns=cols,
    )

    if df.empty:
        raise EmptyDataError(
            f"{where}: no trades returned for OEFOF PCODEs {pcodes}."
        )

    df = _apply_trades_sql_normalisation(df)
    assert_schema(df, TRADES_REQUIRED, where=f"load_trades_sql({pcodes})")
    return df


def _apply_trades_sql_normalisation(df: pd.DataFrame) -> pd.DataFrame:
    """Cast Decimals to float, normalise text columns, parse CDATE."""
    for col in _TRADES_NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "ISIN" in df.columns:
        df["ISIN"] = df["ISIN"].map(normalise_isin)
    for txt in ("CCY", "T", "PCODE_ORIG", "BCODE", "SNAME"):
        if txt in df.columns:
            df[txt] = df[txt].astype("string").str.strip()
    for txt in ("CCY", "T", "PCODE_ORIG", "BCODE"):
        if txt in df.columns:
            df[txt] = df[txt].str.upper()

    if "CDATE" in df.columns:
        df["CDATE"] = pd.to_datetime(df["CDATE"], errors="coerce")

    return df


# ─── tH_VAL: per-position end-of-day historical valuation ───────────────────
#
# ``ccl.dbo.tH_VAL`` stores one row per (PCODE, SCODE, VDATE). Beyond the
# columns mirrored in ``vw_RPT_VAL`` it carries:
#
#   * ``NUCOST_FIFO``   — FIFO per-unit cost basis (HiPort-maintained)
#   * ``ACQ_DATE``      — first acquisition date of the position
#   * ``CNTR_PERF``     — contribution to fund performance (period-to-date)
#   * ``CURRENT_MONTH_PL`` — month-to-date P&L
#   * ``SECTOR1..6``, ``INDUSTRY``, ``GICS_SUBIND`` — sector taxonomy
#
# We expose this as a per-VDATE snapshot so the validator can cross-check
# our WAC against ``NUCOST_FIFO`` (VAL-12), the workbook can render a
# sector breakdown, and a future revision can replace ``ca_overrides.csv``
# with deltas derived from successive ``UNITS`` snapshots.

SQL_H_VAL = "ccl.dbo.tH_VAL"

_TH_VAL_NUMERIC_COLS: tuple[str, ...] = (
    "UNITS", "PTCOST", "PTVALUE", "NUCOST", "NUCOST_FIFO",
    "UNREAL", "CURRENT_MONTH_PL", "CNTR_PERF",
    "PERC_TOTAL_PTVALUE", "PERC_TOTAL_EXPOSURE", "EXPOSURE",
    "LST_PRICE", "PFACTOR", "XRATE",
)

# Required columns the tH_VAL frame must expose (after normalisation).
TH_VAL_REQUIRED: tuple[str, ...] = (
    "PCODE_ORIG", "ISIN", "SNAME", "CCY", "VDATE",
    "UNITS", "NUCOST_FIFO", "PTCOST", "PTVALUE",
)


def load_th_val_sql(
    yyyymmdd: str,
    *,
    conn_str: str,
    fund_pcodes: Iterable[str],
) -> tuple[pd.DataFrame, pd.Timestamp]:
    """Load the tH_VAL position snapshot for ``yyyymmdd`` from SQL.

    Mirrors :func:`load_snapshot_sql` but reads ``tH_VAL`` directly so
    we get FIFO cost, sector taxonomy and contribution metrics that
    ``vw_RPT_VAL`` does not expose.

    Returns:
        Tuple ``(df, vdate)``. ``df`` is normalised (Decimals → float,
        ISIN uppercased, CCY trimmed, VDATE parsed). ``vdate`` is the
        single ``VDATE`` present (raises ``VDateMismatchError`` if not).
    """
    if len(yyyymmdd) != 8 or not yyyymmdd.isdigit():
        raise ValueError(f"yyyymmdd must be 8 digits, got {yyyymmdd!r}")

    pcodes = tuple(p.strip().upper() for p in fund_pcodes if str(p).strip())
    if not pcodes:
        raise ValueError("fund_pcodes must contain at least one PCODE")

    iso_date = f"{yyyymmdd[:4]}-{yyyymmdd[4:6]}-{yyyymmdd[6:8]}"
    where = f"{SQL_H_VAL} VDATE={iso_date}"

    try:
        import pyodbc  # type: ignore
    except ImportError as e:  # pragma: no cover
        raise DataLoadError(
            "pyodbc is required for tH_VAL loading"
        ) from e

    placeholders = ",".join("?" for _ in pcodes)
    sql = (
        "SELECT h.PCODE_ORIG, s.ISIN, h.SNAME, h.CCY, h.VDATE,"
        "       h.UNITS, h.PTCOST, h.PTVALUE, h.NUCOST, h.NUCOST_FIFO,"
        "       h.UNREAL, h.CURRENT_MONTH_PL, h.CNTR_PERF, h.ACQ_DATE,"
        "       h.SECTOR0, h.SECTOR1, h.SECTOR2, h.SECTOR3,"
        "       h.INDUSTRY, h.GICS_SUBIND,"
        "       h.PERC_TOTAL_PTVALUE, h.PERC_TOTAL_EXPOSURE, h.EXPOSURE,"
        "       h.LST_PRICE, h.PFACTOR, h.XRATE, h.CAT, h.SUBCAT,"
        "       s.ISIN_VALID"
        f" FROM {SQL_H_VAL} h"
        f" LEFT JOIN {SQL_SECURITY} s ON s.SCODE = h.SCODE"
        " WHERE CAST(h.VDATE AS date) = CAST(? AS date)"
        f"   AND UPPER(h.PCODE) IN ({placeholders})"
    )
    params = (iso_date, *pcodes)

    try:
        with pyodbc.connect(conn_str, timeout=30) as cn:
            cur = cn.cursor()
            cur.execute(sql, *params)
            cols = [c[0] for c in cur.description]
            rows = cur.fetchall()
    except pyodbc.Error as e:
        raise DataLoadError(f"{where}: SQL query failed ({e})") from e

    df = pd.DataFrame.from_records([tuple(r) for r in rows], columns=cols)

    if df.empty:
        raise EmptyDataError(
            f"{where}: no tH_VAL rows for {pcodes} on {iso_date}."
        )

    for col in _TH_VAL_NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "ISIN" in df.columns:
        df["ISIN"] = df["ISIN"].map(normalise_isin)
    if "VDATE" in df.columns:
        df["VDATE"] = pd.to_datetime(df["VDATE"], errors="coerce")
    if "ACQ_DATE" in df.columns:
        df["ACQ_DATE"] = pd.to_datetime(df["ACQ_DATE"], errors="coerce")
    for txt in ("CCY", "PCODE_ORIG", "SNAME", "CAT", "SUBCAT",
                "SECTOR0", "SECTOR1", "SECTOR2", "SECTOR3", "INDUSTRY",
                "ISIN_VALID"):
        if txt in df.columns:
            df[txt] = df[txt].astype("string").str.strip()
    for txt in ("CCY", "PCODE_ORIG", "ISIN_VALID"):
        if txt in df.columns:
            df[txt] = df[txt].str.upper()

    # Drop rows without ISIN (rare PCODE-level rows like cash classes).
    df = df.loc[df["ISIN"].fillna("") != ""].reset_index(drop=True)

    assert_schema(df, TH_VAL_REQUIRED, where=f"load_th_val_sql({iso_date})")
    vdate = _resolve_single_vdate(df, where=where)
    return df, vdate


# ─── FX fallback: most-recent rate for currencies absent from all snapshots ──

def fetch_missing_fx_rates(
    missing_ccys: list[str],
    *,
    conn_str: str,
    as_of_yyyymmdd: str,
) -> dict[str, float]:
    """Return ``{CCY: XRATE}`` for currencies absent from both snapshots.

    Queries ``vw_RPT_VAL`` for the most recent row with a valid (non-NULL,
    positive) ``XRATE`` for each requested currency, looking back from
    ``as_of_yyyymmdd`` across ALL PCODEs (not just OEFOF).  This handles
    positions that were opened and closed entirely within the YTD period so
    they never appear in the Dec-31 start snapshot or the current end
    snapshot (e.g. Capital Tankers MHY1096C1093 and SED Energy CY0101162119,
    both NOK, bought and sold in 2026).

    Args:
        missing_ccys:   CCY codes (upper-case) not yet in fx_start or fx_end.
        conn_str:       pyodbc connection string.
        as_of_yyyymmdd: Look back from this date (YYYYMMDD) — typically the
                        end-of-period valuation date.

    Returns:
        ``{CCY: XRATE}`` for whichever currencies were found.  CCYs with no
        valid rate anywhere are omitted — the downstream hard error fires for
        those, which is the correct behaviour.
    """
    if not missing_ccys:
        return {}

    try:
        import pyodbc  # type: ignore
    except ImportError:
        return {}

    iso_date = f"{as_of_yyyymmdd[:4]}-{as_of_yyyymmdd[4:6]}-{as_of_yyyymmdd[6:8]}"
    out: dict[str, float] = {}

    try:
        with pyodbc.connect(conn_str, timeout=15) as cn:
            cur = cn.cursor()
            for ccy in missing_ccys:
                c = str(ccy).strip().upper()
                if not c or c in ("EUR", "USD"):
                    continue
                cur.execute(
                    "SELECT TOP 1 XRATE "
                    "FROM ccl.dbo.vw_RPT_VAL "
                    "WHERE CCY = ? "
                    "AND CAST(VDATE AS date) <= CAST(? AS date) "
                    "AND XRATE IS NOT NULL AND XRATE > 0 "
                    "ORDER BY VDATE DESC",
                    c, iso_date,
                )
                row = cur.fetchone()
                if row is not None:
                    try:
                        x = float(row[0])
                        if x > 0:
                            out[c] = x
                    except (TypeError, ValueError):
                        pass
    except Exception:
        pass  # non-fatal — hard FXInvalidError fires downstream if CCY still missing

    return out


def fetch_eur_usd_rate(
    *,
    conn_str: str,
    as_of_yyyymmdd: str,
) -> float | None:
    """Return most recent EUR/USD XRATE at or before ``as_of_yyyymmdd``.

    Used as a fallback when a snapshot has no EUR row so ``resolve_eur_usd``
    cannot read the rate directly.  Returns ``None`` if no usable row found.
    """
    try:
        import pyodbc  # type: ignore
    except ImportError:
        return None

    iso_date = f"{as_of_yyyymmdd[:4]}-{as_of_yyyymmdd[4:6]}-{as_of_yyyymmdd[6:8]}"
    try:
        with pyodbc.connect(conn_str, timeout=15) as cn:
            cur = cn.cursor()
            cur.execute(
                "SELECT TOP 1 XRATE "
                "FROM ccl.dbo.vw_RPT_VAL "
                "WHERE CCY = 'EUR' "
                "AND CAST(VDATE AS date) <= CAST(? AS date) "
                "AND XRATE IS NOT NULL AND XRATE > 0 "
                "ORDER BY VDATE DESC",
                iso_date,
            )
            row = cur.fetchone()
    except Exception:
        return None

    if row is None:
        return None
    try:
        rate = float(row[0])
    except (TypeError, ValueError):
        return None
    return rate if rate > 0 else None
