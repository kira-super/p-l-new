"""Source-file loaders for HP_VAL.xlsm (portfolio snapshots) and
StockTrList Bottler workbook (.xlsb trade blotter).

Boundary rules (BUG-8 prevention):

* ``load_snapshot`` returns ``(df, vdate)``. ``vdate`` is read from the
  ``VDATE`` column of the source — the **only** authoritative date.
* The filename is **never parsed for a date**. The mapping is
  ``find_snapshot_path(date) → path`` (one-way). The inverse is forbidden.
* ``assert_vdate_matches`` is the single function the pipeline calls to
  cross-check that the snapshot it actually got matches the date it asked
  for. Mismatch raises ``VDateMismatchError`` and the pipeline aborts.

Boundary rules (BUG-4 prevention):

* The trades loader **drops** ``TRADEGROSS_USD`` before returning. The
  string does not appear in any DataFrame downstream consumers receive.

Boundary rules (BUG-11 prevention):

* The portfolio loader applies ``PORT_RENAMES`` (PTVALUE → PTVALUE_EUR,
  etc.) before returning. The bare names do not survive.

Boundary rules (BUG-9 prevention):

* The portfolio loader drops every column listed in ``PORT_DROP``
  (currently just ``ORD``) — it cannot reach the classifier.

The legacy "silent empty DataFrame" failure mode is replaced by
``EmptyDataError`` raised whenever a load or post-filter step produces
zero rows where ≥ 1 was expected.
"""

from __future__ import annotations

import calendar
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd

from ..exceptions import (
    DataLoadError,
    EmptyDataError,
    SchemaError,
    VDateMismatchError,
)
from .normalise import (
    PORT_DROP,
    PORT_RENAMES,
    PORT_REQUIRED,
    TRADES_DROP,
    TRADES_RENAMES,
    TRADES_REQUIRED,
    assert_schema,
    normalise_isin,
)


HP_VAL_SHEET = "val_data"
BOTTLER_SHEET = "TR"
BOTTLER_SKIPROWS = 5  # legacy: row 6 is the header


@dataclass(frozen=True)
class LoadReport:
    """Structured result of a load. ``rows`` is the count *after* filtering."""
    path: Path
    rows: int
    vdate: pd.Timestamp | None = None
    note: str = ""


# ─── Snapshot loader ────────────────────────────────────────────────────────

def load_snapshot(path: Path) -> tuple[pd.DataFrame, pd.Timestamp]:
    """Load a HP_VAL.xlsm-style portfolio snapshot.

    Reads the ``val_data`` sheet, applies ``PORT_RENAMES`` and ``PORT_DROP``,
    normalises ISIN, parses ``VDATE``, validates against ``PORT_REQUIRED``,
    and returns ``(df, vdate)``.

    Raises:
        DataLoadError: file/sheet/engine missing.
        SchemaError:   required column missing after rename.
        EmptyDataError: zero rows in the snapshot.
        VDateMismatchError: rows have differing VDATEs (single-snapshot rule).
    """
    p = Path(path)
    if not p.exists():
        raise DataLoadError(f"Snapshot file not found: {p}")

    try:
        df = pd.read_excel(p, sheet_name=HP_VAL_SHEET, engine="openpyxl")
    except ValueError as e:
        # pandas raises ValueError when sheet name is not found
        raise DataLoadError(
            f"{p}: cannot read sheet {HP_VAL_SHEET!r} ({e})"
        ) from e
    except ImportError as e:
        raise DataLoadError(
            f"{p}: missing engine dependency (openpyxl). {e}"
        ) from e

    if df.empty:
        raise EmptyDataError(f"Snapshot {p} is empty (sheet {HP_VAL_SHEET}).")

    df = _apply_port_normalisation(df, where=str(p))

    assert_schema(df, PORT_REQUIRED, where=f"load_snapshot({p.name})")

    vdate = _resolve_single_vdate(df, where=str(p))
    return df, vdate


def _apply_port_normalisation(df: pd.DataFrame, *, where: str) -> pd.DataFrame:
    df = df.rename(columns=PORT_RENAMES)
    df = df.drop(columns=[c for c in PORT_DROP if c in df.columns], errors="ignore")
    if "ISIN" in df.columns:
        df["ISIN"] = df["ISIN"].map(normalise_isin)
    if "VDATE" in df.columns:
        df["VDATE"] = pd.to_datetime(df["VDATE"], errors="coerce")
    if "CCY" in df.columns:
        df["CCY"] = df["CCY"].astype("string").str.strip().str.upper()
    return df


def _resolve_single_vdate(df: pd.DataFrame, *, where: str) -> pd.Timestamp:
    """A snapshot must carry exactly one valuation date. Differing VDATE
    rows mean the source file mixes two snapshots — refuse to proceed."""
    vd = df["VDATE"].dropna()
    if vd.empty:
        raise SchemaError(f"{where}: VDATE column has no valid dates.")
    uniq = pd.Index(vd.unique())
    if len(uniq) > 1:
        raise VDateMismatchError(
            f"{where}: snapshot contains {len(uniq)} distinct VDATE values "
            f"({sorted(str(x) for x in uniq)}). Snapshots must be a single date."
        )
    return pd.Timestamp(uniq[0])


# ─── Date / archive path helpers ────────────────────────────────────────────

def assert_vdate_matches(
    vdate: pd.Timestamp, requested_yyyymmdd: str, *, where: str = ""
) -> None:
    """BUG-8 structural fix.

    The pipeline calls this immediately after ``load_snapshot``: if the
    actual valuation date in the file disagrees with what was requested
    on the command line, abort. The legacy code parsed dates from
    filenames and silently produced reports against the wrong day.
    """
    expected = pd.Timestamp(requested_yyyymmdd)
    if pd.Timestamp(vdate).normalize() != expected.normalize():
        raise VDateMismatchError(
            f"{where or 'snapshot'}: VDATE={pd.Timestamp(vdate).date()} "
            f"does not match requested date {expected.date()}. "
            "BUG-8 guard — refusing to use snapshot at the wrong date."
        )


def find_snapshot_path(yyyymmdd: str, *, archive_root: Path) -> Path:
    """Build the canonical archive path for a given valuation date.

    Layout: ``{archive_root}/{YYYY}/{Mon}/{YYYYMMDD}.xlsm``
            (falls back to ``.xls`` if ``.xlsm`` is absent).

    The function does **not** parse dates from filenames — the mapping is
    one-way (date → path). The path is returned even if the file does not
    exist; the caller decides how to handle that (load_snapshot raises
    DataLoadError on missing files).
    """
    if len(yyyymmdd) != 8 or not yyyymmdd.isdigit():
        raise ValueError(f"yyyymmdd must be 8 digits, got {yyyymmdd!r}")
    year = yyyymmdd[:4]
    mon_idx = int(yyyymmdd[4:6])
    if not 1 <= mon_idx <= 12:
        raise ValueError(f"invalid month in {yyyymmdd!r}")
    mon_abbr = calendar.month_abbr[mon_idx]  # "Jan" .. "Dec"
    folder = Path(archive_root) / year / mon_abbr
    xlsm = folder / f"{yyyymmdd}.xlsm"
    if xlsm.exists():
        return xlsm
    xls = folder / f"{yyyymmdd}.xls"
    if xls.exists():
        return xls
    # Default to the .xlsm path; caller's load_snapshot will raise.
    return xlsm


# ─── Bottler trade-blotter loader ───────────────────────────────────────────

def load_bottler(
    path: Path,
    *,
    fund_pcodes: Iterable[str],
) -> pd.DataFrame:
    """Load the Bottler ``StockTrList`` trade blotter (``.xlsb``) and return
    the trades for the requested PCODE_ORIG values.

    Steps (in order):
      1. Read ``TR`` sheet via ``pyxlsb`` (skipping the 5 legacy banner rows).
      2. Apply ``TRADES_RENAMES``.
      3. Drop ``DELETED == 'Y'`` rows.
      4. Drop ``TRADES_DROP`` columns (esp. ``TRADEGROSS_USD`` — BUG-4).
      5. Filter to ``PCODE_ORIG in fund_pcodes``.
      6. Normalise ISIN, parse ``CDATE``, coerce numeric columns.
      7. Validate ``TRADES_REQUIRED`` schema.

    Raises:
        DataLoadError: file/sheet/engine missing.
        EmptyDataError: zero rows after PCODE filter.
        SchemaError: missing required column after normalisation.
    """
    p = Path(path)
    if not p.exists():
        raise DataLoadError(f"Bottler file not found: {p}")

    try:
        df = pd.read_excel(
            p, sheet_name=BOTTLER_SHEET, engine="pyxlsb",
            skiprows=BOTTLER_SKIPROWS,
        )
    except ValueError as e:
        raise DataLoadError(
            f"{p}: cannot read sheet {BOTTLER_SHEET!r} ({e})"
        ) from e
    except ImportError as e:
        raise DataLoadError(
            f"{p}: missing engine dependency (pyxlsb). {e}"
        ) from e

    if df.empty:
        raise EmptyDataError(f"Bottler {p} produced zero rows from {BOTTLER_SHEET}.")

    df = _apply_trades_normalisation(df)

    pcodes = tuple(fund_pcodes)
    if not pcodes:
        raise ValueError("fund_pcodes is empty — cannot filter trades.")
    if "PCODE_ORIG" not in df.columns:
        raise SchemaError(
            f"load_bottler({p.name}): PCODE_ORIG missing from trades. "
            f"Got: {list(df.columns)}"
        )
    df = df[df["PCODE_ORIG"].isin(pcodes)].reset_index(drop=True)

    if df.empty:
        raise EmptyDataError(
            f"load_bottler({p.name}): no trades matched PCODE_ORIG in {pcodes}."
        )

    assert_schema(df, TRADES_REQUIRED, where=f"load_bottler({p.name})")
    return df


def _apply_trades_normalisation(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns=TRADES_RENAMES)

    # BUG-4: TRADEGROSS_USD must not survive past the loader.
    df = df.drop(columns=[c for c in TRADES_DROP if c in df.columns], errors="ignore")

    # Drop deleted rows.
    if "DELETED" in df.columns:
        keep = df["DELETED"].astype("string").str.strip().str.upper().ne("Y")
        df = df.loc[keep.fillna(True)].reset_index(drop=True)

    # Normalise text columns.
    if "ISIN" in df.columns:
        df["ISIN"] = df["ISIN"].map(normalise_isin)
    for txt in ("CCY", "T", "PCODE_ORIG"):
        if txt in df.columns:
            df[txt] = df[txt].astype("string").str.strip().str.upper()

    # Parse CDATE.
    if "CDATE" in df.columns:
        df["CDATE"] = _parse_cdate(df["CDATE"])

    # Coerce numeric columns.
    for nc in ("UNITS", "GROSSPRICE_LOCAL", "BUY_NET_LOCAL",
               "SELL_NET_LOCAL", "INCOME_LOCAL"):
        if nc in df.columns:
            df[nc] = pd.to_numeric(df[nc], errors="coerce")

    return df


# Excel's date epoch (Windows): Jan 1 1900 with the famous leap-year bug —
# in practice anchor at 1899-12-30 so int 1 → 1899-12-31 and 36526 →
# 2000-01-01. pyxlsb returns CDATE as raw int64 serials; openpyxl on .xlsm
# files returns them already as ``datetime``. Both paths must work.
_EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def _parse_cdate(s: pd.Series) -> pd.Series:
    """Parse a Bottler/Hiport CDATE column to ``datetime64[ns]``.

    Two upstream shapes are supported:
        * already datetime-like (.xlsm via openpyxl) → passed through
          ``pd.to_datetime``.
        * Excel-serial integers (.xlsb via pyxlsb) → treated as days since
          1899-12-30 and converted via timedelta arithmetic.
    """
    if pd.api.types.is_numeric_dtype(s):
        days = pd.to_numeric(s, errors="coerce")
        return _EXCEL_EPOCH + pd.to_timedelta(days, unit="D")
    return pd.to_datetime(s, errors="coerce")
