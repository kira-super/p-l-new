"""Pipeline utility helpers.

Pure utility functions used by :func:`oefof_pl.pipeline.run_pipeline`.
None of these functions have I/O side-effects beyond the specific artefact
they produce (parquet write, CSV write, archive copy, email send).

Moving them here keeps ``pipeline.py`` focused on the orchestration logic
and stage sequencing rather than the implementation of each sub-task.
"""
from __future__ import annotations

import csv as _csv
import os
import shutil
import tempfile
from pathlib import Path
from typing import Sequence

import pandas as pd

from .analyst.map import UNASSIGNED, AnalystMap
from .compute.exposure import PeriodMeta
from .config import Config
from .email_sender.outlook import (
    EmailMessage,
    EmailSendResult,
    build_email_html_body,
    build_summary_html_table,
    normalize_recipients,
    send_via_outlook,
)
from .exceptions import OefofError
from .output.valuation import (
    DATA_FIRST_ROW as _VAL_DATA_FIRST_ROW,
    HPV_FUTVAL_IDX as _VAL_HPV_FUTVAL_IDX,
    HPV_INITIALS_IDX as _VAL_HPV_INITIALS_IDX,
    HPV_SECURITY_NAME_IDX as _VAL_HPV_SECURITY_NAME_IDX,
    HPV_SORT1_IDX as _VAL_HPV_SORT1_IDX,
)
from .pipeline_types import PipelineOptions, StageResult


# ─── Date utilities ──────────────────────────────────────────────────────────

def to_yyyymmdd(ts: pd.Timestamp) -> str:
    return pd.Timestamp(ts).strftime("%Y%m%d")


def resolve_latest_sql_vdate(
    *, conn_str: str, fund_pcodes: tuple[str, ...],
) -> str:
    """Return the latest VDATE (YYYYMMDD) the SQL view holds for the fund.

    Used when the caller didn't pass ``--end-yyyymmdd``: equivalent to the
    legacy "open HP_VAL.xlsm and read the VDATE column" behaviour.
    Tiny query — single date — so it adds milliseconds, not seconds.
    """
    try:
        import pyodbc  # type: ignore
    except ImportError as e:  # pragma: no cover
        raise OefofError(
            "pyodbc is required for SQL snapshot loading."
        ) from e
    pcodes = tuple(p.strip().upper() for p in fund_pcodes if str(p).strip())
    if not pcodes:
        raise OefofError("fund_pcodes is empty; cannot resolve latest VDATE")
    placeholders = ",".join("?" for _ in pcodes)
    sql = (
        "SELECT MAX(VDATE) FROM dbo.vw_RPT_VAL "
        f"WHERE UPPER(PCODE) IN ({placeholders})"
    )
    try:
        with pyodbc.connect(conn_str, timeout=15) as cn:
            row = cn.cursor().execute(sql, pcodes).fetchone()
    except pyodbc.Error as e:
        raise OefofError(f"vw_RPT_VAL: latest-VDATE probe failed ({e})") from e
    if not row or row[0] is None:
        raise OefofError(
            f"vw_RPT_VAL: no rows for fund PCODEs {pcodes}; "
            "server has no snapshot to use as 'latest'."
        )
    return pd.Timestamp(row[0]).strftime("%Y%m%d")


# ─── Snapshot / trade helpers ────────────────────────────────────────────────

def filter_trade_window(
    trades_df: pd.DataFrame, start_vdate: pd.Timestamp, end_vdate: pd.Timestamp,
) -> pd.DataFrame:
    """Keep trades with start_vdate < CDATE ≤ end_vdate.

    Trades on the start date itself are *excluded* (they are part of the
    start-of-day position). Trades on the end date are included.
    """
    if trades_df.empty or "CDATE" not in trades_df.columns:
        return trades_df
    cdate = pd.to_datetime(trades_df["CDATE"], errors="coerce")
    start = pd.Timestamp(start_vdate).normalize()
    end = pd.Timestamp(end_vdate).normalize()
    mask = (cdate.dt.normalize() > start) & (cdate.dt.normalize() <= end)
    return trades_df.loc[mask].reset_index(drop=True)


def filter_snapshot_to_pcodes(
    port_df: pd.DataFrame, fund_pcodes: tuple[str, ...],
) -> pd.DataFrame:
    """Keep only rows whose ``PCODE_ORIG`` is in the fund's PCODE set.

    HP_VAL is firm-wide. Without this filter, ``aggregate_portfolio``
    would sum units across funds we have no trades for, breaking
    every per-ISIN reconciliation downstream.
    """
    if port_df.empty or "PCODE_ORIG" not in port_df.columns:
        return port_df
    pcodes_upper = {str(p).strip().upper() for p in fund_pcodes}
    mask = port_df["PCODE_ORIG"].astype(str).str.strip().str.upper().isin(pcodes_upper)
    return port_df.loc[mask].reset_index(drop=True)


# ─── Analyst enrichment ──────────────────────────────────────────────────────

def enrich_analysts_from_valuation(
    val_rows: Sequence[Sequence[object]],
    pl_df: pd.DataFrame,
    analyst_map: AnalystMap,
    *,
    end_port_df: pd.DataFrame | None = None,
) -> tuple[int, int]:
    """Use HiPort ValuationA's carried-down analyst initials to fill in
    pl_df rows still tagged ``UNASSIGNED`` and persist new ISIN→analyst
    rows back to the analyst map.

    Two sources, in order of priority:

      1. ``end_port_df`` — the raw ``val_data`` snapshot. Column
         ``Initials`` (col AS) is HiPort's per-row analyst code keyed by
         ISIN (col AL). This is the authoritative ISIN→analyst map for
         every position currently in the snapshot.
      2. ValuationA pivot (``val_rows``) — initials in col B carried down
         per group. Keyed by SNAME since the pivot has no ISIN. Used as a
         fallback for positions that exited (no longer in the snapshot)
         but still appear in ``pl_df``.

    Every currently-held ISIN (``Ending Units != 0``) is also persisted
    into the analyst map so the CSV reflects the live portfolio next run,
    even if the analyst is still unknown.

    Returns ``(reassigned, recorded)`` for the stage message.
    """
    if pl_df.empty:
        return (0, 0)

    # Source 1 — direct ISIN → Initials from val_data.
    isin_to_analyst: dict[str, str] = {}
    if end_port_df is not None and not end_port_df.empty \
            and "ISIN" in end_port_df.columns and "Initials" in end_port_df.columns:
        for isin_v, init_v in zip(
            end_port_df["ISIN"].tolist(),
            end_port_df["Initials"].tolist(),
        ):
            try:
                isin_k = str(isin_v).strip().upper()
                init_k = str(init_v).strip()
            except Exception:
                continue
            if not isin_k or isin_k in ("NAN", "NONE"):
                continue
            if not init_k or init_k.lower() in ("nan", "(blank)", "none"):
                continue
            isin_to_analyst.setdefault(isin_k, init_k)

    sname_to_analyst: dict[str, str] = {}
    current_sort1: str | None = None
    current_initials: str | None = None
    for idx, row in enumerate(val_rows):
        if idx + 1 < _VAL_DATA_FIRST_ROW:
            continue
        row_padded = list(row) + [None] * (
            max(0, _VAL_HPV_FUTVAL_IDX + 1 - len(row))
        )
        sort1 = row_padded[_VAL_HPV_SORT1_IDX]
        if sort1 not in (None, ""):
            current_sort1 = str(sort1).strip()
        if sort1 is not None and "Total" in str(sort1):
            current_initials = None
        initials = row_padded[_VAL_HPV_INITIALS_IDX]
        if initials not in (None, ""):
            txt = str(initials).strip()
            # HiPort uses "(blank)" as a group label for un-tagged rows.
            # Treat it like no initials so we don't persist it as an
            # analyst code or override good map entries with it.
            if txt and txt.lower() != "(blank)":
                current_initials = txt
            else:
                current_initials = None
        if current_sort1 != "A":
            continue
        sname = row_padded[_VAL_HPV_SECURITY_NAME_IDX]
        if not sname or not current_initials:
            continue
        key = str(sname).strip().upper()
        if key:
            sname_to_analyst.setdefault(key, current_initials)

    reassigned = 0
    recorded = 0
    for i in pl_df.index:
        cur = str(pl_df.at[i, "Analyst"]).strip().upper()
        isin_val = str(pl_df.at[i, "ISIN"]).strip().upper()
        sname_val = str(pl_df.at[i, "Stock Name"]).strip().upper()
        # Prefer val_data ISIN match; fall back to pivot SNAME match.
        new_a = isin_to_analyst.get(isin_val) or sname_to_analyst.get(sname_val)
        if cur == UNASSIGNED and new_a:
            pl_df.at[i, "Analyst"] = new_a
            reassigned += 1
            cur = new_a
        # Persist any currently-held ISIN into the map so the CSV mirrors
        # the live portfolio. UNASSIGNED entries surface unmapped names.
        try:
            held = float(pl_df.at[i, "Ending Units"]) != 0.0
        except (TypeError, ValueError):
            held = False
        if not held:
            continue
        isin = str(pl_df.at[i, "ISIN"]).strip().upper()
        if not isin:
            continue
        existing = analyst_map.by_isin.get(isin)
        if existing != cur:
            analyst_map.by_isin[isin] = cur
            analyst_map._dirty = True
            recorded += 1
    return (reassigned, recorded)


# ─── Valuation lookup ────────────────────────────────────────────────────────

def build_pl_lookup(pl_df: pd.DataFrame) -> dict[str, list[float]]:
    """Map ISIN → 8 P&L values (M:T) for the ValuationA workbook.

    Order matches ``oefof_pl.output.valuation.VAL_HEADERS_M_TO_T``:
    Realised EUR, Realised %, Unrealised EUR, Unrealised %,
    Income EUR, Income Yield %, Total EUR, Total %.
    """
    if pl_df.empty:
        return {}
    cols = (
        "Realised P&L (EUR)", "Realised P&L (%)",
        "Unrealised P&L (EUR)", "Unrealised P&L (%)",
        "Income (EUR)", "Income Yield (%)",
        "Total P&L (EUR)", "Total P&L (%)",
    )
    missing = [c for c in cols if c not in pl_df.columns]
    if missing:
        raise KeyError(f"build_pl_lookup: pl_df missing columns {missing}")
    out: dict[str, list[float]] = {}
    for _, row in pl_df.iterrows():
        isin = str(row.get("ISIN", "")).strip()
        if not isin:
            continue
        out[isin] = [float(row[c]) if pd.notna(row[c]) else 0.0 for c in cols]
    return out


# ─── I/O helpers ─────────────────────────────────────────────────────────────

def write_parquet_atomic(df: pd.DataFrame, path: Path) -> None:
    """Write ``df`` to parquet via a sibling tempfile + ``os.replace``."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_str = tempfile.mkstemp(
        prefix=path.stem + ".", suffix=".parquet.tmp", dir=str(path.parent),
    )
    os.close(fd)
    tmp = Path(tmp_str)
    try:
        df.to_parquet(tmp, index=False)
        os.replace(tmp, path)
    finally:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass


def archive_pl_report(
    archive_root: Path,
    src_path: Path,
    end_vdate: pd.Timestamp,
    stages: list,
) -> Path | None:
    """Copy the published P&L workbook into the per-month OneDrive archive.

    Layout: ``<archive_root>/<YYYY-MM Month>/P&L-OEFOF-<YYYYMMDD>.xlsx``.
    The end-of-period date drives both the month folder and the filename so
    re-runs overwrite the same file rather than spawning duplicates.
    Failures are recorded as a non-fatal stage and never abort the pipeline.
    """
    end = pd.Timestamp(end_vdate)
    month_dir = archive_root / end.strftime("%Y-%m %B")
    dest = month_dir / f"P&L-OEFOF-{end.strftime('%Y%m%d')}.xlsx"
    try:
        month_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_path, dest)
        stages.append(StageResult(
            "archive-pl-workbook", True, str(dest),
            {"path": str(dest), "month_dir": str(month_dir)},
        ))
        return dest
    except Exception as exc:
        stages.append(StageResult(
            "archive-pl-workbook", False, f"{type(exc).__name__}: {exc}",
        ))
        return None


def save_analyst_map_safely(analyst_map: AnalystMap, stages: list[StageResult]) -> None:
    try:
        analyst_map.save()
        stages.append(StageResult(
            "save-analyst-map", True,
            f"entries={len(analyst_map.by_isin)} dirty={analyst_map.dirty}",
            {"entries": len(analyst_map.by_isin)},
        ))
    except Exception as exc:
        stages.append(StageResult("save-analyst-map", False, str(exc)))


def write_unassigned_isins(
    pl_df: pd.DataFrame, output_dir: Path, ts_str: str,
    stages: list[StageResult],
) -> Path | None:
    """Emit ``out/unassigned_isins_<ts>.csv`` listing positions still UNASSIGNED.

    Format: ISIN, SNAME, Status (Current/Exited), Ending Units, ANALYST.
    The ANALYST column is left blank for the user to fill in; the file is
    designed to be merged back into ``isin_analyst_map.csv`` row-by-row.
    Returns the path written, or None if nothing to write.
    """
    if pl_df is None or pl_df.empty or "Analyst" not in pl_df.columns:
        return None
    try:
        sub = pl_df[pl_df["Analyst"].astype(str).str.upper() == UNASSIGNED]
        if sub.empty:
            return None
        path = output_dir / f"unassigned_isins_{ts_str}.csv"
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as fh:
            w = _csv.writer(fh)
            w.writerow(("ISIN", "SNAME", "Status", "Ending Units", "ANALYST"))
            for _, r in sub.sort_values("Stock Name").iterrows():
                ending = r.get("Ending Units", 0)
                status = "Current" if ending not in (0, 0.0) else "Exited"
                w.writerow((
                    str(r.get("ISIN", "")), str(r.get("Stock Name", "")),
                    status, ending, "",
                ))
        stages.append(StageResult(
            "write-unassigned-isins", True,
            f"file={path.name} rows={len(sub)}",
            {"path": str(path), "rows": int(len(sub))},
        ))
        return path
    except Exception as exc:
        stages.append(StageResult("write-unassigned-isins", False, str(exc)))
        return None


def send_email(
    *,
    options: PipelineOptions,
    cfg: Config,
    period_meta: PeriodMeta,
    attachments: tuple[Path, ...],
    valuation_path: Path,
    outlook_dispatcher,
) -> EmailSendResult:
    to = normalize_recipients(options.email_to or cfg.email_to_default)
    cc = normalize_recipients(options.email_cc)
    subject = options.email_subject or cfg.email_subject_default
    summary_html = build_summary_html_table(valuation_path)
    footer = (
        f"Period: {period_meta.start_date.date()} → {period_meta.end_date.date()}",
        f"Source: {period_meta.snapshot_path}",
        f"Computed at: {period_meta.run_timestamp:%Y-%m-%d %H:%M}",
    )
    html_body = build_email_html_body(options.email_body, summary_html, footer_lines=footer)
    msg = EmailMessage(
        to=to, cc=cc, subject=subject,
        body_text=options.email_body, body_html=html_body,
        attachments=attachments, from_smtp=options.email_from_smtp,
    )
    kwargs = {}
    if outlook_dispatcher is not None:
        kwargs["dispatcher"] = outlook_dispatcher
    return send_via_outlook(msg, **kwargs)
