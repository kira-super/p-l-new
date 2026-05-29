"""Single-file P&L workbook writer.

Pure ``openpyxl`` — no COM, no compute. Consumes:

* ``pl_df``        — output of ``compute.pl.compute_positions``
* ``income_df``    — output of ``compute.pl.compute_positions`` (long form)
* ``trades_df``    — normalised trade blotter
* ``validation``   — ``ValidationResult`` from ``validate.checks.run_all_checks``
* ``period_meta``  — ``PeriodMeta`` (dates, file paths, run timestamp)
* ``valuation_rows`` — optional ValuationA rows to embed (else sheet skipped)

Sheets, in tab order:

    1. Summary             — dashboard (KPI cards, instrument & analyst tables)
    2. ValuationA          — optional, ValuationA reconciliation
    3. <Analyst>           — one sheet per distinct analyst (dynamic)
    4. Current Holdings    — full-detail table, all analysts
    5. Exited Positions    — full-detail table, all analysts
    6. Income & Dividends  — by-position summary
    7. Income Detail       — every income trade
    8. Trade Detail        — every trade
    9. Validation Report   — built last so it can include write-time
   10. Methodology         — narrative + audit metadata

The writer never re-derives a P&L number. Every cell is either a value
straight from ``pl_df`` / ``income_df`` or an arithmetic sum of such
values.

Atomic-write contract: data goes to ``{out}.tmp`` then ``os.replace``s
to ``{out}``. A half-written workbook can never appear on disk.
"""

from __future__ import annotations

import os
import tempfile
import time
from pathlib import Path
from typing import Sequence

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from ..compute.exposure import (  # noqa: F401  (re-exported for backwards compat)
    ExposureBlock,
    ExposureMetrics,
    PeriodMeta,
)
from ..validate.checks import ValidationResult
from .styles import (
    COLOR_TAB_DATA,
    COLUMN_LAYOUT_ANALYST,
    COLUMN_LAYOUT_HOLDINGS,
    COLUMN_LAYOUT_INCOME,
)
from .sheets._helpers import (
    _analyst_order,
    _current_holdings,
    _exited_positions,
    _reduced_positions,
    _write_table,
)
from .sheets.summary import _write_summary
from .sheets.analyst import _write_analyst_sheet
from .sheets.income import _write_income_summary, _write_trade_detail
from .sheets.valuation_a import _write_valuation_a
from .sheets.validation import _write_validation, _write_methodology


# Static (non-dynamic) sheet order. Per-analyst sheets are inserted between
# "Summary" / "ValuationA" and "Current Holdings" — the count and names
# depend on the data so they cannot live in this constant.
SHEET_ORDER: tuple[str, ...] = (
    "Summary",
    "Current Holdings",
    "Exited Positions",
    "Reduced Positions",
    "Income & Dividends",
    "Income Detail",
    "Trade Detail",
    "Validation Report",
    "Methodology",
)


def _apply_live_price_overrides(
    pl_df: pd.DataFrame,
    live_prices: dict[str, float] | None,
) -> pd.DataFrame:
    """Overwrite Last Price (EUR) from live price map when available."""
    if pl_df.empty or not live_prices:
        return pl_df
    if "ISIN" not in pl_df.columns or "Last Price (EUR)" not in pl_df.columns:
        return pl_df

    out = pl_df.copy()
    price_map = {
        str(isin).strip().upper(): float(price)
        for isin, price in live_prices.items()
        if str(isin).strip()
    }
    isin_series = out["ISIN"].astype(str).str.strip().str.upper()
    mapped = isin_series.map(price_map)
    has_live = mapped.notna()
    out.loc[has_live, "Last Price (EUR)"] = mapped.loc[has_live].astype(float)
    return out


# ─── Public entry point ─────────────────────────────────────────────────────

def build_workbook(
    *,
    out_path: Path,
    pl_df: pd.DataFrame,
    income_df: pd.DataFrame,
    trades_df: pd.DataFrame,
    validation: ValidationResult,
    period_meta: PeriodMeta,
    valuation_rows: Sequence[Sequence[object]] | None = None,
    ca_suggestions_path: Path | None = None,
    applied_ca_overrides: tuple = (),       # reserved — shown in Methodology TODO
    ca_override_sources: tuple = (),        # reserved
    th_val_df: pd.DataFrame | None = None,  # reserved — used by ValuationA TODO
    analyst_sname_overrides: dict | None = None,  # reserved
    arb_pairs: tuple[dict[str, object], ...] | None = None,
    live_prices: dict[str, float] | None = None,
) -> Path:
    """Build the workbook and atomically rename it into place."""
    out_path = Path(out_path)
    t0 = time.perf_counter()

    pl_df = _apply_live_price_overrides(pl_df, live_prices)

    wb = Workbook()
    wb.remove(wb.active)

    _write_summary(wb, pl_df, income_df, period_meta)
    if valuation_rows:
        _write_valuation_a(wb, valuation_rows, period_meta, pl_df,
                           th_val_df=th_val_df)
    for analyst_code in _analyst_order(pl_df, period_meta.analyst_codes):
        _write_analyst_sheet(
            wb,
            pl_df,
            analyst_code,
            period_meta,
            arb_pairs=arb_pairs,
            live_prices=live_prices,
        )
    _write_table(wb, "Current Holdings", _current_holdings(pl_df),
                 COLUMN_LAYOUT_HOLDINGS, period_meta,
                 tab_color=COLOR_TAB_DATA)
    _write_table(wb, "Exited Positions", _exited_positions(pl_df),
                 COLUMN_LAYOUT_ANALYST, period_meta,
                 tab_color=COLOR_TAB_DATA)
    _write_table(wb, "Reduced Positions", _reduced_positions(pl_df),
                 COLUMN_LAYOUT_ANALYST, period_meta,
                 tab_color=COLOR_TAB_DATA)
    _write_income_summary(wb, pl_df, period_meta)
    _write_table(wb, "Income Detail", income_df,
                 COLUMN_LAYOUT_INCOME, period_meta,
                 tab_color=COLOR_TAB_DATA)
    _write_trade_detail(wb, trades_df, period_meta)
    _write_methodology(wb, period_meta,
                       applied_ca_overrides=applied_ca_overrides,
                       ca_override_sources=ca_override_sources,
                       analyst_sname_overrides=analyst_sname_overrides or {})

    elapsed = time.perf_counter() - t0
    _write_validation(wb, validation, period_meta, elapsed,
                      ca_suggestions_path=ca_suggestions_path)

    _enforce_sheet_order(wb, period_meta.analyst_codes)
    _atomic_save(wb, out_path)
    return out_path


# ─── Sheet ordering / atomic save ───────────────────────────────────────────

def _enforce_sheet_order(wb: Workbook, analyst_codes: dict[str, str]) -> None:
    """Order: Summary, ValuationA?, <analysts...>,
    Current Holdings, Exited Positions, Income & Dividends, Income Detail,
    Trade Detail, Validation Report, Methodology.
    """
    desired: list[str] = []
    desired.append("Summary")
    if "ValuationA" in wb.sheetnames:
        desired.append("ValuationA")
    known_static = set(SHEET_ORDER) | {"ValuationA"}
    analyst_sheets = [s for s in wb.sheetnames if s not in known_static]
    code_order = list(analyst_codes.keys())
    analyst_sheets.sort(key=lambda s: (
        code_order.index(s) if s in code_order else len(code_order),
        s,
    ))
    desired.extend(analyst_sheets)
    desired.extend([s for s in SHEET_ORDER if s != "Summary"])

    desired = [s for s in desired if s in wb.sheetnames]
    by_name = {ws.title: ws for ws in wb.worksheets}
    wb._sheets = [by_name[name] for name in desired]


def _autofit_all_sheets(wb: Workbook) -> None:
    """Widen any column whose existing width is too narrow to display its
    widest *numeric* cell, so €-amount totals don't render as ######.
    """
    WIDTH_CAP = 22
    for ws in wb.worksheets:
        col_max: dict[int, int] = {}
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.row <= 2:
                    continue
                v = cell.value
                if v is None or not isinstance(v, (int, float)):
                    continue
                fmt = cell.number_format or ""
                if "%" in fmt:
                    s = f"{v * 100:,.2f}%"
                elif "€" in fmt or "EUR" in fmt:
                    s = f"€{abs(v):,.0f}"
                    if v < 0:
                        s = f"({s})"
                else:
                    s = f"{v:,.2f}" if isinstance(v, float) else f"{v:,}"
                length = len(s)
                col = cell.column
                if col_max.get(col, 0) < length:
                    col_max[col] = length
        for col, max_len in col_max.items():
            letter = get_column_letter(col)
            current = ws.column_dimensions[letter].width or 8
            needed = min(WIDTH_CAP, max_len + 2)
            if needed > current:
                ws.column_dimensions[letter].width = needed


def _atomic_save(wb: Workbook, out_path: Path) -> None:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    _autofit_all_sheets(wb)
    fd, tmp_name = tempfile.mkstemp(
        prefix=out_path.stem + ".", suffix=".xlsx.tmp",
        dir=str(out_path.parent),
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        wb.save(str(tmp_path))
        os.replace(tmp_path, out_path)
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
