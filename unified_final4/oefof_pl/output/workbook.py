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
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from ..validate.checks import ValidationResult
from .styles import (
    ALIGN_CENTER,
    ALIGN_KPI_LABEL,
    ALIGN_KPI_VALUE,
    ALIGN_LEFT,
    ALIGN_RIGHT,
    COLOR_BAND_ALT,
    COLOR_HEADER,
    COLOR_HEADER_ACCENT,
    COLOR_KPI_FILL,
    COLOR_KPI_FILL_ACCENT,
    COLOR_MUTED,
    COLOR_NEG,
    COLOR_POS,
    COLOR_TAB_ANALYST,
    COLOR_TAB_AUDIT,
    COLOR_TAB_DASHBOARD,
    COLOR_TAB_DATA,
    COLOR_VALIDATION_FAIL,
    COLOR_VALIDATION_PASS,
    COLOR_VALIDATION_WARN,
    COLUMN_LAYOUT_ANALYST,
    COLUMN_LAYOUT_HOLDINGS,
    COLUMN_LAYOUT_INCOME,
    ColSpec,
    FILL_BAND,
    FILL_HEADER,
    FILL_HEADER_ACCENT,
    FILL_KPI,
    FILL_KPI_ACCENT,
    FILL_SUBTOTAL,
    FILL_TOTAL,
    FONT_BODY,
    FONT_BODY_MUTED,
    FONT_HEADER,
    FONT_NEG,
    FONT_NEG_BOLD,
    FONT_POS,
    FONT_KPI_LABEL,
    FONT_KPI_VALUE,
    FONT_KPI_VALUE_NEG,
    FONT_KPI_VALUE_POS,
    FONT_SECTION,
    FONT_SUBTITLE,
    FONT_TITLE,
    FONT_TOTAL,
    align_for,
    fmt_for,
)


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

# Display names for known analyst codes. Codes not in this map are shown as
# the raw code; "UNASSIGNED" keeps that label.
ANALYST_DISPLAY: dict[str, str] = {
    "IS": "Ian Simmons",
    "HK": "Hayden Kwan",
    "JB": "Julius Bottcher",
    "SB": "Stefan Bottcher",
    "KX": "Karen Xiao",
    "VS": "Vijay Singh",
    "AS": "Alexander Short",
    "VJ": "VJ",
    "UNASSIGNED": "Unassigned",
}

# Stable sheet order for analysts (matches the order the team uses in their
# Excel mock). Anything not in this list goes alphabetically at the end.
ANALYST_TAB_ORDER: tuple[str, ...] = (
    "IS", "HK", "JB", "SB", "KX", "VS", "AS", "VJ", "UNASSIGNED",
)


@dataclass
class ExposureMetrics:
    """Exposure figures for one entity (fund or analyst).

    ``short_eur`` is always a positive absolute magnitude (not signed).
    net  = long − short
    gross = long + short
    """
    long_eur: float = 0.0
    short_eur: float = 0.0
    nav_eur: float = 0.0

    @property
    def net_eur(self) -> float:
        return self.long_eur - self.short_eur

    @property
    def gross_eur(self) -> float:
        return self.long_eur + self.short_eur

    def pct_of_nav(self, value: float) -> float | None:
        return value / self.nav_eur if self.nav_eur > 1e-9 else None


@dataclass
class ExposureBlock:
    """Snapshot or AUM-weighted exposure for the fund and per analyst."""
    label: str = ""
    as_of: pd.Timestamp | None = None
    fund: ExposureMetrics = field(default_factory=ExposureMetrics)
    by_analyst: dict[str, ExposureMetrics] = field(default_factory=dict)


@dataclass(frozen=True)
class PeriodMeta:
    """All non-numeric context the workbook needs."""
    fund_name: str
    start_date: pd.Timestamp
    end_date: pd.Timestamp
    fund_currency: str = "EUR"
    snapshot_path: str = ""
    bottler_path: str = ""
    run_timestamp: datetime = field(default_factory=datetime.now)
    # Optional NAV reconciliation components (EUR). Ordered tuple of
    # (label, value, is_total) rows rendered on the Summary sheet. Bridges
    # the equity-only stock MV to the HP_VAL fund portfolio total.
    nav_components: tuple[tuple[str, float, bool], ...] | None = None
    # Optional total NAV at start of period (EUR). When present, the
    # Summary headline ``Total P&L %`` is reported as Total P&L / NAV start
    # — the apples-to-apples denominator vs. the firm's TWR dashboard,
    # which divides the period P&L by starting fund NAV (stocks + cash
    # − accruals at 31-Dec).
    nav_total_start: float | None = None
    # Current-snapshot exposure (from end-period HiPort snapshot).
    snapshot_exposure: ExposureBlock | None = None
    # AUM-weighted daily average exposure over the full period.
    ytd_weighted_exposure: ExposureBlock | None = None


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
) -> Path:
    """Build the workbook and atomically rename it into place."""
    out_path = Path(out_path)
    t0 = time.perf_counter()

    wb = Workbook()
    wb.remove(wb.active)

    _write_summary(wb, pl_df, income_df, period_meta)
    if valuation_rows:
        _write_valuation_a(wb, valuation_rows, period_meta, pl_df,
                           th_val_df=th_val_df)
    for analyst_code in _analyst_order(pl_df):
        _write_analyst_sheet(wb, pl_df, analyst_code, period_meta)
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

    _enforce_sheet_order(wb)
    _atomic_save(wb, out_path)
    return out_path


# ─── Sheet-name parity check (BUG-2 structural lock) ────────────────────────

def _validate_layout_against_pl_df(pl_df: pd.DataFrame, layout: tuple[ColSpec, ...]) -> None:
    if pl_df.empty:
        return
    missing = [c.df_column for c in layout if c.df_column not in pl_df.columns]
    if missing:
        raise KeyError(
            f"Workbook layout references columns not in pl_df: {missing}. "
            "Either add them to pl_df or remove the ColSpec."
        )


# ─── Aggregation helpers ────────────────────────────────────────────────────

def _current_holdings(pl_df: pd.DataFrame) -> pd.DataFrame:
    if pl_df.empty:
        return pl_df
    return pl_df[pl_df["Ending Units"] != 0].reset_index(drop=True)


def _exited_positions(pl_df: pd.DataFrame) -> pd.DataFrame:
    if pl_df.empty:
        return pl_df
    return pl_df[pl_df["Ending Units"] == 0].reset_index(drop=True)


def _reduced_positions(pl_df: pd.DataFrame) -> pd.DataFrame:
    """Positions still held at period end but with sells inside the period.

    Surfaces partial sell-downs (e.g. VPS Securities) that would otherwise
    be invisible: they don't qualify for ``Exited Positions`` (Ending != 0)
    and inside ``Current Holdings`` their realised P&L from the period's
    sells is not visually distinct from buy-and-hold rows.
    """
    if pl_df.empty:
        return pl_df
    sold = pd.to_numeric(pl_df.get("Units Sold", 0.0), errors="coerce").fillna(0.0)
    held = pl_df["Ending Units"] != 0
    return pl_df[held & (sold > 0)].reset_index(drop=True)


def _with_total_contribution_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add contribution-of-total and absolute-bps columns for display tables.

    Contribution is computed against net ``Total P&L (EUR)`` in the input
    table, so the signed contributions sum to 100% (when table total is
    non-zero). Absolute bps is ``abs(contribution) * 10_000``.
    """
    out = df.copy()
    if "Total P&L (EUR)" not in out.columns:
        out["Contribution to Total (%)"] = None
        out["Absolute Contribution (bps)"] = None
        return out

    total_pl = float(pd.to_numeric(out["Total P&L (EUR)"], errors="coerce").fillna(0.0).sum())
    if abs(total_pl) < 1e-9:
        contrib = pd.Series([None] * len(out), index=out.index, dtype=object)
    else:
        contrib = pd.to_numeric(out["Total P&L (EUR)"], errors="coerce") / total_pl
    out["Contribution to Total (%)"] = contrib
    out["Absolute Contribution (bps)"] = pd.to_numeric(contrib, errors="coerce").abs() * 10_000
    return out


def _analyst_order(pl_df: pd.DataFrame) -> list[str]:
    """Distinct analyst codes in stable display order."""
    if pl_df.empty or "Analyst" not in pl_df.columns:
        return []
    found = {str(a).strip().upper() for a in pl_df["Analyst"].dropna() if str(a).strip()}
    ordered = [a for a in ANALYST_TAB_ORDER if a in found]
    extras = sorted(found - set(ordered))
    return ordered + extras


def _analyst_display(code: str) -> str:
    return ANALYST_DISPLAY.get(code, code)


def _safe_float(v) -> float:
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _pct_or_none(numerator: float, denominator: float) -> float | None:
    """Ratio (NOT percent). Returns None if denominator is ~0 to avoid div-by-zero."""
    if abs(denominator) < 1e-9:
        return None
    return numerator / denominator


def _exposure_for(df: pd.DataFrame) -> dict:
    """Compute long/short/net exposure (EUR) and P&L breakdown.

    Exposure semantics:
      * ORD positions  → Market Value End (EUR), signed by L/S.
      * SWAP / FTSWAP  → |notional| (Cost Basis EUR), signed by L/S.
        HiPort's MV for a swap is the swap's mark-to-market, not the
        underlying notional, so it can be small (or even positive on a
        winning short). Using it as exposure produced nonsense like
        "shorts with positive Short Exposure" and 300%+ Total P&L %.
    """
    if df.empty:
        return dict(
            n=0, long_eur=0.0, short_eur=0.0, net_eur=0.0, gross_eur=0.0,
            realised=0.0, unrealised=0.0, income=0.0, fx_attrib=0.0,
            total=0.0, total_pct=None, cost=0.0,
        )
    mv = pd.to_numeric(df.get("Market Value End (EUR)", 0.0), errors="coerce").fillna(0.0)
    cost = pd.to_numeric(df.get("Cost Basis (EUR)", 0.0), errors="coerce").fillna(0.0).abs()
    instr = df.get("Instrument", "ORD").astype(str).str.upper()
    ls = df.get("L/S", "L").astype(str).str.upper()
    is_deriv = instr.isin(["SWAP", "FTSWAP"])
    sign = ls.map(lambda s: -1.0 if s == "S" else 1.0)
    # ORD: signed MV (MV is already signed for shorts in HiPort, but we
    # take |MV|·sign(L/S) to get a clean L/S split). Derivatives: |notional|·sign.
    exposure_signed = (cost.where(is_deriv, mv.abs())) * sign
    long_eur = float(exposure_signed[exposure_signed > 0].sum())
    # short_eur is stored as a positive absolute value (magnitude of short book)
    # so callers can display it directly. net = long - short; gross = long + short.
    short_eur = abs(float(exposure_signed[exposure_signed < 0].sum()))
    realised = float(pd.to_numeric(df.get("Realised P&L (EUR)", 0.0), errors="coerce").fillna(0.0).sum())
    unrealised = float(pd.to_numeric(df.get("Unrealised P&L (EUR)", 0.0), errors="coerce").fillna(0.0).sum())
    income = float(pd.to_numeric(df.get("Income (EUR)", 0.0), errors="coerce").fillna(0.0).sum())
    # Use the stored cash-flow Total P&L (EUR) column so the Summary ties
    # exactly to the per-stock rows shown in Current Holdings / Exited
    # Positions / per-Analyst sheets. R+U+I differs from this by a small
    # FX-attribution residual (period-FX vs trade-FX), which would
    # otherwise show up as an inflated by-Analyst total.
    total = float(pd.to_numeric(df.get("Total P&L (EUR)", 0.0), errors="coerce").fillna(0.0).sum())
    # FX attribution = the residual that makes R + U + I + FX_attrib = Total.
    # It captures the difference between trade-day FX (baked into realised
    # P&L) and period-end FX (used for unrealised MV translation), plus any
    # rounding. Surfacing it explicitly means the breakdown always sums to
    # the cash-flow Total shown per stock.
    fx_attrib = total - (realised + unrealised + income)
    cost_total = float(cost.sum())
    # Return % uses COST BASIS (capital deployed) as the denominator so the
    # by-Analyst row is methodologically consistent with each position's
    # own Total P&L (%) column (which also divides by Cost Basis). Using
    # end-period gross exposure would silently exclude capital deployed in
    # positions that were exited or trimmed inside the period — VS in
    # particular runs many round-trips, so end-MV under-counts his book
    # while sum-of-cost-basis captures it.
    return dict(
        n=int(len(df)),
        long_eur=long_eur, short_eur=short_eur,
        net_eur=long_eur - short_eur,      # net = long − |short|
        gross_eur=long_eur + short_eur,    # gross = long + |short|
        realised=realised, unrealised=unrealised, income=income,
        fx_attrib=fx_attrib,
        total=total, total_pct=_pct_or_none(total, cost_total),
        cost=cost_total,
    )


# ─── Common sheet helpers ───────────────────────────────────────────────────

_THIN = Side(style="thin", color="D5DDDE")
_MED = Side(style="medium", color=COLOR_HEADER)
BORDER_THIN_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
BORDER_TOP_THICK = Border(top=_MED)
BORDER_KPI = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _setup_sheet(ws: Worksheet, *, tab_color: str | None = None,
                 freeze: str | None = None, hide_gridlines: bool = True,
                 zoom: int = 100) -> None:
    if tab_color:
        ws.sheet_properties.tabColor = tab_color
    if hide_gridlines:
        ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    if freeze:
        ws.freeze_panes = freeze


def _write_title_block(ws: Worksheet, meta: PeriodMeta, *,
                       title: str | None = None, subtitle: str | None = None,
                       col_start: int = 1) -> None:
    title = title or f"{meta.fund_name} — Stock-Level P&L Report"
    subtitle = subtitle or (
        f"Period: {meta.start_date.date():%d %B %Y} → "
        f"{meta.end_date.date():%d %B %Y}  |  Fund Currency: EUR"
    )
    ws.row_dimensions[1].height = 30
    c = ws.cell(1, col_start, title)
    c.font = FONT_TITLE
    c.alignment = ALIGN_LEFT
    c2 = ws.cell(2, col_start, subtitle)
    c2.font = FONT_BODY_MUTED
    c2.alignment = ALIGN_LEFT


def _write_section_header(ws: Worksheet, row: int, *, title: str,
                          span_cols: int, accent: bool = False,
                          col_start: int = 1) -> None:
    fill = FILL_HEADER_ACCENT if accent else FILL_HEADER
    ws.row_dimensions[row].height = 22
    for c in range(col_start, col_start + span_cols):
        cell = ws.cell(row, c)
        cell.fill = fill
        cell.font = FONT_SECTION
        cell.alignment = ALIGN_LEFT
    ws.cell(row, col_start, "  " + title).font = FONT_SECTION


def _kpi_card(ws: Worksheet, *, row: int, col: int, label: str, value,
              kind: str = "eur", colour_signed: bool = False,
              accent: bool = False) -> None:
    """Render a 2-row × 2-col card: label on top, value below.

    Cell range: (row, col)..(row+1, col+1)
    """
    fill = FILL_KPI_ACCENT if accent else FILL_KPI

    # Label cells (row, col)..(row, col+1) merged
    ws.merge_cells(start_row=row, start_column=col,
                   end_row=row, end_column=col + 1)
    lc = ws.cell(row, col, label.upper())
    lc.font = FONT_KPI_LABEL
    lc.alignment = ALIGN_KPI_LABEL
    lc.fill = fill
    lc.border = BORDER_KPI
    ws.cell(row, col + 1).fill = fill
    ws.cell(row, col + 1).border = BORDER_KPI

    # Value cells (row+1, col)..(row+1, col+1) merged
    ws.merge_cells(start_row=row + 1, start_column=col,
                   end_row=row + 1, end_column=col + 1)
    vc = ws.cell(row + 1, col, _coerce_value(value, kind))
    vc.alignment = ALIGN_KPI_VALUE
    vc.fill = fill
    vc.border = BORDER_KPI
    if value is None:
        vc.font = FONT_KPI_VALUE
    elif colour_signed and isinstance(vc.value, (int, float)):
        vc.font = FONT_KPI_VALUE_POS if vc.value >= 0 else FONT_KPI_VALUE_NEG
    else:
        vc.font = FONT_KPI_VALUE
    if kind != "text":
        vc.number_format = fmt_for(kind)
    ws.cell(row + 1, col + 1).fill = fill
    ws.cell(row + 1, col + 1).border = BORDER_KPI

    ws.row_dimensions[row].height = 18
    ws.row_dimensions[row + 1].height = 26


# ─── Exposure block table helper ────────────────────────────────────────────

def _write_exposure_block_table(
    ws: Worksheet,
    block: ExposureBlock,
    analysts: list[str],
    meta: PeriodMeta,
    *,
    start_row: int,
    is_ytd: bool = False,
    col_start: int = 1,
    show_total: bool = True,
) -> int:
    """Write a current-snapshot or time-weighted average exposure table.

    ``col_start`` lets callers shift the table right to skip hidden columns
    (e.g. the hidden ISIN column on per-analyst sheets uses col_start=2).
    ``show_total=False`` suppresses the fund-total row (used on analyst sheets
    where the fund total is irrelevant and confusing).

    Columns: Analyst | Long (EUR) | Long (%) | Short (EUR) | Short (%) |
             Net (EUR) | Net (%) | Gross (EUR) | Gross (%)

    Returns the next available row after the table.
    """
    if block.as_of is not None:
        date_label = "through" if is_ytd else "as of"
        as_of_str = f"  ({date_label} {block.as_of.date():%d %B %Y})"
    else:
        as_of_str = ""
    title = block.label + as_of_str
    _write_section_header(ws, start_row, title=title, span_cols=9,
                          accent=is_ytd, col_start=col_start)
    start_row += 1

    prefix = "YTD Avg " if is_ytd else "Current "
    hdrs = (
        "Analyst",
        f"{prefix}Long ({meta.fund_currency})", f"{prefix}Long (%)",
        f"{prefix}Short ({meta.fund_currency})", f"{prefix}Short (%)",
        f"{prefix}Net ({meta.fund_currency})", f"{prefix}Net (%)",
        f"{prefix}Gross ({meta.fund_currency})", f"{prefix}Gross (%)",
    )
    kinds = ("text", "eur", "pct", "eur", "pct", "eur", "pct", "eur", "pct")
    hdr_fill = FILL_HEADER_ACCENT if is_ytd else FILL_HEADER
    for i, h in enumerate(hdrs, start=col_start):
        cc = ws.cell(start_row, i, h)
        cc.font = FONT_HEADER; cc.fill = hdr_fill; cc.alignment = ALIGN_CENTER
    start_row += 1

    for r_idx, code in enumerate(analysts):
        em = block.by_analyst.get(code, ExposureMetrics())
        vals = (
            _analyst_display(code),
            em.long_eur, em.pct_of_nav(em.long_eur),
            em.short_eur, em.pct_of_nav(em.short_eur),
            em.net_eur, em.pct_of_nav(em.net_eur),
            em.gross_eur, em.pct_of_nav(em.gross_eur),
        )
        for c_idx, (v, k) in enumerate(zip(vals, kinds), start=col_start):
            cell = ws.cell(start_row + r_idx, c_idx, _coerce_value(v, k))
            cell.font = FONT_BODY
            if r_idx % 2 == 1:
                cell.fill = FILL_BAND
            cell.alignment = align_for(k)
            if k != "text":
                cell.number_format = fmt_for(k)

    last_data_row = start_row + len(analysts) - 1

    if show_total:
        # Fund total row — labelled "Fund Total" to distinguish from per-analyst rows.
        em = block.fund
        total_row = last_data_row + 1
        vals = (
            "Fund Total",
            em.long_eur, em.pct_of_nav(em.long_eur),
            em.short_eur, em.pct_of_nav(em.short_eur),
            em.net_eur, em.pct_of_nav(em.net_eur),
            em.gross_eur, em.pct_of_nav(em.gross_eur),
        )
        for c_idx, (v, k) in enumerate(zip(vals, kinds), start=col_start):
            cell = ws.cell(total_row, c_idx, _coerce_value(v, k))
            cell.font = FONT_TOTAL; cell.fill = FILL_TOTAL
            cell.alignment = align_for(k)
            if k != "text":
                cell.number_format = fmt_for(k)
        return total_row + 1

    return last_data_row + 1


# ─── Summary sheet ──────────────────────────────────────────────────────────

def _write_summary(wb: Workbook, pl_df: pd.DataFrame, income_df: pd.DataFrame,
                   meta: PeriodMeta) -> None:
    ws = wb.create_sheet("Summary")
    _setup_sheet(ws, tab_color=COLOR_TAB_DASHBOARD, zoom=100)
    _write_title_block(ws, meta)

    # Column widths sized to the "By Analyst" table headers (longest row in sheet).
    # Col A: analyst name. Cols E-F: "Long/Short Exposure (EUR)" = 20 chars each.
    # Col K: "Income & Financing (EUR)" = 24 chars. Col L: "FX Attribution (EUR)" = 20 chars.
    summary_widths = [18, 11, 17, 13, 21, 22, 20, 22, 14, 16, 26, 22]
    for i, w in enumerate(summary_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # If empty, short-circuit with a clean message and stop.
    if pl_df.empty:
        ws.cell(4, 1, "No positions in this period.").font = FONT_BODY_MUTED
        return

    current = _current_holdings(pl_df)
    exited = _exited_positions(pl_df)
    fund_kpi = _exposure_for(pl_df)
    cur_kpi = _exposure_for(current)
    ex_kpi = _exposure_for(exited)

    # ── KPI cards: 3 columns, 4 rows of cards (left block, cols 1-6) ──
    # Card grid layout: each card occupies 2 columns and 2 rows. Cards are
    # stacked back-to-back (no spacer rows) per user request — the alternating
    # accent fill provides enough visual separation.
    card_rows = (4, 6, 8, 10)
    card_cols = (1, 3, 5)

    # Headline Total P&L %: use the per-strategy weighted % so the headline
    # ties to the strategy rows below (gross-exposure denominator). The
    # NAV-start % is shown separately in the NAV Reconciliation block.
    cards = [
        ("Total Positions", len(pl_df), "int", False, False),
        ("Current Holdings", cur_kpi["n"], "int", False, False),
        ("Exited Positions", ex_kpi["n"], "int", False, False),

        ("Realised P&L (EUR)", fund_kpi["realised"], "eur", True, False),
        ("Unrealised P&L (EUR)", fund_kpi["unrealised"], "eur", True, False),
        ("Income & Financing (EUR)", fund_kpi["income"], "eur", True, False),

        ("FX Attribution (EUR)", fund_kpi["fx_attrib"], "eur", True, False),
        ("Total P&L (EUR)", fund_kpi["total"], "eur", True, True),
        ("Total P&L %", fund_kpi["total_pct"], "pct", True, True),
    ]
    for i, (label, value, kind, signed, accent) in enumerate(cards):
        r = card_rows[i // 3]
        c = card_cols[i % 3]
        _kpi_card(ws, row=r, col=c, label=label, value=value,
                  kind=kind, colour_signed=signed, accent=accent)

    # ── Instrument breakdown table ──
    base_row = card_rows[-1] + 3   # one blank row between KPI cards and table
    _write_section_header(ws, base_row, title="Instrument Breakdown",
                          span_cols=4)
    base_row += 1

    inst_headers = ("Type", "Positions", f"Total P&L ({meta.fund_currency})", "Total P&L %")
    inst_kinds = ("text", "int", "eur", "pct")
    for i, h in enumerate(inst_headers, start=1):
        cc = ws.cell(base_row, i, h); cc.font = FONT_HEADER
        cc.fill = FILL_HEADER; cc.alignment = ALIGN_CENTER
    base_row += 1
    g = pl_df.groupby("Instrument", sort=False)
    rows: list[tuple] = []
    for inst, sub in g:
        e = _exposure_for(sub)
        rows.append((str(inst), e["n"], e["total"], e["total_pct"]))
    rows.append(("TOTAL", fund_kpi["n"], fund_kpi["total"], fund_kpi["total_pct"]))
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(base_row + r_idx, c_idx,
                           _coerce_value(val, inst_kinds[c_idx - 1]))
            kind = inst_kinds[c_idx - 1]
            if r_idx == len(rows) - 1:
                cell.font = FONT_TOTAL
                cell.fill = FILL_TOTAL
            else:
                cell.font = FONT_BODY
                if r_idx % 2 == 1:
                    cell.fill = FILL_BAND
            cell.alignment = align_for(kind)
            if kind != "text":
                cell.number_format = fmt_for(kind)
            if kind in ("eur", "pct") and isinstance(cell.value, (int, float)) and cell.value < 0:
                cell.font = FONT_NEG_BOLD if cell.font is FONT_TOTAL else FONT_NEG
    ws.row_dimensions[base_row - 1].height = 22

    # ════════════════════════════════════════════════════════════════════════
    # PERIOD P&L ATTRIBUTION (YTD)
    # ════════════════════════════════════════════════════════════════════════
    base_row = base_row + len(rows) + 1
    _write_section_header(ws, base_row,
                          title="By Analyst — Period P&L Attribution (YTD)",
                          span_cols=8, accent=True)
    base_row += 1

    # All columns here are PERIOD metrics (YTD, includes exited positions).
    # Return %: uses YTD avg gross (absolute) exposure as denominator when NAV history is
    # available (Stefan's requirement); falls back to cost basis otherwise.
    ytd_exp = meta.ytd_weighted_exposure
    ret_pct_label = ("Return % (YTD Gross Exp)" if ytd_exp is not None
                     else "Total P&L % (Cost Basis)")
    a_headers = ("Analyst", "Open Positions",
                 "Total P&L (EUR)", ret_pct_label,
                 "Realised (EUR)", "Unrealised (EUR)",
                 "Income & Financing (EUR)", "FX Attribution (EUR)")
    a_kinds = ("text", "int", "eur", "pct", "eur", "eur", "eur", "eur")
    for i, h in enumerate(a_headers, start=1):
        cc = ws.cell(base_row, i, h); cc.font = FONT_HEADER
        cc.fill = FILL_HEADER_ACCENT; cc.alignment = ALIGN_CENTER
    base_row += 1

    def _ret_pct(total: float, code: str) -> float | None:
        """Return % — YTD avg gross (absolute) exposure denominator if available, else cost basis."""
        if ytd_exp is not None:
            em = ytd_exp.by_analyst.get(code, ExposureMetrics())
            return _pct_or_none(total, em.gross_eur) if abs(em.gross_eur) > 1e-9 else None
        return None  # will be filled from e["total_pct"] below

    a_rows: list[tuple] = []
    a_rows_with_codes: list[tuple[str, tuple]] = []  # (code, row_tuple) for sorting
    for code in _analyst_order(pl_df):
        sub = pl_df[pl_df["Analyst"].astype(str).str.upper() == code]
        cur_sub = sub[sub["Ending Units"] != 0]
        e = _exposure_for(sub)
        ret = _ret_pct(e["total"], code) if ytd_exp is not None else e["total_pct"]
        row = (_analyst_display(code), int(len(cur_sub)),
               e["total"], ret,
               e["realised"], e["unrealised"], e["income"], e["fx_attrib"])
        a_rows_with_codes.append((code, row))
    
    # Sort by Total P&L (EUR) descending (column index 2)
    a_rows_with_codes.sort(key=lambda x: x[1][2] if isinstance(x[1][2], (int, float)) else 0, 
                           reverse=True)
    
    # Extract sorted codes and rows
    sorted_analyst_codes = [code for code, _ in a_rows_with_codes]
    a_rows = [row for _, row in a_rows_with_codes]
    
    fund_ret = (_pct_or_none(fund_kpi["total"], ytd_exp.fund.gross_eur)
                if ytd_exp is not None and abs(ytd_exp.fund.gross_eur) > 1e-9
                else fund_kpi["total_pct"])
    a_rows.append(("TOTAL", int(len(_current_holdings(pl_df))),
                   fund_kpi["total"], fund_ret,
                   fund_kpi["realised"], fund_kpi["unrealised"],
                   fund_kpi["income"], fund_kpi["fx_attrib"]))
    for r_idx, row in enumerate(a_rows):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(base_row + r_idx, c_idx,
                           _coerce_value(val, a_kinds[c_idx - 1]))
            kind = a_kinds[c_idx - 1]
            if r_idx == len(a_rows) - 1:
                cell.font = FONT_TOTAL; cell.fill = FILL_TOTAL
            else:
                cell.font = FONT_BODY
                if r_idx % 2 == 1:
                    cell.fill = FILL_BAND
            cell.alignment = align_for(kind)
            if kind != "text":
                cell.number_format = fmt_for(kind)
            if kind in ("eur", "pct") and isinstance(cell.value, (int, float)) and cell.value < 0:
                cell.font = FONT_NEG_BOLD if cell.font is FONT_TOTAL else FONT_NEG

    end_row = base_row + len(a_rows)
    analysts = sorted_analyst_codes

    # ════════════════════════════════════════════════════════════════════════
    # CURRENT PORTFOLIO SNAPSHOT
    # ════════════════════════════════════════════════════════════════════════
    sec_row = end_row + 2
    _write_section_header(ws, sec_row,
                          title="Current Portfolio Snapshot",
                          span_cols=9, accent=False)
    sec_row += 1

    if meta.snapshot_exposure is not None:
        sec_row = _write_exposure_block_table(
            ws, meta.snapshot_exposure, analysts, meta, start_row=sec_row, is_ytd=False,
        ) + 1
    else:
        ws.cell(sec_row, 1,
                "Current snapshot unavailable — end-period snapshot not loaded."
                ).font = FONT_BODY_MUTED
        sec_row += 2
    if meta.ytd_weighted_exposure is not None:
        sec_row = _write_exposure_block_table(
            ws, meta.ytd_weighted_exposure, analysts, meta, start_row=sec_row, is_ytd=True,
        ) + 1
        # Return on Gross Exposure (YTD) — period P&L / YTD avg gross (absolute) exposure
        _write_section_header(ws, sec_row, title="Return on Gross Exposure (YTD)",
                              span_cols=4)
        sec_row += 1
        ren_hdrs = ("Analyst", "Total P&L (EUR)", "YTD Avg Gross Exp (EUR)", "Return %")
        ren_kinds = ("text", "eur", "eur", "pct")
        for i, h in enumerate(ren_hdrs, start=1):
            cc = ws.cell(sec_row, i, h)
            cc.font = FONT_HEADER; cc.fill = FILL_HEADER; cc.alignment = ALIGN_CENTER
        sec_row += 1
        ytd = meta.ytd_weighted_exposure
        
        # Build return table data and sort by Return % descending
        ren_rows: list[tuple[str, float, float, float | None]] = []
        for code in analysts:
            sub = pl_df[pl_df["Analyst"].astype(str).str.upper() == code]
            pl_total = float(pd.to_numeric(sub.get("Total P&L (EUR)", 0.0),
                                           errors="coerce").fillna(0.0).sum())
            em = ytd.by_analyst.get(code, ExposureMetrics())
            ret = _pct_or_none(pl_total, em.gross_eur) if abs(em.gross_eur) > 1e-9 else None
            ren_rows.append((_analyst_display(code), pl_total, em.gross_eur, ret))
        
        # Sort by Return % descending (index 3), NaN values last
        ren_rows.sort(key=lambda x: (x[3] is None, -x[3] if x[3] is not None else 0))
        
        # Render sorted rows
        for r_idx, vals in enumerate(ren_rows):
            for c_idx, (v, k) in enumerate(zip(vals, ren_kinds), start=1):
                cell = ws.cell(sec_row + r_idx, c_idx, _coerce_value(v, k))
                cell.font = FONT_BODY
                if r_idx % 2 == 1:
                    cell.fill = FILL_BAND
                cell.alignment = align_for(k)
                if k != "text":
                    cell.number_format = fmt_for(k)
                if k in ("eur", "pct") and isinstance(cell.value, (int, float)) and cell.value < 0:
                    cell.font = FONT_NEG
        
        pl_fund = fund_kpi["total"]
        em_fund = ytd.fund
        ret_fund = _pct_or_none(pl_fund, em_fund.gross_eur) if abs(em_fund.gross_eur) > 1e-9 else None
        for c_idx, (v, k) in enumerate(zip(
            ("TOTAL", pl_fund, em_fund.gross_eur, ret_fund), ren_kinds
        ), start=1):
            cell = ws.cell(sec_row + len(ren_rows), c_idx, _coerce_value(v, k))
            cell.font = FONT_TOTAL; cell.fill = FILL_TOTAL
            cell.alignment = align_for(k)
            if k != "text":
                cell.number_format = fmt_for(k)
            if k in ("eur", "pct") and isinstance(cell.value, (int, float)) and cell.value < 0:
                cell.font = FONT_NEG_BOLD
        sec_row += len(ren_rows) + 2
    elif meta.snapshot_exposure is not None:
        # Snapshot available but no NAV history — show a note
        ws.cell(sec_row, 1,
                "YTD AUM-weighted exposure unavailable — NAV history not loaded "
                "(run with SQL access to NAV.dbo.tNAV)."
                ).font = FONT_BODY_MUTED
        sec_row += 2

    # Footer
    foot_row = sec_row + 1
    fc = ws.cell(foot_row, 1,
                 f"Generated {meta.run_timestamp:%Y-%m-%d %H:%M:%S}  |  "
                 f"Snapshot: {Path(meta.snapshot_path).name}  |  "
                 f"Bottler: {Path(meta.bottler_path).name}")
    fc.font = FONT_BODY_MUTED


# ─── Per-analyst sheet ──────────────────────────────────────────────────────

def _write_analyst_sheet(wb: Workbook, pl_df: pd.DataFrame,
                         analyst_code: str, meta: PeriodMeta) -> None:
    # Tab name = code (short, stable, used by tests and navigation).
    # Sheet title = full name for readability inside the sheet.
    name = _safe_sheet_name(analyst_code)
    ws = wb.create_sheet(name)
    _setup_sheet(ws, tab_color=COLOR_TAB_ANALYST, zoom=100)
    _write_title_block(ws, meta,
                       title=f"Analyst: {_analyst_display(analyst_code)} ({analyst_code})",
                       subtitle=(f"Period: {meta.start_date.date():%d %B %Y} → "
                                 f"{meta.end_date.date():%d %B %Y}  |  Fund Currency: EUR"),
                       col_start=2)

    sub = pl_df[pl_df["Analyst"].astype(str).str.upper() == analyst_code].copy()
    cur = sub[sub["Ending Units"] != 0].reset_index(drop=True)
    exi = sub[sub["Ending Units"] == 0].reset_index(drop=True)
    
    # Sort both sections by Total P&L (EUR) descending
    if "Total P&L (EUR)" in cur.columns:
        cur = cur.sort_values("Total P&L (EUR)", ascending=False, na_position="last").reset_index(drop=True)
    if "Total P&L (EUR)" in exi.columns:
        exi = exi.sort_values("Total P&L (EUR)", ascending=False, na_position="last").reset_index(drop=True)
    cur = _with_total_contribution_columns(cur)
    exi = _with_total_contribution_columns(exi)
    sub = _with_total_contribution_columns(sub)
    layout = COLUMN_LAYOUT_ANALYST
    n_cols = len(layout)

    # Column widths
    for i, spec in enumerate(layout, start=1):
        ws.column_dimensions[get_column_letter(i)].width = spec.width

    # Hide the ISIN column on per-analyst sheets (kept for copy-paste /
    # downstream lookups but visually noisy for the analyst view). It
    # remains visible on the top-level Current Holdings / Exited Positions
    # / Reduced Positions / Trade Detail data sheets.
    isin_col_idx = next(
        (i for i, spec in enumerate(layout, start=1) if spec.df_column == "ISIN"),
        None,
    )
    if isin_col_idx is not None:
        ws.column_dimensions[get_column_letter(isin_col_idx)].hidden = True

    # Top KPI row: period P&L + return-on-net + current exposure.
    # Period P&L covers all positions including those exited inside the period.
    # Current exposure uses only positions held at period end (matching HiPort snapshot).
    e = _exposure_for(sub)
    e_cur = _exposure_for(cur)
    em_ytd = (
        meta.ytd_weighted_exposure.by_analyst.get(analyst_code, ExposureMetrics())
        if meta.ytd_weighted_exposure is not None
        else ExposureMetrics()
    )
    ret_on_gross_ytd = (
        _pct_or_none(e["total"], em_ytd.gross_eur)
        if meta.ytd_weighted_exposure is not None and abs(em_ytd.gross_eur) > 1e-9
        else None
    )
    kpi_row = 4
    # Col 1 is the hidden ISIN column — start all KPI cards at col 2 so they
    # land in the first visible column (Stock Name).
    # Period P&L cards
    _kpi_card(ws, row=kpi_row, col=2, label="Period — Positions",
              value=int(len(cur)), kind="int", accent=False)
    _kpi_card(ws, row=kpi_row, col=4, label="Period — Total P&L (EUR)",
              value=e["total"], kind="eur", colour_signed=True, accent=True)
    _kpi_card(ws, row=kpi_row, col=6, label="Period — Total P&L %",
              value=e["total_pct"], kind="pct", colour_signed=True, accent=True)
    _kpi_card(ws, row=kpi_row, col=8, label="Return on Gross Exposure (YTD)",
              value=ret_on_gross_ytd, kind="pct", colour_signed=True, accent=True)
    # Current snapshot exposure cards
    _kpi_card(ws, row=kpi_row, col=10, label="Current — Net Exposure (EUR)",
              value=e_cur["net_eur"], kind="eur", accent=False)
    _kpi_card(ws, row=kpi_row, col=12, label="Current — Gross Exposure (EUR)",
              value=e_cur["gross_eur"], kind="eur", accent=False)

    # ── Current Holdings section ──
    # Section headers and "(none)" messages use col_start=2 so the title text
    # lands in the first visible column (col 1 is the hidden ISIN column).
    visible_span = n_cols - 1  # cols 2..n_cols
    section_row = kpi_row + 4
    _write_section_header(ws, section_row, title="Current Holdings",
                          span_cols=visible_span, col_start=2)
    header_row = section_row + 1
    _write_layout_headers(ws, header_row, layout)
    next_row = header_row + 1
    if cur.empty:
        empty = ws.cell(next_row, 2, "(none)")
        empty.font = FONT_BODY_MUTED
        next_row += 1
    else:
        next_row = _write_layout_data(ws, cur, layout, start_row=next_row)
        next_row = _write_layout_total(ws, cur, layout, total_row=next_row,
                                        label="Subtotal — Current",
                                        fill=FILL_SUBTOTAL)
    next_row += 1
    
    # Write merged "Total" row before Exited Positions (columns B-P, centered)
    ws.merge_cells(start_row=next_row, start_column=2,
                   end_row=next_row, end_column=n_cols)
    total_cell = ws.cell(next_row, 2, "Total")
    total_cell.font = FONT_TOTAL
    total_cell.fill = FILL_TOTAL
    total_cell.alignment = ALIGN_CENTER
    ws.row_dimensions[next_row].height = 20
    next_row += 1

    # ── Exited Positions section ──
    _write_section_header(ws, next_row, title="Exited Positions",
                          span_cols=visible_span, accent=True, col_start=2)
    header_row = next_row + 1
    # For exited positions Units & Last Px are stale (units=0) — hide them.
    EXITED_SUPPRESS = frozenset({"Ending Units", "Last Price (EUR)"})
    _write_layout_headers(ws, header_row, layout, accent=True,
                          suppress=EXITED_SUPPRESS)
    next_row = header_row + 1
    if exi.empty:
        empty = ws.cell(next_row, 2, "(none)")
        empty.font = FONT_BODY_MUTED
        next_row += 1
    else:
        next_row = _write_layout_data(ws, exi, layout, start_row=next_row,
                                       suppress=EXITED_SUPPRESS)
        next_row = _write_layout_total(ws, exi, layout, total_row=next_row,
                                        label="Subtotal — Exited",
                                        fill=FILL_SUBTOTAL,
                                        suppress=EXITED_SUPPRESS)
    next_row += 2

    # ── Grand total (preceded by a fresh header row so users don't have
    # to scroll back up to remember which column is which) ──
    _write_layout_headers(ws, next_row, layout, accent=True)
    next_row += 1
    # Per-position attributes don't aggregate meaningfully across the book.
    TOTAL_SUPPRESS = frozenset({"CCY", "Instrument", "L/S",
                                 "Ending Units", "Last Price (EUR)",
                                 "Avg Buy Price (EUR)",
                                 "Avg Sell Price (EUR)"})
    next_row = _write_layout_total(ws, sub, layout, total_row=next_row,
                                    label="ANALYST TOTAL", fill=FILL_TOTAL,
                                    suppress=TOTAL_SUPPRESS)

    # ── Current Snapshot & YTD weighted average exposure ──
    # Each block writes its own labelled section header (9 cols wide) so no outer
    # wrapper is needed — that was creating a wide empty "tail" past column I.
    # show_total=False hides the fund-total row; analysts only care about their own.
    next_row += 2
    if meta.snapshot_exposure is not None:
        next_row = _write_exposure_block_table(
            ws, meta.snapshot_exposure, [analyst_code],
            meta,
            start_row=next_row, is_ytd=False, col_start=2, show_total=False,
        ) + 1
    if meta.ytd_weighted_exposure is not None:
        next_row = _write_exposure_block_table(
            ws, meta.ytd_weighted_exposure, [analyst_code],
            meta,
            start_row=next_row, is_ytd=True, col_start=2, show_total=False,
        ) + 1


# ─── Generic table writers ──────────────────────────────────────────────────

def _write_layout_headers(ws: Worksheet, row: int, layout: tuple[ColSpec, ...],
                          *, accent: bool = False,
                          suppress: frozenset[str] = frozenset()) -> None:
    fill = FILL_HEADER_ACCENT if accent else FILL_HEADER
    ws.row_dimensions[row].height = 24
    for col_idx, spec in enumerate(layout, start=1):
        label = "" if spec.df_column in suppress else (spec.display or spec.df_column)
        cell = ws.cell(row, col_idx, label)
        cell.font = FONT_HEADER
        cell.fill = fill
        cell.alignment = ALIGN_CENTER


def _write_layout_data(ws: Worksheet, df: pd.DataFrame,
                       layout: tuple[ColSpec, ...], *, start_row: int,
                       suppress: frozenset[str] = frozenset()) -> int:
    for row_offset, (_, row) in enumerate(df.iterrows()):
        excel_row = start_row + row_offset
        banded = (row_offset % 2 == 1)
        for col_idx, spec in enumerate(layout, start=1):
            if spec.df_column in suppress:
                cell = ws.cell(excel_row, col_idx)
                cell.font = FONT_BODY
                if banded:
                    cell.fill = FILL_BAND
                continue
            value = row.get(spec.df_column, None) if spec.df_column in df.columns else None
            cell = ws.cell(excel_row, col_idx, _coerce_value(value, spec.kind))
            cell.font = FONT_BODY
            cell.number_format = fmt_for(spec.kind)
            cell.alignment = align_for(spec.kind)
            if banded:
                cell.fill = FILL_BAND
            # Sign-coloured P&L cells
            if spec.kind == "eur" and isinstance(cell.value, (int, float)):
                if "P&L" in spec.df_column or "Income" in spec.df_column:
                    if cell.value < 0:
                        cell.font = FONT_NEG
                    elif cell.value > 0 and "P&L" in spec.df_column:
                        cell.font = FONT_POS
    return start_row + len(df)


def _write_layout_total(ws: Worksheet, df: pd.DataFrame,
                        layout: tuple[ColSpec, ...], *, total_row: int,
                        label: str = "Total", fill: PatternFill = FILL_TOTAL,
                        suppress: frozenset[str] = frozenset()) -> int:
    ws.row_dimensions[total_row].height = 20
    for col_idx, spec in enumerate(layout, start=1):
        cell = ws.cell(total_row, col_idx)
        cell.font = FONT_TOTAL
        cell.fill = fill
        cell.alignment = align_for(spec.kind)
        cell.border = BORDER_TOP_THICK
        if spec.df_column in suppress:
            continue
        if spec.total == "label" or col_idx == 1:
            cell.value = label
        elif spec.total == "sum" and spec.df_column in df.columns:
            s = pd.to_numeric(df[spec.df_column], errors="coerce").sum()
            cell.value = float(s)
            cell.number_format = fmt_for(spec.kind)
            if spec.kind == "eur" and float(s) < 0:
                cell.font = FONT_NEG_BOLD
        elif spec.df_column == "Contribution to Total (%)":
            pl = pd.to_numeric(df.get("Total P&L (EUR)",
                                      pd.Series(dtype=float)),
                               errors="coerce").sum()
            if abs(float(pl)) > 1e-9:
                cell.value = 1.0
                cell.number_format = fmt_for(spec.kind)
        elif spec.df_column == "Total P&L (%)":
            pl = pd.to_numeric(df.get("Total P&L (EUR)",
                                       pd.Series(dtype=float)),
                               errors="coerce").sum()
            cb = pd.to_numeric(df.get("Cost Basis (EUR)",
                                       pd.Series(dtype=float)),
                               errors="coerce").sum()
            if cb:
                ratio = float(pl) / abs(float(cb))
                cell.value = ratio
                cell.number_format = fmt_for(spec.kind)
                if ratio < 0:
                    cell.font = FONT_NEG_BOLD
    return total_row + 1


def _write_table(wb: Workbook, sheet_name: str, df: pd.DataFrame,
                 layout: tuple[ColSpec, ...], meta: PeriodMeta,
                 *, tab_color: str | None = None) -> None:
    ws = wb.create_sheet(sheet_name)
    _setup_sheet(ws, tab_color=tab_color)
    _write_title_block(ws, meta,
                       title=f"{sheet_name} — {meta.fund_name}",
                       subtitle=(f"Period: {meta.start_date.date():%d %B %Y} → "
                                 f"{meta.end_date.date():%d %B %Y}"))
    n_cols = len(layout)
    # Column widths
    for i, spec in enumerate(layout, start=1):
        ws.column_dimensions[get_column_letter(i)].width = spec.width

    if df.empty:
        ws.cell(4, 1, "(no rows)").font = FONT_BODY_MUTED
        return

    if any(
        spec.df_column in {"Contribution to Total (%)", "Absolute Contribution (bps)"}
        for spec in layout
    ):
        df = _with_total_contribution_columns(df)

    if sheet_name == "Current Holdings":
        _validate_layout_against_pl_df(df, COLUMN_LAYOUT_HOLDINGS)

    # Sort by Total P&L (EUR) descending
    if "Total P&L (EUR)" in df.columns:
        df = df.sort_values("Total P&L (EUR)", ascending=False, na_position="last").reset_index(drop=True)

    header_row = 4
    _write_layout_headers(ws, header_row, layout)
    
    # Add merged "Total" row right after header (for Exited/Reduced Positions)
    if sheet_name in ("Exited Positions", "Reduced Positions"):
        total_header_row = header_row + 1
        ws.merge_cells(start_row=total_header_row, start_column=2,
                       end_row=total_header_row, end_column=n_cols)
        total_cell = ws.cell(total_header_row, 2, "Total")
        total_cell.font = FONT_TOTAL
        total_cell.fill = FILL_TOTAL
        total_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[total_header_row].height = 20
        next_row = total_header_row + 1
    else:
        next_row = header_row + 1
    
    next_row = _write_layout_data(ws, df, layout, start_row=next_row)
    if any(spec.total != "blank" for spec in layout):
        next_row = _write_layout_total(ws, df, layout, total_row=next_row,
                                        label="Total", fill=FILL_TOTAL)

    # Freeze panes below header row
    ws.freeze_panes = ws.cell(header_row + 1, 1)
    
    # Auto filter range - adjust for Total row if present
    if sheet_name in ("Exited Positions", "Reduced Positions"):
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(n_cols)}{header_row + 1 + len(df)}"
    else:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(n_cols)}{header_row + len(df)}"


# ─── Income summary & trade detail sheets ───────────────────────────────────

def _write_income_summary(wb: Workbook, pl_df: pd.DataFrame,
                          meta: PeriodMeta) -> None:
    ws = wb.create_sheet("Income & Dividends")
    _setup_sheet(ws, tab_color=COLOR_TAB_DATA)
    _write_title_block(ws, meta,
                       title=f"Income & Dividends — {meta.fund_name}",
                       subtitle="By position")

    if pl_df.empty:
        ws.cell(4, 1, "(no positions)").font = FONT_BODY_MUTED
        return

    sub = pl_df[pl_df["Income (EUR)"] != 0][[
        "ISIN", "Stock Name", "Analyst", "CCY",
        "Income (Local)", "Income (EUR)", "Income Yield (%)",
    ]].reset_index(drop=True)

    layout = (
        ColSpec("ISIN",              14, "text",  total="label"),
        ColSpec("Stock Name",        32, "text"),
        ColSpec("Analyst",           10, "text"),
        ColSpec("CCY",                6, "text"),
        ColSpec("Income (Local)",    14, "float"),
        ColSpec("Income (EUR)",      14, "eur",   total="sum",
                display="Income & Financing (EUR)"),
        ColSpec("Income Yield (%)",  14, "pct"),
    )
    n_cols = len(layout)
    for i, spec in enumerate(layout, start=1):
        ws.column_dimensions[get_column_letter(i)].width = spec.width

    if sub.empty:
        ws.cell(4, 1, "(no income)").font = FONT_BODY_MUTED
        return

    header_row = 4
    _write_layout_headers(ws, header_row, layout)
    next_row = header_row + 1
    next_row = _write_layout_data(ws, sub, layout, start_row=next_row)
    next_row = _write_layout_total(ws, sub, layout, total_row=next_row)
    ws.freeze_panes = ws.cell(header_row + 1, 1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(n_cols)}{header_row + len(sub)}"


def _write_trade_detail(wb: Workbook, trades_df: pd.DataFrame,
                        meta: PeriodMeta) -> None:
    ws = wb.create_sheet("Trade Detail")
    _setup_sheet(ws, tab_color=COLOR_TAB_DATA)
    _write_title_block(ws, meta,
                       title=f"Trade Detail — {meta.fund_name}",
                       subtitle="All trades in the analysis window")

    if trades_df.empty:
        ws.cell(4, 1, "(no trades)").font = FONT_BODY_MUTED
        return

    cols = ["TRADE_ID", "PCODE_ORIG", "ISIN", "SNAME", "CCY", "T", "CDATE",
            "UNITS", "GROSSPRICE_LOCAL", "BUY_NET_LOCAL", "SELL_NET_LOCAL",
            "INCOME_LOCAL"]
    sub = trades_df[[c for c in cols if c in trades_df.columns]].copy()
    layout = tuple(
        ColSpec(c, 14, "date" if c == "CDATE"
                       else "int" if c == "TRADE_ID"
                       else "float" if c not in ("PCODE_ORIG", "ISIN", "SNAME",
                                                  "CCY", "T")
                       else "text",
                total="blank")
        for c in sub.columns
    )
    n_cols = len(layout)
    for i, spec in enumerate(layout, start=1):
        ws.column_dimensions[get_column_letter(i)].width = spec.width
    header_row = 4
    _write_layout_headers(ws, header_row, layout)
    next_row = header_row + 1
    next_row = _write_layout_data(ws, sub, layout, start_row=next_row)
    ws.freeze_panes = ws.cell(header_row + 1, 1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(n_cols)}{header_row + len(sub)}"


# ─── ValuationA sheet (optional embed) ──────────────────────────────────────

# Map raw HiPort pivot header → friendly display label.
_VALA_HEADER_RENAME: dict[str, str] = {
    "SORT1":      "Region",
    "Initials":   "Analyst",
    "SORT2":      "Country",
    "SNAME":      "Stock Name",
    "FMCID":      "Identifier",
    "BTICKER":    "Bloomberg",
    "Cost":       "Cost",
    "Nominal":    "Nominal",
    "Price":      "Price",
    "Fut Value":  "Future Value",
    "P/F Value":  "P/F Value",
    "Exposure %": "Exposure %",
}

# Width hints for renamed ValuationA columns (after the two inserted P&L cols).
_VALA_COL_WIDTHS: dict[str, int] = {
    "Region":       8,
    "Analyst":      9,
    "Country":     12,
    "Stock Name":  28,
    "Total P&L (EUR)": 15,
    "Total P&L %":     11,
    "Identifier":  14,
    "Bloomberg":   18,
    "Cost":        14,
    "Nominal":     12,
    "Price":       10,
    "Future Value": 14,
    "P/F Value":   14,
    "Exposure %":  11,
    "Last Price (EUR)": 13,
    "Avg Buy Price (EUR)": 14,
    "Avg Sell Price (EUR)": 14,
}


def _write_valuation_a(wb: Workbook, valuation_rows: Sequence[Sequence[object]],
                        meta: PeriodMeta, pl_df: pd.DataFrame | None = None, *,
                        th_val_df: pd.DataFrame | None = None) -> None:
    """Render ValuationA in HiPort's hierarchical layout: grouped by Analyst,
    then by Country, with per-Country and per-Analyst subtotal rows. Stock
    rows whose Initials are ``(blank)`` are reassigned via the SNAME→analyst
    lookup built from the P&L frame, so they sit under the correct Analyst
    and Country headings instead of bunching at the top.
    """
    _ = th_val_df  # reserved: tH_VAL reconciliation column to be added
    ws = wb.create_sheet("ValuationA")
    _setup_sheet(ws, tab_color=COLOR_TAB_ANALYST)
    _write_title_block(ws, meta,
                       title="ValuationA — HiPort Reconciliation",
                       subtitle=f"As of {meta.end_date.date():%d %B %Y}")
    if not valuation_rows or len(valuation_rows) < 2:
        ws.cell(4, 1, "(no rows)").font = FONT_BODY_MUTED
        return

    # ── 1. Find the actual column-header row in the pivot dump ──
    raw_rows = [list(r) for r in valuation_rows]
    hdr_idx = next(
        (i for i, r in enumerate(raw_rows)
         if r and str(r[0] or "").strip().upper() == "SORT1"),
        0,
    )
    raw_headers = raw_rows[hdr_idx]
    raw_body = raw_rows[hdr_idx + 1:]

    def _idx(name: str) -> int | None:
        target = name.upper()
        for i, h in enumerate(raw_headers):
            if str(h or "").strip().upper() == target:
                return i
        return None

    sort1_i = _idx("SORT1")
    init_i = _idx("INITIALS")
    sort2_i = _idx("SORT2")
    sname_i = _idx("SNAME")
    cost_i = _idx("COST")
    nominal_i = _idx("NOMINAL")
    futv_i = _idx("FUT VALUE")
    pfv_i = _idx("P/F VALUE")
    exp_i = _idx("EXPOSURE %")

    # ── 2. SNAME → (analyst, total_pl_eur, total_pl_pct, last_px, buy_px, sell_px) ──
    sname_lookup: dict[str, tuple[str, float | None, float | None,
                                  float | None, float | None, float | None]] = {}
    if pl_df is not None and not pl_df.empty:
        for _, r in pl_df.iterrows():
            key = str(r.get("Stock Name", "")).strip().upper()
            if not key:
                continue
            analyst = str(r.get("Analyst", "") or "").strip().upper()
            pl_eur = r.get("Total P&L (EUR)")
            pl_pct = r.get("Total P&L (%)")
            last_px = r.get("Last Price (EUR)")
            buy_px = r.get("Avg Buy Price (EUR)")
            sell_px = r.get("Avg Sell Price (EUR)")
            sname_lookup[key] = (
                analyst,
                float(pl_eur) if pd.notna(pl_eur) else None,
                float(pl_pct) if pd.notna(pl_pct) else None,
                float(last_px) if pd.notna(last_px) else None,
                float(buy_px) if pd.notna(buy_px) else None,
                float(sell_px) if pd.notna(sell_px) else None,
            )

    # ── 3. Walk rows, propagate SORT1/Initials/SORT2 through merged cells.
    #       Skip subtotal rows (no SNAME). Reassign blank Initials. ──
    cur_sort1 = cur_init = cur_sort2 = ""
    Stock = dict  # type alias for readability
    stocks: list[Stock] = []
    for r in raw_body:
        if not any(v not in (None, "") for v in r):
            continue
        s1 = str(r[sort1_i] or "").strip() if sort1_i is not None and sort1_i < len(r) else ""
        ini = str(r[init_i] or "").strip() if init_i is not None and init_i < len(r) else ""
        s2 = str(r[sort2_i] or "").strip() if sort2_i is not None and sort2_i < len(r) else ""
        sn = str(r[sname_i] or "").strip() if sname_i is not None and sname_i < len(r) else ""
        if s1:
            cur_sort1 = s1
        if ini:
            cur_init = ini
        # SORT2 is reused for "<Country> Total" rows — only adopt as country
        # when it's a real (non-aggregate) value.
        if s2 and "Total" not in s2:
            cur_sort2 = s2
        if not sn:
            # Pure subtotal/header row → drop (we recompute subtotals below).
            continue
        eff_init = cur_init
        if eff_init in ("", "(blank)"):
            mapped = sname_lookup.get(sn.upper())
            if mapped and mapped[0]:
                eff_init = mapped[0]
        pl_eur, pl_pct = (None, None)
        last_px = buy_px = sell_px = None
        hit = sname_lookup.get(sn.upper())
        if hit:
            pl_eur, pl_pct, last_px, buy_px, sell_px = (
                hit[1], hit[2], hit[3], hit[4], hit[5],
            )
        stocks.append({
            "raw": r, "sort1": cur_sort1, "init": eff_init or "(blank)",
            "country": cur_sort2, "sname": sn,
            "pl_eur": pl_eur, "pl_pct": pl_pct,
            "last_px": last_px, "buy_px": buy_px, "sell_px": sell_px,
        })

    # ── 4. Output header row: rename + drop unused HiPort cols + insert
    #       Total P&L cols after SNAME ──
    # User-requested removals: these HiPort columns are dropped because the
    # equivalent information is already conveyed by the appended Last/Avg-Buy/
    # Avg-Sell price columns plus the per-stock Total P&L numbers.
    _VALA_DROPPED = {"Identifier", "Cost", "Nominal", "Price", "P/F Value"}
    display_headers: list[str] = []
    raw_keep: list[int] = []   # raw header indices kept, in display order
    sname_out = None
    for raw_idx, raw_h in enumerate(raw_headers):
        nice = _VALA_HEADER_RENAME.get(str(raw_h or "").strip(),
                                       str(raw_h or "").strip())
        if nice in _VALA_DROPPED:
            continue
        display_headers.append(nice)
        raw_keep.append(raw_idx)
        if nice == "Stock Name":
            sname_out = len(display_headers) - 1
    if sname_out is not None:
        display_headers[sname_out + 1:sname_out + 1] = [
            "Total P&L (EUR)", "Total P&L %",
        ]
    # Append the three per-stock price columns sourced from pl_df.
    display_headers.extend([
        "Last Price (EUR)", "Avg Buy Price (EUR)", "Avg Sell Price (EUR)",
    ])
    n_cols = len(display_headers)

    header_row = 4
    for i, k in enumerate(display_headers, start=1):
        cc = ws.cell(header_row, i, k)
        cc.font = FONT_HEADER
        cc.fill = FILL_HEADER
        cc.alignment = ALIGN_CENTER
        ws.column_dimensions[get_column_letter(i)].width = _VALA_COL_WIDTHS.get(
            k, max(10, min(20, len(k) + 4)),
        )
    ws.row_dimensions[header_row].height = 24

    # ── 5. Sort: assigned analysts first (alpha), unassigned last; then
    #       country, then stock name. ──
    def _stock_key(s):
        unassigned = s["init"] in ("", "(blank)")
        return (1 if unassigned else 0, s["init"], s["country"], s["sname"].upper())
    stocks.sort(key=_stock_key)

    # Map raw column index → output column index. Iterates the kept raw
    # columns (raw_keep) in display order, skipping the synthetic columns
    # injected into display_headers (Total P&L pair + appended price cols).
    raw_to_out: dict[int, int] = {}
    insert_after = sname_out  # 0-based output index of "Stock Name"
    _SYNTH_HEADERS = {
        "Total P&L (EUR)", "Total P&L %",
        "Last Price (EUR)", "Avg Buy Price (EUR)", "Avg Sell Price (EUR)",
    }
    keep_iter = iter(raw_keep)
    for out_idx, h in enumerate(display_headers):
        if h in _SYNTH_HEADERS:
            continue
        raw_to_out[next(keep_iter)] = out_idx
    pl_eur_out = (insert_after + 1) if insert_after is not None else None
    pl_pct_out = (insert_after + 2) if insert_after is not None else None
    last_px_out = n_cols - 3
    buy_px_out = n_cols - 2
    sell_px_out = n_cols - 1

    # Output column indices (0-based) for the propagated/grouped fields.
    out_init = raw_to_out.get(init_i) if init_i is not None else None
    out_sort1 = raw_to_out.get(sort1_i) if sort1_i is not None else None
    out_sort2 = raw_to_out.get(sort2_i) if sort2_i is not None else None
    out_sname = raw_to_out.get(sname_i) if sname_i is not None else None
    out_cost = raw_to_out.get(cost_i) if cost_i is not None else None
    out_nominal = raw_to_out.get(nominal_i) if nominal_i is not None else None
    out_futv = raw_to_out.get(futv_i) if futv_i is not None else None
    out_pfv = raw_to_out.get(pfv_i) if pfv_i is not None else None
    out_exp = raw_to_out.get(exp_i) if exp_i is not None else None

    # ── 6. Emit hierarchical body with subtotals. ──
    FILL_COUNTRY_TOTAL = PatternFill("solid", fgColor="EAEFF1")
    FILL_ANALYST_TOTAL = PatternFill("solid", fgColor="C8D5DA")

    def _format_cell(cell, value, *, is_pct_eur=False, is_pct=False,
                     is_total=False):
        if value is None or value == "":
            cell.value = ""
            cell.alignment = ALIGN_LEFT
        elif isinstance(value, bool):
            cell.value = "Yes" if value else "No"
            cell.alignment = ALIGN_CENTER
        elif isinstance(value, (int, float)):
            cell.value = float(value)
            if is_pct_eur:
                cell.number_format = fmt_for("eur")
            elif is_pct:
                cell.number_format = fmt_for("pct")
            else:
                cell.number_format = "#,##0.00"
            cell.alignment = ALIGN_RIGHT
            if cell.value < 0 and (is_pct_eur or is_pct):
                cell.font = FONT_NEG_BOLD if is_total else FONT_NEG
                return
        elif isinstance(value, datetime):
            cell.value = value
            cell.number_format = "yyyy-mm-dd"
            cell.alignment = ALIGN_CENTER
        else:
            cell.value = str(value)
            cell.alignment = ALIGN_LEFT
        cell.font = FONT_TOTAL if is_total else FONT_BODY

    def _write_subtotal_row(row_idx: int, label: str, label_col: int,
                             stocks_subset, fill, *, bold_label: bool):
        ws.row_dimensions[row_idx].height = 18
        sums = {
            "cost":    sum(_safe_float(s["raw"][cost_i]) for s in stocks_subset)    if cost_i    is not None else 0.0,
            "nominal": sum(_safe_float(s["raw"][nominal_i]) for s in stocks_subset) if nominal_i is not None else 0.0,
            "futv":    sum(_safe_float(s["raw"][futv_i]) for s in stocks_subset)    if futv_i    is not None else 0.0,
            "pfv":     sum(_safe_float(s["raw"][pfv_i]) for s in stocks_subset)     if pfv_i     is not None else 0.0,
            "exp":     sum(_safe_float(s["raw"][exp_i]) for s in stocks_subset)     if exp_i     is not None else 0.0,
            "pl_eur":  sum(s["pl_eur"] for s in stocks_subset if s["pl_eur"] is not None),
        }
        for c_idx in range(n_cols):
            cell = ws.cell(row_idx, c_idx + 1)
            cell.fill = fill
            cell.border = BORDER_TOP_THICK if bold_label else None
            if c_idx == label_col:
                cell.value = label
                cell.font = FONT_TOTAL if bold_label else FONT_BODY
                cell.alignment = ALIGN_LEFT
            elif c_idx == out_cost and out_cost is not None:
                _format_cell(cell, sums["cost"], is_total=True)
            elif c_idx == out_nominal and out_nominal is not None:
                _format_cell(cell, sums["nominal"], is_total=True)
            elif c_idx == out_futv and out_futv is not None:
                _format_cell(cell, sums["futv"], is_total=True)
            elif c_idx == out_pfv and out_pfv is not None:
                _format_cell(cell, sums["pfv"], is_total=True)
            elif c_idx == out_exp and out_exp is not None:
                _format_cell(cell, sums["exp"], is_pct=True, is_total=True)
            elif c_idx == pl_eur_out and pl_eur_out is not None:
                _format_cell(cell, sums["pl_eur"], is_pct_eur=True, is_total=True)
            else:
                cell.value = ""
                cell.font = FONT_TOTAL if bold_label else FONT_BODY
        return row_idx + 1

    out_row = header_row + 1
    prev_init = None
    prev_country_in_init = None
    init_block_start = None     # first stock-row index of current analyst block
    country_block_start = None  # first stock-row index of current country block
    cur_init_stocks: list = []
    cur_country_stocks: list = []

    def _flush_country(at_row: int) -> int:
        nonlocal prev_country_in_init
        if cur_country_stocks and prev_country_in_init is not None:
            label = f"{prev_country_in_init} Total"
            at_row = _write_subtotal_row(
                at_row, label, out_sort2 if out_sort2 is not None else 0,
                cur_country_stocks, FILL_COUNTRY_TOTAL, bold_label=False,
            )
        return at_row

    def _flush_init(at_row: int) -> int:
        nonlocal prev_init
        if cur_init_stocks and prev_init is not None:
            label = f"{prev_init} Total"
            at_row = _write_subtotal_row(
                at_row, label, out_init if out_init is not None else 0,
                cur_init_stocks, FILL_ANALYST_TOTAL, bold_label=True,
            )
        return at_row

    for s in stocks:
        # Detect analyst boundary.
        if s["init"] != prev_init:
            out_row = _flush_country(out_row)
            out_row = _flush_init(out_row)
            prev_init = s["init"]
            prev_country_in_init = None
            cur_init_stocks = []
            cur_country_stocks = []
            init_block_start = out_row
            show_init = True
        else:
            show_init = False

        # Detect country boundary inside the analyst block.
        if s["country"] != prev_country_in_init:
            out_row = _flush_country(out_row)
            prev_country_in_init = s["country"]
            cur_country_stocks = []
            country_block_start = out_row
            show_country = True
        else:
            show_country = False

        # Write the stock row.
        ws.row_dimensions[out_row].height = 16
        for c_idx in range(n_cols):
            cell = ws.cell(out_row, c_idx + 1)
            if c_idx == out_init:
                cell.value = s["init"] if show_init else ""
                cell.alignment = ALIGN_LEFT
                cell.font = FONT_TOTAL if show_init else FONT_BODY
            elif c_idx == out_sort1:
                cell.value = s["sort1"] if show_init else ""
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_BODY
            elif c_idx == out_sort2:
                cell.value = s["country"] if show_country else ""
                cell.alignment = ALIGN_LEFT
                cell.font = FONT_BODY
            elif c_idx == pl_eur_out:
                _format_cell(cell, s["pl_eur"], is_pct_eur=True)
            elif c_idx == pl_pct_out:
                _format_cell(cell, s["pl_pct"], is_pct=True)
            elif c_idx == last_px_out:
                _format_cell(cell, s["last_px"])
            elif c_idx == buy_px_out:
                _format_cell(cell, s["buy_px"])
            elif c_idx == sell_px_out:
                _format_cell(cell, s["sell_px"])
            else:
                # Pull from raw row by reverse-mapping output → raw index.
                raw_idx = next((ri for ri, oi in raw_to_out.items() if oi == c_idx), None)
                v = s["raw"][raw_idx] if raw_idx is not None and raw_idx < len(s["raw"]) else None
                if c_idx == out_exp:
                    _format_cell(cell, v, is_pct=True)
                else:
                    _format_cell(cell, v)
        cur_country_stocks.append(s)
        cur_init_stocks.append(s)
        out_row += 1

    # Flush the final blocks.
    out_row = _flush_country(out_row)
    out_row = _flush_init(out_row)

    # Grand Total — sum across every stock (matches HiPort's pivot footer).
    if stocks:
        FILL_GRAND_TOTAL = PatternFill("solid", fgColor="2E6168")
        ws.row_dimensions[out_row].height = 22
        sums = {
            "cost":    sum(_safe_float(s["raw"][cost_i]) for s in stocks)    if cost_i    is not None else 0.0,
            "nominal": sum(_safe_float(s["raw"][nominal_i]) for s in stocks) if nominal_i is not None else 0.0,
            "futv":    sum(_safe_float(s["raw"][futv_i]) for s in stocks)    if futv_i    is not None else 0.0,
            "pfv":     sum(_safe_float(s["raw"][pfv_i]) for s in stocks)     if pfv_i     is not None else 0.0,
            "exp":     sum(_safe_float(s["raw"][exp_i]) for s in stocks)     if exp_i     is not None else 0.0,
            "pl_eur":  sum(s["pl_eur"] for s in stocks if s["pl_eur"] is not None),
        }
        WHITE_BOLD = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        for c_idx in range(n_cols):
            cell = ws.cell(out_row, c_idx + 1)
            cell.fill = FILL_GRAND_TOTAL
            cell.border = BORDER_TOP_THICK
            cell.font = WHITE_BOLD
            if c_idx == 0:
                cell.value = "Grand Total"
                cell.alignment = ALIGN_LEFT
            elif c_idx == out_cost and out_cost is not None:
                cell.value = sums["cost"]; cell.number_format = "#,##0.00"; cell.alignment = ALIGN_RIGHT
            elif c_idx == out_nominal and out_nominal is not None:
                cell.value = sums["nominal"]; cell.number_format = "#,##0.00"; cell.alignment = ALIGN_RIGHT
            elif c_idx == out_futv and out_futv is not None:
                cell.value = sums["futv"]; cell.number_format = "#,##0.00"; cell.alignment = ALIGN_RIGHT
            elif c_idx == out_pfv and out_pfv is not None:
                cell.value = sums["pfv"]; cell.number_format = "#,##0.00"; cell.alignment = ALIGN_RIGHT
            elif c_idx == out_exp and out_exp is not None:
                cell.value = sums["exp"]; cell.number_format = fmt_for("pct"); cell.alignment = ALIGN_RIGHT
            elif c_idx == pl_eur_out and pl_eur_out is not None:
                cell.value = sums["pl_eur"]; cell.number_format = fmt_for("eur"); cell.alignment = ALIGN_RIGHT
            else:
                cell.value = ""
        out_row += 1

    ws.freeze_panes = ws.cell(header_row + 1, 1)


# ─── Validation report & methodology ────────────────────────────────────────

def _write_validation(
    wb: Workbook, validation: ValidationResult, meta: PeriodMeta,
    write_seconds: float,
    *, ca_suggestions_path: Path | None = None,
) -> None:
    ws = wb.create_sheet("Validation Report")
    _setup_sheet(ws, tab_color=COLOR_TAB_AUDIT)
    _write_title_block(ws, meta,
                       title="Validation Report",
                       subtitle=f"Workbook write took {write_seconds:.2f}s")

    if validation.has_critical:
        head = "STATUS: CRITICAL FAILURES PRESENT"
        col = COLOR_VALIDATION_FAIL
    elif validation.has_warnings:
        head = "STATUS: PASS WITH WARNINGS"
        col = COLOR_VALIDATION_WARN
    else:
        head = "STATUS: ALL CHECKS PASS"
        col = COLOR_VALIDATION_PASS
    c = ws.cell(3, 1, head); c.font = Font(bold=True, size=12, color=col)

    headers = ("ID", "Check", "Status", "Detail", "Failures")
    header_row = 5
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(header_row, i, h)
        cell.font = FONT_HEADER; cell.fill = FILL_HEADER; cell.alignment = ALIGN_CENTER

    r = header_row + 1
    for chk in validation.checks:
        ws.cell(r, 1, chk.id).font = FONT_BODY
        ws.cell(r, 2, chk.name).font = FONT_BODY
        sc = ws.cell(r, 3, chk.status)
        sc.font = Font(
            bold=True,
            color={"PASS": COLOR_VALIDATION_PASS,
                   "WARN": COLOR_VALIDATION_WARN,
                   "FAIL": COLOR_VALIDATION_FAIL}[chk.status],
        )
        ws.cell(r, 4, chk.detail).font = FONT_BODY
        ws.cell(r, 5, len(chk.failures)).font = FONT_BODY
        r += 1

    # Per-check failure tables. NEVER truncated.
    r += 2
    for chk in validation.checks:
        if not chk.failures:
            continue
        ws.cell(r, 1, f"{chk.id} — {chk.name}: {len(chk.failures)} failure(s)").font = FONT_SUBTITLE
        r += 1
        keys = list(chk.failures[0].keys())
        for i, k in enumerate(keys, start=1):
            cc = ws.cell(r, i, k)
            cc.font = FONT_HEADER; cc.fill = FILL_HEADER; cc.alignment = ALIGN_CENTER
        r += 1
        for f in chk.failures:
            for i, k in enumerate(keys, start=1):
                v = f.get(k, "")
                cell = ws.cell(r, i, v); cell.font = FONT_BODY
                if isinstance(v, float):
                    cell.number_format = fmt_for("float")
            r += 1
        r += 2

    # ── How-to-fix block for VAL-01 ──────────────────────────────────────────
    try:
        val01 = validation.by_id("VAL-01")
    except KeyError:
        val01 = None
    if val01 is not None and val01.status == "FAIL" and val01.failures:
        cell = ws.cell(r, 1, "VAL-01 — How to unblock")
        cell.font = FONT_SUBTITLE
        r += 1
        msg_lines = [
            "Each VAL-01 failure means HiPort end-units do NOT match start + bought - sold.",
            "Causes are almost always corporate actions (bonus, split, consolidation, in-kind delivery)",
            "that hit the snapshot but not the Bottler trade feed.",
            "",
            "TWO WAYS TO FIX:",
            "  1. Permanent: paste the rows below into inputs/ca_overrides.csv",
            "  2. One-shot : save them to a CSV and re-run with",
            "                python -m oefof_pl --apply-suggestions <file>.csv",
        ]
        if ca_suggestions_path is not None:
            msg_lines.append("")
            msg_lines.append(
                f"Auto-generated file ready to use: {ca_suggestions_path}"
            )
            msg_lines.append(
                f"  Quick re-run: python -m oefof_pl --apply-suggestions \"{ca_suggestions_path}\""
            )
        for line in msg_lines:
            ws.cell(r, 1, line).font = FONT_BODY
            r += 1
        r += 1
        # Inline preview of the suggestion rows.
        sugg_headers = ("ISIN", "UNITS (signed Diff)", "Suggested REASON (review!)")
        for i, h in enumerate(sugg_headers, start=1):
            cc = ws.cell(r, i, h)
            cc.font = FONT_HEADER; cc.fill = FILL_HEADER; cc.alignment = ALIGN_CENTER
        r += 1
        for f in val01.failures:
            isin = str(f.get("ISIN", ""))
            diff = f.get("Diff", 0.0)
            sname = str(f.get("Stock Name", "")).strip()
            ws.cell(r, 1, isin).font = FONT_BODY
            uc = ws.cell(r, 2, float(diff)); uc.font = FONT_BODY
            uc.number_format = fmt_for("float")
            ws.cell(r, 3, f"{sname} - TODO confirm with ops").font = FONT_BODY
            r += 1
        r += 2

    for i, w in enumerate((10, 28, 8, 60, 10), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_methodology(wb: Workbook, meta: PeriodMeta, *,
                       applied_ca_overrides: tuple = (),
                       ca_override_sources: tuple = (),
                       analyst_sname_overrides: dict | None = None) -> None:
    ws = wb.create_sheet("Methodology")
    _setup_sheet(ws, tab_color=COLOR_TAB_AUDIT)
    _write_title_block(ws, meta,
                       title="Methodology & Audit Trail",
                       subtitle="Read this once. Then read the validation report.")

    text = [
        ("ORD positions", "bold"),
        ("  Total P&L (EUR) = (End MV − Start MV) + Sell Proceeds − Buy Cost + Income", "body"),
        ("  Cost basis uses Weighted Average Cost (WAC). Sells reduce held units at WAC; "
         "the WAC itself is not changed by sells.", "body"),
        ("  Realised + Unrealised + Income does NOT equal Total. Unrealised is reported "
         "point-in-time (End MV − End WAC Cost), not the period change in unrealised. "
         "The cash-flow Total above is the authoritative period P&L.", "body"),
        ("  Cross-zero sells (long → short in one trade) book the entire realised P&L at "
         "the prior WAC and open the residual short at a negative cost basis. VAL-06 "
         "will WARN; review those rows manually if any appear.", "body"),
        ("", "body"),
        ("SWAP / FTSWAP positions", "bold"),
        ("  Realised  = 0", "body"),
        ("  Unrealised = 0", "body"),
        ("  Total P&L (EUR) = Income (EUR)", "body"),
        ("  Cost Basis (EUR) = |Notional × Units| converted at end FX", "body"),
        ("", "body"),
        ("FX policy", "bold"),
        ("  All monetary components (start MV, end MV, buy, sell, income) are converted to EUR "
         "at the END-period FX rate. This eliminates phantom corporate-action gains that arose "
         "in the legacy system from per-trade USD conversion (BUG-4).", "body"),
        ("", "body"),
        ("Closed positions", "bold"),
        ("  When a position is fully exited inside the period the end-period FX may be missing for "
         "its currency. We then fall back to the start-period FX rate so the realised P&L is still "
         "computed in EUR. (Regression covered by tests/test_pl.py.)", "body"),
        ("", "body"),
        ("Corporate-action trades (BCODE = ZZZZ)", "bold"),
        ("  HiPort marks SCODE migrations and bonus-share events with BCODE='ZZZZ' and price=0. "
         "These are NOT economic transactions; we drop them before WAC. Failing to drop them "
         "would dilute per-share cost to zero and book phantom losses on the matched sell.", "body"),
        ("", "body"),
        ("Manual CA overrides (BCODE = CA_ADJ)", "bold"),
        ("  When HiPort's snapshot reflects a corporate action (bonus, split, consolidation, "
         "in-kind delivery) that did NOT arrive in the Bottler trade feed, the unit drift is "
         "patched via inputs/ca_overrides.csv. Each row is injected as a synthetic T='P' trade "
         "at price 0 dated end-period. WAC handles bonus (units up, cost unchanged) and "
         "consolidation (units down, cost unchanged) correctly. Audit trail = the CSV itself; "
         "synthetic rows carry negative TRADE_IDs and BCODE='CA_ADJ' in Trade Detail.", "body"),
        ("", "body"),
        ("Income column = Income & Financing", "bold"),
        ("  HiPort tags every cash-flow leg with T='I' regardless of whether it is a real "
         "dividend or a CFD/swap financing/borrow charge. Negative values in the Income column "
         "are therefore NOT 'negative dividends' \u2014 they are financing legs on CFD/swap "
         "positions (Goldman/HSBC swaps, MSCI ETF CFDs, etc.). The column is labelled "
         "'Income & Financing (EUR)' to make this explicit. Net of the two flows is the cash "
         "P&L impact regardless of category.", "body"),
        ("", "body"),
        ("Percentage handling", "bold"),
        ("  All percent columns are stored as ratios (0.44 = 44%) and Excel's '0.00%' format does "
         "the display conversion. The validate stage rejects anything outside ±5 (the legacy "
         "×100 bug, BUG-1).", "body"),
        ("", "body"),
        ("Inputs", "bold"),
        (f"  HiPort snapshot: {meta.snapshot_path}", "body"),
        (f"  Bottler trades : {meta.bottler_path}", "body"),
        (f"  Generated      : {meta.run_timestamp:%Y-%m-%d %H:%M:%S}", "body"),
        *([
            ("", "body"),
            ("CA Override Sources", "bold"),
            *[(f"  {p}", "body") for p in ca_override_sources],
            (f"  Applied overrides: {len(applied_ca_overrides)} row(s)", "body"),
        ] if ca_override_sources else []),
        *([
            ("", "body"),
            ("Analyst Name Overrides", "bold"),
            *[(f"  {k} → {v}", "body") for k, v in (analyst_sname_overrides or {}).items()],
        ] if analyst_sname_overrides else []),
    ]
    r = 4
    for line, kind in text:
        cell = ws.cell(r, 1, line)
        if kind == "bold":
            cell.font = FONT_SUBTITLE
        else:
            cell.font = FONT_BODY
        cell.alignment = ALIGN_LEFT
        r += 1
    ws.column_dimensions["A"].width = 130


# ─── Coercion / sheet ordering / atomic save ────────────────────────────────

def _coerce_value(v, kind: str):
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if kind == "date":
        if isinstance(v, pd.Timestamp):
            if pd.isna(v):
                return None
            return v.to_pydatetime()
        return v
    if kind in ("int", "float", "price", "eur", "pct"):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None
    return v if isinstance(v, str) else str(v)


def _safe_sheet_name(name: str) -> str:
    """Excel sheet names: max 31 chars; cannot contain : \\ / ? * [ ]."""
    bad = ':\\/?*[]'
    cleaned = "".join("_" if ch in bad else ch for ch in str(name))
    return cleaned[:31] or "Sheet"


def _enforce_sheet_order(wb: Workbook) -> None:
    """Order: Summary, ValuationA?, <analysts...>,
    Current Holdings, Exited Positions, Income & Dividends, Income Detail,
    Trade Detail, Validation Report, Methodology.
    """
    desired: list[str] = []
    desired.append("Summary")
    if "ValuationA" in wb.sheetnames:
        desired.append("ValuationA")
    # Analyst sheets in display order
    known_static = set(SHEET_ORDER) | {"ValuationA"}
    analyst_sheets = [s for s in wb.sheetnames if s not in known_static]
    # Order analyst sheets by ANALYST_TAB_ORDER display name then alphabetical
    display_order = [_analyst_display(c) for c in ANALYST_TAB_ORDER]
    analyst_sheets.sort(key=lambda s: (
        display_order.index(s) if s in display_order else len(display_order),
        s,
    ))
    desired.extend(analyst_sheets)
    desired.extend([s for s in SHEET_ORDER if s != "Summary"])

    # Drop anything not actually in workbook (defensive)
    desired = [s for s in desired if s in wb.sheetnames]
    by_name = {ws.title: ws for ws in wb.worksheets}
    wb._sheets = [by_name[name] for name in desired]


def _autofit_all_sheets(wb: Workbook) -> None:
    """Widen any column whose existing width is too narrow to display its
    widest *numeric* cell, so €-amount totals don't render as ######.

    Title rows and text-only columns are deliberately excluded — text is
    allowed to overflow visually, and many text columns (ISIN, CCY, L/S)
    have intentionally narrow widths that we don't want to expand.
    """
    WIDTH_CAP = 22
    for ws in wb.worksheets:
        col_max: dict[int, int] = {}
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                # Skip the title block (rows 1-2): long free-text labels
                # in column A that are allowed to overflow visually.
                if cell.row <= 2:
                    continue
                v = cell.value
                if v is None or not isinstance(v, (int, float)):
                    # Text & blanks ignored — we only widen for numeric
                    # ###### avoidance.
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
    # Auto-widen any column whose configured width is too narrow for its
    # widest cell. Prevents ###### display for €-formatted totals and
    # long header labels (e.g. "Total P&L % (on Gross Exposure)").
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
