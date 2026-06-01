"""Workbook styling constants and the ``ColSpec`` schema that drives the
output writer.

Adding a column to ``Current Holdings`` is a one-line change: append a
``ColSpec`` to ``COLUMN_LAYOUT_HOLDINGS``. Header writing, data writing,
and the total row all iterate the same list. The legacy code required
edits in 4 separate places per column — the source of BUG-2.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.styles import Alignment, Font, PatternFill


# ─── Colours (hex without leading #) ────────────────────────────────────────
# Brand palette is built around Fiera teal (#1B474D). All accent colours
# pick from the same hue family so the workbook reads as one design.

COLOR_HEADER = "1B474D"
COLOR_HEADER_ACCENT = "2E7C84"      # mid-teal for sub-section headers
COLOR_TITLE_TEXT = "0F2F33"
COLOR_SUBTITLE_TEXT = "1B474D"
COLOR_MUTED = "7A8B8D"

COLOR_POS = "1F8A4C"               # green for gains
COLOR_NEG = "B83A3A"               # red for losses
COLOR_TOTAL_FILL = "E8F4F5"
COLOR_SUBTOTAL_FILL = "DDEBED"
COLOR_BAND_ALT = "F4F8F9"
COLOR_KPI_FILL = "F8FBFB"
COLOR_KPI_FILL_ACCENT = "EDF5F6"

COLOR_VALIDATION_PASS = "1F8A4C"
COLOR_VALIDATION_WARN = "B58900"
COLOR_VALIDATION_FAIL = "B83A3A"

COLOR_TAB_DASHBOARD = "1B474D"
COLOR_TAB_ANALYST = "2E7C84"
COLOR_TAB_DATA = "7A8B8D"
COLOR_TAB_AUDIT = "9D7B00"


# ─── Fonts / fills ──────────────────────────────────────────────────────────

FONT_BODY = Font(name="Calibri", size=10)
FONT_BODY_MUTED = Font(name="Calibri", size=10, color=COLOR_MUTED)
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_HEADER_ACCENT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_TITLE = Font(name="Calibri", size=20, bold=True, color=COLOR_TITLE_TEXT)
FONT_SUBTITLE = Font(name="Calibri", size=12, bold=True, color=COLOR_SUBTITLE_TEXT)
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_TOTAL = Font(name="Calibri", size=10, bold=True)
FONT_NEG = Font(name="Calibri", size=10, color=COLOR_NEG)
FONT_NEG_BOLD = Font(name="Calibri", size=10, bold=True, color=COLOR_NEG)
FONT_POS = Font(name="Calibri", size=10)
FONT_KPI_LABEL = Font(name="Calibri", size=9, bold=True, color=COLOR_MUTED)
FONT_KPI_VALUE = Font(name="Calibri", size=14, bold=True, color=COLOR_TITLE_TEXT)
FONT_KPI_VALUE_POS = Font(name="Calibri", size=14, bold=True, color=COLOR_TITLE_TEXT)
FONT_KPI_VALUE_NEG = Font(name="Calibri", size=14, bold=True, color=COLOR_NEG)

FILL_HEADER = PatternFill("solid", fgColor=COLOR_HEADER)
FILL_HEADER_ACCENT = PatternFill("solid", fgColor=COLOR_HEADER_ACCENT)
FILL_TOTAL = PatternFill("solid", fgColor=COLOR_TOTAL_FILL)
FILL_SUBTOTAL = PatternFill("solid", fgColor=COLOR_SUBTOTAL_FILL)
FILL_BAND = PatternFill("solid", fgColor=COLOR_BAND_ALT)
FILL_KPI = PatternFill("solid", fgColor=COLOR_KPI_FILL)
FILL_KPI_ACCENT = PatternFill("solid", fgColor=COLOR_KPI_FILL_ACCENT)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_KPI_LABEL = Alignment(horizontal="left", vertical="center", indent=1)
ALIGN_KPI_VALUE = Alignment(horizontal="left", vertical="center", indent=1)


# ─── Borders ────────────────────────────────────────────────────────────────
# openpyxl: import lazily inside the workbook module to avoid bloating
# this module's import surface.


# ─── Number formats ─────────────────────────────────────────────────────────

FMT_TEXT = "@"
FMT_INT = "#,##0;[Red](#,##0);"
FMT_FLOAT = "#,##0.00;[Red](#,##0.00);"
FMT_PRICE = "#,##0.0000;[Red](#,##0.0000);"
FMT_EUR = '"€"#,##0;[Red]("€"#,##0);'  # 3 sections: positive; negative; zero
                                        # The empty trailing section renders
                                        # exact zero as a blank cell instead
                                        # of "€0" — keeps analyst sheets
                                        # readable when whole columns are
                                        # zero (e.g. Income for SB / KX / HK).
FMT_PCT = "0.00%;[Red](0.00%)"   # red negative in parens; stored ratio (0.44) → 44.00%
FMT_DATE = "dd-mm-yyyy"


# ─── ColSpec ────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class ColSpec:
    """One column's full presentation contract.

    Attributes:
        df_column: Source DataFrame column (and the displayed header).
        width: Approximate Excel column width.
        kind: ``'text' | 'int' | 'float' | 'price' | 'eur' | 'pct' | 'date'``.
        total: How the total row is computed for this column.
            * ``'sum'``   — arithmetic sum
            * ``'wavg'``  — weighted average (rare, not used here)
            * ``'blank'`` — leave empty
            * ``'label'`` — write ``'Total'`` (used on the leftmost col)
        display: Optional override for the displayed header text. When None,
            ``df_column`` is used. Use this when the source column name is
            stable but the human-facing label should differ (e.g. an income
            column that actually mixes dividends + CFD financing legs).
    """
    df_column: str
    width: int
    kind: str
    total: str = "blank"
    display: str | None = None


_KIND_TO_FMT = {
    "text": FMT_TEXT,
    "int": FMT_INT,
    "float": FMT_FLOAT,
    "price": FMT_PRICE,
    "eur": FMT_EUR,
    "pct": FMT_PCT,
    "date": FMT_DATE,
}


def fmt_for(kind: str) -> str:
    if kind not in _KIND_TO_FMT:
        raise ValueError(f"Unknown ColSpec kind {kind!r}")
    return _KIND_TO_FMT[kind]


def align_for(kind: str) -> Alignment:
    if kind == "text":
        return ALIGN_LEFT
    if kind == "date":
        return ALIGN_CENTER
    return ALIGN_RIGHT


# ─── Layouts ────────────────────────────────────────────────────────────────
# ``df_column`` strings MUST match ``data.normalise.PL_COLUMNS`` exactly.
# A test asserts column-name parity — if PL_COLUMNS changes and a layout
# entry doesn't, the build fails loudly.

COLUMN_LAYOUT_HOLDINGS: tuple[ColSpec, ...] = (
    ColSpec("ISIN",                       14, "text",  total="label"),
    ColSpec("Stock Name",                 32, "text"),
    ColSpec("Country",                     8, "text"),
    ColSpec("Analyst",                    10, "text"),
    ColSpec("CCY",                         6, "text"),
    ColSpec("Instrument",                  9, "text"),
    ColSpec("L/S",                         5, "text"),
    ColSpec("Starting Units",             14, "float"),
    ColSpec("Ending Units",               14, "float"),
    ColSpec("Units Bought",               14, "float"),
    ColSpec("Units Sold",                 14, "float"),
    ColSpec("Avg Buy Price (EUR)",        16, "price"),
    ColSpec("Avg Sell Price (EUR)",       16, "price"),
    ColSpec("Start Price (Local)",        14, "price"),
    ColSpec("Last Price (EUR)",           14, "price"),
    ColSpec("Market Value Start (EUR)",   18, "eur",  total="sum"),
    ColSpec("Realised P&L (Local)",       16, "float"),
    ColSpec("Realised P&L (EUR)",         16, "eur",  total="sum"),
    ColSpec("Realised P&L (%)",           14, "pct"),
    ColSpec("Unrealised P&L (Local)",     16, "float"),
    ColSpec("Unrealised P&L (EUR)",       16, "eur",  total="sum"),
    ColSpec("Unrealised P&L (%)",         16, "pct"),
    ColSpec("Income (Local)",             14, "float"),
    ColSpec("Income (EUR)",               14, "eur",  total="sum"),
    ColSpec("Dividends (EUR)",            14, "eur",  total="sum"),
    ColSpec("Swap Financing (EUR)",       16, "eur",  total="sum"),
    ColSpec("Income Yield (%)",           14, "pct"),
    ColSpec("Total P&L (EUR)",            14, "eur",  total="sum"),
    ColSpec("Contribution to Total (%)",  14, "pct",
            display="Contribution %"),
    ColSpec("Total P&L (%)",              12, "pct"),
    ColSpec("Last Trade Date",            14, "date"),
    ColSpec("First Trade Date",           14, "date"),
)

COLUMN_LAYOUT_INCOME: tuple[ColSpec, ...] = (
    ColSpec("ISIN",          14, "text",  total="label"),
    ColSpec("Stock Name",    32, "text"),
    ColSpec("Analyst",       10, "text"),
    ColSpec("CCY",            6, "text"),
    ColSpec("Date",          14, "date"),
    ColSpec("Income Type",   14, "text"),
    ColSpec("Income (Local)", 14, "float"),
    ColSpec("Income (EUR)",   14, "eur",  total="sum",
            display="Income & Financing (EUR)"),
    ColSpec("Units",         14, "float"),
)


# Per-analyst & exited layout: friendly, narrow layout with the headline
# P&L columns up front. ISIN kept narrow (copy-paste only). Used by per-
# analyst sheets and the "Exited Positions" sheet for parity.
COLUMN_LAYOUT_ANALYST: tuple[ColSpec, ...] = (
    ColSpec("ISIN",                       12, "text",  total="blank"),  # hidden; label on Stock Name
    ColSpec("Stock Name",                 26, "text",  total="label"),
    ColSpec("Total P&L (EUR)",            15, "eur",  total="sum"),
    ColSpec("Total P&L (%)",              10, "pct",
            display="Total P&L %"),
        ColSpec("Contribution to Total (%)",  14, "int",
            display="Contribution (bps)"),
    ColSpec("Ending Units",               11, "int",
            display="Units"),
    ColSpec("Last Price (EUR)",           11, "price",
            display="Last Px (EUR)"),
    ColSpec("Avg Buy Price (EUR)",        14, "price",
            display="Avg Buy Px (EUR)"),
    ColSpec("Avg Sell Price (EUR)",       14, "price",
            display="Avg Sell Px (EUR)"),
    ColSpec("Realised P&L (EUR)",         14, "eur",  total="sum",
            display="Realised (EUR)"),
    ColSpec("Unrealised P&L (EUR)",       15, "eur",  total="sum",
            display="Unrealised (EUR)"),
    ColSpec("Income (EUR)",               20, "eur",  total="sum",
            display="Income & Financing (EUR)"),
    ColSpec("Dividends (EUR)",            14, "eur",  total="sum",
            display="Dividends (EUR)"),
    ColSpec("Swap Financing (EUR)",       16, "eur",  total="sum",
            display="Swap Financing (EUR)"),
    ColSpec("CCY",                         5, "text"),
    ColSpec("Instrument",                  9, "text",
            display="Type"),
    ColSpec("L/S",                         4, "text"),
)
