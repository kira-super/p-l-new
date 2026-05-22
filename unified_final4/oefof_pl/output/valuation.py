"""ValuationA workbook builder.

Two responsibilities, sharply separated:

1. **Pure post-processing** — given a list of `ValuationA` rows (cols A:L)
   and a per-key lookup of the 8 P&L columns (M:T) computed by the new
   compute layer, produce a 2+N-sheet workbook
   (``Summary``, ``ValuationA``, one tab per analyst). Pure ``openpyxl``,
   no COM, fully unit-tested.
2. **COM extraction** — refresh ``HP_VAL.xlsm!ValuationA`` with the OAKS
   pivot page filter and extract cols A:L. Gated behind a function-local
   ``import win32com`` so the package imports cleanly on non-Windows CI
   and the AST lint at Phase 13 stays green.

Public entry points
-------------------
``run_valuation(...)``
    End-to-end orchestration used by the pipeline (COM + write).

``build_valuation_workbook(...)``
    Pure write-only entry used by tests and any caller that already has
    val-rows + a P&L lookup in memory.

``extract_valuation_a_via_com(...)``
    The COM-only step, exposed so tests can monkey-patch it.

``close_open_excel_workbook(...)``
    Best-effort file-lock workaround; mirrors the legacy ROT walk.
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, Mapping, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

from ..config import PERSON_OVERRIDES_BY_SECURITY, TARGET_PNAME


# ─── Constants ───────────────────────────────────────────────────────────────

VAL_ACOL_COUNT = 12          # ValuationA cols A:L extracted from HP_VAL
VAL_MT_COUNT = 8             # Cols M:T appended from PL lookup
BLANK_PERSON_KEY = "(blank)"

# ── ValuationA row structure (0-based indices into row tuple) ───────────────
HPV_SORT1_IDX = 0   # col A
HPV_INITIALS_IDX = 1   # col B  (analyst initials, carried down)
HPV_SECURITY_NAME_IDX = 3   # col D
HPV_FMCID_IDX = 4   # col E  (lookup key into pl_lookup)
HPV_FUTVAL_IDX = 9  # col J  (Fut Value — sign decides long/short)

DATA_FIRST_ROW = 8   # First data row in ValuationA after pivot headers
HEADER_ROW_M_TO_T = 7   # Row that carries the M:T header labels

# ── Sheet ordering ───────────────────────────────────────────────────────────
SHEET_SUMMARY = "Summary"
SHEET_VALUATION_A = "ValuationA"

# ── Styles ───────────────────────────────────────────────────────────────────
COLOR_HEADER = "1B474D"
COLOR_TOTAL_FILL = "E8F4F5"

HEADER_FILL = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type="solid")
TOTAL_FILL = PatternFill(start_color=COLOR_TOTAL_FILL, end_color=COLOR_TOTAL_FILL, fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name="Calibri", bold=True, size=16, color=COLOR_HEADER)
SUBTITLE_FONT = Font(name="Calibri", size=11, color=COLOR_HEADER)
BODY_FONT = Font(name="Calibri", size=10)
TOTAL_FONT = Font(name="Calibri", bold=True, size=10)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)

FMT_AMOUNT = '#,##0.00;[Red](#,##0.00)'
FMT_PCT = '0.00%;[Red](0.00%)'
FMT_PCT1 = '0.0%;[Red](0.0%)'

# ── M:T headers (filled from pl_lookup) ──────────────────────────────────────
VAL_HEADERS_M_TO_T: tuple[str, ...] = (
    "Realised P&L (EUR)",
    "Realised P&L (%)",
    "Unrealised P&L (EUR)",
    "Unrealised P&L (%)",
    "Income (EUR)",
    "Income Yield (%)",
    "Total P&L (EUR)",
    "Total P&L (%)",
)

# Number format per appended column. EUR amounts vs. percentages alternate.
VAL_FORMATS_M_TO_T: tuple[str, ...] = (
    FMT_AMOUNT, FMT_PCT, FMT_AMOUNT, FMT_PCT,
    FMT_AMOUNT, FMT_PCT, FMT_AMOUNT, FMT_PCT,
)

# ── Detail sheet column layout (one row per analyst's holding) ───────────────
# `source_col` is a 1-based index into the M:T-extended ValuationA row, so
# 1..12 = A:L from COM, 13..20 = M:T from pl_lookup, None = spacer.

@dataclass(frozen=True)
class _DetailCol:
    header: str
    source_col: int | None
    width: int
    align: Alignment
    number_format: str | None = None


DETAIL_COLUMN_LAYOUT: tuple[_DetailCol, ...] = (
    _DetailCol("SORT2",                3,  12, ALIGN_LEFT),
    _DetailCol("SNAME",                4,  24, ALIGN_LEFT),
    _DetailCol("FMCID",                5,  12, ALIGN_LEFT),
    _DetailCol("BTICKER",              6,  12, ALIGN_LEFT),
    _DetailCol("Total P&L (EUR)",      19, 16, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("Total P&L (%)",        20, 13, ALIGN_RIGHT, FMT_PCT),
    _DetailCol("",                     None, 4, ALIGN_LEFT),
    _DetailCol("",                     None, 4, ALIGN_LEFT),
    _DetailCol("Cost",                 7,  13, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("Nominal",              8,  13, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("Price",                9,  11, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("Fut Value",            10, 14, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("P/F Value",            11, 14, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("Exposure %",           12, 12, ALIGN_RIGHT, FMT_PCT),
    _DetailCol("Realised P&L (EUR)",   13, 16, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("Realised P&L (%)",     14, 13, ALIGN_RIGHT, FMT_PCT),
    _DetailCol("Unrealised P&L (EUR)", 15, 17, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("Unrealised P&L (%)",   16, 14, ALIGN_RIGHT, FMT_PCT),
    _DetailCol("Income (EUR)",         17, 13, ALIGN_RIGHT, FMT_AMOUNT),
    _DetailCol("Income Yield (%)",     18, 14, ALIGN_RIGHT, FMT_PCT),
)


# ─── Frozen records ──────────────────────────────────────────────────────────

@dataclass(frozen=True)
class StockRow:
    """A single stock-level row from ValuationA (Sort1 == 'A')."""
    source_row: int
    person: str          # post-override analyst code
    raw_person: str      # original initials before override
    j_val: float
    t_val: float


@dataclass(frozen=True)
class SummaryAcc:
    long_j: float = 0.0
    short_j: float = 0.0
    long_t: float = 0.0
    short_t: float = 0.0


# ─── Pure helpers ────────────────────────────────────────────────────────────

def normalize_person(value: object) -> str:
    """Map empty/missing initials to the shared blank bucket."""
    if value in (None, ""):
        return BLANK_PERSON_KEY
    text = str(value).strip()
    return text if text else BLANK_PERSON_KEY


def resolve_person_override(
    security_name: object,
    default_person: object,
    sname_overrides: Mapping[str, str] | None = None,
) -> str:
    """Apply manual analyst remaps for known SNAME exceptions.

    ``sname_overrides`` is consulted first; if not supplied, falls back to
    the (now typically empty) module-level ``PERSON_OVERRIDES_BY_SECURITY``
    so unit tests that don't construct an ``AnalystMap`` keep working.
    """
    if security_name in (None, ""):
        return normalize_person(default_person)
    key = str(security_name).strip().upper()
    overrides = sname_overrides if sname_overrides is not None else PERSON_OVERRIDES_BY_SECURITY
    return overrides.get(key, normalize_person(default_person))


def ordered_people(stock_rows: Sequence[StockRow]) -> list[str]:
    """Return analyst codes in first-seen order, with the blank bucket last."""
    seen: list[str] = []
    for row in stock_rows:
        if row.person != BLANK_PERSON_KEY and row.person not in seen:
            seen.append(row.person)
    if any(row.person == BLANK_PERSON_KEY for row in stock_rows):
        seen.append(BLANK_PERSON_KEY)
    return seen


def detail_sheet_name(person: str) -> str:
    """Excel-safe worksheet title for an analyst detail tab."""
    return "Unassigned" if person == BLANK_PERSON_KEY else str(person)[:31]


def collect_stock_rows(
    val_rows: Sequence[Sequence[object]],
    analyst_lookup: Mapping[str, str] | None = None,
    sname_overrides: Mapping[str, str] | None = None,
) -> list[StockRow]:
    """Walk ValuationA rows, carry SORT1 and INITIALS forward, keep Sort1=='A'.

    Mirrors the legacy carry-down semantics:
      - Col A (SORT1) sets the active section; resets person on any 'Total'.
      - Col B (INITIALS) carries forward until a new value appears.
      - Only rows with non-blank FMCID, numeric Fut Value, and SORT1=='A'
        are emitted as `StockRow`. Section header / subtotal rows are skipped.

    If ``analyst_lookup`` (FMCID/ISIN -> analyst code) is provided, it is
    used as a fallback whenever HiPort's pivot has no carried-down initials
    for the row. This prevents otherwise-mapped positions from landing in
    the ``Unassigned`` tab simply because HiPort hasn't tagged them yet.
    """
    stock_rows: list[StockRow] = []
    current_sort1: str | None = None
    current_person: object = None

    # Data starts at row index DATA_FIRST_ROW-1 (0-based) within val_rows.
    for idx, row in enumerate(val_rows):
        source_row_1based = idx + 1   # Excel row number for downstream lookups
        if source_row_1based < DATA_FIRST_ROW:
            continue
        row_padded = list(row) + [None] * (max(0, HPV_FUTVAL_IDX + 1 - len(row)))

        sort1 = row_padded[HPV_SORT1_IDX]
        if sort1 not in (None, ""):
            current_sort1 = str(sort1).strip()
        if sort1 is not None and "Total" in str(sort1):
            current_person = None

        initials = row_padded[HPV_INITIALS_IDX]
        if initials not in (None, ""):
            current_person = initials

        fmcid = row_padded[HPV_FMCID_IDX]
        if fmcid in (None, "") or current_sort1 != "A":
            continue

        j_val = row_padded[HPV_FUTVAL_IDX]
        if not isinstance(j_val, (int, float)) or isinstance(j_val, bool):
            continue

        # T-value (Total P&L EUR) is at offset 6 of the M:T block, written
        # by `_write_valuation_a_sheet`. The PL lookup carries it.
        # Here we read it from the (already-extended) row if present, else 0.
        t_val_idx = 12 + 6   # col S (1-based 19) = 0-based 18
        t_val = row_padded[t_val_idx] if len(row_padded) > t_val_idx else None
        t_val_f = float(t_val) if isinstance(t_val, (int, float)) and not isinstance(t_val, bool) else 0.0

        security_name = row_padded[HPV_SECURITY_NAME_IDX]
        # Fallback: if HiPort pivot has no carried-down initials, use the
        # analyst the rest of the pipeline assigned (analyst_map.csv +
        # PERSON_OVERRIDES). Keeps ValuationA tabs in sync with the main
        # P&L workbook even when HiPort hasn't tagged a position yet.
        person_for_row: object = current_person
        if (person_for_row in (None, "") and analyst_lookup):
            key = str(fmcid).strip() if fmcid not in (None, "") else ""
            fb = analyst_lookup.get(key)
            if fb:
                person_for_row = fb
        stock_rows.append(StockRow(
            source_row=source_row_1based,
            person=resolve_person_override(security_name, person_for_row, sname_overrides),
            raw_person=normalize_person(current_person),
            j_val=float(j_val),
            t_val=t_val_f,
        ))
    return stock_rows


def accumulate_by_person(
    stock_rows: Sequence[StockRow], people: Sequence[str],
) -> dict[str, SummaryAcc]:
    """Sum Fut Value (J) and Total P&L EUR (T) per person, split long/short."""
    acc = {p: {"long_j": 0.0, "short_j": 0.0, "long_t": 0.0, "short_t": 0.0} for p in people}
    for row in stock_rows:
        bucket = acc.get(row.person)
        if bucket is None:
            continue
        if row.j_val > 0:
            bucket["long_j"] += row.j_val
            bucket["long_t"] += row.t_val
        elif row.j_val < 0:
            bucket["short_j"] += row.j_val
            bucket["short_t"] += row.t_val
    return {p: SummaryAcc(**v) for p, v in acc.items()}


def extend_val_rows_with_lookup(
    val_rows: Sequence[Sequence[object]],
    pl_lookup: Mapping[str, Sequence[object]],
) -> list[list[object]]:
    """Return a copy of `val_rows` with M:T appended from `pl_lookup`.

    For data rows (>= DATA_FIRST_ROW), look up FMCID (col E) in `pl_lookup`
    and append the 8 trailing values; otherwise pad with None. Header row
    `HEADER_ROW_M_TO_T` (1-based 7) is overwritten with VAL_HEADERS_M_TO_T.
    """
    out: list[list[object]] = []
    for idx, row in enumerate(val_rows):
        excel_row = idx + 1
        base = list(row) + [None] * max(0, VAL_ACOL_COUNT - len(row))
        base = base[:VAL_ACOL_COUNT]
        if excel_row == HEADER_ROW_M_TO_T:
            tail: list[object] = list(VAL_HEADERS_M_TO_T)
        elif excel_row >= DATA_FIRST_ROW:
            fmcid = base[HPV_FMCID_IDX]
            key = str(fmcid).strip() if fmcid not in (None, "") else ""
            vals = pl_lookup.get(key)
            if vals is None:
                tail = [None] * VAL_MT_COUNT
            else:
                tail = list(vals) + [None] * max(0, VAL_MT_COUNT - len(vals))
                tail = tail[:VAL_MT_COUNT]
        else:
            tail = [None] * VAL_MT_COUNT
        out.append(base + tail)
    return out


# ─── Sheet writers ───────────────────────────────────────────────────────────

def _write_valuation_a_sheet(
    ws, extended_rows: Sequence[Sequence[object]], target_pname: str,
) -> int:
    """Write the ValuationA sheet from `extended_rows` (cols A:T)."""
    last_row = len(extended_rows)
    for r_idx, row in enumerate(extended_rows, start=1):
        for c_idx in range(1, VAL_ACOL_COUNT + VAL_MT_COUNT + 1):
            v = row[c_idx - 1] if c_idx - 1 < len(row) else None
            ws.cell(r_idx, c_idx, v)

    ws["B3"] = target_pname

    # Apply M:T number formats to all data rows.
    for col_offset, fmt in zip(range(13, 21), VAL_FORMATS_M_TO_T):
        for r_idx in range(DATA_FIRST_ROW, last_row + 1):
            ws.cell(r_idx, col_offset).number_format = fmt

    ws.freeze_panes = "G8"
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 10
    return last_row


def _write_summary_sheet(
    ws, stock_rows: Sequence[StockRow], target_pname: str,
) -> list[str]:
    """Build the Summary sheet (per-person long/short exposure + P&L)."""
    ws.sheet_properties.tabColor = Color(theme=7, tint=0.0)
    ws.row_dimensions[1].height = 18.0
    ws.row_dimensions[4].height = 15.75
    ws.row_dimensions[6].height = 40.15

    ws.merge_cells("B2:H2")
    ws["B2"] = "OAKS EM Fund - Stock-Level P&L Report"
    ws["B2"].font = TITLE_FONT
    ws.merge_cells("B3:H3")
    ws["B3"] = f"Period: {target_pname}"
    ws["B3"].font = SUBTITLE_FONT

    headers = (
        "Person",
        "Positive Exposure (Long)",
        "Positive Exposure (Long - %)",
        "Negative Exposure (Short)",
        "Negative Exposure (Short -  %)",
        "Net Exposure",
        "Total P&L (Long)",
        "Total P&L (Short)",
    )
    for c, header in enumerate(headers, start=1):
        cell = ws.cell(6, c, header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = ALIGN_CENTER_WRAP

    people = ordered_people(stock_rows)
    acc = accumulate_by_person(stock_rows, people)

    total_long_j = sum(v.long_j for v in acc.values())
    total_short_j = sum(v.short_j for v in acc.values())
    total_long_t = sum(v.long_t for v in acc.values())
    total_short_t = sum(v.short_t for v in acc.values())
    total_net_j = total_long_j + total_short_j

    for row_idx, person in enumerate(people, start=7):
        a = acc[person]
        ws.cell(row_idx, 1, person)
        ws.cell(row_idx, 2, a.long_j)
        ws.cell(row_idx, 3, (a.long_j / total_long_j) if total_long_j else 0.0)
        ws.cell(row_idx, 4, a.short_j)
        ws.cell(row_idx, 5, (a.short_j / total_short_j) if total_short_j else 0.0)
        ws.cell(row_idx, 6, a.long_j + a.short_j)
        ws.cell(row_idx, 7, a.long_t)
        ws.cell(row_idx, 8, a.short_t)

    total_row = 7 + len(people)
    ws.cell(total_row, 1, "TOTAL")
    ws.cell(total_row, 2, total_long_j)
    ws.cell(total_row, 3, (total_long_j / total_net_j) if total_net_j else 0.0)
    ws.cell(total_row, 4, total_short_j)
    ws.cell(total_row, 5, (total_short_j / total_net_j) if total_net_j else 0.0)
    ws.cell(total_row, 6, total_net_j)
    ws.cell(total_row, 7, total_long_t)
    ws.cell(total_row, 8, total_short_t)

    for r in range(7, total_row + 1):
        for c in range(1, 9):
            cell = ws.cell(r, c)
            cell.font = TOTAL_FONT if r == total_row else BODY_FONT
            if r == total_row:
                cell.fill = TOTAL_FILL
            cell.alignment = ALIGN_LEFT if c == 1 else ALIGN_RIGHT
            if c != 1:
                cell.number_format = FMT_PCT1 if c in (3, 5) else FMT_AMOUNT
    return people


def _write_detail_sheet(
    ws, extended_rows: Sequence[Sequence[object]],
    stock_rows: Sequence[StockRow], person: str, target_pname: str,
) -> None:
    """Write one analyst detail tab from `extended_rows` filtered by person."""
    rows_for_person = [row for row in stock_rows if row.person == person]
    display_person = detail_sheet_name(person)

    ws.sheet_properties.tabColor = Color(theme=7, tint=0.0)
    ws.freeze_panes = "I8"
    ws.row_dimensions[1].height = 18.0
    ws.row_dimensions[4].height = 15.75
    ws.row_dimensions[6].height = 40.15

    ws["A2"] = f"{display_person} - Stock-Level Detail"
    ws["A3"] = f"Period: {target_pname}"
    ws["A2"].font = TITLE_FONT
    ws["A3"].font = SUBTITLE_FONT

    for out_col, column in enumerate(DETAIL_COLUMN_LAYOUT, start=1):
        cell = ws.cell(7, out_col, column.header)
        if column.header:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = ALIGN_CENTER_WRAP
        ws.column_dimensions[get_column_letter(out_col)].width = column.width

    out_row = 8
    for row_info in rows_for_person:
        src_row_idx0 = row_info.source_row - 1   # 0-based into extended_rows
        if src_row_idx0 < 0 or src_row_idx0 >= len(extended_rows):
            continue
        src_row = extended_rows[src_row_idx0]
        for out_col, column in enumerate(DETAIL_COLUMN_LAYOUT, start=1):
            src_col = column.source_col
            value = None if src_col is None else (
                src_row[src_col - 1] if src_col - 1 < len(src_row) else None
            )
            dst = ws.cell(out_row, out_col, value)
            dst.font = BODY_FONT
            dst.alignment = column.align
            if src_col is not None and column.number_format is not None:
                dst.number_format = column.number_format
        out_row += 1


# ─── Public write entry ──────────────────────────────────────────────────────

def build_valuation_workbook(
    *,
    out_path: Path | str,
    val_rows: Sequence[Sequence[object]],
    pl_lookup: Mapping[str, Sequence[object]],
    target_pname: str = TARGET_PNAME,
    analyst_lookup: Mapping[str, str] | None = None,
    sname_overrides: Mapping[str, str] | None = None,
) -> Path:
    """Build the valuation workbook from extracted rows + a P&L lookup.

    Sheets, in order: ``Summary``, ``ValuationA``, then one tab per analyst
    appearing in `val_rows` (blank-bucket last as ``Unassigned``).
    Atomic save via tempfile + os.replace.

    Returns the resolved output path.
    """
    out_path = Path(out_path).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    extended = extend_val_rows_with_lookup(val_rows, pl_lookup)
    stock_rows = collect_stock_rows(
        extended, analyst_lookup=analyst_lookup,
        sname_overrides=sname_overrides,
    )

    wb = Workbook()
    # First sheet auto-created — repurpose as Summary so we keep ordering.
    summary_ws = wb.active
    summary_ws.title = SHEET_SUMMARY
    people = _write_summary_sheet(summary_ws, stock_rows, target_pname)

    val_ws = wb.create_sheet(SHEET_VALUATION_A)
    _write_valuation_a_sheet(val_ws, extended, target_pname)

    for person in people:
        ws = wb.create_sheet(detail_sheet_name(person))
        _write_detail_sheet(ws, extended, stock_rows, person, target_pname)

    return _atomic_save(wb, out_path)


def _atomic_save(wb: Workbook, out_path: Path) -> Path:
    """Write `wb` to a sibling tempfile and `os.replace` over `out_path`."""
    fd, tmp_str = tempfile.mkstemp(
        prefix=out_path.stem + ".", suffix=".xlsx.tmp", dir=str(out_path.parent),
    )
    os.close(fd)
    tmp = Path(tmp_str)
    try:
        wb.save(tmp)
        os.replace(tmp, out_path)
    finally:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
    return out_path


# ─── COM extraction (function-local imports — kept off module top level) ─────

def extract_valuation_a_via_sql(
    conn_str: str, *, target_pname: str = TARGET_PNAME,
) -> list[list[object]]:
    """Build the ValuationA row table by querying ``ccl.dbo.vw_RPT_VAL``.

    Returns a list-of-lists with the same shape ``extract_valuation_a_via_com``
    produces (cols A:L), so it is a drop-in replacement for the COM extractor:

      * Rows 1-7 are header / title rows. Row 7 (``HEADER_ROW_M_TO_T``) carries
        the A:L column labels — the M:T labels are filled in later by
        ``extend_val_rows_with_lookup``.
      * Rows 8+ (``DATA_FIRST_ROW``) are the SORT1='A' stock-level rows
        for the OAKS parent fund (PNAME match), one per FMCID, ordered by
        SORT2 then SNAME for stable output.

    Column meanings (1-based):
        A  SORT1          — always 'A' (downstream filter key for stocks)
        B  INITIALS       — left blank; ``analyst_lookup`` fills via FMCID
        C  (unused)       — kept blank to match the legacy pivot layout
        D  SECURITY NAME  — ``SNAME``
        E  FMCID          — primary lookup key into ``pl_lookup``
        F  BTICKER        — ``BTICKER_FULL``
        G  Cost           — ``PTCOST`` (fund CCY = EUR for OEFOF)
        H  Nominal        — ``UNITS``
        I  Price          — ``LST_PRICE`` (local CCY)
        J  Fut Value      — ``EXPOSURE``  (signed — sign decides long/short)
        K  P/F Value      — ``PTVALUE``   (marked value, EUR)
        L  Exposure %     — ``PERC_EXP_PTVALUE`` (already a fraction 0..1)

    The COM-extractor sign convention is preserved by using ``EXPOSURE`` for
    ``Fut Value``: HiPort sets ``EXPOSURE`` to the full notional for short
    futures/swaps (so ``j_val < 0`` correctly buckets them as shorts in
    :func:`accumulate_by_person`), whereas ``PTVALUE`` would only carry the
    small marked-to-market P&L on those short instruments.

    Raises ``DataLoadError`` on any pyodbc failure.
    """
    try:
        import pyodbc  # type: ignore
    except ImportError as exc:   # pragma: no cover — dependency required in prod
        raise ImportError(
            "pyodbc is required for ValuationA SQL extraction. "
            "Install with: pip install pyodbc"
        ) from exc

    sql = (
        "SELECT SORT1, SORT2, SNAME, FMCID, BTICKER_FULL, "
        "PTCOST, UNITS, LST_PRICE, EXPOSURE, PTVALUE, PERC_EXP_PTVALUE "
        "FROM dbo.vw_RPT_VAL "
        "WHERE PNAME = ? "
        "  AND VDATE = (SELECT MAX(VDATE) FROM dbo.vw_RPT_VAL WHERE PNAME = ?) "
        "ORDER BY SORT1, SORT2, SNAME"
    )

    try:
        with pyodbc.connect(conn_str, timeout=30) as cn:
            cur = cn.cursor()
            cur.execute(sql, target_pname, target_pname)
            data = cur.fetchall()
    except Exception as exc:   # pyodbc.Error is the common parent
        # Match existing COM-extractor failure semantics: raise so the
        # pipeline can record the stage failure and continue without
        # the ValuationA tabs.
        raise RuntimeError(
            f"vw_RPT_VAL query for PNAME={target_pname!r} failed: {exc}"
        ) from exc

    # Build header rows 1-7. Row 3 column B is overwritten with
    # ``target_pname`` by ``_write_valuation_a_sheet``; we leave it blank
    # here to avoid double-writing.
    blank12: list[object] = [None] * VAL_ACOL_COUNT
    rows: list[list[object]] = [list(blank12) for _ in range(HEADER_ROW_M_TO_T - 1)]
    header_row: list[object] = [
        "SORT1", "INITIALS", "SORT2", "SNAME", "FMCID", "BTICKER",
        "Cost", "Nominal", "Price", "Fut Value", "P/F Value", "Exposure %",
    ]
    rows.append(header_row)
    assert len(rows) == HEADER_ROW_M_TO_T, (
        f"header block size {len(rows)} != HEADER_ROW_M_TO_T={HEADER_ROW_M_TO_T}"
    )

    def _to_float(v: object) -> float | None:
        if v is None:
            return None
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    for r in data:
        # r = (SORT1, SORT2, SNAME, FMCID, BTICKER_FULL, PTCOST, UNITS,
        #      LST_PRICE, EXPOSURE, PTVALUE, PERC_EXP_PTVALUE)
        rows.append([
            r[0],                          # A SORT1
            None,                          # B INITIALS (analyst_lookup fills)
            r[1],                          # C SORT2 (country grouping)
            r[2],                          # D SECURITY NAME
            r[3],                          # E FMCID
            r[4],                          # F BTICKER
            _to_float(r[5]),               # G Cost
            _to_float(r[6]),               # H Nominal
            _to_float(r[7]),               # I Price
            _to_float(r[8]),               # J Fut Value (EXPOSURE)
            _to_float(r[9]),               # K P/F Value (PTVALUE)
            _to_float(r[10]),              # L Exposure % (already a fraction)
        ])
    return rows


def extract_valuation_a_via_com(
    hp_val_path: Path | str, *, target_pname: str = TARGET_PNAME,
) -> list[list[object]]:
    """Open ``HP_VAL.xlsm``, set ValuationA pivot page → ``target_pname``,
    refresh, and extract cols A:L as a list-of-lists of Python values.

    Raises ``ImportError`` on systems without ``pywin32``.
    """
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:   # pragma: no cover - platform-gated
        raise ImportError(
            "pywin32 is required for ValuationA COM extraction. "
            "Install with: pip install pywin32"
        ) from exc

    src = Path(hp_val_path).resolve()
    if not src.exists():
        raise FileNotFoundError(f"HP_VAL workbook not found: {src}")

    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0)
        ws = wb.Sheets("ValuationA")

        n_pts = ws.PivotTables().Count
        for i in range(1, n_pts + 1):
            pt = ws.PivotTables(i)
            try:
                page_fields = pt.PageFields
                n_pf = page_fields.Count
            except Exception:
                n_pf = 0
            for j in range(1, n_pf + 1):
                pf = page_fields(j)
                try:
                    pf.CurrentPage = target_pname
                except Exception:
                    pass
            pt.RefreshTable()
        xl.Calculate()

        scan_max = 2000
        scan_rng = ws.Range(ws.Cells(DATA_FIRST_ROW, 1), ws.Cells(scan_max, 5))
        scan_vals = scan_rng.Value
        last_data_row = DATA_FIRST_ROW - 1
        if scan_vals is not None:
            for r_offset, row in enumerate(scan_vals):
                if any(v not in (None, "") for v in (row[0], row[3], row[4])):
                    last_data_row = DATA_FIRST_ROW + r_offset

        rng = ws.Range(ws.Cells(1, 1), ws.Cells(last_data_row, VAL_ACOL_COUNT))
        raw = rng.Value
        if raw is None:
            rows: list[list[object]] = []
        elif not isinstance(raw[0], (tuple, list)):
            rows = [[v for v in raw]]
        else:
            rows = [list(r) for r in raw]
        wb.Close(SaveChanges=False)
        return rows
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def close_open_excel_workbook(target_path: Path | str) -> bool:
    """Best-effort: close `target_path` if any running Excel instance has it open.

    Returns True if a workbook was found and closed. Never raises — any COM
    failure is treated as "nothing to close".
    """
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return False

    target = Path(target_path).resolve()
    target_name = target.name.lower()
    closed = False
    pythoncom.CoInitialize()
    try:
        # Strategy 1: Running Object Table walk (catches all Excel instances).
        try:
            ctx = pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable()
            enum_monikers = rot.EnumRunning()
            # Hard cap on iterations — pywin32's IEnumMoniker.Next() is known
            # to return None / empty tuples on exhaustion in some versions
            # instead of raising, which would otherwise spin this loop forever.
            for _ in range(4096):
                try:
                    moniker = enum_monikers.Next()
                except Exception:
                    break
                if moniker is None or moniker == ():
                    break
                # pywin32 wraps IEnumMoniker.Next as either a single moniker
                # or a 1-tuple depending on version; unwrap defensively.
                if isinstance(moniker, tuple):
                    if not moniker:
                        break
                    moniker = moniker[0]
                try:
                    display_name = moniker.GetDisplayName(ctx, None)
                except Exception:
                    continue
                if Path(display_name).name.lower() == target_name:
                    try:
                        unk = rot.GetObject(moniker)
                        wb = win32.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))
                        wb.Close(SaveChanges=False)
                        closed = True
                        break
                    except Exception:
                        continue
        except Exception:
            pass

        # Strategy 2: GetObject fallback (single Excel instance).
        if not closed:
            try:
                xl = win32.GetObject(Class="Excel.Application")
                for wb in xl.Workbooks:
                    try:
                        full = str(Path(str(wb.FullName)).resolve()).lower()
                    except Exception:
                        full = ""
                    if full == str(target).lower() or str(wb.Name).lower() == target_name:
                        wb.Close(SaveChanges=False)
                        closed = True
                        break
            except Exception:
                pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return closed


# ─── Orchestration ───────────────────────────────────────────────────────────

@dataclass(frozen=True)
class ValuationResult:
    out_path: Path
    val_row_count: int
    person_tab_count: int


def run_valuation(
    *,
    source: object,
    out_path: Path | str,
    pl_lookup: Mapping[str, Sequence[object]],
    target_pname: str = TARGET_PNAME,
    extractor=extract_valuation_a_via_sql,
    closer=close_open_excel_workbook,
    analyst_lookup: Mapping[str, str] | None = None,
    sname_overrides: Mapping[str, str] | None = None,
) -> ValuationResult:
    """End-to-end: close any open output, extract ValuationA rows, write.

    ``source`` is the first positional value passed to ``extractor``. For
    the default SQL extractor this is the pyodbc connection string; for
    the legacy COM extractor it is the path to ``HP_VAL.xlsm``. ``extractor``
    and ``closer`` are injected so tests can swap them for in-memory fakes.
    """
    out = Path(out_path).resolve()
    closer(out)
    val_rows = extractor(source, target_pname=target_pname)
    extended = extend_val_rows_with_lookup(val_rows, pl_lookup)
    stock_rows = collect_stock_rows(
        extended, analyst_lookup=analyst_lookup,
        sname_overrides=sname_overrides,
    )
    written = build_valuation_workbook(
        out_path=out, val_rows=val_rows, pl_lookup=pl_lookup,
        target_pname=target_pname, analyst_lookup=analyst_lookup,
        sname_overrides=sname_overrides,
    )
    return ValuationResult(
        out_path=written,
        val_row_count=len(val_rows),
        person_tab_count=len(ordered_people(stock_rows)),
    )
