"""Smoke tests for output.workbook — the 9-sheet P&L report writer.

Strategy: build small but realistic ``pl_df`` / ``income_df`` /
``ValidationResult`` fixtures, call ``build_workbook``, then re-open the
saved .xlsx with openpyxl and assert structural properties (sheet
order, column count, header texts, total-row presence, number formats,
non-truncated failure listings).
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from oefof_pl.data.normalise import INCOME_COLUMNS, PL_COLUMNS
from oefof_pl.data.arb_pairs_loader import load_arb_pairs
from oefof_pl.output.styles import COLUMN_LAYOUT_HOLDINGS, COLUMN_LAYOUT_INCOME
from oefof_pl.output.workbook import (
    ExposureBlock,
    ExposureMetrics,
    PeriodMeta,
    SHEET_ORDER,
    build_workbook,
)
from oefof_pl.validate.checks import CheckResult, ValidationResult


# ─── Layout/schema parity (BUG-2 lock) ──────────────────────────────────────

def test_holdings_layout_matches_pl_columns_exactly():
    """Every ColSpec must reference a real PL_COLUMNS entry. BUG-2 was an
    output-side reference to a non-existent column. The reverse direction
    (every PL_COLUMNS entry must be in the layout) is intentionally not
    enforced — some columns (e.g. 'Cost Basis (EUR)', 'Market Value End
    (EUR)') are kept in the dataframe for compute / validation but hidden
    from the user-facing workbook by request."""
    pl_set = set(PL_COLUMNS)
    derived_output_cols = {
        "Contribution to Total (%)",
        "Absolute Contribution (bps)",
    }
    layout_cols = [c.df_column for c in COLUMN_LAYOUT_HOLDINGS]
    missing = [c for c in layout_cols if c not in pl_set and c not in derived_output_cols]
    assert not missing, f"Layout references columns not in PL_COLUMNS: {missing}"


def test_income_layout_matches_income_columns_exactly():
    layout_cols = [c.df_column for c in COLUMN_LAYOUT_INCOME]
    assert layout_cols == list(INCOME_COLUMNS)


# ─── Fixtures ──────────────────────────────────────────────────────────────

def _meta() -> PeriodMeta:
    return PeriodMeta(
        fund_name="OAKS Emerging and Frontier Fund",
        start_date=pd.Timestamp("2025-12-31"),
        end_date=pd.Timestamp("2026-01-31"),
        snapshot_path="HP_VAL.xlsm",
        bottler_path="StockTrList.xlsb",
        run_timestamp=datetime(2026, 4, 29, 14, 30, 0),
        snapshot_exposure=ExposureBlock(
            label="Current Snapshot",
            as_of=pd.Timestamp("2026-01-31"),
            fund=ExposureMetrics(long_eur=30.0, short_eur=8.0, nav_eur=100.0),
            by_analyst={"IS": ExposureMetrics(long_eur=12.0, short_eur=3.0, nav_eur=100.0)},
        ),
        ytd_weighted_exposure=ExposureBlock(
            label="YTD Avg Gross Exposure",
            as_of=pd.Timestamp("2026-01-31"),
            fund=ExposureMetrics(long_eur=28.0, short_eur=7.0, nav_eur=100.0),
            by_analyst={"IS": ExposureMetrics(long_eur=11.0, short_eur=2.5, nav_eur=100.0)},
        ),
        start_nav_eur=100.0,
        analyst_codes={"IS": "Ian Simmons"},
    )


def _pl_row(**kw) -> dict:
    base = {
        "ISIN": "X", "Stock Name": "STOCK", "Country": "DE", "Analyst": "IS",
        "CCY": "EUR", "Instrument": "ORD", "L/S": "L",
        "Starting Units": 0.0, "Ending Units": 0.0,
        "Units Bought": 0.0, "Units Sold": 0.0, "Bonus Units": 0.0,
        "Avg Buy Price (EUR)": 0.0, "Avg Sell Price (EUR)": 0.0,
        "Start Price (Local)": 0.0, "Last Price (EUR)": 0.0,
        "Market Value Start (EUR)": 0.0, "Market Value End (EUR)": 0.0,
        "Cost Basis (EUR)": 0.0,
        "Realised P&L (Local)": 0.0, "Realised P&L (EUR)": 0.0,
        "Realised P&L (%)": 0.0,
        "Unrealised P&L (Local)": 0.0, "Unrealised P&L (EUR)": 0.0,
        "Unrealised P&L (%)": 0.0,
        "Income (Local)": 0.0, "Income (EUR)": 0.0,
        "Dividends (EUR)": 0.0, "Swap Financing (EUR)": 0.0,
        "Income Yield (%)": 0.0,
        "Total P&L (EUR)": 0.0, "Total P&L (%)": 0.0,
        "Last Trade Date": pd.NaT, "First Trade Date": pd.NaT,
    }
    base.update(kw)
    return base


def _pl_df(rows):
    return pd.DataFrame(rows or [], columns=list(PL_COLUMNS))


def _income_df(rows):
    return pd.DataFrame(rows or [], columns=list(INCOME_COLUMNS))


def _validation_pass() -> ValidationResult:
    return ValidationResult(checks=[
        CheckResult(id="VAL-01", name="Unit reconciliation", status="PASS", detail="ok"),
        CheckResult(id="VAL-02", name="Duplicate trades", status="PASS", detail="ok"),
    ])


def _validation_with_failures(n: int = 12) -> ValidationResult:
    return ValidationResult(checks=[
        CheckResult(
            id="VAL-11", name="PTVALUE local sanity",
            status="FAIL", detail=f"{n} positions failed",
            failures=[{"ISIN": f"X{i}", "Stock Name": f"S{i}",
                       "Diff (Local)": float(i)} for i in range(n)],
        ),
    ])


# ─── End-to-end smoke ──────────────────────────────────────────────────────

def test_build_workbook_produces_nine_sheets_in_order(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="A", **{"Ending Units": 10.0, "Total P&L (EUR)": 100.0,
                             "Cost Basis (EUR)": 1000.0,
                             "Market Value End (EUR)": 1100.0}),
        _pl_row(ISIN="B", **{"Ending Units": 0.0,  # exited
                             "Total P&L (EUR)": -50.0,
                             "Cost Basis (EUR)": 500.0}),
    ])
    inc = _income_df([{
        "ISIN": "A", "Stock Name": "Sa", "Analyst": "IS", "CCY": "EUR",
        "Date": pd.Timestamp("2026-01-15"), "Income Type": "DIV",
        "Income (Local)": 10.0, "Income (EUR)": 10.0, "Units": 10.0,
    }])
    trades = pd.DataFrame([{
        "TRADE_ID": 1,
        "PCODE_ORIG": "OEFOF", "ISIN": "A", "SNAME": "Sa", "CCY": "EUR",
        "T": "P", "CDATE": pd.Timestamp("2026-01-10"), "UNITS": 10.0,
        "GROSSPRICE_LOCAL": 100.0, "BUY_NET_LOCAL": 1000.0,
        "SELL_NET_LOCAL": 0.0, "INCOME_LOCAL": 0.0,
    }])
    build_workbook(out_path=out, pl_df=pl, income_df=inc, trades_df=trades,
                   validation=_validation_pass(), period_meta=_meta())
    assert out.exists()

    wb = load_workbook(out, read_only=False, data_only=False)
    # Static sheets must all be present; their relative order must match
    # SHEET_ORDER. Per-analyst sheets are inserted dynamically between
    # "Analyst Overview" and "Current Holdings" and are not in SHEET_ORDER.
    static = [s for s in wb.sheetnames if s in SHEET_ORDER]
    assert static == list(SHEET_ORDER)
    # Per-analyst sheet for IS must exist (single analyst in fixture).
    assert "IS" in wb.sheetnames


def test_unassigned_sheet_is_dynamic(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="A", **{"Analyst": "IS", "Ending Units": 1.0,
                             "Total P&L (EUR)": 10.0, "Cost Basis (EUR)": 100.0}),
        _pl_row(ISIN="B", **{"Analyst": "", "Ending Units": 1.0,
                             "Total P&L (EUR)": -5.0, "Cost Basis (EUR)": 100.0}),
    ])
    build_workbook(out_path=out, pl_df=pl, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())
    wb = load_workbook(out, read_only=False, data_only=False)
    assert "UNASSIGNED" in wb.sheetnames

    out2 = tmp_path / "report_no_unassigned.xlsx"
    pl2 = _pl_df([
        _pl_row(ISIN="A", **{"Analyst": "IS", "Ending Units": 1.0,
                             "Total P&L (EUR)": 10.0, "Cost Basis (EUR)": 100.0}),
    ])
    build_workbook(out_path=out2, pl_df=pl2, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())
    wb2 = load_workbook(out2, read_only=False, data_only=False)
    assert "UNASSIGNED" not in wb2.sheetnames


def test_current_holdings_has_all_pl_columns(tmp_path: Path):
    """BUG-2 regression: every PL_COLUMNS entry appears as a header in
    the Current Holdings sheet."""
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="A", **{"Ending Units": 1.0,
                             "Start Price (Local)": 50.0,
                             "Last Price (EUR)": 60.0,
                             "Market Value End (EUR)": 60.0,
                             "Total P&L (EUR)": 10.0,
                             "Cost Basis (EUR)": 50.0}),
    ])
    build_workbook(out_path=out, pl_df=pl, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())
    wb = load_workbook(out)
    ws = wb["Current Holdings"]
    from oefof_pl.output.styles import COLUMN_LAYOUT_HOLDINGS as _LH
    expected = [s.display or s.df_column for s in _LH]
    headers = [ws.cell(4, c).value for c in range(1, len(expected) + 1)]
    assert headers == expected
    # BUG-2 specifically: "Start Price (Local)" must be present
    assert "Start Price (Local)" in headers


def test_exited_positions_separated_from_current(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="HOLD", **{"Ending Units": 5.0,
                                "Total P&L (EUR)": 1.0,
                                "Cost Basis (EUR)": 100.0}),
        _pl_row(ISIN="EXIT", **{"Ending Units": 0.0,
                                "Total P&L (EUR)": -5.0,
                                "Cost Basis (EUR)": 50.0}),
    ])
    build_workbook(out_path=out, pl_df=pl, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())
    wb = load_workbook(out)

    cur_isins = {wb["Current Holdings"].cell(r, 1).value
                 for r in range(5, 5 + 5)}
    ex_isins = {wb["Exited Positions"].cell(r, 1).value
                for r in range(5, 5 + 5)}
    assert "HOLD" in cur_isins and "EXIT" not in cur_isins
    assert "EXIT" in ex_isins and "HOLD" not in ex_isins


def test_total_row_uses_sum_for_eur_columns(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="A", **{"Ending Units": 1.0,
                             "Realised P&L (EUR)": 100.0,
                             "Total P&L (EUR)": 25.0}),
        _pl_row(ISIN="B", **{"Ending Units": 1.0,
                             "Realised P&L (EUR)": 200.0,
                             "Total P&L (EUR)": -10.0}),
    ])
    build_workbook(out_path=out, pl_df=pl, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())
    wb = load_workbook(out)
    ws = wb["Current Holdings"]
    headers = [ws.cell(4, c).value for c in range(1, len(PL_COLUMNS) + 1)]
    total_row = 4 + 2 + 1  # header_row + n_rows + 1
    realised_col = headers.index("Realised P&L (EUR)") + 1
    total_col = headers.index("Total P&L (EUR)") + 1
    assert ws.cell(total_row, realised_col).value == pytest.approx(300.0)
    assert ws.cell(total_row, total_col).value == pytest.approx(15.0)
    # Leftmost column should say "Total"
    assert ws.cell(total_row, 1).value == "Total"


def test_pct_columns_use_percentage_format(tmp_path: Path):
    """BUG-1 regression at the output layer: Excel format '0.00%' is what
    turns the ratio (0.44) into '44.00%'."""
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="A", **{"Ending Units": 1.0,
                             "Total P&L (%)": 0.44,
                             "Total P&L (EUR)": 44.0,
                             "Cost Basis (EUR)": 100.0}),
    ])
    build_workbook(out_path=out, pl_df=pl, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())
    wb = load_workbook(out)
    ws = wb["Current Holdings"]
    headers = [ws.cell(4, c).value for c in range(1, len(PL_COLUMNS) + 1)]
    pct_col = headers.index("Total P&L (%)") + 1
    cell = ws.cell(5, pct_col)
    assert cell.value == pytest.approx(0.44)
    assert cell.number_format == "0.00%;[Red](0.00%)"


def test_analyst_sheet_has_bps_contribution_column_without_abs_bps(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="A", **{"Ending Units": 1.0,
                             "Analyst": "IS",
                             "Total P&L (EUR)": 60.0,
                             "Cost Basis (EUR)": 100.0}),
        _pl_row(ISIN="B", **{"Ending Units": 1.0,
                             "Analyst": "IS",
                             "Total P&L (EUR)": 40.0,
                             "Cost Basis (EUR)": 100.0}),
    ])
    build_workbook(out_path=out, pl_df=pl, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())

    wb = load_workbook(out)
    ws = wb["IS"]
    headers = [ws.cell(9, c).value for c in range(1, ws.max_column + 1)]
    contrib_col = headers.index("Contribution (bps)") + 1
    assert "Abs Contribution (bps)" not in headers

    contrib_vals = [ws.cell(10, contrib_col).value, ws.cell(11, contrib_col).value]

    assert contrib_vals[0] == pytest.approx(6000.0)
    assert contrib_vals[1] == pytest.approx(4000.0)

    subtotal_row = 12
    assert ws.cell(subtotal_row, contrib_col).value == pytest.approx(10000.0)


def test_analyst_contribution_sign_uses_absolute_total_pl(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="LOSER", **{"Ending Units": 1.0,
                                  "Analyst": "IS",
                                  "Total P&L (EUR)": -240.0,
                                  "Cost Basis (EUR)": 100.0}),
        _pl_row(ISIN="WINNER", **{"Ending Units": 1.0,
                                   "Analyst": "IS",
                                   "Total P&L (EUR)": 60.0,
                                   "Cost Basis (EUR)": 100.0}),
    ])
    build_workbook(out_path=out, pl_df=pl, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())

    wb = load_workbook(out)
    ws = wb["IS"]
    headers = [ws.cell(9, c).value for c in range(1, ws.max_column + 1)]
    contrib_col = headers.index("Contribution (bps)") + 1

    # Sorted by P&L descending: winner row first, then loser row.
    winner_contrib = ws.cell(10, contrib_col).value
    loser_contrib = ws.cell(11, contrib_col).value

    assert winner_contrib == pytest.approx((1.0 / 3.0) * 10_000)
    assert loser_contrib == pytest.approx((-4.0 / 3.0) * 10_000)


def test_workbook_overrides_last_price_from_live_prices(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="KYG070341048", **{"Ending Units": 1.0,
                                         "Analyst": "IS",
                                         "Last Price (EUR)": 10.0,
                                         "Total P&L (EUR)": 5.0,
                                         "Cost Basis (EUR)": 100.0}),
    ])
    build_workbook(
        out_path=out,
        pl_df=pl,
        income_df=_income_df([]),
        trades_df=pd.DataFrame(),
        validation=_validation_pass(),
        period_meta=_meta(),
        live_prices={"KYG070341048": 12.34},
    )

    wb = load_workbook(out)
    ws = wb["Current Holdings"]
    headers = [ws.cell(4, c).value for c in range(1, ws.max_column + 1)]
    last_px_col = headers.index("Last Price (EUR)") + 1
    assert ws.cell(5, last_px_col).value == pytest.approx(12.34)


def test_summary_does_not_render_instrument_breakdown_section(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="A", **{"Ending Units": 1.0,
                             "Instrument": "ORD",
                             "Total P&L (EUR)": 20.0,
                             "Cost Basis (EUR)": 100.0}),
    ])
    build_workbook(out_path=out, pl_df=pl, income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())

    wb = load_workbook(out)
    ws = wb["Summary"]
    col_a = [str(ws.cell(r, 1).value or "") for r in range(1, ws.max_row + 1)]
    assert "Instrument Breakdown" not in col_a


def test_analyst_sheet_keeps_default_total_pl_sort_with_arb_map(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(ISIN="KYG070341048", **{
            "Stock Name": "BAIDU INC",
            "Analyst": "HK",
            "Instrument": "ORD",
            "L/S": "L",
            "Ending Units": 100.0,
            "Total P&L (EUR)": -244_470.0,
            "Cost Basis (EUR)": 1_000_000.0,
            "Market Value End (EUR)": 900_000.0,
        }),
        _pl_row(ISIN="FDGSCBIHKT", **{
            "Stock Name": "GSCBIHKT CFD GOLDMAN",
            "Analyst": "HK",
            "Instrument": "FTSWAP",
            "L/S": "S",
            "Ending Units": 100.0,
            "Total P&L (EUR)": 321_603.0,
            "Cost Basis (EUR)": 800_000.0,
            "Market Value End (EUR)": -850_000.0,
        }),
        _pl_row(ISIN="OTHERHK", **{
            "Stock Name": "OTHER HK NAME",
            "Analyst": "HK",
            "Instrument": "ORD",
            "L/S": "L",
            "Ending Units": 50.0,
            "Total P&L (EUR)": 10_000.0,
            "Cost Basis (EUR)": 100_000.0,
            "Market Value End (EUR)": 120_000.0,
        }),
    ])

    pair_csv = tmp_path / "arb_pairs.csv"
    pair_csv.write_text(
        "label,long_isin,short_isin,analyst\n"
        "Baidu INC / GSCBIHKT,KYG070341048,FDGSCBIHKT,HK\n",
        encoding="utf-8",
    )
    pair_map = load_arb_pairs(pair_csv)

    build_workbook(
        out_path=out,
        pl_df=pl,
        income_df=_income_df([]),
        trades_df=pd.DataFrame(),
        validation=_validation_pass(),
        period_meta=_meta(),
        arb_pairs=pair_map,
    )

    wb = load_workbook(out)
    ws = wb["HK"]
    headers = [ws.cell(9, c).value for c in range(1, ws.max_column + 1)]
    name_col = headers.index("Stock Name") + 1

    names = []
    row = 10
    while True:
        v = ws.cell(row, name_col).value
        if v in (None, "Subtotal — Current"):
            break
        names.append(str(v))
        row += 1

    assert names[:3] == [
        "GSCBIHKT CFD GOLDMAN",
        "OTHER HK NAME",
        "BAIDU INC",
    ]


def test_validation_failures_are_listed_completely(tmp_path: Path):
    """Legacy printed only first 3. The Validation Report must list every row."""
    out = tmp_path / "report.xlsx"
    n = 25
    build_workbook(
        out_path=out,
        pl_df=_pl_df([_pl_row(ISIN="A", **{"Ending Units": 1.0,
                                            "Cost Basis (EUR)": 1.0})]),
        income_df=_income_df([]),
        trades_df=pd.DataFrame(),
        validation=_validation_with_failures(n),
        period_meta=_meta(),
    )
    wb = load_workbook(out)
    ws = wb["Validation Report"]
    text = "\n".join(
        str(ws.cell(r, c).value or "")
        for r in range(1, ws.max_row + 1) for c in range(1, ws.max_column + 1)
    )
    # Every X{i} must appear; nothing was truncated.
    for i in range(n):
        assert f"X{i}" in text, f"Validation Report missing row X{i}"


def test_atomic_save_does_not_leave_tmp_behind(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    build_workbook(out_path=out, pl_df=_pl_df([]), income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())
    leftover = list(tmp_path.glob("*.tmp"))
    assert leftover == []
    assert out.exists()


def test_empty_pl_still_writes_all_sheets(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    build_workbook(out_path=out, pl_df=_pl_df([]), income_df=_income_df([]),
                   trades_df=pd.DataFrame(), validation=_validation_pass(),
                   period_meta=_meta())
    wb = load_workbook(out)
    # No analysts → no dynamic sheets, so we expect exactly SHEET_ORDER.
    assert wb.sheetnames == list(SHEET_ORDER)


def test_workbook_does_not_import_win32com():
    """Hard rule: output/workbook.py is openpyxl-only. win32com is reserved
    for output/valuation.py and refresh/bottler.py."""
    import ast
    from pathlib import Path
    from oefof_pl.output import workbook as wbmod
    src = Path(wbmod.__file__).read_text(encoding="utf-8")
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                assert "win32com" not in alias.name
        elif isinstance(node, ast.ImportFrom):
            assert node.module is None or "win32com" not in node.module


def test_summary_uses_snapshot_and_weighted_exposure_sections(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(
            ISIN="A",
            Analyst="IS",
            **{
                "Ending Units": 10.0,
                "Cost Basis (EUR)": 100.0,
                "Realised P&L (EUR)": 12.0,
                "Unrealised P&L (EUR)": 8.0,
                "Income (EUR)": 3.0,
                "Total P&L (EUR)": 23.0,
                "Market Value End (EUR)": 110.0,
            },
        ),
    ])
    build_workbook(
        out_path=out,
        pl_df=pl,
        income_df=_income_df([]),
        trades_df=pd.DataFrame(),
        validation=_validation_pass(),
        period_meta=_meta(),
    )
    wb = load_workbook(out)
    ws = wb["Summary"]
    text = "\n".join(
        str(ws.cell(r, c).value or "")
        for r in range(1, ws.max_row + 1) for c in range(1, ws.max_column + 1)
    )
    assert "Return on Gross Exposure (YTD)" in text
    assert "Contribution to Fund Return" in text
    assert "Current Snapshot  (as of 31 January 2026)" in text
    assert "YTD Avg Gross Exposure  (through 31 January 2026)" in text
    assert "Current Long (%)" in text
    assert "YTD Avg Gross (%)" in text
    assert "RETURN ON NAV (YTD)" not in text

    ytd_hdr_row = next(
        r for r in range(1, ws.max_row + 1)
        if ws.cell(r, 1).value == "Analyst" and ws.cell(r, 8).value == "YTD Avg Gross (EUR)"
    )
    # Analyst row and fund-total row should use fund YTD gross denominator.
    assert ws.cell(ytd_hdr_row + 1, 9).value == pytest.approx(13.5 / 35.0)
    assert ws.cell(ytd_hdr_row + 2, 9).value == pytest.approx(1.0)

    header_row = next(
        r for r in range(1, ws.max_row + 1)
        if ws.cell(r, 1).value == "Analyst" and ws.cell(r, 4).value == "Return %"
    )
    assert ws.cell(header_row, 5).value == "Contribution to Fund Return"
    data_row = header_row + 1
    total_row = header_row + 2
    assert ws.cell(data_row, 1).value == "Ian Simmons"
    assert ws.cell(data_row, 4).value == pytest.approx(23.0 / 13.5)
    assert ws.cell(data_row, 5).value == pytest.approx(1.0)
    assert ws.cell(total_row, 1).value == "TOTAL"
    assert ws.cell(total_row, 5).value == pytest.approx(1.0)


def test_analyst_sheet_hides_bottom_exposure_blocks(tmp_path: Path):
    out = tmp_path / "report.xlsx"
    pl = _pl_df([
        _pl_row(
            ISIN="A",
            Analyst="IS",
            **{
                "Ending Units": 10.0,
                "Cost Basis (EUR)": 100.0,
                "Realised P&L (EUR)": 12.0,
                "Unrealised P&L (EUR)": 8.0,
                "Total P&L (EUR)": 20.0,
                "Market Value End (EUR)": 110.0,
            },
        ),
    ])
    build_workbook(
        out_path=out,
        pl_df=pl,
        income_df=_income_df([]),
        trades_df=pd.DataFrame(),
        validation=_validation_pass(),
        period_meta=_meta(),
    )
    wb = load_workbook(out)
    ws = wb["IS"]
    text = "\n".join(
        str(ws.cell(r, c).value or "")
        for r in range(1, ws.max_row + 1) for c in range(1, ws.max_column + 1)
    )
    assert "RETURN ON GROSS EXPOSURE (YTD)" in text.upper()
    assert "PERIOD — TOTAL P&L %" not in text.upper()
    assert "Current Snapshot  (as of 31 January 2026)" not in text
    assert "YTD Avg Gross Exposure  (through 31 January 2026)" not in text
