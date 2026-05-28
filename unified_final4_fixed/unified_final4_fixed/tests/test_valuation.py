"""Tests for ``oefof_pl.output.valuation`` (ValuationA workbook builder)."""
from __future__ import annotations

import ast
import os
from pathlib import Path
from typing import Sequence

import pytest
from openpyxl import load_workbook

from oefof_pl.output import valuation as V
from oefof_pl.output.valuation import (
    BLANK_PERSON_KEY,
    SHEET_SUMMARY,
    SHEET_VALUATION_A,
    StockRow,
    ValuationResult,
    accumulate_by_person,
    build_valuation_workbook,
    close_open_excel_workbook,
    collect_stock_rows,
    detail_sheet_name,
    extend_val_rows_with_lookup,
    normalize_person,
    ordered_people,
    resolve_person_override,
    run_valuation,
)


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _blank_row(n: int = 12) -> list[object]:
    return [None] * n


def _data_row(
    *, sort1: object = None, initials: object = None, sname: object = None,
    fmcid: object = None, fut_value: object = None, cost: float = 0.0,
    nominal: float = 0.0, price: float = 0.0, pf_value: float = 0.0,
    exposure: float = 0.0, sort2: object = None, bticker: object = None,
) -> list[object]:
    """Build a 12-element ValuationA row (cols A:L)."""
    return [
        sort1, initials, sort2, sname, fmcid, bticker,
        cost, nominal, price, fut_value, pf_value, exposure,
    ]


def _sample_val_rows() -> list[list[object]]:
    """Construct val_rows with 7 header rows + 4 stock rows.

    Layout:
      row 1..6: pivot/header padding (mostly blanks)
      row 7   : the M:T-headers row (left blank — caller will overwrite)
      row 8   : SORT1='A', initials='IS', stock with positive Fut Value
      row 9   : carry-down initials, positive Fut Value
      row 10  : SORT1 carries 'A', initials='HK', negative Fut Value
      row 11  : 'Total' row — should reset person and be skipped
    """
    rows: list[list[object]] = []
    for _ in range(6):
        rows.append(_blank_row())
    rows.append(_blank_row())   # row 7
    rows.append(_data_row(sort1="A", initials="IS", sname="ALPHA CORP",
                          fmcid="FMC-001", fut_value=1000.0))
    rows.append(_data_row(sname="BETA LTD", fmcid="FMC-002", fut_value=500.0))
    rows.append(_data_row(initials="HK", sname="GAMMA INC",
                          fmcid="FMC-003", fut_value=-200.0))
    rows.append(_data_row(sort1="Total", sname="Subtotal"))
    return rows


# ─── normalize_person ────────────────────────────────────────────────────────

@pytest.mark.parametrize("value,expected", [
    (None, BLANK_PERSON_KEY),
    ("", BLANK_PERSON_KEY),
    ("   ", BLANK_PERSON_KEY),
    ("IS", "IS"),
    ("  HK  ", "HK"),
    (123, "123"),
])
def test_normalize_person_variants(value, expected):
    assert normalize_person(value) == expected


# ─── resolve_person_override ─────────────────────────────────────────────────

def test_resolve_person_override_passthrough_when_no_match():
    assert resolve_person_override("Some Random Stock", "IS") == "IS"


def test_resolve_person_override_blank_security_returns_default():
    assert resolve_person_override(None, "IS") == "IS"
    assert resolve_person_override("", "  HK ") == "HK"


@pytest.mark.parametrize("sname,expected", [
    ("VICTORY GIANT TECH", "KX"),
    ("victory giant tech", "KX"),
    ("  Shanghai Xizhi Tec ", "HK"),
    ("GSCBIHKT CFD GOLDMAN", "HK"),
])
def test_resolve_person_override_known_securities(sname, expected):
    overrides = {
        "VICTORY GIANT TECH": "KX",
        "SHANGHAI XIZHI TEC": "HK",
        "GSCBIHKT CFD GOLDMAN": "HK",
    }
    assert resolve_person_override(sname, "IS", overrides) == expected


# ─── ordered_people ──────────────────────────────────────────────────────────

def test_ordered_people_first_seen_with_blank_last():
    rows = [
        StockRow(8, "IS", "IS", 1.0, 0.0),
        StockRow(9, "HK", "HK", 1.0, 0.0),
        StockRow(10, BLANK_PERSON_KEY, BLANK_PERSON_KEY, 1.0, 0.0),
        StockRow(11, "IS", "IS", 1.0, 0.0),
        StockRow(12, "KX", "KX", 1.0, 0.0),
    ]
    assert ordered_people(rows) == ["IS", "HK", "KX", BLANK_PERSON_KEY]


def test_ordered_people_no_blank_bucket_when_absent():
    rows = [StockRow(8, "IS", "IS", 1.0, 0.0)]
    assert ordered_people(rows) == ["IS"]


# ─── detail_sheet_name ───────────────────────────────────────────────────────

def test_detail_sheet_name_unassigned_for_blank():
    assert detail_sheet_name(BLANK_PERSON_KEY) == "Unassigned"


def test_detail_sheet_name_truncates_to_31_chars():
    name = "A" * 50
    assert len(detail_sheet_name(name)) == 31


# ─── collect_stock_rows ──────────────────────────────────────────────────────

def test_collect_stock_rows_carries_initials_and_skips_subtotals():
    rows = collect_stock_rows(_sample_val_rows())
    assert [r.source_row for r in rows] == [8, 9, 10]
    assert [r.person for r in rows] == ["IS", "IS", "HK"]
    assert [r.j_val for r in rows] == [1000.0, 500.0, -200.0]


def test_collect_stock_rows_applies_sname_overrides():
    rows = [_blank_row() for _ in range(7)]
    rows.append(_data_row(sort1="A", initials="IS",
                          sname="VICTORY GIANT TECH",
                          fmcid="FMC-X", fut_value=10.0))
    out = collect_stock_rows(rows, sname_overrides={"VICTORY GIANT TECH": "KX"})
    assert len(out) == 1
    assert out[0].person == "KX"
    assert out[0].raw_person == "IS"


def test_collect_stock_rows_skips_rows_outside_section_a():
    rows = [_blank_row() for _ in range(7)]
    rows.append(_data_row(sort1="B", initials="IS", fmcid="X", fut_value=1.0))
    assert collect_stock_rows(rows) == []


def test_collect_stock_rows_requires_numeric_fut_value():
    rows = [_blank_row() for _ in range(7)]
    rows.append(_data_row(sort1="A", initials="IS", fmcid="X", fut_value="not a number"))
    assert collect_stock_rows(rows) == []


# ─── accumulate_by_person ────────────────────────────────────────────────────

def test_accumulate_by_person_splits_long_short():
    rows = [
        StockRow(8, "IS", "IS",  100.0,  10.0),
        StockRow(9, "IS", "IS", -50.0,   -5.0),
        StockRow(10, "HK", "HK",  20.0,    2.0),
    ]
    acc = accumulate_by_person(rows, ["IS", "HK"])
    assert acc["IS"].long_j == 100.0 and acc["IS"].short_j == -50.0
    assert acc["IS"].long_t == 10.0 and acc["IS"].short_t == -5.0
    assert acc["HK"].long_j == 20.0 and acc["HK"].short_j == 0.0


# ─── extend_val_rows_with_lookup ─────────────────────────────────────────────

def test_extend_val_rows_writes_headers_at_row_7():
    rows = _sample_val_rows()
    extended = extend_val_rows_with_lookup(rows, {})
    assert extended[6][12] == "Realised P&L (EUR)"
    assert extended[6][19] == "Total P&L (%)"


def test_extend_val_rows_appends_lookup_for_data_rows():
    rows = _sample_val_rows()
    pl_lookup = {"FMC-001": [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0]}
    extended = extend_val_rows_with_lookup(rows, pl_lookup)
    # row 8 (index 7) FMCID = FMC-001 → trailing eight values
    assert extended[7][12:20] == [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0]
    # row 9 has no entry — all None
    assert extended[8][12:20] == [None] * 8


def test_extend_val_rows_pads_short_lookup_to_eight():
    rows = _sample_val_rows()
    pl_lookup = {"FMC-001": [1.0, 2.0]}
    extended = extend_val_rows_with_lookup(rows, pl_lookup)
    assert extended[7][12:20] == [1.0, 2.0, None, None, None, None, None, None]


# ─── build_valuation_workbook (end-to-end pure write) ────────────────────────

def _pl_lookup_sample() -> dict[str, list[float]]:
    return {
        "FMC-001": [10.0, 0.01, 20.0, 0.02, 5.0, 0.005, 35.0, 0.035],
        "FMC-002": [5.0, 0.005, 10.0, 0.01, 2.0, 0.002, 17.0, 0.017],
        "FMC-003": [-2.0, -0.002, -3.0, -0.003, 0.0, 0.0, -5.0, -0.005],
    }


def test_build_workbook_sheet_order_summary_then_valuation_then_people(tmp_path: Path):
    out = tmp_path / "v.xlsx"
    build_valuation_workbook(
        out_path=out, val_rows=_sample_val_rows(),
        pl_lookup=_pl_lookup_sample(),
    )
    wb = load_workbook(out)
    assert wb.sheetnames[0] == SHEET_SUMMARY
    assert wb.sheetnames[1] == SHEET_VALUATION_A
    # Person tabs follow in first-seen order: IS then HK.
    assert wb.sheetnames[2:] == ["IS", "HK"]


def test_build_workbook_summary_per_person_totals(tmp_path: Path):
    out = tmp_path / "v.xlsx"
    build_valuation_workbook(
        out_path=out, val_rows=_sample_val_rows(),
        pl_lookup=_pl_lookup_sample(),
    )
    wb = load_workbook(out)
    summary = wb[SHEET_SUMMARY]
    # Headers at row 6, data starts row 7. IS first.
    assert summary.cell(7, 1).value == "IS"
    assert summary.cell(7, 2).value == 1500.0   # 1000 + 500 long
    assert summary.cell(7, 4).value == 0.0      # no IS short
    # HK row 8: short only
    assert summary.cell(8, 1).value == "HK"
    assert summary.cell(8, 4).value == -200.0
    # TOTAL row at row 9
    assert summary.cell(9, 1).value == "TOTAL"
    assert summary.cell(9, 2).value == 1500.0
    assert summary.cell(9, 4).value == -200.0
    assert summary.cell(9, 6).value == 1300.0   # net


def test_build_workbook_valuation_a_has_mt_headers_and_format(tmp_path: Path):
    out = tmp_path / "v.xlsx"
    build_valuation_workbook(
        out_path=out, val_rows=_sample_val_rows(),
        pl_lookup=_pl_lookup_sample(),
    )
    wb = load_workbook(out)
    val = wb[SHEET_VALUATION_A]
    assert val.cell(7, 13).value == "Realised P&L (EUR)"
    assert val.cell(7, 20).value == "Total P&L (%)"
    # Row 8 has the FMC-001 lookup applied.
    assert val.cell(8, 13).value == 10.0
    # Number format on a percentage column (col 14 = N).
    assert "%" in val.cell(8, 14).number_format
    # B3 carries the explicit fund label.
    assert val["B3"].value == "OAKS Emerging and Frontier Fund"


def test_build_workbook_detail_sheet_filters_by_person(tmp_path: Path):
    out = tmp_path / "v.xlsx"
    build_valuation_workbook(
        out_path=out, val_rows=_sample_val_rows(),
        pl_lookup=_pl_lookup_sample(),
    )
    wb = load_workbook(out)
    is_ws = wb["IS"]
    # Header row 7, then 2 IS rows.
    assert is_ws.cell(8, 2).value == "ALPHA CORP"
    assert is_ws.cell(9, 2).value == "BETA LTD"
    # No 3rd data row.
    assert is_ws.cell(10, 2).value is None
    hk_ws = wb["HK"]
    assert hk_ws.cell(8, 2).value == "GAMMA INC"


def test_build_workbook_atomic_save_leaves_no_tmp(tmp_path: Path):
    out = tmp_path / "v.xlsx"
    build_valuation_workbook(
        out_path=out, val_rows=_sample_val_rows(),
        pl_lookup=_pl_lookup_sample(),
    )
    leftovers = list(tmp_path.glob("v.*.xlsx.tmp"))
    assert leftovers == []


def test_build_workbook_blank_bucket_becomes_unassigned_tab(tmp_path: Path):
    rows = [_blank_row() for _ in range(7)]
    rows.append(_data_row(sort1="A", initials=None, sname="ORPHAN",
                          fmcid="FMC-Z", fut_value=1.0))
    out = tmp_path / "v.xlsx"
    build_valuation_workbook(out_path=out, val_rows=rows, pl_lookup={})
    wb = load_workbook(out)
    assert wb.sheetnames[-1] == "Unassigned"


# ─── COM gating (AST scan: no top-level win32com import) ─────────────────────

def test_valuation_module_does_not_import_win32com_at_top_level():
    src = Path(V.__file__).read_text(encoding="utf-8")
    tree = ast.parse(src)
    for node in tree.body:   # only top-level
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            mod = getattr(node, "module", None)
            names = [a.name for a in node.names] if isinstance(node, ast.Import) else []
            assert mod != "win32com" and "win32com" not in names, (
                "win32com must not be imported at module top level "
                "(use a function-local import instead)"
            )
            assert mod != "pythoncom" and "pythoncom" not in names, (
                "pythoncom must not be imported at module top level"
            )


# ─── close_open_excel_workbook (graceful when pywin32 missing/COM fails) ─────

def test_close_open_excel_workbook_returns_false_when_pywin32_unavailable(monkeypatch):
    """If win32com import fails, the function must return False, not raise."""
    import builtins
    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name in ("win32com", "win32com.client", "pythoncom"):
            raise ImportError("simulated")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    # Drop any cached modules so the function-local import re-executes.
    import sys
    for mod_name in list(sys.modules):
        if mod_name.startswith("win32com") or mod_name == "pythoncom":
            sys.modules.pop(mod_name, None)
    assert close_open_excel_workbook("nonexistent.xlsx") is False


# ─── run_valuation (orchestration with injected fakes) ───────────────────────

def test_run_valuation_uses_injected_extractor_and_closer(tmp_path: Path):
    calls = {"closer": 0, "extractor": 0}

    def fake_closer(path):
        calls["closer"] += 1
        return False

    def fake_extractor(hp_val_path, *, target_pname):
        calls["extractor"] += 1
        assert target_pname == "OAKS Emerging and Frontier Fund"
        return _sample_val_rows()

    out = tmp_path / "v.xlsx"
    result = run_valuation(
        source=tmp_path / "HP_VAL.xlsm",
        out_path=out,
        pl_lookup=_pl_lookup_sample(),
        extractor=fake_extractor,
        closer=fake_closer,
    )
    assert isinstance(result, ValuationResult)
    assert result.out_path == out.resolve()
    assert result.val_row_count == 11   # 7 padding + 3 stock + 1 subtotal
    assert result.person_tab_count == 2
    assert calls == {"closer": 1, "extractor": 1}
    assert out.exists()


def test_run_valuation_propagates_extractor_error(tmp_path: Path):
    def boom(hp_val_path, *, target_pname):
        raise FileNotFoundError("hp_val missing")

    with pytest.raises(FileNotFoundError):
        run_valuation(
            source=tmp_path / "missing.xlsm",
            out_path=tmp_path / "v.xlsx",
            pl_lookup={},
            extractor=boom,
            closer=lambda p: False,
        )
