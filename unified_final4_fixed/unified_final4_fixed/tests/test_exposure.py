from __future__ import annotations

import pandas as pd
import pytest

from oefof_pl.compute.exposure import build_snapshot_exposure, build_weighted_exposure


def _history_row(vdate: str, *, isin: str, sname: str, ls: str, ptvalue_eur: float) -> dict:
    return {
        "VDATE": pd.Timestamp(vdate),
        "ISIN": isin,
        "SNAME": sname,
        "LS": ls,
        "CAT": "ORD",
        "PTVALUE_EUR": ptvalue_eur,
    }


def test_build_weighted_exposure_uses_daily_nav_weighting_and_latest_snapshot_date():
    history_df = pd.DataFrame([
        _history_row("2026-05-22", isin="A", sname="A", ls="L", ptvalue_eur=100.0),
        _history_row("2026-05-22", isin="B", sname="B", ls="S", ptvalue_eur=-40.0),
        _history_row("2026-05-23", isin="A", sname="A", ls="L", ptvalue_eur=300.0),
        _history_row("2026-05-23", isin="B", sname="B", ls="S", ptvalue_eur=-60.0),
    ])
    pl_df = pd.DataFrame([
        {"ISIN": "A", "Stock Name": "A", "Analyst": "IS"},
        {"ISIN": "B", "Stock Name": "B", "Analyst": "IS"},
    ])

    block = build_weighted_exposure(history_df, pl_df)

    assert block.as_of == pd.Timestamp("2026-05-23")
    assert block.fund.long_eur == pytest.approx(260.0)
    assert block.fund.short_eur == pytest.approx(56.0)
    assert block.by_analyst["IS"].long_eur == pytest.approx(260.0)
    assert block.by_analyst["IS"].short_eur == pytest.approx(56.0)
    assert block.fund.nav_eur == pytest.approx(150.0)


def test_build_weighted_exposure_weights_by_daily_nav():
    history_df = pd.DataFrame([
        _history_row("2026-05-22", isin="A", sname="A", ls="L", ptvalue_eur=100.0),
        _history_row("2026-05-23", isin="A", sname="A", ls="L", ptvalue_eur=300.0),
    ])
    pl_df = pd.DataFrame([
        {"ISIN": "A", "Stock Name": "A", "Analyst": "IS"},
    ])

    block = build_weighted_exposure(history_df, pl_df)

    assert block.as_of == pd.Timestamp("2026-05-23")
    assert block.fund.long_eur == pytest.approx(250.0)
    assert block.fund.short_eur == pytest.approx(0.0)
    assert block.by_analyst["IS"].long_eur == pytest.approx(250.0)
    assert block.by_analyst["IS"].short_eur == pytest.approx(0.0)
    assert block.fund.nav_eur == pytest.approx(200.0)


def test_build_snapshot_exposure_excludes_non_equity_rows():
    snap_df = pd.DataFrame([
        {
            "ISIN": "A", "SNAME": "Alpha", "LS": "L", "CAT": "ORD", "PTVALUE_EUR": 50.0,
        },
        {
            "ISIN": "B", "SNAME": "Beta", "LS": "S", "CAT": "ORD", "PTVALUE_EUR": -20.0,
        },
        {
            "ISIN": "", "SNAME": "CASH GBP", "LS": "", "CAT": "", "PTVALUE_EUR": 15.0,
        },
        {
            "ISIN": "", "SNAME": "ACCRUALS", "LS": "", "CAT": "", "PTVALUE_EUR": -3.0,
        },
    ])
    pl_df = pd.DataFrame([
        {"ISIN": "A", "Stock Name": "Alpha", "Analyst": "VS"},
        {"ISIN": "B", "Stock Name": "Beta", "Analyst": "VS"},
    ])

    block = build_snapshot_exposure(snap_df, pl_df, as_of=pd.Timestamp("2026-05-28"))

    assert block.fund.long_eur == pytest.approx(50.0)
    assert block.fund.short_eur == pytest.approx(20.0)
    assert block.by_analyst["VS"].long_eur == pytest.approx(50.0)
    assert block.by_analyst["VS"].short_eur == pytest.approx(20.0)
    assert "UNASSIGNED" not in block.by_analyst


def test_build_weighted_exposure_nav_uses_all_rows_but_exposure_equity_only():
    history_df = pd.DataFrame([
        {
            "VDATE": "2026-05-22", "ISIN": "A", "SNAME": "Alpha", "LS": "L", "CAT": "ORD", "PTVALUE_EUR": 100.0,
        },
        {
            "VDATE": "2026-05-23", "ISIN": "A", "SNAME": "Alpha", "LS": "L", "CAT": "ORD", "PTVALUE_EUR": 300.0,
        },
        {
            "VDATE": "2026-05-22", "ISIN": "", "SNAME": "CASH", "LS": "", "CAT": "", "PTVALUE_EUR": 50.0,
        },
        {
            "VDATE": "2026-05-23", "ISIN": "", "SNAME": "CASH", "LS": "", "CAT": "", "PTVALUE_EUR": 50.0,
        },
    ])
    history_df["VDATE"] = pd.to_datetime(history_df["VDATE"])
    pl_df = pd.DataFrame([
        {"ISIN": "A", "Stock Name": "Alpha", "Analyst": "IS"},
    ])

    block = build_weighted_exposure(history_df, pl_df)

    assert block.fund.nav_eur == pytest.approx(250.0)
    assert block.fund.long_eur == pytest.approx(240.0)
    assert block.fund.short_eur == pytest.approx(0.0)
    assert block.by_analyst["IS"].long_eur == pytest.approx(block.fund.long_eur)
    assert "UNASSIGNED" not in block.by_analyst


def test_build_weighted_exposure_analyst_sum_equals_fund_total():
    """Analysts joining mid-period must have zero-filled missing days.

    This preserves the weighted-average invariant that analyst sums equal
    fund totals when all analyst series use the same trading-date universe
    and NAV weights.
    """
    history_df = pd.DataFrame([
        {
            "VDATE": "2026-01-02", "ISIN": "A", "SNAME": "Alpha", "LS": "L", "CAT": "ORD", "PTVALUE_EUR": 100.0,
        },
        {
            "VDATE": "2026-01-03", "ISIN": "A", "SNAME": "Alpha", "LS": "L", "CAT": "ORD", "PTVALUE_EUR": 200.0,
        },
        {
            "VDATE": "2026-01-03", "ISIN": "B", "SNAME": "Beta", "LS": "L", "CAT": "ORD", "PTVALUE_EUR": 50.0,
        },
    ])
    history_df["VDATE"] = pd.to_datetime(history_df["VDATE"])
    pl_df = pd.DataFrame([
        {"ISIN": "A", "Stock Name": "Alpha", "Analyst": "VS"},
        {"ISIN": "B", "Stock Name": "Beta", "Analyst": "IS"},
    ])

    block = build_weighted_exposure(history_df, pl_df)

    analyst_long_sum = sum(em.long_eur for em in block.by_analyst.values())
    analyst_short_sum = sum(em.short_eur for em in block.by_analyst.values())
    assert analyst_long_sum == pytest.approx(block.fund.long_eur)
    assert analyst_short_sum == pytest.approx(block.fund.short_eur)
