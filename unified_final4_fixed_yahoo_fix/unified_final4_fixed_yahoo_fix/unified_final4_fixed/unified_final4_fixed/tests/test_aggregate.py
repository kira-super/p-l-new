"""Tests for compute.aggregate — portfolio-line aggregation and FUT escalation."""

from __future__ import annotations

import pandas as pd
import pytest

from oefof_pl.compute.aggregate import AggLine, aggregate_portfolio
from oefof_pl.exceptions import SchemaError


def _port(rows):
    """Build a port_df with all PORT_REQUIRED columns; missing fields filled with defaults."""
    defaults = {
        "ISIN": "X", "SNAME": "Stock", "CCY": "EUR", "XRATE": 1.0,
        "UNITS": 0.0, "PTVALUE_EUR": 0.0, "NUCOST_LOCAL": 0.0,
        "LST_PRICE_LOCAL": 0.0, "CAT": "ORD", "LS": "L",
        "PCODE_ORIG": "OEFOF", "VDATE": pd.Timestamp("2026-01-31"),
        "EXCODE1": "DE",
    }
    out = []
    for r in rows:
        row = dict(defaults)
        row.update(r)
        out.append(row)
    if not out:
        return pd.DataFrame(columns=list(defaults.keys()))
    return pd.DataFrame(out)


def test_empty_returns_empty_dict():
    df = _port([])
    assert aggregate_portfolio(df) == {}


def test_schema_error_on_missing_columns():
    bad = pd.DataFrame({"ISIN": ["X"]})
    with pytest.raises(SchemaError):
        aggregate_portfolio(bad)


def test_single_row_passes_through():
    df = _port([{
        "ISIN": "KR7005930003", "SNAME": "SAMSUNG ELEC", "CCY": "KRW",
        "XRATE": 1300.0, "UNITS": 100.0, "PTVALUE_EUR": 7_000_000.0,
        "NUCOST_LOCAL": 50_000.0, "LST_PRICE_LOCAL": 70_000.0,
        "CAT": "ORD", "LS": "L", "EXCODE1": "KR",
    }])
    out = aggregate_portfolio(df)
    assert "KR7005930003" in out
    line = out["KR7005930003"]
    assert isinstance(line, AggLine)
    assert line.units == 100.0
    assert line.ptvalue_eur == 7_000_000.0
    assert line.nucost_local == 50_000.0
    assert line.cat == "ORD"
    assert line.country == "KR"
    assert line.n_source_rows == 1


def test_multi_row_sums_and_weighted_avg():
    # Two rows for same ISIN: sum units & ptvalue, weighted avg nucost.
    df = _port([
        {"ISIN": "X", "UNITS": 100, "PTVALUE_EUR": 5000, "NUCOST_LOCAL": 40, "LST_PRICE_LOCAL": 50},
        {"ISIN": "X", "UNITS": 100, "PTVALUE_EUR": 5000, "NUCOST_LOCAL": 60, "LST_PRICE_LOCAL": 50},
    ])
    out = aggregate_portfolio(df)
    line = out["X"]
    assert line.units == 200
    assert line.ptvalue_eur == 10000
    # weighted avg: (40*100 + 60*100) / 200 = 50
    assert line.nucost_local == 50.0
    assert line.n_source_rows == 2


def test_isin_normalised_uppercase_stripped():
    df = _port([{"ISIN": " kr7005930003 "}])
    out = aggregate_portfolio(df)
    assert "KR7005930003" in out
    assert " kr7005930003 " not in out


def test_empty_isin_dropped():
    df = _port([
        {"ISIN": "", "UNITS": 10},
        {"ISIN": "ABC", "UNITS": 5},
    ])
    out = aggregate_portfolio(df)
    assert list(out) == ["ABC"]


# ─── BUG-7 (b): FUT escalation ──────────────────────────────────────────────

def test_cat_fut_escalates_when_any_row_is_fut():
    """Vietnam Dairy hybrid: ORD line + FUT line on same ISIN.
    Mixed groups now split into separate ORD and FUT keys."""
    df = _port([
        {"ISIN": "VN-DAIRY", "CAT": "ORD", "UNITS": 100},
        {"ISIN": "VN-DAIRY", "CAT": "FUT", "UNITS": 50},
    ])
    out = aggregate_portfolio(df)
    assert out["VN-DAIRY"].cat == "ORD"
    assert out["VN-DAIRY:FUT"].cat == "FUT"


def test_cat_fut_escalation_independent_of_row_order():
    df_a = _port([{"ISIN": "X", "CAT": "FUT"}, {"ISIN": "X", "CAT": "ORD"}])
    df_b = _port([{"ISIN": "X", "CAT": "ORD"}, {"ISIN": "X", "CAT": "FUT"}])
    out_a = aggregate_portfolio(df_a)
    out_b = aggregate_portfolio(df_b)
    for out in (out_a, out_b):
        assert out["X"].cat == "ORD"
        assert out["X:FUT"].cat == "FUT"


def test_cat_blank_falls_back_to_ord():
    df = _port([{"ISIN": "X", "CAT": ""}])
    assert aggregate_portfolio(df)["X"].cat == "ORD"


# ─── EUR fields on AggLine carry HiPort's pre-converted PTVALUE/PTCOST ──────

def test_aggline_carries_eur_ptvalue_and_ptcost_only():
    """BUG-11 (post-rename): the snapshot pre-converts PTVALUE and PTCOST
    to fund-base EUR, so AggLine MUST carry them as ``ptvalue_eur`` /
    ``ptcost_eur``. No other EUR field is allowed (no eur-converted unit
    cost, etc.) — those are computed downstream with explicit FX args."""
    fields = {f.name for f in AggLine.__dataclass_fields__.values()}
    eur_fields = {f for f in fields if "eur" in f.lower()}
    assert eur_fields == {"ptvalue_eur", "ptcost_eur"}, (
        f"Unexpected EUR field set on AggLine: {eur_fields}"
    )


def test_aggline_is_frozen():
    """Mutation must raise — no silent overwrites of cost basis etc."""
    line = AggLine(
        isin="X", sname="X", ccy="EUR", ls="L", country="DE", cat="ORD",
        units=0, ptvalue_eur=0, ptcost_eur=0, nucost_local=0,
        lst_price_local=0, xrate=1.0, n_source_rows=1,
    )
    with pytest.raises(Exception):
        line.units = 999  # type: ignore[misc]
