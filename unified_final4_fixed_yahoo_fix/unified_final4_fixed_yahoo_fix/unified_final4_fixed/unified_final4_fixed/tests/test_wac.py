"""Tests for compute.wac — weighted-average cost in local currency.

Pins the BUG-5 structural fix: income trades (T='I') are filtered out
before the WAC loop and cannot influence units or cost basis.
"""

from __future__ import annotations

import pandas as pd
import pytest

from oefof_pl.compute.wac import WacResult, compute_wac
from oefof_pl.exceptions import SchemaError


def _trades(rows):
    df = pd.DataFrame(rows, columns=["T", "CDATE", "UNITS", "GROSSPRICE_LOCAL"])
    df["CDATE"] = pd.to_datetime(df["CDATE"])
    return df.sort_values("CDATE").reset_index(drop=True)


# ─── trivial cases ──────────────────────────────────────────────────────────

def test_no_trades_no_starting_position():
    r = compute_wac(0.0, 0.0, _trades([]))
    assert r == WacResult(0, 0, 0, 0, 0, 0, 0, None)


def test_no_trades_with_starting_position_passes_through():
    r = compute_wac(100.0, 5000.0, _trades([]))
    assert r.ending_units == 100.0
    assert r.ending_cost_local == 5000.0
    assert r.realised_local == 0.0
    assert r.units_bought == 0
    assert r.units_sold == 0
    assert r.last_entry_date is None


# ─── buy-only ───────────────────────────────────────────────────────────────

def test_single_buy_from_flat():
    r = compute_wac(0, 0, _trades([("P", "2026-01-10", 100, 50.0)]))
    assert r.ending_units == 100
    assert r.ending_cost_local == pytest.approx(5000.0)
    assert r.avg_buy_price_local == pytest.approx(50.0)
    assert r.units_bought == 100
    assert r.realised_local == 0


def test_two_buys_average_correctly():
    r = compute_wac(0, 0, _trades([
        ("P", "2026-01-10", 100, 50.0),
        ("P", "2026-01-20", 100, 70.0),
    ]))
    assert r.ending_units == 200
    assert r.ending_cost_local == pytest.approx(12000.0)
    # WAC = (50*100 + 70*100) / 200 = 60
    assert r.ending_cost_local / r.ending_units == pytest.approx(60.0)
    assert r.avg_buy_price_local == pytest.approx(60.0)


# ─── buy-then-sell, WAC realisation ─────────────────────────────────────────

def test_buy_then_full_sell_realises_total_pl():
    # 100 @ 50 → sell 100 @ 75 → realised 2500, end flat
    r = compute_wac(0, 0, _trades([
        ("P", "2026-01-10", 100, 50.0),
        ("S", "2026-01-20", 100, 75.0),
    ]))
    assert r.ending_units == 0
    assert r.ending_cost_local == pytest.approx(0.0)
    assert r.realised_local == pytest.approx(2500.0)
    assert r.avg_buy_price_local == pytest.approx(50.0)
    assert r.avg_sell_price_local == pytest.approx(75.0)


def test_partial_sell_at_wac_keeps_remaining_at_wac():
    # 100 @ 50 then 100 @ 70 → WAC 60 → sell 50 @ 80 → realised = 50*(80-60)=1000
    r = compute_wac(0, 0, _trades([
        ("P", "2026-01-10", 100, 50.0),
        ("P", "2026-01-15", 100, 70.0),
        ("S", "2026-01-20",  50, 80.0),
    ]))
    assert r.ending_units == 150
    # remaining cost = 60 * 150
    assert r.ending_cost_local == pytest.approx(60.0 * 150)
    assert r.realised_local == pytest.approx(1000.0)


def test_starting_position_used_as_initial_wac():
    # Start: 100 units @ avg 40 (cost 4000) → buy 100 @ 60 → WAC = (4000+6000)/200 = 50
    # → sell 50 @ 70 → realised = 50 * (70-50) = 1000
    r = compute_wac(100.0, 4000.0, _trades([
        ("P", "2026-01-10", 100, 60.0),
        ("S", "2026-01-15",  50, 70.0),
    ]))
    assert r.realised_local == pytest.approx(1000.0)
    assert r.ending_units == 150
    assert r.ending_cost_local == pytest.approx(50.0 * 150)


# ─── ordering ───────────────────────────────────────────────────────────────

def test_chronological_order_is_respected_via_input_sort():
    # Identical inputs in different order must give the same result IF caller
    # has sorted (which compute_wac trusts). Sort happens in _trades helper.
    a = compute_wac(0, 0, _trades([
        ("P", "2026-01-10", 100, 50.0),
        ("S", "2026-01-15",  50, 80.0),
    ]))
    b = compute_wac(0, 0, _trades([
        ("S", "2026-01-15",  50, 80.0),  # later in source order
        ("P", "2026-01-10", 100, 50.0),  # but earlier in CDATE
    ]))
    assert a == b


# ─── BUG-5 regression: income trades ignored ────────────────────────────────

def test_income_trades_ignored_completely():
    """BUG-5 structural fix. T='I' rows must not change units, cost, or
    avg_buy_price. If they did, the dividend on a 1m-share holding would
    show up as 1m extra units bought at the per-share dividend price."""
    with_income = compute_wac(0, 0, _trades([
        ("P", "2026-01-10", 100, 50.0),
        ("I", "2026-01-12", 100, 1.50),     # huge dividend trap
        ("I", "2026-01-13",   1, 99999.0),  # absurd trap
        ("S", "2026-01-20", 100, 75.0),
    ]))
    without_income = compute_wac(0, 0, _trades([
        ("P", "2026-01-10", 100, 50.0),
        ("S", "2026-01-20", 100, 75.0),
    ]))
    assert with_income == without_income


# ─── short positions ────────────────────────────────────────────────────────

def test_open_short_from_flat():
    r = compute_wac(0, 0, _trades([("S", "2026-01-10", 100, 50.0)]))
    assert r.ending_units == -100
    assert r.ending_cost_local == pytest.approx(-5000.0)
    assert r.realised_local == 0
    assert r.units_sold == 100


def test_short_then_cover_realises_pl():
    # Open short 100 @ 50 → cover (buy) 100 @ 40 → P&L = +1000 (sold high, bought low).
    # In WAC representation: starting_units = -100 (post-sell), cost = -5000.
    # Then a P at 40 increases cost by +4000 → cost=-1000, units=0.
    # That residual −1000 is the realised P&L the model should capture.
    # Note: WAC algorithm here books realised on S, not on cover-P. The
    # cover pushes units back to flat; remaining negative cost equals
    # accumulated profit but isn't booked as realised by compute_wac itself.
    # compute.pl reconciles by treating ending_cost_local at flat as realised.
    # This test pins current behaviour so any future change is intentional.
    r = compute_wac(0, 0, _trades([
        ("S", "2026-01-10", 100, 50.0),
        ("P", "2026-01-15", 100, 40.0),
    ]))
    assert r.ending_units == 0
    # realised inside compute_wac is from S-leg only (sold into flat → 0)
    assert r.realised_local == 0
    # P&L is captured in residual cost: -5000 + 4000 = -1000
    assert r.ending_cost_local == pytest.approx(-1000.0)


# ─── validation ─────────────────────────────────────────────────────────────

def test_schema_error_on_missing_columns():
    bad = pd.DataFrame({"T": ["P"], "UNITS": [100]})  # no CDATE, no price
    with pytest.raises(SchemaError):
        compute_wac(0, 0, bad)


def test_unknown_trade_type_raises():
    df = _trades([("X", "2026-01-10", 100, 50.0)])
    with pytest.raises(ValueError, match="Unknown trade type"):
        compute_wac(0, 0, df)


def test_nan_price_treated_as_zero():
    # A NaN price on a P trade gives 0 cost and 0 wac contribution.
    r = compute_wac(0, 0, _trades([("P", "2026-01-10", 100, float("nan"))]))
    assert r.ending_units == 100
    assert r.ending_cost_local == 0


def test_reentry_after_full_exit_sets_last_entry_date_and_resets_avg_buy():
    r = compute_wac(0.0, 0.0, _trades([
        ("P", "2026-01-10", 100, 50.0),
        ("S", "2026-01-11", 100, 60.0),
        ("P", "2026-01-20", 10, 80.0),
    ]))
    assert r.ending_units == pytest.approx(10.0)
    assert r.avg_buy_price_local == pytest.approx(80.0)
    assert r.last_entry_date == pd.Timestamp("2026-01-20")
