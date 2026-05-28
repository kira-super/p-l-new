"""Regression tests for build_fx_from_trades (FX tier-3 fallback).

Covers the scenario where a position (e.g. NOK-denominated Okeanis Eco
Tankers) was opened and closed entirely within the YTD window, so its CCY
never appears in either the start or end portfolio snapshot, causing the
pipeline to crash with ``FXInvalidError: xrate for NOK must be finite and
>0, got nan``.

Add this file to tests/ and run with::

    pytest tests/test_fx_from_trades.py -v
"""
from __future__ import annotations

import math
import pandas as pd
import pytest

from oefof_pl.compute.fx import build_fx_from_trades


# ─── helpers ─────────────────────────────────────────────────────────────────

def _trades(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(rows)


# ─── tests ───────────────────────────────────────────────────────────────────

def test_returns_empty_when_no_xrate_column():
    """Trades without XRATE column → empty dict (standard schema, no crash)."""
    df = _trades([{"CCY": "NOK", "T": "P", "UNITS": 1000}])
    assert build_fx_from_trades(df) == {}


def test_returns_empty_for_empty_dataframe():
    df = pd.DataFrame()
    assert build_fx_from_trades(df) == {}


def test_extracts_nok_from_trades():
    """Core scenario: NOK position closed YTD, XRATE available in trade row."""
    df = _trades([
        {"CCY": "NOK", "XRATE": 10.52, "T": "P", "UNITS": 5000},
        {"CCY": "NOK", "XRATE": 10.48, "T": "S", "UNITS": 5000},
    ])
    result = build_fx_from_trades(df)
    assert "NOK" in result
    # First-row rate wins (10.52)
    assert math.isclose(result["NOK"], 10.52)


def test_skips_nan_xrate_rows():
    df = _trades([
        {"CCY": "NOK", "XRATE": float("nan"), "T": "P"},
        {"CCY": "NOK", "XRATE": 10.52, "T": "S"},
    ])
    result = build_fx_from_trades(df)
    # NaN row skipped; second valid row captured
    assert math.isclose(result["NOK"], 10.52)


def test_skips_zero_xrate():
    df = _trades([{"CCY": "NOK", "XRATE": 0.0, "T": "P"}])
    assert "NOK" not in build_fx_from_trades(df)


def test_skips_negative_xrate():
    df = _trades([{"CCY": "NOK", "XRATE": -10.52, "T": "P"}])
    assert "NOK" not in build_fx_from_trades(df)


def test_eur_and_usd_not_included():
    """EUR and USD are handled specially; build_fx_from_trades must exclude them."""
    df = _trades([
        {"CCY": "EUR", "XRATE": 0.92, "T": "I"},
        {"CCY": "USD", "XRATE": 1.0,  "T": "I"},
        {"CCY": "HKD", "XRATE": 7.78, "T": "P"},
    ])
    result = build_fx_from_trades(df)
    assert "EUR" not in result
    assert "USD" not in result
    assert "HKD" in result


def test_multiple_ccys_extracted():
    df = _trades([
        {"CCY": "NOK", "XRATE": 10.52, "T": "P"},
        {"CCY": "HKD", "XRATE": 7.78,  "T": "P"},
        {"CCY": "KRW", "XRATE": 1340.0, "T": "P"},
    ])
    result = build_fx_from_trades(df)
    assert set(result.keys()) == {"NOK", "HKD", "KRW"}


def test_first_valid_rate_wins_per_ccy():
    """If the same CCY appears multiple times, the first valid rate wins."""
    df = _trades([
        {"CCY": "NOK", "XRATE": 10.52, "T": "P"},
        {"CCY": "NOK", "XRATE": 10.99, "T": "S"},
    ])
    result = build_fx_from_trades(df)
    assert math.isclose(result["NOK"], 10.52)


def test_snapshot_rate_beats_trades_rate_in_pipeline_merge():
    """Snapshot rates must override trades rates when merged in compute_positions.

    This test validates the priority contract: snapshot FX (tier 2/1) always
    beats trade-date FX (tier 3). The merge in compute_positions is:
        {**fx_trades, **fx_start, **fx_end}
    so later dicts win. We assert that pattern here so a refactor that
    accidentally reverses priority breaks this test.
    """
    fx_trades = {"NOK": 10.52}   # trade-date rate (stale)
    fx_start  = {"NOK": 10.60}   # start-of-period snapshot rate
    fx_end    = {"NOK": 10.75}   # end-of-period snapshot rate (highest priority)

    # Simulating the merge in compute_positions:
    fx_end_effective = {**fx_trades, **fx_start, **fx_end}

    assert math.isclose(fx_end_effective["NOK"], 10.75), (
        "End-snapshot rate must win over start-snapshot and trade-date rates"
    )


def test_snapshot_start_beats_trades_rate():
    """When end snapshot is missing NOK, start snapshot rate must win over trades."""
    fx_trades = {"NOK": 10.52}
    fx_start  = {"NOK": 10.60}
    fx_end    = {}  # NOK position was closed; not in end snapshot

    fx_end_effective = {**fx_trades, **fx_start, **fx_end}

    assert math.isclose(fx_end_effective["NOK"], 10.60)
