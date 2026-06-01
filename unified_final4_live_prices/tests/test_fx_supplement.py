"""Regression tests for build_fx_supplement — the NOK FX fallback fix.

Covers the scenario where a position (e.g. Okeanis Eco Tankers, NOK) has
NULL XRATE in vw_RPT_VAL at Dec-31 but a valid XRATE in tH_VAL for that
same date. Without the supplement, the pipeline crashes at compute-positions
with: "FXInvalidError: xrate for NOK must be finite and >0, got nan".

Run with:
    pytest tests/test_fx_supplement.py -v
"""
from __future__ import annotations

import math
import pandas as pd
import pytest

from oefof_pl.compute.fx import build_fx_supplement


def _df(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(rows)


# ── basic supplement behaviour ────────────────────────────────────────────────

def test_fills_missing_nok_from_th_val():
    """Core regression: NOK absent from base_fx gets filled by extra_df."""
    base = {"USD": 1.0, "EUR": 0.859, "HKD": 7.78}
    extra = _df([
        {"CCY": "NOK", "XRATE": 10.52},
        {"CCY": "HKD", "XRATE": 7.80},   # already in base — must NOT override
    ])
    result = build_fx_supplement(extra, base)

    assert "NOK" in result
    assert math.isclose(result["NOK"], 10.52)
    # HKD must keep the base value (primary snapshot wins)
    assert math.isclose(result["HKD"], 7.78)


def test_base_is_not_mutated():
    base = {"USD": 1.0}
    extra = _df([{"CCY": "NOK", "XRATE": 10.52}])
    build_fx_supplement(extra, base)
    assert "NOK" not in base  # original dict unchanged


def test_returns_copy_of_base_when_extra_is_none():
    base = {"USD": 1.0, "EUR": 0.859}
    result = build_fx_supplement(None, base)
    assert result == base
    assert result is not base  # must be a copy


def test_returns_copy_of_base_when_extra_is_empty():
    base = {"USD": 1.0}
    result = build_fx_supplement(pd.DataFrame(), base)
    assert result == base


def test_returns_copy_of_base_when_extra_lacks_ccy_or_xrate():
    base = {"USD": 1.0}
    extra = _df([{"ISIN": "MHY641771016", "UNITS": 5000}])  # no CCY/XRATE cols
    result = build_fx_supplement(extra, base)
    assert result == base


def test_skips_nan_and_zero_xrate_in_extra():
    base = {"USD": 1.0}
    extra = _df([
        {"CCY": "NOK", "XRATE": float("nan")},
        {"CCY": "SEK", "XRATE": 0.0},
        {"CCY": "DKK", "XRATE": -10.0},
    ])
    result = build_fx_supplement(extra, base)
    # None of the invalid rows should be added
    assert "NOK" not in result
    assert "SEK" not in result
    assert "DKK" not in result


def test_first_valid_row_per_ccy_wins():
    """If the same CCY appears twice in extra, the first valid rate is used."""
    base = {}
    extra = _df([
        {"CCY": "NOK", "XRATE": 10.52},
        {"CCY": "NOK", "XRATE": 10.99},
    ])
    result = build_fx_supplement(extra, base)
    assert math.isclose(result["NOK"], 10.52)


def test_multiple_missing_ccys_filled():
    base = {"USD": 1.0, "EUR": 0.859}
    extra = _df([
        {"CCY": "NOK", "XRATE": 10.52},
        {"CCY": "IDR", "XRATE": 16350.0},
        {"CCY": "VND", "XRATE": 25400.0},
    ])
    result = build_fx_supplement(extra, base)
    assert set(result.keys()) == {"USD", "EUR", "NOK", "IDR", "VND"}


# ── priority contract ────────────────────────────────────────────────────────

def test_pipeline_merge_priority():
    """Full priority chain: end > start > supplement — matches pipeline merge.

    In compute_positions:
        fx_end_effective = {**dict(fx_start), **dict(fx_end)}
    And in pipeline stage 7:
        fx_start = build_fx_supplement(th_val_start_df, fx_start)

    So the priority is: fx_end > fx_start (vw_RPT_VAL) > th_val_start supplement.
    """
    # Simulate: supplement adds NOK at 10.40 (th_val Dec-31 rate)
    base_start = {}
    extra_start = _df([{"CCY": "NOK", "XRATE": 10.40}])
    fx_start = build_fx_supplement(extra_start, base_start)
    assert math.isclose(fx_start["NOK"], 10.40)

    # fx_end has no NOK (position was closed)
    fx_end = {"USD": 1.0, "EUR": 0.859}

    # compute_positions merge: {**fx_start, **fx_end}
    fx_end_effective = {**fx_start, **fx_end}
    # NOK comes from fx_start (the supplemented value)
    assert math.isclose(fx_end_effective["NOK"], 10.40)
