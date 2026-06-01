"""Tests for compute.fx — FX conversion and snapshot resolution.

These tests pin the structural fixes for BUG-1 (silent fallback EUR/USD) and
BUG-2 (FX leaking via globals). If anyone re-introduces a fallback constant
or a module-level cache, ``test_no_fallback_constant_in_source`` and
``test_resolve_eur_usd_missing_raises`` will fail.
"""

from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
import pytest

from oefof_pl.compute import fx
from oefof_pl.exceptions import FXInvalidError, FXMissingError, SchemaError


# ─── local_to_eur_at: arithmetic ─────────────────────────────────────────────

def test_eur_passes_through():
    assert fx.local_to_eur_at(100.0, "EUR", xrate=1.23, eur_usd=0.86) == 100.0


def test_usd_uses_eur_usd():
    # 100 USD * 0.86 EUR/USD = 86 EUR
    assert fx.local_to_eur_at(100.0, "USD", xrate=1.0, eur_usd=0.86) == pytest.approx(86.0)


def test_local_uses_xrate_then_eur_usd():
    # 15000 JPY / (150 JPY/USD) = 100 USD; 100 * 0.86 = 86 EUR
    got = fx.local_to_eur_at(15000.0, "JPY", xrate=150.0, eur_usd=0.86)
    assert got == pytest.approx(86.0)


def test_ccy_case_insensitive():
    a = fx.local_to_eur_at(100.0, "usd", xrate=1.0, eur_usd=0.9)
    b = fx.local_to_eur_at(100.0, "USD", xrate=1.0, eur_usd=0.9)
    assert a == b


def test_zero_amount_returns_zero_even_without_xrate():
    # Should not require a valid xrate when the amount is zero.
    assert fx.local_to_eur_at(0.0, "JPY", xrate=float("nan"), eur_usd=0.86) == 0.0


def test_nan_amount_returns_zero():
    assert fx.local_to_eur_at(float("nan"), "USD", xrate=1.0, eur_usd=0.86) == 0.0


# ─── local_to_eur_at: validation ────────────────────────────────────────────

@pytest.mark.parametrize("bad", [0, -0.5, float("nan"), float("inf"), -float("inf"), None, "x"])
def test_invalid_eur_usd_raises(bad):
    with pytest.raises(FXInvalidError):
        fx.local_to_eur_at(100.0, "USD", xrate=1.0, eur_usd=bad)


@pytest.mark.parametrize("bad", [0, -1, float("nan"), float("inf")])
def test_invalid_xrate_raises_for_non_eur_usd(bad):
    with pytest.raises(FXInvalidError):
        fx.local_to_eur_at(100.0, "JPY", xrate=bad, eur_usd=0.86)


def test_invalid_xrate_does_not_raise_for_eur():
    # EUR path doesn't touch xrate at all.
    assert fx.local_to_eur_at(100.0, "EUR", xrate=float("nan"), eur_usd=0.86) == 100.0


def test_invalid_xrate_does_not_raise_for_usd():
    assert fx.local_to_eur_at(100.0, "USD", xrate=float("nan"), eur_usd=0.86) == pytest.approx(86.0)


# ─── build_fx_table ──────────────────────────────────────────────────────────

def _mini_port(rows):
    return pd.DataFrame(rows)


def test_build_fx_table_basic():
    df = _mini_port([
        {"CCY": "EUR", "XRATE": 0.86},
        {"CCY": "JPY", "XRATE": 150.0},
        {"CCY": "GBP", "XRATE": 0.78},
    ])
    tbl = fx.build_fx_table(df)
    assert tbl["EUR"] == pytest.approx(0.86)
    assert tbl["JPY"] == pytest.approx(150.0)
    assert tbl["GBP"] == pytest.approx(0.78)
    assert tbl["USD"] == 1.0  # always present


def test_build_fx_table_first_value_wins_for_dupes():
    df = _mini_port([
        {"CCY": "JPY", "XRATE": 150.0},
        {"CCY": "JPY", "XRATE": 151.0},
    ])
    tbl = fx.build_fx_table(df)
    assert tbl["JPY"] == 150.0


def test_build_fx_table_skips_nan():
    df = _mini_port([
        {"CCY": "JPY", "XRATE": float("nan")},
        {"CCY": "JPY", "XRATE": 150.0},
    ])
    assert fx.build_fx_table(df)["JPY"] == 150.0


def test_build_fx_table_invalid_xrate_raises():
    df = _mini_port([{"CCY": "JPY", "XRATE": -150.0}])
    with pytest.raises(FXInvalidError):
        fx.build_fx_table(df)


def test_build_fx_table_missing_columns_raises():
    df = pd.DataFrame({"CCY": ["EUR"]})  # no XRATE
    with pytest.raises(SchemaError):
        fx.build_fx_table(df)


# ─── resolve_eur_usd ─────────────────────────────────────────────────────────

def test_resolve_eur_usd_basic():
    df = _mini_port([
        {"CCY": "JPY", "XRATE": 150.0},
        {"CCY": "EUR", "XRATE": 0.868},
    ])
    assert fx.resolve_eur_usd(df) == pytest.approx(0.868)


def test_resolve_eur_usd_case_insensitive_and_padding():
    df = _mini_port([{"CCY": " eur ", "XRATE": 0.9}])
    assert fx.resolve_eur_usd(df) == pytest.approx(0.9)


def test_resolve_eur_usd_missing_raises():
    """BUG-1 regression: no fallback. Missing EUR must raise, not return 0.868018."""
    df = _mini_port([{"CCY": "USD", "XRATE": 1.0}])
    with pytest.raises(FXMissingError):
        fx.resolve_eur_usd(df)


def test_resolve_eur_usd_invalid_raises():
    df = _mini_port([{"CCY": "EUR", "XRATE": 0.0}])
    with pytest.raises(FXInvalidError):
        fx.resolve_eur_usd(df)


def test_resolve_eur_usd_nan_treated_as_missing():
    df = _mini_port([{"CCY": "EUR", "XRATE": float("nan")}])
    with pytest.raises(FXMissingError):
        fx.resolve_eur_usd(df)


def test_resolve_eur_usd_schema_error():
    df = pd.DataFrame({"CCY": ["EUR"]})
    with pytest.raises(SchemaError):
        fx.resolve_eur_usd(df)


# ─── structural regressions (AST-based, ignores docstrings/comments) ────────

import ast as _ast

_FX_SOURCE = Path(fx.__file__).read_text(encoding="utf-8")
_FX_TREE = _ast.parse(_FX_SOURCE)


def _all_numeric_constants(tree: _ast.AST) -> list[float]:
    out: list[float] = []
    for node in _ast.walk(tree):
        if isinstance(node, _ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            out.append(float(node.value))
    return out


def _module_level_names(tree: _ast.Module) -> set[str]:
    names: set[str] = set()
    for node in tree.body:
        if isinstance(node, _ast.Assign):
            for tgt in node.targets:
                if isinstance(tgt, _ast.Name):
                    names.add(tgt.id)
        elif isinstance(node, _ast.AnnAssign) and isinstance(node.target, _ast.Name):
            names.add(node.target.id)
    return names


def test_no_fallback_constant_in_code():
    """BUG-1 regression: the legacy magic EUR/USD fallbacks must not appear as code literals."""
    consts = _all_numeric_constants(_FX_TREE)
    assert 0.868018 not in consts
    assert 0.851462 not in consts


def test_no_module_level_fx_state():
    """BUG-2 regression: no module-level mutable EUR/USD globals."""
    forbidden = {"EUR_USD_END", "EUR_USD_START", "EUR_USD", "fx_start", "fx_end", "FX_START", "FX_END"}
    leaked = forbidden & _module_level_names(_FX_TREE)
    assert not leaked, f"forbidden module-level names in fx.py: {leaked}"


def test_local_to_eur_signature_requires_explicit_args():
    """Caller cannot accidentally rely on a default — every rate is explicit."""
    import inspect
    sig = inspect.signature(fx.local_to_eur_at)
    for name in ("amount", "ccy", "xrate", "eur_usd"):
        assert sig.parameters[name].default is inspect.Parameter.empty, (
            f"{name} must have no default (would re-introduce BUG-2)"
        )
