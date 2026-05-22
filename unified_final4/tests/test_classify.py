"""Tests for compute.classify — instrument-type resolution.

Pins BUG-9 structurally: the function signature does not accept the
legacy ``ORD`` column, so any caller that tries to pass it gets a TypeError
at import / call time rather than a wrong classification.
"""

from __future__ import annotations

import inspect

import pytest

from oefof_pl.compute.classify import (
    FTSWAP,
    INSTRUMENT_TYPES,
    ORD,
    SWAP,
    classify_instrument,
)


FTSWAP_SET = frozenset({"SX7E INDEX", "GSCBIHKT"})


# ─── happy-path classifications ─────────────────────────────────────────────

def test_ord_equity_classified_ord():
    got = classify_instrument(
        cat="ORD", sname="SAMSUNG ELECTRONICS", isin="KR7005930003",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == ORD


def test_fut_cat_classified_swap():
    got = classify_instrument(
        cat="FUT", sname="HAPAG-LLOYD CFD", isin="DE000HLAG475",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == SWAP


def test_cfd_in_name_classified_swap_even_when_cat_blank():
    got = classify_instrument(
        cat=None, sname="HAPAG-LLOYD AG CFD", isin="DE000HLAG475",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == SWAP


def test_isin_fut_suffix_classified_swap_even_when_cat_and_name_are_blank():
    got = classify_instrument(
        cat=None, sname="", isin="KYG8879R1048:FUT",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == SWAP


def test_ftswap_isin_promoted_from_swap():
    got = classify_instrument(
        cat="FUT", sname="SX7E INDEX", isin="SX7E INDEX",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == FTSWAP


def test_ftswap_via_gscbihkt_isin():
    got = classify_instrument(
        cat="FUT", sname="GSCBIHKT CFD GOLDMAN", isin="GSCBIHKT",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == FTSWAP


def test_index_substring_promotes_to_ftswap_even_when_isin_unknown():
    """Safety net for an index CFD not yet added to ftswap_isins."""
    got = classify_instrument(
        cat="FUT", sname="SOME NEW INDEX CFD", isin="UNKNOWN-XYZ",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == FTSWAP


# ─── normalisation ──────────────────────────────────────────────────────────

def test_inputs_are_case_and_whitespace_insensitive():
    got = classify_instrument(
        cat="  fut  ", sname=" some cfd ", isin=" sx7e index ",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == FTSWAP


def test_nan_inputs_default_to_ord():
    got = classify_instrument(
        cat=float("nan"), sname=float("nan"), isin="ABC",
        ftswap_isins=FTSWAP_SET,
    )
    assert got == ORD


def test_ftswap_set_can_be_empty():
    # If config has no FTSWAP ISINs, all derivatives are SWAP unless
    # they match the INDEX/GSCBIHKT name fallback.
    got = classify_instrument(
        cat="FUT", sname="HAPAG CFD", isin="DE0001",
        ftswap_isins=frozenset(),
    )
    assert got == SWAP


def test_returned_value_is_in_known_domain():
    for cat in ("ORD", "FUT", "CFD", "SWAP", None, "BOND"):
        out = classify_instrument(
            cat=cat, sname="X", isin="Y", ftswap_isins=FTSWAP_SET,
        )
        assert out in INSTRUMENT_TYPES


# ─── BUG-9 structural regression ────────────────────────────────────────────

def test_signature_does_not_accept_legacy_ord_column():
    """BUG-9 fix. The legacy ORD column (always 'A' in HP_VAL) must not be
    a parameter — passing it must fail loudly, not be silently ignored
    and not be silently used to force ORD."""
    sig = inspect.signature(classify_instrument)
    forbidden = {"ord", "ORD", "ord_flag", "ord_column", "is_ord"}
    leaked = forbidden & set(sig.parameters)
    assert not leaked, f"forbidden params on classify_instrument: {leaked}"

    # Calling with the legacy kw must raise TypeError, not silently accept.
    with pytest.raises(TypeError):
        classify_instrument(  # type: ignore[call-arg]
            cat="FUT", sname="X", isin="Y",
            ftswap_isins=FTSWAP_SET,
            ord_flag="A",
        )


def test_ftswap_isins_is_keyword_only():
    """Prevent accidental positional drift if someone adds a 4th positional arg."""
    sig = inspect.signature(classify_instrument)
    p = sig.parameters["ftswap_isins"]
    assert p.kind is inspect.Parameter.KEYWORD_ONLY
