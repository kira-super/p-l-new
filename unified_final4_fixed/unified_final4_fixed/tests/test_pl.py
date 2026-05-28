"""Tests for compute.pl — the P&L engine and the 11 BUG-N regressions.

Each `test_bug_NN_*` test pins one structural fix from DESIGN.md §4. If a
future refactor re-introduces the bug, that test fails. The tests use
small hand-built fixtures rather than full HP_VAL files so the arithmetic
is auditable line-by-line.
"""

from __future__ import annotations

import inspect
import math

import pandas as pd
import pytest

from oefof_pl.compute import pl as plmod
from oefof_pl.compute.aggregate import AggLine
from oefof_pl.compute.pl import (
    PCT_FIELDS,
    SNAKE_TO_DISPLAY,
    PositionRow,
    _to_ratio,
    compute_positions,
)
from oefof_pl.exceptions import ComputeIntegrityError, SchemaError


END_TS = pd.Timestamp("2026-01-31")


# ─── fixtures ───────────────────────────────────────────────────────────────

def _trades(rows):
    cols = ["PCODE_ORIG", "ISIN", "SNAME", "CCY", "T", "CDATE", "UNITS",
            "GROSSPRICE_LOCAL", "BUY_NET_LOCAL", "SELL_NET_LOCAL", "INCOME_LOCAL"]
    if not rows:
        df = pd.DataFrame(columns=cols)
        df["TRADE_ID"] = pd.Series(dtype="int64")
        return df
    df = pd.DataFrame(rows, columns=cols)
    df["CDATE"] = pd.to_datetime(df["CDATE"])
    df["TRADE_ID"] = range(1, len(df) + 1)
    return df


def _agg(isin, *, units, ptvalue_local=None, ptvalue_eur=None,
         nucost_local, ccy="EUR",
         cat="ORD", sname="STOCK", lst_price_local=0.0, xrate=1.0,
         country="DE", ls="L"):
    # Tests historically passed ``ptvalue_local`` meaning "the EUR
    # equivalent of the snapshot's PTVALUE for this position" — they all
    # used ccy="EUR" with xrate=1, so the value happened to be the same
    # as the EUR amount. Post-rename, accept either kwarg.
    if ptvalue_eur is None and ptvalue_local is None:
        raise TypeError("_agg needs ptvalue_eur (or legacy ptvalue_local)")
    val = ptvalue_eur if ptvalue_eur is not None else ptvalue_local
    return AggLine(
        isin=isin, sname=sname, ccy=ccy, ls=ls, country=country, cat=cat,
        units=units, ptvalue_eur=val, ptcost_eur=0.0,
        nucost_local=nucost_local, lst_price_local=lst_price_local,
        xrate=xrate, n_source_rows=1,
    )


def _resolver(default="IS"):
    return lambda isin, sname: default


FTSWAP_SET = frozenset({"SX7E INDEX", "GSCBIHKT"})


def _run(start_agg=None, end_agg=None, trades=None, *,
         eur_usd_start=0.85, eur_usd_end=0.86,
         fx_start=None, fx_end=None,
         resolver=None, ftswap=FTSWAP_SET):
    return compute_positions(
        trades_df=trades if trades is not None else _trades([]),
        start_agg=start_agg or {},
        end_agg=end_agg or {},
        fx_start=fx_start or {},
        fx_end=fx_end or {},
        eur_usd_start=eur_usd_start, eur_usd_end=eur_usd_end,
        end_date_ts=END_TS,
        analyst_resolver=resolver or _resolver(),
        ftswap_isins=ftswap,
    )


# ─── _to_ratio helper ───────────────────────────────────────────────────────

def test_to_ratio_basic():
    assert _to_ratio(50, 100) == 0.5
    assert _to_ratio(0, 100) == 0
    assert _to_ratio(50, 0) == 0  # division-by-zero protection
    assert _to_ratio(float("nan"), 100) == 0


# ─── Schema and shape ──────────────────────────────────────────────────────

def test_empty_inputs_produce_empty_outputs_with_correct_columns():
    pl_df, inc_df = _run()
    from oefof_pl.data.normalise import INCOME_COLUMNS, PL_COLUMNS
    assert list(pl_df.columns) == list(PL_COLUMNS)
    assert list(inc_df.columns) == list(INCOME_COLUMNS)
    assert pl_df.empty
    assert inc_df.empty


def test_trades_schema_error():
    bad = pd.DataFrame({"ISIN": ["X"]})
    with pytest.raises(SchemaError):
        _run(trades=bad)


# ─── ORD path: simple buy-then-hold in EUR ──────────────────────────────────

def test_ord_simple_buy_and_hold_eur():
    """Buy 100 shares at 50 EUR; hold to end at 60 EUR.
    Total = 60*100 - 0 + 0 - 5000 + 0 = 1000 EUR (unrealised)."""
    end = {"X": _agg("X", units=100, ptvalue_local=6000, nucost_local=50, ccy="EUR")}
    trades = _trades([
        ("OEFOF", "X", "STOCK", "EUR", "P", "2026-01-10", 100, 50, 5000, 0, 0),
    ])
    pl_df, _ = _run(end_agg=end, trades=trades)
    row = pl_df.iloc[0]
    assert row["Market Value End (EUR)"] == pytest.approx(6000.0)
    assert row["Market Value Start (EUR)"] == pytest.approx(0.0)
    assert row["Total P&L (EUR)"] == pytest.approx(1000.0)
    assert row["Realised P&L (EUR)"] == pytest.approx(0.0)


def test_ord_full_round_trip_eur():
    """Buy 100@50, sell 100@70. realised=2000. total=2000."""
    trades = _trades([
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 100, 50, 5000, 0, 0),
        ("OEFOF", "X", "S", "EUR", "S", "2026-01-20", 100, 70, 0, 7000, 0),
    ])
    pl_df, _ = _run(trades=trades)
    row = pl_df.iloc[0]
    assert row["Realised P&L (EUR)"] == pytest.approx(2000.0)
    assert row["Total P&L (EUR)"] == pytest.approx(2000.0)
    assert row["Ending Units"] == 0


# ─── BUG-1: percentages stored as ratios ───────────────────────────────────

def test_bug_01_total_pct_is_ratio_not_percent():
    """A 44% return must be stored as 0.44, NEVER 44."""
    end = {"X": _agg("X", units=100, ptvalue_local=144, nucost_local=1.0, ccy="EUR")}
    trades = _trades([
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 100, 1.0, 100, 0, 0),
    ])
    pl_df, _ = _run(end_agg=end, trades=trades)
    pct = pl_df.iloc[0]["Total P&L (%)"]
    assert 0 < pct < 1.0  # ratio form
    assert pct == pytest.approx(0.44, abs=0.001)


def test_bug_01_pct_validation_rejects_non_finite_percent():
    """compute._assert_pct_fields_are_ratios catches NaN/Inf in pct cols
    (a divide-by-zero ``_to_ratio`` should have guarded). The data-quality
    >500% threshold lives in VAL-03, not compute, so legitimate large
    gains can flow through to validate and be reported per row."""
    df = pd.DataFrame({col: [0.0] for col in SNAKE_TO_DISPLAY.values()})
    df["Total P&L (%)"] = [float("inf")]
    with pytest.raises(ComputeIntegrityError, match="non-finite"):
        plmod._assert_pct_fields_are_ratios(df)


def test_bug_01_pct_validation_allows_large_but_finite_ratio():
    """A 16x gain (ratio=16) is uncommon but not a compute bug \u2014 it must
    pass compute and be flagged by VAL-03 instead."""
    df = pd.DataFrame({col: [0.0] for col in SNAKE_TO_DISPLAY.values()})
    df["Total P&L (%)"] = [16.0]
    plmod._assert_pct_fields_are_ratios(df)  # must not raise


# ─── BUG-2: Start Price column present ──────────────────────────────────────

def test_bug_02_start_price_column_in_schema():
    pl_df, _ = _run()
    assert "Start Price (Local)" in pl_df.columns


# ─── BUG-3: SWAP total = ΔMV + income (View B identity) ─────────────────────

def test_bug_03_swap_total_equals_delta_mv_plus_income():
    """SWAP/CFD: Total P&L = (mv_end − mv_start) + Income (EUR).

    Daily mark-to-market on a CFD is captured in PTVALUE_EUR by the
    custodian. The period swap P&L is therefore ΔPTVALUE_EUR plus any
    booked income/financing trades — matching the View B identity used by
    ORD positions. (The pre-fix policy hard-coded MTM to zero.)
    """
    end = {"X": _agg("X", units=100, ptvalue_local=999_999, nucost_local=10, ccy="EUR", cat="FUT", sname="X CFD")}
    start = {"X": _agg("X", units=100, ptvalue_local=100_000, nucost_local=10, ccy="EUR", cat="FUT", sname="X CFD")}
    trades = _trades([
        ("OEFOF", "X", "X CFD", "EUR", "I", "2026-01-15", 0, 0, 0, 0, 1234.5),
    ])
    pl_df, _ = _run(start_agg=start, end_agg=end, trades=trades)
    row = pl_df.iloc[0]
    assert row["Instrument"] == "SWAP"
    delta_mv = 999_999 - 100_000
    assert row["Unrealised P&L (EUR)"] == pytest.approx(delta_mv)
    assert row["Realised P&L (EUR)"] == 0.0
    assert row["Total P&L (EUR)"] == pytest.approx(delta_mv + 1234.5)


def test_bug_03_swap_function_signature_excludes_start_fx():
    """Structural: _compute_swap_position must not accept fx_start or eur_usd_start."""
    sig = inspect.signature(plmod._compute_swap_position)
    forbidden = {"fx_start", "eur_usd_start"}
    leaked = forbidden & set(sig.parameters)
    assert not leaked, f"_compute_swap_position must not accept {leaked}"


# ─── BUG-4: trade FX uses uniform end-period rate ───────────────────────────

def test_bug_04_corporate_action_pair_nets_to_zero_at_uniform_fx():
    """A 'sell A then buy B equivalent' on the same day at the same notional
    must net to ~0 EUR, not produce a 200k EUR phantom gain (legacy bug)."""
    # KRW position: sell 100 @ 70_000, buy 100 @ 70_000 same day.
    # All FX uniform so cash flows cancel exactly.
    end = {"X": _agg("X", units=100, ptvalue_local=7_000_000, nucost_local=70_000, ccy="KRW", xrate=1300)}
    trades = _trades([
        ("OEFOF", "X", "S", "KRW", "P", "2026-01-15", 100, 70_000, 7_000_000, 0, 0),
        ("OEFOF", "X", "S", "KRW", "S", "2026-01-15", 100, 70_000, 0, 7_000_000, 0),
    ])
    pl_df, _ = _run(
        start_agg={"X": _agg("X", units=100, ptvalue_local=7_000_000, nucost_local=70_000, ccy="KRW", xrate=1300)},
        end_agg=end, trades=trades,
        fx_start={"KRW": 1300}, fx_end={"KRW": 1300},
        eur_usd_start=0.86, eur_usd_end=0.86,
    )
    # Total should be ~0 (no MV change, no income, buy and sell cancel)
    assert abs(pl_df.iloc[0]["Total P&L (EUR)"]) < 1.0


def test_bug_04_tradegross_usd_not_in_schema():
    """The TRADEGROSS_USD column must not be a required input."""
    from oefof_pl.data.normalise import TRADES_REQUIRED
    assert "TRADEGROSS_USD" not in TRADES_REQUIRED


# ─── BUG-5: SWAP cost basis = |NUCOST × Units| ──────────────────────────────

def test_bug_05_swap_cost_basis_is_notional():
    """Cost basis for SWAP = |NUCOST × Units|, NOT |StartMV| + |BuyEUR|."""
    end = {"X": _agg("X", units=100, ptvalue_local=15_000, nucost_local=120, ccy="EUR", cat="FUT", sname="X CFD")}
    pl_df, _ = _run(end_agg=end, trades=_trades([]))
    row = pl_df.iloc[0]
    assert row["Instrument"] == "SWAP"
    # Notional = 100 * 120 = 12_000
    assert row["Cost Basis (EUR)"] == pytest.approx(12_000.0)


# ─── BUG-6: First Trade Date only for new positions ─────────────────────────

def test_bug_06_first_trade_date_only_for_new_positions():
    """Existing positions (in start_agg) get NaT for First Trade Date."""
    start = {"OLD": _agg("OLD", units=100, ptvalue_local=5000, nucost_local=50)}
    end = {
        "OLD": _agg("OLD", units=100, ptvalue_local=5000, nucost_local=50),
        "NEW": _agg("NEW", units=50, ptvalue_local=2500, nucost_local=50),
    }
    trades = _trades([
        ("OEFOF", "OLD", "X", "EUR", "P", "2026-01-15", 50, 50, 2500, 0, 0),
        ("OEFOF", "NEW", "Y", "EUR", "P", "2026-01-20", 50, 50, 2500, 0, 0),
    ])
    pl_df, _ = _run(start_agg=start, end_agg=end, trades=trades)
    by_isin = pl_df.set_index("ISIN")
    assert pd.isna(by_isin.loc["OLD", "First Trade Date"])
    assert by_isin.loc["NEW", "First Trade Date"] == pd.Timestamp("2026-01-20")


def test_bug_06_reentry_uses_last_entry_date_for_existing_position():
    start = {"X": _agg("X", units=100, ptvalue_local=5_000, nucost_local=50, ccy="EUR")}
    end = {"X": _agg("X", units=20, ptvalue_local=1_200, nucost_local=60, ccy="EUR")}
    trades = _trades([
        ("OEFOF", "X", "STOCK", "EUR", "S", "2026-01-05", 100, 55, 0, 5_500, 0),
        ("OEFOF", "X", "STOCK", "EUR", "P", "2026-01-20", 20, 60, 1_200, 0, 0),
    ])
    pl_df, _ = _run(start_agg=start, end_agg=end, trades=trades)
    row = pl_df.iloc[0]
    assert row["First Trade Date"] == pd.Timestamp("2026-01-20")


def test_round_trip_new_position_uses_gross_buy_as_period_denom():
    trades = _trades([
        ("OEFOF", "X", "STOCK", "EUR", "P", "2026-01-10", 100, 50, 5000, 0, 0),
        ("OEFOF", "X", "STOCK", "EUR", "S", "2026-01-20", 100, 100, 0, 10000, 0),
    ])
    pl_df, _ = _run(trades=trades)
    row = pl_df.iloc[0]
    assert row["Total P&L (EUR)"] == pytest.approx(5000.0)
    assert row["Total P&L (%)"] == pytest.approx(1.0)


def test_colon_suffixed_isin_is_excluded_from_output_positions():
    end = {
        "X:FUT": _agg("X:FUT", units=100, ptvalue_local=1_000, nucost_local=10,
                        ccy="EUR", cat="FUT", sname="X CFD"),
    }
    pl_df, _ = _run(end_agg=end, trades=_trades([]))
    assert pl_df.empty


# ─── BUG-7: Vietnam Dairy hybrid (income uses per-trade CCY) ────────────────

def test_bug_07_income_uses_per_trade_ccy():
    """Position is in VND. An income trade booked in USD must convert at
    USD/EUR, not at VND/EUR (the legacy code used position CCY)."""
    end = {"X": _agg("X", units=100, ptvalue_local=2_500_000, nucost_local=25_000,
                     ccy="VND", xrate=24_000)}
    # 100 USD income trade
    trades = _trades([
        ("OEFOF", "X", "S", "USD", "I", "2026-01-15", 0, 0, 0, 0, 100.0),
    ])
    pl_df, _ = _run(
        end_agg=end, trades=trades,
        fx_end={"VND": 24_000, "USD": 1.0}, eur_usd_end=0.86,
    )
    # 100 USD * 0.86 = 86 EUR. Wrong (legacy) would convert as VND: 100/24000*0.86 ≈ 0.0036 EUR.
    assert pl_df.iloc[0]["Income (EUR)"] == pytest.approx(86.0)


# ─── BUG-8: VDATE is the date authority — tested in test_loader.py ──────────

# (No test here; date authority lives in data.loader. See Phase 8.)


# ─── BUG-9: classifier doesn't use ORD column — tested in test_classify ─────

# (No test here; covered in tests/test_classify.py::test_signature_does_not_accept_legacy_ord_column)


# ─── BUG-10: SWAP realised/unrealised reflect ΔMV (View B) ──────────────────

def test_bug_10_swap_unrealised_equals_delta_ptvalue_when_open():
    """For a swap that's still open at period end, Unrealised = ΔPTVALUE_EUR
    (the daily MTM accumulated over the period), Realised = 0, and
    Total = Unrealised + Income."""
    start = {"X": _agg("X", units=100, ptvalue_local=10_000, nucost_local=120, ccy="EUR", cat="FUT", sname="X CFD")}
    end = {"X": _agg("X", units=100, ptvalue_local=15_000, nucost_local=120, ccy="EUR", cat="FUT", sname="X CFD")}
    trades = _trades([
        ("OEFOF", "X", "X CFD", "EUR", "I", "2026-01-15", 0, 0, 0, 0, 500),
    ])
    pl_df, _ = _run(start_agg=start, end_agg=end, trades=trades)
    row = pl_df.iloc[0]
    delta_mv = 15_000 - 10_000
    assert row["Realised P&L (EUR)"] == 0.0
    assert row["Unrealised P&L (EUR)"] == pytest.approx(delta_mv)
    assert row["Total P&L (EUR)"] == pytest.approx(delta_mv + 500)


def test_bug_10_swap_realised_equals_neg_start_mv_when_closed():
    """For a swap closed during the period (end_units = 0), the entire ΔMV
    is reclassified as Realised — the position dropped from mv_start to 0."""
    start = {"X": _agg("X", units=100, ptvalue_local=20_000, nucost_local=120, ccy="EUR", cat="FUT", sname="X CFD")}
    end = {}  # closed
    trades = _trades([])
    pl_df, _ = _run(start_agg=start, end_agg=end, trades=trades)
    row = pl_df.iloc[0]
    assert row["Ending Units"] == 0.0
    assert row["Realised P&L (EUR)"] == pytest.approx(-20_000.0)
    assert row["Unrealised P&L (EUR)"] == 0.0
    assert row["Total P&L (EUR)"] == pytest.approx(-20_000.0)


# ─── BUG-11: PTVALUE never used as EUR ──────────────────────────────────────

def test_bug_11_ord_ptvalue_used_directly_as_eur():
    """``PTVALUE_EUR`` is exactly that \u2014 a fund-base EUR value the snapshot\n    pre-computes for us. compute.pl must use it directly without\n    re-converting through xrate (that was the legacy double-conversion).\n    """
    end = {"X": _agg("X", units=100, ptvalue_eur=4_630.77, nucost_local=70_000,
                     ccy="KRW", xrate=1300, lst_price_local=70_000)}
    pl_df, _ = _run(
        end_agg=end, fx_end={"KRW": 1300}, eur_usd_end=0.86,
    )
    mv_end = pl_df.iloc[0]["Market Value End (EUR)"]
    # No FX double-conversion: snapshot value is taken as-is.
    assert mv_end == pytest.approx(4_630.77, rel=1e-9)


def test_bug_11_ptvalue_string_absent_from_compute_pl_source():
    """The bare column name 'PTVALUE' (without _LOCAL suffix) must not be
    used as a code identifier or string-literal in compute/pl.py — that
    would mean someone bypassed the rename."""
    import ast
    from pathlib import Path
    tree = ast.parse(Path(plmod.__file__).read_text(encoding="utf-8"))

    bad: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            if node.value == "PTVALUE":
                bad.append(f"string literal 'PTVALUE' at line {node.lineno}")
        elif isinstance(node, ast.Attribute) and node.attr == "PTVALUE":
            bad.append(f".PTVALUE attribute at line {node.lineno}")
        elif isinstance(node, ast.Name) and node.id == "PTVALUE":
            bad.append(f"name PTVALUE at line {node.lineno}")
    assert not bad, f"bare PTVALUE references in compute/pl.py: {bad}"


# ─── pl_df row-order and column-order determinism ───────────────────────────

def test_rows_sorted_by_isin():
    end = {
        "BBB": _agg("BBB", units=10, ptvalue_local=100, nucost_local=10),
        "AAA": _agg("AAA", units=10, ptvalue_local=100, nucost_local=10),
        "CCC": _agg("CCC", units=10, ptvalue_local=100, nucost_local=10),
    }
    pl_df, _ = _run(end_agg=end)
    assert list(pl_df["ISIN"]) == ["AAA", "BBB", "CCC"]


def test_compute_positions_keyword_only_args():
    """All optional arguments must be keyword-only — prevents arg-order bugs."""
    sig = inspect.signature(compute_positions)
    for name in ("fx_start", "fx_end", "eur_usd_start", "eur_usd_end",
                 "end_date_ts", "analyst_resolver", "ftswap_isins"):
        assert sig.parameters[name].kind is inspect.Parameter.KEYWORD_ONLY, (
            f"{name} must be keyword-only"
        )


# ─── income_df shape ────────────────────────────────────────────────────────

def test_income_df_one_row_per_income_trade():
    end = {"X": _agg("X", units=100, ptvalue_local=5000, nucost_local=50, ccy="EUR")}
    trades = _trades([
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 100, 50, 5000, 0, 0),
        ("OEFOF", "X", "S", "EUR", "I", "2026-01-15", 0, 0, 0, 0, 25.0),
        ("OEFOF", "X", "S", "EUR", "I", "2026-01-20", 0, 0, 0, 0, 30.0),
    ])
    _, inc_df = _run(end_agg=end, trades=trades)
    assert len(inc_df) == 2
    assert inc_df["Income (EUR)"].sum() == pytest.approx(55.0)


def test_income_df_preserves_per_trade_ccy():
    """BUG-7 (c) preservation in income_df."""
    end = {"X": _agg("X", units=100, ptvalue_local=5000, nucost_local=50, ccy="VND", xrate=24_000)}
    trades = _trades([
        ("OEFOF", "X", "S", "USD", "I", "2026-01-15", 0, 0, 0, 0, 100.0),
        ("OEFOF", "X", "S", "VND", "I", "2026-01-20", 0, 0, 0, 0, 1_000_000.0),
    ])
    _, inc_df = _run(end_agg=end, trades=trades,
                     fx_end={"VND": 24_000, "USD": 1.0}, eur_usd_end=0.86)
    ccys = set(inc_df["CCY"])
    assert ccys == {"USD", "VND"}


def test_position_closed_during_period_uses_start_fx_when_end_missing():
    """A JPY position held at start, fully sold by end \u2014 the JPY CCY is gone
    from the end snapshot's FX table. compute_positions must merge fx_start
    into fx_end so the closing trade and zero MV don't trip FXInvalidError.
    Real-world case observed against production HP_VAL on 2026-04-28.
    """
    start = {"X": _agg("X", units=100, ptvalue_local=15_000_000,
                       nucost_local=140_000, ccy="JPY", xrate=156.745)}
    # End: position absent. fx_end has no JPY entry.
    trades = _trades([
        ("OEFOF", "X", "TOYOTA", "JPY", "S", "2026-02-15",
         100, 160_000, 0, 16_000_000, 0),
    ])
    pl_df, _ = compute_positions(
        trades_df=trades,
        start_agg=start, end_agg={},
        fx_start={"JPY": 156.745, "USD": 1.0},
        fx_end={"USD": 1.0},                   # JPY missing on purpose
        eur_usd_start=0.86, eur_usd_end=0.86,
        end_date_ts=END_TS,
        analyst_resolver=_resolver(),
        ftswap_isins=FTSWAP_SET,
    )
    assert len(pl_df) == 1
    # Sanity: realised P&L is finite, not NaN \u2014 the fix worked.
    assert pd.notna(pl_df.iloc[0]["Realised P&L (EUR)"])


def test_zzzz_corporate_action_trades_excluded_from_wac():
    """HiPort marks SCODE migrations (bonus shares, ticker changes) with
    ``BCODE='ZZZZ'`` and price=0 P/S pairs. These are NOT economic
    transactions \u2014 they only book a unit-count change. Feeding them into
    WAC would dilute per-share cost to zero and then realise huge phantom
    losses on the matched sell. Real-world case: VN000000VCK5 (VPS
    Securities Bonus) on 2026-03-10/11.
    """
    cols = ["PCODE_ORIG", "ISIN", "SNAME", "CCY", "T", "CDATE", "UNITS",
            "GROSSPRICE_LOCAL", "BUY_NET_LOCAL", "SELL_NET_LOCAL",
            "INCOME_LOCAL", "BCODE"]
    rows = [
        # Real sell at non-zero price.
        ("OEFOF", "X", "X CO", "EUR", "S", "2026-01-15",
         10, 110.0, 0, 1100, 0, "MYUK"),
        # Corporate-action P/S pair at price=0 (would wreck WAC if not skipped).
        ("OEFOF", "X", "X CO", "EUR", "P", "2026-02-01",
         50, 0.0, 0, 0, 0, "ZZZZ"),
        ("OEFOF", "X", "X CO", "EUR", "S", "2026-02-01",
         50, 0.0, 0, 0, 0, "ZZZZ"),
    ]
    trades = pd.DataFrame(rows, columns=cols)
    trades["CDATE"] = pd.to_datetime(trades["CDATE"])
    trades["TRADE_ID"] = range(1, len(trades) + 1)
    start = {"X": _agg("X", units=100, ptvalue_local=10_000,
                       nucost_local=100, ccy="EUR")}
    end = {"X": _agg("X", units=90, ptvalue_local=9_900, nucost_local=100,
                     ccy="EUR")}
    pl_df, _ = compute_positions(
        trades_df=trades, start_agg=start, end_agg=end,
        fx_start={"EUR": 1.0, "USD": 1.0},
        fx_end={"EUR": 1.0, "USD": 1.0},
        eur_usd_start=0.86, eur_usd_end=0.86,
        end_date_ts=END_TS,
        analyst_resolver=_resolver(),
        ftswap_isins=FTSWAP_SET,
    )
    row = pl_df.iloc[0]
    # Realised = 10 \u00d7 (110 \u2212 100) = 100. The ZZZZ pair must contribute zero.
    assert row["Realised P&L (EUR)"] == pytest.approx(100.0)
    # Avg sell price excludes the ZZZZ row.
    assert row["Avg Sell Price (EUR)"] == pytest.approx(110.0)
    assert row["Units Sold"] == pytest.approx(10.0)
    assert row["Units Bought"] == pytest.approx(0.0)


def test_zzzz_income_dividend_preserved():
    """HiPort routes some real cash dividends through pseudo-broker ZZZZ
    with T='I', UNITS = current holding, NUPRICEG = DPS, INCOME_LOCAL =
    real cash. These MUST survive the ZZZZ filter \u2014 stripping them
    silently drops material dividend income (e.g. 1.5B VND for PNJ6).
    Only T='P'/'S' ZZZZ rows (SCODE-migration placeholders) get dropped.
    """
    cols = ["PCODE_ORIG", "ISIN", "SNAME", "CCY", "T", "CDATE", "UNITS",
            "GROSSPRICE_LOCAL", "BUY_NET_LOCAL", "SELL_NET_LOCAL",
            "INCOME_LOCAL", "BCODE"]
    rows = [
        # Real dividend booked through ZZZZ pseudo-broker (must survive).
        ("OEFOF", "Y", "Y CO", "EUR", "I", "2026-02-15",
         100, 2.5, 0, 0, 250.0, "ZZZZ"),
        # SCODE-migration placeholder (must be stripped).
        ("OEFOF", "Y", "Y CO", "EUR", "P", "2026-02-01",
         50, 0.0, 0, 0, 0, "ZZZZ"),
        ("OEFOF", "Y", "Y CO", "EUR", "S", "2026-02-01",
         50, 0.0, 0, 0, 0, "ZZZZ"),
    ]
    trades = pd.DataFrame(rows, columns=cols)
    trades["CDATE"] = pd.to_datetime(trades["CDATE"])
    trades["TRADE_ID"] = range(1, len(trades) + 1)
    start = {"Y": _agg("Y", units=100, ptvalue_local=10_000,
                       nucost_local=100, ccy="EUR")}
    end = {"Y": _agg("Y", units=100, ptvalue_local=11_000, nucost_local=100,
                     ccy="EUR")}
    pl_df, income_df = compute_positions(
        trades_df=trades, start_agg=start, end_agg=end,
        fx_start={"EUR": 1.0, "USD": 1.0},
        fx_end={"EUR": 1.0, "USD": 1.0},
        eur_usd_start=0.86, eur_usd_end=0.86,
        end_date_ts=END_TS,
        analyst_resolver=_resolver(),
        ftswap_isins=FTSWAP_SET,
    )
    row = pl_df.iloc[0]
    assert row["Income (Local)"] == pytest.approx(250.0)
    assert row["Income (EUR)"] == pytest.approx(250.0)


# ─── Dividends vs Swap Financing split ──────────────────────────────────────

def test_ord_income_goes_to_dividends_not_swap_financing():
    """ORD positions: income lands in Dividends (EUR); Swap Financing (EUR) = 0."""
    end = {"A": _agg("A", units=100, ptvalue_eur=11_000, nucost_local=100, ccy="EUR")}
    start = {"A": _agg("A", units=100, ptvalue_eur=10_000, nucost_local=100, ccy="EUR")}
    trades = _trades([
        ("OEFOF", "A", "STOCK A", "EUR", "I", "2026-01-15", 0, 0, 0, 0, 500.0),
    ])
    pl_df, _ = _run(start_agg=start, end_agg=end, trades=trades)
    row = pl_df.iloc[0]
    assert row["Instrument"] == "ORD"
    assert row["Income (EUR)"] == pytest.approx(500.0)
    assert row["Dividends (EUR)"] == pytest.approx(500.0)
    assert row["Swap Financing (EUR)"] == pytest.approx(0.0)


def test_swap_income_goes_to_swap_financing_not_dividends():
    """SWAP/FTSWAP positions: income lands in Swap Financing (EUR); Dividends (EUR) = 0."""
    end = {"X": _agg("X", units=100, ptvalue_eur=5_000, nucost_local=50, ccy="EUR",
                     cat="FUT", sname="X CFD")}
    start = {"X": _agg("X", units=100, ptvalue_eur=4_000, nucost_local=50, ccy="EUR",
                       cat="FUT", sname="X CFD")}
    trades = _trades([
        ("OEFOF", "X", "X CFD", "EUR", "I", "2026-01-15", 0, 0, 0, 0, -300.0),
    ])
    pl_df, _ = _run(start_agg=start, end_agg=end, trades=trades)
    row = pl_df.iloc[0]
    assert row["Instrument"] == "SWAP"
    assert row["Income (EUR)"] == pytest.approx(-300.0)
    assert row["Swap Financing (EUR)"] == pytest.approx(-300.0)
    assert row["Dividends (EUR)"] == pytest.approx(0.0)


def test_mixed_book_dividends_and_financing_are_independent():
    """ORD and SWAP in the same book: Dividends and Swap Financing don't bleed."""
    start = {
        "A": _agg("A", units=100, ptvalue_eur=10_000, nucost_local=100, ccy="EUR"),
        "X": _agg("X", units=50,  ptvalue_eur=2_000,  nucost_local=40,  ccy="EUR",
                  cat="FUT", sname="X CFD"),
    }
    end = {
        "A": _agg("A", units=100, ptvalue_eur=11_000, nucost_local=100, ccy="EUR"),
        "X": _agg("X", units=50,  ptvalue_eur=2_500,  nucost_local=40,  ccy="EUR",
                  cat="FUT", sname="X CFD"),
    }
    trades = _trades([
        ("OEFOF", "A", "STOCK A", "EUR", "I", "2026-01-10", 0, 0, 0, 0, 200.0),
        ("OEFOF", "X", "X CFD",   "EUR", "I", "2026-01-15", 0, 0, 0, 0, -80.0),
    ])
    pl_df, _ = _run(start_agg=start, end_agg=end, trades=trades)
    a = pl_df[pl_df["ISIN"] == "A"].iloc[0]
    x = pl_df[pl_df["ISIN"] == "X"].iloc[0]
    # ORD position A
    assert a["Dividends (EUR)"] == pytest.approx(200.0)
    assert a["Swap Financing (EUR)"] == pytest.approx(0.0)
    # SWAP position X
    assert x["Swap Financing (EUR)"] == pytest.approx(-80.0)
    assert x["Dividends (EUR)"] == pytest.approx(0.0)
    # Income (EUR) is the aggregate — unchanged by the split
    assert a["Income (EUR)"] == pytest.approx(200.0)
    assert x["Income (EUR)"] == pytest.approx(-80.0)
