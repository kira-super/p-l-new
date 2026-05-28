"""Tests for validate.checks — VAL-01..VAL-13."""

from __future__ import annotations

import pandas as pd
import pytest

from oefof_pl.compute.aggregate import AggLine
from oefof_pl.validate.checks import (
    CheckResult,
    ValidationResult,
)
from oefof_pl.validate import checks as _checks_mod


def run_all_checks(*, pl_df, trades_df, start_agg, end_agg, eur_usd_end=0.86):
    """Test wrapper supplying a default ``eur_usd_end`` for VAL-11."""
    return _checks_mod.run_all_checks(
        pl_df=pl_df, trades_df=trades_df,
        start_agg=start_agg, end_agg=end_agg,
        eur_usd_end=eur_usd_end,
    )


# ─── helpers ────────────────────────────────────────────────────────────────

def _agg(isin, *, units=0.0, ptvalue_eur=0.0, nucost_local=0.0,
         lst_price_local=0.0, ccy="EUR", cat="ORD", sname="STOCK",
         xrate=1.0, country="DE", ls="L"):
    return AggLine(
        isin=isin, sname=sname, ccy=ccy, ls=ls, country=country, cat=cat,
        units=units, ptvalue_eur=ptvalue_eur, ptcost_eur=0.0,
        nucost_local=nucost_local, lst_price_local=lst_price_local,
        xrate=xrate, n_source_rows=1,
    )


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


def _pl_row(**kw):
    """Build a P&L row dict with sensible defaults."""
    base = {
        "ISIN": "X", "Stock Name": "STOCK", "Country": "DE", "Analyst": "IS",
        "CCY": "EUR", "Instrument": "ORD", "L/S": "L",
        "Starting Units": 0.0, "Ending Units": 0.0,
        "Units Bought": 0.0, "Units Sold": 0.0,
        "Avg Buy Price (EUR)": 0.0, "Avg Sell Price (EUR)": 0.0,
        "Start Price (Local)": 0.0, "Last Price (EUR)": 0.0,
        "Market Value Start (EUR)": 0.0, "Market Value End (EUR)": 0.0,
        "Cost Basis (EUR)": 0.0,
        "Realised P&L (Local)": 0.0, "Realised P&L (EUR)": 0.0, "Realised P&L (%)": 0.0,
        "Unrealised P&L (Local)": 0.0, "Unrealised P&L (EUR)": 0.0, "Unrealised P&L (%)": 0.0,
        "Income (Local)": 0.0, "Income (EUR)": 0.0, "Income Yield (%)": 0.0,
        "Total P&L (EUR)": 0.0, "Total P&L (%)": 0.0,
        "Last Trade Date": pd.NaT, "First Trade Date": pd.NaT,
    }
    base.update(kw)
    return base


def _pl_df(rows):
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


# ─── Result wrappers ────────────────────────────────────────────────────────

def test_validation_result_has_critical_and_warnings():
    r = ValidationResult(checks=[
        CheckResult(id="A", name="a", status="PASS", detail=""),
        CheckResult(id="B", name="b", status="WARN", detail=""),
    ])
    assert r.has_critical is False
    assert r.has_warnings is True

    r2 = ValidationResult(checks=[
        CheckResult(id="A", name="a", status="FAIL", detail=""),
    ])
    assert r2.has_critical is True


def test_validation_result_by_id():
    r = ValidationResult(checks=[
        CheckResult(id="VAL-01", name="x", status="PASS", detail=""),
    ])
    assert r.by_id("VAL-01").status == "PASS"
    with pytest.raises(KeyError):
        r.by_id("VAL-99")


def test_run_all_checks_returns_expected_order():
    res = run_all_checks(
        pl_df=pd.DataFrame(),
        trades_df=_trades([]),
        start_agg={},
        end_agg={},
    )
    ids = [c.id for c in res.checks]
    # VAL-11 must run BEFORE VAL-03
    assert ids.index("VAL-11") < ids.index("VAL-03")
    expected = ["VAL-01", "VAL-02", "VAL-11", "VAL-03", "VAL-04",
                "VAL-05", "VAL-06", "VAL-07", "VAL-08", "VAL-09", "VAL-10", "VAL-13"]
    assert ids == expected


# ─── VAL-01 unit reconciliation ─────────────────────────────────────────────

def test_val_01_pass_balanced():
    start = {"X": _agg("X", units=100)}
    end = {"X": _agg("X", units=120)}
    trades = _trades([
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 50, 0, 0, 0, 0),
        ("OEFOF", "X", "S", "EUR", "S", "2026-01-15", 30, 0, 0, 0, 0),
    ])
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=trades,
                         start_agg=start, end_agg=end)
    assert res.by_id("VAL-01").status == "PASS"


def test_val_01_fail_short_units():
    start = {"X": _agg("X", units=100)}
    end = {"X": _agg("X", units=999)}  # impossible
    trades = _trades([])
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=trades,
                         start_agg=start, end_agg=end)
    c = res.by_id("VAL-01")
    assert c.status == "FAIL"
    assert len(c.failures) == 1
    assert c.failures[0]["ISIN"] == "X"
    assert res.has_critical is True


def test_val_01_ignores_colon_suffixed_synthetic_isin():
    start = {"X:FUT": _agg("X:FUT", units=100)}
    end = {"X:FUT": _agg("X:FUT", units=0)}
    trades = _trades([])
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=trades,
                         start_agg=start, end_agg=end)
    assert res.by_id("VAL-01").status == "PASS"


# ─── VAL-02 duplicate trades ────────────────────────────────────────────────

def test_val_02_pass_unique():
    trades = _trades([
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 50, 100, 5000, 0, 0),
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-11", 50, 100, 5000, 0, 0),
    ])
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=trades,
                         start_agg={}, end_agg={})
    assert res.by_id("VAL-02").status == "PASS"


def test_val_02_fail_duplicates_listed_completely():
    trades = _trades([
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 50, 100, 5000, 0, 0),
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 50, 100, 5000, 0, 0),
        ("OEFOF", "Y", "S", "EUR", "P", "2026-01-10", 25, 50, 1250, 0, 0),
        ("OEFOF", "Y", "S", "EUR", "P", "2026-01-10", 25, 50, 1250, 0, 0),
        ("OEFOF", "Y", "S", "EUR", "P", "2026-01-10", 25, 50, 1250, 0, 0),
    ])
    # Simulate a loader double-read: identical TRADE_IDs across rows.
    trades["TRADE_ID"] = [1, 1, 2, 2, 2]
    res = run_all_checks(
        pl_df=pd.DataFrame(),
        trades_df=trades,
        # Provide aggs so VAL-01 doesn't FAIL on these phantom trades.
        start_agg={"X": _agg("X", units=0), "Y": _agg("Y", units=0)},
        end_agg={"X": _agg("X", units=100), "Y": _agg("Y", units=75)},
    )
    c = res.by_id("VAL-02")
    assert c.status == "FAIL"
    # Both groups appear: 2 X rows + 3 Y rows = 5 failure entries (NOT truncated to 3)
    assert len(c.failures) == 5


# ─── VAL-11 PTVALUE EUR sanity (BUG-11 regression) ──────────────────────────
# Expected PTVALUE_EUR = (units * lst_price_local / xrate) * eur_usd_end
# For KRW: 100 * 70_000 / 1300 * 0.86 ≈ 4_630.77 EUR

def test_val_11_pass_when_ptvalue_eur_matches_units_price_fx():
    end = {"X": _agg("X", units=100, lst_price_local=70_000.0,
                     ptvalue_eur=4_630.77, ccy="KRW", xrate=1300.0)}
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=_trades([]),
                         start_agg={"X": _agg("X", units=100)},
                         end_agg=end)
    assert res.by_id("VAL-11").status == "PASS"


def test_val_11_fail_when_ptvalue_is_local_not_eur():
    """KRW position units=100, last_price=70_000. PTVALUE_EUR must be the
    EUR equivalent (~4630). If somebody fed in raw KRW (7M) as if it were
    EUR, VAL-11 must FAIL (BUG-11 reversed: was 'EUR-as-local', now 'local-as-EUR')."""
    end = {"X": _agg("X", units=100, lst_price_local=70_000.0,
                     ptvalue_eur=7_000_000.0,  # wrong: raw KRW posing as EUR
                     ccy="KRW", xrate=1300.0)}
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=_trades([]),
                         start_agg={"X": _agg("X", units=100)},
                         end_agg=end)
    c = res.by_id("VAL-11")
    assert c.status == "FAIL"
    assert c.failures[0]["ISIN"] == "X"
    assert res.has_critical is True


def test_val_11_skips_swap_positions():
    end = {"S": _agg("S", units=100, lst_price_local=70.0,
                     ptvalue_eur=999_999.0, ccy="USD", cat="SWAP")}
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=_trades([]),
                         start_agg={}, end_agg=end)
    assert res.by_id("VAL-11").status == "PASS"


# ─── VAL-03 row integrity ───────────────────────────────────────────────────

def test_val_03_pass_finite_row():
    df = _pl_df([_pl_row(ISIN="X", **{"Total P&L (EUR)": 100.0, "Total P&L (%)": 0.05})])
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    assert res.by_id("VAL-03").status == "PASS"


def test_val_03_fails_on_nonfinite_value():
    df = _pl_df([_pl_row(ISIN="X", **{"Total P&L (EUR)": float("inf")})])
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    assert res.by_id("VAL-03").status == "FAIL"


def test_val_03_fails_on_percentage_stored_as_percent_bug_1():
    """If somebody stores 44 instead of 0.44 (BUG-1), VAL-03 catches it."""
    df = _pl_df([_pl_row(ISIN="X", **{"Total P&L (%)": 44.0})])
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    c = res.by_id("VAL-03")
    assert c.status == "FAIL"
    assert "BUG-1" in c.failures[0]["Issues"]


# ─── VAL-04 SWAP identity ───────────────────────────────────────────────────

def test_val_04_pass_when_total_equals_components():
    df = _pl_df([_pl_row(ISIN="S", Instrument="SWAP",
                         **{"Total P&L (EUR)": 150.0,
                            "Realised P&L (EUR)": 60.0,
                            "Unrealised P&L (EUR)": 40.0,
                            "Income (EUR)": 50.0})])
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    assert res.by_id("VAL-04").status == "PASS"


def test_val_04_fail_when_swap_total_differs_from_components():
    df = _pl_df([_pl_row(ISIN="S", Instrument="SWAP",
                         **{"Total P&L (EUR)": 1000.0,
                            "Realised P&L (EUR)": 0.0,
                            "Unrealised P&L (EUR)": 0.0,
                            "Income (EUR)": 50.0})])
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    c = res.by_id("VAL-04")
    assert c.status == "FAIL"
    assert c.failures[0]["ISIN"] == "S"


# ─── VAL-05..VAL-10 (WARN) ──────────────────────────────────────────────────

def test_val_05_warns_when_income_trade_has_units():
    trades = _trades([
        ("OEFOF", "X", "S", "EUR", "I", "2026-01-10", 1_000_000, 0, 0, 0, 500),
    ])
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=trades,
                         start_agg={"X": _agg("X")}, end_agg={"X": _agg("X")})
    assert res.by_id("VAL-05").status == "WARN"


def test_val_06_warns_on_negative_cost_basis():
    df = _pl_df([_pl_row(ISIN="X", **{"Cost Basis (EUR)": -100.0})])
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    assert res.by_id("VAL-06").status == "WARN"


def test_val_07_warns_on_unassigned_analyst():
    df = _pl_df([_pl_row(ISIN="X", Analyst="UNASSIGNED")])
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    assert res.by_id("VAL-07").status == "WARN"


def test_val_08_warns_on_trade_ccy_mismatch():
    end = {"X": _agg("X", units=100, ccy="EUR")}
    trades = _trades([
        ("OEFOF", "X", "S", "USD", "P", "2026-01-10", 100, 1, 100, 0, 0),
    ])
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=trades,
                         start_agg={}, end_agg=end)
    # Don't assert VAL-01 status; we only care VAL-08 fired
    assert res.by_id("VAL-08").status == "WARN"


def test_val_09_warns_on_trades_only_isin():
    trades = _trades([
        ("OEFOF", "Z", "S", "EUR", "P", "2026-01-10", 1, 1, 1, 0, 0),
        ("OEFOF", "Z", "S", "EUR", "S", "2026-01-15", 1, 1, 0, 1, 0),
    ])
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=trades,
                         start_agg={}, end_agg={})
    assert res.by_id("VAL-09").status == "WARN"


def test_val_10_warns_on_phantom_zero_position():
    df = _pl_df([_pl_row(ISIN="X")])  # all zeros
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    assert res.by_id("VAL-10").status == "WARN"


def test_val_13_warns_when_open_position_has_zero_avg_buy():
    df = _pl_df([
        _pl_row(
            ISIN="X",
            **{
                "Ending Units": 100.0,
                "Units Bought": 100.0,
                "Avg Buy Price (EUR)": 0.0,
            },
        )
    ])
    res = run_all_checks(pl_df=df, trades_df=_trades([]),
                         start_agg={}, end_agg={})
    c = res.by_id("VAL-13")
    assert c.status == "WARN"
    assert c.failures[0]["ISIN"] == "X"


# ─── End-to-end: critical fail short-circuits Excel write decision ──────────

def test_critical_failure_signals_no_excel_should_be_written():
    """The whole point of VAL-* is that has_critical is the gate Pipeline.run
    uses to decide whether to write the workbook."""
    start = {"X": _agg("X", units=100)}
    end = {"X": _agg("X", units=999)}  # unit recon FAIL
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=_trades([]),
                         start_agg=start, end_agg=end)
    assert res.has_critical is True
    # The downstream contract:
    if not res.has_critical:
        pytest.fail("Pipeline would have written Excel — but VAL-01 failed.")


# ─── Failures are NEVER truncated (legacy first-3 bug) ──────────────────────

def test_failures_are_never_truncated_to_three():
    """Legacy code printed only first 3 failures. We list all of them."""
    end = {f"X{i}": _agg(f"X{i}", units=100, lst_price_local=70_000.0,
                         ptvalue_eur=7_000_000.0, ccy="KRW", xrate=1300.0) for i in range(10)}
    res = run_all_checks(pl_df=pd.DataFrame(), trades_df=_trades([]),
                         start_agg={f"X{i}": _agg(f"X{i}", units=100) for i in range(10)},
                         end_agg=end)
    c = res.by_id("VAL-11")
    assert c.status == "FAIL"
    assert len(c.failures) == 10  # NOT 3
