"""Validation gates VAL-01..VAL-13.

Run after ``compute.pl.compute_positions`` and **before** the Excel writer.
A FAIL result blocks Excel writing entirely — the user gets a single-sheet
failure report instead of a wrong P&L workbook. WARN results are recorded
but do not block.

Run order is significant. VAL-11 (PTVALUE local sanity) runs **before**
VAL-03 (per-row integrity) because if PTVALUE_EUR is wrong, every
downstream P&L number is wrong and VAL-03 would fail uselessly.

Failure lists are **complete** — never truncated. The legacy code's
"first 3 only" truncation is the reason BUG-7 went undetected for months.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Iterable, Literal

import pandas as pd

from ..compute.aggregate import AggLine
from ..compute.classify import FTSWAP, ORD, SWAP


Status = Literal["PASS", "WARN", "FAIL"]


@dataclass(frozen=True)
class CheckResult:
    """Outcome of a single validation check.

    ``failures`` is the *complete* list of offending rows — each is a dict
    keyed for inclusion in the failure-report sheet. Never truncated.
    """
    id: str
    name: str
    status: Status
    detail: str
    failures: list[dict] = field(default_factory=list)


@dataclass(frozen=True)
class ValidationResult:
    """The result of running all checks. Order preserved per ``run_order``."""
    checks: list[CheckResult]

    @property
    def has_critical(self) -> bool:
        """True if any check FAILed. Pipeline must NOT write the P&L workbook."""
        return any(c.status == "FAIL" for c in self.checks)

    @property
    def has_warnings(self) -> bool:
        return any(c.status == "WARN" for c in self.checks)

    def by_id(self, check_id: str) -> CheckResult:
        for c in self.checks:
            if c.id == check_id:
                return c
        raise KeyError(check_id)


# ─── Public entry point ─────────────────────────────────────────────────────

def run_all_checks(
    *,
    pl_df: pd.DataFrame,
    trades_df: pd.DataFrame,
    start_agg: dict[str, AggLine],
    end_agg: dict[str, AggLine],
    eur_usd_end: float,
    th_val_df: pd.DataFrame | None = None,
) -> ValidationResult:
    """Run VAL-01..VAL-13 in the documented order. Returns ``ValidationResult``."""
    checks: list[CheckResult] = []

    checks.append(_val_01_unit_reconciliation(start_agg, end_agg, trades_df))
    checks.append(_val_02_duplicate_trades(trades_df))
    checks.append(_val_11_ptvalue_eur_sanity(end_agg, eur_usd_end))
    checks.append(_val_03_row_integrity(pl_df, trades_df))
    checks.append(_val_04_swap_total_equals_income(pl_df))

    checks.append(_val_05_income_with_units(trades_df))
    checks.append(_val_06_negative_cost_basis(pl_df))
    checks.append(_val_07_unassigned_analyst(pl_df))
    checks.append(_val_08_trades_ccy_consistency(trades_df, end_agg))
    checks.append(_val_09_trades_only_isins(trades_df, start_agg, end_agg))
    checks.append(_val_10_zero_units_with_no_trades(pl_df, trades_df))

    # VAL-12 requires tH_VAL data; skip cleanly if unavailable.
    if th_val_df is not None and not th_val_df.empty:
        checks.append(_val_12_isin_valid(th_val_df))

    checks.append(_val_13_zero_avg_buy_with_open_units(pl_df))

    return ValidationResult(checks=checks)


# ─── VAL-01 unit reconciliation (FAIL) ──────────────────────────────────────

_UNIT_TOL = 1e-3


def _val_01_unit_reconciliation(
    start_agg: dict[str, AggLine],
    end_agg: dict[str, AggLine],
    trades_df: pd.DataFrame,
) -> CheckResult:
    """end_units ≈ start_units + Σbuys − Σsells per ISIN, tol 0.001."""
    failures: list[dict] = []

    all_isins = set(start_agg) | set(end_agg)
    if not trades_df.empty:
        from ..data.normalise import normalise_isin
        norm_isins = trades_df["ISIN"].map(normalise_isin)
        all_isins |= set(norm_isins.unique()) - {""}

    bought = _sum_by_isin(trades_df, t_code="P")
    sold = _sum_by_isin(trades_df, t_code="S")

    for isin in sorted(all_isins):
        if ":" in isin:
            continue
        s_units = float(start_agg[isin].units) if isin in start_agg else 0.0
        e_units = float(end_agg[isin].units) if isin in end_agg else 0.0
        b = bought.get(isin, 0.0)
        sl = sold.get(isin, 0.0)
        expected = s_units + b - sl
        diff = e_units - expected
        if abs(diff) > _UNIT_TOL:
            sname = (
                end_agg[isin].sname if isin in end_agg
                else (start_agg[isin].sname if isin in start_agg else "")
            )
            failures.append({
                "ISIN": isin, "Stock Name": sname,
                "Start Units": s_units, "Bought": b, "Sold": sl,
                "Expected End Units": expected, "Actual End Units": e_units,
                "Diff": diff,
            })

    if failures:
        return CheckResult(
            id="VAL-01", name="Unit reconciliation", status="FAIL",
            detail=f"{len(failures)} ISIN(s) failed unit reconciliation.",
            failures=failures,
        )
    return CheckResult(
        id="VAL-01", name="Unit reconciliation", status="PASS",
        detail=f"All {len(all_isins)} ISIN(s) reconcile within {_UNIT_TOL}.",
    )


# ─── VAL-02 duplicate trades (FAIL) ─────────────────────────────────────────

def _val_02_duplicate_trades(trades_df: pd.DataFrame) -> CheckResult:
    """Two trades sharing the same bottler ``TRADE_ID`` mean the loader
    double-read a row (or the bottler export double-printed it). Coarser
    keys would flag legitimate parallel-lot swap accruals (two open lots
    with identical PCODE/ISIN/CDATE/T/UNITS/GROSSPRICE on the same day)."""
    if trades_df.empty:
        return CheckResult(
            id="VAL-02", name="Duplicate trades", status="PASS",
            detail="No trades to check.",
        )

    if "TRADE_ID" not in trades_df.columns:
        return CheckResult(
            id="VAL-02", name="Duplicate trades", status="WARN",
            detail="Cannot check — TRADE_ID column missing.",
        )

    dup_mask = trades_df.duplicated(subset=["TRADE_ID"], keep=False)
    if not dup_mask.any():
        return CheckResult(
            id="VAL-02", name="Duplicate trades", status="PASS",
            detail=f"No duplicate TRADE_IDs among {len(trades_df)} trades.",
        )

    detail_cols = [c for c in (
        "TRADE_ID", "PCODE_ORIG", "ISIN", "CDATE", "T", "UNITS", "GROSSPRICE_LOCAL",
    ) if c in trades_df.columns]
    dups = trades_df.loc[dup_mask, detail_cols].sort_values(detail_cols)
    failures = dups.to_dict(orient="records")
    return CheckResult(
        id="VAL-02", name="Duplicate trades", status="FAIL",
        detail=f"{len(failures)} duplicate trade row(s) detected.",
        failures=failures,
    )


# ─── VAL-11 PTVALUE_EUR sanity (FAIL) — runs before VAL-03 ────────────────

_PTVALUE_REL_TOL = 0.01    # 1%
_PTVALUE_ABS_TOL = 1.0     # absolute floor in EUR


def _val_11_ptvalue_eur_sanity(
    end_agg: dict[str, AggLine], eur_usd_end: float,
) -> CheckResult:
    """For ORD positions: PTVALUE_EUR ≈ (UNITS × LST_PRICE_LOCAL / XRATE) × EUR_USD.

    HiPort writes PTVALUE in the fund base currency (EUR). If this check
    fails the snapshot is in a different unit (legacy data, or a column-
    rename slip in :mod:`oefof_pl.data.normalise`) and every Total P&L is
    wrong (BUG-11). Runs before VAL-03 to avoid compounding noise.

    The EUR-leg of the conversion uses the position's own snapshot XRATE;
    EUR/USD is the period-end rate.
    """
    import math as _math

    failures: list[dict] = []
    if not _math.isfinite(eur_usd_end) or eur_usd_end <= 0:
        return CheckResult(
            id="VAL-11", name="PTVALUE EUR sanity", status="WARN",
            detail=f"eur_usd_end is invalid ({eur_usd_end!r}); cannot run check.",
        )

    for isin, agg in end_agg.items():
        if agg.cat in ("FUT", "CFD", "SWAP"):
            continue
        if agg.units == 0 or agg.lst_price_local == 0:
            continue

        ccy_u = (agg.ccy or "").strip().upper()
        units_x_price_local = agg.units * agg.lst_price_local

        if ccy_u == "EUR":
            expected_eur = units_x_price_local
        elif ccy_u == "USD":
            expected_eur = units_x_price_local * eur_usd_end
        else:
            if not _math.isfinite(agg.xrate) or agg.xrate <= 0:
                continue  # FX missing — caught elsewhere
            expected_eur = (units_x_price_local / agg.xrate) * eur_usd_end

        diff = agg.ptvalue_eur - expected_eur
        denom = max(abs(expected_eur), _PTVALUE_ABS_TOL)
        if abs(diff) / denom > _PTVALUE_REL_TOL and abs(diff) > _PTVALUE_ABS_TOL:
            failures.append({
                "ISIN": isin, "Stock Name": agg.sname, "CCY": agg.ccy,
                "Units": agg.units, "Last Price (Local)": agg.lst_price_local,
                "XRATE (local/USD)": agg.xrate,
                "Expected PTVALUE_EUR": expected_eur,
                "Actual PTVALUE_EUR": agg.ptvalue_eur,
                "Diff (EUR)": diff,
            })

    if failures:
        return CheckResult(
            id="VAL-11", name="PTVALUE EUR sanity", status="FAIL",
            detail=(
                f"{len(failures)} ORD position(s) have PTVALUE_EUR that "
                "does not match (Units × Last Price) converted to EUR within 1%. "
                "Likely cause: PTVALUE is in a different currency than expected (BUG-11)."
            ),
            failures=failures,
        )
    return CheckResult(
        id="VAL-11", name="PTVALUE EUR sanity", status="PASS",
        detail="All ORD PTVALUE_EUR values reconcile to (Units × Last Price) in EUR.",
    )


# ─── VAL-03 row-level integrity (FAIL) ──────────────────────────────────────

def _val_03_row_integrity(
    pl_df: pd.DataFrame, trades_df: pd.DataFrame,
) -> CheckResult:
    """Every monetary cell must be finite. Every percentage must be a ratio
    (|x| ≤ 5). The compute layer also enforces these but we re-check at the
    output boundary so a future refactor can't slip through.

    The ±500% guard exists to catch BUG-1 (percentage stored as percent
    instead of ratio). For ISINs that received a CA_ADJ synthetic trade
    (bonus / split / consolidation) the % can legitimately exceed 500%
    because cost basis stays at the cash-paid level while market value
    reflects the diluted share count. In that case we still record the row
    but downgrade the whole check to WARN so the workbook can build.
    Non-finite values and non-CA rows still FAIL.
    """
    if pl_df.empty:
        return CheckResult(
            id="VAL-03", name="Row integrity", status="PASS",
            detail="No rows to check.",
        )

    monetary_cols = [
        "Market Value Start (EUR)", "Market Value End (EUR)",
        "Cost Basis (EUR)", "Realised P&L (EUR)", "Unrealised P&L (EUR)",
        "Income (EUR)", "Dividends (EUR)", "Swap Financing (EUR)",
        "Total P&L (EUR)",
    ]
    pct_cols = [
        "Realised P&L (%)", "Unrealised P&L (%)",
        "Income Yield (%)", "Total P&L (%)",
    ]

    ca_isins: set[str] = set()
    if not trades_df.empty and "BCODE" in trades_df.columns and "ISIN" in trades_df.columns:
        ca_mask = trades_df["BCODE"].astype(str).str.upper() == "CA_ADJ"
        ca_isins = set(trades_df.loc[ca_mask, "ISIN"].dropna().astype(str))

    failures: list[dict] = []
    has_critical = False

    for _, row in pl_df.iterrows():
        isin = str(row.get("ISIN", ""))
        is_ca = isin in ca_isins
        cost_basis = float(row.get("Cost Basis (EUR)") or 0)
        is_zero_cost = abs(cost_basis) < 1e-9
        bad_fields: list[str] = []
        row_critical = False
        for c in monetary_cols:
            if c not in pl_df.columns:
                continue
            v = row[c]
            if v is None or (isinstance(v, float) and not math.isfinite(v)):
                bad_fields.append(f"{c}={v!r}")
                row_critical = True
        for c in pct_cols:
            if c not in pl_df.columns:
                continue
            v = row[c]
            if v is None or (isinstance(v, float) and not math.isfinite(v)):
                if not is_zero_cost:  # NaN pct is expected when cost basis = 0
                    bad_fields.append(f"{c}={v!r}")
                    row_critical = True
            elif isinstance(v, (int, float)) and abs(v) > 5.0:
                if is_ca:
                    bad_fields.append(
                        f"{c}={v} (>500%, allowed: ISIN has CA_ADJ trade — "
                        "bonus/split dilutes WAC vs MV)"
                    )
                else:
                    bad_fields.append(
                        f"{c}={v} (>500%, percentage stored as percent? — BUG-1)"
                    )
                    row_critical = True
        if bad_fields:
            failures.append({
                "ISIN": isin,
                "Stock Name": row.get("Stock Name", ""),
                "Issues": "; ".join(bad_fields),
            })
            if row_critical:
                has_critical = True

    if not failures:
        return CheckResult(
            id="VAL-03", name="Row-level integrity", status="PASS",
            detail=f"All {len(pl_df)} row(s) finite and percentages in ratio form.",
        )
    if has_critical:
        return CheckResult(
            id="VAL-03", name="Row-level integrity", status="FAIL",
            detail=f"{len(failures)} P&L row(s) have non-finite or out-of-range values.",
            failures=failures,
        )
    return CheckResult(
        id="VAL-03", name="Row-level integrity", status="WARN",
        detail=(
            f"{len(failures)} P&L row(s) have |%| > 500% — all are ISINs with "
            "CA_ADJ trades (bonus / split / consolidation), so the large % is "
            "a legitimate consequence of WAC dilution, not BUG-1. Review the "
            "row(s) but the workbook is safe to publish."
        ),
        failures=failures,
    )


# ─── VAL-04 SWAP identity Total == Realised + Unrealised + Income ──────────

_SWAP_TOL = 0.01  # EUR


def _val_04_swap_total_equals_income(pl_df: pd.DataFrame) -> CheckResult:
    """For SWAP/FTSWAP rows: Total P&L (EUR) must equal
    Realised + Unrealised + Income (EUR), matching the View B identity used
    on the ORD path. The MTM legs are sourced from \u0394PTVALUE_EUR; income
    is sourced from booked income trades."""
    if pl_df.empty or "Instrument" not in pl_df.columns:
        return CheckResult(
            id="VAL-04", name="SWAP identity", status="PASS",
            detail="No SWAP rows to check.",
        )

    swap_rows = pl_df[pl_df["Instrument"].isin([SWAP, FTSWAP])]
    if swap_rows.empty:
        return CheckResult(
            id="VAL-04", name="SWAP identity", status="PASS",
            detail="No SWAP / FTSWAP positions.",
        )

    failures: list[dict] = []
    for _, row in swap_rows.iterrows():
        total = float(row["Total P&L (EUR)"])
        realised = float(row["Realised P&L (EUR)"])
        unrealised = float(row["Unrealised P&L (EUR)"])
        income = float(row["Income (EUR)"])
        expected = realised + unrealised + income
        diff = total - expected
        if abs(diff) > _SWAP_TOL:
            failures.append({
                "ISIN": row["ISIN"], "Stock Name": row["Stock Name"],
                "Instrument": row["Instrument"],
                "Total P&L (EUR)": total,
                "Realised (EUR)": realised,
                "Unrealised (EUR)": unrealised,
                "Income (EUR)": income,
                "Expected (R+U+I)": expected,
                "Diff (EUR)": diff,
            })

    if failures:
        return CheckResult(
            id="VAL-04", name="SWAP identity", status="FAIL",
            detail=(
                f"{len(failures)} SWAP/FTSWAP row(s) violate "
                "Total = Realised + Unrealised + Income."
            ),
            failures=failures,
        )
    return CheckResult(
        id="VAL-04", name="SWAP identity", status="PASS",
        detail=(
            f"All {len(swap_rows)} SWAP/FTSWAP row(s) satisfy "
            "Total == Realised + Unrealised + Income."
        ),
    )


# ─── VAL-05..VAL-10 (WARN) ──────────────────────────────────────────────────

def _val_05_income_with_units(trades_df: pd.DataFrame) -> CheckResult:
    """Income trades (T='I') with non-zero UNITS are suspicious.

    The legacy 'income contains 1m share dividend' bug had I-rows with
    huge UNITS. With BUG-5 fixed in WAC they no longer affect cost basis,
    but they should still be flagged for the analyst to review.
    """
    if trades_df.empty or "T" not in trades_df.columns:
        return CheckResult(id="VAL-05", name="Income trades have units",
                           status="PASS", detail="N/A.")
    inc = trades_df[trades_df["T"] == "I"]
    bad = inc[pd.to_numeric(inc["UNITS"], errors="coerce").fillna(0).abs() > 0]
    if bad.empty:
        return CheckResult(id="VAL-05", name="Income trades have units",
                           status="PASS", detail="No income rows with UNITS.")
    return CheckResult(
        id="VAL-05", name="Income trades have units", status="WARN",
        detail=f"{len(bad)} income trade(s) carry non-zero UNITS.",
        failures=bad[["ISIN", "CDATE", "UNITS", "INCOME_LOCAL"]].to_dict(orient="records"),
    )


def _val_06_negative_cost_basis(pl_df: pd.DataFrame) -> CheckResult:
    if pl_df.empty or "Cost Basis (EUR)" not in pl_df.columns:
        return CheckResult(id="VAL-06", name="Negative cost basis",
                           status="PASS", detail="N/A.")
    bad = pl_df[pl_df["Cost Basis (EUR)"] < 0]
    if bad.empty:
        return CheckResult(id="VAL-06", name="Negative cost basis",
                           status="PASS", detail="No negative cost bases.")
    return CheckResult(
        id="VAL-06", name="Negative cost basis", status="WARN",
        detail=f"{len(bad)} position(s) have negative cost basis.",
        failures=bad[["ISIN", "Stock Name", "Cost Basis (EUR)"]].to_dict(orient="records"),
    )


def _val_07_unassigned_analyst(pl_df: pd.DataFrame) -> CheckResult:
    if pl_df.empty or "Analyst" not in pl_df.columns:
        return CheckResult(id="VAL-07", name="Unassigned analysts",
                           status="PASS", detail="N/A.")
    bad = pl_df[pl_df["Analyst"].astype("string").str.upper() == "UNASSIGNED"]
    if bad.empty:
        return CheckResult(id="VAL-07", name="Unassigned analysts",
                           status="PASS", detail="All positions have an analyst.")
    return CheckResult(
        id="VAL-07", name="Unassigned analysts", status="WARN",
        detail=(
            f"{len(bad)} position(s) have UNASSIGNED analyst — "
            "fill in isin_analyst_map.csv."
        ),
        failures=bad[["ISIN", "Stock Name"]].to_dict(orient="records"),
    )


def _val_08_trades_ccy_consistency(
    trades_df: pd.DataFrame,
    end_agg: dict[str, AggLine],
) -> CheckResult:
    """Trade CCY differs from end-snapshot CCY for the same ISIN. Legitimate
    for hybrid positions (Vietnam Dairy USD income on a VND ORD); flag for
    review only."""
    if trades_df.empty or not end_agg:
        return CheckResult(id="VAL-08", name="Trade CCY mismatch",
                           status="PASS", detail="N/A.")

    from ..data.normalise import normalise_isin
    bad: list[dict] = []
    for isin, ccy in zip(trades_df["ISIN"].map(normalise_isin), trades_df["CCY"]):
        if not isin or isin not in end_agg:
            continue
        agg_ccy = end_agg[isin].ccy
        trade_ccy = (str(ccy) if ccy is not None else "").strip().upper()
        if trade_ccy and agg_ccy and trade_ccy != agg_ccy:
            bad.append({"ISIN": isin, "Position CCY": agg_ccy, "Trade CCY": trade_ccy})

    if not bad:
        return CheckResult(id="VAL-08", name="Trade CCY mismatch",
                           status="PASS", detail="All trade CCYs match position CCY.")
    # Dedupe
    seen = set()
    deduped = []
    for r in bad:
        k = (r["ISIN"], r["Trade CCY"])
        if k not in seen:
            seen.add(k)
            deduped.append(r)
    return CheckResult(
        id="VAL-08", name="Trade CCY mismatch", status="WARN",
        detail=f"{len(deduped)} ISIN/CCY pair(s) where trade CCY ≠ position CCY.",
        failures=deduped,
    )


def _val_09_trades_only_isins(
    trades_df: pd.DataFrame,
    start_agg: dict[str, AggLine],
    end_agg: dict[str, AggLine],
) -> CheckResult:
    """ISINs that have trades but appear in neither snapshot — usually a
    full round-trip that closed before the end snapshot (legitimate), but
    worth flagging."""
    if trades_df.empty:
        return CheckResult(id="VAL-09", name="Trades-only ISINs",
                           status="PASS", detail="No trades.")
    from ..data.normalise import normalise_isin
    trade_isins = set(trades_df["ISIN"].map(normalise_isin)) - {""}
    snapshot_isins = set(start_agg) | set(end_agg)
    orphans = sorted(trade_isins - snapshot_isins)
    if not orphans:
        return CheckResult(id="VAL-09", name="Trades-only ISINs",
                           status="PASS", detail="All traded ISINs appear in a snapshot.")
    return CheckResult(
        id="VAL-09", name="Trades-only ISINs", status="WARN",
        detail=f"{len(orphans)} ISIN(s) have trades but no snapshot row.",
        failures=[{"ISIN": x} for x in orphans],
    )


def _val_10_zero_units_with_no_trades(
    pl_df: pd.DataFrame,
    trades_df: pd.DataFrame,
) -> CheckResult:
    """Position with 0 ending units AND 0 starting units AND no trades —
    this is dead data that shouldn't show up at all."""
    if pl_df.empty:
        return CheckResult(id="VAL-10", name="Phantom zero positions",
                           status="PASS", detail="N/A.")

    zero = pl_df[
        (pl_df["Ending Units"] == 0)
        & (pl_df["Starting Units"] == 0)
        & (pl_df["Units Bought"] == 0)
        & (pl_df["Units Sold"] == 0)
    ]
    if zero.empty:
        return CheckResult(id="VAL-10", name="Phantom zero positions",
                           status="PASS", detail="No phantom rows.")
    return CheckResult(
        id="VAL-10", name="Phantom zero positions", status="WARN",
        detail=f"{len(zero)} position(s) have zero units and no trades.",
        failures=zero[["ISIN", "Stock Name"]].to_dict(orient="records"),
    )


# ─── VAL-12 ISIN_VALID gate (WARN) ──────────────────────────────────────────

def _val_12_isin_valid(th_val_df: pd.DataFrame) -> CheckResult:
    """Flag positions whose security carries ``tSecurity.ISIN_VALID='N'``.

    ISIN_VALID='N' means CCL ops marked the ISIN as suspect (e.g. composite
    ticker, custom internal SCODE, post-merger placeholder). These rows
    can still flow through the P&L, but they should be reviewed because
    downstream pricing systems may not recognise them.
    """
    if "ISIN_VALID" not in th_val_df.columns:
        return CheckResult(
            id="VAL-12", name="ISIN_VALID gate", status="PASS",
            detail="ISIN_VALID column not present in tH_VAL frame.",
        )
    bad = th_val_df.loc[
        th_val_df["ISIN_VALID"].astype("string").str.upper().fillna("Y") == "N"
    ]
    if bad.empty:
        return CheckResult(
            id="VAL-12", name="ISIN_VALID gate", status="PASS",
            detail=f"All {len(th_val_df)} positions have valid ISINs.",
        )
    cols = [c for c in ("ISIN", "SNAME", "CCY", "UNITS", "PTVALUE") if c in bad.columns]
    failures = bad[cols].to_dict(orient="records")
    return CheckResult(
        id="VAL-12", name="ISIN_VALID gate", status="WARN",
        detail=f"{len(bad)} position(s) flagged ISIN_VALID='N'.",
        failures=failures,
    )


def _val_13_zero_avg_buy_with_open_units(pl_df: pd.DataFrame) -> CheckResult:
    """Warn on rows with open units and trades but zero average buy price.

    This catches suspicious zero-cost injections that leave an open position
    with a zero average buy (for example, mispriced CA conversion rows).
    """
    if pl_df.empty:
        return CheckResult(
            id="VAL-13", name="Zero avg-buy with open units", status="PASS",
            detail="No rows to check.",
        )

    required = {"Ending Units", "Units Bought", "Avg Buy Price (EUR)", "ISIN", "Stock Name"}
    if not required.issubset(set(pl_df.columns)):
        return CheckResult(
            id="VAL-13", name="Zero avg-buy with open units", status="PASS",
            detail="Required columns missing; check skipped.",
        )

    ending = pd.to_numeric(pl_df["Ending Units"], errors="coerce").fillna(0.0)
    bought = pd.to_numeric(pl_df["Units Bought"], errors="coerce").fillna(0.0)
    avg_buy = pd.to_numeric(pl_df["Avg Buy Price (EUR)"], errors="coerce").fillna(0.0)

    bad = pl_df.loc[
        (ending.abs() > 1e-9)
        & (bought > 1e-9)
        & (avg_buy.abs() <= 1e-9)
    ]
    if bad.empty:
        return CheckResult(
            id="VAL-13", name="Zero avg-buy with open units", status="PASS",
            detail="No open positions with zero average buy price.",
        )

    cols = [c for c in ("ISIN", "Stock Name", "Ending Units", "Units Bought", "Avg Buy Price (EUR)") if c in bad.columns]
    return CheckResult(
        id="VAL-13", name="Zero avg-buy with open units", status="WARN",
        detail=f"{len(bad)} position(s) have open units with zero average buy price.",
        failures=bad[cols].to_dict(orient="records"),
    )


# ─── helpers ────────────────────────────────────────────────────────────────

def _sum_by_isin(trades_df: pd.DataFrame, *, t_code: str) -> dict[str, float]:
    if trades_df.empty or "T" not in trades_df.columns:
        return {}
    sub = trades_df.loc[trades_df["T"] == t_code]
    # Mirror compute.pl and derive_ca_auto: BCODE='ZZZZ' P/S rows are
    # corporate-action placeholders that get stripped before WAC. They
    # MUST also be stripped here, otherwise VAL-01's view diverges from
    # the WAC walk and from derive_ca_auto, leading to either spurious
    # PASS (when ZZZZ legs net to the snapshot delta) or double-correction
    # when auto CA fills the gap.
    if not sub.empty and "BCODE" in sub.columns:
        bcode_u = sub["BCODE"].astype(str).str.upper()
        sub = sub.loc[bcode_u != "ZZZZ"]
    if sub.empty:
        return {}
    from ..data.normalise import normalise_isin
    g = sub.groupby(sub["ISIN"].map(normalise_isin))["UNITS"].sum()
    return {k: float(v) for k, v in g.items() if k}
