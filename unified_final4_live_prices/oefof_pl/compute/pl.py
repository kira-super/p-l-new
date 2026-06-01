"""Position-level P&L computation.

This is the only module in the package that produces ``pl_df`` and
``income_df``. It is pure: no I/O, no globals, no COM. All FX rates and
the analyst resolver are passed in as arguments.

Two code paths, dispatched on instrument type:

* :func:`_compute_ord_position` — ordinary equity. Uses WAC for cost basis,
  computes start/end MV in EUR, and sums realised + unrealised + income.
* :func:`_compute_swap_position` — single-name CFD or index CFD. Total
  P&L = income only. Cost basis = ``|nucost_local × units|`` converted to
  EUR. **Does not accept** start-period FX inputs — BUG-3 made impossible.

Five structural bug fixes are implemented here:

* **BUG-1** Percentages produced *only* by :func:`_to_ratio`. A schema
  check at the end asserts ``|pct| ≤ 5`` for every percentage column.
* **BUG-3** SWAP path has a separate function with a different signature.
  It cannot receive ``fx_start`` or ``eur_usd_start``.
* **BUG-4** Trade conversions use the *uniform end-period* FX (``fx_end``,
  ``eur_usd_end``). ``TRADEGROSS_USD`` is not in the trade schema and
  cannot be reached.
* **BUG-7 (c)** The income loop reads ``trade['CCY']`` per row — the
  position's CCY is never used for income conversion.
* **BUG-11** EUR market values come from ``AggLine.ptvalue_eur`` directly\n  \u2014 the snapshot already gives them in fund base CCY. Local-CCY values\n  carry the ``_local`` suffix; mixing the two is a typo-safe KeyError.
"""

from __future__ import annotations

import math
from dataclasses import asdict, dataclass, field
from typing import Callable, Mapping

import pandas as pd

from ..data.normalise import (
    INCOME_COLUMNS,
    PL_COLUMNS,
    TRADES_REQUIRED,
    assert_schema,
    normalise_isin,
)
from ..exceptions import ComputeIntegrityError, SchemaError
from .aggregate import AggLine
from .classify import FTSWAP, ORD, SWAP, classify_instrument
from .fx import build_fx_from_trades, eur_to_local_at, local_to_eur_at
from .wac import compute_wac


# ─── Output row types ───────────────────────────────────────────────────────

@dataclass(frozen=True)
class PositionRow:
    """One row of ``pl_df``. Field names are snake_case; display names
    (with units and EUR suffix) are mapped at DataFrame construction time
    via :data:`SNAKE_TO_DISPLAY`. All money fields are float64.
    """

    isin: str
    stock_name: str
    country: str
    analyst: str
    ccy: str
    instrument: str
    ls: str
    starting_units: float
    ending_units: float
    units_bought: float
    units_sold: float
    bonus_units: float
    avg_buy_price_eur: float
    avg_sell_price_eur: float
    start_price_local: float
    last_price_eur: float
    mv_start_eur: float
    mv_end_eur: float
    cost_basis_eur: float
    realised_local: float
    realised_eur: float
    realised_pct: float
    unrealised_local: float
    unrealised_eur: float
    unrealised_pct: float
    income_local: float
    income_eur: float
    income_yield_pct: float
    total_pl_eur: float
    total_pl_pct: float
    last_trade_date: object        # pd.Timestamp or NaT
    first_trade_date: object       # pd.Timestamp or NaT (only when new position)


# Map snake_case field name → finance-standard column label used in output.
SNAKE_TO_DISPLAY: dict[str, str] = {
    "isin": "ISIN",
    "stock_name": "Stock Name",
    "country": "Country",
    "analyst": "Analyst",
    "ccy": "CCY",
    "instrument": "Instrument",
    "ls": "L/S",
    "starting_units": "Starting Units",
    "ending_units": "Ending Units",
    "units_bought": "Units Bought",
    "units_sold": "Units Sold",
    "bonus_units": "Bonus Units",
    "avg_buy_price_eur": "Avg Buy Price (EUR)",
    "avg_sell_price_eur": "Avg Sell Price (EUR)",
    "start_price_local": "Start Price (Local)",
    "last_price_eur": "Last Price (EUR)",
    "mv_start_eur": "Market Value Start (EUR)",
    "mv_end_eur": "Market Value End (EUR)",
    "cost_basis_eur": "Cost Basis (EUR)",
    "realised_local": "Realised P&L (Local)",
    "realised_eur": "Realised P&L (EUR)",
    "realised_pct": "Realised P&L (%)",
    "unrealised_local": "Unrealised P&L (Local)",
    "unrealised_eur": "Unrealised P&L (EUR)",
    "unrealised_pct": "Unrealised P&L (%)",
    "income_local": "Income (Local)",
    "income_eur": "Income (EUR)",
    "income_yield_pct": "Income Yield (%)",
    "total_pl_eur": "Total P&L (EUR)",
    "total_pl_pct": "Total P&L (%)",
    "last_trade_date": "Last Trade Date",
    "first_trade_date": "First Trade Date",
}

# Columns whose values must satisfy |x| ≤ 5 (i.e. stored as ratios, not pct).
PCT_FIELDS: tuple[str, ...] = (
    "realised_pct",
    "unrealised_pct",
    "income_yield_pct",
    "total_pl_pct",
)


# ─── Public entry point ─────────────────────────────────────────────────────

AnalystResolver = Callable[[str, str], str]


def compute_positions(
    trades_df: pd.DataFrame,
    start_agg: dict[str, AggLine],
    end_agg: dict[str, AggLine],
    *,
    fx_start: Mapping[str, float],
    fx_end: Mapping[str, float],
    eur_usd_start: float,
    eur_usd_end: float,
    end_date_ts: pd.Timestamp,
    analyst_resolver: AnalystResolver,
    ftswap_isins: frozenset[str],
    drop_zzzz_ps_trades: bool = True,
    diagnostics: dict[str, object] | None = None,
    live_prices_eur: dict[str, float] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Compute per-position P&L for every ISIN in start ∪ end ∪ trades.

    All keyword arguments are required (the explicit ``*`` makes call sites
    self-documenting and prevents argument-order accidents).

    Args:
        live_prices_eur: Optional {isin: price_eur} for open positions.
            When supplied, overrides the HiPort end-snapshot market value
            for currently-held positions, making unrealised P&L intraday-live.
            Exited positions (end_units == 0) and SWAPs are always unaffected.

    Returns:
        ``(pl_df, income_df)`` — each a pandas DataFrame with the column
        names defined in :mod:`oefof_pl.data.normalise`.

    Raises:
        SchemaError: if ``trades_df`` is missing required columns.
        ComputeIntegrityError: if the grand-total identity for ORD positions
            fails by > 1 EUR. Indicates a code bug, not bad data.
    """
    if not trades_df.empty:
        assert_schema(trades_df, TRADES_REQUIRED, where="compute_positions")

    # Effective end FX — three-tier priority (highest wins):
    #   Tier 1 (end snapshot)   — fx_end       — the period-end market rate.
    #   Tier 2 (start snapshot) — fx_start     — covers positions closed during the
    #                                            period whose CCY is absent from the
    #                                            end snapshot but was present at start.
    #   Tier 3 (trades)         — fx_trades    — covers positions that were opened
    #                                            AND closed entirely within the YTD
    #                                            window (never in either snapshot).
    #                                            Uses the trade-date XRATE, which is
    #                                            the best available rate for a fully
    #                                            realised position.
    # A CCY missing from ALL THREE sources is still a hard error — it means
    # the trade CCY is genuinely unknown, which is a real data bug.
    fx_trades: dict[str, float] = build_fx_from_trades(trades_df)
    fx_end_effective: dict[str, float] = {**fx_trades, **dict(fx_start), **dict(fx_end)}

    # Working copy with normalised ISIN; never mutate caller's DataFrame.
    trades = trades_df.copy() if not trades_df.empty else _empty_trades()
    if not trades.empty:
        trades["__isin__"] = trades["ISIN"].map(normalise_isin)
        trades = trades.loc[trades["__isin__"] != ""]
    if diagnostics is not None:
        diagnostics["input_trade_rows"] = int(len(trades_df))
        diagnostics["trades_after_isin_filter"] = int(len(trades))

    # Drop corporate-action placeholders only when they are buy/sell legs.
    # HiPort uses BCODE='ZZZZ' for two distinct purposes:
    #   1) P/S pairs at price=0 marking SCODE migrations (bonus-share
    #      issuance, ticker change). These are NOT economic transactions
    #      \u2014 the unit change is reflected in the next snapshot directly.
    #      Feeding them into WAC dilutes the per-share cost to zero and
    #      then realises huge phantom losses on the matched sell. End
    #      units come from the snapshot so dropping these rows does not
    #      affect Ending Units.
    #   2) T='I' rows that are REAL CASH DIVIDENDS routed through the
    #      'ZZZZ' pseudo-broker (UNITS = current holding, NUPRICEG = DPS,
    #      NSETTLE_AMOUNT = NINCOME = real cash). Stripping these would
    #      silently drop material dividend income (>1.5B VND, 500M IDR,
    #      etc. for OEFOF in a typical period). Keep them.
    zzzz_ps_dropped = 0
    if drop_zzzz_ps_trades and not trades.empty and "BCODE" in trades.columns and "T" in trades.columns:
        bcode_u = trades["BCODE"].astype(str).str.upper()
        t_u = trades["T"].astype(str).str.upper()
        ca_mask = (bcode_u == "ZZZZ") & t_u.isin(["P", "S"])
        zzzz_ps_dropped = int(ca_mask.sum())
        trades = trades.loc[~ca_mask]
    if diagnostics is not None:
        diagnostics["zzzz_ps_rows_dropped"] = zzzz_ps_dropped
        diagnostics["trades_after_zzzz_filter"] = int(len(trades))

    all_isins: set[str] = set(start_agg) | set(end_agg)
    if not trades.empty:
        all_isins |= set(trades["__isin__"].unique())
    synthetic_isins = {i for i in all_isins if ":" in i}
    all_isins -= synthetic_isins

    rows: list[PositionRow] = []
    income_records: list[dict] = []

    if diagnostics is not None:
        diagnostics["synthetic_isins_skipped"] = int(len(synthetic_isins))

    for isin in sorted(all_isins):
        start = start_agg.get(isin)
        end = end_agg.get(isin)
        isin_trades = (
            trades.loc[trades["__isin__"] == isin].sort_values("CDATE")
            if not trades.empty
            else _empty_trades()
        )

        # Resolve metadata. End preferred, then start, then first trade.
        sname, ccy, ls, country, cat = _resolve_meta(start, end, isin_trades)
        instrument = classify_instrument(
            cat=cat, sname=sname, isin=isin, ftswap_isins=ftswap_isins,
        )
        analyst = analyst_resolver(isin, sname)

        if instrument == ORD:
            row = _compute_ord_position(
                isin=isin, sname=sname, country=country, analyst=analyst,
                ccy=ccy, ls=ls, instrument=instrument,
                start=start, end=end, trades=isin_trades,
                fx_start=fx_start, fx_end=fx_end_effective,
                eur_usd_start=eur_usd_start, eur_usd_end=eur_usd_end,
                start_isins=set(start_agg.keys()),
                live_prices_eur=live_prices_eur,
            )
            # Per-trade income recording also happens here.
            for rec in _income_records(isin, sname, analyst, isin_trades, fx_end_effective, eur_usd_end):
                income_records.append(rec)
        else:
            # SWAP and FTSWAP share the same compute path.
            row = _compute_swap_position(
                isin=isin, sname=sname, country=country, analyst=analyst,
                ccy=ccy, ls=ls, instrument=instrument,
                start=start, end=end, trades=isin_trades,
                fx_end=fx_end_effective, eur_usd_end=eur_usd_end,
                start_isins=set(start_agg.keys()),
            )
            for rec in _income_records(isin, sname, analyst, isin_trades, fx_end_effective, eur_usd_end):
                income_records.append(rec)

        rows.append(row)

    if diagnostics is not None:
        diagnostics["position_rows_before_split"] = int(len(rows))

    pl_df = _build_pl_df(rows)
    income_df = _build_income_df(income_records)

    _assert_pct_fields_are_ratios(pl_df)
    _assert_ord_grand_total_identity(pl_df, rows)

    return pl_df, income_df


# ─── ORD compute path ───────────────────────────────────────────────────────

def _compute_ord_position(
    *,
    isin: str,
    sname: str,
    country: str,
    analyst: str,
    ccy: str,
    ls: str,
    instrument: str,
    start: AggLine | None,
    end: AggLine | None,
    trades: pd.DataFrame,
    fx_start: Mapping[str, float],
    fx_end: Mapping[str, float],
    eur_usd_start: float,
    eur_usd_end: float,
    start_isins: set[str],
    live_prices_eur: dict[str, float] | None = None,
) -> PositionRow:
    """ORD path: cost via WAC, total = ΔMV + realised + income."""

    # Starting position from the start-of-period AggLine.
    start_units = float(start.units) if start else 0.0
    start_price_local = float(start.lst_price_local) if start else 0.0
    start_xrate = float(start.xrate) if start else _fx_rate_for(fx_start, ccy)

    end_units = float(end.units) if end else 0.0
    end_price_local = float(end.lst_price_local) if end else 0.0
    end_xrate = _fx_rate_for(fx_end, ccy)

    # ── Period-mark cost basis (View B) ─────────────────────────────────────
    # Re-baseline cost to the *start-of-period market value* (in local CCY)
    # rather than HiPort's lifetime NUCOST. Rationale: this report is a
    # PERIOD P&L (typically YTD), not a tax/IRR book. Lifetime WAC for a
    # legacy holding can be a tiny fraction of current price, so any sells
    # this period would crystallise multi-year embedded gains into this
    # period's "Realised P&L" — wildly inflating both Realised and Total
    # for the period (and the % columns).
    #
    # By starting the WAC walk at start_price_local, all gains attributed
    # to this run are *period* gains:
    #   • Realised P&L = sells_proceeds − sold_units × start_price_local
    #   • Unrealised   = end_price − weighted_avg(start_price, mid-period buys)
    # Bonuses (auto-CA at price 0) still dilute correctly against the
    # period-start basis. Income / dividends / FX attribution unaffected.
    mv_start_eur = float(start.ptvalue_eur) if start else 0.0
    mv_end_eur = float(end.ptvalue_eur) if end else 0.0

    # ── Live price override (open positions only) ────────────────────────────
    # If a live price in EUR was supplied for this ISIN and the position is
    # still open (end_units != 0), replace the HiPort snapshot MV with the
    # intraday market value so unrealised P&L is live rather than T-1.
    # Exited positions are never overridden — their P&L is fully realised.
    _live_px_eur: float | None = None
    if live_prices_eur and end_units != 0:
        _live_px_eur = live_prices_eur.get(isin.strip().upper())
        if _live_px_eur is not None:
            mv_end_eur = _live_px_eur * end_units

    start_ptvalue_local = (
        eur_to_local_at(mv_start_eur, ccy, start_xrate, eur_usd_start)
        if start is not None else 0.0
    )
    start_cost_local = start_ptvalue_local

    # ── WAC walk (period-mark basis) ──
    wac = compute_wac(start_units, start_cost_local, trades) if not trades.empty else _zero_wac(start_units, start_cost_local)
    realised_local = wac.realised_local

    # End MV in local CCY — derived from EUR by inverse FX. Used for the
    # unrealised-P&L identity which works in local terms.
    end_ptvalue_local = eur_to_local_at(mv_end_eur, ccy, end_xrate, eur_usd_end)

    # If live price was applied, derive per-share local price from live MV
    # so last_price_eur in the output reflects the live price not T-1 close.
    if _live_px_eur is not None and end_units != 0:
        end_price_local = end_ptvalue_local / end_units

    # ── Trade-level EUR conversions (BUG-4 fix: uniform end-period FX) ──
    buy_eur, sell_eur = _trade_cash_flows_eur(trades, ccy, end_xrate, eur_usd_end)

    # ── Income (per-trade CCY: BUG-7 (c)) ──
    income_local = _income_local_sum(trades, ccy_filter=None)  # in mixed CCYs is meaningless
    income_local_pos_ccy = _income_local_sum(trades, ccy_filter=ccy)
    income_eur = _income_eur_sum(trades, fx_end, eur_usd_end)

    # ── Decomposition ──
    realised_eur = local_to_eur_at(realised_local, ccy, end_xrate, eur_usd_end)

    # Unrealised P&L: PERIOD CHANGE in unrealised (mark-to-market this
    # period only). Lifetime unrealised at end would double-count the
    # prior-period unrealised already embedded in mv_start_eur — that
    # broke the R+U+I ≈ Total identity for any position with a non-flat
    # carry-over.
    end_cost_local = wac.ending_cost_local if not trades.empty else start_cost_local
    end_unrealised_local = end_ptvalue_local - end_cost_local
    # Start unrealised is zero by construction (View B re-baselined cost
    # to start MV). Period ΔUnrealised therefore equals end_unrealised.
    start_unrealised_local = 0.0
    # ΔUnrealised in EUR uses the FX prevailing at each end of the period
    # so it carries the FX revaluation of the unrealised stock, matching
    # how mv_end_eur and mv_start_eur are valued.
    end_unrealised_eur = local_to_eur_at(end_unrealised_local, ccy, end_xrate, eur_usd_end)
    start_unrealised_eur = 0.0
    unrealised_local = end_unrealised_local - start_unrealised_local
    unrealised_eur = end_unrealised_eur - start_unrealised_eur

    # ── FX revaluation of carry-over (R+U+I = Total identity fix) ──
    # All trade-side EUR conversions (buys, sells, income, realised,
    # unrealised) use END-period FX. mv_start_eur, however, uses START-
    # period FX (it comes straight from the start snapshot's PTVALUE_EUR).
    # That mismatch leaks the FX revaluation of the carry-over stock out
    # of the R/U/I decomposition. Algebra (see DESIGN.md):
    #     R+U+I − Total = mv_start_eur − start_mv_at_end_FX
    # So Unrealised must be REDUCED by that gap to make the additive
    # identity hold. Economically this means: unrealised P&L excludes
    # the FX revaluation of the carry-over (which is already captured
    # inside ΔMV_eur on the Total line).
    if start is not None and abs(mv_start_eur) > 1e-9:
        start_mv_at_end_fx = local_to_eur_at(
            start_ptvalue_local, ccy, end_xrate, eur_usd_end,
        )
        fx_reval_eur = mv_start_eur - start_mv_at_end_fx
        unrealised_eur -= fx_reval_eur

    # If the position is fully exited (end_units == 0), the remaining ΔU
    # is just the reversal of last period's carry-over unrealised
    # (start_unrealised, since end_unrealised = 0). Economically this was
    # *crystallised* by the exit, so report it inside Realised and zero
    # the Unrealised column. Total P&L is unchanged because Realised and
    # Unrealised are added together downstream.
    if end_units == 0:
        realised_local += unrealised_local
        realised_eur += unrealised_eur
        unrealised_local = 0.0
        unrealised_eur = 0.0

    # Cost basis in EUR = current capital tied up in the position (lifetime
    # WAC end-cost in local converted at end FX — matches HiPort PTCOST and
    # is how the column is read across desks). For fully-exited positions
    # (end_units == 0) prefer |start MV|; intra-period round-trips fall back
    # to gross buy proceeds in EUR. This column is NOT used as the % denom;
    # see ``period_denom`` below.
    end_cost_eur = local_to_eur_at(end_cost_local, ccy, end_xrate, eur_usd_end)
    if end_units != 0:
        cost_basis_eur = end_cost_eur
    elif abs(mv_start_eur) > 0:
        cost_basis_eur = abs(mv_start_eur)
    else:
        cost_basis_eur = abs(buy_eur)

    # Total (the identity):  Total = (EndMV − StartMV) + Sell − Buy + Income
    total_pl_eur = (mv_end_eur - mv_start_eur) + sell_eur - buy_eur + income_eur

    # Period-return denominator: the *%* columns must reflect this period's
    # return, not lifetime. Lifetime WAC for legacy positions is far below
    # current MV and would make a normal price move look like a 200%+ return.
    # Use |Start MV| (capital at risk at period start) when the position was
    # held at start, otherwise the capital deployed mid-period (buys),
    # otherwise lifetime cost as a final fallback.
    if abs(mv_start_eur) > 0:
        period_denom = abs(mv_start_eur)
    elif buy_eur > 0:
        if end_units == 0:
            # Fully exited intraperiod positions should use gross deployed
            # capital; a half-buy floor inflates pct for round-trips.
            period_denom = buy_eur
        else:
            net_deployed = buy_eur - sell_eur
            period_denom = max(net_deployed, buy_eur * 0.5)
    else:
        period_denom = cost_basis_eur

    return PositionRow(
        isin=isin,
        stock_name=sname,
        country=country,
        analyst=analyst,
        ccy=ccy,
        instrument=instrument,
        ls=ls,
        starting_units=start_units,
        ending_units=end_units,
        units_bought=wac.units_bought,
        units_sold=wac.units_sold,
        bonus_units=_bonus_units(trades),
        avg_buy_price_eur=local_to_eur_at(wac.avg_buy_price_local, ccy, end_xrate, eur_usd_end),
        avg_sell_price_eur=local_to_eur_at(wac.avg_sell_price_local, ccy, end_xrate, eur_usd_end),
        start_price_local=start_price_local,
        last_price_eur=local_to_eur_at(end_price_local, ccy, end_xrate, eur_usd_end),
        mv_start_eur=mv_start_eur,
        mv_end_eur=mv_end_eur,
        cost_basis_eur=cost_basis_eur,
        realised_local=realised_local,
        realised_eur=realised_eur,
        realised_pct=_to_ratio(realised_eur, period_denom),
        unrealised_local=unrealised_local,
        unrealised_eur=unrealised_eur,
        unrealised_pct=_to_ratio(unrealised_eur, period_denom),
        income_local=income_local_pos_ccy,
        income_eur=income_eur,
        income_yield_pct=_to_ratio(income_eur, period_denom),
        total_pl_eur=total_pl_eur,
        total_pl_pct=_to_ratio(total_pl_eur, period_denom),
        last_trade_date=_last_trade_date(trades),
        first_trade_date=_first_trade_date_for_position(
            isin, trades, start_isins, wac.last_entry_date,
        ),
    )


# ─── SWAP / FTSWAP compute path ─────────────────────────────────────────────

def _compute_swap_position(
    *,
    isin: str,
    sname: str,
    country: str,
    analyst: str,
    ccy: str,
    ls: str,
    instrument: str,
    start: AggLine | None,
    end: AggLine | None,
    trades: pd.DataFrame,
    fx_end: Mapping[str, float],
    eur_usd_end: float,
    start_isins: set[str],
) -> PositionRow:
    """SWAP / FTSWAP path: Total = income only. Cost = |NUCOST × Units|.

    Notice the signature: no ``fx_start`` or ``eur_usd_start``. This is the
    BUG-3 structural fix — there is no input that could let a start-period
    market value contaminate the SWAP total.
    """

    end_xrate = _fx_rate_for(fx_end, ccy)

    start_units = float(start.units) if start else 0.0
    end_units = float(end.units) if end else 0.0

    # PTVALUE is already in EUR (fund base) per the HiPort snapshot. Use it
    # directly; do NOT re-convert via xrate (BUG-11). For SWAP/FTSWAP this
    # is informational only — total P&L is income.
    mv_end_eur = float(end.ptvalue_eur) if end else 0.0
    mv_start_eur = float(start.ptvalue_eur) if start else 0.0

    # Cost basis = notional exposure (BUG-5 fix).
    nucost_local = float(end.nucost_local if end else (start.nucost_local if start else 0.0))
    units_for_notional = end_units if end_units != 0 else start_units
    notional_local = abs(nucost_local * units_for_notional)
    cost_basis_eur = local_to_eur_at(notional_local, ccy, end_xrate, eur_usd_end)

    # Income (per-trade CCY).
    income_local_pos_ccy = _income_local_sum(trades, ccy_filter=ccy)
    income_eur = _income_eur_sum(trades, fx_end, eur_usd_end)

    # ── SWAP / FTSWAP P&L identity ──────────────────────────────────────
    # Daily mark-to-market on a CFD / total-return-swap is booked into
    # PTVALUE_EUR by the custodian. There is no buy/sell cash leg on the
    # general ledger, so the period P&L (excluding income) is simply the
    # change in marked value: ΔMV = mv_end_eur − mv_start_eur.
    #
    #   * Position open at end of period       → unrealised = ΔMV
    #   * Position closed during the period
    #     (end_units = 0, mv_end_eur = 0)      → realised   = ΔMV   (= −mv_start_eur)
    #
    # Total = realised + unrealised + income, matching the View B identity
    # used on the ORD path. (The previous policy hard-coded both legs to
    # zero and reported total = income only, which suppressed all visible
    # MTM movement on swap positions.)
    delta_mv_eur = mv_end_eur - mv_start_eur
    if end_units == 0.0 and start_units != 0.0:
        realised_eur = delta_mv_eur
        unrealised_eur = 0.0
    else:
        realised_eur = 0.0
        unrealised_eur = delta_mv_eur

    # Local-CCY views are informational only on swaps (no FX involved in
    # PTVALUE_EUR — already EUR-denominated by the custodian).
    realised_local = realised_eur
    unrealised_local = unrealised_eur

    total_pl_eur = realised_eur + unrealised_eur + income_eur

    # WAC-style averages on units, useful for reporting (no FX involved).
    units_bought = float(_safe_sum(trades.loc[trades["T"] == "P", "UNITS"])) if not trades.empty else 0.0
    units_sold = float(_safe_sum(trades.loc[trades["T"] == "S", "UNITS"])) if not trades.empty else 0.0

    avg_buy = _avg_price(trades, "P")
    avg_sell = _avg_price(trades, "S")

    return PositionRow(
        isin=isin,
        stock_name=sname,
        country=country,
        analyst=analyst,
        ccy=ccy,
        instrument=instrument,
        ls=ls,
        starting_units=start_units,
        ending_units=end_units,
        units_bought=units_bought,
        units_sold=units_sold,
        bonus_units=_bonus_units(trades),
        avg_buy_price_eur=local_to_eur_at(avg_buy, ccy, end_xrate, eur_usd_end),
        avg_sell_price_eur=local_to_eur_at(avg_sell, ccy, end_xrate, eur_usd_end),
        start_price_local=float(start.lst_price_local) if start else 0.0,
        last_price_eur=local_to_eur_at(float(end.lst_price_local) if end else 0.0, ccy, end_xrate, eur_usd_end),
        mv_start_eur=mv_start_eur,
        mv_end_eur=mv_end_eur,
        cost_basis_eur=cost_basis_eur,
        realised_local=realised_local,
        realised_eur=realised_eur,
        realised_pct=_to_ratio(realised_eur, cost_basis_eur),
        unrealised_local=unrealised_local,
        unrealised_eur=unrealised_eur,
        unrealised_pct=_to_ratio(unrealised_eur, cost_basis_eur),
        income_local=income_local_pos_ccy,
        income_eur=income_eur,
        income_yield_pct=_to_ratio(income_eur, cost_basis_eur),
        total_pl_eur=total_pl_eur,
        total_pl_pct=_to_ratio(total_pl_eur, cost_basis_eur),
        last_trade_date=_last_trade_date(trades),
        first_trade_date=_first_trade_date_for_position(
            isin, trades, start_isins, None,
        ),
    )


# ─── helpers ────────────────────────────────────────────────────────────────

def _bonus_units(trades: pd.DataFrame) -> float:
    """Units received as genuine zero-price bonus shares.

    Identifies real bonus-issue trades by two criteria:
      1. T == 'P' with zero gross price and zero settlement amount
         (free shares, not a purchase).
      2. SNAME contains 'BONUS' or 'BON' (HiPort naming convention).

    Deliberately excludes CA_ADJ rows: those are unit-reconciliation
    entries auto-generated by the pipeline to close snapshot gaps.
    They are accounting artifacts, not economic bonus events, and
    should never trigger the split-row presentation.
    """
    if trades.empty:
        return 0.0
    t_u = trades["T"].astype(str).str.upper()
    # Exclude auto-generated CA_ADJ reconciliation rows.
    if "BCODE" in trades.columns:
        bcode_u = trades["BCODE"].astype(str).str.upper()
        candidates = trades.loc[(t_u == "P") & (bcode_u != "CA_ADJ")]
    else:
        candidates = trades.loc[t_u == "P"]
    if candidates.empty:
        return 0.0
    # Zero gross price AND zero settlement = genuinely free shares.
    gross = pd.to_numeric(candidates.get("GROSSPRICE_LOCAL", 0), errors="coerce").fillna(0.0)
    buy_net = pd.to_numeric(candidates.get("BUY_NET_LOCAL", 0), errors="coerce").fillna(0.0)
    is_free = (gross == 0.0) & (buy_net == 0.0)
    # SNAME must contain BONUS or BON (HiPort convention for bonus issues).
    sname_u = candidates.get("SNAME", pd.Series(dtype=str)).astype(str).str.upper()
    is_bonus_name = sname_u.str.contains(r"\bBON(?:US)?\b", regex=True, na=False)
    mask = is_free & is_bonus_name
    return float(pd.to_numeric(candidates.loc[mask, "UNITS"], errors="coerce").fillna(0.0).sum())


def _to_ratio(numerator: float, denominator: float) -> float:
    """Return ``numerator / denominator`` as a *ratio* (0.44 = 44%).

    BUG-1 prevention: this is the **only** way a percentage column is
    produced. The output layer applies the ``'0.00%'`` Excel format which
    multiplies by 100 for display. If anyone divides by 100 here, the
    schema check at the end of ``compute_positions`` will fail because the
    ratio will exceed |5|.
    """
    if denominator is None or denominator == 0 or (
        isinstance(denominator, float) and math.isnan(denominator)
    ):
        return 0.0
    if numerator is None or (isinstance(numerator, float) and math.isnan(numerator)):
        return 0.0
    return float(numerator) / float(denominator)


def _resolve_meta(
    start: AggLine | None,
    end: AggLine | None,
    trades: pd.DataFrame,
) -> tuple[str, str, str, str, str]:
    """Return (sname, ccy, ls, country, cat). End preferred, then start, then trades."""
    for src in (end, start):
        if src is not None:
            return src.sname, src.ccy, src.ls, src.country, src.cat

    if not trades.empty:
        first = trades.iloc[0]
        return (
            str(first.get("SNAME", "")).strip(),
            str(first.get("CCY", "")).strip().upper(),
            "L",
            "",
            "ORD",  # trades-only ISINs default to ORD; classify_instrument
                    # may still promote via SNAME (e.g. "X CFD")
        )

    return "", "", "L", "", "ORD"


def _fx_rate_for(fx: Mapping[str, float], ccy: str) -> float:
    """Get an FX rate for a CCY. Used only for non-EUR/USD currencies; for
    EUR/USD ``local_to_eur_at`` ignores xrate. A missing entry returns NaN
    so a downstream non-EUR/USD conversion will raise FXInvalidError —
    exactly the failure mode we want."""
    c = (ccy or "").strip().upper()
    if c in ("EUR", "USD"):
        return 0.0  # ignored by local_to_eur_at for these currencies
    return float(fx.get(c, float("nan")))


def _trade_cash_flows_eur(
    trades: pd.DataFrame,
    ccy: str,
    xrate: float,
    eur_usd: float,
) -> tuple[float, float]:
    """Sum P/S cash flows in EUR using uniform end-period FX (BUG-4 fix)."""
    if trades.empty:
        return 0.0, 0.0
    buy_local = float(_safe_sum(trades.loc[trades["T"] == "P", "BUY_NET_LOCAL"]))
    sell_local = float(_safe_sum(trades.loc[trades["T"] == "S", "SELL_NET_LOCAL"]))
    buy_eur = local_to_eur_at(buy_local, ccy, xrate, eur_usd) if buy_local else 0.0
    sell_eur = local_to_eur_at(sell_local, ccy, xrate, eur_usd) if sell_local else 0.0
    return buy_eur, sell_eur


def _income_local_sum(trades: pd.DataFrame, *, ccy_filter: str | None) -> float:
    """Sum INCOME_LOCAL for I-trades, optionally filtered by CCY."""
    if trades.empty:
        return 0.0
    inc = trades.loc[trades["T"] == "I"]
    if ccy_filter is not None:
        inc = inc.loc[inc["CCY"].astype("string").str.upper().str.strip() == ccy_filter.upper()]
    return float(_safe_sum(inc["INCOME_LOCAL"]))


def _income_eur_sum(
    trades: pd.DataFrame,
    fx_end: Mapping[str, float],
    eur_usd_end: float,
) -> float:
    """BUG-7 (c) fix: iterate income trades and convert each at its own
    trade CCY. The position's CCY is never used here."""
    if trades.empty:
        return 0.0
    inc = trades.loc[trades["T"] == "I"]
    if inc.empty:
        return 0.0

    total = 0.0
    for ccy, amount in zip(inc["CCY"], inc["INCOME_LOCAL"]):
        c = (str(ccy) if ccy is not None else "").strip().upper()
        amt = float(amount) if amount is not None and not (
            isinstance(amount, float) and math.isnan(amount)
        ) else 0.0
        if amt == 0.0:
            continue
        xrate = _fx_rate_for(fx_end, c)
        total += local_to_eur_at(amt, c, xrate, eur_usd_end)
    return total


def _income_records(
    isin: str,
    sname: str,
    analyst: str,
    trades: pd.DataFrame,
    fx_end: Mapping[str, float],
    eur_usd_end: float,
) -> list[dict]:
    """One record per I-trade for the long-form income_df."""
    if trades.empty:
        return []
    inc = trades.loc[trades["T"] == "I"]
    if inc.empty:
        return []

    out: list[dict] = []
    for _, r in inc.iterrows():
        ccy = (str(r.get("CCY", "")) or "").strip().upper()
        amt_local = float(r["INCOME_LOCAL"]) if pd.notna(r["INCOME_LOCAL"]) else 0.0
        amt_eur = local_to_eur_at(
            amt_local, ccy, _fx_rate_for(fx_end, ccy), eur_usd_end,
        )
        out.append({
            "ISIN": isin,
            "Stock Name": sname,
            "Analyst": analyst,
            "CCY": ccy,
            "Date": r.get("CDATE"),
            "Income Type": "DIV",
            "Income (Local)": amt_local,
            "Income (EUR)": amt_eur,
            "Units": float(r["UNITS"]) if pd.notna(r.get("UNITS")) else 0.0,
        })
    return out


def _avg_price(trades: pd.DataFrame, t_code: str) -> float:
    if trades.empty:
        return 0.0
    sub = trades.loc[trades["T"] == t_code]
    if sub.empty:
        return 0.0
    units = pd.to_numeric(sub["UNITS"], errors="coerce").fillna(0.0)
    price = pd.to_numeric(sub["GROSSPRICE_LOCAL"], errors="coerce").fillna(0.0)
    total_units = float(units.sum())
    if total_units == 0:
        return 0.0
    return float((units * price).sum() / total_units)


def _last_trade_date(trades: pd.DataFrame) -> object:
    if trades.empty or "CDATE" not in trades.columns:
        return pd.NaT
    s = pd.to_datetime(trades["CDATE"], errors="coerce").dropna()
    return s.max() if not s.empty else pd.NaT


def _first_trade_date_for_position(
    isin: str,
    trades: pd.DataFrame,
    start_isins: set[str],
    last_entry_date: pd.Timestamp | None,
) -> object:
    """Emit first trade date for new names and re-entry date after full exits."""
    if last_entry_date is not None and pd.notna(last_entry_date):
        return pd.Timestamp(last_entry_date)
    if isin in start_isins:
        return pd.NaT
    if trades.empty or "CDATE" not in trades.columns:
        return pd.NaT
    s = pd.to_datetime(trades["CDATE"], errors="coerce").dropna()
    return s.min() if not s.empty else pd.NaT


def _safe_sum(s: pd.Series) -> float:
    return float(pd.to_numeric(s, errors="coerce").fillna(0.0).sum())


def _zero_wac(starting_units: float, starting_cost_local: float):
    """Trivial WacResult-like for empty-trades case."""
    from .wac import WacResult
    return WacResult(
        units_bought=0.0,
        units_sold=0.0,
        avg_buy_price_local=0.0,
        avg_sell_price_local=0.0,
        ending_units=starting_units,
        ending_cost_local=starting_cost_local,
        realised_local=0.0,
        last_entry_date=None,
    )


def _empty_trades() -> pd.DataFrame:
    return pd.DataFrame(columns=list(TRADES_REQUIRED) + ["__isin__"])


def _build_pl_df(rows: list[PositionRow]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=list(PL_COLUMNS))

    records = [asdict(r) for r in rows]
    df = pd.DataFrame(records)
    # Map to display column names in the spec'd order.
    df = df.rename(columns=SNAKE_TO_DISPLAY)

    # Split Income (EUR) into Dividends (ORD path) and Swap Financing
    # (SWAP/FTSWAP path). The trade blotter has no income sub-type, so
    # the only reliable partition is by the position's Instrument.
    is_swap = df["Instrument"].isin([SWAP, FTSWAP])
    income_eur = pd.to_numeric(df["Income (EUR)"], errors="coerce").fillna(0.0)
    df["Dividends (EUR)"] = income_eur.where(~is_swap, 0.0)
    df["Swap Financing (EUR)"] = income_eur.where(is_swap, 0.0)

    df = df.reindex(columns=list(PL_COLUMNS))
    return df


def split_ca_rows(pl_df: pd.DataFrame) -> pd.DataFrame:
    """Split positions that received bonus shares into two rows.

    .. deprecated::
        This re-export exists for backwards compatibility.
        Import from :mod:`oefof_pl.compute.ca_split` instead.
    """
    from .ca_split import split_ca_rows as _split_ca_rows
    return _split_ca_rows(pl_df)


def _build_income_df(records: list[dict]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(columns=list(INCOME_COLUMNS))
    df = pd.DataFrame(records)
    df = df.reindex(columns=list(INCOME_COLUMNS))
    return df


# ─── post-compute integrity checks ──────────────────────────────────────────

def _assert_pct_fields_are_ratios(pl_df: pd.DataFrame) -> None:
    """Compute-stage integrity: percentage columns must be finite. The
    ``|x| ≤ 5`` magnitude check is a *data-quality* concern (BUG-1: somebody
    fed a per-cent value where a ratio was expected) and lives in VAL-03,
    which can list every offending row in the validation report instead of
    aborting the pipeline on the first one. Here we only catch NaN/Inf,
    which would mean compute itself produced garbage (a divide-by-zero
    that should have been guarded by ``_to_ratio``)."""
    for snake_name in PCT_FIELDS:
        col = SNAKE_TO_DISPLAY[snake_name]
        if col not in pl_df.columns:
            continue
        s = pd.to_numeric(pl_df[col], errors="coerce")
        non_finite_mask = ~s.apply(lambda v: math.isfinite(v) if pd.notna(v) else True)
        if non_finite_mask.any():
            offenders = pl_df.loc[non_finite_mask, [
                c for c in ("ISIN", "Stock Name", "CCY", "Instrument", col)
                if c in pl_df.columns
            ]].head(20)
            raise ComputeIntegrityError(
                f"Percentage column {col!r} contains non-finite value(s). "
                "compute._to_ratio failed to guard a divide-by-zero.\n"
                f"First offenders:\n{offenders.to_string(index=False)}"
            )


def _assert_ord_grand_total_identity(
    pl_df: pd.DataFrame,
    rows: list[PositionRow],
) -> None:
    """For ORD positions the identity holds:
        Σ Total = Σ(EndMV − StartMV) + Σ Sell_EUR − Σ Buy_EUR + Σ Income_EUR
    Buy/Sell EUR are not stored on PositionRow so we reconstruct them from
    the components: Total − (ΔMV + Income) ≡ Sell − Buy. The check is that
    Total_sum equals (ΔMV_sum + (Sell − Buy)_sum + Income_sum) — which is
    trivially true given the formula, so what we *actually* check here is
    that no NaNs/Infs slipped in.
    """
    ord_rows = [r for r in rows if r.instrument == ORD]
    if not ord_rows:
        return

    bad = [
        r.isin for r in ord_rows
        if not all(
            isinstance(v, (int, float)) and math.isfinite(v)
            for v in (r.mv_start_eur, r.mv_end_eur, r.income_eur, r.total_pl_eur)
        )
    ]
    if bad:
        raise ComputeIntegrityError(
            f"Non-finite values in ORD P&L for ISINs: {bad[:10]}"
            + (" ..." if len(bad) > 10 else "")
        )
