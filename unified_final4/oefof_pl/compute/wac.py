"""Weighted-average cost (WAC) tracker for ORD instruments.

The WAC is computed *in local currency*. EUR conversion happens later, in
``compute.pl``, using a single end-period FX rate per currency (BUG-4 fix).
This split is what prevents the legacy "phantom corporate-action gain" bug:
WAC has no FX inputs, so trade-day FX cannot leak into the cost basis.

Algorithm — per ISIN, in CDATE order:

* **P (purchase)**  Update WAC:
      new_units    = old_units + trade_units
      new_cost     = old_cost  + trade_units × gross_price_local
      new_wac      = new_cost / new_units            (if new_units != 0)

* **S (sell)**      Realise P&L at current WAC, do NOT change WAC:
      realised    += trade_units × (sell_price_local − wac)
      remaining    = old_units − trade_units
      cost_basis  *= remaining / old_units           (proportional)

* **I (income)**    *Ignored entirely* — income is not a position event.
                    This is the BUG-5 structural fix: income trades cannot
                    pollute WAC because they're filtered out before the loop.

The function takes pre-filtered, pre-sorted trades for a single ISIN. The
caller (``compute.pl``) is responsible for grouping. This keeps the
function pure, easy to unit-test, and free of any DataFrame indexing tricks.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

import pandas as pd

from ..exceptions import SchemaError


# Trade-type codes recognised by this module. Anything else is rejected so a
# new type added upstream cannot silently slip past WAC.
_BUY = "P"
_SELL = "S"
_INCOME = "I"
_KNOWN_TYPES = frozenset({_BUY, _SELL, _INCOME})


# Required columns on the per-ISIN trades DataFrame.
_REQUIRED_TRADE_COLS = ("T", "CDATE", "UNITS", "GROSSPRICE_LOCAL")


@dataclass(frozen=True)
class WacResult:
    """Result of walking one ISIN's trades through WAC.

    All monetary fields are in local currency. Conversion to EUR happens in
    ``compute.pl`` using the end-period FX rate.
    """

    units_bought: float          # sum of P units
    units_sold: float            # sum of S units
    avg_buy_price_local: float   # weighted by units, P trades only
    avg_sell_price_local: float  # weighted by units, S trades only
    ending_units: float          # starting + bought − sold
    ending_cost_local: float     # WAC × ending_units (≥0 by construction)
    realised_local: float        # Σ (sell_price − wac_at_sell) × sell_units
    last_entry_date: pd.Timestamp | None  # first buy date of the current open leg


def compute_wac(
    starting_units: float,
    starting_cost_local: float,
    trades_for_isin: pd.DataFrame,
) -> WacResult:
    """Walk one ISIN's trades through the WAC algorithm.

    Args:
        starting_units:      Units held at period start (signed: long > 0,
                             short < 0).
        starting_cost_local: Total cost basis at period start, in local CCY.
                             Sign matches ``starting_units``: a 100-unit long
                             position costing 50/share has cost = +5000; a
                             100-unit short with 50/share proceeds has cost
                             = −5000.
        trades_for_isin:     DataFrame of trades for *one* ISIN, already
                             sorted by ``CDATE`` ascending. Must contain
                             columns ``T``, ``CDATE``, ``UNITS``,
                             ``GROSSPRICE_LOCAL``. Income rows (T='I') are
                             ignored. ``UNITS`` is positive for both P and S
                             rows (direction is encoded by ``T``).

    Returns:
        WacResult with all per-ISIN aggregates needed by compute.pl.

    Raises:
        SchemaError: if a required column is missing.
        ValueError: if a trade has an unknown ``T`` value.
    """
    _assert_trade_schema(trades_for_isin)

    # Drop income rows up-front. They never affect units or WAC.
    pos_trades = trades_for_isin.loc[trades_for_isin["T"] != _INCOME]

    units = float(starting_units)
    cost = float(starting_cost_local)

    units_bought = 0.0
    units_sold = 0.0
    buy_notional = 0.0   # Σ units × price for P trades, in local
    units_bought_this_leg = 0.0
    buy_notional_this_leg = 0.0
    last_closed_avg_buy = 0.0
    sell_notional = 0.0  # Σ units × price for S trades, in local
    realised = 0.0
    last_entry_date: pd.Timestamp | None = None

    for t, cdate, qty, price in zip(
        pos_trades["T"].tolist(),
        pos_trades["CDATE"].tolist(),
        pos_trades["UNITS"].tolist(),
        pos_trades["GROSSPRICE_LOCAL"].tolist(),
    ):
        if t not in _KNOWN_TYPES:
            raise ValueError(f"Unknown trade type {t!r} (expected P/S/I)")

        q = _safe_float(qty)
        p = _safe_float(price)

        if t == _BUY:
            if abs(units) < 1e-9:
                ts = pd.to_datetime(cdate, errors="coerce")
                last_entry_date = ts if pd.notna(ts) else None
                units_bought_this_leg = 0.0
                buy_notional_this_leg = 0.0
            units_bought += q
            buy_notional += q * p
            units_bought_this_leg += q
            buy_notional_this_leg += q * p
            units += q
            cost += q * p

        elif t == _SELL:
            units_sold += q
            sell_notional += q * p

            if units == 0:
                # Selling into a flat book — treat as opening a short.
                # WAC of the short is the sell price.
                realised += 0.0
                units -= q
                cost -= q * p
            else:
                wac = cost / units  # WAC just before this sell
                realised += q * (p - wac)
                # Reduce position proportionally; cost stays at WAC × remaining.
                remaining = units - q
                if abs(units) < 1e-12:
                    cost = 0.0
                else:
                    cost = wac * remaining
                units = remaining
                if abs(units) < 1e-9:
                    units = 0.0
                    cost = 0.0
                    if units_bought_this_leg > 0:
                        last_closed_avg_buy = buy_notional_this_leg / units_bought_this_leg
                    units_bought_this_leg = 0.0
                    buy_notional_this_leg = 0.0

    if units_bought_this_leg:
        avg_buy = buy_notional_this_leg / units_bought_this_leg
    else:
        avg_buy = last_closed_avg_buy
    avg_sell = sell_notional / units_sold if units_sold else 0.0

    return WacResult(
        units_bought=units_bought,
        units_sold=units_sold,
        avg_buy_price_local=avg_buy,
        avg_sell_price_local=avg_sell,
        ending_units=units,
        ending_cost_local=cost,
        realised_local=realised,
        last_entry_date=last_entry_date,
    )


# ─── helpers ─────────────────────────────────────────────────────────────────

def _assert_trade_schema(df: pd.DataFrame) -> None:
    missing = [c for c in _REQUIRED_TRADE_COLS if c not in df.columns]
    if missing:
        raise SchemaError(
            f"compute_wac: trades_for_isin missing columns {missing}; "
            f"got {list(df.columns)}"
        )


def _safe_float(x: object) -> float:
    """Convert to float, treating None/NaN as 0.0."""
    if x is None:
        return 0.0
    try:
        f = float(x)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return 0.0
    if math.isnan(f):
        return 0.0
    return f
