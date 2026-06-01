"""Canonical schemas and normalisation helpers for source DataFrames.

This module is the **single source of truth** for what columns must exist in
the DataFrames that flow between layers, and for the renames that prevent
the legacy local-vs-EUR confusion (BUG-11) and per-trade-FX confusion
(BUG-4).

Two structural rules enforced here:

1. Source-name suffixes spell out the unit-of-measure so compute code
   cannot mix currencies.  In HiPort/HP_VAL the snapshot fields are NOT
   uniform: ``PTVALUE`` and ``PTCOST`` are pre-converted to the fund base
   currency (EUR), while ``NUCOST``, ``LST_PRICE``, ``BUY NET``, ``SELL NET``,
   ``GROSSPRICE`` and ``INCOME`` are in the position's local CCY. The
   loader renames accordingly to ``_EUR`` or ``_LOCAL`` so any compute
   code that reaches for the wrong unit gets a ``KeyError``. The bare
   strings (``PTVALUE`` etc.) do not survive past the loader.

2. ``TRADEGROSS_USD`` is not in the trade schema. It is dropped at load time
   so compute code cannot use it for FX conversion (it is the source of the
   phantom corporate-action gains in BUG-4).

The function ``assert_schema`` is the boundary check used by every consumer
of these DataFrames.
"""

from __future__ import annotations

from typing import Iterable

import pandas as pd

from ..exceptions import SchemaError


# ─── Portfolio snapshot schema ───────────────────────────────────────────────

# Source columns (HP_VAL!val_data) → canonical names. Anything not listed is
# either passed through unchanged (extra columns are allowed) or dropped at
# the loader's discretion. Critically, the local-currency monetary columns
# are renamed so downstream code cannot mistake them for EUR amounts.
PORT_RENAMES: dict[str, str] = {
    "PTVALUE":   "PTVALUE_EUR",
    "PTCOST":    "PTCOST_EUR",
    "NUCOST":    "NUCOST_LOCAL",
    "LST_PRICE": "LST_PRICE_LOCAL",
}

# Columns the loader guarantees to deliver, in this exact name and dtype
# family. Validated by `assert_schema(df, PORT_REQUIRED)`.
PORT_REQUIRED: tuple[str, ...] = (
    "ISIN",
    "SNAME",
    "CCY",
    "XRATE",
    "UNITS",
    "PTVALUE_EUR",
    "NUCOST_LOCAL",
    "LST_PRICE_LOCAL",
    "CAT",          # sole instrument-type input — see classify.py
    "LS",
    "PCODE_ORIG",
    "VDATE",        # authoritative valuation date — never derived from filename
    "EXCODE1",      # → Country
)

# Columns deliberately stripped from any portfolio DataFrame at load time.
# `ORD` is the always-'A' classification trap (BUG-9).
PORT_DROP: tuple[str, ...] = (
    "ORD",
)


# ─── Trade blotter schema ────────────────────────────────────────────────────

TRADES_RENAMES: dict[str, str] = {
    "ID":         "TRADE_ID",   # bottler row primary key — used by VAL-02 to detect true duplicates
    "PCODE":      "PCODE_ORIG",  # raw bottler column is PCODE, not PCODE_ORIG
    "GROSSPRICE": "GROSSPRICE_LOCAL",
    "NETPRICE":   "NETPRICE_LOCAL",
    "BUY NET":    "BUY_NET_LOCAL",
    "SELL NET":   "SELL_NET_LOCAL",
    "INCOME":     "INCOME_LOCAL",
    "BCOMM":      "BCOMM_LOCAL",
    "EXPENSES":   "EXPENSES_LOCAL",
    "SHORT_NAME": "SNAME",      # unify with portfolio snapshot
}

TRADES_REQUIRED: tuple[str, ...] = (
    "TRADE_ID",
    "PCODE_ORIG",
    "ISIN",
    "SNAME",
    "CCY",
    "T",            # P / S / I
    "CDATE",        # parsed datetime
    "UNITS",        # always positive in source
    "GROSSPRICE_LOCAL",
    "BUY_NET_LOCAL",
    "SELL_NET_LOCAL",
    "INCOME_LOCAL",
)

# Dropped at load time. `TRADEGROSS_USD` is the BUG-4 trap: any compute code
# that tries to reach it will get KeyError.
TRADES_DROP: tuple[str, ...] = (
    "TRADEGROSS_USD",
)


# ─── P&L output schema (final user-facing column names) ──────────────────────
# These names are finance-standard and client-presentable per the §18
# decisions. The internal `pl_df` returned by compute.pl uses these exact
# strings as column labels — no further translation in the output layer.

PL_COLUMNS: tuple[str, ...] = (
    "ISIN",
    "Stock Name",
    "Country",
    "Analyst",
    "CCY",
    "Instrument",
    "L/S",
    "Starting Units",
    "Ending Units",
    "Units Bought",
    "Units Sold",
    "Bonus Units",
    "Avg Buy Price (EUR)",
    "Avg Sell Price (EUR)",
    "Start Price (Local)",
    "Last Price (EUR)",
    "Market Value Start (EUR)",
    "Market Value End (EUR)",
    "Cost Basis (EUR)",
    "Realised P&L (Local)",
    "Realised P&L (EUR)",
    "Realised P&L (%)",
    "Unrealised P&L (Local)",
    "Unrealised P&L (EUR)",
    "Unrealised P&L (%)",
    "Income (Local)",
    "Income (EUR)",
    "Dividends (EUR)",
    "Swap Financing (EUR)",
    "Income Yield (%)",
    "Total P&L (EUR)",
    "Total P&L (%)",
    "Last Trade Date",
    "First Trade Date",
)

INCOME_COLUMNS: tuple[str, ...] = (
    "ISIN",
    "Stock Name",
    "Analyst",
    "CCY",
    "Date",
    "Income Type",
    "Income (Local)",
    "Income (EUR)",
    "Units",
)


# ─── Schema validator ────────────────────────────────────────────────────────

def assert_schema(df: pd.DataFrame, required: Iterable[str], *, where: str) -> None:
    """Raise ``SchemaError`` if any required column is missing.

    Args:
        df: The DataFrame to validate.
        required: Iterable of column names that must be present.
        where: Human-readable identifier of the call site (used in the error
            message — e.g. ``"compute.pl.compute_positions input"``).

    Extra columns are allowed; this only enforces presence, not exclusion.
    """
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SchemaError(
            f"{where}: missing required columns {missing}. "
            f"Got: {list(df.columns)}"
        )


def normalise_isin(value: object) -> str:
    """Canonical ISIN key used across snapshots/trades.

    Preserves semantic derivative suffixes such as ``:FUT`` so
    ``KYG8879R1048`` and ``KYG8879R1048:FUT`` remain distinct positions.
    """
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    raw = str(value).strip().upper()
    if not raw:
        return ""

    # Normalise colon-separated forms while preserving the suffix itself.
    if ":" in raw:
        left, right = raw.split(":", 1)
        left = left.strip()
        right = right.strip()
        if left and right:
            return f"{left}:{right}"
    return raw
