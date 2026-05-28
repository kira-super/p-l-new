"""Pure FX conversion functions.

Convention (from HP_VAL):
    XRATE   = units of *local* currency per 1 USD   (e.g. JPY/USD ≈ 150)
    EUR_USD = units of EUR per 1 USD                (e.g. ≈ 0.86)

Therefore:
    USD     → EUR : amount * eur_usd
    Local   → USD : amount / xrate
    Local   → EUR : (amount / xrate) * eur_usd
    EUR     → EUR : amount

Design rules — every one is a deliberate fix to a v1 bug:

* **No globals.** Every conversion takes its rates as explicit arguments. The
  legacy code mutated module-level `EUR_USD_END`/`EUR_USD_START` and let
  them leak into helpers (BUG-2). Here, if you don't pass the rate, your
  code won't compile.

* **No fallback constants.** `resolve_eur_usd` raises ``FXMissingError`` if
  the snapshot does not contain an EUR row. The legacy
  ``fx.get('EUR', 0.868018)`` was the root of BUG-1 (silent wrong rate after
  HP_VAL renames CCY column).

* **No date-dependent state.** Each call carries its own (xrate, eur_usd)
  pair. The "start vs end" distinction lives entirely in the caller, which
  selects the right snapshot.
"""

from __future__ import annotations

import math
from typing import Mapping

import pandas as pd

from ..exceptions import FXInvalidError, FXMissingError, SchemaError


# ─── Single-amount converter ─────────────────────────────────────────────────

def local_to_eur_at(
    amount: float,
    ccy: str,
    xrate: float,
    eur_usd: float,
) -> float:
    """Convert ``amount`` in ``ccy`` to EUR using the supplied rates.

    Args:
        amount: Monetary amount in ``ccy``.
        ccy:    ISO-ish currency code as it appears in HP_VAL (``EUR``,
                ``USD``, ``GBP``, ``JPY``, ...). Case-insensitive.
        xrate:  Local currency per 1 USD for ``ccy`` (ignored for EUR/USD).
        eur_usd: EUR per 1 USD.

    Returns:
        EUR-equivalent of ``amount``.

    Raises:
        FXInvalidError: ``eur_usd`` is non-finite, ≤ 0; or ``xrate`` is
            non-finite or ≤ 0 for a non-EUR/non-USD currency.
    """
    ccy_u = (ccy or "").strip().upper()

    if not _finite_pos(eur_usd):
        raise FXInvalidError(f"eur_usd must be finite and >0, got {eur_usd!r}")

    if amount == 0 or amount is None or (isinstance(amount, float) and math.isnan(amount)):
        return 0.0

    if ccy_u == "EUR":
        return float(amount)

    if ccy_u == "USD":
        return float(amount) * float(eur_usd)

    if not _finite_pos(xrate):
        raise FXInvalidError(
            f"xrate for {ccy_u} must be finite and >0, got {xrate!r}"
        )

    return (float(amount) / float(xrate)) * float(eur_usd)


def eur_to_local_at(
    amount_eur: float,
    ccy: str,
    xrate: float,
    eur_usd: float,
) -> float:
    """Inverse of :func:`local_to_eur_at`.

    Used when the snapshot already gives a value in EUR (PTVALUE_EUR) but
    a downstream identity needs the same value in the position's local CCY
    (e.g. unrealised = end_mv − end_cost, both in local).
    """
    ccy_u = (ccy or "").strip().upper()

    if not _finite_pos(eur_usd):
        raise FXInvalidError(f"eur_usd must be finite and >0, got {eur_usd!r}")

    if amount_eur == 0 or amount_eur is None or (
        isinstance(amount_eur, float) and math.isnan(amount_eur)
    ):
        return 0.0

    if ccy_u == "EUR":
        return float(amount_eur)

    if ccy_u == "USD":
        return float(amount_eur) / float(eur_usd)

    if not _finite_pos(xrate):
        raise FXInvalidError(
            f"xrate for {ccy_u} must be finite and >0, got {xrate!r}"
        )

    return (float(amount_eur) / float(eur_usd)) * float(xrate)


# ─── Snapshot extractors ─────────────────────────────────────────────────────

def build_fx_table(port_df: pd.DataFrame) -> dict[str, float]:
    """Return ``{CCY: XRATE}`` for every currency present in ``port_df``.

    EUR is included with whatever XRATE the snapshot carries (typically the
    EUR/USD rate itself, since EUR is the local currency for an EUR row).
    USD is forced to 1.0 if absent — it has no XRATE in HP_VAL because USD
    *is* the reference for XRATE.

    Raises:
        SchemaError: if ``CCY`` or ``XRATE`` columns are missing.
        FXInvalidError: if any non-EUR/USD currency has a non-positive XRATE.
    """
    if "CCY" not in port_df.columns or "XRATE" not in port_df.columns:
        raise SchemaError(
            "build_fx_table: port_df missing CCY/XRATE columns; "
            f"got {list(port_df.columns)}"
        )

    out: dict[str, float] = {"USD": 1.0}
    seen: dict[str, float] = {}

    for ccy, xrate in zip(port_df["CCY"], port_df["XRATE"]):
        if ccy is None:
            continue
        c = str(ccy).strip().upper()
        if not c:
            continue
        if isinstance(xrate, float) and math.isnan(xrate):
            continue
        try:
            x = float(xrate)
        except (TypeError, ValueError):
            continue

        if c in seen:
            # XRATE for the same CCY can vary slightly across rows
            # (different timestamps in the same snapshot). First value wins;
            # variance is checked below.
            continue
        seen[c] = x

    for c, x in seen.items():
        if c in ("USD",):
            continue
        if c == "EUR":
            # EUR's XRATE is by definition EUR/USD; allow it through.
            if not _finite_pos(x):
                raise FXInvalidError(f"EUR XRATE invalid in snapshot: {x!r}")
        else:
            if not _finite_pos(x):
                raise FXInvalidError(f"XRATE for {c} invalid in snapshot: {x!r}")
        out[c] = x

    return out


def resolve_eur_usd(port_df: pd.DataFrame) -> float:
    """Return the EUR/USD rate (EUR per 1 USD) read from the snapshot.

    Raises:
        FXMissingError: if no EUR row exists in ``port_df``. **No fallback
            constant is ever used** — this is the BUG-1 fix. A bad pipeline
            run must fail loudly, not silently emit a wrong report.
        FXInvalidError: if EUR is present but its XRATE is non-finite or ≤ 0.
        SchemaError: if required columns are missing.
    """
    if "CCY" not in port_df.columns or "XRATE" not in port_df.columns:
        raise SchemaError(
            "resolve_eur_usd: port_df missing CCY/XRATE columns; "
            f"got {list(port_df.columns)}"
        )

    eur_rows = port_df.loc[
        port_df["CCY"].astype("string").str.upper().str.strip() == "EUR",
        "XRATE",
    ]
    eur_rows = eur_rows.dropna()
    if eur_rows.empty:
        raise FXMissingError(
            "EUR/USD rate not found in portfolio snapshot. "
            "No fallback is applied — fix the source data and re-run."
        )

    rate = float(eur_rows.iloc[0])
    if not _finite_pos(rate):
        raise FXInvalidError(f"EUR/USD rate is invalid: {rate!r}")
    return rate


def build_fx_from_trades(trades_df: pd.DataFrame) -> dict[str, float]:
    """Extract ``{CCY: XRATE}`` from the optional ``XRATE`` column in trades.

    Used as the **lowest-priority** fallback tier for currencies that are
    absent from both the start- and end-of-period snapshots.  The canonical
    case is a position that was opened *and* fully closed within the YTD
    window: it never appears in either snapshot, so ``build_fx_table`` never
    sees its CCY, but the trade rows still carry the trade-date XRATE.

    Only rows where both ``CCY`` and ``XRATE`` are present, non-null, and
    where XRATE is finite and positive are included.  Rows with NaN/zero XRATE
    are silently skipped — the caller's ``_fx_rate_for`` will still raise a
    clear ``FXInvalidError`` if the CCY ends up being needed at compute time.

    Args:
        trades_df: The trades DataFrame.  If it has no ``XRATE`` column (the
            column is optional in ``TRADES_REQUIRED``) an empty dict is
            returned, so callers that pass the standard trades frame never
            need a special case.

    Returns:
        ``{CCY: XRATE}`` mapping, possibly empty.
    """
    if trades_df.empty or "CCY" not in trades_df.columns or "XRATE" not in trades_df.columns:
        return {}

    out: dict[str, float] = {}
    for ccy, xrate in zip(trades_df["CCY"], trades_df["XRATE"]):
        if ccy is None:
            continue
        c = str(ccy).strip().upper()
        if not c or c in ("EUR", "USD"):
            continue  # EUR/USD handled specially; USD is always 1.0
        if c in out:
            continue  # first valid rate wins (chronological order not guaranteed here)
        try:
            x = float(xrate)
        except (TypeError, ValueError):
            continue
        if _finite_pos(x):
            out[c] = x

    return out


def build_fx_supplement(
    extra_df: pd.DataFrame,
    base_fx: dict[str, float],
) -> dict[str, float]:
    """Fill gaps in ``base_fx`` using a secondary snapshot (e.g. ``tH_VAL``).

    Returns a **new** dict equal to ``base_fx`` plus any additional
    ``{CCY: XRATE}`` entries found in ``extra_df`` for CCYs not already
    present in ``base_fx``.  Entries already in ``base_fx`` are never
    overwritten — the primary snapshot always wins.

    This is used to recover FX rates for currencies whose XRATE is
    ``NULL``/``NaN`` in the primary portfolio snapshot (``vw_RPT_VAL``).
    The canonical case is a position (e.g. a NOK-denominated stock closed
    during the YTD period) where HiPort writes ``NULL`` for XRATE in the
    Dec-31 start snapshot but ``tH_VAL`` carries a valid rate for the same
    position on that date.

    Args:
        extra_df:  Secondary snapshot DataFrame with ``CCY`` and ``XRATE``
                   columns. Missing or empty → returns ``base_fx`` unchanged.
        base_fx:   Already-built FX table (from :func:`build_fx_table`).

    Returns:
        Copy of ``base_fx`` supplemented with any new valid rates from
        ``extra_df``.  The original ``base_fx`` dict is not mutated.
    """
    if extra_df is None or extra_df.empty:
        return dict(base_fx)
    if "CCY" not in extra_df.columns or "XRATE" not in extra_df.columns:
        return dict(base_fx)

    result = dict(base_fx)

    for ccy, xrate in zip(extra_df["CCY"], extra_df["XRATE"]):
        if ccy is None:
            continue
        c = str(ccy).strip().upper()
        if not c or c in result:
            continue  # already covered by the primary snapshot
        try:
            x = float(xrate)
        except (TypeError, ValueError):
            continue
        if _finite_pos(x):
            result[c] = x

    return result


# ─── helpers ─────────────────────────────────────────────────────────────────

def _finite_pos(x: object) -> bool:
    """True iff ``x`` is a finite, strictly-positive number."""
    try:
        f = float(x)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return False
    return math.isfinite(f) and f > 0.0
