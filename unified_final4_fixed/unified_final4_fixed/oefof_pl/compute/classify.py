"""Instrument classification: ORD vs SWAP vs FTSWAP.

The classification drives the entire P&L methodology:

* **ORD**    Ordinary equity. Total = price MV change + realised + income.
             Cost basis from WAC.
* **SWAP**   Single-name CFD/total-return swap. Total = income only (daily
             MTM is captured in the income column). Cost basis = notional
             exposure |NUCOST × Units|.
* **FTSWAP** Index-level CFD (Euro Stoxx Banks, GSCBIHKT). Same compute
             path as SWAP but reported on a separate sheet.

Inputs accepted by ``classify_instrument``:

    cat   — value of the ``CAT`` column from HP_VAL ('ORD' / 'FUT' / etc.)
    sname — security short name (used for the 'CFD' substring fallback)
    isin  — security ISIN (used for the FTSWAP membership check)

**The legacy ``ORD`` column is intentionally NOT a parameter.** That column
is always 'A' in HP_VAL and any classifier that uses it would always return
ORD (BUG-9). To make the bug structurally impossible, the function refuses
to accept any extra positional or keyword arguments beyond the three above.
A unit test asserts the signature is locked.
"""

from __future__ import annotations

from typing import Iterable

# ─── Output domain ──────────────────────────────────────────────────────────

ORD = "ORD"
SWAP = "SWAP"
FTSWAP = "FTSWAP"
INSTRUMENT_TYPES: frozenset[str] = frozenset({ORD, SWAP, FTSWAP})


# ─── CAT values that mark a derivative position in HP_VAL ───────────────────
# Anything in this set classifies as a swap (then specialised to FTSWAP via
# the ISIN membership check). Add to this set, not to ad-hoc string checks
# elsewhere.
_DERIV_CAT_VALUES: frozenset[str] = frozenset({"FUT", "CFD", "SWAP"})
_DERIV_ISIN_SUFFIXES: tuple[str, ...] = (":FUT", ":CFD", ":SWAP")


def classify_instrument(
    cat: object,
    sname: object,
    isin: object,
    *,
    ftswap_isins: Iterable[str],
) -> str:
    """Classify a position as ORD, SWAP, or FTSWAP.

    Args:
        cat:   Value of the ``CAT`` column from HP_VAL.
        sname: Security short name. Used only for the 'CFD' substring
               fallback (covers rows where CAT is missing but the name
               clearly indicates a CFD).
        isin:  Security ISIN. Membership in ``ftswap_isins`` promotes a
               SWAP to FTSWAP.
        ftswap_isins: The set of ISINs that should be reported as FTSWAP
               instead of SWAP. Passed in by the caller (``Config.ftswap_isins``)
               so this module owns no fund-specific data.

    Returns:
        One of ``ORD``, ``SWAP``, ``FTSWAP``.

    Note:
        ``ORD`` here is the *return value* meaning ordinary equity, not the
        always-'A' ``ORD`` column from HP_VAL. The column is forbidden as
        an input — see module docstring.
    """
    cat_u = _norm(cat)
    sname_u = _norm(sname)
    isin_u = _norm(isin)
    ftswap_set = {_norm(x) for x in ftswap_isins}

    is_deriv = (
        (cat_u in _DERIV_CAT_VALUES)
        or ("CFD" in sname_u)
        or any(isin_u.endswith(suffix) for suffix in _DERIV_ISIN_SUFFIXES)
    )

    if not is_deriv:
        return ORD

    # Promote derivatives that are index-level CFDs to FTSWAP.
    if isin_u in ftswap_set:
        return FTSWAP
    # Belt-and-braces: the legacy code also pattern-matched on 'INDEX' or
    # 'GSCBIHKT' substrings in SNAME. Keep the safety net but anchor it on
    # the explicit ftswap_isins set whenever possible.
    if "INDEX" in sname_u or "GSCBIHKT" in sname_u:
        return FTSWAP

    return SWAP


# ─── helpers ────────────────────────────────────────────────────────────────

def _norm(x: object) -> str:
    """Stripped, upper-cased string. Empty for None/NaN."""
    if x is None:
        return ""
    try:
        # NaN check without importing pandas (we get plain Python objects
        # from itertuples or list comprehensions in callers).
        if isinstance(x, float) and x != x:  # NaN
            return ""
    except TypeError:
        pass
    return str(x).strip().upper()
