"""Portfolio-line aggregation: collapse one HP_VAL snapshot into per-ISIN
``AggLine`` records.

A single ISIN can appear on multiple HP_VAL rows (different sub-accounts,
different fee classes, both an ORD line and a FUT line for hybrid stocks
like Vietnam Dairy). This module sums them into a single canonical record
so ``compute.pl`` works one ISIN at a time.

Fields are local-currency only — EUR conversion happens later, in
``compute.pl``, using ``compute.fx.local_to_eur_at``.

Two structural bug fixes live in this module:

* **BUG-7 (b)** — ``cat`` is escalated to ``'FUT'`` if any source row for
  the ISIN was FUT. This prevents the Vietnam Dairy hybrid (an ORD line +
  a FUT line) from being mis-classified as pure ORD.
* **BUG-11 origin** — fields carry their unit-of-measure in the suffix.
  ``ptvalue_eur`` and ``ptcost_eur`` are EUR (already converted by HiPort
  in the snapshot); ``nucost_local``, ``lst_price_local`` are local CCY.
  Mixing them is a structural bug.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

import pandas as pd

from ..exceptions import SchemaError
from ..data.normalise import PORT_REQUIRED, assert_schema, normalise_isin


@dataclass(frozen=True)
class AggLine:
    """One ISIN's aggregated portfolio snapshot in local currency.

    ``cat`` is the CAT value after FUT-escalation. ``n_source_rows`` is kept
    for diagnostics: a value > 1 indicates a hybrid or multi-account
    holding, which is useful when validating reconciliations.

    ``ptvalue_eur`` / ``ptcost_eur`` come straight from HiPort already in
    EUR and must NOT be re-converted. ``nucost_local`` and
    ``lst_price_local`` are per-share in the position's local CCY. Mixing
    these units is what BUG-11 was about — the suffix is now mandatory.
    """

    isin: str
    sname: str
    ccy: str
    ls: str
    country: str
    cat: str            # post-escalation
    units: float
    ptvalue_eur: float        # already in fund base CCY (EUR) in the snapshot
    ptcost_eur: float         # already in fund base CCY (EUR) in the snapshot
    nucost_local: float       # weighted avg by |units|, per share, in local CCY
    lst_price_local: float
    xrate: float              # local CCY per 1 USD; same for all rows of an ISIN
    n_source_rows: int


# CAT values that promote a hybrid holding to derivative-status.
_DERIV_CATS: frozenset[str] = frozenset({"FUT", "CFD", "SWAP"})


def aggregate_portfolio(port_df: pd.DataFrame) -> dict[str, AggLine]:
    """Collapse a portfolio snapshot into ``{ISIN: AggLine}``.

    Args:
        port_df: Output of ``data.loader.load_snapshot`` — already schema-
            validated, with ``_LOCAL`` renames applied. Must contain the
            columns listed in :data:`oefof_pl.data.normalise.PORT_REQUIRED`.

    Returns:
        Dict keyed by canonical ISIN. ISINs that normalise to empty are
        dropped (they cannot be looked up later).
    """
    assert_schema(port_df, PORT_REQUIRED, where="aggregate_portfolio")

    if port_df.empty:
        return {}

    # Normalise once into a working frame.
    work = port_df.copy()
    work["__isin__"] = work["ISIN"].map(normalise_isin)
    work = work.loc[work["__isin__"] != ""]
    if work.empty:
        return {}

    out: dict[str, AggLine] = {}

    for isin, grp in work.groupby("__isin__", sort=False):
        # If the group mixes ORD and derivative (FUT/CFD/SWAP) rows under the
        # same ISIN, split them rather than collapsing. HiPort can store an
        # equity and its paired CFD under the same ISIN with different CAT
        # values (e.g. TIME INTERCONNECT KYG8879R1048). Collapsing them sums
        # their PTVALUEs (producing a huge phantom unrealised) and loses the
        # instrument distinction. Split: ORD rows keep the bare ISIN key,
        # derivative rows get an "ISIN:FUT" key so classify_instrument sees
        # the correct CAT for each sub-position.
        cats_u = grp["CAT"].astype(str).str.strip().str.upper()
        is_deriv_row = cats_u.isin(_DERIV_CATS)
        if is_deriv_row.any() and (~is_deriv_row).any():
            sub_groups: list[tuple[str, "pd.DataFrame"]] = [
                (isin,         grp.loc[~is_deriv_row]),
                (isin + ":FUT", grp.loc[is_deriv_row]),
            ]
        else:
            sub_groups = [(isin, grp)]

        for sub_isin, sub_grp in sub_groups:
            if sub_grp.empty:
                continue
            first = sub_grp.iloc[0]

            units = float(_safe_sum(sub_grp["UNITS"]))
            ptvalue = float(_safe_sum(sub_grp["PTVALUE_EUR"]))

            # PTCOST may not survive every loader path; treat as optional.
            ptcost = (
                float(_safe_sum(sub_grp["PTCOST_EUR"]))
                if "PTCOST_EUR" in sub_grp.columns
                else 0.0
            )

            # NUCOST = weighted average per |units|, matching legacy
            # `_aggregate_port`. Falls back to first-row value if all units
            # are zero (degenerate case).
            nucost = _weighted_avg(sub_grp["NUCOST_LOCAL"], sub_grp["UNITS"].abs())
            if math.isnan(nucost):
                nucost = _safe_float(first["NUCOST_LOCAL"])

            lst_price = _safe_float(first["LST_PRICE_LOCAL"])
            xrate = _first_valid_xrate(sub_grp["XRATE"])
            cat = _escalate_cat(sub_grp["CAT"])

            out[sub_isin] = AggLine(
                isin=sub_isin,
                sname=str(first["SNAME"]).strip() if pd.notna(first["SNAME"]) else "",
                ccy=str(first["CCY"]).strip().upper() if pd.notna(first["CCY"]) else "",
                ls=str(first["LS"]).strip().upper() if pd.notna(first["LS"]) else "L",
                country=str(first["EXCODE1"]).strip() if pd.notna(first["EXCODE1"]) else "",
                cat=cat,
                units=units,
                ptvalue_eur=ptvalue,
                ptcost_eur=ptcost,
                nucost_local=float(nucost),
                lst_price_local=lst_price,
                xrate=xrate,
                n_source_rows=int(len(sub_grp)),
            )

    return out


# ─── helpers ────────────────────────────────────────────────────────────────



def _escalate_cat(cats: pd.Series) -> str:
    """Return 'FUT' (or first deriv match) if any row is a derivative,
    else the first non-empty CAT value, else 'ORD'."""
    seen_deriv: str | None = None
    first_nonempty: str | None = None

    for v in cats:
        if v is None or (isinstance(v, float) and math.isnan(v)):
            continue
        s = str(v).strip().upper()
        if not s:
            continue
        if first_nonempty is None:
            first_nonempty = s
        if s in _DERIV_CATS and seen_deriv is None:
            seen_deriv = s

    if seen_deriv is not None:
        return seen_deriv
    return first_nonempty or "ORD"


def _safe_sum(s: pd.Series) -> float:
    return float(pd.to_numeric(s, errors="coerce").fillna(0.0).sum())


def _safe_float(x: object) -> float:
    if x is None:
        return 0.0
    try:
        f = float(x)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return 0.0
    return 0.0 if math.isnan(f) else f


def _first_valid_xrate(s: pd.Series) -> float:
    """Return the first finite, positive XRATE in *s*; 0.0 if none.

    Hiport snapshots sometimes carry NaN XRATE on basket / synthetic rows
    (e.g. short JPY index lines). The first row of an ISIN group can be
    one of those, so picking row[0] blindly leaves the AggLine carrying a
    NaN xrate that later trips ``local_to_eur_at`` with a confusing error.
    Picking the first valid value is the structural fix.
    """
    v = pd.to_numeric(s, errors="coerce")
    valid = v[v.notna() & (v > 0)]
    if valid.empty:
        return 0.0
    return float(valid.iloc[0])


def _weighted_avg(values: pd.Series, weights: pd.Series) -> float:
    v = pd.to_numeric(values, errors="coerce")
    w = pd.to_numeric(weights, errors="coerce").fillna(0.0)
    mask = v.notna() & (w != 0)
    if not mask.any():
        return float("nan")
    total_w = float(w[mask].sum())
    if total_w == 0:
        return float("nan")
    return float((v[mask] * w[mask]).sum() / total_w)
