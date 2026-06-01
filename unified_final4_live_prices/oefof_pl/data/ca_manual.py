"""Manual/curated CA overrides: load, apply, write, and bonus-price helpers."""

from __future__ import annotations

import csv
import warnings
from pathlib import Path

import pandas as pd

from .ca_history import append_history
from .ca_types import BonusPriceOverride, CaOverride
from .normalise import normalise_isin


def load_ca_overrides(path: Path) -> list[CaOverride]:
    """Parse a CSV in the override format. Missing file = empty list."""
    p = Path(path)
    if not p.exists():
        return []

    out: list[CaOverride] = []
    with p.open("r", encoding="utf-8", newline="") as fh:
        reader = csv.reader(fh)
        for row in reader:
            if not row:
                continue
            first = (row[0] or "").strip()
            if not first or first.startswith("#"):
                continue
            isin = normalise_isin(first)
            if not isin or isin.upper() == "ISIN":
                continue
            if len(row) < 2:
                continue
            try:
                units = float((row[1] or "").strip())
            except ValueError:
                continue
            if units == 0:
                continue

            # Backward compatible schema:
            #   ISIN,UNITS,REASON
            # New schema:
            #   ISIN,UNITS,PRICE,REASON
            price = 0.0
            reason = ""
            if len(row) >= 4:
                try:
                    price = float((row[2] or "").strip() or 0.0)
                except ValueError:
                    price = 0.0
                reason = (row[3] or "").strip()
            elif len(row) >= 3:
                reason = (row[2] or "").strip()

            out.append(CaOverride(isin=isin, units=units, price=price, reason=reason))
    return out


def load_overrides_merged(*paths: Path) -> list[CaOverride]:
    """Load and merge multiple override files.

    When the same ISIN appears in more than one file, units are **summed**
    (not replaced). This is required by the auto-retry flow in ``app.py``:
    the auto-generated ``ca_suggestions_*.csv`` lists the *residual* drift
    that survived after the curated ``inputs/ca_overrides.csv`` was already
    applied, so the residual must add to — not overwrite — the curated value.
    Reasons are joined with ' | '. Net-zero ISINs after summing are dropped.
    """
    merged: dict[tuple[str, float], CaOverride] = {}
    for p in paths:
        for ov in load_ca_overrides(p):
            price_key = round(float(ov.price), 8)
            key = (ov.isin, price_key)
            prev = merged.get(key)
            if prev is None:
                merged[key] = ov
            else:
                new_units = prev.units + ov.units
                reason = " | ".join(r for r in (prev.reason, ov.reason) if r)
                merged[key] = CaOverride(
                    isin=ov.isin, units=new_units, price=price_key, reason=reason,
                )
    rows = [ov for ov in merged.values() if ov.units != 0]
    return sorted(rows, key=lambda x: (x.isin, x.price))


def overrides_to_trades(
    overrides: list[CaOverride],
    *,
    end_vdate: pd.Timestamp,
    template_columns: list[str],
    base_id: int = -1_000_000,
) -> pd.DataFrame:
    """Materialise the overrides as a trade DataFrame ready for ``concat``.

    ``base_id`` is the starting (negative) TRADE_ID for the synthetic rows;
    callers that inject in multiple batches (e.g. manual then auto) must
    pass disjoint ranges to avoid VAL-02 duplicate-TRADE_ID failures.
    """
    if not overrides:
        return pd.DataFrame(columns=template_columns)

    rows = []
    for i, ov in enumerate(overrides):
        if ":" in ov.isin:
            warnings.warn(
                f"Skipping synthetic CA override ISIN {ov.isin!r}; colon-suffixed "
                "keys are not valid for valuation checks.",
                stacklevel=2,
            )
            continue
        price_local = float(ov.price)
        buy_net_local = float(ov.units) * price_local
        rows.append({
            "TRADE_ID": base_id - i,
            "PCODE_ORIG": "OEFOF",
            "ISIN": ov.isin,
            "SNAME": "(CA adjustment)",
            "CCY": "",
            "T": "P",
            "CDATE": end_vdate,
            "UNITS": float(ov.units),
            "GROSSPRICE_LOCAL": price_local,
            "NETPRICE_LOCAL": price_local,
            "BUY_NET_LOCAL": buy_net_local,
            "SELL_NET_LOCAL": 0.0,
            "INCOME_LOCAL": 0.0,
            "BCOMM_LOCAL": 0.0,
            "EXPENSES_LOCAL": 0.0,
            "BCODE": "CA_ADJ",
            "DELETED": "N",
            "_CA_REASON": ov.reason,
        })

    df = pd.DataFrame(rows)
    for col in template_columns:
        if col not in df.columns:
            df[col] = pd.NA
    return df


def load_bonus_prices(path: Path) -> list[BonusPriceOverride]:
    """Load optional bonus-price overrides from CSV.

    Schema: ISIN,DATE,PRICE[,REASON]
    Missing file => empty list.
    """
    p = Path(path)
    if not p.exists():
        return []

    out: list[BonusPriceOverride] = []
    with p.open("r", encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            isin = normalise_isin(row.get("ISIN", ""))
            if not isin:
                continue
            if ":" in isin:
                continue
            try:
                cdate = pd.Timestamp(row.get("DATE", "")).normalize()
            except Exception:
                continue
            if pd.isna(cdate):
                continue
            try:
                price = float(row.get("PRICE", 0) or 0)
            except (TypeError, ValueError):
                continue
            if price <= 0:
                continue
            out.append(BonusPriceOverride(
                isin=isin,
                cdate=cdate,
                price=price,
                reason=(row.get("REASON") or "").strip(),
            ))
    return out


def apply_bonus_price_overrides(
    trades_df: pd.DataFrame,
    overrides: list[BonusPriceOverride],
) -> tuple[pd.DataFrame, int]:
    """Patch zero-price BONUS buy trades with configured local prices.

    This updates only ``GROSSPRICE_LOCAL`` and does not alter units or cash-flow
    columns, so unit reconciliation and cash identities remain unchanged.
    """
    if trades_df.empty or not overrides:
        return trades_df, 0

    required = {"ISIN", "CDATE", "T", "SNAME", "GROSSPRICE_LOCAL"}
    if not required.issubset(set(trades_df.columns)):
        return trades_df, 0

    out = trades_df.copy()
    t_u = out["T"].astype(str).str.upper()
    sname_u = out["SNAME"].astype(str).str.upper()
    gross = pd.to_numeric(out["GROSSPRICE_LOCAL"], errors="coerce").fillna(0.0)
    cdate = pd.to_datetime(out["CDATE"], errors="coerce").dt.normalize()
    isin = out["ISIN"].map(normalise_isin)

    is_bonus = sname_u.str.contains(r"\bBON(?:US)?\b", regex=True, na=False)
    base_mask = (t_u == "P") & (gross == 0.0) & is_bonus

    patched = 0
    for ov in overrides:
        mask = base_mask & (isin == ov.isin) & (cdate == ov.cdate)
        hits = int(mask.sum())
        if hits <= 0:
            continue
        out.loc[mask, "GROSSPRICE_LOCAL"] = float(ov.price)
        patched += hits

    return out, patched


def write_ca_overrides(path: Path, overrides: list[CaOverride]) -> None:
    """Rewrite ``ca_overrides.csv`` with the given list of overrides.

    Keeps the curated header comment so the file stays human-readable.
    Existing content is replaced; callers should merge first via
    ``load_overrides_merged`` if they want to preserve prior rows.
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8", newline="") as fh:
        fh.write("# OEFOF corporate-action unit overrides.\n")
        fh.write("# Each row injects a synthetic T='P' trade at price 0 on the period-end date.\n")
        fh.write("# UNITS is signed: positive = bonus/new units received, negative = consolidation.\n")
        fh.write("# Auto-rows from the pipeline are marked [AUTO]; review periodically.\n")
        fh.write("#\n")
        fh.write("# ISIN, UNITS, PRICE, REASON\n")
        w = csv.writer(fh)
        w.writerow(["ISIN", "UNITS", "PRICE", "REASON"])
        for ov in sorted(overrides, key=lambda x: x.isin):
            units_str = f"{ov.units:.6f}".rstrip("0").rstrip(".")
            price_str = f"{ov.price:.6f}".rstrip("0").rstrip(".")
            w.writerow([ov.isin, units_str, price_str, ov.reason])


def persist_auto_ca(
    *,
    suggestions_path: Path,
    overrides_path: Path,
    history_path: Path,
    period_end: pd.Timestamp,
) -> list[CaOverride]:
    """Merge auto-generated suggestions into the permanent ``ca_overrides.csv``.

    Called after a successful auto-retry so the adjustments survive future
    runs without needing to re-discover them.  Returns the list of newly
    persisted overrides (empty if suggestions_path does not exist).
    """
    if not Path(suggestions_path).exists():
        return []

    new_overrides = load_ca_overrides(suggestions_path)
    if not new_overrides:
        return []

    # Tag them as auto so the audit trail is clear
    new_overrides = [
        CaOverride(isin=ov.isin, units=ov.units, price=ov.price,
                   reason=f"[AUTO {period_end.date()}] {ov.reason}")
        for ov in new_overrides
    ]

    # Merge with existing overrides (same-ISIN units are summed)
    existing = load_ca_overrides(overrides_path)
    merged = load_overrides_merged(overrides_path, suggestions_path)

    # Build merged list: existing rows updated with new values, new ISINs appended
    existing_isins = {ov.isin for ov in existing}
    new_isins = {ov.isin for ov in new_overrides} - existing_isins

    final: list[CaOverride] = []
    for ov in existing:
        # If this ISIN has a new adjustment, use the merged (summed) value
        merged_match = next((m for m in merged if m.isin == ov.isin), None)
        if merged_match and ov.isin in {n.isin for n in new_overrides}:
            final.append(CaOverride(
                isin=ov.isin, units=merged_match.units, price=merged_match.price,
                reason=f"{ov.reason} | [AUTO {period_end.date()}] +{next(n.units for n in new_overrides if n.isin == ov.isin):.4g}u",
            ))
        else:
            final.append(ov)
    # Append brand-new ISINs
    for ov in new_overrides:
        if ov.isin in new_isins:
            final.append(ov)

    write_ca_overrides(overrides_path, final)
    append_history(history_path, new_overrides, period_end=period_end)
    return new_overrides
