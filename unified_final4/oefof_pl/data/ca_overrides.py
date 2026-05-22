"""Corporate-action unit overrides.

This module owns three artefacts:

* ``inputs/ca_overrides.csv``      — curated, permanent overrides (you edit it).
* ``inputs/ca_overrides_history.csv`` — append-only audit log; one row per
                                        applied override per period. Used to
                                        flag *recurring* CAs in suggestions.
* ``out/ca_suggestions_<ts>.csv``   — auto-generated when VAL-01 fails. Ready
                                        to merge into ``ca_overrides.csv`` or
                                        re-fed via ``--apply-suggestions PATH``.

Synthetic trade rows look like a normal P trade so :func:`compute.wac.compute_wac`
treats them correctly (bonus dilutes WAC, consolidation scales WAC up).
``BCODE='CA_ADJ'`` is non-ZZZZ so the WAC stripper in compute.pl leaves
them alone, but the marker keeps them auditable in Trade Detail / VAL-02.
"""

from __future__ import annotations

import csv
import re
import warnings
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .normalise import normalise_isin


@dataclass(frozen=True)
class CaOverride:
    isin: str
    units: float       # signed
    reason: str
    price: float = 0.0


@dataclass(frozen=True)
class CaSuggestion:
    """One auto-detected unit drift, presented to the user for sign-off."""
    isin: str
    sname: str
    units: float                     # signed (= VAL-01 Diff)
    start_units: float
    expected_end_units: float
    actual_end_units: float
    recurrence_count: int            # times same-ISIN/same-direction appeared in history
    last_seen_period: str            # "" if never; YYYY-MM-DD otherwise
    suggested_reason: str            # heuristic, includes recurrence hint


@dataclass(frozen=True)
class HistoryEntry:
    isin: str
    units: float
    price: float
    reason: str
    period_end: str                  # YYYY-MM-DD


@dataclass(frozen=True)
class BonusPriceOverride:
    isin: str
    cdate: pd.Timestamp
    price: float
    reason: str


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


def derive_ca_auto(
    *,
    start_units: dict[str, float],
    end_units: dict[str, float],
    trades_df: pd.DataFrame,
    manual_isins: set[str] | None = None,
    tol: float = 1e-3,
) -> list[CaOverride]:
    """Derive CA adjustments from snapshot deltas.

    For each ISIN, compares ``end_units`` against
    ``start_units + Σ(P trades) − Σ(S trades)`` (real + ZZZZ trades both
    counted, mirroring VAL-01's view). Any non-zero residual is emitted
    as a synthetic ``CaOverride`` so VAL-01 reconciles by construction
    and ``ca_overrides.csv`` need only be edited for the rare cases the
    snapshot itself is suspect.

    Manual overrides take precedence: ISINs already covered by
    ``ca_overrides.csv`` are skipped (the manual value wins).
    """
    if manual_isins is None:
        manual_isins = set()

    bought: dict[str, float] = {}
    sold: dict[str, float] = {}
    if not trades_df.empty and {"T", "ISIN", "UNITS"}.issubset(trades_df.columns):
        # Mirror the WAC stripper in compute.pl: BCODE='ZZZZ' P/S rows are
        # corporate-action placeholders that get dropped before WAC runs,
        # so they MUST NOT count towards the bought/sold totals here either
        # — otherwise the residual nets to ~0 and no synthetic CA is emitted,
        # but VAL-01 still fails because WAC sees only the real (non-ZZZZ)
        # trades. Bug seen in LOTUS RESOURCES, PHUNHUAN JEWELRY,
        # VPS SECURITIES BONUS, UNITED INTL TRANSP (May-2026 close).
        df = trades_df
        if "BCODE" in df.columns:
            bcode_u = df["BCODE"].astype(str).str.upper()
            t_u = df["T"].astype(str).str.upper()
            ca_mask = (bcode_u == "ZZZZ") & t_u.isin(["P", "S"])
            df = df.loc[~ca_mask]
        norm = df["ISIN"].map(normalise_isin)
        t = df["T"].astype(str).str.upper()
        u = pd.to_numeric(df["UNITS"], errors="coerce").fillna(0.0)
        for isin, ti, ui in zip(norm, t, u):
            if not isin:
                continue
            if ti == "P":
                bought[isin] = bought.get(isin, 0.0) + float(ui)
            elif ti == "S":
                sold[isin] = sold.get(isin, 0.0) + float(ui)

    auto: list[CaOverride] = []
    all_isins = set(start_units) | set(end_units) | set(bought) | set(sold)
    for isin in sorted(all_isins):
        if ":" in isin:
            continue
        if isin in manual_isins:
            continue
        s = float(start_units.get(isin, 0.0))
        e = float(end_units.get(isin, 0.0))
        b = float(bought.get(isin, 0.0))
        sl = float(sold.get(isin, 0.0))
        diff = e - (s + b - sl)
        if abs(diff) > tol:
            auto.append(CaOverride(
                isin=isin, units=diff,
                price=0.0,
                reason="(auto-derived from snapshot delta)",
            ))
    return auto


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


# ─── History (audit log) ────────────────────────────────────────────────────

_HISTORY_HEADER = ("period_end", "isin", "units", "price", "reason")


def load_history(path: Path) -> list[HistoryEntry]:
    """Read ``ca_overrides_history.csv``. Missing file = empty list."""
    p = Path(path)
    if not p.exists():
        return []
    out: list[HistoryEntry] = []
    with p.open("r", encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            try:
                isin = normalise_isin(row.get("isin", ""))
                if not isin:
                    continue
                out.append(HistoryEntry(
                    isin=isin,
                    units=float(row.get("units", 0) or 0),
                    price=float(row.get("price", 0) or 0),
                    reason=(row.get("reason") or "").strip(),
                    period_end=(row.get("period_end") or "").strip(),
                ))
            except (TypeError, ValueError):
                continue
    return out


def append_history(
    path: Path, applied: list[CaOverride], *, period_end: pd.Timestamp,
) -> None:
    """Append-only writer. Creates the file (with header) on first call."""
    if not applied:
        return
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)

    existing = load_history(p)
    existing_keys = {(h.period_end, h.isin) for h in existing}
    period_str = period_end.strftime("%Y-%m-%d") if hasattr(period_end, "strftime") else str(period_end)
    to_write: list[CaOverride] = []
    for ov in applied:
        k = (period_str, ov.isin)
        if k in existing_keys:
            continue
        existing_keys.add(k)
        to_write.append(ov)
    if not to_write:
        return

    write_header = not p.exists()
    with p.open("a", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        if write_header:
            w.writerow(_HISTORY_HEADER)
        for ov in to_write:
            w.writerow([
                period_str,
                ov.isin,
                f"{ov.units:.6f}".rstrip("0").rstrip("."),
                f"{ov.price:.6f}".rstrip("0").rstrip("."),
                ov.reason,
            ])


def deduplicate_history(path: Path) -> int:
    """Remove duplicate history rows by (period_end, isin), keeping first."""
    p = Path(path)
    if not p.exists():
        return 0
    rows = load_history(p)
    if not rows:
        return 0

    seen: set[tuple[str, str]] = set()
    kept: list[HistoryEntry] = []
    for row in rows:
        key = (row.period_end, row.isin)
        if key in seen:
            continue
        seen.add(key)
        kept.append(row)

    removed = len(rows) - len(kept)
    if removed <= 0:
        return 0

    with p.open("w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(_HISTORY_HEADER)
        for row in kept:
            w.writerow([
                row.period_end,
                row.isin,
                f"{row.units:.6f}".rstrip("0").rstrip("."),
                f"{row.price:.6f}".rstrip("0").rstrip("."),
                row.reason,
            ])
    return removed


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


# ─── Suggestions (auto-detected from VAL-01 failures) ───────────────────────

def build_suggestions(
    val01_failures: list[dict],
    history: list[HistoryEntry],
) -> list[CaSuggestion]:
    """Convert VAL-01 failure rows into suggestions enriched with recurrence."""
    # Pre-index history by ISIN with same-direction matches.
    by_isin: dict[str, list[HistoryEntry]] = {}
    for h in history:
        by_isin.setdefault(h.isin, []).append(h)

    out: list[CaSuggestion] = []
    for f in val01_failures:
        isin = str(f.get("ISIN", "")).strip()
        if not isin:
            continue
        if ":" in isin:
            continue
        diff = float(f.get("Diff", 0.0))
        if diff == 0:
            continue
        prior = by_isin.get(isin, [])
        same_dir = [h for h in prior if (h.units > 0) == (diff > 0)]
        same_dir.sort(key=lambda h: h.period_end, reverse=True)
        n = len(same_dir)
        last = same_dir[0].period_end if same_dir else ""
        # Heuristic reason
        if n >= 2:
            hint = f"recurring CA — last seen {last} ({n} prior occurrences)"
        elif n == 1:
            hint = f"matches prior CA on {last} (1 occurrence)"
        else:
            hint = "TODO confirm with ops"
        out.append(CaSuggestion(
            isin=isin,
            sname=str(f.get("Stock Name", "")).strip(),
            units=diff,
            start_units=float(f.get("Start Units", 0.0)),
            expected_end_units=float(f.get("Expected End Units", 0.0)),
            actual_end_units=float(f.get("Actual End Units", 0.0)),
            recurrence_count=n,
            last_seen_period=last,
            suggested_reason=hint,
        ))
    return out


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


def write_suggestions_csv(path: Path, suggestions: list[CaSuggestion]) -> None:
    """Write suggestions in the override CSV format so it can be fed straight
    to ``--apply-suggestions PATH`` after review.
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8", newline="") as fh:
        fh.write("# Auto-generated CA suggestions from VAL-01 failures.\n")
        fh.write("# Review each row, edit/delete as needed, then re-run with:\n")
        fh.write("#   python -m oefof_pl --apply-suggestions <this file>\n")
        fh.write("# Or paste accepted rows into inputs/ca_overrides.csv.\n")
        fh.write("#\n")
        fh.write("# Recurrence column shows how many prior periods had a same-direction\n")
        fh.write("# CA on this ISIN; a recurrence >= 1 is a strong hint the row is real.\n")
        fh.write("#\n")
        fh.write("# ISIN, UNITS, PRICE, REASON\n")
        w = csv.writer(fh)
        w.writerow(["ISIN", "UNITS", "PRICE", "REASON"])
        for s in suggestions:
            tag = f"[recurrence={s.recurrence_count}]" if s.recurrence_count else "[NEW]"
            reason = f"{tag} {s.sname} — {s.suggested_reason}".strip()
            w.writerow([s.isin, f"{s.units:.6f}".rstrip("0").rstrip("."), "0", reason])
