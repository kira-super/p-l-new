"""Auto-derived CA overrides and suggestions (built from VAL-01 failures)."""

from __future__ import annotations

import csv
from pathlib import Path

import pandas as pd

from .ca_types import CaOverride, CaSuggestion, HistoryEntry
from .normalise import normalise_isin


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
