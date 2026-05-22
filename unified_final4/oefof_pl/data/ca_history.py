"""CA history: append-only audit log for applied overrides."""

from __future__ import annotations

import csv
from pathlib import Path

import pandas as pd

from .ca_types import CaOverride, HistoryEntry

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
                from .normalise import normalise_isin
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
