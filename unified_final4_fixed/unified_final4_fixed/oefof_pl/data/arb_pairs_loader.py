from __future__ import annotations

import csv
from pathlib import Path

from .normalise import normalise_isin


_REQUIRED_COLUMNS = ("label", "long_isin", "short_isin", "analyst")


def _blank_to_none(raw: str | None) -> str | None:
    text = "" if raw is None else str(raw).strip()
    return text or None


def load_arb_pairs(path: Path) -> tuple[dict[str, object], ...]:
    path = Path(path)
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = tuple(reader.fieldnames or ())
        missing = [name for name in _REQUIRED_COLUMNS if name not in fieldnames]
        if missing:
            raise ValueError(f"arb pair registry missing required columns: {missing}")

        pairs: list[dict[str, object]] = []
        for row in reader:
            if row is None:
                continue
            label = str(row.get("label", "")).strip()
            long_isin = normalise_isin(row.get("long_isin"))
            short_isin = normalise_isin(row.get("short_isin"))
            analyst = str(row.get("analyst", "")).strip().upper()
            if not label and not long_isin and not short_isin and not analyst:
                continue
            pairs.append({
                "label": label,
                "long_isin": long_isin,
                "short_isin": short_isin,
                "analyst": analyst,
                "long_yahoo_ticker": _blank_to_none(row.get("long_yahoo_ticker")),
                "short_yahoo_ticker": _blank_to_none(row.get("short_yahoo_ticker")),
            })
    return tuple(pairs)