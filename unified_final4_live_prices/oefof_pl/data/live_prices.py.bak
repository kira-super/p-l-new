from __future__ import annotations

import csv
from pathlib import Path

from .normalise import normalise_isin


def _extract_last_price(fast_info) -> float | None:
    if fast_info is None:
        return None
    if hasattr(fast_info, "last_price"):
        try:
            value = fast_info.last_price
            return float(value) if value is not None else None
        except Exception:
            return None
    if isinstance(fast_info, dict):
        for key in ("last_price", "lastPrice"):
            value = fast_info.get(key)
            if value is not None:
                try:
                    return float(value)
                except (TypeError, ValueError):
                    return None
    return None


def load_yahoo_ticker_map(path: Path) -> dict[str, str]:
    """Load optional ISIN->Yahoo ticker mapping from CSV.

    Expected columns: isin,yahoo_ticker
    Missing file is treated as empty mapping.
    """
    csv_path = Path(path)
    if not csv_path.exists():
        return {}

    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = {str(name).strip().lower() for name in (reader.fieldnames or [])}
        if not {"isin", "yahoo_ticker"}.issubset(fieldnames):
            raise ValueError("yahoo_tickers.csv missing required columns: ['isin', 'yahoo_ticker']")

        out: dict[str, str] = {}
        for row in reader:
            if row is None:
                continue
            isin = normalise_isin(row.get("isin"))
            ticker = str(row.get("yahoo_ticker", "")).strip()
            if not isin or not ticker:
                continue
            out[isin] = ticker
    return out


def fetch_live_prices(
    arb_pairs: tuple[dict[str, object], ...],
    *,
    yahoo_tickers: dict[str, str] | None = None,
) -> dict[str, float]:
    try:
        import yfinance  # type: ignore
    except Exception:
        return {}

    isin_to_ticker: dict[str, str] = {}
    for pair in arb_pairs:
        for isin_key, ticker_key in (
            ("long_isin", "long_yahoo_ticker"),
            ("short_isin", "short_yahoo_ticker"),
        ):
            isin = str(pair.get(isin_key, "")).strip().upper()
            ticker = str(pair.get(ticker_key, "") or "").strip()
            if isin and ticker:
                isin_to_ticker[isin] = ticker
    for isin, ticker in (yahoo_tickers or {}).items():
        isin_key = normalise_isin(isin)
        ticker_key = str(ticker or "").strip()
        if isin_key and ticker_key:
            isin_to_ticker[isin_key] = ticker_key

    prices: dict[str, float] = {}
    for isin, ticker in isin_to_ticker.items():
        try:
            ticker_obj = yfinance.Ticker(str(ticker))
            price = _extract_last_price(getattr(ticker_obj, "fast_info", None))
            if price is not None:
                prices[isin] = price
        except Exception:
            continue
    return prices