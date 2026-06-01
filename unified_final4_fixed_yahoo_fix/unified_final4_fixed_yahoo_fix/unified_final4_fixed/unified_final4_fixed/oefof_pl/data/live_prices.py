from __future__ import annotations

import csv
import time
from pathlib import Path

from .normalise import normalise_isin

# ---------------------------------------------------------------------------
# Yahoo ticker CSV loader
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Direct Yahoo Finance price fetch (no yfinance dependency)
# ---------------------------------------------------------------------------

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}

_RETRIES = 3
_RETRY_DELAY = 2.0

# Two query endpoints — try both in order if the first fails
_YAHOO_URLS = [
    "https://query1.finance.yahoo.com/v8/finance/chart/{ticker}",
    "https://query2.finance.yahoo.com/v8/finance/chart/{ticker}",
]


def _fetch_ticker_price(ticker: str) -> float | None:
    """Fetch last price for a ticker via Yahoo Finance JSON API."""
    try:
        import requests  # type: ignore
    except ImportError:
        print("requests not installed — cannot fetch live prices")
        return None

    last_exc: Exception | None = None
    for attempt in range(_RETRIES):
        for url_template in _YAHOO_URLS:
            url = url_template.format(ticker=ticker)
            try:
                r = requests.get(url, headers=_HEADERS, timeout=10)
                r.raise_for_status()
                data = r.json()
                price = (
                    data.get("chart", {})
                    .get("result", [{}])[0]
                    .get("meta", {})
                    .get("regularMarketPrice")
                )
                if price is not None:
                    return float(price)
            except Exception as exc:
                last_exc = exc
                continue  # try next URL

        if attempt < _RETRIES - 1:
            time.sleep(_RETRY_DELAY)

    if last_exc is not None:
        print(f"Failed to get ticker '{ticker}' reason: {last_exc}")
    return None


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def fetch_live_prices(
    arb_pairs: tuple[dict[str, object], ...],
    *,
    yahoo_tickers: dict[str, str] | None = None,
) -> dict[str, float]:
    """Return {isin: last_price} for all tickers resolvable via Yahoo Finance.

    Silently returns {} if requests is not installed or all fetches fail —
    the pipeline continues without live prices.
    """
    try:
        import requests  # noqa: F401 — just check it's available
    except ImportError:
        return {}

    # Build ISIN -> ticker map from arb pairs + explicit yahoo_tickers CSV
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
        price = _fetch_ticker_price(ticker)
        if price is not None:
            prices[isin] = price

    return prices