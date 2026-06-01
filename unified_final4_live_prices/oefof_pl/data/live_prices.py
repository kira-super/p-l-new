"""Live price fetcher for open positions.

Fetches current market prices from Yahoo Finance for a given set of ISINs.
Uses the OpenFIGI API to resolve ISIN -> ticker automatically (no manual CSV
maintenance required), with ``yahoo_tickers.csv`` as a manual override layer.

Price fetch uses Yahoo Finance's JSON chart API directly -- no yfinance
dependency, which was blocked by the Fiera corporate proxy.

Returns prices in local currency (the position's trading currency).
FX conversion to EUR happens in the pipeline via ``_live_prices_to_eur``.
"""

from __future__ import annotations

import csv
import time
from pathlib import Path

from .normalise import normalise_isin


# ---------------------------------------------------------------------------
# Manual ticker CSV loader (override layer)
# ---------------------------------------------------------------------------

def load_yahoo_ticker_map(path: Path) -> dict[str, str]:
    """Load optional ISIN -> Yahoo ticker overrides from CSV.

    Expected columns: isin, yahoo_ticker
    Missing file is silently treated as empty mapping.
    """
    csv_path = Path(path)
    if not csv_path.exists():
        return {}

    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = {str(n).strip().lower() for n in (reader.fieldnames or [])}
        if not {"isin", "yahoo_ticker"}.issubset(fieldnames):
            raise ValueError(
                "yahoo_tickers.csv missing required columns: ['isin', 'yahoo_ticker']"
            )
        out: dict[str, str] = {}
        for row in reader:
            if row is None:
                continue
            isin = normalise_isin(row.get("isin"))
            ticker = str(row.get("yahoo_ticker", "")).strip()
            if isin and ticker:
                out[isin] = ticker
    return out


# ---------------------------------------------------------------------------
# Yahoo Finance direct price fetch (no yfinance)
# ---------------------------------------------------------------------------

_YAHOO_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}
_RETRIES = 3
_RETRY_DELAY = 2.0
_YAHOO_URLS = [
    "https://query1.finance.yahoo.com/v8/finance/chart/{ticker}",
    "https://query2.finance.yahoo.com/v8/finance/chart/{ticker}",
]


def _fetch_ticker_price(ticker: str) -> float | None:
    """Fetch regularMarketPrice for a Yahoo ticker via direct JSON API."""
    try:
        import requests  # type: ignore
    except ImportError:
        return None

    last_exc: Exception | None = None
    for attempt in range(_RETRIES):
        for url_tmpl in _YAHOO_URLS:
            url = url_tmpl.format(ticker=ticker)
            try:
                r = requests.get(url, headers=_YAHOO_HEADERS, timeout=10, verify=False)
                r.raise_for_status()
                data = r.json()
                price = (
                    data.get("chart", {})
                    .get("result", [{}])[0]
                    .get("meta", {})
                    .get("regularMarketPrice")
                )
                if price is not None:
                    px = float(price)
                    # Yahoo returns London prices in pence (GBX), not GBP.
                    # HiPort stores them in GBP. Divide by 100.
                    if ticker.upper().endswith(".L"):
                        px = px / 100.0
                    return px
            except Exception as exc:
                last_exc = exc
                continue
        if attempt < _RETRIES - 1:
            time.sleep(_RETRY_DELAY)

    if last_exc:
        print(f"[live_prices] failed to fetch '{ticker}': {last_exc}")
    return None


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def fetch_live_prices(
    arb_pairs: tuple[dict[str, object], ...],
    *,
    yahoo_tickers: dict[str, str] | None = None,
    open_isins: list[str] | None = None,
    openfigi_tickers: dict[str, str] | None = None,
) -> dict[str, float]:
    """Return {isin: last_price_local_ccy} for all resolvable ISINs.

    Ticker resolution priority (highest wins):
      1. yahoo_tickers (manual CSV overrides)
      2. openfigi_tickers (auto-resolved by pipeline before this call)
      3. arb_pairs inline ticker fields

    Args:
        arb_pairs:         Arb pairs config (existing interface, kept for compat).
        yahoo_tickers:     Manual ISIN -> ticker overrides from yahoo_tickers.csv.
        open_isins:        Full list of open-position ISINs to fetch prices for.
        openfigi_tickers:  Auto-resolved {isin: ticker} from OpenFIGI lookup.

    Returns:
        {isin: price_in_local_ccy} -- empty dict if requests unavailable or
        all fetches fail. Never raises.
    """
    try:
        import requests  # noqa: F401
    except ImportError:
        return {}

    # Build unified ISIN -> ticker map (priority: manual > openfigi > arb pairs)
    isin_to_ticker: dict[str, str] = {}

    # Lowest priority: arb pairs inline tickers
    for pair in arb_pairs:
        for isin_key, ticker_key in (
            ("long_isin", "long_yahoo_ticker"),
            ("short_isin", "short_yahoo_ticker"),
        ):
            isin = str(pair.get(isin_key, "")).strip().upper()
            ticker = str(pair.get(ticker_key, "") or "").strip()
            if isin and ticker:
                isin_to_ticker[isin] = ticker

    # Mid priority: OpenFIGI auto-resolved tickers
    for isin, ticker in (openfigi_tickers or {}).items():
        isin_k = normalise_isin(isin)
        ticker_v = str(ticker or "").strip()
        if isin_k and ticker_v:
            isin_to_ticker[isin_k] = ticker_v

    # Highest priority: manual CSV overrides
    for isin, ticker in (yahoo_tickers or {}).items():
        isin_k = normalise_isin(isin)
        ticker_v = str(ticker or "").strip()
        if isin_k and ticker_v:
            isin_to_ticker[isin_k] = ticker_v

    # Restrict to open positions only if provided
    if open_isins is not None:
        open_set = {normalise_isin(i) for i in open_isins if i}
        isin_to_ticker = {k: v for k, v in isin_to_ticker.items() if k in open_set}

    prices: dict[str, float] = {}
    for isin, ticker in isin_to_ticker.items():
        price = _fetch_ticker_price(ticker)
        if price is not None:
            prices[isin] = price

    if isin_to_ticker:
        print(f"[live_prices] fetched {len(prices)}/{len(isin_to_ticker)} prices")

    return prices
