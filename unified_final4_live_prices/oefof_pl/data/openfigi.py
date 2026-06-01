"""OpenFIGI ISIN → Yahoo ticker resolver.

Batch-maps ISINs to Yahoo Finance ticker symbols using the free OpenFIGI API.
Results are cached to ``inputs/openfigi_cache.json`` so repeat runs don't
re-fetch unchanged mappings.

Limits (no API key):  25 ISINs per request, ~10 req/minute.
Limits (free API key): 100 ISINs per request, higher rate limit.
Get a free key at https://www.openfigi.com/api — optional but recommended.

The resolver picks the best listing for each ISIN using this priority:
  1. Manual override from yahoo_tickers.csv (always wins)
  2. Primary exchange matching the position's CCY (e.g. HKD → HKEX)
  3. US listing (most liquid, Yahoo tickers work cleanly)
  4. Any other listing

Yahoo ticker format: if micCode is XHKG → append ".HK", XNSE → ".NS", etc.
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Mapping

# MIC code → Yahoo Finance suffix
# Covers the main markets in an EM/frontier fund universe
_MIC_TO_YAHOO_SUFFIX: dict[str, str] = {
    "XHKG": ".HK",    # Hong Kong
    "XSHG": ".SS",    # Shanghai
    "XSHE": ".SZ",    # Shenzhen
    "XBOM": ".BO",    # Bombay/BSE
    "XNSE": ".NS",    # NSE India
    "XKRX": ".KS",    # Korea
    "XTAI": ".TW",    # Taiwan
    "XBKK": ".BK",    # Thailand
    "XIDX": ".JK",    # Indonesia
    "XKLS": ".KL",    # Malaysia
    "XPHS": ".PS",    # Philippines
    "XVNM": ".VN",    # Vietnam (HoSE)
    "XSTC": ".VN",    # Vietnam (HNX)
    "XCAI": ".CA",    # Egypt
    "XJSE": ".JO",    # South Africa
    "XNAI": ".NR",    # Nigeria
    "XKAR": ".KA",    # Pakistan
    "XDHA": ".BD",    # Bangladesh
    "XCOL": ".CM",    # Colombia
    "XBUE": ".BA",    # Argentina
    "XSAU": ".SR",    # Saudi Arabia (Tadawul)
    "XDFM": ".AE",    # Dubai
    "XADS": ".AD",    # Abu Dhabi
    "XCAS": ".CA",    # Casablanca
    "XWAR": ".WA",    # Warsaw
    "XBUD": ".BD",    # Budapest
    "XPRA": ".PR",    # Prague
    "XIST": ".IS",    # Istanbul
    # US exchanges — no suffix needed
    "XNYS": "",
    "XNAS": "",
    "ARCX": "",       # NYSE Arca
    "XASE": "",       # NYSE American
    # London
    "XLON": ".L",
    # Frankfurt
    "XFRA": ".F",
    # Paris
    "XPAR": ".PA",
    # Amsterdam
    "XAMS": ".AS",
}

# CCY → preferred MIC codes (in priority order)
_CCY_TO_MIC: dict[str, list[str]] = {
    "HKD": ["XHKG"],
    "CNY": ["XSHG", "XSHE"],
    "INR": ["XNSE", "XBOM"],
    "KRW": ["XKRX"],
    "TWD": ["XTAI"],
    "THB": ["XBKK"],
    "IDR": ["XIDX"],
    "MYR": ["XKLS"],
    "PHP": ["XPHS"],
    "VND": ["XVNM", "XSTC"],
    "EGP": ["XCAI"],
    "ZAR": ["XJSE"],
    "NGN": ["XNAI"],
    "PKR": ["XKAR"],
    "BDT": ["XDHA"],
    "SAR": ["XSAU"],
    "AED": ["XDFM", "XADS"],
    "MAD": ["XCAS"],
    "PLN": ["XWAR"],
    "HUF": ["XBUD"],
    "CZK": ["XPRA"],
    "TRY": ["XIST"],
    "GBP": ["XLON"],
    "USD": ["XNYS", "XNAS", "ARCX"],
    "EUR": ["XPAR", "XFRA", "XAMS"],
}

_OPENFIGI_URL = "https://api.openfigi.com/v3/mapping"
_BATCH_SIZE = 10        # max without API key (25 requires key)
_RETRY_DELAY = 6.0      # seconds between batches (rate limit without key)
_REQUEST_TIMEOUT = 15   # seconds


def _ticker_from_figi_result(result: dict) -> str | None:
    """Extract a Yahoo-format ticker from one OpenFIGI result entry."""
    ticker = str(result.get("ticker") or "").strip()
    mic = str(result.get("exchCode") or result.get("primaryExchCode") or "").strip().upper()
    if not ticker:
        return None
    suffix = _MIC_TO_YAHOO_SUFFIX.get(mic, None)
    if suffix is None:
        # Unknown exchange — skip rather than guess
        return None
    return f"{ticker}{suffix}"


def _pick_best_ticker(
    figi_data: list[dict],
    ccy: str | None,
) -> str | None:
    """From a list of FIGI matches for one ISIN, pick the best Yahoo ticker."""
    if not figi_data:
        return None

    ccy_u = (ccy or "").strip().upper()
    preferred_mics = set(_CCY_TO_MIC.get(ccy_u, []))
    us_mics = {"XNYS", "XNAS", "ARCX", "XASE"}

    candidates: list[tuple[int, str]] = []  # (priority, ticker)
    for entry in figi_data:
        mic = str(entry.get("exchCode") or entry.get("primaryExchCode") or "").strip().upper()
        ticker = _ticker_from_figi_result(entry)
        if not ticker:
            continue
        # Security type filter — only equities
        sec_type = str(entry.get("securityType") or entry.get("securityType2") or "").lower()
        if any(x in sec_type for x in ("option", "warrant", "future", "etf", "fund")):
            continue
        if mic in preferred_mics:
            priority = 0
        elif mic in us_mics:
            priority = 1
        else:
            priority = 2
        candidates.append((priority, ticker))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]


def _load_cache(cache_path: Path) -> dict[str, str]:
    """Load {isin: yahoo_ticker} cache from disk."""
    if not cache_path.exists():
        return {}
    try:
        with cache_path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_cache(cache_path: Path, cache: dict[str, str]) -> None:
    """Persist cache to disk atomically."""
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = cache_path.with_suffix(".tmp")
        with tmp.open("w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2, sort_keys=True)
        tmp.replace(cache_path)
    except Exception as exc:
        print(f"[openfigi] cache save failed: {exc}")


def lookup_tickers(
    isins: list[str],
    *,
    isin_to_ccy: Mapping[str, str] | None = None,
    cache_path: Path | None = None,
    api_key: str | None = None,
    manual_overrides: dict[str, str] | None = None,
) -> dict[str, str]:
    """Map ISINs to Yahoo Finance tickers via OpenFIGI.

    Args:
        isins:            List of ISINs to resolve.
        isin_to_ccy:      Optional {isin: ccy} to guide exchange selection.
        cache_path:       Path to JSON cache file. None = no caching.
        api_key:          Optional OpenFIGI API key for higher rate limits.
        manual_overrides: {isin: ticker} that always win (from yahoo_tickers.csv).

    Returns:
        {isin: yahoo_ticker} for every ISIN that could be resolved.
        ISINs that couldn't be mapped are simply absent from the dict.
    """
    try:
        import requests  # type: ignore
    except ImportError:
        print("[openfigi] requests not installed — skipping ticker lookup")
        return {}

    overrides = {str(k).strip().upper(): str(v).strip()
                 for k, v in (manual_overrides or {}).items() if v}
    isin_to_ccy_clean = {str(k).strip().upper(): str(v).strip().upper()
                         for k, v in (isin_to_ccy or {}).items()}
    isins_clean = [str(i).strip().upper() for i in isins if str(i).strip()]

    # Load cache
    cache: dict[str, str] = _load_cache(cache_path) if cache_path else {}

    # Which ISINs still need looking up?
    to_fetch = [
        i for i in isins_clean
        if i not in overrides and i not in cache
    ]

    if to_fetch:
        headers = {"Content-Type": "application/json"}
        if api_key:
            headers["X-OPENFIGI-APIKEY"] = api_key

        fetched: dict[str, str] = {}
        for batch_start in range(0, len(to_fetch), _BATCH_SIZE):
            batch = to_fetch[batch_start: batch_start + _BATCH_SIZE]
            payload = [{"idType": "ID_ISIN", "idValue": isin} for isin in batch]
            try:
                resp = requests.post(
                    _OPENFIGI_URL,
                    headers=headers,
                    json=payload,
                    timeout=_REQUEST_TIMEOUT,
                )
                if resp.status_code == 429:
                    print(f"[openfigi] rate limited — waiting {_RETRY_DELAY*2}s")
                    time.sleep(_RETRY_DELAY * 2)
                    resp = requests.post(
                        _OPENFIGI_URL,
                        headers=headers,
                        json=payload,
                        timeout=_REQUEST_TIMEOUT,
                    )
                resp.raise_for_status()
                results = resp.json()  # list of {"data": [...]} or {"error": "..."}
                for isin, result in zip(batch, results):
                    if "error" in result:
                        continue
                    data = result.get("data") or []
                    ccy = isin_to_ccy_clean.get(isin)
                    ticker = _pick_best_ticker(data, ccy)
                    if ticker:
                        fetched[isin] = ticker
            except Exception as exc:
                print(f"[openfigi] batch {batch_start//  _BATCH_SIZE + 1} failed: {exc}")

            # Rate-limit pause between batches
            if batch_start + _BATCH_SIZE < len(to_fetch):
                time.sleep(_RETRY_DELAY)

        if fetched:
            cache.update(fetched)
            if cache_path:
                _save_cache(cache_path, cache)
            print(f"[openfigi] resolved {len(fetched)}/{len(to_fetch)} new tickers")
        else:
            if to_fetch:
                print(f"[openfigi] could not resolve any of {len(to_fetch)} ISINs")

    # Assemble final result: manual overrides > cache
    out: dict[str, str] = {}
    for isin in isins_clean:
        if isin in overrides:
            out[isin] = overrides[isin]
        elif isin in cache:
            out[isin] = cache[isin]

    return out