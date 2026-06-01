from __future__ import annotations

import types
from pathlib import Path

from oefof_pl.data.live_prices import fetch_live_prices, load_yahoo_ticker_map


class _FakeTicker:
    def __init__(self, ticker: str):
        self.fast_info = {"last_price": float(len(ticker))}


def test_load_yahoo_ticker_map_reads_isin_mapping(tmp_path: Path):
    path = tmp_path / "yahoo_tickers.csv"
    path.write_text(
        "isin,yahoo_ticker\n"
        "KYG070341048,9888.HK\n"
        "SA16CI8KMOH3,4321.SR\n",
        encoding="utf-8",
    )

    out = load_yahoo_ticker_map(path)

    assert out["KYG070341048"] == "9888.HK"
    assert out["SA16CI8KMOH3"] == "4321.SR"


def test_load_yahoo_ticker_map_missing_file_returns_empty(tmp_path: Path):
    out = load_yahoo_ticker_map(tmp_path / "missing.csv")
    assert out == {}


def test_fetch_live_prices_uses_arb_pairs_and_explicit_map(monkeypatch):
    fake_yfinance = types.SimpleNamespace(Ticker=lambda ticker: _FakeTicker(ticker))
    monkeypatch.setitem(__import__("sys").modules, "yfinance", fake_yfinance)

    arb_pairs = (
        {
            "long_isin": "KYG070341048",
            "short_isin": "",
            "long_yahoo_ticker": "9888.HK",
            "short_yahoo_ticker": None,
        },
    )
    explicit = {"SA16CI8KMOH3": "4321.SR"}

    out = fetch_live_prices(arb_pairs, yahoo_tickers=explicit)

    assert out["KYG070341048"] == 7.0
    assert out["SA16CI8KMOH3"] == 7.0
