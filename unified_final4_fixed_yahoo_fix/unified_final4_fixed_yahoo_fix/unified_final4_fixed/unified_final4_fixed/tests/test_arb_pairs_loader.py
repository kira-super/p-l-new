from __future__ import annotations

from pathlib import Path

import pytest

from oefof_pl.data.arb_pairs_loader import load_arb_pairs


def test_load_arb_pairs_valid_csv(tmp_path: Path):
    path = tmp_path / "arb_pairs.csv"
    path.write_text(
        "label,long_isin,short_isin,analyst,long_yahoo_ticker,short_yahoo_ticker\n"
        "Pair A,AA1,BB1,HK,1111.HK,2222.HK\n",
        encoding="utf-8",
    )
    out = load_arb_pairs(path)
    assert len(out) == 1
    assert out[0]["label"] == "Pair A"
    assert out[0]["long_yahoo_ticker"] == "1111.HK"
    assert out[0]["short_yahoo_ticker"] == "2222.HK"


def test_load_arb_pairs_blank_tickers_become_none(tmp_path: Path):
    path = tmp_path / "arb_pairs.csv"
    path.write_text(
        "label,long_isin,short_isin,analyst,long_yahoo_ticker,short_yahoo_ticker\n"
        "Pair A,AA1,BB1,HK,,\n",
        encoding="utf-8",
    )
    out = load_arb_pairs(path)
    assert out[0]["long_yahoo_ticker"] is None
    assert out[0]["short_yahoo_ticker"] is None


def test_load_arb_pairs_missing_optional_columns_use_defaults(tmp_path: Path):
    path = tmp_path / "arb_pairs.csv"
    path.write_text(
        "label,long_isin,short_isin,analyst\n"
        "Pair A,AA1,BB1,HK\n",
        encoding="utf-8",
    )
    out = load_arb_pairs(path)
    assert out[0]["long_yahoo_ticker"] is None
    assert out[0]["short_yahoo_ticker"] is None


def test_load_arb_pairs_missing_required_columns_raises_clear_error(tmp_path: Path):
    path = tmp_path / "arb_pairs.csv"
    path.write_text(
        "label,long_isin,analyst\n"
        "Pair A,AA1,HK\n",
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="missing required columns"):
        load_arb_pairs(path)