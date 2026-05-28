from __future__ import annotations

from pathlib import Path

import pandas as pd

from oefof_pl.data.ca_overrides import (
    BonusPriceOverride,
    CaOverride,
    apply_bonus_price_overrides,
    append_history,
    deduplicate_history,
    load_bonus_prices,
    load_ca_overrides,
    load_history,
    load_overrides_merged,
    overrides_to_trades,
)


def test_load_ca_overrides_supports_legacy_three_column_schema(tmp_path: Path):
    p = tmp_path / "ca.csv"
    p.write_text("ISIN,UNITS,REASON\nX,10,test reason\n", encoding="utf-8")

    rows = load_ca_overrides(p)
    assert len(rows) == 1
    assert rows[0].isin == "X"
    assert rows[0].units == 10
    assert rows[0].price == 0
    assert rows[0].reason == "test reason"


def test_load_ca_overrides_parses_price_column(tmp_path: Path):
    p = tmp_path / "ca.csv"
    p.write_text("ISIN,UNITS,PRICE,REASON\nX,10,17.92,test reason\n", encoding="utf-8")

    rows = load_ca_overrides(p)
    assert len(rows) == 1
    assert rows[0].price == 17.92


def test_overrides_to_trades_uses_price_and_skips_colon_isin():
    overrides = [
        CaOverride(isin="X", units=5.0, reason="ok", price=12.5),
        CaOverride(isin="X:FUT", units=-5.0, reason="skip", price=0.0),
    ]
    df = overrides_to_trades(
        overrides,
        end_vdate=pd.Timestamp("2026-05-21"),
        template_columns=["TRADE_ID", "ISIN", "T", "UNITS", "GROSSPRICE_LOCAL"],
    )
    assert len(df) == 1
    assert df.iloc[0]["ISIN"] == "X"
    assert float(df.iloc[0]["GROSSPRICE_LOCAL"]) == 12.5
    assert float(df.iloc[0]["NETPRICE_LOCAL"]) == 12.5
    assert float(df.iloc[0]["BUY_NET_LOCAL"]) == 62.5


def test_append_history_is_idempotent_by_period_and_isin(tmp_path: Path):
    p = tmp_path / "history.csv"
    rows = [CaOverride(isin="X", units=10.0, reason="first", price=1.23)]

    append_history(p, rows, period_end=pd.Timestamp("2026-05-21"))
    append_history(p, rows, period_end=pd.Timestamp("2026-05-21"))

    hist = load_history(p)
    assert len(hist) == 1
    assert hist[0].period_end == "2026-05-21"
    assert hist[0].isin == "X"
    assert hist[0].price == 1.23


def test_deduplicate_history_keeps_first_row(tmp_path: Path):
    p = tmp_path / "history.csv"
    p.write_text(
        "period_end,isin,units,price,reason\n"
        "2026-05-21,X,10,1.0,first\n"
        "2026-05-21,X,10,2.0,second\n"
        "2026-05-21,Y,5,0.0,other\n",
        encoding="utf-8",
    )

    removed = deduplicate_history(p)
    assert removed == 1

    hist = load_history(p)
    assert len(hist) == 2
    assert hist[0].isin == "X"
    assert hist[0].price == 1.0


def test_load_overrides_merged_keeps_distinct_prices_for_same_isin(tmp_path: Path):
    p = tmp_path / "ca.csv"
    p.write_text(
        "ISIN,UNITS,PRICE,REASON\n"
        "X,10,1.0,first\n"
        "X,20,2.0,second\n",
        encoding="utf-8",
    )
    rows = load_overrides_merged(p)
    assert len(rows) == 2
    assert {(r.isin, r.units, r.price) for r in rows} == {
        ("X", 10.0, 1.0),
        ("X", 20.0, 2.0),
    }


def test_load_bonus_prices_parses_valid_rows(tmp_path: Path):
    p = tmp_path / "bonus.csv"
    p.write_text(
        "ISIN,DATE,PRICE,REASON\n"
        "VN000000VCK5,2026-03-11,24.5,March bonus\n"
        "VN000000PNJ6,2026-04-24,95.0,April bonus\n",
        encoding="utf-8",
    )
    rows = load_bonus_prices(p)
    assert len(rows) == 2
    assert rows[0].isin == "VN000000VCK5"
    assert rows[0].cdate == pd.Timestamp("2026-03-11")
    assert rows[0].price == 24.5


def test_apply_bonus_price_overrides_patches_only_matching_bonus_rows():
    df = pd.DataFrame([
        {
            "ISIN": "VN000000VCK5",
            "CDATE": pd.Timestamp("2026-03-11"),
            "T": "P",
            "SNAME": "VPS SECURITIES BONUS",
            "GROSSPRICE_LOCAL": 0.0,
        },
        {
            "ISIN": "VN000000VCK5",
            "CDATE": pd.Timestamp("2026-03-11"),
            "T": "S",
            "SNAME": "VPS SECURITIES BONUS",
            "GROSSPRICE_LOCAL": 0.0,
        },
        {
            "ISIN": "VN000000VCK5",
            "CDATE": pd.Timestamp("2026-03-12"),
            "T": "P",
            "SNAME": "VPS SECURITIES BONUS",
            "GROSSPRICE_LOCAL": 0.0,
        },
    ])
    patched_df, patched = apply_bonus_price_overrides(
        df,
        [
            BonusPriceOverride(
                isin="VN000000VCK5",
                cdate=pd.Timestamp("2026-03-11"),
                price=24.5,
                reason="March bonus",
            )
        ],
    )
    assert patched == 1
    assert float(patched_df.iloc[0]["GROSSPRICE_LOCAL"]) == 24.5
    assert float(patched_df.iloc[1]["GROSSPRICE_LOCAL"]) == 0.0
    assert float(patched_df.iloc[2]["GROSSPRICE_LOCAL"]) == 0.0
