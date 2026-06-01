from __future__ import annotations

import pandas as pd
import pytest
from openpyxl import Workbook

from oefof_pl.compute.pairs import PairLegMetrics, PairMetrics, compute_pair_metrics
from oefof_pl.output.sheets.pairs_section import write_pairs_section


def test_compute_pair_metrics_calculates_ratios_and_leg_fields():
    pl_df = pd.DataFrame([
        {
            "ISIN": "KYG070341048",
            "Stock Name": "BAIDU INC",
            "Instrument": "ORD",
            "Ending Units": 100.0,
            "Avg Buy Price (EUR)": 7.5,
            "Avg Sell Price (EUR)": 0.0,
            "Last Price (EUR)": 9.5,
            "Cost Basis (EUR)": 750.0,
            "Market Value End (EUR)": 900.0,
            "Market Value Start (EUR)": 800.0,
            "Realised P&L (EUR)": 10.0,
            "Unrealised P&L (EUR)": 20.0,
            "Swap Financing (EUR)": 0.0,
            "Total P&L (EUR)": 35.0,
        },
        {
            "ISIN": "FDGSCBIHKT",
            "Stock Name": "GSCBIHKT CFD GOLDMAN",
            "Instrument": "FTSWAP",
            "Ending Units": 100.0,
            "Avg Buy Price (EUR)": 0.0,
            "Avg Sell Price (EUR)": 8.4,
            "Last Price (EUR)": 8.0,
            "Cost Basis (EUR)": 1_800.0,
            "Market Value End (EUR)": -600.0,
            "Market Value Start (EUR)": -700.0,
            "Realised P&L (EUR)": -3.0,
            "Unrealised P&L (EUR)": 2.0,
            "Swap Financing (EUR)": 312.0,
            "Total P&L (EUR)": 0.0,
        },
    ])
    pairs = compute_pair_metrics(
        pl_df,
        ({
            "label": "Baidu INC / GSCBIHKT",
            "long_isin": "KYG070341048",
            "short_isin": "FDGSCBIHKT",
            "analyst": "HK",
        },),
        live_prices={"KYG070341048": 11.0},
    )
    assert len(pairs) == 1
    pair = pairs[0]
    assert pair.long_leg.avg_buy_price_eur == pytest.approx(7.5)
    assert pair.short_leg.avg_sell_price_eur == pytest.approx(8.4)
    assert pair.long_leg.last_price_eur == pytest.approx(11.0)
    assert pair.short_leg.swap_financing_eur == pytest.approx(312.0)
    assert pair.combined_total_pl_eur == pytest.approx(35.0)
    assert pair.ratio_end == pytest.approx(900.0 / 1_800.0)
    assert pair.ratio_start == pytest.approx(800.0 / 700.0)
    assert pair.ratio_delta == pytest.approx((900.0 / 1_800.0) - (800.0 / 700.0))


def test_write_pairs_section_renders_pair_block_for_analyst():
    wb = Workbook()
    ws = wb.active
    ws.title = "HK"
    pair = PairMetrics(
        label="Baidu INC / GSCBIHKT",
        analyst="HK",
        long_leg=PairLegMetrics(
            isin="KYG070341048",
            stock_name="BAIDU INC",
            instrument="ORD",
            units=100.0,
            avg_buy_price_eur=7.5,
            avg_sell_price_eur=None,
            last_price_eur=11.0,
            cost_basis_eur=750.0,
            market_value_end_eur=900.0,
            market_value_start_eur=800.0,
            realised_pl_eur=10.0,
            unrealised_pl_eur=20.0,
            swap_financing_eur=0.0,
            total_pl_eur=35.0,
        ),
        short_leg=PairLegMetrics(
            isin="FDGSCBIHKT",
            stock_name="GSCBIHKT CFD GOLDMAN",
            instrument="FTSWAP",
            units=100.0,
            avg_buy_price_eur=None,
            avg_sell_price_eur=8.4,
            last_price_eur=None,
            cost_basis_eur=1_800.0,
            market_value_end_eur=-600.0,
            market_value_start_eur=-700.0,
            realised_pl_eur=-3.0,
            unrealised_pl_eur=2.0,
            swap_financing_eur=312.0,
            total_pl_eur=0.0,
        ),
        combined_total_pl_eur=35.0,
        ratio_end=0.5,
        ratio_start=800.0 / 700.0,
        ratio_delta=0.5 - (800.0 / 700.0),
    )

    next_row = write_pairs_section(ws, [pair], "HK", start_row=5, col_start=2)

    assert next_row > 5
    assert ws.cell(5, 2).value == "  Pairs & Arb Trades"
    assert ws.cell(6, 2).value == "  Baidu INC / GSCBIHKT"
    assert ws.cell(7, 7).value == "Current Price (Local)"
    assert ws.cell(7, 12).value == "MV / Notional (EUR)"
    assert ws.cell(8, 3).value == "BAIDU INC"
    assert ws.cell(9, 3).value == "GSCBIHKT CFD GOLDMAN"
    assert ws.cell(8, 11).value == 35.0
    assert ws.cell(9, 10).value == 312.0
    assert ws.cell(9, 12).value == 1800.0
    assert ws.cell(10, 2).value == "Combined"
    assert ws.cell(11, 2).value == "Current Ratio (Long MV / Short Notional)"
    assert ws.cell(11, 3).value == "Opening Ratio (L/S)"
    assert ws.cell(11, 4).value == "Ratio Change"
    assert ws.cell(12, 2).value == 0.5
    assert ws.cell(12, 3).value == pytest.approx(800.0 / 700.0)


def test_write_pairs_section_collapses_ratio_monitor_when_opening_unavailable():
    wb = Workbook()
    ws = wb.active
    ws.title = "HK"
    pair = PairMetrics(
        label="Baidu INC / GSCBIHKT",
        analyst="HK",
        long_leg=PairLegMetrics(
            isin="KYG070341048",
            stock_name="BAIDU INC",
            instrument="ORD",
            units=100.0,
            avg_buy_price_eur=7.5,
            avg_sell_price_eur=None,
            last_price_eur=11.0,
            cost_basis_eur=750.0,
            market_value_end_eur=900.0,
            market_value_start_eur=0.0,
            realised_pl_eur=10.0,
            unrealised_pl_eur=20.0,
            swap_financing_eur=0.0,
            total_pl_eur=35.0,
        ),
        short_leg=PairLegMetrics(
            isin="FDGSCBIHKT",
            stock_name="GSCBIHKT CFD GOLDMAN",
            instrument="FTSWAP",
            units=100.0,
            avg_buy_price_eur=None,
            avg_sell_price_eur=8.4,
            last_price_eur=None,
            cost_basis_eur=1_800.0,
            market_value_end_eur=9.0,
            market_value_start_eur=0.0,
            realised_pl_eur=-3.0,
            unrealised_pl_eur=2.0,
            swap_financing_eur=312.0,
            total_pl_eur=0.0,
        ),
        combined_total_pl_eur=35.0,
        ratio_end=0.5,
        ratio_start=None,
        ratio_delta=None,
    )

    write_pairs_section(ws, [pair], "HK", start_row=5, col_start=2)

    assert ws.cell(11, 2).value == "Current Ratio (Long MV / Short Notional)"
    assert ws.cell(11, 3).value is None
    assert ws.cell(12, 2).value == 0.5
    assert ws.cell(12, 3).value is None