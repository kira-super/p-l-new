from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class PairLegMetrics:
    isin: str
    stock_name: str
    instrument: str
    units: float
    avg_buy_price_eur: float | None
    avg_sell_price_eur: float | None
    last_price_eur: float | None
    cost_basis_eur: float
    market_value_end_eur: float
    market_value_start_eur: float
    realised_pl_eur: float
    unrealised_pl_eur: float
    swap_financing_eur: float
    total_pl_eur: float


@dataclass(frozen=True)
class PairMetrics:
    label: str
    analyst: str
    long_leg: PairLegMetrics
    short_leg: PairLegMetrics
    combined_total_pl_eur: float
    ratio_end: float | None
    ratio_start: float | None
    ratio_delta: float | None


def _sum_numeric(df: pd.DataFrame, column: str) -> float:
    if column not in df.columns or df.empty:
        return 0.0
    return float(pd.to_numeric(df[column], errors="coerce").fillna(0.0).sum())


def _first_text(df: pd.DataFrame, column: str) -> str:
    if column not in df.columns or df.empty:
        return ""
    for value in df[column]:
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def _first_price(df: pd.DataFrame, column: str) -> float | None:
    if column not in df.columns or df.empty:
        return None
    series = pd.to_numeric(df[column], errors="coerce")
    valid = series[series.notna()]
    if valid.empty:
        return None
    return float(valid.iloc[0])


def _ratio(long_mv: float, short_mv: float) -> float | None:
    denom = abs(float(short_mv))
    if denom < 1e-9:
        return None
    return float(long_mv) / denom


def _ratio_short_denom(leg: PairLegMetrics) -> float:
    if leg.instrument.upper() == "FTSWAP" and abs(leg.cost_basis_eur) > 1e-9:
        return leg.cost_basis_eur
    return leg.market_value_end_eur


def _build_leg(pl_df: pd.DataFrame, isin: str, live_prices: dict[str, float]) -> PairLegMetrics:
    isin_upper = str(isin).strip().upper()
    mask = pl_df.get("ISIN", pd.Series(dtype=object)).astype(str).str.strip().str.upper() == isin_upper
    sub = pl_df.loc[mask].copy()
    last_price = live_prices.get(isin_upper)
    if last_price is None:
        last_price = _first_price(sub, "Last Price (EUR)")
    return PairLegMetrics(
        isin=isin_upper,
        stock_name=_first_text(sub, "Stock Name"),
        instrument=_first_text(sub, "Instrument"),
        units=_sum_numeric(sub, "Ending Units"),
        avg_buy_price_eur=_first_price(sub, "Avg Buy Price (EUR)"),
        avg_sell_price_eur=_first_price(sub, "Avg Sell Price (EUR)"),
        last_price_eur=last_price,
        cost_basis_eur=_sum_numeric(sub, "Cost Basis (EUR)"),
        market_value_end_eur=_sum_numeric(sub, "Market Value End (EUR)"),
        market_value_start_eur=_sum_numeric(sub, "Market Value Start (EUR)"),
        realised_pl_eur=_sum_numeric(sub, "Realised P&L (EUR)"),
        unrealised_pl_eur=_sum_numeric(sub, "Unrealised P&L (EUR)"),
        swap_financing_eur=_sum_numeric(sub, "Swap Financing (EUR)"),
        total_pl_eur=_sum_numeric(sub, "Total P&L (EUR)"),
    )


def compute_pair_metrics(
    pl_df: pd.DataFrame,
    arb_pairs: tuple[dict[str, object], ...],
    live_prices: dict[str, float] | None = None,
) -> list[PairMetrics]:
    live_prices = {str(k).strip().upper(): float(v) for k, v in (live_prices or {}).items()}
    pairs: list[PairMetrics] = []
    for pair in arb_pairs:
        long_isin = str(pair.get("long_isin", "")).strip().upper()
        short_isin = str(pair.get("short_isin", "")).strip().upper()
        if not long_isin or not short_isin:
            continue
        long_leg = _build_leg(pl_df, long_isin, live_prices)
        short_leg = _build_leg(pl_df, short_isin, live_prices)
        ratio_end = _ratio(long_leg.market_value_end_eur, _ratio_short_denom(short_leg))
        ratio_start = _ratio(long_leg.market_value_start_eur, short_leg.market_value_start_eur)
        ratio_delta = None if ratio_end is None or ratio_start is None else ratio_end - ratio_start
        pairs.append(PairMetrics(
            label=str(pair.get("label", f"{long_isin}/{short_isin}")),
            analyst=str(pair.get("analyst", "")).strip().upper(),
            long_leg=long_leg,
            short_leg=short_leg,
            combined_total_pl_eur=long_leg.total_pl_eur + short_leg.total_pl_eur,
            ratio_end=ratio_end,
            ratio_start=ratio_start,
            ratio_delta=ratio_delta,
        ))
    return pairs