"""Bonus-share CA row splitter.

After the main P&L computation, positions that received bonus shares need to
be presented as two rows in the output workbook:

* The *original holding* row — keeps its realized / unrealized figures.
* The *bonus shares* row — zero cost basis, pure unrealized gain.

This is a pure DataFrame transformation with no I/O dependency, so it lives
here rather than inside the larger :mod:`compute.pl` orchestrator.
"""
from __future__ import annotations

import pandas as pd


def split_ca_rows(pl_df: pd.DataFrame) -> pd.DataFrame:
    """Split positions that received bonus shares into two rows:
    one for the original holding's P&L (with its realized/unrealized),
    one for the bonus shares received at zero cost (pure unrealized gain).

    This prevents a confusing presentation where a realized loss from selling
    a declining original holding appears on the same row as free bonus shares.
    """
    ca_mask = pd.to_numeric(pl_df.get("Bonus Units", 0), errors="coerce").fillna(0) > 0
    if not ca_mask.any():
        return pl_df

    result_rows: list[dict] = []
    for _, row in pl_df.iterrows():
        bonus = float(pd.to_numeric(row.get("Bonus Units", 0), errors="coerce") or 0)
        if bonus <= 0:
            result_rows.append(row.to_dict())
            continue

        total_end = float(row["Ending Units"])
        orig_end = total_end - bonus
        bonus_frac = bonus / total_end if total_end > 0 else 0.0

        mv_end_total = float(row.get("Market Value End (EUR)") or 0)
        bonus_mv = mv_end_total * bonus_frac
        orig_mv = mv_end_total - bonus_mv

        total_unreal_eur = float(row["Unrealised P&L (EUR)"])
        # If the original position is fully exited, all unrealized belongs to
        # the bonus shares; otherwise split proportionally by market value.
        if orig_end <= 0:
            bonus_unreal_eur = total_unreal_eur
            orig_unreal_eur = 0.0
        else:
            bonus_unreal_eur = bonus_mv
            orig_unreal_eur = total_unreal_eur - bonus_unreal_eur

        total_unreal_local = float(row.get("Unrealised P&L (Local)") or 0)
        if orig_end <= 0:
            bonus_unreal_local = total_unreal_local
            orig_unreal_local = 0.0
        elif abs(total_unreal_eur) > 1e-9:
            bonus_unreal_local = total_unreal_local * (bonus_unreal_eur / total_unreal_eur)
            orig_unreal_local = total_unreal_local - bonus_unreal_local
        else:
            bonus_unreal_local = total_unreal_local * bonus_frac
            orig_unreal_local = total_unreal_local - bonus_unreal_local

        realized_eur = float(row["Realised P&L (EUR)"])
        income_eur = float(row.get("Income (EUR)") or 0)
        orig_cost = float(row.get("Cost Basis (EUR)") or 0)

        # ── original holding row ──────────────────────────────────────────────
        orig = row.to_dict()
        orig["Ending Units"] = orig_end
        orig["Units Bought"] = max(0.0, float(row["Units Bought"]) - bonus)
        orig["Avg Buy Price (EUR)"] = 0.0 if orig["Units Bought"] == 0 else float(row["Avg Buy Price (EUR)"])
        orig["Market Value End (EUR)"] = orig_mv
        orig["Unrealised P&L (EUR)"] = orig_unreal_eur
        orig["Unrealised P&L (Local)"] = orig_unreal_local
        orig["Unrealised P&L (%)"] = orig_unreal_eur / orig_cost if abs(orig_cost) > 1e-9 else float("nan")
        orig_total = realized_eur + orig_unreal_eur + income_eur
        orig["Total P&L (EUR)"] = orig_total
        orig["Total P&L (%)"] = orig_total / orig_cost if abs(orig_cost) > 1e-9 else float("nan")
        orig["Bonus Units"] = 0.0

        # ── bonus shares row ──────────────────────────────────────────────────
        bon = row.to_dict()
        bon["Stock Name"] = str(row["Stock Name"]) + " (bonus)"
        bon["Starting Units"] = 0.0
        bon["Ending Units"] = bonus
        bon["Units Bought"] = bonus
        bon["Units Sold"] = 0.0
        bon["Avg Buy Price (EUR)"] = 0.0
        bon["Avg Sell Price (EUR)"] = float("nan")
        bon["Start Price (Local)"] = 0.0
        bon["Market Value Start (EUR)"] = 0.0
        bon["Market Value End (EUR)"] = bonus_mv
        bon["Cost Basis (EUR)"] = 0.0
        bon["Realised P&L (EUR)"] = 0.0
        bon["Realised P&L (Local)"] = 0.0
        bon["Realised P&L (%)"] = float("nan")
        bon["Unrealised P&L (EUR)"] = bonus_unreal_eur
        bon["Unrealised P&L (Local)"] = bonus_unreal_local
        bon["Unrealised P&L (%)"] = float("nan")
        bon["Income (EUR)"] = 0.0
        bon["Income (Local)"] = 0.0
        bon["Dividends (EUR)"] = 0.0
        bon["Swap Financing (EUR)"] = 0.0
        bon["Income Yield (%)"] = float("nan")
        bon["Total P&L (EUR)"] = bonus_unreal_eur
        bon["Total P&L (%)"] = float("nan")
        bon["First Trade Date"] = pd.NaT
        bon["Last Trade Date"] = pd.NaT
        bon["Bonus Units"] = bonus

        result_rows.append(orig)
        result_rows.append(bon)

    return pd.DataFrame(result_rows).reindex(columns=pl_df.columns).reset_index(drop=True)
