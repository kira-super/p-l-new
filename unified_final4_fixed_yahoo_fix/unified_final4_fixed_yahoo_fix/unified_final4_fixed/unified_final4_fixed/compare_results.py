import pandas as pd
import numpy as np

def get_metrics(df):
    total_pl = df["Total P&L (EUR)"].sum()
    cost_basis = df["Cost Basis (EUR)"].sum()
    realised = df["Realised P&L (EUR)"].sum()
    unrealised = df["Unrealised P&L (EUR)"].sum()
    income = df["Income (EUR)"].sum()
    # If Total P&L = Realised + Unrealised + Income + FX
    fx = total_pl - (realised + unrealised + income)
    total_pct = (total_pl / cost_basis) if cost_basis != 0 else 0
    return {
        "Total": total_pl,
        "TotalPct": total_pct,
        "Realised": realised,
        "Unrealised": unrealised,
        "Income": income,
        "FX": fx,
        "CostBasis": cost_basis
    }

df20 = pd.read_parquet("out/pl_results_20260520.parquet")
df21 = pd.read_parquet("out/pl_results_20260521.parquet")

m20 = get_metrics(df20)
m21 = get_metrics(df21)

print("--- TOTALS ---")
print(f"Date         {'Total':>15} {'TotalPct':>10} {'Realised':>15} {'Unrealised':>15} {'Income':>15} {'FX':>12}")
print(f"20260520     {m20['Total']:15,.2f} {m20['TotalPct']:10.2%} {m20['Realised']:15,.2f} {m20['Unrealised']:15,.2f} {m20['Income']:15,.2f} {m20['FX']:12,.2f}")
print(f"20260521     {m21['Total']:15,.2f} {m21['TotalPct']:10.2%} {m21['Realised']:15,.2f} {m21['Unrealised']:15,.2f} {m21['Income']:15,.2f} {m21['FX']:12,.2f}")
print(f"Delta        {m21['Total']-m20['Total']:15,.2f} {m21['TotalPct']-m20['TotalPct']:10.2%} {m21['Realised']-m20['Realised']:15,.2f} {m21['Unrealised']-m20['Unrealised']:15,.2f} {m21['Income']-m20['Income']:15,.2f} {m21['FX']-m20['FX']:12,.2f}")

# Group by Analyst
a20 = df20.groupby("Analyst")["Total P&L (EUR)"].sum()
a21 = df21.groupby("Analyst")["Total P&L (EUR)"].sum()
analyst_delta = (a21 - a20).fillna(a21).fillna(-a20).sort_values(ascending=False)

print("\n--- ANALYST DELTA ---")
for analyst, delta in analyst_delta.items():
    print(f"{analyst:<25}: {delta:15,.2f}")

# Group by ISIN
# Join them to get full info
cols = ["ISIN", "Stock Name", "Analyst", "Total P&L (EUR)", "Realised P&L (EUR)", "Unrealised P&L (EUR)", "Income (EUR)"]
merged = pd.merge(
    df20[cols], 
    df21[cols], 
    on=["ISIN", "Stock Name", "Analyst"], 
    how="outer", 
    suffixes=("_20", "_21")
).fillna(0)

merged["delta_total"] = merged["Total P&L (EUR)_21"] - merged["Total P&L (EUR)_20"]
merged["delta_realised"] = merged["Realised P&L (EUR)_21"] - merged["Realised P&L (EUR)_20"]
merged["delta_unrealised"] = merged["Unrealised P&L (EUR)_21"] - merged["Unrealised P&L (EUR)_20"]
merged["delta_income"] = merged["Income (EUR)_21"] - merged["Income (EUR)_20"]
# For ISIN FX delta, we can compute it if we want, but instruction asks for these specifically.
# Wait, "delta_fx" was requested in the table columns. 
# We calculate FX per ISIN as Total - (Realised + Unrealised + Income)
merged["fx_20"] = merged["Total P&L (EUR)_20"] - (merged["Realised P&L (EUR)_20"] + merged["Unrealised P&L (EUR)_20"] + merged["Income (EUR)_20"])
merged["fx_21"] = merged["Total P&L (EUR)_21"] - (merged["Realised P&L (EUR)_21"] + merged["Unrealised P&L (EUR)_21"] + merged["Income (EUR)_21"])
merged["delta_fx"] = merged["fx_21"] - merged["fx_20"]

top15 = merged.sort_values("delta_total", ascending=False).head(15)

print("\n--- TOP 15 ISIN DELTAS ---")
header = f"{'ISIN':<15} {'Stock Name':<25} {'Analyst':<15} {'Total':>12} {'Realised':>12} {'Unrealised':>12} {'Income':>10} {'FX':>10}"
print(header)
for _, r in top15.iterrows():
    print(f"{r['ISIN']:<15} {str(r['Stock Name'])[:24]:<25} {str(r['Analyst'])[:14]:<15} {r['delta_total']:12,.2f} {r['delta_realised']:12,.2f} {r['delta_unrealised']:12,.2f} {r['delta_income']:10,.2f} {r['delta_fx']:10,.2f}")

