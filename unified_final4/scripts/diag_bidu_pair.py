"""BIDU pair P&L: long BAIDU INC HK + short GSCBIHKT CFD."""
import pandas as pd

pd.set_option("display.float_format", lambda x: f"{x:,.2f}")
pd.set_option("display.width", 240)
pd.set_option("display.max_columns", None)

df = pd.read_parquet("oefof_pl/out/pl_results.parquet")
pair = df[df["Stock Name"].isin(["BAIDU INC", "GSCBIHKT CFD GOLDMAN"])][[
    "Stock Name", "L/S", "Instrument", "Starting Units", "Ending Units",
    "Cost Basis (EUR)", "Market Value Start (EUR)", "Market Value End (EUR)",
    "Realised P&L (EUR)", "Unrealised P&L (EUR)", "Income (EUR)",
    "Total P&L (EUR)",
]]
print(pair.to_string(index=False))
print()
print(f"PAIR TOTAL P&L (EUR): {pair['Total P&L (EUR)'].sum():,.2f}")
