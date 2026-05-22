import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
from oefof_pl.config import load_default_config
from oefof_pl.data.sql_loader import load_snapshot_sql, load_trades_sql
from oefof_pl.compute.aggregate import aggregate_portfolio
from oefof_pl.data.ca_overrides import derive_ca_auto

cfg = load_default_config()
start, _ = load_snapshot_sql("20251231", conn_str=cfg.sql_conn_str, oefof_pcodes=cfg.oefof_pcodes)
end, _ = load_snapshot_sql("20260504", conn_str=cfg.sql_conn_str, oefof_pcodes=cfg.oefof_pcodes)
tr = load_trades_sql(conn_str=cfg.sql_conn_str, oefof_pcodes=cfg.oefof_pcodes)
tr["CDATE"] = pd.to_datetime(tr["CDATE"], errors="coerce")
tr = tr[(tr["CDATE"] > pd.Timestamp("2025-12-31")) & (tr["CDATE"] <= pd.Timestamp("2026-05-04"))]

s_agg = aggregate_portfolio(start)
e_agg = aggregate_portfolio(end)
s_units = {i: l.units for i, l in s_agg.items()}
e_units = {i: l.units for i, l in e_agg.items()}

targets = ["AU0000058737", "SA000A0MWH44", "VN000000PNJ6", "VN000000VCK5"]
for isin in targets:
    s = s_units.get(isin)
    e = e_units.get(isin)
    sub = tr[tr["ISIN"] == isin]
    p = sub[sub["T"] == "P"]["UNITS"].sum()
    sv = sub[sub["T"] == "S"]["UNITS"].sum()
    print(f"{isin}: start={s!r}, end={e!r}, trades={len(sub)}, P={p}, S={sv}")
    print(sub[["TRADE_ID","CDATE","T","UNITS","BCODE","PCODE_ORIG","SNAME"]].to_string(index=False))
    print()

auto = derive_ca_auto(start_units=s_units, end_units=e_units, trades_df=tr)
hits = [a for a in auto if a.isin in set(targets)]
print()
print("AUTO HITS:", hits)
print("Total auto generated:", len(auto))
print()
# Sample a few autos
for a in auto[:8]:
    print(f"  auto: {a.isin} units={a.units} reason={a.reason}")
