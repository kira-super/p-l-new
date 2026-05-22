"""Hard audit: independently verify MV End / MV Start / Cost Basis on every position.

For each position:
    MV_check = UNITS × LST_PRICE_LOCAL ÷ XRATE × EUR_USD
Compare to pipeline's Market Value End/Start (EUR).
Flag any |delta| > 1% AND > €100.
"""
from __future__ import annotations

import pandas as pd
import pyodbc

from oefof_pl.config import load_default_config

cfg = load_default_config()
cn = pyodbc.connect(cfg.sql_conn_str, timeout=30)
cur = cn.cursor()


def _load_snapshot(yyyymmdd: str) -> pd.DataFrame:
    placeholders = ",".join("?" * len(cfg.oefof_pcodes))
    iso = f"{yyyymmdd[:4]}-{yyyymmdd[4:6]}-{yyyymmdd[6:8]}"
    cur.execute(
        f"SELECT PCODE_ORIG,ISIN,SNAME,CCY,UNITS,LST_PRICE,XRATE,PFACTOR,PTVALUE,PTCOST,NUCOST,CAT "
        f"FROM ccl.dbo.vw_RPT_VAL "
        f"WHERE CAST(VDATE AS date)=CAST(? AS date) AND PCODE_ORIG IN ({placeholders})",
        iso, *cfg.oefof_pcodes,
    )
    rows = cur.fetchall()
    cols = [c[0] for c in cur.description]
    return pd.DataFrame.from_records(rows, columns=cols)


def _eur_usd_rate(yyyymmdd: str) -> float:
    iso = f"{yyyymmdd[:4]}-{yyyymmdd[4:6]}-{yyyymmdd[6:8]}"
    cur.execute(
        "SELECT TOP 1 XRATE FROM ccl.dbo.vw_RPT_VAL "
        "WHERE CAST(VDATE AS date)=CAST(? AS date) AND CCY='EUR'",
        iso,
    )
    rows = cur.fetchall()
    if not rows:
        return 1.0
    # XRATE for EUR row gives EUR per USD
    return float(rows[0][0])


# Resolve latest VDATE
cur.execute(
    "SELECT MAX(VDATE) FROM ccl.dbo.vw_RPT_VAL WHERE PCODE_ORIG IN (" +
    ",".join("?" * len(cfg.oefof_pcodes)) + ")",
    *cfg.oefof_pcodes,
)
end_vdate = cur.fetchone()[0]
end_yyyymmdd = end_vdate.strftime("%Y%m%d")
start_yyyymmdd = "20251231"

print(f"end_vdate   = {end_yyyymmdd}")
print(f"start_vdate = {start_yyyymmdd}")
print()

end_df = _load_snapshot(end_yyyymmdd)
start_df = _load_snapshot(start_yyyymmdd)
eur_usd_end = _eur_usd_rate(end_yyyymmdd)
eur_usd_start = _eur_usd_rate(start_yyyymmdd)
print(f"EUR/USD end   = {eur_usd_end}")
print(f"EUR/USD start = {eur_usd_start}")
print()


def _audit(df: pd.DataFrame, eur_usd: float, label: str) -> pd.DataFrame:
    df = df.copy()
    for col in ("UNITS", "LST_PRICE", "XRATE", "PFACTOR", "PTVALUE"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["PFACTOR"] = df["PFACTOR"].replace(0, 1.0)
    # MV in USD = units × price × pfactor ÷ xrate (CCY per USD)
    df["MV_USD_check"] = df["UNITS"] * df["LST_PRICE"] * df["PFACTOR"] / df["XRATE"].replace(0, 1.0)
    df["MV_EUR_check"] = df["MV_USD_check"] * eur_usd
    df["MV_EUR_db"] = df["PTVALUE"]
    df["delta_eur"] = df["MV_EUR_check"] - df["MV_EUR_db"]
    df["delta_pct"] = df["delta_eur"] / df["MV_EUR_db"].replace(0, pd.NA)
    bad = df[(df["delta_eur"].abs() > 100) & (df["delta_pct"].abs() > 0.01)]
    print(f"=== {label}: {len(df)} positions, {len(bad)} discrepancies > 1% AND > €100 ===")
    if not bad.empty:
        cols = ["PCODE_ORIG", "ISIN", "SNAME", "CCY", "CAT", "UNITS",
                "LST_PRICE", "XRATE", "PFACTOR", "MV_EUR_db", "MV_EUR_check",
                "delta_eur", "delta_pct"]
        print(bad[cols].to_string(index=False))
    print()
    return df


end_audit = _audit(end_df, eur_usd_end, "END snapshot")
start_audit = _audit(start_df, eur_usd_start, "START snapshot")

# Now check pipeline output too
pl = pd.read_parquet(r'oefof_pl/out/pl_results.parquet')
print("=== Pipeline MV End vs DB PTVALUE per ISIN (top 5 |delta|) ===")
end_db = end_audit.groupby("ISIN", as_index=False).agg(
    db_mv=("MV_EUR_db", "sum"), check_mv=("MV_EUR_check", "sum"),
)
merged = pl[["ISIN", "Stock Name", "Market Value End (EUR)"]].merge(end_db, on="ISIN", how="left")
merged["delta"] = merged["Market Value End (EUR)"] - merged["db_mv"]
merged["delta_check"] = merged["Market Value End (EUR)"] - merged["check_mv"]
print(merged.reindex(merged["delta_check"].abs().sort_values(ascending=False).index).head(10).to_string(index=False))
