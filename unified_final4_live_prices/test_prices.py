"""
test_prices.py — Compare Yahoo closing price on HiPort VDATE vs HiPort LST_PRICE
                 for ALL ISINs in the analyst map (not just current portfolio).

Run from project root:
    python test_prices.py [--all]

Without --all: only current open positions (from HiPort latest snapshot)
With --all:    all ISINs in isin_analyst_map.csv, pulling last known HiPort price
"""

import sys
import warnings
import requests
import pyodbc
import pandas as pd
from datetime import datetime, timedelta

warnings.filterwarnings("ignore")

MODE_ALL = "--all" in sys.argv

# ─── Config ──────────────────────────────────────────────────────────────────

CONN_STR = (
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=fcedata01.fieracapital.ca;"
    "Database=ccl;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

FUND_PCODES = ("OEFOF", "OEFOGSSC", "OEFOGSSW")

YAHOO_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}

# ─── Step 1: Load ISINs ───────────────────────────────────────────────────────

print("Connecting to HiPort SQL...")
pcodes_sql = ", ".join(f"'{p}'" for p in FUND_PCODES)

with pyodbc.connect(CONN_STR) as conn:
    latest_vdate = pd.read_sql(
        f"SELECT TOP 1 VDATE FROM ccl.dbo.vw_RPT_VAL WHERE PCODE IN ({pcodes_sql}) ORDER BY VDATE DESC",
        conn
    ).iloc[0, 0]
    print(f"Latest VDATE: {latest_vdate}")

    if MODE_ALL:
        # Load analyst map ISINs + get last known price from HiPort for each
        analyst_df = pd.read_csv("isin_analyst_map.csv")
        all_isins = analyst_df["ISIN"].dropna().str.strip().str.upper().unique().tolist()
        isins_sql = ", ".join(f"'{i}'" for i in all_isins)

        # Get the most recent price per ISIN across all history
        snap_df = pd.read_sql(f"""
            SELECT v.ISIN, v.SNAME, v.CCY, v.LST_PRICE, v.UNITS, v.VDATE
            FROM ccl.dbo.vw_RPT_VAL v
            INNER JOIN (
                SELECT ISIN, MAX(VDATE) AS MAX_VDATE
                FROM ccl.dbo.vw_RPT_VAL
                WHERE PCODE IN ({pcodes_sql})
                  AND ISIN IN ({isins_sql})
                GROUP BY ISIN
            ) m ON v.ISIN = m.ISIN AND v.VDATE = m.MAX_VDATE
            WHERE v.PCODE IN ({pcodes_sql})
        """, conn)
        print(f"Loaded {len(snap_df)} ISINs from analyst map (last known HiPort price)\n")
    else:
        snap_df = pd.read_sql(f"""
            SELECT ISIN, SNAME, CCY, LST_PRICE, UNITS, VDATE='{latest_vdate}'
            FROM ccl.dbo.vw_RPT_VAL
            WHERE PCODE IN ({pcodes_sql})
              AND VDATE = '{latest_vdate}'
              AND ISIN IS NOT NULL
              AND UNITS != 0
            ORDER BY ISIN
        """, conn)
        print(f"Loaded {len(snap_df)} open positions\n")

# For Yahoo historical fetch, use latest_vdate as the reference
vdate = pd.Timestamp(latest_vdate)
period_end   = int((vdate + timedelta(days=1)).timestamp())
period_start = int((vdate - timedelta(days=5)).timestamp())

# ─── Step 2: Load existing ticker map ────────────────────────────────────────

ticker_df = pd.read_csv("oefof_pl/data/yahoo_tickers.csv")
ticker_map = dict(zip(
    ticker_df["isin"].str.strip().str.upper(),
    ticker_df["yahoo_ticker"].str.strip()
))
print(f"Loaded {len(ticker_map)} existing ticker mappings\n")

# ─── Step 3: OpenFIGI for unmapped ISINs ─────────────────────────────────────

snap_dedup = snap_df.drop_duplicates(subset="ISIN")
all_isins_in_snap = snap_dedup["ISIN"].str.strip().str.upper().tolist()
unmapped = [i for i in all_isins_in_snap if i not in ticker_map and str(i).strip()]

if unmapped:
    print(f"Attempting OpenFIGI lookup for {len(unmapped)} unmapped ISINs...")

    MIC_SUFFIX = {
        "XHKG": ".HK", "XSHG": ".SS", "XSHE": ".SZ",
        "XBOM": ".BO", "XNSE": ".NS", "XKRX": ".KS",
        "XTAI": ".TW", "XBKK": ".BK", "XIDX": ".JK",
        "XKLS": ".KL", "XPHS": ".PS", "XVNM": ".VN",
        "XSAU": ".SR", "XDFM": ".AE", "XADS": ".AD",
        "XJSE": ".JO", "XCAI": ".CA", "XIST": ".IS",
        "XNAI": ".NR", "XKAR": ".KA", "XWAR": ".WA",
        "XNYS": "", "XNAS": "", "ARCX": "",
        "XLON": ".L", "XFRA": ".F", "XPAR": ".PA",
        "XBUD": ".BD", "XPRA": ".PR", "XBUE": ".BA",
    }

    import time
    import json
    from pathlib import Path

    cache_path = Path("inputs/openfigi_cache.json")
    cache = json.loads(cache_path.read_text()) if cache_path.exists() else {}

    to_fetch = [i for i in unmapped if i not in cache]
    BATCH = 10

    for i in range(0, len(to_fetch), BATCH):
        batch = to_fetch[i:i + BATCH]
        payload = [{"idType": "ID_ISIN", "idValue": isin} for isin in batch]
        try:
            resp = requests.post(
                "https://api.openfigi.com/v3/mapping",
                headers={"Content-Type": "application/json"},
                json=payload, timeout=15, verify=False,
            )
            resp.raise_for_status()
            results = resp.json()
            for isin, result in zip(batch, results):
                if "error" in result:
                    cache[isin] = None
                    continue
                data = result.get("data") or []
                best = None
                for entry in data:
                    sec = str(entry.get("securityType") or "").lower()
                    if any(x in sec for x in ("option", "warrant", "future", "etf", "fund")):
                        continue
                    mic = str(entry.get("exchCode") or "").upper()
                    if mic not in MIC_SUFFIX:
                        continue
                    ticker = str(entry.get("ticker") or "").strip()
                    if ticker:
                        best = f"{ticker}{MIC_SUFFIX[mic]}"
                        break
                cache[isin] = best
            resolved = sum(1 for isin in batch if cache.get(isin))
            print(f"  batch {i//BATCH + 1}: resolved {resolved}/{len(batch)}")
        except Exception as e:
            print(f"  batch {i//BATCH + 1} failed: {e}")
        if i + BATCH < len(to_fetch):
            time.sleep(6)

    # Save updated cache
    cache_path.write_text(json.dumps(cache, indent=2, sort_keys=True))

    # Merge into ticker_map
    for isin, ticker in cache.items():
        if ticker and isin not in ticker_map:
            ticker_map[isin] = ticker

    print(f"After OpenFIGI: {len(ticker_map)} total tickers mapped\n")

# ─── Step 4: Fetch Yahoo historical close ────────────────────────────────────

def yahoo_close_on_date(ticker: str, p1: int, p2: int) -> tuple[float | None, str | None]:
    for url_base in [
        "https://query1.finance.yahoo.com",
        "https://query2.finance.yahoo.com",
    ]:
        url = f"{url_base}/v8/finance/chart/{ticker}?period1={p1}&period2={p2}&interval=1d"
        try:
            r = requests.get(url, headers=YAHOO_HEADERS, timeout=10, verify=False)
            r.raise_for_status()
            result = r.json().get("chart", {}).get("result", [{}])[0]
            timestamps = result.get("timestamp", [])
            closes = result.get("indicators", {}).get("quote", [{}])[0].get("close", [])
            for ts, close in reversed(list(zip(timestamps, closes))):
                if close is not None:
                    date_str = datetime.utcfromtimestamp(ts).strftime("%Y-%m-%d")
                    px = float(close)
                    if ticker.upper().endswith(".L"):
                        px = px / 100.0
                    return px, date_str
        except Exception:
            continue
    return None, None

print(f"Fetching Yahoo prices ({vdate.date()} window)...")
yahoo_prices: dict[str, tuple] = {}
for isin, ticker in ticker_map.items():
    if isin not in [r["ISIN"].strip().upper() for _, r in snap_dedup.iterrows()]:
        continue
    price, date_str = yahoo_close_on_date(ticker, period_start, period_end)
    yahoo_prices[isin] = (price, date_str)
    status = f"{price:.4f} ({date_str})" if price is not None else "FAILED"
    print(f"  {isin:20s} [{ticker:20s}]: {status}")

# ─── Step 5: Build comparison table ──────────────────────────────────────────

rows = []
for _, row in snap_dedup.iterrows():
    isin      = str(row["ISIN"]).strip().upper()
    ticker    = ticker_map.get(isin, "—")
    hiport_px = float(row["LST_PRICE"]) if pd.notna(row["LST_PRICE"]) else None
    yahoo_px, yahoo_date = yahoo_prices.get(isin, (None, None))
    ccy       = str(row["CCY"]).strip()

    diff_pct = (yahoo_px - hiport_px) / hiport_px * 100 if hiport_px and yahoo_px else None

    if ticker == "—":
        status = "NO TICKER"
    elif yahoo_px is None:
        status = "FETCH FAIL"
    else:
        status = "OK"

    rows.append({
        "ISIN":       isin,
        "Name":       str(row["SNAME"])[:28],
        "CCY":        ccy,
        "Ticker":     ticker,
        "HiPort Px":  round(hiport_px, 4) if hiport_px else None,
        "Yahoo Px":   round(yahoo_px, 4)  if yahoo_px is not None else None,
        "Yahoo Date": yahoo_date or "—",
        "Diff %":     round(diff_pct, 2)  if diff_pct is not None else None,
        "Status":     status,
    })

result_df = pd.DataFrame(rows).sort_values(
    "Diff %", key=lambda x: x.abs(), ascending=False, na_position="last"
)

# ─── Step 6: Print + save ────────────────────────────────────────────────────

pd.set_option("display.max_rows", 500)
pd.set_option("display.width", 140)
pd.set_option("display.float_format", "{:.4f}".format)

print("\n" + "="*120)
print(result_df.to_string(index=False))
print("="*120)

ok   = result_df[result_df["Status"] == "OK"]
fail = result_df[result_df["Status"] != "OK"]
print(f"\nSummary: {len(ok)} OK  |  {len(fail)} missing/failed")

if len(ok):
    big = ok[ok["Diff %"].abs() > 5]
    print(f"Large discrepancies (>5%): {len(big)}")
    if len(big):
        print("\n" + big[["ISIN","Name","CCY","Ticker","HiPort Px","Yahoo Px","Yahoo Date","Diff %"]].to_string(index=False))

result_df.to_csv("test_prices_output.csv", index=False)
print("\nSaved to test_prices_output.csv")

# ─── Step 7: Save updated ticker map ─────────────────────────────────────────

if MODE_ALL and unmapped:
    new_map = {k: v for k, v in ticker_map.items() if v}
    out_df = pd.DataFrame([{"isin": k, "yahoo_ticker": v} for k, v in sorted(new_map.items())])
    out_df.to_csv("oefof_pl/data/yahoo_tickers.csv", index=False)
    print(f"Updated yahoo_tickers.csv with {len(out_df)} entries")