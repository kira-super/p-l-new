import os
import re
import warnings
import pandas as pd
import requests

warnings.filterwarnings("ignore")

def fetch_yahoo_historical_close(ticker: str, target_date: str) -> float | None:
    """Queries Yahoo Finance for the closing price on the target date."""
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    try:
        t_epoch = int(pd.Timestamp(target_date).timestamp())
        # 7-day lookback window to bridge holidays/weekends
        params = {"period1": t_epoch - 432000, "period2": t_epoch + 172800, "interval": "1d"}
        r = requests.get(url, headers=headers, params=params, timeout=7, verify=False)
        res = r.json()["chart"]["result"][0]
        closes = res["indicators"]["quote"][0]["close"]
        valid_closes = [c for c in closes if c is not None]
        return valid_closes[-1] if valid_closes else None
    except Exception:
        return None

def build_comparison_analysis():
    print("=" * 80)
    print("        OAKS UNIVERSE MULTI-SOURCE PRICE VARIANCE ENGINE")
    print("=" * 80)

    # Validate asset files inside the running subfolder
    if not os.path.exists("hiport_all_isins.csv"):
        print("[CRITICAL] Run the engine inside 'unified_final4_live_prices' containing 'hiport_all_isins.csv'.")
        return

    hiport_df = pd.read_csv("hiport_all_isins.csv")
    
    ticker_map = {}
    if os.path.exists("oefof_pl/data/yahoo_tickers.csv"):
        tm_df = pd.read_csv("oefof_pl/data/yahoo_tickers.csv")
        
        # DEFINITIVE FIX: Force all column headers to be completely lowercase
        tm_df.columns = tm_df.columns.str.lower()
        
        if "isin" in tm_df.columns and "yahoo_ticker" in tm_df.columns:
            ticker_map = dict(zip(tm_df["isin"], tm_df["yahoo_ticker"]))

    target_date = "2026-05-31"
    valid_positions = hiport_df[hiport_df["LST_PRICE"].notna()].drop_duplicates(subset=["ISIN"])
    
    matrix_rows = []
    
    # Track the Excel row numbering (Row 1 is headers, data rows begin on Row 2)
    row_num = 2 
    
    print(f"Generating live-calculating formula framework for {len(valid_positions)} rows...")
    
    for _, row in valid_positions.iterrows():
        isin = row["ISIN"]
        sname = row.get("SNAME", "UNKNOWN")
        ccy = row.get("CCY", "EUR")
        hiport_price = float(row["LST_PRICE"])
        
        yahoo_ticker = ticker_map.get(isin, "")
        yahoo_price = ""
        google_ticker = ""
        
        # Fetch matching Yahoo historical closing values
        if pd.notna(yahoo_ticker) and yahoo_ticker != "":
            y_val = fetch_yahoo_historical_close(yahoo_ticker, target_date)
            if y_val is not None:
                # Standard 100x penny stock rule conversion for London (.L) listings
                yahoo_price = y_val / 100.0 if str(yahoo_ticker).endswith(".L") else y_val

        # Infer explicit Google Ticker formats based on exchange rules
        if isin.startswith("AEA") or isin.startswith("AEE") or isin.startswith("AEN"):
            google_ticker = f"ADX:{sname.split()[0]}"
        elif isin.startswith("PHY"):
            google_ticker = f"PSE:{sname.split()[0]}"
        elif isin.startswith("SA"):
            numeric_code = re.findall(r"\d+", str(yahoo_ticker))
            google_ticker = f"TADAWUL:{numeric_code[0]}" if numeric_code else f"TADAWUL:{sname.split()[0]}"
        elif isin.startswith("PL"):
            google_ticker = f"WSE:{sname.split()[0]}"
        else:
            google_ticker = str(yahoo_ticker).replace(".HK", ":HKG").replace(".NS", ":NSE").replace(".SA", ":BVMF")

        # Map sheet columns: D=HiPort Price, F=Yahoo Price, H=Google Price
        google_price_formula = f'=INDEX(GOOGLEFINANCE("{google_ticker}", "close", DATE(2026,5,31)), 2, 2)' if google_ticker else ""
        
        # 1. Variance: HiPort vs Yahoo
        diff_hp_vs_yh = f'=IF(ISNUMBER(F{row_num}), ABS(D{row_num}-F{row_num})/D{row_num}, "")'
        
        # 2. Variance: HiPort vs Google
        diff_hp_vs_go = f'=IF(ISNUMBER(H{row_num}), ABS(D{row_num}-H{row_num})/D{row_num}, "")'
        
        # 3. Variance: Yahoo vs Google
        diff_yh_vs_go = f'=IF(AND(ISNUMBER(F{row_num}), ISNUMBER(H{row_num})), ABS(F{row_num}-H{row_num})/F{row_num}, "")'
        
        # 4. Global Average: Three-way pricing summary evaluation
        avg_three_way = f'=AVERAGE(D{row_num}, IFERROR(F{row_num}, D{row_num}), IFERROR(H{row_num}, D{row_num}))'

        matrix_rows.append({
            "ISIN": isin,
            "Security Name": sname[:25],
            "Currency": ccy,
            "HiPort Price": hiport_price,                # Column D
            "Yahoo Ticker": yahoo_ticker,
            "Yahoo Price": yahoo_price,                  # Column F
            "Google Ticker": google_ticker,
            "Google Price": google_price_formula,        # Column H
            "Diff HP vs YH": diff_hp_vs_yh,              # Column I
            "Diff HP vs GO": diff_hp_vs_go,              # Column J
            "Diff YH vs GO": diff_yh_vs_go,              # Column K
            "Avg HP YH GO": avg_three_way                # Column L
        })
        row_num += 1

    output_df = pd.DataFrame(matrix_rows)
    os.makedirs("out", exist_ok=True)
    
    file_path = "out/valuation_variance_analysis.csv"
    output_df.to_csv(file_path, index=False)
    
    print("\n" + "-" * 80)
    print(f"[SUCCESS] Multi-source validation analysis matrix file created at: {file_path}")
    print("-> Next Step: Upload/Import this file directly into Google Sheets.")
    print("-> Every comparative deviation matrix path will calculate out live.")
    print("-" * 80)

if __name__ == "__main__":
    build_comparison_analysis()