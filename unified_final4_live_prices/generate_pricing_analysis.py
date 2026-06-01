import os
import re
import warnings
import pandas as pd
import requests

warnings.filterwarnings("ignore")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Verified hard overrides for Google Finance tickers where name-based guessing fails.
ISIN_TO_GOOGLE_TICKER = {
    # UAE (ADX/DFM)
    "AEA000201011": "ADX:ADCB",
    "AEA000801018": "ADX:ADIB",
    "AEN000101016": "ADX:FAB",
    "AEA006101017": "ADX:ADNOCDRILL",
    "AEA007501017": "ADX:YAHSAT",
    "AEE000401019": "ADX:EAND",
    "AEE01195A234": "ADX:ADNOCGAS",
    "AEE01268A239": "ADX:ADNOCLS",
    "AEE01135A222": "ADX:AMERICANA",
    "AEF000901015": "ADX:FERTIGLOBE",
    "AEE01356D236": "DFM:DUBAITAXI",
    "AEE01377S248": "DFM:SPINNEYS",
    "AEE01110S227": "DFM:SALIK",
    "AEE01134E227": "DFM:EMPOWER",
    # Saudi (Tadawul)
    "SA16CI8KMOH3": "TADAWUL:9649",
    "SA15M1HH2NH5": "TADAWUL:2222",
    "SA15T1L22JH8": "TADAWUL:1831",
    "SA15ED94KR18": "TADAWUL:9543",
    "SA165I21VPH1": "TADAWUL:4083",
    "SA1610O13M15": "TADAWUL:2100",
    "SA154HG210H6": "TADAWUL:4161",
    "SA15LGLI0N19": "TADAWUL:4110",
    # HK / India / Poland / Vietnam
    "KYG070341048": "HKG:9888",
    "KYG371091086": "HKG:1448",
    "CNE100004272": "HKG:9633",
    "INE474Q01031": "NSE:MEDANTA",
    "INE066P01011": "NSE:INOXWIND",
    "VN000000FPT1": "HOSE:FPT",
    "VN000000SSI1": "HOSE:SSI",
    "PLBNFTS00018": "WSE:BFT",
    "PLDINPL00011": "WSE:DNP",
}


def infer_google_ticker(isin: str, security_name: str, yahoo_ticker: str) -> str:
    """Returns the best Google ticker candidate for an instrument."""
    if isin in ISIN_TO_GOOGLE_TICKER:
        return ISIN_TO_GOOGLE_TICKER[isin]

    ticker = "" if pd.isna(yahoo_ticker) else str(yahoo_ticker).strip()
    if ticker:
        suffix_to_exchange = {
            ".HK": "HKG",
            ".NS": "NSE",
            ".SA": "BVMF",
            ".BO": "BSE",
            ".WA": "WSE",
        }
        for suffix, exchange in suffix_to_exchange.items():
            if ticker.endswith(suffix):
                base_symbol = ticker[: -len(suffix)]
                return f"{exchange}:{base_symbol}"

    # Avoid incorrect first-word guessing for UAE names (e.g., ADX:FIRST / ADX:ABU).
    if isin.startswith(("AEA", "AEE", "AEN", "AEF")):
        return ""

    if isin.startswith("SA"):
        numeric_code = re.findall(r"\d+", ticker)
        return f"TADAWUL:{numeric_code[0]}" if numeric_code else ""

    if isin.startswith("PHY"):
        first_word = str(security_name).split()[0] if str(security_name).split() else ""
        return f"PSE:{first_word}" if first_word else ""

    if isin.startswith("PL"):
        first_word = str(security_name).split()[0] if str(security_name).split() else ""
        return f"WSE:{first_word}" if first_word else ""

    return ""

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

    hiport_path = os.path.join(BASE_DIR, "hiport_all_isins.csv")
    ticker_map_path = os.path.join(BASE_DIR, "oefof_pl", "data", "yahoo_tickers.csv")
    out_dir = os.path.join(BASE_DIR, "out")

    # Validate required assets relative to the script location
    if not os.path.exists(hiport_path):
        print(f"[CRITICAL] Missing required input file: {hiport_path}")
        return

    hiport_df = pd.read_csv(hiport_path)
    
    ticker_map = {}
    if os.path.exists(ticker_map_path):
        tm_df = pd.read_csv(ticker_map_path)
        
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

        # Infer Google ticker with deterministic mapping and exchange suffix translation.
        google_ticker = infer_google_ticker(isin=isin, security_name=sname, yahoo_ticker=yahoo_ticker)

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
    os.makedirs(out_dir, exist_ok=True)
    
    file_path = os.path.join(out_dir, "valuation_variance_analysis.csv")
    output_df.to_csv(file_path, index=False)
    
    print("\n" + "-" * 80)
    print(f"[SUCCESS] Multi-source validation analysis matrix file created at: {file_path}")
    print("-> Next Step: Upload/Import this file directly into Google Sheets.")
    print("-> Every comparative deviation matrix path will calculate out live.")
    print("-" * 80)

if __name__ == "__main__":
    build_comparison_analysis()