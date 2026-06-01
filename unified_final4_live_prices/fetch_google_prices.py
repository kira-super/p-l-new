import json
import os
from typing import Iterable

import pandas as pd
import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


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


def _parse_google_info_payload(payload: str) -> list[dict]:
    """Parse historical Google finance/info response formats."""
    text = payload.strip()

    # Older format is prefixed with //
    if text.startswith("//"):
        text = text[2:].strip()

    # Some responses may contain XSSI prefix
    if text.startswith(")]}'"):
        lines = text.splitlines()
        text = "\n".join(lines[1:]).strip() if len(lines) > 1 else ""

    if not text:
        return []

    try:
        parsed = json.loads(text)
        return parsed if isinstance(parsed, list) else []
    except json.JSONDecodeError:
        return []


def get_google_finance_batch(tickers: Iterable[str]) -> dict[str, float]:
    """Fetches last-trade prices for Google tickers via finance/info endpoint."""
    normalized_tickers = [str(t).strip() for t in tickers if str(t).strip()]
    if not normalized_tickers:
        return {}

    ticker_string = ",".join(normalized_tickers)
    url = f"https://www.google.com/finance/info?q={ticker_string}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    }

    prices: dict[str, float] = {}
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code != 200:
            return prices

        data = _parse_google_info_payload(response.text)
        for item in data:
            exchange = item.get("e")
            symbol = item.get("t")
            last_trade = item.get("l")
            if not exchange or not symbol or last_trade is None:
                continue
            try:
                prices[f"{exchange}:{symbol}"] = float(str(last_trade).replace(",", ""))
            except ValueError:
                continue
    except requests.RequestException:
        return prices

    return prices


def _load_input_matrix() -> tuple[pd.DataFrame, str] | tuple[None, None]:
    candidates = [
        "out/universe_valuation_matrix.csv",
        "out/valuation_variance_analysis.csv",
    ]
    for candidate in candidates:
        candidate_path = os.path.join(BASE_DIR, candidate)
        if os.path.exists(candidate_path):
            return pd.read_csv(candidate_path), candidate_path
    return None, None


def patch_matrix_with_google() -> None:
    print("=" * 80)
    print("        OAKS HARDCODED UNIVERSE MULTI-FEED PRICE INJECTION ENGINE")
    print("=" * 80)

    matrix_df, source_path = _load_input_matrix()
    if matrix_df is None:
        print("[CRITICAL] Could not find out/universe_valuation_matrix.csv or out/valuation_variance_analysis.csv")
        return

    required_cols = {"ISIN", "Google Ticker"}
    missing_cols = required_cols.difference(set(matrix_df.columns))
    if missing_cols:
        print(f"[CRITICAL] Missing required columns in {source_path}: {sorted(missing_cols)}")
        return

    hiport_col = ""
    for candidate in ("HiPort Price", "HiPort Close"):
        if candidate in matrix_df.columns:
            hiport_col = candidate
            break
    if not hiport_col:
        print(f"[CRITICAL] Missing HiPort value column in {source_path}: expected 'HiPort Price' or 'HiPort Close'")
        return

    print(f"Loaded matrix source: {source_path}")
    print("Mapping explicit structural ticker overrides...")

    google_tickers_to_query: list[str] = []
    for idx, row in matrix_df.iterrows():
        isin = str(row["ISIN"]).strip()
        mapped_ticker = ISIN_TO_GOOGLE_TICKER.get(isin)

        if mapped_ticker:
            matrix_df.at[idx, "Google Ticker"] = mapped_ticker
            google_tickers_to_query.append(mapped_ticker)
        else:
            existing_ticker = str(row.get("Google Ticker", "")).strip()
            if existing_ticker and existing_ticker.upper() != "NO TICKER":
                google_tickers_to_query.append(existing_ticker)

    deduped_tickers = list(dict.fromkeys(google_tickers_to_query))
    print(f"Fetching Google prices for {len(deduped_tickers)} tickers...")

    google_price_cache: dict[str, float] = {}
    batch_size = 50
    for i in range(0, len(deduped_tickers), batch_size):
        batch = deduped_tickers[i : i + batch_size]
        google_price_cache.update(get_google_finance_batch(batch))

    print("Injecting prices and recalculating variance columns...")
    for idx, row in matrix_df.iterrows():
        g_ticker = str(matrix_df.at[idx, "Google Ticker"]).strip()
        if not g_ticker:
            continue

        if g_ticker in google_price_cache:
            g_price = google_price_cache[g_ticker]
            matrix_df.at[idx, "Google Price"] = f"{g_price:.2f}"

            hiport_price = pd.to_numeric(row.get(hiport_col), errors="coerce")
            if pd.notna(hiport_price) and float(hiport_price) > 0:
                g_var = (abs(float(hiport_price) - g_price) / float(hiport_price)) * 100

                if "HiPort vs Yahoo Var" in matrix_df.columns:
                    existing = str(row.get("HiPort vs Yahoo Var", ""))
                    matrix_df.at[idx, "HiPort vs Yahoo Var"] = f"{existing} | G_Var: {g_var:.2f}%"
                else:
                    matrix_df.at[idx, "HiPort vs Google Var"] = f"{g_var:.2f}%"
        else:
            if g_ticker.upper() != "NO TICKER":
                matrix_df.at[idx, "Google Price"] = "UNAVAILABLE_HISTORICAL"

    output_path = os.path.join(BASE_DIR, "out", "final_hardcoded_valuation_matrix.csv")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    matrix_df.to_csv(output_path, index=False)

    print("\n" + "-" * 70)
    print(f"[SUCCESS] Completed integration. Saved to: {output_path}")
    print("-" * 70)


if __name__ == "__main__":
    patch_matrix_with_google()