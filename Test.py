import yfinance as yf
import pandas as pd
import requests, time
from datetime import datetime

# -----------------------------
# NSE API Fallback (with cookies)
# -----------------------------
def get_price_from_nse(symbol):
    try:
        url = f"https://www.nseindia.com/api/quote-equity?symbol={symbol}"
        session = requests.Session()
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json",
            "Referer": "https://www.nseindia.com/"
        }
        session.headers.update(headers)
        session.get("https://www.nseindia.com", timeout=10)  # get cookies
        r = session.get(url, timeout=10)
        data = r.json()
        return data['priceInfo']['lastPrice']
    except Exception as e:
        print(f"NSE fetch failed for {symbol} → {e}")
        return None

# -----------------------------
# Yahoo fetch (fix tz bug)
# -----------------------------
def get_returns_yahoo(symbol: str):
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period="1y", interval="1d")
        if hist.empty:
            print(f"Yahoo empty for {symbol}")
            return None

        # Remove timezone for comparisons
        hist.index = hist.index.tz_localize(None)
        current = hist['Close'].iloc[-1]

        def pct(days):
            if len(hist) > days:
                return (current / hist['Close'].iloc[-days] - 1) * 100
            return None

        start_year = datetime(datetime.today().year, 1, 1)
        ytd_data = hist[hist.index >= start_year]
        ytd_price = ytd_data['Close'].iloc[0] if not ytd_data.empty else None

        return {
            "Current Price (₹)": round(float(current), 2),
            "1W %": round(pct(5), 2) if pct(5) else None,
            "2W %": round(pct(10), 2) if pct(10) else None,
            "1M %": round(pct(21), 2) if pct(21) else None,
            "6M %": round(pct(126), 2) if pct(126) else None,
            "YTD %": round((current / ytd_price - 1) * 100, 2) if ytd_price else None,
            "Last Updated": datetime.today().strftime("%d-%m-%Y")
        }
    except Exception as e:
        print(f"Yahoo error for {symbol}: {e}")
        return None

# -----------------------------
# Trend logic
# -----------------------------
def compute_trend(row):
    try:
        if row.get("YTD %") is None or pd.isna(row.get("YTD %")):
            return "Unknown"
        if row["YTD %"] > 10:
            return "Bullish"
        elif row["YTD %"] < -10:
            return "Bearish"
        else:
            return "Neutral"
    except:
        return "Unknown"

# -----------------------------
# Update Excel
# -----------------------------
def update_excel(input_file, output_file):
    df = pd.read_excel(input_file)

    # Clean column names
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.replace('\u00A0', ' ', regex=True)

    print("Detected columns:", df.columns.tolist())

    # Auto-detect Stock Name column (first column in file)
    stock_col = df.columns[0]

    # Ensure required columns exist
    for col in ["Symbol", "Current Price (₹)", "1W %", "2W %", "1M %", "6M %", "YTD %", "Trend", "Last Updated"]:
        if col not in df.columns:
            df[col] = None

    # Custom verified mapping (extend as needed)
    mapping = {
        "Bajaj Auto Ltd": "BAJAJ-AUTO.NS",
        "Havells India Ltd": "HAVELLS.NS",
        "Lupin Ltd": "LUPIN.NS",
        "Astral Ltd": "ASTRAL.NS",
        "Marksans Pharma Ltd": "MARKSANS.NS",
        "Zydus Lifesciences": "ZYDUSLIFE.NS",
        "NBCC": "NBCC.NS",
        "Finolex Cable": "FINCABLES.NS",
        "Rashtriya Chemicals and Fertilizers Ltd": "RCF.NS",
        "Deepak Frtlsrs and Ptrchmcls Corp Ltd": "DEEPAKFERT.NS",
        # Add more mappings here...
    }

    for idx, row in df.iterrows():
        name = str(row[stock_col]).strip()
        if not name or name.lower() == 'nan':
            continue

        symbol = mapping.get(name)
        if not symbol:
            # Fallback guess → first word + ".NS"
            symbol = name.split(" ")[0].upper() + ".NS"

        print(f"\nFetching: {name} ({symbol})")

        res = None
        if symbol.endswith(".NS"):
            res = get_returns_yahoo(symbol)
            time.sleep(0.5)

        if not res:
            nse_symbol = symbol.replace(".NS", "")
            price = get_price_from_nse(nse_symbol)
            if price:
                res = {
                    "Current Price (₹)": price,
                    "1W %": None, "2W %": None, "1M %": None,
                    "6M %": None, "YTD %": None,
                    "Last Updated": datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                }

        if not res:
            res = {"Error": "Data not found"}

        # Update DataFrame row
        for col in res:
            df.at[idx, col] = res[col]
        df.at[idx, "Symbol"] = symbol
        df.at[idx, "Trend"] = compute_trend(res)

    # Save new Excel
    df.to_excel(output_file, index=False)
    print(f"\n✅ Updated data written to {output_file}")

# -----------------------------
# Run
# -----------------------------
update_excel("Stock-List.xlsx", "Stock-List-Updated.xlsx")
