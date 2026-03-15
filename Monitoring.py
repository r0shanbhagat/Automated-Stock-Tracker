import pandas as pd
import os
import re


def process_portfolio_stocks(excel_file_path, threshold=30000):
    """
    Portfolio scanner with configurable threshold + count reporting
    """

    # 🆕 CONFIGURABLE PRICE THRESHOLD
    MARKET_VALUE_THRESHOLD = threshold  # Change to 20000, 25000 anytime!

    STOCK_CODES_TEXT = """
	NETSTO		ANGBRO	COMAGE	IIFWEA	MCX
      
    """

    # Advanced extraction + cleaning
    stock_codes = []
    lines = STOCK_CODES_TEXT.strip().split()

    for item in lines:
        clean_item = re.sub(r'\[.*?\]', '', item).strip()
        comma_items = clean_item.split(',')
        for comma_item in comma_items:
            comma_item = comma_item.strip()
            if not comma_item:
                continue
            if comma_item in ['BUY', 'OB']:
                continue
            clean_stock = re.sub(r'-[A-Z]+$', '', comma_item).strip()
            clean_stock = re.sub(r'[^\w\s/]', '', clean_stock).strip()

            if clean_stock:
                if '/' in clean_stock:
                    split_stocks = clean_stock.split('/')
                    for split_stock in split_stocks:
                        final_stock = re.sub(r'[^\w]', '', split_stock).strip()
                        if final_stock and len(final_stock) > 1:
                            stock_codes.append(final_stock.upper())
                else:
                    final_stock = re.sub(r'[^\w]', '', clean_stock).strip()
                    if final_stock and len(final_stock) > 1:
                        stock_codes.append(final_stock.upper())

    # Remove duplicates
    seen = set()
    unique_stock_codes = []
    for stock in stock_codes:
        if stock not in seen:
            unique_stock_codes.append(stock)
            seen.add(stock)

    stock_codes = unique_stock_codes
    print(f"📊 Processing {len(stock_codes)} cleaned stock codes: {stock_codes}")

    try:
        engines = ['openpyxl', 'xlrd', 'calamine', 'odf']
        df = None

        for engine in engines:
            try:
                df = pd.read_excel(excel_file_path, engine=engine)
                print(f"✅ Loaded with {engine} engine")
                break
            except:
                continue

        if df is None:
            print("Trying CSV fallback...")
            df = pd.read_csv(excel_file_path)
            print("✅ Loaded as CSV")

        stock_col = df.iloc[:, 0]  # Column A
        value_col = df.iloc[:, 7]  # Column H

        low_value_stocks = []
        high_value_stocks = []  # 🆕 NEW: Track stocks >= threshold
        not_found_stocks = []

        for stock_code in stock_codes:
            matching_rows = stock_col[stock_col.astype(str).str.strip().str.upper() == stock_code]

            if not matching_rows.empty:
                row_index = matching_rows.index[0]
                market_value = value_col.iloc[row_index]

                if pd.notna(market_value) and isinstance(market_value, (int, float)):
                    if market_value < MARKET_VALUE_THRESHOLD:
                        low_value_stocks.append({
                            'stock_code': stock_code,
                            'market_value': market_value
                        })
                    else:
                        # 🆕 NEW: Stock found but >= threshold
                        high_value_stocks.append({
                            'stock_code': stock_code,
                            'market_value': market_value
                        })
            else:
                not_found_stocks.append(stock_code)

        print(f"\n✅ Found {len(low_value_stocks)} stocks with market value < ₹{MARKET_VALUE_THRESHOLD:,}")

        return {
            'total_checked': len(stock_codes),
            'low_value_stocks': low_value_stocks,
            'high_value_stocks': high_value_stocks,  # 🆕 NEW
            'not_found_stocks': not_found_stocks,
            'all_stock_codes': stock_codes,
            'threshold': MARKET_VALUE_THRESHOLD  # 🆕 NEW: Store for display
        }

    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return None


# 🚀 RUN IT
if __name__ == "__main__":
    excel_file = "PortFolioEqtSummary.csv"
    result = process_portfolio_stocks(excel_file, threshold=30000)

    if result:
        threshold_display = f"₹{result['threshold']:,}"

        print(f"\n🎯 LOW VALUE STOCKS (< {threshold_display}): [{len(result['low_value_stocks'])}]")
        print("-" * 40)
        for stock in result['low_value_stocks']:
            print(f"💰 {stock['stock_code']:15} | ₹{stock['market_value']:>10,.0f}")

        if result['not_found_stocks']:
            print(f"\n❌ STOCKS NOT FOUND IN PORTFOLIO: [{len(result['not_found_stocks'])}]")
            print("-" * 40)
            for stock in result['not_found_stocks']:
                print(f"🔍 {stock:15} | Not in your holdings")

        # 🆕 NEW: Display high value stocks
        if result['high_value_stocks']:
            print(f"\n⚠️  STOCKS ABOVE THRESHOLD (≥ {threshold_display}): [{len(result['high_value_stocks'])}]")
            print("-" * 40)
            for stock in result['high_value_stocks']:
                print(f"📈 {stock['stock_code']:15} | ₹{stock['market_value']:>10,.0f}")

        print("\n" + "=" * 50)
        print("📊 SUMMARY REPORT")
        print("=" * 50)
        print(f"✅ STOCKS FOUND IN PORTFOLIO < {threshold_display}: {len(result['low_value_stocks'])}")
        print(f"⚠️  STOCKS FOUND ≥ {threshold_display}: {len(result['high_value_stocks'])}")
        print(f"❌ STOCKS NOT FOUND IN PORTFOLIO: {len(result['not_found_stocks'])}")
        print(f"📈 Total Stocks Checked: {result['total_checked']}")
        print("=" * 50)
