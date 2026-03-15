import pandas as pd
import os
import re


def process_portfolio_stocks(excel_file_path, threshold=30000):
    """
      Portfolio scanner with configurable threshold + count reporting
    """
    STOCK_NAMES_TEXT = """
Indo Tech.Trans.
Sharda Motor
Cigniti Tech.
BLS Internat.
Banco Products
Premier Polyfilm
Fiem Industries
Gandhi Spl. Tube
Ceinsys Tech
Dodla Dairy
Dynamic Cables
Caplin Point Lab
Interarch Build.
Shri Ahimsa
Saksoft
Antelopus Selan
    """

    MARKET_VALUE_THRESHOLD = threshold
    stock_names = [name.strip().strip('"') for name in STOCK_NAMES_TEXT.strip().split('\n') if name.strip()]
    print(f"📊 Processing {len(stock_names)} stock names")

    # File existence check
    if not os.path.exists(excel_file_path):
        print(f"❌ File not found: {os.path.abspath(excel_file_path)}")
        return None

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

        name_col = df.iloc[:, 1]  # Column B
        code_col = df.iloc[:, 0]  # Column A
        value_col = df.iloc[:, 7]  # Column H

        # Pre-compute Excel first words for efficiency
        excel_first_words = name_col.astype(str).str.strip().str.upper().str.split().str[0]

        low_value_stocks = []
        high_value_stocks = []
        not_found_stocks = []

        for stock_name in stock_names:
            found_match = False
            name_upper = stock_name.upper().strip()

            # STEP 1: Exact match
            matching_rows = name_col[
                name_col.astype(str).str.strip().str.upper() == name_upper
                ]

            # STEP 2: Full name contains (if multi-word input)
            if matching_rows.empty and len(stock_name.split()) > 1:
                full_pattern = r'\b' + re.escape(name_upper) + r'\b'
                matching_rows = name_col[
                    name_col.astype(str).str.contains(full_pattern, case=False, na=False, regex=True)
                ]

            # STEP 3: Input matches FIRST WORD of Excel name (NEW! Works for ALL inputs)
            if matching_rows.empty:
                if name_upper in excel_first_words.values:
                    matching_rows = name_col[
                        excel_first_words == name_upper
                        ]

            # STEP 4: Input first word matches Excel first word (multi-word fallback)
            if matching_rows.empty and len(stock_name.split()) > 1:
                input_first_word = stock_name.split()[0].strip().upper()
                if input_first_word in excel_first_words.values:
                    matching_rows = name_col[
                        excel_first_words == input_first_word
                        ]

            if not matching_rows.empty:
                row_index = matching_rows.index[0]
                stock_code = code_col.iloc[row_index]
                excel_name = name_col.iloc[row_index]

                # Safe market value extraction
                try:
                    market_value = float(value_col.iloc[row_index])
                except (ValueError, TypeError):
                    print(f"⚠️ Skipping {stock_name}: Invalid market value")
                    not_found_stocks.append(stock_name)
                    continue

                found_match = True

                if market_value < MARKET_VALUE_THRESHOLD:
                    low_value_stocks.append({
                        'stock_name': stock_name,
                        'stock_code': stock_code,
                        'excel_name': excel_name,
                        'market_value': market_value
                    })
                else:
                    high_value_stocks.append({
                        'stock_name': stock_name,
                        'stock_code': stock_code,
                        'excel_name': excel_name,
                        'market_value': market_value
                    })

            if not found_match:
                not_found_stocks.append(stock_name)

        print(f"\n✅ Found {len(low_value_stocks)} stocks < ₹{MARKET_VALUE_THRESHOLD:,}")

        # # Export low value stocks to CSV
        # if low_value_stocks:
        #     output_df = pd.DataFrame(low_value_stocks)
        #     output_df.to_csv('low_value_stocks.csv', index=False)
        #     print("💾 Results saved to low_value_stocks.csv")

        return {
            'total_checked': len(stock_names),
            'low_value_stocks': low_value_stocks,
            'high_value_stocks': high_value_stocks,
            'not_found_stocks': not_found_stocks,
            'threshold': MARKET_VALUE_THRESHOLD
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
        print("-" * 90)
        for stock in result['low_value_stocks']:
            print(
                f"💰 {stock['stock_name']:<15} | {stock['stock_code']:<10} | {stock['excel_name']:<35} | ₹{stock['market_value']:>12,.0f}")

        if result['not_found_stocks']:
            print(f"\n❌ STOCKS NOT FOUND IN PORTFOLIO: [{len(result['not_found_stocks'])}]")
            print("-" * 90)
            for stock_name in result['not_found_stocks']:
                print(f"🔍 {stock_name:<15}                           | Not in your holdings")

        if result['high_value_stocks']:
            print(f"\n⚠️  STOCKS ABOVE THRESHOLD (≥ {threshold_display}): [{len(result['high_value_stocks'])}]")
            print("-" * 90)
            for stock in result['high_value_stocks']:
                print(
                    f"📈 {stock['stock_name']:<15} | {stock['stock_code']:<10} | {stock['excel_name']:<35} | ₹{stock['market_value']:>12,.0f}")

        print("\n" + "=" * 90)
        print("📊 SUMMARY REPORT")
        print("=" * 90)
        print(f"✅ STOCKS FOUND IN PORTFOLIO < {threshold_display}: {len(result['low_value_stocks'])}")
        print(f"⚠️  STOCKS FOUND ≥ {threshold_display}: {len(result['high_value_stocks'])}")
        print(f"❌ STOCKS NOT FOUND IN PORTFOLIO: {len(result['not_found_stocks'])}")
        print(f"📈 Total Stocks Checked: {result['total_checked']}")
        print("=" * 90)
