import yfinance as yf
import pandas as pd
import requests, time
from datetime import datetime

from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill


# ---------------------------------
# NSE API Fallback (with cookies)
# ---------------------------------
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
        # Fetch homepage first for session cookies
        session.get("https://www.nseindia.com", timeout=10)
        r = session.get(url, timeout=10)
        data = r.json()
        return data['priceInfo']['lastPrice']
    except Exception as e:
        print(f"NSE fetch failed for {symbol} → {e}")
        return None


# ---------------------------------
# Yahoo fetch (with tz bug fix)
# ---------------------------------
def get_returns_yahoo(symbol):
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period="1y", interval="1d")
        if hist.empty:
            print(f"Yahoo returned empty for {symbol}")
            return None

        # Fix timezone mismatch
        hist.index = hist.index.tz_localize(None)

        current = hist['Close'].iloc[-1]

        # For 1D, should be (-1) vs (-2)
        def pct(days):
            if len(hist) > days:
                return (hist['Close'].iloc[-1] / hist['Close'].iloc[-2] - 1) * 100 if days == 1 else \
                    (current / hist['Close'].iloc[-days] - 1) * 100
            return None

        start_year = datetime(datetime.today().year, 1, 1)
        ytd_data = hist[hist.index >= start_year]
        ytd_price = ytd_data['Close'].iloc[0] if not ytd_data.empty else None
        ytd_price_perct = round((current / ytd_price - 1) * 100, 2)
        qrt_price_perct = round(pct(63), 2) if pct(63) is not None else None

        return {
            "Current Price (₹)": round(float(current), 2),
            "1D %": round(pct(1), 2) if pct(1) else None,
            "1W %": round(pct(5), 2) if pct(5) else None,
            "2W %": round(pct(10), 2) if pct(10) else None,
            "1M %": round(pct(21), 2) if pct(21) else None,
            "3M %":qrt_price_perct,
            "6M %": round(pct(126), 2) if pct(126) else None,
            "YTD %": ytd_price_perct if ytd_price else None,
            "Trend": compute_trend(qrt_price_perct) if qrt_price_perct is not None else None,
            "Last Updated": datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        }
    except Exception as e:
        print(f"Error fetching from Yahoo for {symbol} → {e}")
        return None


# -----------------------------
# Trend logic
# -----------------------------
def compute_trend(ytd_pricePerct):
    try:
        if ytd_pricePerct is None:
            return "Unknown"
        if ytd_pricePerct > 10:
            return "Bullish"
        elif ytd_pricePerct < -10:
            return "Bearish"
        else:
            return "Neutral"
    except:
        return "Unknown"


# ---------------------------------
# Example Stock Mapping
# (⚠️ You will complete this with all 87 from Excel)
# ---------------------------------
stocks = {
    "Marksans Pharma Ltd": "MARKSANS.NS",
    "Astral Ltd": "ASTRAL.NS",
    "Ice Make Refrigeration Ltd": "ICEMAKE.NS",
    "Mahanagar Gas Ltd": "MGL.NS",
    "Sanofi India Ltd": "SANOFI.NS",
    "Clean Science and Technology Ltd": "CLEAN.NS",
    "Cigniti Technologies Ltd": "CIGNITITEC.NS",
    "Symphony Ltd": "SYMPHONY.NS",
    "Finolex Cables Ltd": "FINCABLES.NS",
    "Bajaj Auto Ltd": "BAJAJ-AUTO.NS",
    "Rashtriya Chemicals and Fertilizers Ltd": "RCF.NS",
    "Deepak Fertilizers and Petrochemicals Corp Ltd": "DEEPAKFERT.NS",
    "Gujarat State Fertilizers & Chemicals Ltd": "GSFC.NS",
    "Thangamayil Jewellery Ltd": "THANGAMAYL.NS",
    "Fineotex Chemical Ltd": "FCL.NS",
    "Alkyl Amines Chemicals Ltd": "ALKYLAMINE.NS",
    "Havells India Ltd": "HAVELLS.NS",
    "Gujarat Alkalies and Chemicals Ltd": "GUJALKALI.NS",
    "Chemfab Alkalis Ltd": "CHEMFAB.NS",
    "Ajanta Pharma Ltd": "AJANTPHARM.NS",
    "Happiest Minds Technologies Ltd": "HAPPSTMNDS.NS",
    "Venus Pipes and Tubes Ltd": "VENUSPIPES.NS",
    "TALBROS Automotive Components Ltd": "TALBROAUTO.NS",
    "Zydus Lifesciences Ltd": "ZYDUSLIFE.NS",
    "Faze Three Ltd": "FAZE3Q.NS",
    "Automotive Stampings and Assemblies Ltd": "AUTOIND.NS",
    "Chennai Petroleum Corporation Ltd": "CHENNPETRO.NS",
    "Intellect Design Arena Ltd": "INTELLECT.NS",
    "Thejo Engineering Ltd": "THEJO.NS",
    "Insecticides (India) Ltd": "INSECTICID.NS",
    "Axiscades Technologies Ltd": "AXISCADES.NS",
    "Eimco Elecon (India) Ltd": "EIMCOELECO.NS",
    "Waaree Energies Ltd": "WAAREEENER.NS",
    "DCX Systems Ltd": "DCXINDIA.NS",
    "Ramkrishna Forgings Ltd": "RKFORGE.NS",
    "Gujarat Narmada Valley Fertilizers & Chemicals Ltd": "GNFC.NS",
    "JK Tyre & Industries Ltd": "JKTYRE.NS",
    "Oracle Financial Services Software Ltd": "OFSS.NS",
    "Aavas Financiers Ltd": "AAVAS.NS",
    "Bharat Bijlee Ltd": "BBL.NS",
    "Zydus Wellness Ltd": "ZYDUSWELL.NS",
    "Alivus Life Sciences Ltd": "ALIVUS.NS",
    "Hexaware Technologies Ltd": "HEXT.NS",
    "LIC Housing Finance Ltd": "LICHSGFIN.NS",
    "Tech Mahindra Ltd": "TECHM.NS",
    "Suprajit Engineering Ltd": "SUPRAJIT.NS",
    "IIFL Capital Services Ltd": "IIFLCAPS.NS",
    "Welspun Investments and Commercials Ltd": "WELINV.NS",
    "Lupin Ltd": "LUPIN.NS",
    "Paras Defence and Space Technologies Ltd": "PARAS.NS",
    "Unicommerce eSolutions Ltd": "UNIECOM.NS",
    "Isgec Heavy Engineering Ltd": "ISGEC.NS",
    "Apeejay Surrendra Park Hotels Ltd": "PARKHOTELS.NS",
    "Indo US Bio-Tech Ltd": "INDOUS.NS",
    "Moil Ltd": "MOIL.NS",
    "Hindustan Zinc Ltd": "HINDZINC.NS",
    "Kernex Microsystems (India) Ltd": "KERNEX.NS",
    "Camlin Fine Sciences Ltd": "CAMLINFINE.NS",
    "RHI Magnesita India Ltd": "RHIM.NS",
    "Anik Industries Ltd": "ANIKINDS.NS",
    "Time Technoplast Ltd": "TIMETECHNO.NS",
    "R R Kabel Ltd": "RRKABEL.NS",
    "Capacite Infraprojects Ltd": "CAPACITE.NS",
    "Indian Hume Pipe Company Ltd": "INDIANHUME.NS",
    "Pudumjee Paper Products Ltd": "PDMJEPAPER.NS",
    "Tarc Ltd": "TARC.NS",
    "Pokarna Ltd": "POKARNA.NS",
    "Brigade Enterprises Ltd": "BRIGADE.NS",
    "Info Edge (India) Ltd": "NAUKRI.NS",
    "Awfis Space Solutions Ltd": "AWFIS.NS",
    "Welspun Enterprises Ltd": "WELENT.NS",
    "Ems Ltd": "EMSLIMITED.NS",
    "Cohance Lifesciences Ltd": "COHANCE.NS",
    "Dynacons Systems and Solutions Ltd": "DSSL.NS",
    "AIA Engineering Ltd": "AIAENG.NS",
    "Pearl Global Industries Ltd": "PGIL.NS",
    "Hindustan Oil Exploration Company Ltd": "HINDOILEXP.NS",
    "Exide Industries Ltd": "EXIDEIND.NS",
    "Surya Roshni Ltd": "SURYAROSNI.NS",
    "Birla Corporation Ltd": "BIRLACORPN.NS",
    "Indo Count Industries Ltd": "ICIL.NS",
    "Atul Auto Ltd": "ATULAUTO.NS",
    "Crompton Greaves Consumer Electricals Ltd": "CROMPTON.NS",
    "Tata Motors Ltd": "TATAMOTORS.NS",
    "ACC Ltd": "ACC.NS",
    "Chambal Fertilisers and Chemicals Ltd": "CHAMBLFERT.NS",
    "Tejas Networks Ltd": "TEJASNET.NS",
    "Carborundum Universal Ltd": "CARBORUNIV.NS",
    "Kewal Kiran Clothing Ltd": "KKCL.NS",
    "Mangalore Refinery and Petrochemicals Ltd": "MRPL.NS",
    "Inox Wind Ltd": "INOXWIND.NS",
    "Max Estates Ltd": "MAXESTATES.NS",
    "Granules India Ltd": "GRANULES.NS",
    "Galaxy Surfactants Ltd": "GALAXYSURF.NS",
    "Indraprastha Gas Ltd": "IGL.NS",
    "BASF India Ltd": "BASF.NS",
    "Birlanu Ltd": "BIRLANU.NS",
    "Nitin Spinners Ltd": "NITINSPIN.NS",
    "TAJ GVK Hotels and Resorts Ltd": "TAJGVK.NS",
    "Pix Transmissions Ltd": "PIXTRANS.NS",
    "Trident Ltd": "TRIDENT.NS",
    "TVS Holdings Ltd": "TVSHLTD.NS",
    "Piramal Enterprises Ltd": "PEL.NS",
    "Motilal Oswal Nasdaq Q50 ETF": "MONQ50.NS",
    "Jbm Auto Ltd": "JBMA.NS",
    "Rane Brake Lining Ltd": "RBL.NS",
    "Gala Precision Engineering Ltd": "GALAPREC.NS",
    "Indoco Remedies Ltd": "INDOCO.NS",
    "Motilal Oswal Nifty Realty ETF": "MOREALTY.NS",
    "Gujarat Fluorochemicals Ltd": "FLUOROCHEM.NS",
    "Century Plyboards (India) Ltd": "CENTURYPLY.NS",
    "Westlife Foodworld Ltd": "WESTLIFE.NS",
    "Monarch Networth Capital Ltd": "MONARCH.NS",
    "JITF Infralogistics Ltd": "JITFINFRA.NS",
    "Rategain Travel Technologies Ltd": "RATEGAIN.NS",
    "Swan Energy Ltd": "SWANENERGY.NS",
    "Firstsource Solutions Ltd": "FSL.NS",
    "Sonata Software Ltd": "SONATSOFTW.NS",
    "Yasho Industries Ltd": "YASHO.NS",
    "Route Mobile Ltd": "ROUTE.NS",
    "Bata India Ltd": "BATAINDIA.NS",
    "Colgate-Palmolive (India) Ltd": "COLPAL.NS",
    "Refex Industries Ltd": "REFEX.NS",
    "Sona BLW Precision Forgings Ltd": "SONACOMS.NS",
    "Embassy Office Parks REIT": "EMBASSY.NS",
    "Birlasoft Ltd": "BSOFT.NS",
    "Ceigall India Ltd": "CEIGALL.NS",
    "Tata Consultancy Services Ltd": "TCS.NS",
    "Network People Services Technologies Ltd": "NPST.NS",
    "Mrs. Bectors Food Specialities Ltd": "BECTORFOOD.NS",
    "Voltamp Transformers Ltd": "VOLTAMP.NS",
    "Page Industries Ltd": "PAGEIND.NS",
    "ABB India Ltd": "ABB.NS",
    "AstraZeneca Pharma India Ltd": "ASTRAZEN.NS",
    "Wendt (India) Ltd": "WENDT.NS",
    "Procter & Gamble Hygiene & Health Care Ltd": "PGHH.NS",
    "Honeywell Automation India Ltd": "HONAUT.NS",
    "DISA India Ltd": "DISAQ.BO",
    "Orissa Minerals Development Company Ltd": "ORISSAMINE.NS",
    "GRP Ltd": "GRPLTD.NS",
    "Polyplex Corporation Ltd": "POLYPLEX.NS",
    "Ratnamani Metals & Tubes Ltd": "RATNAMANI.NS",
    "United Breweries Ltd": "UBL.NS",
    "Garware Hi-Tech Films Ltd": "GRWRHITECH.NS",
    "Deepak Nitrite Ltd": "DEEPAKNTR.NS",
    "Bajaj Electricals Ltd": "BAJAJELEC.NS",
    "Chemplast Sanmar Ltd": "CHEMPLASTS.NS",
    "Phoenix Mills Ltd": "PHOENIXLTD.NS",
    "Grindwell Norton Ltd": "GRINDWELL.NS",
    "KPIT Technologies Ltd": "KPITTECH.NS",
    "Syngene International Ltd": "SYNGENE.NS",
    "KNR Constructions Ltd": "KNRCON.NS",
    "Sundram Fasteners Ltd": "SUNDRMFAST.NS",
    "ZF Commercial Vehicle Control System India Ltd": "ZFCVINDIA.NS",
    "Rolex Rings Ltd": "ROLEXRINGS.NS",
    "Indo Tech Transformers Ltd": "INDOTECH.NS",
    "JSW Holdings Ltd": "JSWHL.NS",
    "Piramal Pharma Ltd": "PPLPHARMA.NS",
    "Crisil Ltd": "CRISIL.NS",
    "Varun Beverages Ltd": "VBL.NS",
    "Bajaj Holdings & Investment Ltd": "BAJAJHLDNG.NS",
    "RPG Life Sciences Ltd": "RPGLIFE.NS",
    "Bharat Rasayan Ltd": "BHARATRAS.NS",
    "Tata Elxsi Ltd": "TATAELXSI.NS",
    "Persistent Systems Ltd": "PERSISTENT.NS",
    "Trent Ltd": "TRENT.NS",
    "Sanofi Consumer Healthcare India Ltd": "SANOFICONR.NS",
    "Sundaram Finance Ltd": "SUNDARMFIN.NS",
    "Akzo Nobel India Ltd": "AKZOINDIA.NS",
    "Thermax Ltd": "THERMAX.NS",
    "GlaxoSmithKline Pharmaceuticals Ltd": "GLAXO.NS",
    "Mankind Pharma Ltd": "MANKIND.NS",
    "Mastek Ltd": "MASTEK.NS",
    "Angel One Ltd": "ANGELONE.NS",
    "Poly Medicure Ltd": "POLYMED.NS",
    "Interarch Building Solutions Ltd": "INTERARCH.NS",
    "Alkem Laboratories Ltd": "ALKEM.NS",
    "Narayana Hrudayalaya Ltd": "NH.NS",
    "Epigral Ltd": "EPIGRAL.NS",
    "Concord Biotech Ltd": "CONCORDBIO.NS",
    "D P Abhushan Ltd": "DPABHUSHAN.NS",
    "VA Tech Wabag Ltd": "WABAG.NS",
    "Balaji Amines Ltd": "BALAMINES.NS",
    "IPCA Laboratories Ltd": "IPCALAB.NS",
    "Websol Energy Systems Ltd": "WEBELSOLAR.NS",
    "Torrent Power Ltd": "TORNTPOWER.NS",
    "Aurionpro Solutions Ltd": "AURIONPRO.NS",
    "Godrej Consumer Products Ltd": "GODREJCP.NS",
    "Lodha Developers Ltd": "LODHA.NS",
    "Action Construction Equipment Ltd": "ACE.NS",
    "Ramco Cements Ltd": "RAMCOCEM.NS",
    "Associated Alcohols & Breweries Ltd": "ASALCBR.NS",
    "S.P. Apparels Ltd": "SPAL.NS",
    "Gokaldas Exports Ltd": "GOKEX.NS",
    "Sumitomo Chemical India Ltd": "SUMICHEM.NS",
    "Berger Paints India Ltd": "BERGEPAINT.NS",
    "Triveni Turbine Ltd": "TRITURBINE.NS",
    "Transformers and Rectifiers (India) Ltd": "TARIL.NS",
    "MSTC Ltd": "MSTCLTD.NS",
    "Jash Engineering Ltd": "JASH.NS",
    "Kalyan Jewellers India Ltd": "KALYANKJIL.NS",
    "HPL Electric & Power Ltd": "HPL.NS",
    "Epack Durable Ltd": "EPACK.NS",
    "Panama Petrochem Ltd": "PANAMAPET.NS",
    "Jindal SAW Ltd": "JINDALSAW.NS",
    "Gail (India) Ltd": "GAIL.NS",
    "Cera Sanitaryware Ltd": "CERA.NS",
    "Bayer Cropscience Ltd": "BAYERCROP.NS",
    "TCPL Packaging Ltd": "TCPLPACK.NS",
    "Power Mech Projects Ltd": "POWERMECH.NS",
    "Summit Securities Ltd": "SUMMITSEC.NS",
    "Anup Engineering Ltd": "ANUP.NS",
    "Balkrishna Industries Ltd": "BALKRISIND.NS",
    "MPS Ltd": "MPSLTD.NS",
    "Bajaj Finserv Ltd": "BAJAJFINSV.NS",
    "ICICI Lombard General Insurance Co Ltd": "ICICIGI.NS",
    "Blue Star Ltd": "BLUESTARCO.NS",
    "Oberoi Realty Ltd": "OBEROIRLTY.NS",
    "Mallcom (India) Ltd": "MALLCOM.NS",
    "Max Healthcare Institute Ltd": "MAXHEALTH.NS",
    "Amara Raja Energy & Mobility Ltd": "ARE&M.NS",
    "Dabur India Ltd": "DABUR.NS",
    "BLS International Services Ltd": "BLS.NS",
    "Petronet LNG Ltd": "PETRONET.NS",
    "Indian Railway Finance Corp Ltd": "IRFC.NS",
    "Power Grid Corporation of India Ltd": "POWERGRID.NS",
    "Pidilite Industries Ltd": "PIDILITIND.NS",
    "Federal Bank Ltd": "FEDERALBNK.NS"

}

# ---------------------------------
# Master Loop (Yahoo → NSE fallback)
# ---------------------------------
results = {}
for name, symbol in stocks.items():
    print(f"\nFetching {name} ({symbol})...")
    res = None

    if symbol.endswith(".NS"):  # Yahoo first
        res = get_returns_yahoo(symbol)
        time.sleep(0.5)

    if not res:  # NSE fallback
        nse_symbol = symbol.replace(".NS", "")
        price = get_price_from_nse(nse_symbol)
        if price:
            res = {
                "Current Price (₹)": price,
                "1D %": None, "1W %": None,
                "2W %": None, "1M %": None,
                "3M %": None, "6M %": None,
                "YTD %": None,
                "Last Updated": datetime.today().strftime("%d-%m-%Y")
            }

    if not res:
        res = {"Error": "Data not found"}

    results[name] = {"Symbol": symbol.replace(".NS", ""), **res}

# ---------------------------------
# Export to Excel
# ---------------------------------
excelName = "Stock-List_" + datetime.now().strftime("%d-%m-%Y") + ".xlsx"
df = pd.DataFrame(results).T
df.to_excel(excelName)
print(f"\n {excelName} created successfully!")

# ---------------------------------
# Apply Conditional Formatting to 'Trend' Column
# ---------------------------------
wb = load_workbook(excelName)
ws = wb.active

# Find 'Trend' column
trend_col_idx = None
for idx, cell in enumerate(ws[1], start=1):
    if cell.value == "Trend":
        trend_col_idx = idx
        break

if trend_col_idx:
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=trend_col_idx)
        if cell.value == "Bullish":
            cell.fill = green_fill
        elif cell.value == "Bearish":
            cell.fill = red_fill
        # Neutral/Unknown are left unchanged

wb.save(excelName)
print(f"\n✅ Trend colors applied in {excelName}")
