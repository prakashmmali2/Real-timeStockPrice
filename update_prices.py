import pandas as pd
import yfinance as yf
import re
import time
from datetime import datetime

# === File Paths ===
excel_file = "SKV Sheet-1.xlsx"       # Original client-editable file
output_csv = "SKV Sheet-1-Updated.csv"  # CSV for Power Query

# === Load Excel ===
df = pd.read_excel(excel_file)

# === Clean Yahoo Stock Symbols ===
def clean_symbol(sym):
    if not isinstance(sym, str) or sym.strip() == "":
        return None
    sym = sym.strip().upper()
    sym = re.sub(r"^\$+", "", sym)
    sym = sym.replace("_", "-")
    sym = re.sub(r"[^A-Z0-9\-]", "", sym)
    return sym + ".NS"  # NSE format

# Create symbol column if not exists
if "Yahoo Symbol" not in df.columns:
    df["Yahoo Symbol"] = df["Stock Name"].apply(clean_symbol)
else:
    df["Yahoo Symbol"] = df["Stock Name"].apply(clean_symbol)

# === Fetch Updated Prices ===
new_prices = []
failed_symbols = []

for symbol in df["Yahoo Symbol"]:
    if pd.isna(symbol):
        new_prices.append(None)
        continue
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period="1d")
        if hist.empty:
            hist = ticker.history(period="5d")
        if not hist.empty:
            last_close = round(hist["Close"].dropna().iloc[-1], 2)
            new_prices.append(last_close)
        else:
            new_prices.append(None)
            failed_symbols.append(symbol)
    except Exception:
        new_prices.append(None)
        failed_symbols.append(symbol)
    time.sleep(0.3)  # avoid rate limits

# Update only the "Last Close Price" column
df["Last Close Price"] = new_prices

# === Save back to Excel ===
df.to_excel(excel_file, index=False)

# === Save to CSV for Power Query ===
df.to_csv(output_csv, index=False)

print(f"✅ Prices updated at {datetime.now()}")
print(f"📂 Excel file updated: {excel_file}")
print(f"📂 CSV file updated: {output_csv}")

if failed_symbols:
    print("\n⚠️ Failed to fetch:")
    for sym in sorted(set(failed_symbols)):
        print(" -", sym)

# Prevent GitHub Actions JSON-to-Python errors
execution_count = None
