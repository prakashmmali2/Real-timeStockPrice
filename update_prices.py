import pandas as pd
import yfinance as yf
import re
import time
from datetime import datetime

# === Load Excel ===
file_path = "SKV Sheet-1.xlsx"  # Uploaded file name
df = pd.read_excel(file_path)

# === Clean Yahoo Stock Symbols ===
def clean_symbol(sym):
    if not isinstance(sym, str):
        return None
    sym = sym.strip().upper()
    sym = re.sub(r"^\$+", "", sym)
    sym = sym.replace("_", "-")
    sym = re.sub(r"[^A-Z0-9\-]", "", sym)
    return sym + ".NS"  # NSE format

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
    time.sleep(0.3)

df["Last Close Price"] = new_prices

# === Save to CSV (no Change % column) ===
output_file = "SKV Sheet-1-Updated.csv"
df.to_csv(output_file, index=False)

print(f"✅ CSV saved at {datetime.now()} — open in Excel and apply filter on 'Last Close Price' column")

if failed_symbols:
    print("\n⚠️ Failed to fetch:")
    for sym in sorted(set(failed_symbols)):
        print(" -", sym)

# Prevent GitHub Actions JSON-to-Python errors
execution_count = None 
