import pandas as pd
import yfinance as yf
import re
import time
from datetime import datetime
import os

# === File Paths ===
excel_file = "SKV Sheet-1.xlsx"       # Client editable file
output_csv = "SKV Sheet-1-Updated.csv"  # For Power Query

# === Load latest Excel from repo ===
if not os.path.exists(excel_file):
    raise FileNotFoundError(f"{excel_file} not found in repo!")

df = pd.read_excel(excel_file)

# === Clean Yahoo Stock Symbols ===
def clean_symbol(sym):
    if not isinstance(sym, str) or sym.strip() == "":
        return None
    sym = sym.strip().upper()
    sym = re.sub(r"^\$+", "", sym)
    sym = sym.replace("_", "-")
    sym = re.sub(r"[^A-Z0-9\-]", "", sym)
    if not sym.endswith(".NS"):
        sym += ".NS"
    return sym

# Create/Update Yahoo Symbol column
df["Yahoo Symbol"] = df["Stock Name"].apply(clean_symbol)

# === Fetch Updated Prices ===
new_prices = {}
failed_symbols = []

for symbol in df["Yahoo Symbol"]:
    if pd.isna(symbol):
        continue
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period="1d")
        if hist.empty:
            hist = ticker.history(period="5d")
        if not hist.empty:
            last_close = round(hist["Close"].dropna().iloc[-1], 2)
            new_prices[symbol] = last_close
        else:
            failed_symbols.append(symbol)
    except Exception:
        failed_symbols.append(symbol)
    time.sleep(0.3)  # avoid Yahoo Finance rate limits

# === Update only prices in existing rows ===
df["Last Close Price"] = df.apply(
    lambda row: new_prices.get(row["Yahoo Symbol"], row.get("Last Close Price")),
    axis=1
)

# === Save updated data back ===
df.to_excel(excel_file, index=False)
df.to_csv(output_csv, index=False)

print(f"✅ Prices updated at {datetime.now()}")
print(f"📂 Excel file updated: {excel_file}")
print(f"📂 CSV file updated: {output_csv}")

if failed_symbols:
    print("\n⚠️ Failed to fetch:")
    for sym in sorted(set(failed_symbols)):
        print(" -", sym)

execution_count = None
