import pandas as pd
import yfinance as yf
import re
import time
from datetime import datetime
import os

# === Master file ===
updated_file = "SKV Sheet_Updated PM.xlsx"  # Client edits here
output_csv = "SKV Sheet-1-Updated.csv"      # Output for Power Query

if not os.path.exists(updated_file):
    raise FileNotFoundError(f"{updated_file} not found in repo!")

# === Load Updated file (.xlsx uses openpyxl) ===
df = pd.read_excel(updated_file, engine="openpyxl")

# === Clean Yahoo Stock Symbols ===
def clean_symbol(sym):
    if not isinstance(sym, str) or sym.strip() == "":
        return None
    sym = sym.strip().upper()
    sym = re.sub(r"^\$+", "", sym)   # remove leading $
    sym = sym.replace("_", "-")      # replace _ with -
    sym = re.sub(r"[^A-Z0-9\-]", "", sym)  # allow only letters, digits, dash
    if not sym.endswith(".NS"):
        sym += ".NS"
    return sym

df["Yahoo Symbol"] = df["Stock Name"].apply(clean_symbol)

# === Fetch Latest Prices ===
new_prices = {}
failed_symbols = []

for symbol in df["Yahoo Symbol"]:
    if pd.isna(symbol):
        failed_symbols.append(symbol)
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

    time.sleep(0.3)  # avoid Yahoo rate limits

# === Update Prices Column ===
def update_price(row):
    sym = row["Yahoo Symbol"]
    if sym in new_prices:
        return new_prices[sym]
    elif sym in failed_symbols:
        return "Not available"
    else:
        return row.get("Last Close Price")

df["Last Close Price"] = df.apply(update_price, axis=1)

# === Save Updated File ===
df.to_excel(updated_file, index=False)  # overwrite same file
df.to_csv(output_csv, index=False)

print(f"✅ Prices updated at {datetime.now()}")
print(f"📂 Excel updated: {updated_file}")
print(f"📂 CSV updated: {output_csv}")

if failed_symbols:
    print("\n⚠️ These symbols could not be fetched:")
    for sym in sorted(set(filter(None, failed_symbols))):
        print(" -", sym)
