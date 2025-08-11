import pandas as pd
import yfinance as yf
import re
import time
from datetime import datetime

# === Load Excel ===
file_path = "SKV Sheet-1.xlsx"  # Original master file in repo
df = pd.read_excel(file_path)

# === Clean Yahoo Stock Symbols ===
def clean_symbol(sym):
    if not isinstance(sym, str):
        return None
    sym = sym.strip().upper()
    sym = re.sub(r"^\$+", "", sym)
    sym = sym.replace("_", "-")
    sym = re.sub(r"[^A-Z0-9\-]", "", sym)
    return sym + ".NS"

df["Yahoo Symbol"] = df["Stock Name"].apply(clean_symbol)

# === Rename 'Last Close Price' to 'Previous Price' if exists ===
if "Last Close Price" in df.columns:
    df.rename(columns={"Last Close Price": "Previous Price"}, inplace=True)
else:
    df["Previous Price"] = None

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

# === Calculate % Change ===
def calc_change_percent(new, old):
    if pd.isna(new) or pd.isna(old) or old == 0:
        return None
    return round(((new - old) / old) * 100, 2)

df["% Change"] = df.apply(lambda row: calc_change_percent(row["Last Close Price"], row["Previous Price"]), axis=1)

# === Flag Rows ===
df["Flag"] = df["% Change"].apply(lambda x: "Highlight" if isinstance(x, float) and abs(x) > 2.5 else "")

# === Save Updated CSV (overwrite existing updated file) ===
output_file = "SKV Sheet-1-Updated.csv"
df.to_csv(output_file, index=False)

print(f"✅ CSV updated at {datetime.now()}")

if failed_symbols:
    print("\n⚠️ Failed to fetch:")
    for sym in sorted(set(failed_symbols)):
        print(" -", sym)
