import os, re, time
from datetime import datetime

import pandas as pd
import yfinance as yf

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule

# ========= SETTINGS =========
MASTER_XLSX = "SKV Sheet_Updated PM.xlsx"    # client edits this
OUTPUT_CSV  = "SKV Sheet-1-Updated.csv"      # for Power Query / web
THRESHOLD   = 0.025                           # 2.5%

# ========= VALIDATE FILE =========
if not os.path.exists(MASTER_XLSX):
    raise FileNotFoundError(f"{MASTER_XLSX} not found in repo!")

# ========= LOAD =========
df = pd.read_excel(MASTER_XLSX, engine="openpyxl")

# Ensure required columns exist
required_cols = [
    "Stock Name","Time Frame","Zone","Entry Price","Stop Loss","Legout Date",
    "Validation Issue","Zone Perfection","Entry Date","Status","Diff","Qty",
    "Tgt","Yahoo Symbol","Last Close Price"
]
for col in required_cols:
    if col not in df.columns:
        df[col] = None

# ========= CLEAN SYMBOLS =========
def clean_symbol(sym):
    if not isinstance(sym, str) or sym.strip() == "":
        return None
    s = sym.strip().upper()
    s = re.sub(r"^\$+", "", s)
    s = s.replace("_", "-")
    s = re.sub(r"[^A-Z0-9\-\.]", "", s)
    if not s.endswith(".NS"):
        s += ".NS"
    return s

# If Yahoo Symbol missing, build from Stock Name
df["Yahoo Symbol"] = df["Yahoo Symbol"].where(df["Yahoo Symbol"].notna(), df["Stock Name"])
df["Yahoo Symbol"] = df["Yahoo Symbol"].apply(clean_symbol)

# ========= FETCH PRICES =========
new_prices = {}
failed = []

for symbol in df["Yahoo Symbol"]:
    if pd.isna(symbol) or not isinstance(symbol, str) or symbol.strip()=="":
        failed.append(symbol)
        continue
    try:
        t = yf.Ticker(symbol)
        hist = t.history(period="1d")
        if hist.empty:
            hist = t.history(period="5d")
        if not hist.empty:
            price = float(hist["Close"].dropna().iloc[-1])
            new_prices[symbol] = round(price, 2)
        else:
            failed.append(symbol)
    except Exception:
        failed.append(symbol)
    time.sleep(0.3)  # be gentle with Yahoo

def set_price(row):
    sym = row.get("Yahoo Symbol")
    if isinstance(sym, str) and sym in new_prices:
        return new_prices[sym]
    if isinstance(sym, str) and sym in failed:
        return "Not available"
    # keep existing if neither fetched nor failed
    return row.get("Last Close Price")

df["Last Close Price"] = df.apply(set_price, axis=1)

# ========= DIFF % helper (numeric rows only) =========
def diff_pct(row):
    entry = row.get("Entry Price")
    last  = row.get("Last Close Price")
    try:
        entry = float(entry)
        last  = float(last)
        if entry and entry != 0:
            return round((last - entry) / entry * 100, 2)
    except Exception:
        return None
    return None

df["Diff %"] = df.apply(diff_pct, axis=1)  # optional helper column

# ========= SAVE: Excel + CSV =========
df.to_excel(MASTER_XLSX, index=False)
df.to_csv(OUTPUT_CSV, index=False)

# ========= ADD CONDITIONAL FORMATTING (green/red rows) =========
# We re-open with openpyxl to apply formatting that persists in Excel
wb = load_workbook(MASTER_XLSX)
ws = wb.active

# Find column letters by header names (robust to column order)
from openpyxl.utils.cell import get_column_letter

headers = {cell.value: cell.column for cell in ws[1] if cell.value}
def col_letter(col_name):
    idx = headers.get(col_name)
    return get_column_letter(idx) if idx else None

col_entry = col_letter("Entry Price")
col_last  = col_letter("Last Close Price")

if col_entry and col_last:
    start_row = 2
    end_row = ws.max_row
    end_col_letter = get_column_letter(ws.max_column)
    full_range = f"A{start_row}:{end_col_letter}{end_row}"

    # formulas use absolute columns with relative row
    green_formula = f"=AND(ISNUMBER(${col_last}{start_row}),ISNUMBER(${col_entry}{start_row}),${col_last}{start_row}>={col_entry}{start_row}*(1+{THRESHOLD}))"
    red_formula   = f"=AND(ISNUMBER(${col_last}{start_row}),ISNUMBER(${col_entry}{start_row}),${col_last}{start_row}<={col_entry}{start_row}*(1-{THRESHOLD}))"

    green_rule = FormulaRule(formula=[green_formula],
                             fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                             stopIfTrue=True)
    red_rule   = FormulaRule(formula=[red_formula],
                             fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
                             stopIfTrue=True)

    ws.conditional_formatting.add(full_range, green_rule)
    ws.conditional_formatting.add(full_range, red_rule)

wb.save(MASTER_XLSX)

# ========= LOG =========
print(f"✅ Prices updated at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print(f"📂 Excel updated: {MASTER_XLSX}")
print(f"📂 CSV updated:   {OUTPUT_CSV}")

if failed:
    print("\n⚠️ Not available:")
    for s in sorted(set([x for x in failed if x])):
        print(" -", s)
