import os, re, time, subprocess
from datetime import datetime

import pandas as pd
import yfinance as yf

# ========= SETTINGS =========
MASTER_CSV = "SKV Sheet_Updated PM.csv"      # Input CSV file
OUTPUT_CSV = "SKV Sheet-1-Updated.csv"       # Output CSV file
THRESHOLD  = 0.025                           # 2.5% for green/red highlights
SLEEP_SEC  = 0.3                             # polite delay for Yahoo
AUTO_PUSH  = False                           # Git push (disable if not using Git)

# ========= HELPERS =========
def log(msg): print(msg, flush=True)

def run(cmd, cwd=None):
    try:
        p = subprocess.Popen(
            cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True
        )
        out, err = p.communicate()
        return p.returncode, out.decode(errors="ignore"), err.decode(errors="ignore")
    except Exception as e:
        return 1, "", str(e)

def git_commit_push(files, message):
    if not AUTO_PUSH:
        log("ℹ️ AUTO_PUSH=False — skipping git push.")
        return

    run("git add " + " ".join(f'"{f}"' for f in files))
    rc, out, err = run(f'git commit -m "{message}"')
    if rc != 0 and "nothing to commit" in (out + err).lower():
        log("ℹ️ No changes to commit.")
    else:
        log(out or err)

    rc, out, err = run("git push")
    if rc != 0:
        log(f"⚠️ git push failed:\n{err or out}")
    else:
        log("🚀 Pushed to remote.")

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

# ========= VALIDATE FILE =========
if not os.path.exists(MASTER_CSV):
    raise FileNotFoundError(f"{MASTER_CSV} not found in directory!")

# ========= LOAD CSV =========
df = pd.read_csv(MASTER_CSV)

# Add missing expected columns if they don't exist
for col in ["Yahoo Symbol", "Last Close Price", "Diff %", "Diff", "Tgt"]:
    if col not in df.columns:
        df[col] = None

# Clean Yahoo symbols
df["Yahoo Symbol"] = df["Yahoo Symbol"].where(df["Yahoo Symbol"].notna(), df["Stock Name"])
df["Yahoo Symbol"] = df["Yahoo Symbol"].apply(clean_symbol)

# ========= FETCH LATEST PRICES =========
new_prices = {}
failed = []

symbols = df["Yahoo Symbol"]
for symbol in symbols:
    if pd.isna(symbol):
        failed.append(symbol)
        continue
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period="1d")
        if hist.empty:
            hist = ticker.history(period="5d")
        if not hist.empty:
            price = round(float(hist["Close"].dropna().iloc[-1]), 2)
            new_prices[symbol] = price
        else:
            failed.append(symbol)
    except Exception:
        failed.append(symbol)
    time.sleep(SLEEP_SEC)

def set_price(row):
    sym = row.get("Yahoo Symbol")
    if isinstance(sym, str) and sym in new_prices:
        return new_prices[sym]
    if isinstance(sym, str) and sym in failed:
        return "Not available"
    return row.get("Last Close Price")

df["Last Close Price"] = df.apply(set_price, axis=1)

# ========= CALCULATE Diff, Tgt, Diff % =========
df["Diff"] = df.apply(
    lambda row: row["Entry Price"] - row["Stop Loss"]
    if pd.notna(row["Entry Price"]) and pd.notna(row["Stop Loss"]) else None,
    axis=1
)

df["Tgt"] = df["Diff"].apply(lambda x: round(x * 5, 2) if pd.notna(x) else None)

def diff_pct(row):
    try:
        entry = float(row.get("Entry Price"))
        last = float(row.get("Last Close Price"))
        if entry != 0:
            return round((last - entry) / entry * 100, 2)
    except:
        return None
    return None

df["Diff %"] = df.apply(diff_pct, axis=1)

# ========= SAVE CSV =========
df.to_csv(OUTPUT_CSV, index=False)
log(f"✅ CSV updated: {OUTPUT_CSV}")

# ========= GIT PUSH =========
git_commit_push([OUTPUT_CSV], f"Auto update @ {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# ========= SUMMARY =========
if failed:
    log("\n⚠️ Failed to fetch prices for:")
    for s in sorted(set([x for x in failed if x])):
        log(f" - {s}")
