import yfinance as yf
import pandas as pd
from datetime import datetime
import os

# Ensure script runs from its own directory
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# -----------------------
# CONFIGURATION
# -----------------------
TICKERS = [
    "OVL", "VOO", "OVS", "IJR", "OVF", "IEFA", "IEMG", "OVB",
    "AGG", "OVM", "MUB", "OVT", "VCSH", "OVLH", "KHPI",
    "JEPI", "SPY", "QQQ"
]

SAVE_FOLDER = r"C:\Users\nakin\OneDrive - lsfunds.com\Adjusted Close Tickers"

# -----------------------

# Ensure folder exists
os.makedirs(SAVE_FOLDER, exist_ok=True)

# ----------------------------------------
# DOWNLOAD MULTIPLE DAYS (IMPORTANT CHANGE)
# ----------------------------------------
data = yf.download(
    tickers=TICKERS,
    period="7d",          # enough to cover weekends/holidays
    interval="1d",
    auto_adjust=False,
    group_by="ticker"
)

# ----------------------------------------
# DETERMINE PRIOR TRADING DAY (KEY LOGIC)
# ----------------------------------------
# All tickers share the same date index
available_dates = data[TICKERS[0]].index

today = pd.Timestamp(datetime.now().date())
prior_trading_day = available_dates[available_dates < today].max()

# ----------------------------------------
# BUILD OUTPUT ROWS
# ----------------------------------------
rows = []

for ticker in TICKERS:
    try:
        adj_close = data[ticker].loc[prior_trading_day, "Adj Close"]
        rows.append({
            "Date": prior_trading_day.strftime("%Y-%m-%d"),
            "Ticker": ticker,
            "Adjusted Close": round(float(adj_close), 4)
        })
    except Exception:
        rows.append({
            "Date": prior_trading_day.strftime("%Y-%m-%d"),
            "Ticker": ticker,
            "Adjusted Close": None
        })

df = pd.DataFrame(rows)

# ----------------------------------------
# SAVE FILE (DATED BY PRIOR TRADING DAY)
# ----------------------------------------
date_str = prior_trading_day.strftime("%Y-%m-%d")
file_path = os.path.join(
    SAVE_FOLDER,
    f"Adjusted.Close_{date_str}.xlsx"
)

df.to_excel(file_path, index=False)

# ----------------------------------------
# LOG SUCCESS
# ----------------------------------------
with open(os.path.join(SAVE_FOLDER, "task_log.txt"), "a") as f:
    f.write(f"{datetime.now()} - Saved adjusted closes for {date_str}\n")

print(f"Saved file to: {file_path}")
