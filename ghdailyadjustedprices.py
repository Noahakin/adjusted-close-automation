import yfinance as yf
import pandas as pd
from datetime import datetime
import os

# -----------------------
# CONFIGURATION
# -----------------------
TICKERS = [
    "OVL", "VOO", "OVS", "IJR", "OVF", "IEFA", "IEMG", "OVB",
    "AGG", "OVM", "MUB", "OVT", "VCSH", "OVLH", "KHPI",
    "JEPI", "SPY", "QQQ"
]

OUTPUT_DIR = "output"   # GitHub runner local directory
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -----------------------
# DOWNLOAD DATA
# -----------------------
data = yf.download(
    tickers=TICKERS,
    period="7d",
    interval="1d",
    auto_adjust=False,
    group_by="ticker"
)

# -----------------------
# PRIOR TRADING DAY
# -----------------------
available_dates = data[TICKERS[0]].index
today = pd.Timestamp(datetime.utcnow().date())
prior_trading_day = available_dates[available_dates < today].max()

# -----------------------
# BUILD DATAFRAME
# -----------------------
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

# -----------------------
# SAVE FILE
# -----------------------
date_str = prior_trading_day.strftime("%Y-%m-%d")
file_path = os.path.join(OUTPUT_DIR, f"Adjusted_Close_{date_str}.xlsx")
df.to_excel(file_path, index=False)

print(f"File successfully created: {file_path}")
