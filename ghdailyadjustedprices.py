import yfinance as yf
import pandas as pd
from datetime import datetime
import os
import requests

# -----------------------
# CONFIGURATION
# -----------------------
TICKERS = [
    "OVL", "VOO", "OVS", "IJR", "OVF", "IEFA", "IEMG", "OVB",
    "AGG", "OVM", "MUB", "OVT", "VCSH", "OVLH", "KHPI",
    "JEPI", "SPY", "QQQ"
]

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

ONEDRIVE_FOLDER = "/Adjusted Close Tickers"

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

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
file_name = f"Adjusted_Close_{date_str}.xlsx"
file_path = os.path.join(OUTPUT_DIR, file_name)

df.to_excel(file_path, index=False)
print(f"File created: {file_path}")

# -----------------------
# AUTHENTICATE TO GRAPH
# -----------------------
token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

token_data = {
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "scope": "https://graph.microsoft.com/.default",
    "grant_type": "client_credentials",
}

token_response = requests.post(token_url, data=token_data)

print("Token status code:", token_response.status_code)
print("Token response:", token_response.text)

token_response.raise_for_status()

token_json = token_response.json()

if "access_token" not in token_json:
    raise RuntimeError(f"No access_token returned: {token_json}")

access_token = token_json["access_token"]


# -----------------------
# UPLOAD TO ONEDRIVE
# -----------------------
upload_url = (
    "https://graph.microsoft.com/v1.0/me/drive/root:"
    f"{ONEDRIVE_FOLDER}/{file_name}:/content"
)

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/octet-stream",
}

with open(file_path, "rb") as f:
    upload_response = requests.put(upload_url, headers=headers, data=f)

upload_response.raise_for_status()
print("File uploaded to OneDrive successfully.")

