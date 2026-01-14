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

ONEDRIVE_USER = "nakin@lsfunds.com"

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

ONEDRIVE_FOLDER = "/Adjusted Close Tickers"

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

# -----------------------
# DOWNLOAD 5 YEARS OF DATA
# -----------------------
data = yf.download(
    tickers=TICKERS,
    period="5y",
    interval="1d",
    auto_adjust=False,
    group_by="ticker",
    threads=True
)

# -----------------------
# EXTRACT ADJUSTED CLOSE (WIDE FORMAT)
# -----------------------
adj_close = pd.DataFrame()

for ticker in TICKERS:
    if ticker in data:
        adj_close[ticker] = data[ticker]["Adj Close"]

# Ensure Date index is clean
adj_close.index = adj_close.index.strftime("%Y-%m-%d")
adj_close.index.name = "Date"

# -----------------------
# SAVE FILE
# -----------------------
today_str = datetime.utcnow().strftime("%Y-%m-%d")
file_name = f"Adjust_Close_5Y_{today_str}.xlsx"
file_path = os.path.join(OUTPUT_DIR, file_name)

adj_close.to_excel(file_path)
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
token_response.raise_for_status()

access_token = token_response.json()["access_token"]

# -----------------------
# UPLOAD TO ONEDRIVE
# -----------------------
upload_url = (
    f"https://graph.microsoft.com/v1.0/users/{ONEDRIVE_USER}"
    f"/drive/root:{ONEDRIVE_FOLDER}/{file_name}:/content"
)

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/octet-stream",
}

with open(file_path, "rb") as f:
    upload_response = requests.put(upload_url, headers=headers, data=f)

upload_response.raise_for_status()
print("File uploaded to OneDrive successfully.")

