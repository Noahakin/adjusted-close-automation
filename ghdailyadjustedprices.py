import yfinance as yf
import pandas as pd
from datetime import datetime
import os
import smtplib
from email.message import EmailMessage

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

# Email config (use GitHub Secrets)
EMAIL_SENDER = "nakin@lsfunds.com"
EMAIL_RECEIVER = "nakin@lsfunds.com"
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]  # stored securely in GitHub

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

# -----------------------
# EMAIL FILE
# -----------------------
msg = EmailMessage()
msg["Subject"] = f"Adjusted Closing Prices – {date_str}"
msg["From"] = EMAIL_SENDER
msg["To"] = EMAIL_RECEIVER

msg.set_content(
    f"Attached are the adjusted closing prices for the prior trading day ({date_str})."
)

with open(file_path, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(file_path)
    )

with smtplib.SMTP("smtp.office365.com", 587) as smtp:
    smtp.starttls()
    smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
    smtp.send_message(msg)


print(f"Email sent with attachment: {file_path}")
