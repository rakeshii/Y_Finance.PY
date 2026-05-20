import os
import requests
import pandas as pd
import yfinance as yf
import streamlit as st
from datetime import datetime, date, timedelta

# ================= CONFIG =================
EXPORT_DIR = "exports"
os.makedirs(EXPORT_DIR, exist_ok=True)


# ================= UTILITIES =================
def safe_df(df):
    return df if isinstance(df, pd.DataFrame) else pd.DataFrame()


def remove_timezone(df):

    if df.empty:
        return df

    # Remove timezone from index
    if hasattr(df.index, 'tz') and df.index.tz is not None:
        df.index = df.index.tz_localize(None)

    # Remove timezone from datetime columns
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.tz_localize(None)

    return df


# ================= YAHOO API =================
@st.cache_data(ttl=3600)
def fetch_yahoo_statistics(ticker):

    url = f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{ticker}"

    modules = [
        "defaultKeyStatistics",
        "financialData",
        "summaryDetail",
        "calendarEvents"
    ]

    params = {"modules": ",".join(modules)}

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json"
    }

    try:
        r = requests.get(url, headers=headers, params=params, timeout=10)
        data = r.json()

        if "quoteSummary" not in data:
            return pd.DataFrame()

        if not data["quoteSummary"]["result"]:
            return pd.DataFrame()

        result = data["quoteSummary"]["result"][0]

        rows = []

        for module in result:

            for key, value in result[module].items():

                if isinstance(value, dict):
                    val = value.get("fmt") or value.get("raw")
                else:
                    val = value

                rows.append({
                    "Metric": key,
                    "Value": val,
                    "Category": module
                })

        return pd.DataFrame(rows)

    except Exception as e:
        st.error(f"Yahoo API error: {e}")
        return pd.DataFrame()


# ================= EXPORT FUNCTION =================
def fetch_and_export(ticker, start_date, end_date):

    stock = yf.Ticker(ticker)

    # Historical Data using Calendar Dates
    history = safe_df(
        stock.history(
            start=start_date,
            end=end_date
        )
    )

    balance = safe_df(stock.balance_sheet)
    income = safe_df(stock.financials)
    cashflow = safe_df(stock.cashflow)

    stats = fetch_yahoo_statistics(ticker)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    filename = f"{ticker}_YahooFinance_{timestamp}.xlsx"
    filepath = os.path.join(EXPORT_DIR, filename)

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:

        # Price History
        if not history.empty:
            history = history.sort_index(ascending=False)
            history = remove_timezone(history)
            history.to_excel(writer, sheet_name="Price_History")

        # Balance Sheet
        if not balance.empty:
            balance = remove_timezone(balance)
            balance.to_excel(writer, sheet_name="Balance_Sheet")

        # Income Statement
        if not income.empty:
            income = remove_timezone(income)
            income.to_excel(writer, sheet_name="Income_Statement")

        # Cashflow
        if not cashflow.empty:
            cashflow = remove_timezone(cashflow)
            cashflow.to_excel(writer, sheet_name="Cashflow")

        # Statistics
        if not stats.empty:
            stats = remove_timezone(stats)
            stats.to_excel(writer, sheet_name="Key_Statistics", index=False)

    return filepath


# ================= STREAMLIT UI =================
st.set_page_config(
    page_title="Yahoo Finance Data Exporter",
    layout="centered"
)

st.title("📊 Yahoo Finance Financial Data Exporter")

st.write("Fetch financial statements and historical data from Yahoo Finance.")

# -------- Ticker Input --------
ticker = st.text_input(
    "Enter Company Ticker",
    placeholder="Example: RELIANCE.NS or AAPL"
)

# -------- Date Selection --------
st.subheader("📅 Select Historical Data Duration")

default_end = date.today()
default_start = default_end - timedelta(days=365)

col1, col2 = st.columns(2)

with col1:
    start_date = st.date_input(
        "Start Date",
        value=default_start
    )

with col2:
    end_date = st.date_input(
        "End Date",
        value=default_end
    )

# -------- Fetch Button --------
if st.button("Fetch & Export Data"):

    if ticker:

        if start_date >= end_date:
            st.error("Start Date must be before End Date.")

        else:

            with st.spinner("Fetching data from Yahoo Finance..."):

                try:

                    filepath = fetch_and_export(
                        ticker.upper(),
                        start_date,
                        end_date
                    )

                    st.success("Data exported successfully!")

                    with open(filepath, "rb") as f:

                        st.download_button(
                            label="📥 Download Excel File",
                            data=f,
                            file_name=os.path.basename(filepath),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                except Exception as e:
                    st.error(f"Error: {e}")

    else:
        st.warning("Please enter a valid ticker.")