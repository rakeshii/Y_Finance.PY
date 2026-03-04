# 📊 FinScraper — Yahoo Finance Exporter

A sleek **Streamlit** web app that lets you fetch, visualize, and export Yahoo Finance data to beautifully formatted **Excel files** — in seconds.

---

## ✨ Features

| Feature | Details |
|---|---|
| 📥 **Excel Export** | Price History, Balance Sheet, Income Statement, Cashflow, Key Statistics, Analysis |
| 📈 **Interactive Charts** | 1-year price & volume, Revenue & Profit, Balance Sheet — powered by Plotly |
| 🌍 **Global Tickers** | US stocks, Indian NSE (`.NS`) & BSE (`.BO`), and all Yahoo Finance tickers |
| ⚡ **Batch Processing** | Enter multiple tickers at once and export them all |

---

## 🚀 Quick Start

### 1. Clone the repo
```bash
git clone https://github.com/<your-username>/Y_Finance.PY.git
cd Y_Finance.PY
```

### 2. Create a virtual environment (recommended)
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the app
```bash
streamlit run Yahoo_Scrapper_Strmlt.py
```

The app will open at **http://localhost:8501** in your browser.

---

## 🖥️ Usage

1. Enter one or more ticker symbols in the sidebar (e.g. `AAPL`, `MSFT`, `RELIANCE.NS`)
2. Toggle **Include Analysis Sheet** and **Show Charts** as needed
3. Click **⚡ Fetch & Export**
4. View live charts and download individual Excel files per ticker

### Indian Stock Tickers
- **NSE:** append `.NS` → `INFY.NS`, `TCS.NS`, `RELIANCE.NS`
- **BSE:** append `.BO` → `INFY.BO`, `500209.BO`

---

## 📦 Tech Stack

- [Streamlit](https://streamlit.io/) — UI framework
- [yfinance](https://github.com/ranaroussi/yfinance) — Yahoo Finance API wrapper
- [Plotly](https://plotly.com/python/) — Interactive charts
- [BeautifulSoup4](https://www.crummy.com/software/BeautifulSoup/) — HTML scraping for Key Statistics
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel file generation
- [pandas](https://pandas.pydata.org/) — Data processing

---

## ☁️ Deploy to Streamlit Cloud

1. Fork / push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → set main file to `Yahoo_Scrapper_Strmlt.py`
4. Click **Deploy** 🎉

---

## 📄 License

MIT License — free to use and modify.
