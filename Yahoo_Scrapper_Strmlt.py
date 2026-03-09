import io
import time
import requests
import pandas as pd
import yfinance as yf
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ========== PAGE CONFIG ==========
st.set_page_config(
    page_title="FinScraper — Yahoo Finance Exporter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========== CUSTOM CSS ==========
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Syne', sans-serif;
}

.stApp {
    background: #0a0e1a;
    color: #e8eaf0;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #0f1626 !important;
    border-right: 1px solid #1e2d4a;
}

/* Header */
.fin-header {
    padding: 2rem 0 1rem 0;
    border-bottom: 1px solid #1e2d4a;
    margin-bottom: 2rem;
}
.fin-header h1 {
    font-family: 'Syne', sans-serif;
    font-size: 2.8rem;
    font-weight: 800;
    background: linear-gradient(135deg, #4fc3f7 0%, #00e5ff 50%, #69f0ae 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin: 0;
    letter-spacing: -1px;
}
.fin-header p {
    color: #5a7a9a;
    font-family: 'Space Mono', monospace;
    font-size: 0.85rem;
    margin-top: 0.4rem;
    letter-spacing: 0.05em;
}

/* Metric cards */
.metric-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 1rem;
    margin: 1.5rem 0;
}
.metric-card {
    background: #111827;
    border: 1px solid #1e2d4a;
    border-radius: 12px;
    padding: 1.2rem;
    position: relative;
    overflow: hidden;
}
.metric-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, #4fc3f7, #69f0ae);
}
.metric-label {
    font-family: 'Space Mono', monospace;
    font-size: 0.68rem;
    color: #5a7a9a;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-bottom: 0.5rem;
}
.metric-value {
    font-family: 'Syne', sans-serif;
    font-size: 1.4rem;
    font-weight: 700;
    color: #e8eaf0;
}
.metric-value.positive { color: #69f0ae; }
.metric-value.negative { color: #ff5252; }

/* Ticker badge */
.ticker-badge {
    display: inline-block;
    background: linear-gradient(135deg, #1a2744, #0f1e38);
    border: 1px solid #2a4070;
    border-radius: 8px;
    padding: 0.3rem 0.8rem;
    font-family: 'Space Mono', monospace;
    font-size: 0.85rem;
    color: #4fc3f7;
    margin: 0.2rem;
}

/* Section heading */
.section-title {
    font-family: 'Syne', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: #4fc3f7;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin: 2rem 0 1rem 0;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid #1e2d4a;
}

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #4fc3f7, #00e5ff) !important;
    color: #0a0e1a !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.6rem 2rem !important;
    font-size: 1rem !important;
    letter-spacing: 0.03em !important;
    transition: all 0.2s !important;
    width: 100%;
}
.stDownloadButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 24px rgba(79, 195, 247, 0.35) !important;
}

/* Text input */
.stTextArea textarea {
    background: #111827 !important;
    border: 1px solid #1e2d4a !important;
    border-radius: 10px !important;
    color: #e8eaf0 !important;
    font-family: 'Space Mono', monospace !important;
    font-size: 0.9rem !important;
}
.stTextArea textarea:focus {
    border-color: #4fc3f7 !important;
    box-shadow: 0 0 0 2px rgba(79,195,247,0.15) !important;
}

/* Primary button */
.stButton > button {
    background: linear-gradient(135deg, #1a2744, #0f1e38) !important;
    color: #4fc3f7 !important;
    border: 1px solid #2a4070 !important;
    border-radius: 10px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important;
    width: 100%;
    padding: 0.6rem !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: #0f1626;
    border-radius: 10px;
    padding: 4px;
    gap: 4px;
    border: 1px solid #1e2d4a;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Syne', sans-serif !important;
    color: #5a7a9a !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}
.stTabs [aria-selected="true"] {
    background: #1a2744 !important;
    color: #4fc3f7 !important;
}

/* Progress / spinner */
.stProgress > div > div {
    background: linear-gradient(90deg, #4fc3f7, #69f0ae) !important;
}

/* Dataframe */
.stDataFrame {
    border: 1px solid #1e2d4a !important;
    border-radius: 10px !important;
}

/* Status box */
.status-box {
    background: #0f1e38;
    border: 1px solid #1e2d4a;
    border-left: 3px solid #4fc3f7;
    border-radius: 8px;
    padding: 0.8rem 1.2rem;
    font-family: 'Space Mono', monospace;
    font-size: 0.8rem;
    color: #5a9acc;
    margin: 0.5rem 0;
}
.status-box.success { border-left-color: #69f0ae; color: #69f0ae; }
.status-box.error   { border-left-color: #ff5252; color: #ff5252; }
.status-box.warning { border-left-color: #ffb300; color: #ffb300; }
</style>
""", unsafe_allow_html=True)

# ========== KNOWN SECTION HEADINGS ==========
STATS_SECTION_NAMES = [
    "Valuation Measures", "Fiscal Year", "Profitability",
    "Management Effectiveness", "Income Statement", "Balance Sheet",
    "Cash Flow Statement", "Stock Price History", "Share Statistics",
    "Dividends & Splits",
]
ANALYSIS_SECTION_NAMES = [
    "Earnings Estimate", "Revenue Estimate", "Earnings History",
    "EPS Trend", "EPS Revisions", "Growth Estimates",
]

# ========== RATE LIMIT CONFIG ==========
INTER_TICKER_DELAY   = 3.0   # seconds between tickers
RETRY_DELAYS         = [5, 15, 30]  # exponential back-off per attempt
SCRAPE_DELAY         = 2.0   # seconds between the two scrape calls per ticker

# ========== SCRAPING ==========
def safe_df(df):
    return df if isinstance(df, pd.DataFrame) else pd.DataFrame()

def scrape_yahoo_sections(url, section_names=None, retries=3):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "DNT": "1", "Connection": "close"
    }
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            sections = []
            for i, table in enumerate(soup.find_all("table")):
                heading = None
                heading_tag = table.find_previous_sibling(["h3", "h2", "strong", "span"])
                if heading_tag:
                    text = heading_tag.get_text(strip=True)
                    if text and len(text) > 2 and text.lower() not in ("--", ""):
                        heading = text
                if not heading and section_names and i < len(section_names):
                    heading = section_names[i]
                if not heading:
                    heading = f"Section {i + 1}"
                rows = []
                for tr in table.find_all("tr"):
                    cols = [td.get_text(strip=True) for td in tr.find_all("td")]
                    if cols:
                        rows.append(cols)
                if rows:
                    sections.append((heading, pd.DataFrame(rows)))
            return sections
        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code == 429:
                wait = RETRY_DELAYS[min(attempt, len(RETRY_DELAYS) - 1)]
                time.sleep(wait)
            else:
                break
        except Exception:
            wait = RETRY_DELAYS[min(attempt, len(RETRY_DELAYS) - 1)]
            time.sleep(wait)
    return []

# ========== CACHED DATA FETCHERS ==========
# TTL = 3600 seconds (1 hour) — data won't be re-fetched on every Streamlit rerun
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_ticker_data(ticker: str):
    """Fetch all yfinance data for a ticker. Cached for 1 hour."""
    for attempt in range(3):
        try:
            stock = yf.Ticker(ticker)

            history_df = safe_df(stock.history(period="1y"))
            if not history_df.empty:
                try:
                    history_df.index = history_df.index.tz_localize(None)
                except Exception:
                    history_df.index = history_df.index.tz_convert(None)
                history_df = history_df.sort_index(ascending=False)

            balance_sheet_df = safe_df(stock.balance_sheet)
            financials_df    = safe_df(stock.financials)
            cashflow_df      = safe_df(stock.cashflow)
            info             = stock.info or {}

            return {
                "history":       history_df,
                "balance_sheet": balance_sheet_df,
                "financials":    financials_df,
                "cashflow":      cashflow_df,
                "info":          info,
                "error":         None,
            }
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "Too Many Requests" in err_str or "Rate" in err_str:
                wait = RETRY_DELAYS[min(attempt, len(RETRY_DELAYS) - 1)]
                time.sleep(wait)
            else:
                return {"error": err_str}
    return {"error": "Rate limited after 3 retries. Please wait a few minutes and try again."}


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_scraped_sections(ticker: str, include_analysis: bool):
    """Scrape Yahoo Finance statistics & analysis pages. Cached for 1 hour."""
    stats_url    = f"https://finance.yahoo.com/quote/{ticker}/key-statistics"
    analysis_url = f"https://finance.yahoo.com/quote/{ticker}/analysis"

    stats_sections = scrape_yahoo_sections(stats_url, STATS_SECTION_NAMES)

    # Small delay between the two HTTP requests to avoid hammering Yahoo
    time.sleep(SCRAPE_DELAY)

    analysis_sections = []
    if include_analysis:
        analysis_sections = scrape_yahoo_sections(analysis_url, ANALYSIS_SECTION_NAMES)

    return stats_sections, analysis_sections


# ========== EXCEL BUILDER ==========
def build_excel(ticker, history_df, balance_sheet_df,
                financials_df, cashflow_df,
                yahoo_stats_sections, yahoo_analysis_sections):
    buf = io.BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if not history_df.empty:
            h = history_df.round(2)
            h.to_excel(writer, sheet_name="Price_History")

        def format_sheet(df, sheet_name):
            d = df.apply(pd.to_numeric, errors='coerce') / 1000
            d = d.round(1)
            d.to_excel(writer, sheet_name=sheet_name)
            sheet = writer.sheets[sheet_name]
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row,
                                       min_col=1, max_col=sheet.max_column):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.0'

        if not balance_sheet_df.empty:
            format_sheet(balance_sheet_df, "Balance_Sheet")
        if not financials_df.empty:
            format_sheet(financials_df, "Income_Stmnt")
        if not cashflow_df.empty:
            format_sheet(cashflow_df, "Cashflow")

        def write_sectioned_sheet(sheet_name, sections):
            sheet = writer.book.create_sheet(sheet_name)
            hdr_fill = PatternFill("solid", fgColor="0F3460")
            bold = Font(bold=True, color="4FC3F7", size=11)
            for heading, df in sections:
                sheet.append([heading])
                r = sheet.max_row
                ncols = max(2, df.shape[1])
                sheet.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
                cell = sheet.cell(row=r, column=1)
                cell.font = bold
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal="left", vertical="center")
                for row in df.itertuples(index=False):
                    sheet.append(list(row))
                sheet.append([])
            for col in sheet.columns:
                max_len = max((len(str(c.value or "")) for c in col), default=10)
                sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)

        if yahoo_stats_sections:
            write_sectioned_sheet("Key_Statistics", yahoo_stats_sections)
        if yahoo_analysis_sections:
            write_sectioned_sheet("Analysis", yahoo_analysis_sections)

    buf.seek(0)
    return buf

# ========== CHARTS ==========
def make_price_chart(history_df, ticker):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=history_df.index, y=history_df["Close"],
        mode="lines", name="Close",
        line=dict(color="#4fc3f7", width=2),
        fill="tozeroy",
        fillcolor="rgba(79,195,247,0.07)"
    ))
    fig.add_trace(go.Bar(
        x=history_df.index, y=history_df["Volume"],
        name="Volume", yaxis="y2",
        marker_color="rgba(105,240,174,0.25)"
    ))
    fig.update_layout(
        title=f"{ticker} — 1 Year Price & Volume",
        paper_bgcolor="#0a0e1a", plot_bgcolor="#0f1626",
        font=dict(family="Syne", color="#e8eaf0"),
        xaxis=dict(gridcolor="#1e2d4a", showgrid=True),
        yaxis=dict(gridcolor="#1e2d4a", title="Price"),
        yaxis2=dict(overlaying="y", side="right", title="Volume",
                    showgrid=False, tickfont=dict(color="#69f0ae")),
        legend=dict(bgcolor="#0f1626", bordercolor="#1e2d4a"),
        hovermode="x unified",
        height=420,
    )
    return fig

def make_financials_chart(financials_df, ticker):
    if financials_df.empty:
        return None
    rows = ["Total Revenue", "Gross Profit", "Net Income"]
    available = [r for r in rows if r in financials_df.index]
    if not available:
        return None
    df = financials_df.loc[available].T / 1e9
    df.index = [str(c)[:10] for c in df.index]
    colors = ["#4fc3f7", "#69f0ae", "#ff9800"]
    fig = go.Figure()
    for i, col in enumerate(df.columns):
        fig.add_trace(go.Bar(name=col, x=df.index, y=df[col],
                             marker_color=colors[i % len(colors)]))
    fig.update_layout(
        title=f"{ticker} — Revenue & Profit (USD Billion)",
        barmode="group",
        paper_bgcolor="#0a0e1a", plot_bgcolor="#0f1626",
        font=dict(family="Syne", color="#e8eaf0"),
        xaxis=dict(gridcolor="#1e2d4a"),
        yaxis=dict(gridcolor="#1e2d4a", title="USD Billion"),
        legend=dict(bgcolor="#0f1626", bordercolor="#1e2d4a"),
        height=380,
    )
    return fig

def make_balance_chart(balance_df, ticker):
    if balance_df.empty:
        return None
    rows = ["Total Assets", "Total Liabilities Net Minority Interest", "Stockholders Equity"]
    labels = ["Total Assets", "Total Liabilities", "Stockholders Equity"]
    available = [(r, l) for r, l in zip(rows, labels) if r in balance_df.index]
    if not available:
        return None
    df_plot = pd.DataFrame({
        l: balance_df.loc[r].values / 1e9
        for r, l in available
    }, index=[str(c)[:10] for c in balance_df.columns])
    colors = ["#4fc3f7", "#ff5252", "#69f0ae"]
    fig = go.Figure()
    for i, col in enumerate(df_plot.columns):
        fig.add_trace(go.Bar(name=col, x=df_plot.index, y=df_plot[col],
                             marker_color=colors[i]))
    fig.update_layout(
        title=f"{ticker} — Balance Sheet (USD Billion)",
        barmode="group",
        paper_bgcolor="#0a0e1a", plot_bgcolor="#0f1626",
        font=dict(family="Syne", color="#e8eaf0"),
        xaxis=dict(gridcolor="#1e2d4a"),
        yaxis=dict(gridcolor="#1e2d4a", title="USD Billion"),
        legend=dict(bgcolor="#0f1626", bordercolor="#1e2d4a"),
        height=380,
    )
    return fig

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("""
    <div style='padding: 1rem 0;'>
        <div style='font-family: Syne, sans-serif; font-size: 1.5rem; font-weight: 800;
                    background: linear-gradient(135deg, #4fc3f7, #69f0ae);
                    -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>
            📊 FinScraper
        </div>
        <div style='font-family: Space Mono, monospace; font-size: 0.7rem; color: #5a7a9a;
                    margin-top: 0.3rem;'>Yahoo Finance Exporter</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**Enter Tickers**")
    ticker_input = st.text_area(
        label="tickers",
        placeholder="AAPL\nMSFT\nRELIANCE.NS\nCIPLA.NS",
        height=160,
        label_visibility="collapsed"
    )

    include_analysis = st.checkbox("Include Analysis Sheet", value=True)
    include_charts   = st.checkbox("Show Charts in App",    value=True)

    fetch_btn = st.button("⚡ Fetch & Export", use_container_width=True)

    st.markdown("---")
    st.markdown("""
    <div style='font-family: Space Mono, monospace; font-size: 0.7rem; color: #3a5a7a; line-height: 1.8;'>
    ℹ️ Supports any Yahoo Finance ticker<br>
    🇮🇳 Indian stocks: add .NS (NSE) or .BO (BSE)<brko>
    📦 One Excel file per ticker<br>
    ⚡ Data cached for 1 hour<br>
    </div>
    """, unsafe_allow_html=True)

# ========== MAIN AREA ==========
st.markdown("""
<div class='fin-header'>
    <h1>FinScraper</h1>
    <p>YAHOO FINANCE DATA EXPORTER &nbsp;·&nbsp; EXCEL + CHARTS</p>
</div>
""", unsafe_allow_html=True)

if not fetch_btn:
    col1, col2, col3 = st.columns(3)
    for col, icon, title, desc in [
        (col1, "📥", "Excel Export", "All sheets: Price History, Balance Sheet, Income Statement, Cashflow, Key Statistics, Analysis"),
        (col2, "📈", "Interactive Charts", "1-year price & volume, Revenue & Profit, Balance Sheet visualized with Plotly"),
        (col3, "🌍", "Global Tickers", "Supports US, Indian (NSE/BSE), and all Yahoo Finance listed stocks"),
    ]:
        with col:
            st.markdown(f"""
            <div class='metric-card' style='text-align:center; padding: 2rem 1rem;'>
                <div style='font-size:2rem; margin-bottom:0.8rem;'>{icon}</div>
                <div style='font-family:Syne,sans-serif; font-weight:700; font-size:1rem;
                            color:#e8eaf0; margin-bottom:0.5rem;'>{title}</div>
                <div style='font-family:Space Mono,monospace; font-size:0.72rem;
                            color:#5a7a9a; line-height:1.6;'>{desc}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("""
    <div style='margin-top:3rem; text-align:center; font-family:Space Mono,monospace;
                font-size:0.8rem; color:#3a5a7a;'>
        ← Enter tickers in the sidebar and click Fetch & Export
    </div>
    """, unsafe_allow_html=True)

else:
    tickers = [t.strip().upper() for t in ticker_input.replace(",", "\n").splitlines() if t.strip()]
    if not tickers:
        st.error("Please enter at least one ticker symbol.")
        st.stop()

    st.markdown(
        "**Fetching:** " + " ".join([f"<span class='ticker-badge'>{t}</span>" for t in tickers]),
        unsafe_allow_html=True
    )
    st.markdown("")

    for idx, ticker in enumerate(tickers):
        st.markdown(f"<div class='section-title'>📌 {ticker}</div>", unsafe_allow_html=True)
        log = st.empty()

        # ── Add inter-ticker delay (skip for the very first ticker) ──
        if idx > 0:
            log.markdown(
                f"<div class='status-box warning'>⏳ Waiting {INTER_TICKER_DELAY}s before next request to avoid rate limits...</div>",
                unsafe_allow_html=True
            )
            time.sleep(INTER_TICKER_DELAY)

        try:
            log.markdown(
                f"<div class='status-box'>⏳ Fetching yfinance data for {ticker}...</div>",
                unsafe_allow_html=True
            )

            # ── Cached yfinance call ──
            result = fetch_ticker_data(ticker)

            if result.get("error"):
                log.markdown(
                    f"<div class='status-box error'>❌ Failed for {ticker}: {result['error']}</div>",
                    unsafe_allow_html=True
                )
                st.markdown("<br>", unsafe_allow_html=True)
                continue

            history_df       = result["history"]
            balance_sheet_df = result["balance_sheet"]
            financials_df    = result["financials"]
            cashflow_df      = result["cashflow"]
            info             = result["info"]

            log.markdown(
                f"<div class='status-box'>⏳ Scraping Key Statistics page for {ticker}...</div>",
                unsafe_allow_html=True
            )

            # ── Cached scrape call ──
            yahoo_stats_sections, yahoo_analysis_sections = fetch_scraped_sections(
                ticker, include_analysis
            )

            log.markdown(
                f"<div class='status-box success'>✅ Data fetched successfully for {ticker}</div>",
                unsafe_allow_html=True
            )

            # ── KEY METRICS ROW ──
            price      = info.get("currentPrice") or info.get("regularMarketPrice", "N/A")
            mktcap     = info.get("marketCap", 0)
            mktcap_str = f"${mktcap/1e9:.2f}B" if mktcap else "N/A"
            pe         = info.get("trailingPE", "N/A")
            change     = info.get("regularMarketChangePercent", None)
            change_str = f"{change*100:+.2f}%" if change else "N/A"
            sector     = info.get("sector", "N/A")
            cls        = "positive" if (change or 0) >= 0 else "negative"

            st.markdown(f"""
            <div class='metric-grid'>
                <div class='metric-card'>
                    <div class='metric-label'>Current Price</div>
                    <div class='metric-value'>{price}</div>
                </div>
                <div class='metric-card'>
                    <div class='metric-label'>Day Change</div>
                    <div class='metric-value {cls}'>{change_str}</div>
                </div>
                <div class='metric-card'>
                    <div class='metric-label'>Market Cap</div>
                    <div class='metric-value'>{mktcap_str}</div>
                </div>
                <div class='metric-card'>
                    <div class='metric-label'>Trailing P/E</div>
                    <div class='metric-value'>{round(pe,2) if isinstance(pe, float) else pe}</div>
                </div>
                <div class='metric-card'>
                    <div class='metric-label'>Sector</div>
                    <div class='metric-value' style='font-size:1rem;'>{sector}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── CHARTS ──
            if include_charts and not history_df.empty:
                tab1, tab2, tab3 = st.tabs(["📈 Price & Volume", "💰 Revenue & Profit", "🏦 Balance Sheet"])
                with tab1:
                    st.plotly_chart(make_price_chart(history_df, ticker), use_container_width=True)
                with tab2:
                    fig2 = make_financials_chart(financials_df, ticker)
                    if fig2: st.plotly_chart(fig2, use_container_width=True)
                    else:    st.info("Financials data not available.")
                with tab3:
                    fig3 = make_balance_chart(balance_sheet_df, ticker)
                    if fig3: st.plotly_chart(fig3, use_container_width=True)
                    else:    st.info("Balance sheet data not available.")

            # ── BUILD & DOWNLOAD EXCEL ──
            excel_buf = build_excel(
                ticker, history_df,
                balance_sheet_df, financials_df, cashflow_df,
                yahoo_stats_sections, yahoo_analysis_sections
            )
            fname = f"{ticker.replace('.','_')}_YahooFinance_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                label=f"⬇️  Download {ticker} Excel",
                data=excel_buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{ticker}"
            )

        except Exception as e:
            log.markdown(
                f"<div class='status-box error'>❌ Failed for {ticker}: {e}</div>",
                unsafe_allow_html=True
            )

        st.markdown("<br>", unsafe_allow_html=True)
