import datetime as dt
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import yfinance as yf

# ------------------------------------------------------------------------------
# 1. PUSLAPIO KONFIGŪRACIJA IR TEMA
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Akcijų Analizės Agentas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .stApp {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
    }
    [data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 1px solid #E2E8F0 !important;
    }
    p, span, label, [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p {
        color: #0F172A !important;
    }
    label, [data-testid="stWidgetLabel"] {
        font-weight: 600 !important;
        margin-bottom: 4px !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.4rem !important;
        font-weight: 700 !important;
        color: #0F172A !important;
    }
    [data-testid="stMetricLabel"] p {
        color: #475569 !important;
    }
    h1, h2, h3, h4, h5, h6 {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    .stTextInput input {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1.5px solid #0F172A !important;
        border-radius: 10px !important;
    }

    div[data-testid="stNumberInput"] > div > div,
    div[data-testid="stSelectbox"] > div > div {
        background-color: #FFFFFF !important;
        border: 1.5px solid #0F172A !important;
        border-radius: 10px !important;
        overflow: hidden !important;
    }

    div[data-testid="stNumberInput"] button {
        background-color: #E2E8F0 !important;
        color: #0F172A !important;
        border: none !important;
        border-left: 1px solid #E2E8F0 !important;
        font-weight: bold !important;
        transition: background-color 0.2s ease !important;
    }
    div[data-testid="stNumberInput"] button:hover {
        background-color: #CBD5E1 !important;
    }

    div[data-testid="stSelectbox"] [data-baseweb="select"] > div:last-child {
        background-color: #E2E8F0 !important;
        transition: background-color 0.2s ease !important;
        padding-left: 12px !important;
        padding-right: 12px !important;
    }
    div[data-testid="stSelectbox"] [data-baseweb="select"]:hover > div:last-child {
        background-color: #CBD5E1 !important;
    }

    div[data-testid="stNumberInput"] button svg,
    div[data-testid="stNumberInput"] button:disabled svg,
    div[data-testid="stSelectbox"] svg {
        fill: #0F172A !important;
        color: #0F172A !important;
        stroke: #0F172A !important;
    }

    div[data-testid="stNumberInput"] button:disabled {
        background-color: #E2E8F0 !important;
        opacity: 0.7 !important;
    }
    div[data-testid="stNumberInput"] input,
    div[data-testid="stSelectbox"] input {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        font-size: 16px !important;
    }

    div.stDownloadButton > button {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        width: 100% !important;
        padding: 0.5rem 1rem !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #E2E8F0 !important;
        border-color: #CBD5E1 !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

st.title("📊 Akcijų Analizės Agentas")

# ------------------------------------------------------------------------------
# 2. ŠONINĖ JUOSTA (PARAMETRAI IR MYGTUKAS)
# ------------------------------------------------------------------------------
st.sidebar.header("⚙️ Parametrai")

tickers_raw = st.sidebar.text_input(
    "Įveskite tickerį (arba kelis, atskirtus kableliu):", "GOOGL"
)
st.sidebar.caption(
    "💡 Pirmasis įrašytas tickeris gauna pilną, išsamią analizę. "
    "Papildomi tickeriai (pvz. `GOOGL, MSFT, AAPL`) bus parodyti palyginimo lentelėje ir grafike."
)

PERIOD_OPTIONS = {
    "1 mėnuo": "1mo",
    "6 mėnesiai": "6mo",
    "1 metai": "1y",
    "5 metai": "5y",
    "Max": "max",
}

sidebar_download_placeholder = st.sidebar.empty()

# Pagrindinio ir palyginamųjų tickerių išskyrimas
tickers_list = [x.strip().upper() for x in tickers_raw.split(",") if x.strip()]
ticker_input = tickers_list[0] if tickers_list else ""
compare_tickers = tickers_list[1:] if len(tickers_list) > 1 else []


# ------------------------------------------------------------------------------
# PAGALBINĖS FUNKCIJOS
# ------------------------------------------------------------------------------
def calc_cagr(start_val, end_val, periods=3):
    try:
        if start_val is not None and end_val is not None:
            start_val, end_val = float(start_val), float(end_val)
            if start_val > 0 and end_val > 0 and periods > 0:
                return ((end_val / start_val) ** (1 / periods) - 1) * 100
    except (ValueError, TypeError, ZeroDivisionError):
        pass
    return None


def format_number(val, is_currency=True, currency="USD"):
    if val is None or pd.isna(val):
        return "N/A"
    try:
        val = float(val)
    except (ValueError, TypeError):
        return "N/A"

    abs_val = abs(val)
    if abs_val >= 1e12:
        res = f"{val / 1e12:.2f}T"
    elif abs_val >= 1e9:
        res = f"{val / 1e9:.2f}B"
    elif abs_val >= 1e6:
        res = f"{val / 1e6:.2f}M"
    else:
        res = f"{val:,.2f}"
    return f"{res} {currency}" if is_currency else res


def classify_insider(text):
    t = str(text or "").lower()
    if any(w in t for w in ["sale", "sell", "s - sale", "disposition"]):
        return "Pardavimas"
    if any(w in t for w in ["purchase", "buy", "p - purchase", "acquisition"]):
        return "Pirkimas"
    return "Kita"


def parse_value(val_str):
    if not val_str or val_str == "N/A" or val_str == "-":
        return None
    try:
        cleaned = (
            str(val_str)
            .replace("%", "")
            .replace("USD", "")
            .replace("EUR", "")
            .replace("+", "")
            .replace(",", "")
            .strip()
        )
        return float(cleaned)
    except Exception:
        return None


def fmt_val(v, decimals=2, is_pct=False):
    if v is None or pd.isna(v):
        return "N/A"
    try:
        val = float(v)
        if is_pct:
            return f"{val:.2f}%"
        return f"{val:.{decimals}f}"
    except (ValueError, TypeError):
        return str(v)


def fmt_large(v):
    if v is None or pd.isna(v):
        return "N/A"
    try:
        return f"{float(v):,.0f}"
    except (ValueError, TypeError):
        return str(v)


def compute_div_yield_pct(info):
    """Grąžina dividendų pajamingumą procentais arba None."""
    dy = info.get("dividendYield")
    if dy is None or pd.isna(dy):
        return None
    dy = float(dy)
    return dy * 100 if dy < 0.1 else dy


def quarter_label(col):
    """Formatuoja stulpelio datą kaip 'YYYY Qn'."""
    ts = pd.to_datetime(col)
    q = (ts.month - 1) // 3 + 1
    return f"{ts.year} Q{q}"


def extract_financials_series(inc_df, label_func, max_cols=4):
    """Ištraukia (laikotarpių žymes, pajamas, grynąjį pelną) iš income statement DataFrame.

    Praleidžia stulpelius (laikotarpius), kuriems Yahoo Finance neturi jokių
    pajamų/pelno reikšmių, kad grafike neatsirastų tuščių/nulinių stulpelių.
    """
    labels, revenues, net_incomes = [], [], []
    if inc_df is None or inc_df.empty:
        return labels, revenues, net_incomes

    valid_cols = sorted([c for c in inc_df.columns if pd.notna(c)])
    rev_row_candidates = ["Total Revenue", "Operating Revenue"]
    ni_row_candidates = ["Net Income", "Net Income Common Stockholders"]

    r_series, ni_series = None, None
    for rc in rev_row_candidates:
        if rc in inc_df.index:
            r_series = inc_df.loc[rc]
            break
    for nic in ni_row_candidates:
        if nic in inc_df.index:
            ni_series = inc_df.loc[nic]
            break

    cols_with_data = []
    for col in valid_cols:
        r_val = r_series.loc[col] if r_series is not None and col in r_series.index else None
        ni_val = ni_series.loc[col] if ni_series is not None and col in ni_series.index else None
        if (r_val is not None and pd.notna(r_val)) or (ni_val is not None and pd.notna(ni_val)):
            cols_with_data.append(col)

    recent_cols = (
        cols_with_data[-max_cols:] if len(cols_with_data) >= max_cols else cols_with_data
    )
    for col in recent_cols:
        labels.append(label_func(col))
        r_val = (
            r_series.loc[col] if r_series is not None and col in r_series.index else 0
        )
        ni_val = (
            ni_series.loc[col] if ni_series is not None and col in ni_series.index else 0
        )
        revenues.append(float(r_val) if pd.notna(r_val) else 0.0)
        net_incomes.append(float(ni_val) if pd.notna(ni_val) else 0.0)

    return labels, revenues, net_incomes


REC_MAP = {
    "strong_buy": "STRONG BUY",
    "buy": "BUY",
    "hold": "HOLD",
    "underperform": "UNDERPERFORM",
    "sell": "SELL",
    "strong_sell": "STRONG SELL",
}


def style_main_metrics(df, price_val):
    styles = pd.DataFrame(
        "background-color: #ffffff; color: #0f172a;",
        index=df.index,
        columns=df.columns,
    )
    for idx, row in df.iterrows():
        metric = str(row["Rodiklis"])
        val_str = str(row["Reikšmė"])
        num = parse_value(val_str)

        is_good = False
        if num is not None:
            if metric == "Forward P/E" and num < 18:
                is_good = True
            elif metric == "Total Debt/Equity" and num <= 50:
                is_good = True
            elif metric == "ROE" and num >= 20:
                is_good = True
            elif metric == "MGMT owns" and num > 1.0:
                is_good = True
            elif metric == "Dividend Yield" and num >= 5.0:
                is_good = True
            elif (
                    metric in ["3 Year EPS Growth Rate", "3-Year Sales Growth Rate"]
                    and num > 25
            ):
                is_good = True
            elif metric == "1y Target Est" and price_val is not None and num > price_val:
                is_good = True

        if is_good:
            styles.loc[idx, "Reikšmė"] = (
                "color: #15803d; font-weight: bold; background-color: #f0fdf4;"
            )

    return styles


def style_price_change(df):
    styles = pd.DataFrame(
        "background-color: #ffffff; color: #0f172a;",
        index=df.index,
        columns=df.columns,
    )
    for idx, row in df.iterrows():
        val_str = str(row["Pokytis %"])
        num = parse_value(val_str)

        if num is not None:
            if num > 0:
                styles.loc[idx, "Pokytis %"] = (
                    "color: #15803d; font-weight: bold; background-color: #f0fdf4;"
                )
            elif num < 0:
                styles.loc[idx, "Pokytis %"] = (
                    "color: #b91c1c; font-weight: bold; background-color: #fef2f2;"
                )

    return styles


def style_insider_df(df):
    return pd.DataFrame(
        "background-color: #ffffff; color: #0f172a;",
        index=df.index,
        columns=df.columns,
    )


# ------------------------------------------------------------------------------
# DUOMENŲ TRAUKIMO FUNKCIJOS (CACHED)
# ------------------------------------------------------------------------------
@st.cache_data(ttl=600, show_spinner=False)
def validate_ticker(ticker: str) -> bool:
    """Greitas patikrinimas, ar tickeris egzistuoja."""
    try:
        t = yf.Ticker(ticker)
        if t.fast_info.get("lastPrice") is not None:
            return True
        h = t.history(period="1d")
        return not h.empty
    except Exception:
        return False


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_info(ticker: str):
    return yf.Ticker(ticker).info or {}


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_history(ticker: str, period: str):
    return yf.Ticker(ticker).history(period=period)


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_income_stmt(ticker: str):
    return yf.Ticker(ticker).income_stmt


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_quarterly_income_stmt(ticker: str):
    return yf.Ticker(ticker).quarterly_income_stmt


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_insider_transactions(ticker: str):
    return yf.Ticker(ticker).insider_transactions


# ------------------------------------------------------------------------------
# 3. DUOMENŲ TRAUKIMAS IR APDOROJIMAS (PAGRINDINIS TICKERIS)
# ------------------------------------------------------------------------------
if ticker_input:

    if not validate_ticker(ticker_input):
        st.error(
            f"❌ Tickeris **{ticker_input}** nerastas arba šiuo metu neturi rinkos duomenų. "
            f"Patikrinkite simbolį — pavyzdžiui: `AAPL`, `MSFT`, `GOOGL`, `TSLA`, `NVDA`."
        )
        st.stop()

    with st.spinner(f"Renkami ir skaičiuojami {ticker_input} duomenys..."):
        try:
            info = fetch_info(ticker_input)
            hist_5y = fetch_history(ticker_input, "5y")

            price = info.get("currentPrice") or info.get("regularMarketPrice")
            if (price is None or pd.isna(price)) and not hist_5y.empty:
                price = float(hist_5y["Close"].iloc[-1])

            if price is None or hist_5y.empty:
                st.error("Nepavyko rasti rinkos duomenų šiam tickeriui.")
                st.stop()

            cur = info.get("currency", "USD")
            long_name = info.get("longName", ticker_input)

            eps_growth_3y = None
            sales_growth_3y = None

            financials_years = []
            revenues_vals = []
            net_income_vals = []
            financials_periods_q = []
            revenues_vals_q = []
            net_income_vals_q = []

            try:
                inc = fetch_income_stmt(ticker_input)
                if inc is not None and not inc.empty:
                    financials_years, revenues_vals, net_income_vals = extract_financials_series(
                        inc, lambda c: pd.to_datetime(c).strftime("%Y"), max_cols=4
                    )

                    valid_cols = sorted([c for c in inc.columns if pd.notna(c)])
                    if len(valid_cols) >= 4:
                        col_latest = valid_cols[-1]
                        col_base = valid_cols[-4]

                        for cand in ["Diluted EPS", "Basic EPS"]:
                            if cand in inc.index:
                                eps_base_val = inc.loc[cand, col_base]
                                eps_latest_val = inc.loc[cand, col_latest]
                                eps_growth_3y = calc_cagr(eps_base_val, eps_latest_val, 3)
                                break

                        for cand in ["Total Revenue", "Operating Revenue"]:
                            if cand in inc.index:
                                rev_base_val = inc.loc[cand, col_base]
                                rev_latest_val = inc.loc[cand, col_latest]
                                sales_growth_3y = calc_cagr(rev_base_val, rev_latest_val, 3)
                                break
            except Exception:
                pass

            try:
                inc_q = fetch_quarterly_income_stmt(ticker_input)
                if inc_q is not None and not inc_q.empty:
                    financials_periods_q, revenues_vals_q, net_income_vals_q = (
                        extract_financials_series(inc_q, quarter_label, max_cols=8)
                    )
            except Exception:
                pass

            insider_df = None
            total_buy, total_sell, net_flow = 0.0, 0.0, 0.0
            monthly_insider_buy = {}
            monthly_insider_sell = {}

            try:
                ins = fetch_insider_transactions(ticker_input)
                if ins is not None and not ins.empty:
                    ins = ins.copy()
                    date_col = None
                    for col in ["Start Date", "Date", "Transaction Date"]:
                        if col in ins.columns:
                            date_col = col
                            break

                    if date_col:
                        ins[date_col] = pd.to_datetime(ins[date_col], errors="coerce")
                        ins = ins.dropna(subset=[date_col])
                        cutoff = pd.Timestamp.now() - pd.Timedelta(days=365)
                        ins_12m = ins[ins[date_col] >= cutoff].sort_values(
                            date_col, ascending=False
                        )

                        if not ins_12m.empty:
                            rows = []
                            for _, r in ins_12m.iterrows():
                                val = (
                                    float(r.get("Value")) if pd.notna(r.get("Value")) else 0.0
                                )
                                txt = str(r.get("Transaction") or r.get("Text") or "")
                                tx_type = classify_insider(txt)

                                if tx_type == "Pirkimas":
                                    total_buy += val
                                elif tx_type == "Pardavimas":
                                    total_sell += val

                                m_key = r[date_col].strftime("%Y-%m")
                                if tx_type == "Pirkimas":
                                    monthly_insider_buy[m_key] = (
                                        monthly_insider_buy.get(m_key, 0.0) + val
                                    )
                                elif tx_type == "Pardavimas":
                                    monthly_insider_sell[m_key] = (
                                        monthly_insider_sell.get(m_key, 0.0) + val
                                    )

                                rows.append(
                                    {
                                        "Data": r[date_col].strftime("%Y-%m-%d"),
                                        "Asmuo": r.get("Insider", "N/A"),
                                        "Pareigos": r.get("Position", ""),
                                        "Sandoris": txt,
                                        "Akcijos": (
                                            f"{r.get('Shares'):,.0f}"
                                            if pd.notna(r.get("Shares"))
                                            else "N/A"
                                        ),
                                        "Suma": f"{val:,.2f}" if pd.notna(val) else "0.00",
                                    }
                                )

                            insider_df = pd.DataFrame(rows)
                            net_flow = total_buy - total_sell
            except Exception:
                pass

            # 1 SEKCIJA: AKCIJOS KAINOS ISTORIJA (GRAFIKAS)
            price_chg_pct = None
            price_chg_abs = None
            if len(hist_5y) > 1:
                prev_close = float(hist_5y["Close"].iloc[-2])
                if prev_close:
                    price_chg_abs = price - prev_close
                    price_chg_pct = (price_chg_abs / prev_close) * 100

            header_col1, header_col2 = st.columns([3, 1])
            with header_col1:
                st.subheader(f"{long_name} ({ticker_input})")
            with header_col2:
                st.metric(
                    label=f"Kaina ({cur})",
                    value=f"{price:,.2f}",
                    delta=(
                        f"{price_chg_abs:+.2f} ({price_chg_pct:+.2f}%)"
                        if price_chg_pct is not None
                        else None
                    ),
                )

            chart_placeholder = st.empty()
            period_selector_placeholder = st.container()

            with period_selector_placeholder:
                period_label = st.selectbox(
                    "📅 Kainos grafiko laikotarpis:",
                    list(PERIOD_OPTIONS.keys()),
                    index=3,
                    key="price_chart_period",
                )
            selected_period = PERIOD_OPTIONS[period_label]
            hist_display = (
                hist_5y
                if selected_period == "5y"
                else fetch_history(ticker_input, selected_period)
            )
            if hist_display is None or hist_display.empty:
                hist_display = hist_5y

            fig_price = go.Figure()
            fig_price.add_trace(
                go.Scatter(
                    x=hist_display.index,
                    y=hist_display["Close"],
                    mode="lines",
                    name=f"Kaina ({cur})",
                    line=dict(color="#38bdf8", width=2),
                )
            )
            fig_price.update_layout(
                title=f"Kainos istorija ({period_label})",
                paper_bgcolor="#0F172A",
                plot_bgcolor="#0F172A",
                font=dict(color="#FFFFFF"),
                margin=dict(l=20, r=20, t=40, b=20),
                xaxis=dict(
                    showgrid=True, gridcolor="#334155", title="Data", color="#FFFFFF"
                ),
                yaxis=dict(
                    showgrid=True,
                    gridcolor="#334155",
                    title=f"Kaina ({cur})",
                    color="#FFFFFF",
                ),
            )
            with chart_placeholder:
                st.plotly_chart(fig_price, use_container_width=True)

            st.divider()

            # 2 SEKCIJA: KAINOS POKYTIS
            st.subheader("KAINOS POKYTIS")
            df_price = pd.DataFrame()
            if not hist_5y.empty and len(hist_5y) > 1:
                hist_5y_local = hist_5y.copy()
                hist_5y_local.index = hist_5y_local.index.tz_localize(None)
                curr_p = price
                now_dt = hist_5y_local.index[-1]

                def get_historical_price(days_back=None, ytd=False):
                    if ytd:
                        target_dt = pd.Timestamp(year=now_dt.year, month=1, day=1)
                    else:
                        target_dt = now_dt - pd.Timedelta(days=days_back)
                    sub = hist_5y_local[hist_5y_local.index <= target_dt]
                    if not sub.empty:
                        return float(sub["Close"].iloc[-1])
                    return float(hist_5y_local["Close"].iloc[0])

                p_1m = get_historical_price(days_back=30)
                p_6m = get_historical_price(days_back=182)
                p_ytd = get_historical_price(ytd=True)
                p_1y = get_historical_price(days_back=365)
                p_5y = float(hist_5y_local["Close"].iloc[0])

                periods_data = [
                    ("5y", p_5y),
                    ("1y", p_1y),
                    ("YTD", p_ytd),
                    ("6M", p_6m),
                    ("1M", p_1m),
                ]

                price_change_list = []
                for label, past_p in periods_data:
                    if past_p and past_p > 0:
                        chg_pct = ((curr_p - past_p) / past_p) * 100
                        chg_str = f"{chg_pct:+.2f}%"
                    else:
                        chg_str = "N/A"

                    price_change_list.append(
                        {
                            "Laikotarpis": label,
                            "Kaina praeityje (USD)": (
                                f"{past_p:.2f}"
                                if isinstance(past_p, (int, float))
                                else "N/A"
                            ),
                            "Pokytis %": chg_str,
                        }
                    )

                df_price = pd.DataFrame(price_change_list)
                st.dataframe(
                    df_price.style.apply(style_price_change, axis=None),
                    use_container_width=True,
                    hide_index=True,
                )

            st.divider()

            # 3 SEKCIJA: PAGRINDINIAI RODIKLIAI
            st.subheader("PAGRINDINIAI RODIKLIAI")

            mgmt_own = info.get("heldPercentInsiders")
            mgmt_own_val = mgmt_own if mgmt_own is not None else "N/A"

            div_yield_val = compute_div_yield_pct(info)
            if div_yield_val is None:
                div_yield_val = "N/A"

            d_e = info.get("debtToEquity")

            main_metrics_data = [
                {
                    "Rodiklis": "PRICE",
                    "Reikšmė": fmt_val(price),
                    "Šaltinis / pastaba": "Yahoo Finance - currentPrice",
                },
                {
                    "Rodiklis": "Forward P/E",
                    "Reikšmė": fmt_val(info.get("forwardPE")),
                    "Šaltinis / pastaba": "Yahoo Finance - forwardPE",
                },
                {
                    "Rodiklis": "3 Year EPS Growth Rate",
                    "Reikšmė": (
                        f"{eps_growth_3y:.2f}%"
                        if eps_growth_3y is not None
                        else "N/A"
                    ),
                    "Šaltinis / pastaba": (
                        "Apskaičiuota (CAGR) iš Yahoo Finance finansinių ataskaitų"
                    ),
                },
                {
                    "Rodiklis": "3-Year Sales Growth Rate",
                    "Reikšmė": (
                        f"{sales_growth_3y:.2f}%"
                        if sales_growth_3y is not None
                        else "N/A"
                    ),
                    "Šaltinis / pastaba": (
                        "Apskaičiuota (CAGR) iš Yahoo Finance finansinių ataskaitų"
                    ),
                },
                {
                    "Rodiklis": "Cash",
                    "Reikšmė": fmt_large(info.get("totalCash")),
                    "Šaltinis / pastaba": "Yahoo Finance - totalCash",
                },
                {
                    "Rodiklis": "Debt",
                    "Reikšmė": fmt_large(info.get("totalDebt")),
                    "Šaltinis / pastaba": "Yahoo Finance - totalDebt",
                },
                {
                    "Rodiklis": "Total Debt/Equity",
                    "Reikšmė": (
                        f"{d_e:.2f}%"
                        if d_e is not None and not pd.isna(d_e)
                        else "N/A"
                    ),
                    "Šaltinis / pastaba": "Yahoo Finance - debtToEquity (mrq)",
                },
                {
                    "Rodiklis": "Book Value Per Share",
                    "Reikšmė": fmt_val(info.get("bookValue")),
                    "Šaltinis / pastaba": "Yahoo Finance - bookValue (mrq)",
                },
                {
                    "Rodiklis": "ROE",
                    "Reikšmė": (
                        f"{info.get('returnOnEquity') * 100:.2f}%"
                        if info.get("returnOnEquity") is not None
                        else "N/A"
                    ),
                    "Šaltinis / pastaba": "Yahoo Finance - returnOnEquity (%)",
                },
                {
                    "Rodiklis": "MGMT owns",
                    "Reikšmė": (
                        f"{mgmt_own_val * 100:.2f}%"
                        if isinstance(mgmt_own_val, (int, float))
                        else str(mgmt_own_val)
                    ),
                    "Šaltinis / pastaba": "Yahoo Finance - heldPercentInsiders (%)",
                },
                {
                    "Rodiklis": "Dividend Yield",
                    "Reikšmė": (
                        f"{div_yield_val:.2f}%"
                        if isinstance(div_yield_val, (int, float))
                        else str(div_yield_val)
                    ),
                    "Šaltinis / pastaba": "Yahoo Finance - dividendYield (%)",
                },
                {
                    "Rodiklis": "1y Target Est",
                    "Reikšmė": fmt_val(info.get("targetMeanPrice")),
                    "Šaltinis / pastaba": "Yahoo Finance - targetMeanPrice",
                },
            ]

            df_main = pd.DataFrame(main_metrics_data)

            st.dataframe(
                df_main.style.apply(
                    lambda _: style_main_metrics(df_main, price), axis=None
                ),
                use_container_width=True,
                hide_index=True,
            )

            st.divider()

            # 4 SEKCIJA: REVENUES, NET INCOME (GRAFIKAS SU DATA LABELS)
            st.subheader("📊 REVENUES, NET INCOME")

            has_annual_fin = bool(financials_years and revenues_vals and net_income_vals)
            has_quarterly_fin = bool(
                financials_periods_q and revenues_vals_q and net_income_vals_q
            )

            if has_annual_fin or has_quarterly_fin:
                fin_chart_placeholder = st.empty()
                fin_period_placeholder = st.container()

                fin_period_options = []
                if has_annual_fin:
                    fin_period_options.append("Metai")
                if has_quarterly_fin:
                    fin_period_options.append("Ketvirčiai")

                with fin_period_placeholder:
                    fin_period_choice = st.radio(
                        "Rodyti pagal:",
                        fin_period_options,
                        horizontal=True,
                        key="fin_period_choice",
                        label_visibility="collapsed",
                    )

                if fin_period_choice == "Ketvirčiai":
                    labels_fin = financials_periods_q
                    rev_fin = revenues_vals_q
                    ni_fin = net_income_vals_q
                    x_title = "Ketvirtis"
                else:
                    labels_fin = financials_years
                    rev_fin = revenues_vals
                    ni_fin = net_income_vals
                    x_title = "Metai"

                rev_b_formatted = [f"{v / 1e9:.1f}B" for v in rev_fin]
                ni_b_formatted = [f"{v / 1e9:.1f}B" for v in ni_fin]

                fig_rev_ni = go.Figure()
                fig_rev_ni.add_trace(
                    go.Bar(
                        x=labels_fin,
                        y=[v / 1e9 for v in rev_fin],
                        name="Revenues (USD mlrd.)",
                        marker_color="#38bdf8",
                        text=rev_b_formatted,
                        textposition="auto",
                    )
                )
                fig_rev_ni.add_trace(
                    go.Bar(
                        x=labels_fin,
                        y=[v / 1e9 for v in ni_fin],
                        name="Net Income (USD mlrd.)",
                        marker_color="#34d399",
                        text=ni_b_formatted,
                        textposition="auto",
                    )
                )
                fig_rev_ni.update_layout(
                    barmode="group",
                    paper_bgcolor="#0F172A",
                    plot_bgcolor="#0F172A",
                    font=dict(color="#FFFFFF"),
                    margin=dict(l=20, r=20, t=30, b=20),
                    xaxis=dict(
                        title=x_title, showgrid=True, gridcolor="#334155", color="#FFFFFF"
                    ),
                    yaxis=dict(
                        title="Suma (USD mlrd.)",
                        showgrid=True,
                        gridcolor="#334155",
                        color="#FFFFFF",
                    ),
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1,
                        font=dict(color="#FFFFFF"),
                    ),
                )
                with fin_chart_placeholder:
                    st.plotly_chart(fig_rev_ni, use_container_width=True)
            else:
                st.info("Finansinių ataskaitų duomenų pajamų ir pelno grafikui nerasta.")

            st.divider()

            # 5 SEKCIJA: INSIDER PREKYBA
            st.subheader("👥 INSIDER PREKYBA (paskutiniai 12 mėn.)")

            col_i1, col_i2, col_i3 = st.columns(3)
            col_i1.metric("Iš viso pirkta", format_number(total_buy, currency=cur))
            col_i2.metric("Iš viso parduota", format_number(total_sell, currency=cur))
            col_i3.metric("Grynasis srautas", format_number(net_flow, currency=cur))

            if insider_df is not None and not insider_df.empty:
                st.dataframe(
                    insider_df.style.apply(style_insider_df, axis=None),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("Insider prekybos duomenų už paskutinius 12 mėn. nerasta.")

            st.divider()

            # 6 SEKCIJA: MĖNESINĖ INSIDER APYVARTA VS KAINA
            st.subheader("📈 MĖNESINĖ INSIDER APYVARTA VS. KAINA")
            df_chart = pd.DataFrame()
            if not hist_5y.empty:
                hist_5y_naive = hist_5y.copy()
                hist_5y_naive.index = hist_5y_naive.index.tz_localize(None)
                end_dt = hist_5y_naive.index[-1]
                try:
                    m_ends = pd.date_range(end=end_dt, periods=13, freq="ME")
                except ValueError:
                    m_ends = pd.date_range(end=end_dt, periods=13, freq="M")

                chart_data = []
                for me in m_ends:
                    sub_p = hist_5y_naive[hist_5y_naive.index <= me]
                    p_at = float(sub_p["Close"].iloc[-1]) if not sub_p.empty else 0.0
                    m_str = me.strftime("%Y-%m")
                    buy_at = monthly_insider_buy.get(m_str, 0.0)
                    sell_at = monthly_insider_sell.get(m_str, 0.0)
                    chart_data.append(
                        {
                            "Mėnuo": m_str,
                            "Kaina (USD)": round(p_at, 2),
                            "Insider pirkimai (USD)": round(buy_at, 2),
                            "Insider pardavimai (USD)": round(sell_at, 2),
                        }
                    )

                df_chart = pd.DataFrame(chart_data)

                fig = make_subplots(specs=[[{"secondary_y": True}]])
                fig.add_trace(
                    go.Bar(
                        x=df_chart["Mėnuo"],
                        y=df_chart["Insider pirkimai (USD)"],
                        name=f"Insider pirkimai ({cur})",
                        marker_color="#34d399",
                        opacity=0.9,
                    ),
                    secondary_y=False,
                )
                fig.add_trace(
                    go.Bar(
                        x=df_chart["Mėnuo"],
                        y=df_chart["Insider pardavimai (USD)"],
                        name=f"Insider pardavimai ({cur})",
                        marker_color="#ef4444",
                        opacity=0.9,
                    ),
                    secondary_y=False,
                )
                fig.add_trace(
                    go.Scatter(
                        x=df_chart["Mėnuo"],
                        y=df_chart["Kaina (USD)"],
                        name=f"Akcijos Kaina ({cur})",
                        mode="lines+markers",
                        line=dict(color="#38bdf8", width=3),
                        marker=dict(size=6),
                    ),
                    secondary_y=True,
                )
                fig.update_layout(
                    barmode="group",
                    paper_bgcolor="#0F172A",
                    plot_bgcolor="#0F172A",
                    font=dict(color="#FFFFFF"),
                    margin=dict(l=20, r=20, t=30, b=20),
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1,
                        font=dict(color="#FFFFFF"),
                    ),
                )
                st.plotly_chart(fig, use_container_width=True)

            st.divider()

            # 7 SEKCIJA: ANALITIKŲ REKOMENDACIJA
            st.subheader("🎯 ANALITIKŲ REKOMENDACIJA")

            rec_key = (info.get("recommendationKey") or "").lower()
            rec_lt = REC_MAP.get(rec_key, rec_key.upper() if rec_key else "N/A")
            rec_mean_val = info.get("recommendationMean")
            analyst_count_val = info.get("numberOfAnalystOpinions")

            col_rec1, col_rec2, col_rec3 = st.columns(3)
            col_rec1.metric("Rekomendacija", rec_lt)
            col_rec2.metric(
                "Vidutinis balas",
                f"{rec_mean_val:.2f}" if rec_mean_val else "N/A",
            )
            col_rec3.metric(
                "Analitikų skaičius",
                str(analyst_count_val) if analyst_count_val else "N/A",
            )

            st.divider()

            # 8 SEKCIJA: PALYGINIMAS SU KITAIS TICKERIAIS (jei įvesti)
            if compare_tickers:
                st.subheader("🆚 PALYGINIMAS SU KITAIS TICKERIAIS")

                all_compare = [ticker_input] + compare_tickers
                compare_rows = []
                normalized_frames = {}
                invalid_compare = []

                with st.spinner("Renkami palyginimo duomenys..."):
                    for tk in all_compare:
                        if not validate_ticker(tk):
                            invalid_compare.append(tk)
                            continue
                        try:
                            c_info = fetch_info(tk)
                            c_hist = fetch_history(tk, selected_period)
                            if c_hist is None or c_hist.empty:
                                invalid_compare.append(tk)
                                continue

                            c_price = (
                                c_info.get("currentPrice")
                                or c_info.get("regularMarketPrice")
                                or float(c_hist["Close"].iloc[-1])
                            )
                            c_rec_key = (c_info.get("recommendationKey") or "").lower()
                            c_rec = REC_MAP.get(c_rec_key, c_rec_key.upper() if c_rec_key else "N/A")
                            c_div = compute_div_yield_pct(c_info)

                            compare_rows.append(
                                {
                                    "Tickeris": tk,
                                    "Pavadinimas": c_info.get("longName", tk),
                                    "Kaina": fmt_val(c_price),
                                    "Forward P/E": fmt_val(c_info.get("forwardPE")),
                                    "ROE %": (
                                        f"{c_info.get('returnOnEquity') * 100:.2f}"
                                        if c_info.get("returnOnEquity")
                                        else "N/A"
                                    ),
                                    "Div. Yield %": (
                                        f"{c_div:.2f}" if c_div is not None else "N/A"
                                    ),
                                    "1y Target": fmt_val(c_info.get("targetMeanPrice")),
                                    "Rekomendacija": c_rec,
                                }
                            )
                            normalized_frames[tk] = c_hist["Close"] / c_hist["Close"].iloc[0] * 100
                        except Exception:
                            invalid_compare.append(tk)

                if normalized_frames:
                    colors = ["#38bdf8", "#34d399", "#f472b6", "#fbbf24", "#a78bfa", "#f87171", "#22d3ee"]
                    fig_cmp = go.Figure()
                    for i, (tk, series) in enumerate(normalized_frames.items()):
                        fig_cmp.add_trace(
                            go.Scatter(
                                x=series.index,
                                y=series.values,
                                mode="lines",
                                name=tk,
                                line=dict(color=colors[i % len(colors)], width=2),
                            )
                        )
                    fig_cmp.update_layout(
                        title=f"Normalizuota kaina (bazė=100) — {period_label}",
                        paper_bgcolor="#0F172A",
                        plot_bgcolor="#0F172A",
                        font=dict(color="#FFFFFF"),
                        margin=dict(l=20, r=20, t=40, b=20),
                        xaxis=dict(showgrid=True, gridcolor="#334155", color="#FFFFFF"),
                        yaxis=dict(
                            showgrid=True,
                            gridcolor="#334155",
                            title="Indeksas (100 = laikotarpio pradžia)",
                            color="#FFFFFF",
                        ),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1,
                            font=dict(color="#FFFFFF"),
                        ),
                    )
                    st.plotly_chart(fig_cmp, use_container_width=True)

                if compare_rows:
                    df_compare = pd.DataFrame(compare_rows)
                    st.dataframe(df_compare, use_container_width=True, hide_index=True)

                if invalid_compare:
                    st.warning(
                        f"⚠️ Nepavyko rasti duomenų šiems tickeriams: {', '.join(invalid_compare)}"
                    )

                st.divider()

            # ------------------------------------------------------------------------------
            # 10. EXCEL EKSPORTAS NAUDOJANT OPENPYXL
            # ------------------------------------------------------------------------------
            output = BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            sheet_name_clean = ticker_input[:31]
            ws.title = sheet_name_clean

            ws.views.sheetView[0].showGridLines = True

            font_family = "Arial"
            title_font = Font(name=font_family, size=14, bold=True, color="1E3A8A")
            subtitle_font = Font(name=font_family, size=9, italic=True, color="475569")
            section_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
            header_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
            regular_font = Font(name=font_family, size=10, color="0F172A")
            bold_font = Font(name=font_family, size=10, bold=True, color="0F172A")

            section_fill = PatternFill(
                start_color="1E3A8A", end_color="1E3A8A", fill_type="solid"
            )
            header_fill = PatternFill(
                start_color="334155", end_color="334155", fill_type="solid"
            )
            zebra_fill = PatternFill(
                start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"
            )

            thin_border_side = Side(border_style="thin", color="CBD5E1")
            border_all = Border(
                left=thin_border_side,
                right=thin_border_side,
                top=thin_border_side,
                bottom=thin_border_side,
            )

            current_row = 1

            ws.cell(
                row=current_row, column=1, value=f"{long_name} ({ticker_input})"
            ).font = title_font
            current_row += 1
            current_time_str = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
            ws.cell(
                row=current_row,
                column=1,
                value=(
                    f"Duomenys gauti: {current_time_str} | Šaltinis: Yahoo Finance |"
                    f" Valiuta: {cur}"
                ),
            ).font = subtitle_font
            current_row += 2

            def write_section_header(title_text, max_cols=6):
                global current_row
                ws.merge_cells(
                    start_row=current_row,
                    start_column=1,
                    end_row=current_row,
                    end_column=max_cols,
                )
                cell = ws.cell(row=current_row, column=1, value=title_text)
                cell.font = section_font
                cell.fill = section_fill
                cell.alignment = Alignment(horizontal="left", vertical="center")
                current_row += 1

            def write_table_headers(headers):
                global current_row
                for col_idx, h in enumerate(headers, 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=h)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(
                        horizontal="center" if col_idx > 1 else "left", vertical="center"
                    )
                    cell.border = border_all
                current_row += 1

            # 1. PAGRINDINIAI RODIKLIAI
            write_section_header("PAGRINDINIAI RODIKLIAI")
            write_table_headers(["Rodiklis", "Reikšmė", "Šaltinis / pastaba"])

            for idx, r in df_main.iterrows():
                c1 = ws.cell(row=current_row, column=1, value=r["Rodiklis"])
                c2 = ws.cell(row=current_row, column=2, value=r["Reikšmė"])
                c3 = ws.cell(row=current_row, column=3, value=r["Šaltinis / pastaba"])

                for c in [c1, c2, c3]:
                    c.font = regular_font
                    c.border = border_all
                    if idx % 2 == 1:
                        c.fill = zebra_fill
                c2.alignment = Alignment(horizontal="right")
                current_row += 1
            current_row += 1

            # 2. KAINOS POKYTIS
            write_section_header("KAINOS POKYTIS")
            write_table_headers(["Laikotarpis", "Kaina praeityje (USD)", "Pokytis %"])

            if not df_price.empty:
                for idx, r in df_price.iterrows():
                    c1 = ws.cell(row=current_row, column=1, value=r["Laikotarpis"])
                    c2 = ws.cell(
                        row=current_row, column=2, value=r["Kaina praeityje (USD)"]
                    )
                    c3 = ws.cell(row=current_row, column=3, value=r["Pokytis %"])

                    for c in [c1, c2, c3]:
                        c.font = regular_font
                        c.border = border_all
                        if idx % 2 == 1:
                            c.fill = zebra_fill
                    c2.alignment = Alignment(horizontal="right")
                    c3.alignment = Alignment(horizontal="right")
                    current_row += 1
            current_row += 1

            # 3. INSIDER PREKYBA (paskutiniai 12 mėn.)
            write_section_header("INSIDER PREKYBA (paskutiniai 12 mėn.)")
            insider_summary = [
                ["Iš viso pirkta", format_number(total_buy, currency=cur)],
                ["Iš viso parduota", format_number(total_sell, currency=cur)],
                ["Grynasis srautas", format_number(net_flow, currency=cur)],
            ]
            for idx, (lbl, val) in enumerate(insider_summary):
                c1 = ws.cell(row=current_row, column=1, value=lbl)
                c2 = ws.cell(row=current_row, column=2, value=val)
                c1.font = bold_font
                c2.font = bold_font
                c1.border = border_all
                c2.border = border_all
                c2.alignment = Alignment(horizontal="right")
                current_row += 1
            current_row += 1

            if insider_df is not None and not insider_df.empty:
                write_table_headers(["Data", "Asmuo", "Pareigos", "Sandoris", "Akcijos", "Suma"])
                for idx, r in insider_df.iterrows():
                    c1 = ws.cell(row=current_row, column=1, value=r["Data"])
                    c2 = ws.cell(row=current_row, column=2, value=r["Asmuo"])
                    c3 = ws.cell(row=current_row, column=3, value=r["Pareigos"])
                    c4 = ws.cell(row=current_row, column=4, value=r["Sandoris"])
                    c5 = ws.cell(row=current_row, column=5, value=r["Akcijos"])
                    c6 = ws.cell(row=current_row, column=6, value=r["Suma"])

                    for col_idx, c in enumerate([c1, c2, c3, c4, c5, c6], 1):
                        c.font = regular_font
                        c.border = border_all
                        if idx % 2 == 1:
                            c.fill = zebra_fill
                        if col_idx in [5, 6]:
                            c.alignment = Alignment(horizontal="right")
                    current_row += 1
            else:
                c = ws.cell(row=current_row, column=1, value="Insider prekybos duomenų už paskutinius 12 mėn. nerasta.")
                c.font = regular_font
                current_row += 1
            current_row += 1

            # 4. MĖNESINĖ INSIDER APYVARTA VS KAINA
            write_section_header("MĖNESINĖ INSIDER APYVARTA VS. KAINA")
            if not df_chart.empty:
                write_table_headers(
                    [
                        "Mėnuo",
                        f"Kaina ({cur})",
                        f"Insider pirkimai ({cur})",
                        f"Insider pardavimai ({cur})",
                    ]
                )
                for idx, r in df_chart.iterrows():
                    c1 = ws.cell(row=current_row, column=1, value=r["Mėnuo"])
                    c2 = ws.cell(row=current_row, column=2, value=r["Kaina (USD)"])
                    c3 = ws.cell(row=current_row, column=3, value=r["Insider pirkimai (USD)"])
                    c4 = ws.cell(row=current_row, column=4, value=r["Insider pardavimai (USD)"])

                    for col_idx, c in enumerate([c1, c2, c3, c4], 1):
                        c.font = regular_font
                        c.border = border_all
                        if idx % 2 == 1:
                            c.fill = zebra_fill
                        if col_idx in [2, 3, 4]:
                            c.alignment = Alignment(horizontal="right")
                    current_row += 1
            else:
                c = ws.cell(row=current_row, column=1, value="Mėnesinių duomenų nerasta.")
                c.font = regular_font
                current_row += 1
            current_row += 1

            # 5. ANALITIKŲ REKOMENDACIJA
            write_section_header("ANALITIKŲ REKOMENDACIJA")
            write_table_headers(["Rodiklis / Aprašymas", "Reikšmė", "Pastaba"])

            rec_rows = [
                [
                    "Rekomendacija (Buy/Hold/Sell)",
                    rec_lt,
                    "Yahoo Finance - recommendationKey",
                ],
                [
                    "Vidutinis balas (1=Strong Buy, 5=Strong Sell)",
                    f"{rec_mean_val:.2f}" if rec_mean_val else "N/A",
                    "Yahoo Finance - recommendationMean",
                ],
                [
                    "Analitikų skaičius",
                    analyst_count_val if analyst_count_val else "N/A",
                    "Yahoo Finance - numberOfAnalystOpinions",
                ],
            ]
            for idx, row_data in enumerate(rec_rows):
                for col_idx, val in enumerate(row_data, 1):
                    c = ws.cell(row=current_row, column=col_idx, value=val)
                    c.font = regular_font
                    c.border = border_all
                    if idx % 2 == 1:
                        c.fill = zebra_fill
                    if col_idx == 2:
                        c.alignment = Alignment(horizontal="right")
                current_row += 1

            # 6. PALYGINIMAS SU KITAIS TICKERIAIS (jei taikoma)
            if compare_tickers and "compare_rows" in dir() and compare_rows:
                current_row += 1
                write_section_header("PALYGINIMAS SU KITAIS TICKERIAIS")
                comp_headers = list(compare_rows[0].keys())
                write_table_headers(comp_headers)
                for idx, row_dict in enumerate(compare_rows):
                    for col_idx, key in enumerate(comp_headers, 1):
                        c = ws.cell(row=current_row, column=col_idx, value=row_dict[key])
                        c.font = regular_font
                        c.border = border_all
                        if idx % 2 == 1:
                            c.fill = zebra_fill
                        if col_idx > 2:
                            c.alignment = Alignment(horizontal="right")
                    current_row += 1

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 5, 15)

            wb.save(output)
            output.seek(0)

            sidebar_download_placeholder.download_button(
                label="📥 Parsisiųsti Excel",
                data=output,
                file_name=f"{ticker_input}_analize_{dt.date.today()}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )

        except Exception as e:
            st.error(f"Klaida apdorojant duomenis: {e}")
