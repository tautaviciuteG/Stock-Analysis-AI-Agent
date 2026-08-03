import datetime as dt
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import requests
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
    .stApp { background-color: #FFFFFF !important; color: #0F172A !important; }
    [data-testid="stSidebar"] { background-color: #F8FAFC !important; border-right: 1px solid #E2E8F0 !important; }
    p, span, label, [data-testid="stWidgetLabel"] p { color: #0F172A !important; }
    label, [data-testid="stWidgetLabel"] { font-weight: 600 !important; margin-bottom: 4px !important; }
    [data-testid="stMetricValue"] { font-size: 1.4rem !important; font-weight: 700 !important; color: #0F172A !important; }
    [data-testid="stMetricLabel"] p { color: #475569 !important; }
    h1, h2, h3, h4, h5, h6 { color: #0F172A !important; font-weight: 700 !important; }
    .stTextInput input { background-color: #FFFFFF !important; color: #0F172A !important; border: 1.5px solid #0F172A !important; border-radius: 10px !important; }
    div.stDownloadButton > button { background-color: #F1F5F9 !important; color: #0F172A !important; border: 1px solid #E2E8F0 !important; border-radius: 12px !important; font-weight: 600 !important; width: 100% !important; padding: 0.5rem 1rem !important; }
    div.stDownloadButton > button:hover { background-color: #E2E8F0 !important; border-color: #CBD5E1 !important; }
</style>
""",
    unsafe_allow_html=True,
)

st.title("📊 Akcijų Analizės Agentas")

# ------------------------------------------------------------------------------
# 2. ŠONINĖ JUOSTA
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

# ------------------------------------------------------------------------------
# AI KONFIGŪRACIJA (Mistral AI - europietiškas, OpenAI-suderinamas API)
# ------------------------------------------------------------------------------
st.sidebar.divider()
st.sidebar.header("🤖 AI Agentas (Mistral AI)")
api_key = st.sidebar.text_input("Įveskite Mistral API Raktą", type="password")

# Užfiksuotas vienintelis modelis (be galimybės keisti sąsajoje)
SELECTED_MODEL = "mistral-small-latest"

st.sidebar.markdown(
    """
🔑 **Kaip gauti API raktą?**  
Nemokamą Mistral API raktą galite susikurti per [console.mistral.ai](https://console.mistral.ai/) 

🔒 **Privatumas ir Saugumas:**  
Jūsų įvestas API raktas ir užklausos **niekur nėra kaupiami ar išsaugomi šioje programoje**. 
Tačiau  Mistral gali naudoti jūsų užklausas modelio tobulinimui.

⚠️ **Apie klaidas:** nemokamas planas turi maždaug ~1 mlrd. žetonų/mėn., bet griežtą 
per-sekundę limitą (~1-2 užklausos/sek.). Jei gaunate 429 klaidą, tiesiog palaukite kelias sekundes.
"""
)

tickers_list = [x.strip().upper() for x in tickers_raw.split(",") if x.strip()]
ticker_input = tickers_list[0] if tickers_list else ""
compare_tickers = tickers_list[1:] if len(tickers_list) > 1 else []


# ------------------------------------------------------------------------------
# PAGALBINĖS FUNKCIJOS
# ------------------------------------------------------------------------------
def query_ai(prompt: str, system_prompt: str = "") -> str:
    """Išsiunčia užklausą į Mistral AI chat completions API.

    Naudoja pokalbio istoriją iš st.session_state, kad AI atsimintų ankstesnius
    klausimus tos pačios sesijos metu. Be automatinio retry - viena klaida iškart
    parodoma vartotojui su aiškiu paaiškinimu.
    """
    if not api_key:
        return "⚠️ Šoninėje juostoje įveskite savo **Mistral API Raktą**."

    url = "https://api.mistral.ai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
    }

    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})

    # Pridedame pokalbio istoriją iš sesijos būsenos - RAKTAS PRIKLAUSO NUO
    # DABARTINIO TICKERIO, kad skirtingų akcijų pokalbiai nesimaišytų
    chat_key = f"chat_messages_{ticker_input}"
    if chat_key in st.session_state:
        for msg in st.session_state[chat_key]:
            messages.append({"role": msg["role"], "content": msg["content"]})

    # Pridedame dabartinį klausimą, jei jo dar nėra sąrašo gale
    if not messages or messages[-1]["content"] != prompt:
        messages.append({"role": "user", "content": prompt})

    payload = {
        "model": SELECTED_MODEL,
        "messages": messages,
    }

    try:
        res = requests.post(url, headers=headers, json=payload, timeout=45)

        if res.status_code == 200:
            data = res.json()
            choices = data.get("choices", [])
            if choices:
                return choices[0].get("message", {}).get("content", "Nepavyko sugeneruoti atsakymo.")
            return "Nepavyko sugeneruoti atsakymo."

        elif res.status_code == 401:
            return (
                "⚠️ **Neteisingas Mistral API raktas (401).**\n\n"
                "Patikrinkite, ar raktas nukopijuotas tiksliai (be tarpų ar papildomų simbolių) "
                "iš [console.mistral.ai](https://console.mistral.ai/api-keys)."
            )
        elif res.status_code == 429:
            return (
                f"⚠️ **Viršytas „{SELECTED_MODEL}\" užklausų limitas (429).**\n\n"
                "Mistral nemokamas planas turi griežtą per-sekundę limitą "
                "(paprastai 1-2 užklausos/sek.). Palaukite kelias sekundes ir bandykite dar kartą. "
                "Jei tai kartojasi nuolat, patikrinkite savo limitus "
                "[console.mistral.ai](https://console.mistral.ai/) → Limits skiltyje."
            )
        elif res.status_code == 422:
            return (
                f"⚠️ **Netinkama užklausa (422).**\n\n"
                f"Galimai modelis „{SELECTED_MODEL}\" neegzistuoja arba yra rašybos klaida. "
                f"Atsakymas: {res.text}"
            )
        else:
            return f"⚠️ API Klaida ({res.status_code}): {res.text}"
    except Exception as e:
        return f"❌ Nepavyko pasiekti Mistral API. Klaida: {e}"


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


def fmt_unix_date(v):
    """Konvertuoja yfinance unix timestamp (sekundėmis) į skaitomą datą."""
    if v is None or pd.isna(v):
        return "N/A"
    try:
        return pd.to_datetime(int(v), unit="s").strftime("%Y-%m-%d")
    except (ValueError, TypeError, OSError):
        return "N/A"


def compute_div_yield_pct(info):
    dy = info.get("dividendYield")
    if dy is None or pd.isna(dy):
        return None
    dy = float(dy)
    return dy * 100 if dy < 0.1 else dy


def quarter_label(col):
    ts = pd.to_datetime(col)
    q = (ts.month - 1) // 3 + 1
    return f"{ts.year} Q{q}"


def extract_financials_series(inc_df, label_func, max_cols=4):
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
        r_val = (
            r_series.loc[col]
            if r_series is not None and col in r_series.index
            else None
        )
        ni_val = (
            ni_series.loc[col]
            if ni_series is not None and col in ni_series.index
            else None
        )
        if (r_val is not None and pd.notna(r_val)) or (
                ni_val is not None and pd.notna(ni_val)
        ):
            cols_with_data.append(col)

    recent_cols = (
        cols_with_data[-max_cols:]
        if len(cols_with_data) >= max_cols
        else cols_with_data
    )
    for col in recent_cols:
        labels.append(label_func(col))
        r_val = (
            r_series.loc[col]
            if r_series is not None and col in r_series.index
            else 0
        )
        ni_val = (
            ni_series.loc[col]
            if ni_series is not None and col in ni_series.index
            else 0
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
            elif (
                    metric == "1y Target Est"
                    and price_val is not None
                    and num > price_val
            ):
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
# DUOMENŲ TRAUKIMO FUNKCIJOS
# ------------------------------------------------------------------------------
@st.cache_data(ttl=600, show_spinner=False)
def validate_ticker(ticker: str) -> bool:
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
# 3. DUOMENŲ TRAUKIMAS IR APDOROJIMAS
# ------------------------------------------------------------------------------
if ticker_input:

    if not validate_ticker(ticker_input):
        st.error(
            f"❌ Tickeris **{ticker_input}** nerastas arba šiuo metu neturi rinkos duomenų."
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

            financials_years, revenues_vals, net_income_vals = [], [], []
            financials_periods_q, revenues_vals_q, net_income_vals_q = (
                [],
                [],
                [],
            )

            try:
                inc = fetch_income_stmt(ticker_input)
                if inc is not None and not inc.empty:
                    financials_years, revenues_vals, net_income_vals = (
                        extract_financials_series(
                            inc,
                            lambda c: pd.to_datetime(c).strftime("%Y"),
                            max_cols=4,
                        )
                    )
                    valid_cols = sorted([c for c in inc.columns if pd.notna(c)])
                    if len(valid_cols) >= 4:
                        col_latest, col_base = valid_cols[-1], valid_cols[-4]
                        for cand in ["Diluted EPS", "Basic EPS"]:
                            if cand in inc.index:
                                eps_growth_3y = calc_cagr(
                                    inc.loc[cand, col_base],
                                    inc.loc[cand, col_latest],
                                    3,
                                )
                                break
                        for cand in ["Total Revenue", "Operating Revenue"]:
                            if cand in inc.index:
                                sales_growth_3y = calc_cagr(
                                    inc.loc[cand, col_base],
                                    inc.loc[cand, col_latest],
                                    3,
                                )
                                break
            except Exception:
                pass

            try:
                inc_q = fetch_quarterly_income_stmt(ticker_input)
                if inc_q is not None and not inc_q.empty:
                    financials_periods_q, revenues_vals_q, net_income_vals_q = (
                        extract_financials_series(
                            inc_q, quarter_label, max_cols=8
                        )
                    )
            except Exception:
                pass

            insider_df = None
            total_buy, total_sell, net_flow = 0.0, 0.0, 0.0
            ins_12m_raw = pd.DataFrame()
            monthly_insiders = pd.DataFrame()

            try:
                ins = fetch_insider_transactions(ticker_input)
                if ins is not None and not ins.empty:
                    ins = ins.copy()
                    date_col = next(
                        (
                            c
                            for c in [
                            "Start Date",
                            "Date",
                            "Transaction Date",
                        ]
                            if c in ins.columns
                        ),
                        None,
                    )
                    if date_col:
                        ins[date_col] = pd.to_datetime(
                            ins[date_col], errors="coerce"
                        )
                        ins = ins.dropna(subset=[date_col])
                        cutoff = pd.Timestamp.now() - pd.Timedelta(days=365)
                        ins_12m = ins[ins[date_col] >= cutoff].sort_values(
                            date_col, ascending=False
                        )
                        ins_12m_raw = ins_12m.copy()

                        if not ins_12m.empty:
                            rows = []
                            for _, r in ins_12m.iterrows():
                                val = (
                                    float(r.get("Value"))
                                    if pd.notna(r.get("Value"))
                                    else 0.0
                                )
                                txt = str(
                                    r.get("Transaction")
                                    or r.get("Text")
                                    or ""
                                )
                                tx_type = classify_insider(txt)

                                if tx_type == "Pirkimas":
                                    total_buy += val
                                elif tx_type == "Pardavimas":
                                    total_sell += val

                                rows.append(
                                    {
                                        "Data": r[date_col].strftime(
                                            "%Y-%m-%d"
                                        ),
                                        "Asmuo": r.get("Insider", "N/A"),
                                        "Pareigos": r.get("Position", ""),
                                        "Sandoris": txt,
                                        "Akcijos": (
                                            f"{r.get('Shares'):,.0f}"
                                            if pd.notna(r.get("Shares"))
                                            else "N/A"
                                        ),
                                        "Suma": (
                                            f"{val:,.2f}"
                                            if pd.notna(val)
                                            else "0.00"
                                        ),
                                        "TxType": tx_type,
                                        "RawValue": val,
                                        "Month": r[date_col].strftime("%Y-%m"),
                                    }
                                )
                            insider_df = pd.DataFrame(rows)
                            net_flow = total_buy - total_sell

                            if "Month" in insider_df.columns:
                                monthly_insiders = (
                                    insider_df.groupby(["Month", "TxType"])[
                                        "RawValue"
                                    ]
                                    .sum()
                                    .unstack(fill_value=0)
                                    .reset_index()
                                )
                                if "Pirkimas" not in monthly_insiders.columns:
                                    monthly_insiders["Pirkimas"] = 0.0
                                if "Pardavimas" not in monthly_insiders.columns:
                                    monthly_insiders["Pardavimas"] = 0.0
                                monthly_insiders = monthly_insiders.sort_values(
                                    "Month"
                                )

                            insider_df = insider_df.drop(
                                columns=["TxType", "RawValue", "Month"],
                                errors="ignore",
                            )
            except Exception:
                pass

            # 1 SEKCIJA: KAINOS ISTORIJA
            price_chg_pct, price_chg_abs = None, None
            if len(hist_5y) > 1:
                prev_close = float(hist_5y["Close"].iloc[-2])
                if prev_close:
                    price_chg_abs = price - prev_close
                    price_chg_pct = (price_chg_abs / prev_close) * 100

            # After/Before hours kaina (jei Yahoo Finance ją teikia)
            market_state = info.get("marketState", "")
            post_price = info.get("postMarketPrice")
            post_change = info.get("postMarketChange")
            post_change_pct = info.get("postMarketChangePercent")
            pre_price = info.get("preMarketPrice")
            pre_change = info.get("preMarketChange")
            pre_change_pct = info.get("preMarketChangePercent")

            extended_label, extended_price, extended_delta = None, None, None
            if post_price is not None and not pd.isna(post_price):
                extended_label = "After Hours"
                extended_price = post_price
                if post_change is not None and post_change_pct is not None:
                    extended_delta = f"{post_change:+.2f} ({post_change_pct:+.2f}%)"
            elif pre_price is not None and not pd.isna(pre_price):
                extended_label = "Before Hours (Pre-Market)"
                extended_price = pre_price
                if pre_change is not None and pre_change_pct is not None:
                    extended_delta = f"{pre_change:+.2f} ({pre_change_pct:+.2f}%)"

            if extended_label:
                header_col1, header_col2, header_col3 = st.columns([3, 1, 1])
            else:
                header_col1, header_col2 = st.columns([3, 1])
                header_col3 = None

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
            if extended_label and header_col3 is not None:
                with header_col3:
                    st.metric(
                        label=f"{extended_label} ({cur})",
                        value=f"{extended_price:,.2f}",
                        delta=extended_delta,
                    )

            chart_placeholder = st.empty()
            with st.container():
                period_label = st.radio(
                    "Kainos grafiko laikotarpis:",
                    list(PERIOD_OPTIONS.keys()),
                    index=3,
                    horizontal=True,
                    key="price_chart_period",
                    label_visibility="collapsed",
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
                legend=dict(font=dict(color="#FFFFFF")),
                margin=dict(l=20, r=20, t=40, b=20),
                xaxis=dict(
                    showgrid=True,
                    gridcolor="#334155",
                    title="Data",
                    color="#FFFFFF",
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
                    target_dt = (
                        pd.Timestamp(year=now_dt.year, month=1, day=1)
                        if ytd
                        else now_dt - pd.Timedelta(days=days_back)
                    )
                    sub = hist_5y_local[hist_5y_local.index <= target_dt]
                    return (
                        float(sub["Close"].iloc[-1])
                        if not sub.empty
                        else float(hist_5y_local["Close"].iloc[0])
                    )

                periods_data = [
                    ("5y", float(hist_5y_local["Close"].iloc[0])),
                    ("1y", get_historical_price(days_back=365)),
                    ("YTD", get_historical_price(ytd=True)),
                    ("6M", get_historical_price(days_back=182)),
                    ("1M", get_historical_price(days_back=30)),
                ]
                df_price = pd.DataFrame(
                    [
                        {
                            "Laikotarpis": label,
                            "Kaina praeityje (USD)": (
                                f"{past_p:.2f}"
                                if isinstance(past_p, (int, float))
                                else "N/A"
                            ),
                            "Pokytis %": (
                                f"{((curr_p - past_p) / past_p) * 100:+.2f}%"
                                if past_p and past_p > 0
                                else "N/A"
                            ),
                        }
                        for label, past_p in periods_data
                    ]
                )

                st.dataframe(
                    df_price.style.apply(style_price_change, axis=None),
                    use_container_width=True,
                    hide_index=True,
                )

            st.divider()

            # 3 SEKCIJA: PAGRINDINIAI RODIKLIAI
            st.subheader("PAGRINDINIAI RODIKLIAI")
            mgmt_own_val = info.get("heldPercentInsiders", "N/A")
            div_yield_val = compute_div_yield_pct(info) or "N/A"
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
                    "Šaltinis / pastaba": "Apskaičiuota CAGR",
                },
                {
                    "Rodiklis": "3-Year Sales Growth Rate",
                    "Reikšmė": (
                        f"{sales_growth_3y:.2f}%"
                        if sales_growth_3y is not None
                        else "N/A"
                    ),
                    "Šaltinis / pastaba": "Apskaičiuota CAGR",
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
                    "Šaltinis / pastaba": "Yahoo Finance - debtToEquity",
                },
                {
                    "Rodiklis": "Book Value Per Share",
                    "Reikšmė": fmt_val(info.get("bookValue")),
                    "Šaltinis / pastaba": "Yahoo Finance - bookValue",
                },
                {
                    "Rodiklis": "ROE",
                    "Reikšmė": (
                        f"{info.get('returnOnEquity') * 100:.2f}%"
                        if info.get("returnOnEquity") is not None
                        else "N/A"
                    ),
                    "Šaltinis / pastaba": "Yahoo Finance - ROE",
                },
                {
                    "Rodiklis": "MGMT owns",
                    "Reikšmė": (
                        f"{mgmt_own_val * 100:.2f}%"
                        if isinstance(mgmt_own_val, (int, float))
                        else str(mgmt_own_val)
                    ),
                    "Šaltinis / pastaba": "Yahoo Finance - heldPercentInsiders",
                },
                {
                    "Rodiklis": "Dividend Yield",
                    "Reikšmė": (
                        f"{div_yield_val:.2f}%"
                        if isinstance(div_yield_val, (int, float))
                        else str(div_yield_val)
                    ),
                    "Šaltinis / pastaba": "Yahoo Finance - dividendYield",
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

            # 4 SEKCIJA: REVENUES, NET INCOME
            st.subheader("📊 REVENUES, NET INCOME")
            has_annual_fin = bool(
                financials_years and revenues_vals and net_income_vals
            )
            has_quarterly_fin = bool(
                financials_periods_q
                and revenues_vals_q
                and net_income_vals_q
            )

            if has_annual_fin or has_quarterly_fin:
                fin_chart_placeholder = st.empty()
                fin_period_options = []
                if has_annual_fin:
                    fin_period_options.append("Metai")
                if has_quarterly_fin:
                    fin_period_options.append("Ketvirčiai")

                fin_period_choice = st.radio(
                    "Rodyti pagal:",
                    fin_period_options,
                    horizontal=True,
                    label_visibility="collapsed",
                )
                labels_fin, rev_fin, ni_fin, x_title = (
                    (
                        financials_periods_q,
                        revenues_vals_q,
                        net_income_vals_q,
                        "Ketvirtis",
                    )
                    if fin_period_choice == "Ketvirčiai"
                    else (
                        financials_years,
                        revenues_vals,
                        net_income_vals,
                        "Metai",
                    )
                )

                fig_rev_ni = go.Figure()
                fig_rev_ni.add_trace(
                    go.Bar(
                        x=labels_fin,
                        y=[v / 1e9 for v in rev_fin],
                        name="Revenues (USD mlrd.)",
                        marker_color="#38bdf8",
                        text=[f"{v / 1e9:.1f}B" for v in rev_fin],
                        textposition="auto",
                    )
                )
                fig_rev_ni.add_trace(
                    go.Bar(
                        x=labels_fin,
                        y=[v / 1e9 for v in ni_fin],
                        name="Net Income (USD mlrd.)",
                        marker_color="#34d399",
                        text=[f"{v / 1e9:.1f}B" for v in ni_fin],
                        textposition="auto",
                    )
                )
                fig_rev_ni.update_layout(
                    barmode="group",
                    paper_bgcolor="#0F172A",
                    plot_bgcolor="#0F172A",
                    font=dict(color="#FFFFFF"),
                    legend=dict(font=dict(color="#FFFFFF")),
                    margin=dict(l=20, r=20, t=30, b=20),
                    xaxis=dict(
                        title=x_title,
                        showgrid=True,
                        gridcolor="#334155",
                        color="#FFFFFF",
                    ),
                    yaxis=dict(
                        title="Suma (USD mlrd.)",
                        showgrid=True,
                        gridcolor="#334155",
                        color="#FFFFFF",
                    ),
                )
                with fin_chart_placeholder:
                    st.plotly_chart(fig_rev_ni, use_container_width=True)

            st.divider()

            # 5 SEKCIJA: INSIDER PREKYBA
            st.subheader("👥 INSIDER PREKYBA (paskutiniai 12 mėn.)")
            c1, c2, c3 = st.columns(3)
            c1.metric("Iš viso pirkta", format_number(total_buy, currency=cur))
            c2.metric(
                "Iš viso parduota", format_number(total_sell, currency=cur)
            )
            c3.metric("Grynasis srautas", format_number(net_flow, currency=cur))

            hist_1y = fetch_history(ticker_input, "1y")
            if not hist_1y.empty:
                fig_insider = make_subplots(specs=[[{"secondary_y": True}]])

                fig_insider.add_trace(
                    go.Scatter(
                        x=hist_1y.index,
                        y=hist_1y["Close"],
                        mode="lines",
                        name=f"Kaina ({cur})",
                        line=dict(color="#38bdf8", width=2.5),
                    ),
                    secondary_y=True,
                )

                if not ins_12m_raw.empty:
                    date_col_ins = next(
                        (
                            c
                            for c in [
                            "Start Date",
                            "Date",
                            "Transaction Date",
                        ]
                            if c in ins_12m_raw.columns
                        ),
                        None,
                    )
                    if date_col_ins:
                        hist_clean = hist_1y.copy()
                        hist_clean.index = hist_clean.index.tz_localize(None).normalize()

                        if (
                                not monthly_insiders.empty
                                and "Month" in monthly_insiders.columns
                        ):
                            fig_insider.add_trace(
                                go.Bar(
                                    x=monthly_insiders["Month"],
                                    y=monthly_insiders["Pirkimas"],
                                    name="Pirkta suma (stulpeliai)",
                                    marker_color="#10B981",
                                    opacity=0.6,
                                ),
                                secondary_y=False,
                            )
                            fig_insider.add_trace(
                                go.Bar(
                                    x=monthly_insiders["Month"],
                                    y=monthly_insiders["Pardavimas"],
                                    name="Parduota suma (stulpeliai)",
                                    marker_color="#EF4444",
                                    opacity=0.6,
                                ),
                                secondary_y=False,
                            )

                        buys_x, buys_y, buys_txt = [], [], []
                        sells_x, sells_y, sells_txt = [], [], []

                        for _, row in ins_12m_raw.iterrows():
                            tx_type = classify_insider(
                                row.get("Transaction") or row.get("Text")
                            )
                            dt_val = (
                                pd.to_datetime(row[date_col_ins])
                                .tz_localize(None)
                                .normalize()
                            )
                            ins_name = row.get("Insider", "Asmuo")
                            val_str = (
                                f"{row.get('Value'):,.0f} USD"
                                if pd.notna(row.get("Value"))
                                else "N/A"
                            )

                            match_idx = hist_clean.index.get_indexer(
                                [dt_val], method="nearest"
                            )[0]
                            if match_idx >= 0:
                                p_val = hist_clean["Close"].iloc[match_idx]
                                hover_text = f"<b>{ins_name}</b><br>Sandoris: {tx_type}<br>Suma: {val_str}"

                                if tx_type == "Pirkimas":
                                    buys_x.append(dt_val)
                                    buys_y.append(p_val)
                                    buys_txt.append(hover_text)
                                elif tx_type == "Pardavimas":
                                    sells_x.append(dt_val)
                                    sells_y.append(p_val)
                                    sells_txt.append(hover_text)

                        if buys_x:
                            fig_insider.add_trace(
                                go.Scatter(
                                    x=buys_x,
                                    y=buys_y,
                                    mode="markers",
                                    name="Pirkimo kaina",
                                    marker=dict(
                                        color="#10B981",
                                        size=11,
                                        symbol="triangle-up",
                                        line=dict(width=1, color="#FFFFFF"),
                                    ),
                                    hovertext=buys_txt,
                                    hoverinfo="text+x+y",
                                ),
                                secondary_y=True,
                            )
                        if sells_x:
                            fig_insider.add_trace(
                                go.Scatter(
                                    x=sells_x,
                                    y=sells_y,
                                    mode="markers",
                                    name="Pardavimo kaina",
                                    marker=dict(
                                        color="#EF4444",
                                        size=11,
                                        symbol="triangle-down",
                                        line=dict(width=1, color="#FFFFFF"),
                                    ),
                                    hovertext=sells_txt,
                                    hoverinfo="text+x+y",
                                ),
                                secondary_y=True,
                            )

                fig_insider.update_layout(
                    title="Insider prekybos apimtys (stulpeliai) ir akcijos kaina (linija)",
                    paper_bgcolor="#0F172A",
                    plot_bgcolor="#0F172A",
                    font=dict(color="#FFFFFF"),
                    legend=dict(font=dict(color="#FFFFFF")),
                    margin=dict(l=20, r=20, t=40, b=20),
                    barmode="group",
                )

                fig_insider.update_xaxes(
                    title="Data",
                    showgrid=True,
                    gridcolor="#334155",
                    color="#FFFFFF",
                )
                fig_insider.update_yaxes(
                    title_text="Sandorių suma ($/EUR)",
                    showgrid=True,
                    gridcolor="#334155",
                    color="#FFFFFF",
                    secondary_y=False,
                )
                fig_insider.update_yaxes(
                    title_text=f"Akcijos kaina ({cur})",
                    showgrid=False,
                    color="#38bdf8",
                    secondary_y=True,
                )

                st.plotly_chart(fig_insider, use_container_width=True)
            else:
                st.info("💡 Pastarųjų 12 mėnesių kainos grafiko nerasta.")

            if insider_df is not None and not insider_df.empty:
                st.dataframe(
                    insider_df.style.apply(style_insider_df, axis=None),
                    use_container_width=True,
                    hide_index=True,
                )

            st.divider()

            # 6 SEKCIJA: ANALITIKŲ REKOMENDACIJA
            st.subheader("🎯 ANALITIKŲ REKOMENDACIJA")
            rec_key = (info.get("recommendationKey") or "").lower()
            rec_lt = REC_MAP.get(
                rec_key, rec_key.upper() if rec_key else "N/A"
            )

            col_rec1, col_rec2, col_rec3 = st.columns(3)
            col_rec1.metric("Rekomendacija", rec_lt)
            col_rec2.metric(
                "Vidutinis balas",
                (
                    f"{info.get('recommendationMean'):.2f}"
                    if info.get("recommendationMean")
                    else "N/A"
                ),
            )
            col_rec3.metric(
                "Analitikų skaičius",
                (
                    str(info.get("numberOfAnalystOpinions"))
                    if info.get("numberOfAnalystOpinions")
                    else "N/A"
                ),
            )

            st.divider()

            # ------------------------------------------------------------------
            # PAPILDOMA SEKCIJA: TICKERIŲ PALYGINIMAS (Jei įvesti keli tickeriai)
            # ------------------------------------------------------------------
            if compare_tickers:
                st.subheader("⚖️ TICKERIŲ PALYGINIMAS")
                all_tickers_to_compare = [ticker_input] + compare_tickers
                comp_data = []

                fig_comp = go.Figure()

                for t_sym in all_tickers_to_compare:
                    if validate_ticker(t_sym):
                        t_info = fetch_info(t_sym)
                        t_hist = fetch_history(t_sym, "1y")

                        t_price = t_info.get("currentPrice") or t_info.get(
                            "regularMarketPrice"
                        )
                        if (
                                t_price is None or pd.isna(t_price)
                        ) and not t_hist.empty:
                            t_price = float(t_hist["Close"].iloc[-1])

                        if not t_hist.empty:
                            start_p = t_hist["Close"].iloc[0]
                            pct_series = (
                                                 (t_hist["Close"] - start_p) / start_p
                                         ) * 100
                            fig_comp.add_trace(
                                go.Scatter(
                                    x=t_hist.index,
                                    y=pct_series,
                                    mode="lines",
                                    name=t_sym,
                                )
                            )

                        t_dy = compute_div_yield_pct(t_info)
                        comp_data.append(
                            {
                                "Tickeris": t_sym,
                                "Pavadinimas": t_info.get("shortName", t_sym),
                                "Kaina": fmt_val(t_price),
                                "Forward P/E": fmt_val(t_info.get("forwardPE")),
                                "ROE": (
                                    f"{t_info.get('returnOnEquity') * 100:.2f}%"
                                    if t_info.get("returnOnEquity")
                                    else "N/A"
                                ),
                                "Div. Yield": (
                                    f"{t_dy:.2f}%" if t_dy else "N/A"
                                ),
                                "Rekomendacija": t_info.get(
                                    "recommendationKey", "N/A"
                                ).upper(),
                            }
                        )

                if comp_data:
                    st.dataframe(
                        pd.DataFrame(comp_data),
                        use_container_width=True,
                        hide_index=True,
                    )

                fig_comp.update_layout(
                    title="1 metų kainos pokytis (%) palyginimas",
                    paper_bgcolor="#0F172A",
                    plot_bgcolor="#0F172A",
                    font=dict(color="#FFFFFF"),
                    legend=dict(font=dict(color="#FFFFFF")),
                    xaxis=dict(
                        showgrid=True, gridcolor="#334155", color="#FFFFFF"
                    ),
                    yaxis=dict(
                        title="Pokytis (%)",
                        showgrid=True,
                        gridcolor="#334155",
                        color="#FFFFFF",
                    ),
                )
                st.plotly_chart(fig_comp, use_container_width=True)
                st.divider()

            # 7 SEKCIJA: AI AGENTO ANALIZĖ IR CHAT
            st.subheader("🤖 AI AGENTO ANALIZĖ IR CHAT")
            st.caption(f"Naudojamas modelis: `{SELECTED_MODEL}` (Mistral AI)")

            context_summary = f"""
            ĮMONĖ: {long_name} ({ticker_input})
            Dabartinė kaina: {price} {cur}
            Analitikų rekomendacija: {rec_lt}

            RODIKLIAI:
            - Forward P/E: {info.get('forwardPE')}
            - 3-Yr Sales Growth: {sales_growth_3y}%
            - 3-Yr EPS Growth: {eps_growth_3y}%
            - ROE: {info.get('returnOnEquity')}
            - Debt to Equity: {d_e}
            - 1-Yr Target Price: {info.get('targetMeanPrice')}
            - Dividend Yield: {div_yield_val}%

            DIVIDENDŲ DUOMENYS (iš Yahoo Finance, jei įmonė moka dividendus):
            - Paskutinio dividendo suma: {info.get('lastDividendValue', 'N/A')}
            - Paskutinio dividendo data: {fmt_unix_date(info.get('lastDividendDate'))}
            - Ex-dividend data (paskutinė žinoma): {fmt_unix_date(info.get('exDividendDate'))}
            - Metinė dividendo norma: {info.get('dividendRate', 'N/A')}
            - Payout ratio: {info.get('payoutRatio', 'N/A')}
            (Pastaba: šios datos yra PASKUTINĖS ŽINOMOS iš Yahoo Finance, o ne būtinai
            būsimo mokėjimo grafikas - tai nurodyk atsakyme, jei aktualu.)
            """

            col_ai1, col_ai2 = st.columns([1, 2])

            with col_ai1:
                st.markdown("### Generuoti įžvalgas")
                if st.button(
                        "🚀 Sugeneruoti AI Analizę", use_container_width=True
                ):
                    with st.spinner("Mistral AI analizuoja duomenis..."):
                        system_prompt = (
                            "Tu esi patyręs finansų analitikas, kalbantis lietuviškai. "
                            "Analizuok pateiktus duomenis ir pateik stiprybes, rizikas bei "
                            "apibendrinimą. Prireikus gali papildyti analizę savo bendrosiomis "
                            "žiniomis apie įmonę ar sektorių, aiškiai atskirdamas, kas yra "
                            "pateikti tikslūs duomenys, o kas - tavo bendra žinia."
                        )
                        user_prompt = f"Remdamasis šiais duomenimis, pateik trumpą įmonės apžvalgą:\n\n{context_summary}"
                        st.session_state[f"ai_summary_{ticker_input}"] = (
                            query_ai(user_prompt, system_prompt)
                        )

            with col_ai2:
                if f"ai_summary_{ticker_input}" in st.session_state:
                    st.markdown("### AI Analizės Ataskaita")
                    st.info(st.session_state[f"ai_summary_{ticker_input}"])

            st.markdown("---")
            st.markdown("#### 💬 Užduokite klausimą agentui apie šią akciją")

            chat_key = f"chat_messages_{ticker_input}"
            if chat_key not in st.session_state:
                st.session_state[chat_key] = []

            for message in st.session_state[chat_key]:
                with st.chat_message(message["role"]):
                    st.markdown(message["content"])

            if user_query := st.chat_input(
                    f"Paklausk ko nors apie {ticker_input}..."
            ):
                st.session_state[chat_key].append(
                    {"role": "user", "content": user_query}
                )
                with st.chat_message("user"):
                    st.markdown(user_query)

                with st.chat_message("assistant"):
                    with st.spinner("Ieškoma informacijos..."):
                        chat_system_prompt = (
                            "Tu esi patyręs finansų analitikas ir bendro pobūdžio AI asistentas, "
                            "kalbantis lietuviškai. Atsakinėk į BET KOKIUS vartotojo klausimus - "
                            "tiek apie šią konkrečią akciją, tiek bendro pobūdžio finansų, "
                            "investavimo ar kitas temas - naudodamasis savo bendrosiomis žiniomis, "
                            "lygiai kaip įprastas pokalbių asistentas (pvz. ChatGPT ar Gemini).\n\n"
                            f"Apie šią konkrečią akciją turi šiuos šviežius duomenis iš Yahoo Finance:\n"
                            f"{context_summary}\n\n"
                            "Kai klausimas susijęs su šia akcija - PIRMIAUSIA naudok aukščiau "
                            "pateiktus duomenis. Jei jų trūksta konkrečiam faktui (pvz. tikslios "
                            "būsimos datos), papildyk atsakymą savo bendrosiomis žiniomis, bet aiškiai "
                            "nurodyk, kad tai bendra informacija, o ne patvirtintas Yahoo Finance faktas. "
                            "Kai klausimas nesusijęs su šia akcija - atsakyk laisvai, kaip į bet kokį "
                            "kitą klausimą, remdamasis savo žiniomis. Neatsisakyk atsakyti vien todėl, "
                            "kad tikslaus fakto nėra pateiktuose duomenyse - pasakyk, ką žinai, ir "
                            "pažymėk, kur reikėtų patikrinti naujausią informaciją (pvz. jei tavo "
                            "žinios gali būti pasenusios)."
                        )
                        full_res = query_ai(user_query, chat_system_prompt)
                        st.markdown(full_res)
                        st.session_state[chat_key].append(
                            {"role": "assistant", "content": full_res}
                        )

        except Exception as e:
            st.error(f"Klaida apdorojant duomenis: {e}")