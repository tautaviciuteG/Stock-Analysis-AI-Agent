import datetime as dt
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
    /* Pagrindinio puslapio fonas ir tekstas */
    .stApp {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
    }

    /* Šoninė juosta */
    [data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 1px solid #E2E8F0 !important;
    }

    /* Įvedimo laukelių pavadinimai */
    p, span, label, [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p {
        color: #0F172A !important;
    }

    label, [data-testid="stWidgetLabel"] {
        font-weight: 600 !important;
        margin-bottom: 4px !important;
    }

    /* Metrikų kortelės */
    [data-testid="stMetricValue"] {
        font-size: 1.4rem !important;
        font-weight: 700 !important;
        color: #0F172A !important;
    }
    [data-testid="stMetricLabel"] p {
        color: #475569 !important;
    }

    /* Antraštės */
    h1, h2, h3, h4, h5, h6 {
        color: #0F172A !important;
        font-weight: 700 !important;
    }

    /* Tekstinis įvedimo laukas */
    .stTextInput input {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1.5px solid #0F172A !important;
        border-radius: 10px !important;
    }

    /* 1. EPS RATING (Number Input) */
    div[data-testid="stNumberInput"] > div > div {
        background-color: #FFFFFF !important;
        border: 1.5px solid #0F172A !important;
        border-radius: 10px !important;
        overflow: hidden !important;
    }
    div[data-testid="stNumberInput"] input {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        font-size: 16px !important;
    }
    div[data-testid="stNumberInput"] button {
        background-color: #E2E8F0 !important;
        color: #0F172A !important;
        border: none !important;
        border-left: 1px solid #CBD5E1 !important;
    }
    div[data-testid="stNumberInput"] button:hover {
        background-color: #CBD5E1 !important;
    }
    div[data-testid="stNumberInput"] button svg {
        fill: #0F172A !important;
        color: #0F172A !important;
    }

    /* 2. SMR RATING (Selectbox) */
    div[data-testid="stSelectbox"] > div > div {
        background-color: #FFFFFF !important;
        border: 1.5px solid #0F172A !important;
        border-radius: 10px !important;
        overflow: hidden !important;
    }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        border-radius: 8px !important;
    }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: none !important;
    }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div > div:last-child {
        background-color: #E2E8F0 !important;
        border-left: 1px solid #CBD5E1 !important;
        padding-left: 10px !important;
        padding-right: 10px !important;
    }
    div[data-testid="stSelectbox"] svg {
        fill: #0F172A !important;
        color: #0F172A !important;
    }

    /* Išskleidžiamo meniu (Dropdown) stiliai */
    div[data-baseweb="popover"] ul {
        background-color: #FFFFFF !important;
        border: 1.5px solid #0F172A !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="popover"] li {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
    }
    div[data-baseweb="popover"] li:hover,
    div[data-baseweb="popover"] li[aria-selected="true"] {
        background-color: #E2E8F0 !important;
        color: #0F172A !important;
    }

    /* LENTELIŲ STILIAI */
    div[data-testid="stDataFrame"], .stDataFrame {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        padding: 4px !important;
    }
    div[data-testid="stDataFrame"] iframe {
        background-color: #FFFFFF !important;
    }
    .stTable table {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #E2E8F0 !important;
        border-collapse: collapse !important;
    }
    .stTable th {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        font-weight: 600 !important;
        border: 1px solid #E2E8F0 !important;
    }
    .stTable td {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #E2E8F0 !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

st.title("📊 Akcijų Analizės Agentas")

# ------------------------------------------------------------------------------
# 2. ŠONINĖ JUOSTA (PARAMETRAI)
# ------------------------------------------------------------------------------
st.sidebar.header("⚙️ Parametrai")
ticker_input = (
    st.sidebar.text_input("Įveskite akcijos tickerį:", "GOOGL").strip().upper()
)

st.sidebar.subheader("✍️ Rankinis IBD įvedimas")
eps_rating_input = st.sidebar.number_input(
    "EPS rating (1-99 balai):", min_value=0, max_value=99, value=0, step=1
)
smr_rating_input = st.sidebar.selectbox(
    "SMR Rating (A-E balai):", ["-", "A", "B", "C", "D", "E"]
)


# Pagalbinės funkcijos
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
# 3. DUOMENŲ TRAUKIMAS IR APDOROJIMAS
# ------------------------------------------------------------------------------
if ticker_input:
  with st.spinner(f"Renkami ir skaičiuojami {ticker_input} duomenys..."):
    try:
      t = yf.Ticker(ticker_input)
      info = t.info or {}

      hist_5y = t.history(period="5y")

      price = info.get("currentPrice") or info.get("regularMarketPrice")
      if (price is None or pd.isna(price)) and not hist_5y.empty:
        price = float(hist_5y["Close"].iloc[-1])

      if price is None or hist_5y.empty:
        st.error("Nepavyko rasti rinkos duomenų šiam tickeriui.")
        st.stop()

      cur = info.get("currency", "USD")
      long_name = info.get("longName", ticker_input)

      # FINANCIALS (CAGR)
      eps_growth_3y = None
      sales_growth_3y = None

      try:
        inc = t.income_stmt
        if inc is not None and not inc.empty:
          valid_cols = sorted([c for c in inc.columns if pd.notna(c)])
          if len(valid_cols) >= 4:
            col_latest = valid_cols[-1]
            col_base = valid_cols[-4]

            for cand in ["Diluted EPS", "Basic EPS"]:
              if cand in inc.index:
                b_val = inc.loc[cand, col_base]
                l_val = inc.loc[cand, col_latest]
                eps_growth_3y = calc_cagr(b_val, l_val, 3)
                break

            for cand in ["Total Revenue", "Operating Revenue"]:
              if cand in inc.index:
                b_val = inc.loc[cand, col_base]
                l_val = inc.loc[cand, col_latest]
                sales_growth_3y = calc_cagr(b_val, l_val, 3)
                break
      except Exception:
        pass

      # 1 SEKCIJA: AKCIJOS KAINOS ISTORIJA (GRAFIKAS)
      st.subheader(f"{long_name} ({ticker_input})")
      fig_price = go.Figure()
      fig_price.add_trace(
          go.Scatter(
              x=hist_5y.index,
              y=hist_5y["Close"],
              mode="lines",
              name=f"Kaina ({cur})",
              line=dict(color="#38bdf8", width=2),
          )
      )
      fig_price.update_layout(
          paper_bgcolor="#0F172A",
          plot_bgcolor="#0F172A",
          font=dict(color="#FFFFFF"),
          margin=dict(l=20, r=20, t=30, b=20),
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
      st.plotly_chart(fig_price, use_container_width=True)

      st.divider()

      # 2 SEKCIJA: KAINOS POKYTIS (NUO 1M IKI 5Y)
      st.subheader("📉 KAINOS POKYTIS")

      df_price = pd.DataFrame()
      if not hist_5y.empty and len(hist_5y) > 1:
        hist_5y.index = hist_5y.index.tz_localize(None)
        curr_p = price
        now_dt = hist_5y.index[-1]

        def get_historical_price(days_back=None, ytd=False):
          if ytd:
            target_dt = pd.Timestamp(year=now_dt.year, month=1, day=1)
          else:
            target_dt = now_dt - pd.Timedelta(days=days_back)
          sub = hist_5y[hist_5y.index <= target_dt]
          if not sub.empty:
            return float(sub["Close"].iloc[-1])
          return float(hist_5y["Close"].iloc[0])

        p_1m = get_historical_price(days_back=30)
        p_6m = get_historical_price(days_back=182)
        p_ytd = get_historical_price(ytd=True)
        p_1y = get_historical_price(days_back=365)
        p_5y = float(hist_5y["Close"].iloc[0])

        periods_data = [
            ("1M", p_1m),
            ("6M", p_6m),
            ("YTD", p_ytd),
            ("1y", p_1y),
            ("5y", p_5y),
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
                  "Kaina praeityje (USD)": f"{past_p:.2f}",
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
      st.subheader("📌 PAGRINDINIAI RODIKLIAI")

      mgmt_own = info.get("heldPercentInsiders")
      mgmt_own_str = f"{mgmt_own * 100:.2f}%" if mgmt_own is not None else "N/A"

      div_yield = info.get("dividendYield")
      if div_yield is not None:
        div_yield_val = div_yield * 100 if div_yield < 1 else div_yield
        div_yield_str = f"{div_yield_val:.2f}%"
      else:
        div_yield_str = "N/A"

      d_e = info.get("debtToEquity")
      d_e_str = f"{d_e:.2f}%" if d_e is not None else "N/A"

      main_metrics_data = [
          {
              "Rodiklis": "PRICE",
              "Reikšmė": f"{price:.2f} {cur}" if price else "N/A",
              "Šaltinis / pastaba": "Yahoo Finance - currentPrice",
          },
          {
              "Rodiklis": "Forward P/E",
              "Reikšmė": (
                  f"{info.get('forwardPE'):.2f}"
                  if info.get("forwardPE")
                  else "N/A"
              ),
              "Šaltinis / pastaba": "Yahoo Finance - forwardPE",
          },
          {
              "Rodiklis": "EPS rating",
              "Reikšmė": (
                  str(eps_rating_input) if eps_rating_input > 0 else "N/A"
              ),
              "Šaltinis / pastaba": "IBD (Investors.com) - rankinis įvedimas",
          },
          {
              "Rodiklis": "3 Year EPS Growth Rate",
              "Reikšmė": (
                  f"{eps_growth_3y:.2f}%"
                  if eps_growth_3y is not None
                  else "N/A"
              ),
              "Šaltinis / pastaba": (
                  "Apskaičiuota (CAGR) iš Yahoo Finance ataskaitų"
              ),
          },
          {
              "Rodiklis": "SMR Rating",
              "Reikšmė": smr_rating_input if smr_rating_input != "-" else "N/A",
              "Šaltinis / pastaba": "IBD (Investors.com) - rankinis įvedimas",
          },
          {
              "Rodiklis": "3-Year Sales Growth Rate",
              "Reikšmė": (
                  f"{sales_growth_3y:.2f}%"
                  if sales_growth_3y is not None
                  else "N/A"
              ),
              "Šaltinis / pastaba": (
                  "Apskaičiuota (CAGR) iš Yahoo Finance ataskaitų"
              ),
          },
          {
              "Rodiklis": "Cash",
              "Reikšmė": format_number(info.get("totalCash"), currency=cur),
              "Šaltinis / pastaba": "Yahoo Finance - totalCash",
          },
          {
              "Rodiklis": "Debt",
              "Reikšmė": format_number(info.get("totalDebt"), currency=cur),
              "Šaltinis / pastaba": "Yahoo Finance - totalDebt",
          },
          {
              "Rodiklis": "Total Debt/Equity",
              "Reikšmė": d_e_str,
              "Šaltinis / pastaba": "Yahoo Finance - debtToEquity (mrq)",
          },
          {
              "Rodiklis": "Book Value Per Share",
              "Reikšmė": (
                  f"{info.get('bookValue'):.2f} {cur}"
                  if info.get("bookValue")
                  else "N/A"
              ),
              "Šaltinis / pastaba": "Yahoo Finance - bookValue (mrq)",
          },
          {
              "Rodiklis": "ROE",
              "Reikšmė": (
                  f"{info.get('returnOnEquity') * 100:.2f}%"
                  if info.get("returnOnEquity")
                  else "N/A"
              ),
              "Šaltinis / pastaba": "Yahoo Finance - returnOnEquity",
          },
          {
              "Rodiklis": "MGMT owns",
              "Reikšmė": mgmt_own_str,
              "Šaltinis / pastaba": "Yahoo Finance - heldPercentInsiders",
          },
          {
              "Rodiklis": "Dividend Yield",
              "Reikšmė": div_yield_str,
              "Šaltinis / pastaba": "Yahoo Finance - dividendYield",
          },
          {
              "Rodiklis": "1y Target Est",
              "Reikšmė": (
                  f"{info.get('targetMeanPrice'):.2f} {cur}"
                  if info.get("targetMeanPrice")
                  else "N/A"
              ),
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
      try:
        inc = t.income_stmt
        if inc is not None and not inc.empty:
          rev_s = None
          for cand in ["Total Revenue", "Operating Revenue", "Revenue"]:
            if cand in inc.index:
              rev_s = inc.loc[cand]
              break
          ni_s = None
          for cand in [
              "Net Income",
              "Net Income Common Stockholders",
              "Net Income Including Noncontrolling Interests",
          ]:
            if cand in inc.index:
              ni_s = inc.loc[cand]
              break

          if rev_s is not None and ni_s is not None:
            cols_sorted = sorted([c for c in inc.columns if pd.notna(c)])[-5:]
            years = [pd.to_datetime(c).strftime("%Y") for c in cols_sorted]
            rev_vals = [rev_s[c] / 1e9 for c in cols_sorted]
            ni_vals = [ni_s[c] / 1e9 for c in cols_sorted]

            fig_fin = go.Figure()
            fig_fin.add_trace(
                go.Bar(
                    x=years,
                    y=rev_vals,
                    name=f"Revenues ({cur} mlrd.)",
                    marker_color="#38bdf8",
                    text=[f"{v:.1f}B" for v in rev_vals],
                    textposition="auto",
                )
            )
            fig_fin.add_trace(
                go.Bar(
                    x=years,
                    y=ni_vals,
                    name=f"Net Income ({cur} mlrd.)",
                    marker_color="#10b981",
                    text=[f"{v:.1f}B" for v in ni_vals],
                    textposition="auto",
                )
            )
            fig_fin.update_layout(
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
            fig_fin.update_xaxes(
                showgrid=True, gridcolor="#334155", color="#FFFFFF", title="Metai"
            )
            fig_fin.update_yaxes(
                showgrid=True,
                gridcolor="#334155",
                color="#FFFFFF",
                title=f"Suma ({cur} mlrd.)",
            )
            st.plotly_chart(fig_fin, use_container_width=True)
          else:
            st.info(
                "Nepavyko aptikti pajamų arba grynojo pelno eilutes ataskaitose."
            )
        else:
          st.info("Finansinių ataskaitų duomenų nerasta.")
      except Exception as e:
        st.info(f"Nepavyko sugeneruoti Revenues ir Net Income grafiko: {e}")

      st.divider()

      # 5 SEKCIJA: INSIDER PREKYBA
      st.subheader("👥 INSIDER PREKYBA (paskutiniai 12 mėn.)")

      insider_df = None
      total_buy, total_sell, net_flow = 0.0, 0.0, 0.0
      monthly_insider_vol = {}

      try:
        ins = t.insider_transactions
        if ins is not None and not ins.empty:
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
                monthly_insider_vol[m_key] = (
                    monthly_insider_vol.get(m_key, 0.0) + val
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
                        "Suma": format_number(val, currency=cur),
                    }
                )

              insider_df = pd.DataFrame(rows)
              net_flow = total_buy - total_sell
      except Exception:
        pass

      col_i1, col_i2, col_i3 = st.columns(3)
      col_i1.metric("Iš viso pirkta", format_number(total_buy, currency=cur))
      col_i2.metric("Iš viso parduota", format_number(total_sell, currency=cur))
      col_i3.metric(
          "Grynasis srautas (pirkta-parduota)",
          format_number(net_flow, currency=cur),
      )

      if insider_df is not None and not insider_df.empty:
        st.dataframe(
            insider_df.style.apply(style_insider_df, axis=None),
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.info("Insider prekybos duomenų už paskutinius 12 mėn. nerasta.")

      st.divider()

      # 6 SEKCIJA: GRAFIKAS (MĖNESINĖ INSIDER APYVARTA VS. KAINA)
      st.subheader("📈 MĖNESINĖ INSIDER APYVARTA VS. KAINA")

      if not hist_5y.empty:
        end_dt = hist_5y.index[-1]

        try:
          m_ends = pd.date_range(end=end_dt, periods=13, freq="ME")
        except ValueError:
          m_ends = pd.date_range(end=end_dt, periods=13, freq="M")

        chart_data = []
        for me in m_ends:
          sub_p = hist_5y[hist_5y.index <= me]
          p_at = float(sub_p["Close"].iloc[-1]) if not sub_p.empty else 0.0
          m_str = me.strftime("%Y-%m")
          v_at = monthly_insider_vol.get(m_str, 0.0)
          chart_data.append(
              {"Mėnuo": m_str, "Kaina": p_at, "Insider Apyvarta": v_at}
          )

        df_chart = pd.DataFrame(chart_data)

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        fig.add_trace(
            go.Bar(
                x=df_chart["Mėnuo"],
                y=df_chart["Insider Apyvarta"],
                name=f"Insider Apyvarta ({cur})",
                marker_color="#38bdf8",
                opacity=0.85,
            ),
            secondary_y=False,
        )

        fig.add_trace(
            go.Scatter(
                x=df_chart["Mėnuo"],
                y=df_chart["Kaina"],
                name=f"Akcijos Kaina ({cur})",
                mode="lines+markers",
                line=dict(color="#10b981", width=3),
                marker=dict(size=6),
            ),
            secondary_y=True,
        )

        fig.update_layout(
            paper_bgcolor="#0F172A",
            plot_bgcolor="#0F172A",
            font=dict(color="#FFFFFF"),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(color="#FFFFFF"),
            ),
            margin=dict(l=20, r=20, t=30, b=20),
        )

        fig.update_xaxes(
            title_text="Mėnuo", showgrid=True, gridcolor="#334155", color="#FFFFFF"
        )
        fig.update_yaxes(
            title_text=f"Insider Apyvarta ({cur})",
            secondary_y=False,
            showgrid=True,
            gridcolor="#334155",
            color="#FFFFFF",
        )
        fig.update_yaxes(
            title_text=f"Akcijos Kaina ({cur})",
            secondary_y=True,
            showgrid=False,
            color="#FFFFFF",
        )

        st.plotly_chart(fig, use_container_width=True)
      else:
        st.info("Nepavyko sugeneruoti grafiko, nes nėra istorinių duomenų.")

      st.divider()

      # 7 SEKCIJA: ANALITIKŲ REKOMENDACIJA (PERKELTA Į PATĮ GALĄ)
      st.subheader("🎯 ANALITIKŲ REKOMENDACIJA")

      rec_key = info.get("recommendationKey", "").lower()
      rec_map = {
          "strong_buy": "STIPRIAI PIRKTI",
          "buy": "PIRKTI",
          "hold": "LAIKYTI",
          "underperform": "MAŽIAU PIRKTI",
          "sell": "PARDUOTI",
          "strong_sell": "STIPRIAI PARDUOTI",
      }
      rec_lt = rec_map.get(rec_key, rec_key.upper() if rec_key else "N/A")

      col_rec1, col_rec2, col_rec3 = st.columns(3)
      col_rec1.metric("Rekomendacija", rec_lt)
      col_rec2.metric(
          "Vidutinis balas (1=Strong Buy, 5=Sell)",
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

    except Exception as e:
      st.error(f"Klaida renkant duomenis: {e}")