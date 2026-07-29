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
    div[data-testid="stSelectbox"] > div > div {
        background-color: #FFFFFF !important;
        border: 1.5px solid #0F172A !important;
        border-radius: 10px !important;
        overflow: hidden !important;
    }
    /* Stilius Excel parsisiuntimo mygtukui pagal nuotrauką */
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

# Vietos žymeklis mygtukui šoninėje juostoje (bus sugeneruotas vėliau, kai paruošim duomenis)
sidebar_download_placeholder = st.sidebar.empty()


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

      eps_growth_3y = None
      sales_growth_3y = None
      eps_base_val, eps_latest_val = None, None
      rev_base_val, rev_latest_val = None, None
      eps_base_note, eps_latest_note = "", ""
      rev_base_note, rev_latest_note = "", ""

      financials_years = []
      revenues_vals = []
      net_income_vals = []

      try:
        inc = t.income_stmt
        if inc is not None and not inc.empty:
          valid_cols = sorted([c for c in inc.columns if pd.notna(c)])

          rev_row_candidates = ["Total Revenue", "Operating Revenue"]
          ni_row_candidates = ["Net Income", "Net Income Common Stockholders"]

          r_series, ni_series = None, None
          for rc in rev_row_candidates:
            if rc in inc.index:
              r_series = inc.loc[rc]
              break
          for nic in ni_row_candidates:
            if nic in inc.index:
              ni_series = inc.loc[nic]
              break

          recent_cols = valid_cols[-4:] if len(valid_cols) >= 4 else valid_cols
          for col in recent_cols:
            yr_str = pd.to_datetime(col).strftime("%Y")
            financials_years.append(yr_str)
            r_val = r_series.loc[col] if r_series is not None and col in r_series.index else 0
            ni_val = ni_series.loc[col] if ni_series is not None and col in ni_series.index else 0
            revenues_vals.append(float(r_val) if pd.notna(r_val) else 0.0)
            net_income_vals.append(float(ni_val) if pd.notna(ni_val) else 0.0)

          if len(valid_cols) >= 4:
            col_latest = valid_cols[-1]
            col_base = valid_cols[-4]
            dt_latest = pd.to_datetime(col_latest).strftime("%Y-%m-%d")
            dt_base = pd.to_datetime(col_base).strftime("%Y-%m-%d")

            for cand in ["Diluted EPS", "Basic EPS"]:
              if cand in inc.index:
                eps_base_val = inc.loc[cand, col_base]
                eps_latest_val = inc.loc[cand, col_latest]
                eps_growth_3y = calc_cagr(eps_base_val, eps_latest_val, 3)
                eps_base_note = f"Yahoo Finance - Diluted/Basic EPS, FY {dt_base}"
                eps_latest_note = (
                    f"Yahoo Finance - Diluted/Basic EPS, FY {dt_latest}"
                )
                break

            for cand in ["Total Revenue", "Operating Revenue"]:
              if cand in inc.index:
                rev_base_val = inc.loc[cand, col_base]
                rev_latest_val = inc.loc[cand, col_latest]
                sales_growth_3y = calc_cagr(rev_base_val, rev_latest_val, 3)
                rev_base_note = f"Yahoo Finance - Total Revenue, FY {dt_base}"
                rev_latest_note = f"Yahoo Finance - Total Revenue, FY {dt_latest}"
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

      # 2 SEKCIJA: KAINOS POKYTIS
      st.subheader("KAINOS POKYTIS")
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
                      f"{past_p:.2f}" if isinstance(past_p, (int, float)) else "N/A"
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

      div_yield = info.get("dividendYield")
      if div_yield is not None:
        div_yield_val = div_yield * 100 if div_yield < 1 else div_yield
      else:
        div_yield_val = "N/A"

      d_e = info.get("debtToEquity")

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
              "Rodiklis": "EPS rating",
              "Reikšmė": eps_rating_input if eps_rating_input > 0 else "N/A",
              "Šaltinis / pastaba": (
                  "IBD (Investors.com) - reikia rankinio įvedimo, 1-99 balas"
              ),
          },
          {
              "Rodiklis": "EPS (prieš 3 FY)",
              "Reikšmė": fmt_val(eps_base_val),
              "Šaltinis / pastaba": eps_base_note,
          },
          {
              "Rodiklis": "EPS (paskutiniai FY)",
              "Reikšmė": fmt_val(eps_latest_val),
              "Šaltinis / pastaba": eps_latest_note,
          },
          {
              "Rodiklis": "3 Year EPS Growth Rate",
              "Reikšmė": (
                  f"{eps_growth_3y:.2f}"
                  if eps_growth_3y is not None
                  else "N/A"
              ),
              "Šaltinis / pastaba": (
                  "Apskaičiuota (CAGR) iš Yahoo Finance finansinių ataskaitų"
              ),
          },
          {
              "Rodiklis": "SMR Rating",
              "Reikšmė": smr_rating_input if smr_rating_input != "-" else "N/A",
              "Šaltinis / pastaba": (
                  "IBD (Investors.com) - reikia rankinio įvedimo, A-E balas"
              ),
          },
          {
              "Rodiklis": "Pajamos (prieš 3 FY)",
              "Reikšmė": fmt_large(rev_base_val),
              "Šaltinis / pastaba": rev_base_note,
          },
          {
              "Rodiklis": "Pajamos (paskutiniai FY)",
              "Reikšmė": fmt_large(rev_latest_val),
              "Šaltinis / pastaba": rev_latest_note,
          },
          {
              "Rodiklis": "3-Year Sales Growth Rate",
              "Reikšmė": (
                  f"{sales_growth_3y:.2f}"
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
              "Reikšmė": fmt_val(d_e),
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
                  f"{info.get('returnOnEquity') * 100:.2f}"
                  if info.get("returnOnEquity")
                  else "N/A"
              ),
              "Šaltinis / pastaba": "Yahoo Finance - returnOnEquity (%)",
          },
          {
              "Rodiklis": "MGMT owns",
              "Reikšmė": (
                  f"{mgmt_own_val * 100:.2f}"
                  if isinstance(mgmt_own_val, (int, float))
                  else str(mgmt_own_val)
              ),
              "Šaltinis / pastaba": "Yahoo Finance - heldPercentInsiders (%)",
          },
          {
              "Rodiklis": "Dividend Yield",
              "Reikšmė": (
                  f"{div_yield_val:.2f}"
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

      # 4 SEKCIJA: REVENUES, NET INCOME (GRAFIKAS SU DATA LABELS) - PERKELTAS ČIA
      st.subheader("📊 REVENUES, NET INCOME")
      if financials_years and revenues_vals and net_income_vals:
        rev_b_formatted = [f"{v / 1e9:.1f}B" for v in revenues_vals]
        ni_b_formatted = [f"{v / 1e9:.1f}B" for v in net_income_vals]

        fig_rev_ni = go.Figure()
        fig_rev_ni.add_trace(
            go.Bar(
                x=financials_years,
                y=[v / 1e9 for v in revenues_vals],
                name="Revenues (USD mlrd.)",
                marker_color="#38bdf8",
                text=rev_b_formatted,
                textposition="auto",
            )
        )
        fig_rev_ni.add_trace(
            go.Bar(
                x=financials_years,
                y=[v / 1e9 for v in net_income_vals],
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
            xaxis=dict(title="Metai", showgrid=True, gridcolor="#334155", color="#FFFFFF"),
            yaxis=dict(title="Suma (USD mlrd.)", showgrid=True, gridcolor="#334155", color="#FFFFFF"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_rev_ni, use_container_width=True)
      else:
        st.info("Finansinių ataskaitų duomenų pajamų ir pelno grafikui nerasta.")

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
                        "Suma": f"{val:,.2f}" if pd.notna(val) else "0.00",
                    }
                )

              insider_df = pd.DataFrame(rows)
              net_flow = total_buy - total_sell
      except Exception:
        pass

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
              {
                  "Mėnuo": m_str,
                  "Kaina (USD)": round(p_at, 2),
                  "Insider apyvarta (USD)": round(v_at, 2),
              }
          )

        df_chart = pd.DataFrame(chart_data)

        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(
            go.Bar(
                x=df_chart["Mėnuo"],
                y=df_chart["Insider apyvarta (USD)"],
                name=f"Insider Apyvarta ({cur})",
                marker_color="#38bdf8",
                opacity=0.85,
            ),
            secondary_y=False,
        )
        fig.add_trace(
            go.Scatter(
                x=df_chart["Mėnuo"],
                y=df_chart["Kaina (USD)"],
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
            margin=dict(l=20, r=20, t=30, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)

      st.divider()

      # 7 SEKCIJA: ANALITIKŲ REKOMENDACIJA (PAČIOJE APAČIOJE)
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
          "Vidutinis balas",
          f"{info.get('recommendationMean'):.2f}"
          if info.get("recommendationMean")
          else "N/A",
      )
      col_rec3.metric(
          "Analitikų skaičius",
          str(info.get("numberOfAnalystOpinions"))
          if info.get("numberOfAnalystOpinions")
          else "N/A",
      )

      # ------------------------------------------------------------------------------
      # 8. EXCEL EKSPORTAS NAUDOJANT OPENPYXL
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

      def write_section_header(title_text):
        global current_row
        ws.merge_cells(
            start_row=current_row,
            start_column=1,
            end_row=current_row,
            end_column=3,
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

      if "df_price" in locals() and not df_price.empty:
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

      # 3. ANALITIKŲ REKOMENDACIJA
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
              (
                  f"{info.get('recommendationMean'):.2f}"
                  if info.get("recommendationMean")
                  else "N/A"
              ),
              "Yahoo Finance - recommendationMean",
          ],
          [
              "Analitikų skaičius",
              info.get("numberOfAnalystOpinions")
              if info.get("numberOfAnalystOpinions")
              else "N/A",
              "Yahoo Finance - numberOfAnalystOpinions",
          ],
      ]
      for idx, row_data in enumerate(rec_rows):
        for col_idx, val in enumerate(row_data, 1):
          cell = ws.cell(row=current_row, column=col_idx, value=val)
          cell.font = regular_font
          cell.border = border_all
          if idx % 2 == 1:
            cell.fill = zebra_fill
          if col_idx == 2:
            cell.alignment = Alignment(horizontal="right")
        current_row += 1
      current_row += 1

      # 4. INSIDER PREKYBA
      write_section_header("INSIDER PREKYBA (paskutiniai 12 mėn.)")
      insider_headers = ["Data", "Asmuo", "Pareigos", "Sandoris", "Akcijos", "Suma"]
      for col_idx, h in enumerate(insider_headers, 1):
        cell = ws.cell(row=current_row, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border_all
      current_row += 1

      if insider_df is not None and not insider_df.empty:
        for idx, r in insider_df.iterrows():
          c1 = ws.cell(row=current_row, column=1, value=r["Data"])
          c2 = ws.cell(row=current_row, column=2, value=r["Asmuo"])
          c3 = ws.cell(row=current_row, column=3, value=r["Pareigos"])
          c4 = ws.cell(row=current_row, column=4, value=r["Sandoris"])
          c5 = ws.cell(row=current_row, column=5, value=r["Akcijos"])
          c6 = ws.cell(row=current_row, column=6, value=r["Suma"])

          for c in [c1, c2, c3, c4, c5, c6]:
            c.font = regular_font
            c.border = border_all
            if idx % 2 == 1:
              c.fill = zebra_fill
          c5.alignment = Alignment(horizontal="right")
          c6.alignment = Alignment(horizontal="right")
          current_row += 1

      summary_items = [
          ("Iš viso pirkta", f"{total_buy:,.2f}"),
          ("Iš viso parduota", f"{total_sell:,.2f}"),
          ("Grynasis srautas (pirkta - parduota)", f"{net_flow:,.2f}"),
      ]
      for label, val in summary_items:
        ws.cell(row=current_row, column=1, value=label).font = bold_font
        ws.cell(row=current_row, column=1).border = border_all
        c_val = ws.cell(row=current_row, column=2, value=val)
        c_val.font = bold_font
        c_val.border = border_all
        c_val.alignment = Alignment(horizontal="right")
        current_row += 1
      current_row += 1

      # 5. GRAFIKO DUOMENYS
      write_section_header("MĖNESINĖ INSIDER APYVARTA VS. KAINA")
      write_table_headers(["Mėnuo", "Kaina (USD)", "Insider apyvarta (USD)"])

      if "df_chart" in locals() and not df_chart.empty:
        for idx, r in df_chart.iterrows():
          c1 = ws.cell(row=current_row, column=1, value=r["Mėnuo"])
          c2 = ws.cell(
              row=current_row, column=2, value=f"{r['Kaina (USD)']:,.2f}"
          )
          c3 = ws.cell(
              row=current_row,
              column=3,
              value=f"{r['Insider apyvarta (USD)']:,.2f}",
          )

          for c in [c1, c2, c3]:
            c.font = regular_font
            c.border = border_all
            if idx % 2 == 1:
              c.fill = zebra_fill
          c2.alignment = Alignment(horizontal="right")
          c3.alignment = Alignment(horizontal="right")
          current_row += 1

      for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
          try:
            if cell.value:
              val_str = str(cell.value)
              if len(val_str) > max_length:
                max_length = len(val_str)
          except:
            pass
        adjusted_width = max(max_length + 4, 12)
        ws.column_dimensions[col_letter].width = adjusted_width

      wb.save(output)
      processed_data = output.getvalue()

      # Mygtukas įdėtas šoninėje juostoje po SMR Rating su norimu pavadinimu "Parsisiųsti XLS"
      sidebar_download_placeholder.download_button(
          label="⬇️ Parsisiųsti XLS",
          data=processed_data,
          file_name=f"{ticker_input}_analize.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    except Exception as e:
      st.error(f"Klaida renkant duomenis: {e}")