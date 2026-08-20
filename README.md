# 📊 Stock Analysis Agent

An interactive **Streamlit** application for fundamental and technical stock analysis, with a Lithuanian-language UI. The app automatically pulls data from **Yahoo Finance**, calculates key growth metrics, and visualizes them, while an integrated **Mistral AI** agent generates in-depth analysis and answers questions in a chat interface.

---

## ✨ Key Features

### 📈 Price Chart (Yahoo Finance style)
* Interactive **Plotly** price chart with period buttons: `1D`, `5D`, `1M`, `6M`, `YTD`, `1Y`, `5Y`, `All`.
* **1D/5D** views use real intraday (minute-level) data, including pre-market and post-market sessions.
* Volume bar subplot below the price line.
* Dynamic **percentage change badge** that updates based on the selected period.
* Optional overlay indicators: **50-day moving average**, **200-day moving average**, **Relative Strength vs. S&P 500**, **MACD (12, 26, 9)**.
* Header metrics: current price with change, volume vs. average volume, and — when trading outside regular hours — pre-/post-market price.

### 📌 Key Metrics
* Automatically calculated and color-coded (green = favorable) metrics: Forward P/E, 3-year EPS growth, 3-year sales growth (CAGR or linear regression trend), cash, debt, Total Debt/Equity, book value per share, ROE, percent held by insiders, dividend yield, 1-year target price.

### ⚖️ Multi-Ticker Comparison
* Entering multiple tickers (comma-separated) gives the first one a full analysis, while all others are shown in:
  * a comparison table (price, P/E, ROE, dividend yield, recommendation);
  * a normalized 1-year price change (%) chart.

### 📊 Revenue & Net Income
* Bar chart with a toggle between **annual** and **quarterly** financials (up to the last 4 years / 8 quarters).

### 👥 Insider Transactions
* Summary of total bought, sold, and net flow over the last 12 months.
* Dual-axis chart: monthly transaction amounts (bars) and stock price (line), with buy/sell markers.
* Detailed transaction table.

### 🎯 Analyst Recommendations
* Consensus recommendation, average score, and number of analysts.


### 🤖 AI Agent (Mistral AI)
* The user enters their own **Mistral API key** in the sidebar (the key and requests are never stored by the app).
* A **"Generate AI Analysis"** button produces a structured 8-section report (metrics, financial health, dividends, insider activity, strengths, risks, analyst recommendations, summary).
* Built-in **chat window** for free-form questions about the selected stock; chat history is kept in session state separately for each ticker.
* Pulls the **full, real historical dividend payment record** from Yahoo Finance (date + amount per share) — used as a reliable source for the AI agent's answers.

---

## 🛠️ Tech Stack
* **Python**
* **Streamlit** — user interface
* **yfinance** — financial data from Yahoo Finance
* **Plotly** — interactive charts
* **Pandas / NumPy** — data processing and calculations
* **Mistral AI API** — AI analysis and chat agent
* Code generated with AI assistance.

---

To use the AI analysis and chat features, you'll need a free Mistral API key from [console.mistral.ai](https://console.mistral.ai/).
