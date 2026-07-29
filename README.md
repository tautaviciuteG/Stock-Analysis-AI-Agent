# 📊 Stock Analysis Agent

A modern and interactive **Streamlit** application designed for fundamental and technical stock analysis. The tool automatically fetches financial data, calculates key growth metrics, and presents them in a clean, visual interface.

---

## ✨ Key Features

* **📈 Price History:** Interactive 5-year stock price chart displaying the company name and ticker.
* **📌 Key Metrics:** Automatically calculated and displayed Forward P/E, EPS growth (CAGR), SMR/EPS IBD ratings, debt metrics (Total Debt/Equity), ROE, dividend yield, and more (with color-coded evaluations).
* **✍️ Manual IBD Input:** The sidebar allows manual entry for EPS rating (1–99) and SMR Rating (A–E).
* **📉 Price Changes:** Detailed table showing price changes over 5 years, 1 year, YTD, 6 months, and 1 month.
* **🎯 Analyst Recommendations:** Consensus recommendations, average score, and number of analysts.
* **📊 Revenues & Net Income:** Bar charts of financial statements for the last 5 years.
* **👥 Insider Trading:** Summary of management transactions over the last 12 months, net flow calculation, and a detailed table.
* **📈 Monthly Insider Turnover vs. Price:** Dual-axis chart comparing management stock activity with price changes over time.

---

## 🛠️ Technologies Used

* **Python**
* **Streamlit** (For user interface)
* **yfinance** (For fetching financial data from Yahoo Finance)
* **Plotly** (For interactive charts)
* **Pandas** (For data processing and analysis)
