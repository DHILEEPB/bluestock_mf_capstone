# Mutual Fund Analytics & Quantitative Risk Platform
## Capstone Project (Days 1–7)

An enterprise-grade data engineering, quantitative analytics, and business intelligence platform built to ingest mutual fund data, structure an SQLite Star Schema, calculate risk-adjusted returns, evaluate tail risk, and output structured dashboards.

---

## 🏗️ Project Architecture & Folder Structure

```
bluestock_mf_capstone/
├── data/
│   ├── raw/                   ← Drop raw AMFI CSV data files here
│   ├── processed/             ← Cleaned metrics CSVs and Parquet files
│   └── db/                    ← SQLite database directory (mutual_funds.db)
│
├── notebooks/                 ← Jupyter notebooks containing executed results
│   ├── 01_data_ingestion.ipynb ← Stage 1: Data ingestion & schema checks
│   ├── 02_data_cleaning.ipynb  ← Stage 2: Data cleaning & standardization
│   ├── 03_eda_analysis.ipynb   ← Stage 3: Exploratory charts & visualizations
│   ├── 04_performance_analytics.ipynb ← Stage 4: CAGR, Sharpe, Alpha, Beta ratios
│   └── 05_advanced_analytics.ipynb   ← Stage 5: VaR, CVaR, HHI, cohort analyses
│
├── scripts/                   ← Production python scripts
│   ├── etl_pipeline.py        ← End-to-end data ingestion, cleaning & load
│   ├── live_nav_fetch.py      ← Fetches latest NAV prices via MFAPI
│   ├── compute_metrics.py     ← Computes performance scorecard & risk ratios
│   └── recommender.py         ← Quant-based portfolio recommendation engine
│
├── sql/                       ← Database schema & queries
│   ├── schema.sql             ← SQLite Star Schema DDL statements
│   └── queries.sql            ← Common analytical query inventory
│
├── dashboard/                 ← Power BI Dashboard files
│   └── bluestock_mf.pbix      ← Interactive Power BI dashboard deck
│
├── reports/                   ← Final PDF and PPTX deliverables
│   ├── Final_Report.pdf       ← Comprehensive 15-20 page project report
│   └── Presentation.pptx      ← 12-Slide executive widescreen presentation
│
├── README.md                  ← This developer guide
├── requirements.txt           ← External library dependencies
└── .gitignore                 ← Git exclusion mapping
```

---

## ⚡ Quick Start & Setup Instructions

### 1. Prerequisite Libraries
Install all required libraries via `pip` (standard python 3.11 environment):
```bash
pip install -r requirements.txt
```

### 2. Execute ETL Pipeline
To ingest the raw data, perform data cleaning (handling duplicates, outliers, and parsing dates), apply `sql/schema.sql`, and load the database:
```bash
python scripts/etl_pipeline.py
```
This will generate the SQLite database at `data/db/mutual_funds.db` and save clean Parquet/CSV files in `data/processed/`.

### 3. Compute Metrics and Ratios
To run the quantitative calculations (Sharpe, Sortino, Alpha, Beta, Tracking Error, HHI, VaR, CVaR, SIP continuity, and cohorts) and generate plots:
```bash
python scripts/compute_metrics.py
```
This exports report CSVs to `data/processed/` and renders benchmark comparison plots to `reports/`.

### 4. Fetch Live NAV
To fetch the latest NAV snapshot directly from the MFAPI:
```bash
python scripts/live_nav_fetch.py
```

### 5. Run Recommendations
To run the quant-based fund recommendation engine:
```bash
python scripts/recommender.py
```

---

## 📊 Analytics Modules Summary

### A. SQL Star Schema (`data/db/mutual_funds.db`)
Our database represents a dimensional model optimized for query performance:
*   **Dimensions**: `dim_amc`, `dim_scheme`, `dim_date`, `dim_investor`.
*   **Facts**: 
    *   `fact_nav_history` (daily Net Asset Values grain).
    *   `fact_investor_transactions` (individual buy/sell trade ledger).
    *   `fact_scheme_performance` (quarterly return snapshots).
*   **Analytical Views**:
    *   `v_latest_nav`: Fetches the most recent Net Asset Value and absolute change for each scheme.
    *   `v_monthly_transactions`: Summarizes transaction volume, transaction counts, and average investments.
    *   `v_scheme_performance_latest`: Compiles the most recent quarterly snapshot of Returns, Sharpe ratios, and AUM values.

### B. Performance Scorecard (Day 4)
*   **KPIs calculated**: 3Y CAGR, Sharpe Ratio, Sortino Ratio, CAPM Alpha & Beta (regression vs Nifty 100 benchmark), Maximum Drawdown, and Tracking Error.
*   **Fund Scorecard Weights**: 30% 3Y Return Rank, 25% Sharpe Rank, 20% Alpha Rank, 15% Expense Rank (Inverse), 10% Drawdown Rank (Inverse).
*   **Scorecard Winner**: **DSP Top 100 Equity Fund - Direct Growth** (Score: 1.90, Rank 1).

### C. Advanced Risk Analytics (Day 6)
*   **Value at Risk (95% VaR)** & **Conditional VaR (95% CVaR)**: Identifies tail risk losses on daily returns (e.g. ICICI Tech Fund VaR = 2.33% daily vs. Mirae Large Cap = 2.08%).
*   **Portfolio Sector HHI Concentration**: Evaluates sector diversification (HHI < 1500 = Low, >= 2500 = High Concentration). The Tech sectoral fund shows high concentration (HHI = 7,338.00).
*   **SIP Continuity Analysis**: Tracks monthly systematic payment streaks and churn rates.
*   **Recommender API**: Production module `recommender.py` accepts investor risk profiles and queries optimal funds.

---

## 🎨 Power BI Visual Configuration
Custom branding rules are defined in `reports/bluestock_theme.json` to apply a dark-mode carbon theme (`#0F111A` background with electric cyan and green accents).

### 4-Page Specifications:
*   **Page 1: Industry Overview**: KPI cards, monthly AUM inflow trends, and AMC market share treemaps.
*   **Page 2: Fund Performance**: Risk-return bubbles (Beta vs CAGR), dynamic growth curves indexed to base ₹100, and scorecard data bars.
*   **Page 3: Investor Analytics**: State-wise heatmaps, transaction splits, age cohort columns.
*   **Page 4: SIP & Market Trends**: SIP volumes vs Nifty index, quarterly matrix heatmaps.
*   *Interactive Features*: Hidden tooltip canvas for risk details and Drill-Through links for transaction sheets.

---

## 🏆 Portfolio Quant Ratios Leaderboard

| AMFI Code | Scheme Name | 3Y CAGR | Sharpe | CAPM Beta | Max Drawdown | Sector HHI | Scorecard Rank |
|---|---|---|---|---|---|---|---|
| **120594** | DSP Top 100 Equity Fund | 39.78% | 1.22 | -0.04 | -24.73% | 1827.16 | **#1** |
| **120503** | Mirae Asset Large Cap Fund | 33.15% | 1.07 | -0.03 | -17.47% | 1827.16 | **#2** |
| **119551** | Axis Bluechip Fund | 29.86% | 0.96 | 0.02 | -23.78% | 1827.16 | **#3** |
| **122639** | SBI Small Cap Fund | 14.66% | 0.41 | 0.05 | -31.97% | 1650.00 | **#5** |
| **118989** | ICICI Pru Technology Fund | 2.13% | -0.09 | 0.04 | -29.88% | 7338.00 | **#6** |
