"""
compute_metrics.py
==================
Mutual Fund Analytics Capstone Project -- Day 4 & 6 Quantitative Metrics
----------------------------------------------------------------------
Purpose : Calculate CAGRs, Sharpe/Sortino ratios, CAPM Alpha/Beta, MDD, Tracking Error,
          Weighted Scorecard, Historical VaR/CVaR, Sector HHI, Cohorts, and SIP streaks.
          Saves CSVs to data/processed/ and renders performance plots.
"""

import os
import sys
import logging
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import scipy.stats as stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine, text

# Setup paths
PROJECT_ROOT = Path(__file__).resolve().parent.parent
PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
REPORTS_DIR = PROJECT_ROOT / "reports"
DB_PATH = PROJECT_ROOT / "data" / "db" / "mutual_funds.db"

for d in (PROCESSED_DIR, REPORTS_DIR):
    d.mkdir(parents=True, exist_ok=True)

# Configure Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("compute_metrics")

# Constants
RISK_FREE_RATE_ANN = 0.065
TRADING_DAYS_YEAR = 252

def load_data_from_db():
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database not found at {DB_PATH}. Run etl_pipeline.py first.")
    
    engine = create_engine(f"sqlite:///{DB_PATH}")
    
    # Load daily NAV and schemes
    with engine.connect() as conn:
        df_nav = pd.read_sql(
            "SELECT n.scheme_code, s.scheme_name, d.full_date, n.nav "
            "FROM fact_nav_history n "
            "JOIN dim_scheme s ON n.scheme_code = s.scheme_code "
            "JOIN dim_date d ON n.date_id = d.date_id", conn
        )
        df_perf_snapshots = pd.read_sql("SELECT * FROM fact_scheme_performance", conn)
        df_txn = pd.read_sql(
            "SELECT t.*, d.full_date as transaction_date "
            "FROM fact_investor_transactions t "
            "JOIN dim_date d ON t.date_id = d.date_id", conn
        )
        df_master = pd.read_sql("SELECT * FROM dim_scheme", conn)
        
    engine.dispose()
    
    # Parse dates
    df_nav["nav_date"] = pd.to_datetime(df_nav["full_date"])
    df_txn["txn_date"] = pd.to_datetime(df_txn["transaction_date"])
    
    return df_nav, df_perf_snapshots, df_txn, df_master

def load_benchmarks():
    # Look for cached benchmark files
    n50_path = PROCESSED_DIR / "nifty50.csv"
    n100_path = PROCESSED_DIR / "nifty100.csv"
    
    if n50_path.exists() and n100_path.exists():
        logger.info("Loading cached benchmark files from processed/...")
        nifty50_df = pd.read_csv(n50_path, parse_dates=["Date"]).set_index("Date")
        nifty100_df = pd.read_csv(n100_path, parse_dates=["Date"]).set_index("Date")
    else:
        logger.info("Downloading benchmarks from yfinance...")
        import yfinance as yf
        nifty50_df = yf.download("^NSEI", start="2022-01-01", end="2024-12-31")
        nifty100_df = yf.download("^CNX100", start="2022-01-01", end="2024-12-31")
        
        # Flatten MultiIndex if present
        if isinstance(nifty50_df.columns, pd.MultiIndex):
            nifty50_df.columns = nifty50_df.columns.get_level_values(0)
        if isinstance(nifty100_df.columns, pd.MultiIndex):
            nifty100_df.columns = nifty100_df.columns.get_level_values(0)
            
        nifty50_df = nifty50_df[["Close"]].reset_index()
        nifty100_df = nifty100_df[["Close"]].reset_index()
        
        nifty50_df.to_csv(n50_path, index=False)
        nifty100_df.to_csv(n100_path, index=False)
        
        nifty50_df.set_index("Date", inplace=True)
        nifty100_df.set_index("Date", inplace=True)
        
    nifty50_df.sort_index(inplace=True)
    nifty100_df.sort_index(inplace=True)
    
    nifty50_df["daily_return"] = nifty50_df["Close"].pct_change()
    nifty100_df["daily_return"] = nifty100_df["Close"].pct_change()
    
    return nifty50_df, nifty100_df

def compute_performance_ratios(df_nav, df_perf_snapshots, nifty100_df):
    logger.info("Computing Performance Metrics...")
    
    # Calculate daily returns
    df_nav.sort_values(["scheme_code", "nav_date"], inplace=True)
    df_nav["daily_return"] = df_nav.groupby("scheme_code")["nav"].pct_change()
    
    daily_rf = RISK_FREE_RATE_ANN / TRADING_DAYS_YEAR
    unique_schemes = df_nav["scheme_code"].unique()
    
    n100_ret = nifty100_df["daily_return"].dropna()
    results = []
    
    for sc in unique_schemes:
        s_nav = df_nav[df_nav["scheme_code"] == sc].copy()
        s_name = s_nav["scheme_name"].iloc[0]
        
        # 1. CAGR 3Y (2022-2024 range)
        s_nav.sort_values("nav_date", inplace=True)
        first_nav = s_nav["nav"].iloc[0]
        last_nav = s_nav["nav"].iloc[-1]
        days = (s_nav["nav_date"].iloc[-1] - s_nav["nav_date"].iloc[0]).days
        cagr_3y = (last_nav / first_nav) ** (365.25 / days) - 1
        
        # 2. CAGR 1Y (2024 range)
        s_nav_2024 = s_nav[s_nav["nav_date"].dt.year == 2024]
        cagr_1y = (s_nav_2024["nav"].iloc[-1] / s_nav_2024["nav"].iloc[0]) ** (365.25 / (s_nav_2024["nav_date"].iloc[-1] - s_nav_2024["nav_date"].iloc[0]).days) - 1 if not s_nav_2024.empty else np.nan
        
        # 3. CAGR 5Y (from database fact table / snapshot)
        p_row = df_perf_snapshots[df_perf_snapshots["scheme_code"] == sc]
        cagr_5y = p_row["returns_5y"].iloc[0] / 100.0 if not p_row.empty else np.nan
        expense_ratio = p_row["expense_ratio"].iloc[0] if not p_row.empty else np.nan
        
        # 4. Sharpe
        s_returns = s_nav["daily_return"].dropna()
        excess = s_returns - daily_rf
        std_dev = s_returns.std()
        sharpe = (excess.mean() / std_dev) * np.sqrt(TRADING_DAYS_YEAR) if std_dev > 0 else np.nan
        
        # 5. Sortino
        downside_diffs = np.minimum(excess, 0.0)
        downside_dev = np.sqrt(np.mean(downside_diffs ** 2))
        sortino = (excess.mean() / downside_dev) * np.sqrt(TRADING_DAYS_YEAR) if downside_dev > 0 else np.nan
        
        # 6. CAPM Alpha & Beta
        merged = pd.merge(s_nav[["nav_date", "daily_return"]].dropna(), n100_ret.to_frame("n100"), left_on="nav_date", right_index=True)
        if len(merged) > 30:
            slope, intercept, r_val, _, _ = stats.linregress(merged["n100"] - daily_rf, merged["daily_return"] - daily_rf)
            beta = slope
            alpha = intercept * TRADING_DAYS_YEAR
        else:
            beta, alpha = np.nan, np.nan
            
        # 7. Max Drawdown
        peaks = s_nav["nav"].cummax()
        drawdowns = (s_nav["nav"] - peaks) / peaks
        max_dd = drawdowns.min()
        
        # 8. Tracking Error
        if len(merged) > 30:
            active_ret = merged["daily_return"] - merged["n100"]
            tracking_error = active_ret.std() * np.sqrt(TRADING_DAYS_YEAR)
        else:
            tracking_error = np.nan
            
        results.append({
            "scheme_code": sc,
            "scheme_name": s_name,
            "cagr_1y": cagr_1y,
            "cagr_3y": cagr_3y,
            "cagr_5y": cagr_5y,
            "sharpe_ratio": sharpe,
            "sortino_ratio": sortino,
            "beta": beta,
            "alpha_annual": alpha,
            "max_drawdown": max_dd,
            "tracking_error": tracking_error,
            "expense_ratio": expense_ratio
        })
        
    df_metrics = pd.DataFrame(results)
    
    # Save CSV reports to data/processed
    df_metrics[["scheme_code", "scheme_name", "sharpe_ratio"]].to_csv(PROCESSED_DIR / "sharpe_ratio.csv", index=False)
    df_metrics[["scheme_code", "scheme_name", "sortino_ratio"]].to_csv(PROCESSED_DIR / "sortino_ratio.csv", index=False)
    df_metrics[["scheme_code", "scheme_name", "beta", "alpha_annual"]].to_csv(PROCESSED_DIR / "alpha_beta.csv", index=False)
    
    return df_metrics

def generate_fund_scorecard(df_metrics):
    logger.info("Generating Weighted Fund Scorecard...")
    
    scorecard = df_metrics.copy()
    
    # Ranks (1 is best, 10 is worst)
    scorecard["rank_returns"] = scorecard["cagr_3y"].rank(ascending=False, method="min")
    scorecard["rank_sharpe"] = scorecard["sharpe_ratio"].rank(ascending=False, method="min")
    scorecard["rank_alpha"] = scorecard["alpha_annual"].rank(ascending=False, method="min")
    scorecard["rank_expense"] = scorecard["expense_ratio"].rank(ascending=True, method="min")
    scorecard["rank_drawdown"] = scorecard["max_drawdown"].rank(ascending=False, method="min")
    
    # Score calculation
    scorecard["scorecard_score"] = (
        0.30 * scorecard["rank_returns"] +
        0.25 * scorecard["rank_sharpe"] +
        0.20 * scorecard["rank_alpha"] +
        0.15 * scorecard["rank_expense"] +
        0.10 * scorecard["rank_drawdown"]
    )
    scorecard["final_scorecard_rank"] = scorecard["scorecard_score"].rank(ascending=True, method="min").astype(int)
    scorecard.sort_values("final_scorecard_rank", inplace=True)
    
    scorecard.to_csv(PROCESSED_DIR / "fund_scorecard.csv", index=False)
    logger.info("Saved fund_scorecard.csv to data/processed/.")
    return scorecard

def compute_advanced_risk_metrics(df_nav, df_txn):
    logger.info("Computing Advanced Risk Metrics (VaR, CVaR, HHI, Continuity)...")
    
    # 1. Historical VaR & CVaR (95%)
    df_nav.sort_values(["scheme_code", "nav_date"], inplace=True)
    df_nav["daily_return"] = df_nav.groupby("scheme_code")["nav"].pct_change()
    
    var_cvar_results = []
    unique_schemes = df_nav["scheme_code"].unique()
    
    for sc in unique_schemes:
        s_ret = df_nav[df_nav["scheme_code"] == sc]["daily_return"].dropna().values
        s_name = df_nav[df_nav["scheme_code"] == sc]["scheme_name"].iloc[0]
        
        if len(s_ret) > 30:
            var_95 = -np.percentile(s_ret, 5)
            cvar_mask = s_ret <= -var_95
            cvar_95 = -s_ret[cvar_mask].mean() if cvar_mask.sum() > 0 else var_95
        else:
            var_95, cvar_95 = np.nan, np.nan
            
        var_cvar_results.append({
            "scheme_code": sc,
            "scheme_name": s_name,
            "historical_var_95": var_95,
            "conditional_var_95": cvar_95
        })
    df_var_cvar = pd.DataFrame(var_cvar_results)
    df_var_cvar.to_csv(PROCESSED_DIR / "var_cvar_report.csv", index=False)
    
    # 2. Sector HHI concentration report
    # Build simulated holdings weights per scheme
    holdings = []
    sectors = ["Financial Services", "Technology", "Healthcare", "Energy", "Consumer Goods"]
    for sc in unique_schemes:
        s_name = df_nav[df_nav["scheme_code"] == sc]["scheme_name"].iloc[0]
        if "Technology" in s_name:
            # high tech concentration
            weights = [0.05, 0.85, 0.03, 0.02, 0.05]
        else:
            weights = [0.35, 0.20, 0.15, 0.15, 0.15]
            
        for i, sector in enumerate(sectors):
            holdings.append({
                "scheme_code": sc,
                "scheme_name": s_name,
                "sector": sector,
                "weight_pct": weights[i]
            })
    df_holdings = pd.DataFrame(holdings)
    df_holdings.to_csv(PROCESSED_DIR / "portfolio_holdings.csv", index=False)
    
    hhi_results = []
    for sc in unique_schemes:
        s_name = df_nav[df_nav["scheme_code"] == sc]["scheme_name"].iloc[0]
        w_pct = df_holdings[df_holdings["scheme_code"] == sc]["weight_pct"].values * 100
        hhi = np.sum(w_pct ** 2)
        level = "High Concentration" if hhi >= 2500 else "Moderate Concentration" if hhi >= 1500 else "Low Concentration"
        hhi_results.append({
            "scheme_code": sc,
            "scheme_name": s_name,
            "sector_hhi": hhi,
            "concentration_level": level
        })
    df_hhi = pd.DataFrame(hhi_results)
    df_hhi.to_csv(PROCESSED_DIR / "sector_hhi_report.csv", index=False)
    
    # 3. Investor Cohort Analysis
    df_txn["txn_date"] = pd.to_datetime(df_txn["txn_date"])
    first_txn = df_txn.groupby("investor_id")["txn_date"].min().reset_index()
    first_txn.rename(columns={"txn_date": "first_txn_date"}, inplace=True)
    first_txn["cohort_quarter"] = first_txn["first_txn_date"].dt.to_period("Q").astype(str)
    
    df_cohort = df_txn.merge(first_txn[["investor_id", "cohort_quarter"]], on="investor_id")
    cohort_summary = []
    for quarter, group in df_cohort.groupby("cohort_quarter"):
        investor_count = group["investor_id"].nunique()
        total_inflow = group[group["transaction_type"].isin(["SIP", "Lumpsum", "Switch-In"])]["amount_inr"].sum()
        total_outflow = group[group["transaction_type"] == "Redemption"]["amount_inr"].sum()
        cohort_summary.append({
            "cohort_quarter": quarter,
            "unique_investors": investor_count,
            "total_inflow_inr": total_inflow,
            "total_outflow_inr": total_outflow,
            "net_inflow_inr": total_inflow - total_outflow
        })
    df_cohort_report = pd.DataFrame(cohort_summary)
    df_cohort_report.to_csv(PROCESSED_DIR / "cohort_analysis.csv", index=False)

    # 4. SIP Continuity Analysis
    df_sip = df_txn[df_txn["transaction_type"] == "SIP"].copy()
    df_sip["year_month"] = df_sip["txn_date"].dt.to_period("M")
    sip_accounts = []
    for (inv_id, sc), group in df_sip.groupby(["investor_id", "scheme_code"]):
        months = sorted(group["year_month"].unique())
        first_month = months[0]
        last_month = months[-1]
        expected_months = (last_month - first_month).n + 1
        actual_months = len(months)
        continuity_rate = actual_months / expected_months if expected_months > 0 else 1.0
        
        streak = 1
        max_streak = 1
        for i in range(1, len(months)):
            if (months[i] - months[i-1]).n == 1:
                streak += 1
            else:
                streak = 1
            max_streak = max(max_streak, streak)
            
        status = "Active" if last_month >= pd.Period("2024-10", "M") else "Inactive"
        sip_accounts.append({
            "investor_id": inv_id,
            "scheme_code": sc,
            "first_sip_month": str(first_month),
            "last_sip_month": str(last_month),
            "expected_months": expected_months,
            "actual_months": actual_months,
            "continuity_rate": continuity_rate,
            "max_consecutive_streak": max_streak,
            "status": status
        })
    df_sip_report = pd.DataFrame(sip_accounts)
    df_sip_report.to_csv(PROCESSED_DIR / "sip_continuity_report.csv", index=False)
    logger.info("Advanced risk analytics completed and reports exported.")

def generate_visuals(df_nav, nifty100_df):
    logger.info("Generating visual charts...")
    
    # 1. Cumulative Growth Chart
    cum_df = pd.pivot_table(df_nav, index="nav_date", columns="scheme_name", values="nav")
    cum_df = cum_df.ffill().bfill()
    cum_growth = (cum_df / cum_df.iloc[0]) * 100.0
    
    n100_close = nifty100_df["Close"].ffill().bfill()
    n100_normalized = (n100_close / n100_close.iloc[0]) * 100.0
    cum_growth["NIFTY 100 Benchmark"] = n100_normalized
    
    plt.figure(figsize=(12, 6), facecolor="#10121C")
    ax = plt.gca()
    ax.set_facecolor("#1A1D2B")
    ax.grid(True, color="#2E3440", linestyle=":", alpha=0.5)
    
    for col in cum_growth.columns:
        if col == "NIFTY 100 Benchmark":
            ax.plot(cum_growth.index, cum_growth[col], label="Nifty 100 Benchmark", color="red", linestyle="--", linewidth=2.0)
        else:
            short_name = col.split("-")[0].strip()[:20]
            ax.plot(cum_growth.index, cum_growth[col], label=short_name, alpha=0.7, linewidth=1.0)
            
    ax.set_title("Cumulative Growth of ₹100 Investment (2022–2024)", color="white", fontsize=12, fontweight="bold", pad=12)
    ax.set_xlabel("Date", color="lightgray")
    ax.set_ylabel("Value (₹)", color="lightgray")
    ax.tick_params(colors="lightgray", labelsize=8)
    for s in ax.spines.values():
        s.set_edgecolor("#2E3440")
    ax.legend(bbox_to_anchor=(1.01, 1), loc='upper left', facecolor="#1A1D2B", edgecolor="#2E3440", labelcolor="white", fontsize=8)
    
    plt.savefig(str(REPORTS_DIR / "benchmark_comparison.png"), dpi=150, bbox_inches="tight", facecolor="#10121C")
    plt.close()
    
    # 2. Rolling Sharpe Ratio Chart
    df_returns = df_nav.copy()
    df_pivot = df_returns.pivot(index="nav_date", columns="scheme_name", values="daily_return")
    df_pivot = df_pivot.ffill().bfill()
    
    daily_rf = RISK_FREE_RATE_ANN / TRADING_DAYS_YEAR
    rolling_mean = df_pivot.rolling(window=90).mean()
    rolling_std = df_pivot.rolling(window=90).std()
    rolling_sharpe = ((rolling_mean - daily_rf) / rolling_std) * np.sqrt(TRADING_DAYS_YEAR)
    
    plt.figure(figsize=(12, 6), facecolor="#10121C")
    ax = plt.gca()
    ax.set_facecolor("#1A1D2B")
    ax.grid(True, color="#2E3440", linestyle=":", alpha=0.5)
    
    for col in rolling_sharpe.columns:
        short_name = col.split("-")[0].strip()[:20]
        ax.plot(rolling_sharpe.index, rolling_sharpe[col], label=short_name, alpha=0.7, linewidth=1.0)
        
    ax.set_title("Rolling 90-Day Annualized Sharpe Ratio", color="white", fontsize=12, fontweight="bold", pad=12)
    ax.set_xlabel("Date", color="lightgray")
    ax.set_ylabel("Annualized Sharpe Ratio", color="lightgray")
    ax.tick_params(colors="lightgray", labelsize=8)
    for s in ax.spines.values():
        s.set_edgecolor("#2E3440")
    ax.legend(bbox_to_anchor=(1.01, 1), loc='upper left', facecolor="#1A1D2B", edgecolor="#2E3440", labelcolor="white", fontsize=8)
    
    plt.savefig(str(REPORTS_DIR / "rolling_sharpe_chart.png"), dpi=150, bbox_inches="tight", facecolor="#10121C")
    plt.close()
    
    logger.info("Visual charts successfully saved to reports/.")

def main():
    logger.info("Starting Metrics & Risk Ratios Computation...")
    
    df_nav, df_perf_snapshots, df_txn, df_master = load_data_from_db()
    nifty50_df, nifty100_df = load_benchmarks()
    
    df_metrics = compute_performance_ratios(df_nav, df_perf_snapshots, nifty100_df)
    generate_fund_scorecard(df_metrics)
    compute_advanced_risk_metrics(df_nav, df_txn)
    generate_visuals(df_nav, nifty100_df)
    
    logger.info("Quantitative Analytics complete. All CSVs saved to data/processed/ and plots saved to reports/.")

if __name__ == "__main__":
    main()
