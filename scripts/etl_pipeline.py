"""
etl_pipeline.py
===============
Mutual Fund Analytics Capstone Project -- Day 1-3 Ingestion, Cleaning & Loading
-----------------------------------------------------------------------------
Purpose : Ingest raw datasets, run cleaning/validation routines, and load
          the star schema dimensional model into SQLite (data/db/mutual_funds.db).
"""

import os
import sys
import string
import random
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
from sqlalchemy import create_engine, text

# Setup paths
PROJECT_ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = PROJECT_ROOT / "data" / "raw"
PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
DB_DIR = PROJECT_ROOT / "data" / "db"
SQL_DIR = PROJECT_ROOT / "sql"
SCHEMA_PATH = SQL_DIR / "schema.sql"

for d in (RAW_DIR, PROCESSED_DIR, DB_DIR, SQL_DIR):
    d.mkdir(parents=True, exist_ok=True)

DB_PATH = DB_DIR / "mutual_funds.db"

# Configure Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s -- %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("etl_pipeline")

# Reference Data
SCHEMES = [
    {"scheme_code": 119551, "scheme_name": "Axis Bluechip Fund - Direct Growth", "category": "Equity", "sub_category": "Large Cap", "amc_code": "AXIS", "benchmark": "NIFTY 50"},
    {"scheme_code": 120503, "scheme_name": "Mirae Asset Large Cap Fund - Direct Growth", "category": "Equity", "sub_category": "Large Cap", "amc_code": "MIRAE", "benchmark": "NIFTY 50"},
    {"scheme_code": 122639, "scheme_name": "SBI Small Cap Fund - Direct Growth", "category": "Equity", "sub_category": "Small Cap", "amc_code": "SBI", "benchmark": "NIFTY SMALLCAP 100"},
    {"scheme_code": 125497, "scheme_name": "HDFC Mid-Cap Opportunities Fund - Direct Growth", "category": "Equity", "sub_category": "Mid Cap", "amc_code": "HDFC", "benchmark": "NIFTY MIDCAP 100"},
    {"scheme_code": 118989, "scheme_name": "ICICI Prudential Technology Fund - Direct Growth", "category": "Equity", "sub_category": "Sectoral", "amc_code": "ICICI", "benchmark": "NIFTY IT"},
    {"scheme_code": 119598, "scheme_name": "Kotak Bluechip Fund - Direct Growth", "category": "Equity", "sub_category": "Large Cap", "amc_code": "KOTAK", "benchmark": "NIFTY 50"},
    {"scheme_code": 120465, "scheme_name": "Nippon India Large Cap Fund - Direct Growth", "category": "Equity", "sub_category": "Large Cap", "amc_code": "NIPPON", "benchmark": "NIFTY 50"},
    {"scheme_code": 120594, "scheme_name": "DSP Top 100 Equity Fund - Direct Growth", "category": "Equity", "sub_category": "Large Cap", "amc_code": "DSP", "benchmark": "NIFTY 50"},
    {"scheme_code": 120716, "scheme_name": "Canara Robeco Bluechip Equity Fund - Direct Growth", "category": "Equity", "sub_category": "Large Cap", "amc_code": "CANARA", "benchmark": "NIFTY 50"},
    {"scheme_code": 135781, "scheme_name": "Parag Parikh Flexi Cap Fund - Direct Growth", "category": "Equity", "sub_category": "Flexi Cap", "amc_code": "PPFAS", "benchmark": "NIFTY 500"},
]

AMC_DATA = [
    {"amc_code": "AXIS",   "amc_name": "Axis Mutual Fund",              "founded_year": 2009},
    {"amc_code": "MIRAE",  "amc_name": "Mirae Asset Mutual Fund",       "founded_year": 2008},
    {"amc_code": "SBI",    "amc_name": "SBI Mutual Fund",               "founded_year": 1987},
    {"amc_code": "HDFC",   "amc_name": "HDFC Mutual Fund",              "founded_year": 2000},
    {"amc_code": "ICICI",  "amc_name": "ICICI Prudential Mutual Fund",  "founded_year": 1993},
    {"amc_code": "KOTAK",  "amc_name": "Kotak Mahindra Mutual Fund",    "founded_year": 1998},
    {"amc_code": "NIPPON", "amc_name": "Nippon India Mutual Fund",      "founded_year": 1995},
    {"amc_code": "DSP",    "amc_name": "DSP Investment Managers",       "founded_year": 1996},
    {"amc_code": "CANARA", "amc_name": "Canara Robeco Mutual Fund",     "founded_year": 1993},
    {"amc_code": "PPFAS",  "amc_name": "PPFAS Mutual Fund",             "founded_year": 2013},
]

# =============================================================================
# Synthetic Data Generators (Runs if raw CSVs are missing)
# =============================================================================

def generate_synthetic_raw():
    logger.info("Raw files missing. Generating synthetic raw data...")
    random.seed(42)
    np.random.seed(42)
    
    # 1. NAV History
    dates = pd.date_range(start="2022-01-01", end="2024-12-31", freq="D")
    nav_rows = []
    for scheme in SCHEMES:
        nav = round(random.uniform(50, 300), 4)
        for date in dates:
            daily_return = np.random.normal(loc=0.0004, scale=0.012)
            nav = max(round(nav * (1 + daily_return), 4), 1.0)
            nav_rows.append({
                "Scheme Code": scheme["scheme_code"],
                "Scheme Name": scheme["scheme_name"],
                "NAV Date": date.strftime("%d-%b-%Y"),
                "NAV": nav,
                "Repurchase Price": round(nav * 0.9985, 4),
                "Sale Price": round(nav * 1.0, 4),
                "AMC Code": scheme["amc_code"],
                "Category": scheme["category"],
                "Sub Category": scheme["sub_category"]
            })
    df_nav = pd.DataFrame(nav_rows)
    
    # Inject bad dates / invalid scheme code lines intentionally for cleaning demonstration
    idx = np.random.choice(df_nav.index, size=15, replace=False)
    df_nav.loc[idx[:5], "NAV"] = np.nan
    df_nav.loc[idx[5:10], "NAV Date"] = "BAD_DATE"
    df_nav.loc[idx[10:], "Scheme Code"] = -999
    df_nav.to_csv(RAW_DIR / "nav_history.csv", index=False)
    logger.info("Saved raw/nav_history.csv")

    # 2. Investor Transactions
    cities = ["Mumbai", "Delhi", "Bengaluru", "Hyderabad", "Chennai", "Pune", "Kolkata", "Ahmedabad", "Jaipur", "Lucknow"]
    states = ["Maharashtra", "Delhi", "Karnataka", "Telangana", "Tamil Nadu", "Maharashtra", "West Bengal", "Gujarat", "Rajasthan", "Uttar Pradesh"]
    investors = [{"id": f"INV{i:04d}", "name": f"Investor_{i}", "city": cities[i%10], "state": states[i%10], "risk": random.choice(["Conservative", "Moderate", "Aggressive"])} for i in range(1, 201)]
    
    txn_rows = []
    base_dt = datetime(2022, 1, 1)
    for i in range(1, 2001):
        inv = random.choice(investors)
        scheme = random.choice(SCHEMES)
        txn_type = random.choice(["SIP", "Lumpsum", "Redemption", "Switch-In", "Switch-Out"])
        amount = round(random.choice([500, 1000, 2000, 5000, 10000, 50000]) * random.uniform(0.9, 1.1), 2)
        nav_val = round(random.uniform(50, 350), 4)
        units = round(amount / nav_val, 4)
        txn_date = base_dt + timedelta(days=random.randint(0, 1095))
        txn_rows.append({
            "Transaction ID": f"TXN{i:07d}",
            "Investor ID": inv["id"],
            "Investor Name": inv["name"],
            "Investor Type": "Individual",
            "City": inv["city"],
            "State": inv["state"],
            "Risk Profile": inv["risk"],
            "Scheme Code": scheme["scheme_code"],
            "Scheme Name": scheme["scheme_name"],
            "AMC Code": scheme["amc_code"],
            "Transaction Type": txn_type,
            "Transaction Date": txn_date.strftime("%d-%b-%Y"),
            "Amount (INR)": amount,
            "Units": units,
            "NAV at Transaction": nav_val
        })
    df_txn = pd.DataFrame(txn_rows)
    # Inject bad dates / null amounts to demonstrate cleaning
    idx_txn = np.random.choice(df_txn.index, size=20, replace=False)
    df_txn.loc[idx_txn[:10], "Transaction Date"] = "INVALID"
    df_txn.loc[idx_txn[10:], "Amount (INR)"] = np.nan
    df_txn.to_csv(RAW_DIR / "investor_transactions.csv", index=False)
    logger.info("Saved raw/investor_transactions.csv")

    # 3. Scheme Performance
    perf_rows = []
    quarters = pd.date_range(start="2022-01-01", end="2024-12-31", freq="QE")
    for scheme in SCHEMES:
        for q in quarters:
            base_ret = np.random.normal(0.02, 0.05)
            perf_rows.append({
                "Scheme Code": scheme["scheme_code"],
                "Scheme Name": scheme["scheme_name"],
                "AMC Code": scheme["amc_code"],
                "Category": scheme["category"],
                "Sub Category": scheme["sub_category"],
                "As of Date": q.strftime("%d-%b-%Y"),
                "Returns 1M (%)": round(base_ret * 0.5 + np.random.normal(0, 0.02), 4),
                "Returns 3M (%)": round(base_ret * 1.5 + np.random.normal(0, 0.03), 4),
                "Returns 6M (%)": round(base_ret * 3.0 + np.random.normal(0, 0.04), 4),
                "Returns 1Y (%)": round(base_ret * 6.0 + np.random.normal(0, 0.06), 4),
                "Returns 3Y (%)": round(base_ret * 18 + np.random.normal(0, 0.08), 4),
                "Returns 5Y (%)": round(base_ret * 30 + np.random.normal(0, 0.10), 4),
                "Sharpe Ratio": round(np.random.uniform(0.3, 2.5), 4),
                "Expense Ratio (%)": round(np.random.uniform(0.10, 1.20), 4),
                "AUM (Cr)": round(np.random.uniform(500, 50000), 2),
                "Benchmark": scheme["benchmark"]
            })
    df_perf = pd.DataFrame(perf_rows)
    df_perf.to_csv(RAW_DIR / "scheme_performance.csv", index=False)
    logger.info("Saved raw/scheme_performance.csv")

# =============================================================================
# Ingestion & Cleaning Engine
# =============================================================================

def clean_data():
    logger.info("Running data cleaning pipeline...")
    
    # 1. Clean NAV History
    df_nav = pd.read_csv(RAW_DIR / "nav_history.csv")
    df_nav.columns = df_nav.columns.str.strip().str.lower().str.replace(r"[^a-z0-9]+", "_", regex=True).str.strip("_")
    df_nav.drop_duplicates(subset=["scheme_code", "nav_date"], keep="first", inplace=True)
    df_nav["nav_date"] = pd.to_datetime(df_nav["nav_date"], format="%d-%b-%Y", errors="coerce")
    df_nav.dropna(subset=["nav_date", "nav", "scheme_code"], inplace=True)
    df_nav = df_nav[df_nav["scheme_code"] > 0]
    df_nav = df_nav[df_nav["nav_date"].dt.dayofweek < 5] # Drop weekends
    df_nav.to_csv(PROCESSED_DIR / "nav_history_clean.csv", index=False)
    df_nav.to_parquet(PROCESSED_DIR / "nav_history_clean.parquet", index=False)
    logger.info(f"Cleaned NAV History: {len(df_nav)} rows saved.")

    # 2. Clean Transactions
    df_txn = pd.read_csv(RAW_DIR / "investor_transactions.csv")
    df_txn.columns = df_txn.columns.str.strip().str.lower().str.replace(r"[^a-z0-9]+", "_", regex=True).str.strip("_")
    df_txn.drop_duplicates(subset=["transaction_id"], keep="first", inplace=True)
    df_txn["transaction_date"] = pd.to_datetime(df_txn["transaction_date"], format="%d-%b-%Y", errors="coerce")
    df_txn.dropna(subset=["transaction_date", "amount_inr", "units"], inplace=True)
    df_txn = df_txn[(df_txn["amount_inr"] > 0) & (df_txn["units"] > 0)]
    df_txn.to_csv(PROCESSED_DIR / "investor_transactions_clean.csv", index=False)
    df_txn.to_parquet(PROCESSED_DIR / "investor_transactions_clean.parquet", index=False)
    logger.info(f"Cleaned Transactions: {len(df_txn)} rows saved.")

    # 3. Clean Performance
    df_perf = pd.read_csv(RAW_DIR / "scheme_performance.csv")
    df_perf.columns = df_perf.columns.str.strip().str.lower().str.replace(r"[^a-z0-9]+", "_", regex=True).str.strip("_")
    df_perf.drop_duplicates(subset=["scheme_code", "as_of_date"], keep="last", inplace=True)
    df_perf["as_of_date"] = pd.to_datetime(df_perf["as_of_date"], format="%d-%b-%Y", errors="coerce")
    df_perf.dropna(subset=["as_of_date", "scheme_code"], inplace=True)
    df_perf.to_csv(PROCESSED_DIR / "scheme_performance_clean.csv", index=False)
    df_perf.to_parquet(PROCESSED_DIR / "scheme_performance_clean.parquet", index=False)
    logger.info(f"Cleaned Performance snapshots: {len(df_perf)} rows saved.")

# =============================================================================
# SQLite Load Star Schema Engine
# =============================================================================

def apply_schema(engine):
    if not SCHEMA_PATH.exists():
        raise FileNotFoundError(f"schema.sql not found at {SCHEMA_PATH}")
    
    with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
        ddl = f.read()
    
    statements = [s.strip() for s in ddl.split(";") if s.strip()]
    with engine.begin() as conn:
        for stmt in statements:
            # dim_date and its index is populated and dropped via pandas directly
            if "CREATE TABLE dim_date" in stmt or "idx_date" in stmt:
                continue
            try:
                conn.execute(text(stmt))
            except Exception as e:
                logger.warning(f"Non-fatal schema load warning: {str(e)[:80]}")
    logger.info("schema.sql applied to database.")

def populate_dim_date(engine):
    logger.info("Populating dim_date (2020-2030)...")
    dates = pd.date_range(start="2020-01-01", end="2030-12-31", freq="D")
    month_names = {1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June",
                   7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December"}
    
    rows = []
    for dt in dates:
        month = dt.month
        year = dt.year
        fy = year if month >= 4 else year - 1
        fq_map = {4: 1, 5: 1, 6: 1, 7: 2, 8: 2, 9: 2, 10: 3, 11: 3, 12: 3, 1: 4, 2: 4, 3: 4}
        rows.append({
            "date_id": int(dt.strftime("%Y%m%d")),
            "full_date": dt.strftime("%Y-%m-%d"),
            "day": dt.day,
            "month": month,
            "month_name": month_names[month],
            "quarter": dt.quarter,
            "year": year,
            "day_of_week": dt.dayofweek,
            "day_name": dt.day_name(),
            "is_weekend": int(dt.dayofweek >= 5),
            "is_month_end": int(dt == dt + pd.offsets.MonthEnd(0)),
            "is_quarter_end": int(dt == dt + pd.offsets.QuarterEnd(0)),
            "is_year_end": int(dt.month == 12 and dt.day == 31),
            "fiscal_year": fy,
            "fiscal_quarter": fq_map[month]
        })
    df_date = pd.DataFrame(rows)
    with engine.begin() as conn:
        conn.execute(text("DROP TABLE IF EXISTS dim_date"))
    df_date.to_sql("dim_date", engine, if_exists="replace", index=False, chunksize=100)
    logger.info(f"dim_date populated: {len(df_date)} rows.")

def load_star_schema(engine):
    logger.info("Loading cleaned datasets into dimensions and fact tables...")
    
    # 1. dim_amc
    df_amc = pd.DataFrame(AMC_DATA)
    df_amc.insert(0, "amc_id", range(1, len(df_amc) + 1))
    df_amc["amc_country"] = "India"
    df_amc["is_active"] = 1
    df_amc["created_at"] = datetime.now().isoformat(timespec="seconds")
    df_amc["updated_at"] = datetime.now().isoformat(timespec="seconds")
    df_amc.to_sql("dim_amc", engine, if_exists="replace", index=False)
    
    amc_map = {row["amc_code"]: row["amc_id"] for _, row in df_amc.iterrows()}
    logger.info("dim_amc loaded.")
    
    # 2. dim_scheme
    scheme_rows = []
    for s in SCHEMES:
        scheme_rows.append({
            "scheme_code": s["scheme_code"],
            "scheme_name": s["scheme_name"],
            "amc_id": amc_map.get(s["amc_code"], 1),
            "category": s["category"],
            "sub_category": s["sub_category"],
            "scheme_type": "Open Ended",
            "benchmark": s["benchmark"],
            "launch_date": None,
            "is_direct": 1,
            "is_active": 1,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "updated_at": datetime.now().isoformat(timespec="seconds")
        })
    df_scheme = pd.DataFrame(scheme_rows)
    df_scheme.to_sql("dim_scheme", engine, if_exists="replace", index=False)
    logger.info("dim_scheme loaded.")
    
    # Date ID mapping helper
    with engine.connect() as conn:
        res = conn.execute(text("SELECT date_id, full_date FROM dim_date"))
        date_map = {row.full_date: row.date_id for row in res}
        
    def get_date_id(dt_series):
        return pd.to_datetime(dt_series).dt.strftime("%Y-%m-%d").map(date_map)

    # 3. dim_investor (extracted from transactions)
    df_txn_clean = pd.read_csv(PROCESSED_DIR / "investor_transactions_clean.csv")
    df_inv = df_txn_clean[["investor_id", "investor_name", "investor_type", "city", "state", "risk_profile"]].drop_duplicates(subset=["investor_id"]).copy()
    df_inv["country"] = "India"
    df_inv["kyc_status"] = "Verified"
    df_inv["is_active"] = 1
    df_inv["created_at"] = datetime.now().isoformat(timespec="seconds")
    df_inv["updated_at"] = datetime.now().isoformat(timespec="seconds")
    df_inv.to_sql("dim_investor", engine, if_exists="replace", index=False)
    logger.info("dim_investor loaded.")

    # 4. fact_nav_history
    df_nav_clean = pd.read_csv(PROCESSED_DIR / "nav_history_clean.csv")
    df_nav_clean["date_id"] = get_date_id(df_nav_clean["nav_date"])
    df_nav_clean.dropna(subset=["date_id"], inplace=True)
    
    df_nav_clean.sort_values(["scheme_code", "nav_date"], inplace=True)
    df_nav_clean["nav_change"] = df_nav_clean.groupby("scheme_code")["nav"].diff().round(4)
    df_nav_clean["nav_pct_change"] = df_nav_clean.groupby("scheme_code")["nav"].pct_change().mul(100).round(4)
    
    fact_nav = df_nav_clean[["scheme_code", "date_id", "nav", "repurchase_price", "sale_price", "nav_change", "nav_pct_change"]].copy()
    fact_nav["source"] = "AMFI"
    fact_nav["loaded_at"] = datetime.now().isoformat(timespec="seconds")
    fact_nav.to_sql("fact_nav_history", engine, if_exists="replace", index=False)
    logger.info(f"fact_nav_history loaded: {len(fact_nav)} rows.")

    # 5. fact_investor_transactions
    df_txn_clean["date_id"] = get_date_id(df_txn_clean["transaction_date"])
    df_txn_clean.dropna(subset=["date_id"], inplace=True)
    
    fact_txn = df_txn_clean.rename(columns={
        "transaction_id": "txn_id",
        "amount_inr": "amount_inr",
        "nav_at_transaction": "nav_at_txn"
    })[["txn_id", "investor_id", "scheme_code", "date_id", "transaction_type", "amount_inr", "units", "nav_at_txn"]].copy()
    fact_txn["loaded_at"] = datetime.now().isoformat(timespec="seconds")
    fact_txn.to_sql("fact_investor_transactions", engine, if_exists="replace", index=False)
    logger.info(f"fact_investor_transactions loaded: {len(fact_txn)} rows.")

    # 6. fact_scheme_performance
    df_perf_clean = pd.read_csv(PROCESSED_DIR / "scheme_performance_clean.csv")
    df_perf_clean["date_id"] = get_date_id(df_perf_clean["as_of_date"])
    df_perf_clean.dropna(subset=["date_id"], inplace=True)
    
    fact_perf = df_perf_clean.rename(columns={
        "returns_1y": "returns_1y",
        "returns_3y": "returns_3y",
        "returns_5y": "returns_5y",
        "expense_ratio": "expense_ratio",
        "aum_cr": "aum_cr"
    })[["scheme_code", "date_id", "returns_1y", "returns_3y", "returns_5y", "sharpe_ratio", "expense_ratio", "aum_cr", "benchmark"]].copy()
    fact_perf["loaded_at"] = datetime.now().isoformat(timespec="seconds")
    fact_perf.to_sql("fact_scheme_performance", engine, if_exists="replace", index=False)
    logger.info(f"fact_scheme_performance loaded: {len(fact_perf)} rows.")

def run_validation(engine):
    logger.info("=" * 60)
    logger.info("ETL PIPELINE VALIDATION CHECKS")
    logger.info("=" * 60)
    with engine.connect() as conn:
        for t in ["dim_amc", "dim_scheme", "dim_date", "dim_investor", "fact_nav_history", "fact_investor_transactions", "fact_scheme_performance"]:
            cnt = conn.execute(text(f"SELECT COUNT(*) FROM {t}")).scalar()
            logger.info(f"  {t:<30} : {cnt} rows")

def main():
    logger.info("Starting Master ETL Pipeline...")
    
    if not (RAW_DIR / "nav_history.csv").exists() or not (RAW_DIR / "investor_transactions.csv").exists():
        generate_synthetic_raw()
        
    clean_data()
    
    # Establish engine
    db_url = f"sqlite:///{DB_PATH}"
    engine = create_engine(db_url)
    
    try:
        apply_schema(engine)
        populate_dim_date(engine)
        load_star_schema(engine)
        run_validation(engine)
        logger.info("Master ETL Pipeline successfully executed and loaded database.")
    except Exception as e:
        logger.error(f"Fatal ETL error: {e}", exc_info=True)
        sys.exit(1)
    finally:
        engine.dispose()

if __name__ == "__main__":
    main()
