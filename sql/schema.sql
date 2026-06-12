-- =============================================================================
-- schema.sql
-- =============================================================================
-- Mutual Fund Analytics Capstone Project -- Day 2
-- -----------------------------------------------------------------------------
-- Purpose  : Define the complete SQLite star schema for the mutual fund
--            analytics database (mutual_funds.db).
--
-- Schema   : Star schema with 4 dimension tables and 3 fact tables.
-- Author   : Capstone Project -- Senior Data Engineer
-- Created  : 2026-06-12
-- Standard : ANSI SQL (SQLite compatible)
--
-- Execution:
--     sqlite3 mutual_funds.db < sql/schema.sql
-- OR via Python:
--     python scripts/load_to_sqlite.py  (runs this file automatically)
--
-- Table Map:
--     Dimensions  : dim_amc, dim_scheme, dim_date, dim_investor
--     Facts       : fact_nav_history, fact_investor_transactions,
--                   fact_scheme_performance
-- =============================================================================

PRAGMA journal_mode = WAL;          -- Write-Ahead Logging for better concurrency
PRAGMA foreign_keys = ON;           -- Enforce referential integrity
PRAGMA synchronous  = NORMAL;       -- Balance safety and speed

-- =============================================================================
-- DROP EXISTING TABLES (safe re-run)
-- =============================================================================

DROP TABLE IF EXISTS fact_scheme_performance;
DROP TABLE IF EXISTS fact_investor_transactions;
DROP TABLE IF EXISTS fact_nav_history;
DROP TABLE IF EXISTS dim_investor;
DROP TABLE IF EXISTS dim_date;
DROP TABLE IF EXISTS dim_scheme;
DROP TABLE IF EXISTS dim_amc;

-- =============================================================================
-- DIMENSION TABLES
-- =============================================================================

-- -----------------------------------------------------------------------------
-- dim_amc
-- Asset Management Company master dimension.
-- One row per unique AMC.
-- -----------------------------------------------------------------------------
CREATE TABLE dim_amc (
    amc_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    amc_code        TEXT    NOT NULL UNIQUE,            -- Short code e.g. 'HDFC'
    amc_name        TEXT    NOT NULL,                   -- Full name
    amc_country     TEXT    NOT NULL DEFAULT 'India',
    sebi_reg_no     TEXT,                               -- SEBI registration number
    website         TEXT,
    founded_year    INTEGER,
    is_active       INTEGER NOT NULL DEFAULT 1          -- 1=active, 0=inactive
                    CHECK (is_active IN (0, 1)),
    created_at      TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    updated_at      TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

CREATE INDEX idx_amc_code ON dim_amc (amc_code);

-- -----------------------------------------------------------------------------
-- dim_scheme
-- Mutual fund scheme master dimension.
-- One row per unique scheme.
-- -----------------------------------------------------------------------------
CREATE TABLE dim_scheme (
    scheme_code     INTEGER PRIMARY KEY,                -- AMFI scheme code (natural key)
    scheme_name     TEXT    NOT NULL,
    amc_id          INTEGER NOT NULL
                    REFERENCES dim_amc (amc_id)
                    ON DELETE RESTRICT ON UPDATE CASCADE,
    category        TEXT    NOT NULL,                   -- e.g. 'Equity', 'Debt', 'Hybrid'
    sub_category    TEXT    NOT NULL,                   -- e.g. 'Large Cap', 'Flexi Cap'
    scheme_type     TEXT    NOT NULL DEFAULT 'Open Ended',
    benchmark       TEXT,                               -- e.g. 'NIFTY 50'
    launch_date     TEXT,                               -- ISO date string
    is_direct       INTEGER NOT NULL DEFAULT 1
                    CHECK (is_direct IN (0, 1)),        -- 1=Direct, 0=Regular
    is_active       INTEGER NOT NULL DEFAULT 1
                    CHECK (is_active IN (0, 1)),
    created_at      TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    updated_at      TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

CREATE INDEX idx_scheme_amc    ON dim_scheme (amc_id);
CREATE INDEX idx_scheme_cat    ON dim_scheme (category);
CREATE INDEX idx_scheme_subcat ON dim_scheme (sub_category);

-- -----------------------------------------------------------------------------
-- dim_date
-- Pre-populated date dimension spanning the analytics window.
-- One row per calendar date.
-- -----------------------------------------------------------------------------
CREATE TABLE dim_date (
    date_id         INTEGER PRIMARY KEY,                -- Surrogate: YYYYMMDD integer
    full_date       TEXT    NOT NULL UNIQUE,            -- ISO date: 'YYYY-MM-DD'
    day             INTEGER NOT NULL CHECK (day BETWEEN 1 AND 31),
    month           INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    month_name      TEXT    NOT NULL,                   -- 'January', 'February' ...
    quarter         INTEGER NOT NULL CHECK (quarter BETWEEN 1 AND 4),
    year            INTEGER NOT NULL,
    day_of_week     INTEGER NOT NULL CHECK (day_of_week BETWEEN 0 AND 6),
                                                        -- 0=Mon ... 6=Sun
    day_name        TEXT    NOT NULL,                   -- 'Monday' ... 'Sunday'
    is_weekend      INTEGER NOT NULL DEFAULT 0
                    CHECK (is_weekend IN (0, 1)),
    is_month_end    INTEGER NOT NULL DEFAULT 0
                    CHECK (is_month_end IN (0, 1)),
    is_quarter_end  INTEGER NOT NULL DEFAULT 0
                    CHECK (is_quarter_end IN (0, 1)),
    is_year_end     INTEGER NOT NULL DEFAULT 0
                    CHECK (is_year_end IN (0, 1)),
    fiscal_year     INTEGER NOT NULL,                   -- Indian FY: Apr-Mar
    fiscal_quarter  INTEGER NOT NULL CHECK (fiscal_quarter BETWEEN 1 AND 4)
);

CREATE INDEX idx_date_full    ON dim_date (full_date);
CREATE INDEX idx_date_year    ON dim_date (year);
CREATE INDEX idx_date_quarter ON dim_date (quarter, year);

-- -----------------------------------------------------------------------------
-- dim_investor
-- Investor profile dimension.
-- One row per unique investor.
-- -----------------------------------------------------------------------------
CREATE TABLE dim_investor (
    investor_id     TEXT    PRIMARY KEY,                -- e.g. 'INV00001'
    investor_name   TEXT    NOT NULL,
    investor_type   TEXT    NOT NULL
                    CHECK (investor_type IN ('Individual', 'HUF', 'Corporate', 'Unknown')),
    city            TEXT,
    state           TEXT,
    country         TEXT    NOT NULL DEFAULT 'India',
    risk_profile    TEXT
                    CHECK (risk_profile IN ('Conservative', 'Moderate',
                                            'Aggressive', 'Unknown', NULL)),
    kyc_status      TEXT    NOT NULL DEFAULT 'Verified'
                    CHECK (kyc_status IN ('Verified', 'Pending', 'Rejected')),
    is_active       INTEGER NOT NULL DEFAULT 1
                    CHECK (is_active IN (0, 1)),
    created_at      TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    updated_at      TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

CREATE INDEX idx_investor_type  ON dim_investor (investor_type);
CREATE INDEX idx_investor_state ON dim_investor (state);

-- =============================================================================
-- FACT TABLES
-- =============================================================================

-- -----------------------------------------------------------------------------
-- fact_nav_history
-- Daily NAV records for each scheme.
-- Grain: one row per scheme per trading day.
-- -----------------------------------------------------------------------------
CREATE TABLE fact_nav_history (
    nav_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    scheme_code     INTEGER NOT NULL
                    REFERENCES dim_scheme (scheme_code)
                    ON DELETE RESTRICT ON UPDATE CASCADE,
    date_id         INTEGER NOT NULL
                    REFERENCES dim_date (date_id)
                    ON DELETE RESTRICT ON UPDATE CASCADE,
    nav             REAL    NOT NULL CHECK (nav > 0),
    repurchase_price REAL   CHECK (repurchase_price IS NULL OR repurchase_price > 0),
    sale_price       REAL   CHECK (sale_price IS NULL OR sale_price > 0),
    nav_change      REAL,                               -- Absolute change from prev day
    nav_pct_change  REAL,                               -- % change from prev day
    source          TEXT    NOT NULL DEFAULT 'AMFI',
    loaded_at       TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    UNIQUE (scheme_code, date_id)                       -- Natural key constraint
);

CREATE INDEX idx_nav_scheme ON fact_nav_history (scheme_code);
CREATE INDEX idx_nav_date   ON fact_nav_history (date_id);
CREATE INDEX idx_nav_scheme_date ON fact_nav_history (scheme_code, date_id);

-- -----------------------------------------------------------------------------
-- fact_investor_transactions
-- Each investor buy/sell/switch/SIP transaction.
-- Grain: one row per transaction.
-- -----------------------------------------------------------------------------
CREATE TABLE fact_investor_transactions (
    txn_id          TEXT    PRIMARY KEY,                -- e.g. 'TXN0000001'
    investor_id     TEXT    NOT NULL
                    REFERENCES dim_investor (investor_id)
                    ON DELETE RESTRICT ON UPDATE CASCADE,
    scheme_code     INTEGER NOT NULL
                    REFERENCES dim_scheme (scheme_code)
                    ON DELETE RESTRICT ON UPDATE CASCADE,
    date_id         INTEGER NOT NULL
                    REFERENCES dim_date (date_id)
                    ON DELETE RESTRICT ON UPDATE CASCADE,
    transaction_type TEXT   NOT NULL
                    CHECK (transaction_type IN ('SIP', 'Lumpsum', 'Redemption',
                                               'Switch-In', 'Switch-Out', 'Unknown')),
    amount_inr      REAL    NOT NULL CHECK (amount_inr >= 0),
    units           REAL    NOT NULL CHECK (units >= 0),
    nav_at_txn      REAL    NOT NULL CHECK (nav_at_txn > 0),
    loaded_at       TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now'))
);

CREATE INDEX idx_txn_investor ON fact_investor_transactions (investor_id);
CREATE INDEX idx_txn_scheme   ON fact_investor_transactions (scheme_code);
CREATE INDEX idx_txn_date     ON fact_investor_transactions (date_id);
CREATE INDEX idx_txn_type     ON fact_investor_transactions (transaction_type);
CREATE INDEX idx_txn_investor_scheme ON fact_investor_transactions (investor_id, scheme_code);

-- -----------------------------------------------------------------------------
-- fact_scheme_performance
-- Quarterly scheme performance metrics (returns, AUM, ratios).
-- Grain: one row per scheme per quarter-end date.
-- -----------------------------------------------------------------------------
CREATE TABLE fact_scheme_performance (
    perf_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    scheme_code     INTEGER NOT NULL
                    REFERENCES dim_scheme (scheme_code)
                    ON DELETE RESTRICT ON UPDATE CASCADE,
    date_id         INTEGER NOT NULL
                    REFERENCES dim_date (date_id)
                    ON DELETE RESTRICT ON UPDATE CASCADE,
    returns_1m      REAL,                               -- 1-month return (%)
    returns_3m      REAL,                               -- 3-month return (%)
    returns_6m      REAL,                               -- 6-month return (%)
    returns_1y      REAL,                               -- 1-year return (%)
    returns_3y      REAL,                               -- 3-year CAGR (%)
    returns_5y      REAL,                               -- 5-year CAGR (%)
    sharpe_ratio    REAL,                               -- Risk-adjusted return
    expense_ratio   REAL    CHECK (expense_ratio IS NULL OR
                                  (expense_ratio BETWEEN 0.01 AND 3.00)),
    aum_cr          REAL    CHECK (aum_cr IS NULL OR aum_cr >= 0),
                                                        -- Assets Under Mgmt (Crores)
    benchmark       TEXT,
    loaded_at       TEXT    NOT NULL
                    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S', 'now')),
    UNIQUE (scheme_code, date_id)
);

CREATE INDEX idx_perf_scheme ON fact_scheme_performance (scheme_code);
CREATE INDEX idx_perf_date   ON fact_scheme_performance (date_id);

-- =============================================================================
-- VIEWS  (pre-built for common analytical queries)
-- =============================================================================

-- Latest NAV per scheme
CREATE VIEW IF NOT EXISTS v_latest_nav AS
SELECT
    s.scheme_code,
    s.scheme_name,
    a.amc_name,
    s.category,
    s.sub_category,
    d.full_date    AS nav_date,
    n.nav,
    n.nav_change,
    n.nav_pct_change
FROM fact_nav_history  n
JOIN dim_scheme        s ON s.scheme_code = n.scheme_code
JOIN dim_amc           a ON a.amc_id      = s.amc_id
JOIN dim_date          d ON d.date_id     = n.date_id
WHERE n.date_id = (
    SELECT MAX(date_id) FROM fact_nav_history h2
    WHERE h2.scheme_code = n.scheme_code
);

-- Monthly transaction summary
CREATE VIEW IF NOT EXISTS v_monthly_transactions AS
SELECT
    d.year,
    d.month,
    d.month_name,
    t.transaction_type,
    s.category,
    COUNT(*)            AS txn_count,
    SUM(t.amount_inr)   AS total_amount_inr,
    SUM(t.units)        AS total_units,
    AVG(t.amount_inr)   AS avg_amount_inr
FROM fact_investor_transactions t
JOIN dim_date    d ON d.date_id     = t.date_id
JOIN dim_scheme  s ON s.scheme_code = t.scheme_code
GROUP BY d.year, d.month, d.month_name, t.transaction_type, s.category;

-- Scheme performance summary (latest quarter)
CREATE VIEW IF NOT EXISTS v_scheme_performance_latest AS
SELECT
    s.scheme_code,
    s.scheme_name,
    a.amc_name,
    s.category,
    s.sub_category,
    d.full_date       AS as_of_date,
    p.returns_1m,
    p.returns_3m,
    p.returns_6m,
    p.returns_1y,
    p.returns_3y,
    p.returns_5y,
    p.sharpe_ratio,
    p.expense_ratio,
    p.aum_cr
FROM fact_scheme_performance p
JOIN dim_scheme  s ON s.scheme_code = p.scheme_code
JOIN dim_amc     a ON a.amc_id      = s.amc_id
JOIN dim_date    d ON d.date_id     = p.date_id
WHERE p.date_id = (
    SELECT MAX(date_id) FROM fact_scheme_performance p2
    WHERE p2.scheme_code = p.scheme_code
);

-- =============================================================================
-- END OF SCHEMA
-- =============================================================================
