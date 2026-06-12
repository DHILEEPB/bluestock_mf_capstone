-- =============================================================================
-- queries.sql
-- =============================================================================
-- Mutual Fund Analytics Capstone Project -- Day 2
-- -----------------------------------------------------------------------------
-- Purpose  : Production-ready analytical SQL queries for mutual fund analytics.
--            All queries run against mutual_funds.db (SQLite).
--
-- Author   : Capstone Project -- Senior Data Engineer
-- Created  : 2026-06-12
-- DB       : SQLite 3.x
--
-- Usage    :
--     sqlite3 mutual_funds.db < sql/queries.sql
-- OR individual queries via Python / DB Browser for SQLite.
-- =============================================================================

-- Enable useful SQLite settings
PRAGMA foreign_keys = ON;

-- =============================================================================
-- SECTION 1: NAV ANALYTICS
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Q1. Latest NAV for all active schemes
-- Purpose: Dashboard overview card data
-- -----------------------------------------------------------------------------
SELECT
    scheme_code,
    scheme_name,
    amc_name,
    category,
    sub_category,
    nav_date,
    nav,
    ROUND(nav_change,     4) AS nav_change_abs,
    ROUND(nav_pct_change, 4) AS nav_change_pct
FROM v_latest_nav
ORDER BY amc_name, scheme_name;

-- -----------------------------------------------------------------------------
-- Q2. Top 5 Best Performing Schemes (1-Day NAV % Change)
-- Purpose: Daily performance leaderboard
-- -----------------------------------------------------------------------------
SELECT
    scheme_name,
    amc_name,
    category,
    nav,
    ROUND(nav_pct_change, 4) AS pct_change_1d
FROM v_latest_nav
ORDER BY nav_pct_change DESC
LIMIT 5;

-- -----------------------------------------------------------------------------
-- Q3. 52-Week High & Low NAV per Scheme
-- Purpose: Range analysis for investor decisions
-- -----------------------------------------------------------------------------
SELECT
    s.scheme_code,
    s.scheme_name,
    s.category,
    ROUND(MAX(n.nav), 4)  AS nav_52w_high,
    ROUND(MIN(n.nav), 4)  AS nav_52w_low,
    ROUND(MAX(n.nav) - MIN(n.nav), 4)              AS nav_52w_range,
    ROUND((MAX(n.nav) - MIN(n.nav)) / MIN(n.nav) * 100, 2)
                          AS nav_52w_pct_range
FROM fact_nav_history n
JOIN dim_scheme s ON s.scheme_code = n.scheme_code
JOIN dim_date   d ON d.date_id     = n.date_id
WHERE d.full_date >= DATE('now', '-365 days')
GROUP BY s.scheme_code, s.scheme_name, s.category
ORDER BY nav_52w_pct_range DESC;

-- -----------------------------------------------------------------------------
-- Q4. Month-over-Month NAV Change per Scheme
-- Purpose: Trend analysis for time-series visualisations
-- -----------------------------------------------------------------------------
SELECT
    s.scheme_name,
    d.year,
    d.month_name,
    ROUND(AVG(n.nav), 4)          AS avg_monthly_nav,
    ROUND(MIN(n.nav), 4)          AS min_monthly_nav,
    ROUND(MAX(n.nav), 4)          AS max_monthly_nav,
    COUNT(*)                       AS trading_days
FROM fact_nav_history n
JOIN dim_scheme s ON s.scheme_code = n.scheme_code
JOIN dim_date   d ON d.date_id     = n.date_id
GROUP BY s.scheme_code, s.scheme_name, d.year, d.month
ORDER BY s.scheme_name, d.year, d.month;

-- -----------------------------------------------------------------------------
-- Q5. Daily NAV Volatility (Standard Deviation) -- Rolling 30-day window
-- Note: SQLite lacks STDEV; computed as population std dev manually.
-- Purpose: Risk assessment
-- -----------------------------------------------------------------------------
SELECT
    s.scheme_code,
    s.scheme_name,
    d.full_date,
    ROUND(n.nav, 4)        AS nav,
    ROUND(
        (SELECT SQRT(
            AVG((n2.nav - avg_nav.avg_v) * (n2.nav - avg_nav.avg_v))
        )
        FROM fact_nav_history n2
        JOIN dim_date d2 ON d2.date_id = n2.date_id
        CROSS JOIN (
            SELECT AVG(n3.nav) AS avg_v
            FROM fact_nav_history n3
            JOIN dim_date d3 ON d3.date_id = n3.date_id
            WHERE n3.scheme_code = n.scheme_code
              AND d3.full_date BETWEEN DATE(d.full_date, '-30 days') AND d.full_date
        ) avg_nav
        WHERE n2.scheme_code = n.scheme_code
          AND d2.full_date BETWEEN DATE(d.full_date, '-30 days') AND d.full_date
    ), 4) AS rolling_30d_stddev
FROM fact_nav_history n
JOIN dim_scheme s ON s.scheme_code = n.scheme_code
JOIN dim_date   d ON d.date_id     = n.date_id
WHERE d.full_date >= DATE('now', '-90 days')
ORDER BY s.scheme_name, d.full_date;

-- =============================================================================
-- SECTION 2: SCHEME PERFORMANCE ANALYTICS
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Q6. Best Schemes by 1-Year Returns (Latest Quarter)
-- Purpose: Performance ranking for investors
-- -----------------------------------------------------------------------------
SELECT
    scheme_name,
    amc_name,
    category,
    sub_category,
    as_of_date,
    ROUND(returns_1y,  2) AS returns_1y_pct,
    ROUND(returns_3y,  2) AS returns_3y_cagr,
    ROUND(returns_5y,  2) AS returns_5y_cagr,
    ROUND(sharpe_ratio, 4) AS sharpe_ratio,
    ROUND(expense_ratio, 4) AS expense_ratio_pct,
    ROUND(aum_cr, 2)      AS aum_crores
FROM v_scheme_performance_latest
ORDER BY returns_1y DESC;

-- -----------------------------------------------------------------------------
-- Q7. Category-wise Average Returns (Latest Quarter)
-- Purpose: Asset allocation insights
-- -----------------------------------------------------------------------------
SELECT
    category,
    sub_category,
    COUNT(DISTINCT p.scheme_code)     AS scheme_count,
    ROUND(AVG(p.returns_1m), 2)       AS avg_returns_1m,
    ROUND(AVG(p.returns_3m), 2)       AS avg_returns_3m,
    ROUND(AVG(p.returns_1y), 2)       AS avg_returns_1y,
    ROUND(AVG(p.returns_3y), 2)       AS avg_returns_3y,
    ROUND(AVG(p.sharpe_ratio), 4)     AS avg_sharpe,
    ROUND(SUM(p.aum_cr), 2)           AS total_aum_cr
FROM fact_scheme_performance p
JOIN dim_scheme s ON s.scheme_code = p.scheme_code
JOIN dim_date   d ON d.date_id     = p.date_id
WHERE p.date_id = (SELECT MAX(date_id) FROM fact_scheme_performance)
GROUP BY s.category, s.sub_category
ORDER BY avg_returns_1y DESC;

-- -----------------------------------------------------------------------------
-- Q8. AMC-wise Total AUM and Market Share
-- Purpose: Competitive landscape analysis
-- -----------------------------------------------------------------------------
SELECT
    a.amc_name,
    COUNT(DISTINCT s.scheme_code)            AS total_schemes,
    ROUND(SUM(p.aum_cr), 2)                  AS total_aum_cr,
    ROUND(
        SUM(p.aum_cr) * 100.0 /
        SUM(SUM(p.aum_cr)) OVER (), 2
    )                                        AS market_share_pct
FROM fact_scheme_performance p
JOIN dim_scheme s ON s.scheme_code = p.scheme_code
JOIN dim_amc    a ON a.amc_id      = s.amc_id
WHERE p.date_id = (SELECT MAX(date_id) FROM fact_scheme_performance)
GROUP BY a.amc_id, a.amc_name
ORDER BY total_aum_cr DESC;

-- -----------------------------------------------------------------------------
-- Q9. Risk-Return Scatter Data (Sharpe vs 1Y Returns)
-- Purpose: Data for Plotly scatter visualisation (Day 4)
-- -----------------------------------------------------------------------------
SELECT
    p.scheme_code,
    s.scheme_name,
    s.sub_category,
    a.amc_name,
    ROUND(p.returns_1y,   2) AS returns_1y,
    ROUND(p.sharpe_ratio, 4) AS sharpe_ratio,
    ROUND(p.expense_ratio, 4) AS expense_ratio,
    ROUND(p.aum_cr, 2)       AS aum_cr
FROM v_scheme_performance_latest p
JOIN dim_scheme s ON s.scheme_code = p.scheme_code
JOIN dim_amc    a ON a.amc_id      = s.amc_id
ORDER BY returns_1y DESC;

-- -----------------------------------------------------------------------------
-- Q10. Quarter-over-Quarter Returns Trend per Scheme
-- Purpose: Time-series line chart data
-- -----------------------------------------------------------------------------
SELECT
    s.scheme_name,
    d.year,
    d.quarter,
    d.full_date           AS quarter_end_date,
    ROUND(p.returns_1m, 2) AS returns_1m,
    ROUND(p.returns_3m, 2) AS returns_3m,
    ROUND(p.returns_1y, 2) AS returns_1y,
    ROUND(p.aum_cr, 2)    AS aum_cr
FROM fact_scheme_performance p
JOIN dim_scheme s ON s.scheme_code = p.scheme_code
JOIN dim_date   d ON d.date_id     = p.date_id
ORDER BY s.scheme_name, d.year, d.quarter;

-- =============================================================================
-- SECTION 3: INVESTOR TRANSACTION ANALYTICS
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Q11. Transaction Volume and Value by Type (All Time)
-- Purpose: Business overview KPIs
-- -----------------------------------------------------------------------------
SELECT
    transaction_type,
    COUNT(*)                           AS txn_count,
    ROUND(SUM(amount_inr), 2)          AS total_amount_inr,
    ROUND(AVG(amount_inr), 2)          AS avg_amount_inr,
    ROUND(SUM(units), 4)               AS total_units,
    COUNT(DISTINCT investor_id)        AS unique_investors,
    COUNT(DISTINCT scheme_code)        AS unique_schemes
FROM fact_investor_transactions
GROUP BY transaction_type
ORDER BY total_amount_inr DESC;

-- -----------------------------------------------------------------------------
-- Q12. Monthly SIP Inflow Trend
-- Purpose: SIP growth tracking
-- -----------------------------------------------------------------------------
SELECT
    d.year,
    d.month,
    d.month_name,
    COUNT(*)                        AS sip_count,
    ROUND(SUM(t.amount_inr), 2)     AS sip_inflow_inr,
    COUNT(DISTINCT t.investor_id)   AS unique_sip_investors
FROM fact_investor_transactions t
JOIN dim_date d ON d.date_id = t.date_id
WHERE t.transaction_type = 'SIP'
GROUP BY d.year, d.month, d.month_name
ORDER BY d.year, d.month;

-- -----------------------------------------------------------------------------
-- Q13. Top 10 Investors by Total Investment
-- Purpose: High-value investor identification
-- -----------------------------------------------------------------------------
SELECT
    i.investor_id,
    i.investor_name,
    i.investor_type,
    i.city,
    i.risk_profile,
    COUNT(DISTINCT t.scheme_code)    AS schemes_invested,
    COUNT(*)                         AS total_transactions,
    ROUND(SUM(t.amount_inr), 2)      AS total_invested_inr,
    ROUND(AVG(t.amount_inr), 2)      AS avg_txn_size_inr
FROM fact_investor_transactions t
JOIN dim_investor i ON i.investor_id = t.investor_id
WHERE t.transaction_type IN ('SIP', 'Lumpsum')
GROUP BY i.investor_id, i.investor_name, i.investor_type,
         i.city, i.risk_profile
ORDER BY total_invested_inr DESC
LIMIT 10;

-- -----------------------------------------------------------------------------
-- Q14. State-wise Investment Distribution
-- Purpose: Geographic heatmap data
-- -----------------------------------------------------------------------------
SELECT
    i.state,
    COUNT(DISTINCT t.investor_id)    AS unique_investors,
    COUNT(*)                         AS total_txns,
    ROUND(SUM(t.amount_inr), 2)      AS total_invested_inr,
    ROUND(AVG(t.amount_inr), 2)      AS avg_investment_inr
FROM fact_investor_transactions t
JOIN dim_investor i ON i.investor_id = t.investor_id
GROUP BY i.state
ORDER BY total_invested_inr DESC;

-- -----------------------------------------------------------------------------
-- Q15. Investor Portfolio -- Units Held per Scheme
--      (Net units = Buy units - Sell units per investor-scheme pair)
-- Purpose: Portfolio positions for Day 5 risk calculations
-- -----------------------------------------------------------------------------
SELECT
    t.investor_id,
    i.investor_name,
    t.scheme_code,
    s.scheme_name,
    s.category,
    ROUND(
        SUM(CASE WHEN t.transaction_type IN ('SIP', 'Lumpsum', 'Switch-In')
                 THEN t.units ELSE 0 END) -
        SUM(CASE WHEN t.transaction_type IN ('Redemption', 'Switch-Out')
                 THEN t.units ELSE 0 END),
    4) AS net_units_held,
    ROUND(
        SUM(CASE WHEN t.transaction_type IN ('SIP', 'Lumpsum', 'Switch-In')
                 THEN t.amount_inr ELSE 0 END),
    2) AS total_invested_inr
FROM fact_investor_transactions t
JOIN dim_investor i ON i.investor_id = t.investor_id
JOIN dim_scheme   s ON s.scheme_code = t.scheme_code
GROUP BY t.investor_id, i.investor_name, t.scheme_code, s.scheme_name, s.category
HAVING net_units_held > 0
ORDER BY t.investor_id, net_units_held DESC;

-- =============================================================================
-- SECTION 4: COMBINED / CROSS-TABLE ANALYTICS
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Q16. Current Portfolio Value per Investor
--      (Net units held x Latest NAV)
-- Purpose: Portfolio valuation for investor dashboard
-- -----------------------------------------------------------------------------
SELECT
    pos.investor_id,
    pos.investor_name,
    pos.scheme_code,
    pos.scheme_name,
    pos.net_units_held,
    pos.total_invested_inr,
    ln.nav                                                  AS current_nav,
    ROUND(pos.net_units_held * ln.nav, 2)                  AS current_value_inr,
    ROUND(pos.net_units_held * ln.nav - pos.total_invested_inr, 2) AS pnl_inr,
    ROUND(
        (pos.net_units_held * ln.nav - pos.total_invested_inr)
        / NULLIF(pos.total_invested_inr, 0) * 100, 2
    )                                                       AS pnl_pct
FROM (
    -- Reuse portfolio positions from Q15
    SELECT
        t.investor_id,
        i.investor_name,
        t.scheme_code,
        s.scheme_name,
        ROUND(
            SUM(CASE WHEN t.transaction_type IN ('SIP','Lumpsum','Switch-In')
                     THEN t.units ELSE 0 END) -
            SUM(CASE WHEN t.transaction_type IN ('Redemption','Switch-Out')
                     THEN t.units ELSE 0 END), 4) AS net_units_held,
        ROUND(
            SUM(CASE WHEN t.transaction_type IN ('SIP','Lumpsum','Switch-In')
                     THEN t.amount_inr ELSE 0 END), 2) AS total_invested_inr
    FROM fact_investor_transactions t
    JOIN dim_investor i ON i.investor_id = t.investor_id
    JOIN dim_scheme   s ON s.scheme_code = t.scheme_code
    GROUP BY t.investor_id, i.investor_name, t.scheme_code, s.scheme_name
    HAVING net_units_held > 0
) pos
JOIN v_latest_nav ln ON ln.scheme_code = pos.scheme_code
ORDER BY pos.investor_id, pnl_pct DESC;

-- -----------------------------------------------------------------------------
-- Q17. Category-wise Portfolio Allocation (Aggregate)
-- Purpose: Asset allocation pie chart data
-- -----------------------------------------------------------------------------
SELECT
    s.category,
    s.sub_category,
    ROUND(SUM(t.amount_inr), 2)             AS total_invested_inr,
    ROUND(
        SUM(t.amount_inr) * 100.0 /
        SUM(SUM(t.amount_inr)) OVER (), 2
    )                                       AS allocation_pct
FROM fact_investor_transactions t
JOIN dim_scheme s ON s.scheme_code = t.scheme_code
WHERE t.transaction_type IN ('SIP', 'Lumpsum')
GROUP BY s.category, s.sub_category
ORDER BY total_invested_inr DESC;

-- -----------------------------------------------------------------------------
-- Q18. Data Quality Summary (all tables)
-- Purpose: Data monitoring / observability
-- -----------------------------------------------------------------------------
SELECT 'fact_nav_history'          AS table_name,
       COUNT(*)                    AS total_rows,
       COUNT(DISTINCT scheme_code) AS unique_schemes,
       COUNT(DISTINCT date_id)     AS unique_dates,
       ROUND(MIN(nav), 4)          AS min_nav,
       ROUND(MAX(nav), 4)          AS max_nav,
       MIN(loaded_at)              AS first_load,
       MAX(loaded_at)              AS last_load
FROM fact_nav_history
UNION ALL
SELECT 'fact_investor_transactions',
       COUNT(*),
       COUNT(DISTINCT scheme_code),
       COUNT(DISTINCT date_id),
       ROUND(MIN(amount_inr), 2),
       ROUND(MAX(amount_inr), 2),
       MIN(loaded_at),
       MAX(loaded_at)
FROM fact_investor_transactions
UNION ALL
SELECT 'fact_scheme_performance',
       COUNT(*),
       COUNT(DISTINCT scheme_code),
       COUNT(DISTINCT date_id),
       ROUND(MIN(returns_1y), 4),
       ROUND(MAX(returns_1y), 4),
       MIN(loaded_at),
       MAX(loaded_at)
FROM fact_scheme_performance;

-- =============================================================================
-- END OF QUERIES
-- =============================================================================
