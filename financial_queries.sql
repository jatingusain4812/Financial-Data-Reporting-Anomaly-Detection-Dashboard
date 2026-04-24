-- ============================================================
-- FINANCIAL DATA REPORTING & ANOMALY DETECTION
-- MySQL Schema + Reconciliation Queries
-- Tools: MySQL, Python, Excel, Power BI
-- ============================================================

CREATE DATABASE IF NOT EXISTS financial_audit_db;
USE financial_audit_db;

DROP TABLE IF EXISTS financial_data;

CREATE TABLE financial_data (
    record_id        VARCHAR(10)     PRIMARY KEY,
    month            VARCHAR(7),
    region           VARCHAR(20),
    product          VARCHAR(20),
    channel          VARCHAR(20),
    units_sold       INT,
    unit_price       DECIMAL(10,2),
    revenue          DECIMAL(14,2),
    cost             DECIMAL(14,2),
    profit           DECIMAL(14,2)
);

-- ── RECONCILIATION QUERIES ────────────────────────────────────

-- 1. Revenue by Region
SELECT region,
  COUNT(*) AS records,
  SUM(revenue) AS total_revenue,
  SUM(profit) AS total_profit,
  ROUND(SUM(profit)/NULLIF(SUM(revenue),0)*100, 2) AS margin_pct
FROM financial_data
WHERE region IS NOT NULL AND region != ''
GROUP BY region ORDER BY total_revenue DESC;

-- 2. Negative Revenue Detection
SELECT record_id, month, region, product, revenue
FROM financial_data WHERE revenue < 0 ORDER BY revenue ASC;

-- 3. Cost Exceeds Revenue
SELECT record_id, month, region, product, revenue, cost,
  ROUND(cost - revenue, 2) AS excess_cost
FROM financial_data WHERE cost > revenue ORDER BY excess_cost DESC;

-- 4. Revenue Spike (Z-Score > 3)
SELECT f.record_id, f.region, f.product, f.revenue,
  ROUND((f.revenue - s.avg_rev) / NULLIF(s.std_rev, 0), 2) AS z_score
FROM financial_data f
JOIN (
  SELECT product, AVG(revenue) AS avg_rev, STDDEV(revenue) AS std_rev
  FROM financial_data GROUP BY product
) s ON f.product = s.product
WHERE f.revenue > s.avg_rev + 3 * s.std_rev
ORDER BY z_score DESC;

-- 5. Units = 0 but Revenue > 0
SELECT record_id, month, region, product, units_sold, revenue
FROM financial_data WHERE units_sold = 0 AND revenue > 0;

-- 6. Monthly Trend
SELECT month,
  SUM(revenue) AS total_revenue,
  SUM(profit) AS total_profit,
  ROUND(SUM(profit)/NULLIF(SUM(revenue),0)*100, 2) AS margin_pct
FROM financial_data
GROUP BY month ORDER BY month;

-- 7. Product Anomaly Summary
SELECT product,
  COUNT(*) AS total_records,
  SUM(CASE WHEN revenue < 0 THEN 1 ELSE 0 END) AS neg_revenue,
  SUM(CASE WHEN cost > revenue THEN 1 ELSE 0 END) AS cost_exceeds,
  SUM(CASE WHEN units_sold = 0 AND revenue > 0 THEN 1 ELSE 0 END) AS units_zero
FROM financial_data
GROUP BY product
ORDER BY (neg_revenue + cost_exceeds + units_zero) DESC;
