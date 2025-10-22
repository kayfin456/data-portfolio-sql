-- Purpose: Calculate weekly call center performance metrics (Service Level, Volume, AHT, Abandon Rate)
-- Demonstrates advanced aggregation, CASE mapping, and conditional percentage calculations.

SELECT
  DATE(DATE_TRUNC('week', call_timestamp)) AS week_start,
  
  CASE
    WHEN skill IN ('Member Support') THEN 'Inbound Support'
    WHEN skill IN ('Back Office Operations') THEN 'Back Office'
    WHEN skill IN ('Collections_Low_Risk', 'Collections_Regional', 'Collections_Remote') THEN 'Collections - Low Risk'
    WHEN skill IN ('Collections_High_Risk') THEN 'Collections - High Risk'
    WHEN skill IN ('Customer Acquisition', 'Customer Retention', 'Credit Bureau', 'Escalations', 'Servicing', 'Small Business', 'Outbound Sales') THEN 'Member Engagement'
    WHEN skill IN ('Patient Solutions', 'Education Finance') THEN 'Partner Finance'
    WHEN skill IN ('Banking Ops', 'Wire Transfers', 'Customer Service', 'New Accounts', 'Online Banking') THEN 'Digital Banking'
    WHEN skill IN ('Auto Origination', 'Auto Verification', 'Auto Outbound') THEN 'Auto Lending'
    WHEN skill IN ('Fraud Investigation', 'Fraud Prevention') THEN 'Fraud'
    ELSE 'Other'
  END AS department,
  
  -- Service Level: % of calls answered within 30 seconds
  CAST(
    SUM(CASE WHEN speed_of_answer <= 30 THEN 1 ELSE 0 END) * 100.0 /
    NULLIF(SUM(CASE WHEN handle_time IS NOT NULL THEN 1 ELSE 0 END), 0)
    AS DOUBLE
  ) AS service_level,
  
  -- Total call volume
  COUNT(*) AS total_calls,
  
  -- Average handle time (AHT)
  AVG(handle_time) AS avg_handle_time,
  
  -- Abandon rate: % of calls abandoned after 30 seconds
  CAST(
    SUM(CASE WHEN time_to_abandon IS NOT NULL AND time_to_abandon >= 30 THEN 1 ELSE 0 END) * 100.0 /
    COUNT(*) AS DOUBLE
  ) AS abandon_rate
  
FROM company_data.call_center_logs
WHERE call_timestamp >= DATE_TRUNC('week', DATE_ADD('week', -13, CURRENT_DATE))
GROUP BY 1, 2
ORDER BY week_start DESC;
