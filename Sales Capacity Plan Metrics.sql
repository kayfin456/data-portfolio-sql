-- Purpose: Analyze call outcomes and performance by campaign category.
-- Demonstrates advanced CASE logic, pattern matching, and window functions
-- for calculating relative performance metrics.

SELECT
  campaign_group,
  termination_bucket,
  call_count,
  ROUND(
    100.0 * call_count / SUM(call_count) OVER (PARTITION BY campaign_group),
    2
  ) AS pct_of_campaign_calls,
  avg_total_seconds,
  avg_wrap_seconds,
  avg_effective_duration_seconds
FROM (
  SELECT
    -- Categorize campaigns into broader groups based on name patterns
    CASE
      WHEN LOWER(campaign_name) LIKE '%lead_tracking%' THEN 'Lead_Tracking'
      WHEN LOWER(campaign_name) LIKE '%multi_channel%' THEN 'All_Channels'
      WHEN LOWER(campaign_name) LIKE 'funnel%' THEN 'Funnel_Accounts'
      WHEN campaign_name IN ('Afternoon_Batch', 'Morning_Batch', 'Noon_Batch') THEN 'Batch_Campaigns'
      WHEN LOWER(campaign_name) LIKE '%manual%' THEN 'Manual_Outreach'
      ELSE 'Other'
    END AS campaign_group,

    -- Simplify termination codes into broader categories
    CASE
      WHEN termination_code IN (
        'AGENT - CUSTOMER 1',
        'AGENT - CUSTOMER 3',
        'AGENT - CUSTOMER 4',
        'AGENT - CUSTOMER RPC 2',
        'AGENT - CUSTOMER RPC 4'
      ) THEN 'Other'
      WHEN termination_code IN (
        'AGENT - CUSTOMER 2',
        'AGENT - Left Message',
        'AGENT - No Message Left'
      ) THEN 'Voicemail'
      WHEN termination_code = 'AGENT - CUSTOMER RPC 3' THEN 'Connected Authenticated'
      WHEN termination_code IN (
        'Busy',
        'Hung Up Early',
        'Invalid Number',
        'No Answer'
      ) THEN 'Unreachable'
      ELSE 'Other'
    END AS termination_bucket,

    COUNT(*) AS call_count,
    ROUND(AVG(CAST(total_duration_ms AS DOUBLE)), 2) AS avg_total_seconds,
    ROUND(AVG(CAST(wrap_duration_ms AS DOUBLE)), 2) AS avg_wrap_seconds,
    ROUND(
      AVG(CAST(total_duration_ms AS DOUBLE) + CAST(wrap_duration_ms AS DOUBLE)),
      2
    ) AS avg_effective_duration_seconds

  FROM company_data.call_results
  WHERE CAST(call_date AS DATE) >= DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '3' MONTH
    AND total_duration_ms IS NOT NULL
    AND wrap_duration_ms IS NOT NULL
    AND termination_code NOT IN (
      'Operator Transfer',
      'Operator Transfer (Agent Abandoned)'
    )
    AND (
      LOWER(campaign_name) LIKE '%lead_tracking%'
      OR LOWER(campaign_name) LIKE '%multi_channel%'
      OR LOWER(campaign_name) LIKE 'funnel%'
      OR campaign_name IN ('Afternoon_Batch', 'Morning_Batch', 'Noon_Batch')
      OR LOWER(campaign_name) LIKE '%manual%'
    )

  GROUP BY
    CASE
      WHEN LOWER(campaign_name) LIKE '%lead_tracking%' THEN 'Lead_Tracking'
      WHEN LOWER(campaign_name) LIKE '%multi_channel%' THEN 'All_Channels'
      WHEN LOWER(campaign_name) LIKE 'funnel%' THEN 'Funnel_Accounts'
      WHEN campaign_name IN ('Afternoon_Batch', 'Morning_Batch', 'Noon_Batch') THEN 'Batch_Campaigns'
      WHEN LOWER(campaign_name) LIKE '%manual%' THEN 'Manual_Outreach'
      ELSE 'Other'
    END,
    CASE
      WHEN termination_code IN (
        'AGENT - CUSTOMER 1',
        'AGENT - CUSTOMER 3',
        'AGENT - CUSTOMER 4',
        'AGENT - CUSTOMER RPC 2',
        'AGENT - CUSTOMER RPC 4'
      ) THEN 'Other'
      WHEN termination_code IN (
        'AGENT - CUSTOMER 2',
        'AGENT - Left Message',
        'AGENT - No Message Left'
      ) THEN 'Voicemail'
      WHEN termination_code = 'AGENT - CUSTOMER RPC 3' THEN 'Connected Authenticated'
      WHEN termination_code IN (
        'Busy',
        'Hung Up Early',
        'Invalid Number',
        'No Answer'
      ) THEN 'Unreachable'
      ELSE 'Other'
    END
) AS grouped
ORDER BY campaign_group, avg_effective_duration_seconds DESC;
