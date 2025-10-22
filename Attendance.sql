-- Purpose: Generate a daily schedule summary per employee, including earliest start and latest end times.
-- Demonstrates CTE chaining, time parsing, filtering, and metadata enrichment.

WITH deduped_planned_activity AS (
  -- CTE 1: Remove duplicate planned activity records
  SELECT DISTINCT employee_id, activity, start_time, end_time, period
  FROM company_data.planned_activity
),

parsed AS (
  -- CTE 2: Parse datetime strings into separate date and time components
  SELECT 
    pa.employee_id,
    pa.activity,
    TRY(CAST(DATE_PARSE(pa.start_time, '%m-%d-%Y %H:%i:%s.%f') AS DATE)) AS start_date,
    FORMAT_DATETIME(DATE_PARSE(pa.start_time, '%m-%d-%Y %H:%i:%s.%f'), 'HH:mm:ss') AS start_time_only,
    FORMAT_DATETIME(DATE_PARSE(pa.end_time, '%m-%d-%Y %H:%i:%s.%f'), 'HH:mm:ss') AS end_time_only,
    TRY(CAST(pa.period AS DATE)) AS period
  FROM deduped_planned_activity pa
  WHERE TRY(CAST(DATE_PARSE(pa.start_time, '%m-%d-%Y %H:%i:%s.%f') AS DATE)) 
        BETWEEN DATE '2025-06-01' AND DATE '2025-06-30'
),

latest_periods AS (
  -- CTE 3: Keep only latest schedule version per employee per day
  SELECT employee_id, start_date, MAX(period) AS latest_period
  FROM parsed
  GROUP BY employee_id, start_date
),

filtered AS (
  -- CTE 4: Filter to latest records
  SELECT p.*
  FROM parsed p
  JOIN latest_periods l
    ON p.employee_id = l.employee_id
   AND p.start_date = l.start_date
   AND p.period = l.latest_period
),

no_unapproved_remote AS (
  -- CTE 5: Remove unapproved remote work records
  SELECT *
  FROM filtered
  WHERE activity != 'Remote Work - Unapproved'
),

start_end AS (
  -- CTE 6: Aggregate to get earliest start and latest end per employee per day
  SELECT 
    employee_id,
    start_date AS work_date,
    FORMAT_DATETIME(
      CAST(CONCAT('1970-01-01 ', MIN(start_time_only)) AS TIMESTAMP) - INTERVAL '7' HOUR,
      'HH:mm:ss'
    ) AS scheduled_start,
    FORMAT_DATETIME(
      CAST(CONCAT('1970-01-01 ', MAX(end_time_only)) AS TIMESTAMP) - INTERVAL '7' HOUR,
      'HH:mm:ss'
    ) AS scheduled_end
  FROM no_unapproved_remote
  GROUP BY employee_id, start_date
),

username_per_employee AS (
  -- CTE 7: Map usernames to employees
  SELECT employee_id, MAX(username) AS username
  FROM company_data.employee_info
  GROUP BY employee_id
),

joined_agent_info AS (
  -- CTE 8: Enrich with role and organizational data
  SELECT 
    se.work_date,
    se.scheduled_start,
    se.scheduled_end,
    u.username,
    ag.supervisor_name,
    ag.team_name,
    ag.department,
    ag.organization,
    ag.employment_status,
    ag.role_function
  FROM start_end se
  LEFT JOIN username_per_employee u
    ON se.employee_id = u.employee_id
  LEFT JOIN company_data.agent_info ag
    ON u.username = ag.email
   AND se.work_date = ag.record_date
  WHERE ag.organization = 'ExampleCorp'
)

-- Final output: daily employee schedules with team context
SELECT *
FROM joined_agent_info
ORDER BY work_date, team_name;
