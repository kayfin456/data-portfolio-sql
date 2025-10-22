-- Purpose: Analyze employee adherence and activity metrics by day,
-- joining adherence records with occupancy data to evaluate productivity.

WITH LatestAdherence AS (
    -- CTE 1: Gets the most recent adherence record per employee per day
    SELECT 
        adh.employee_id,
        adh.first_name, 
        adh.last_name,
        CAST(DATE_PARSE(adh.day, '%m-%d-%Y') AS DATE) AS record_date,
        adh.in_adherence,
        adh.non_adherence,
        adh.period AS capture_period,
        emp.username,
        ROW_NUMBER() OVER (
            PARTITION BY adh.employee_id, CAST(DATE_PARSE(adh.day, '%m-%d-%Y') AS DATE)
            ORDER BY adh.period DESC
        ) AS rn
    FROM 
        company_data.adherence adh
    JOIN 
        company_data.employee_info emp
        ON adh.employee_id = emp.employee_id
),

FilteredAdherence AS (
    -- CTE 2: Keeps only the most recent record per employee per day
    SELECT *
    FROM LatestAdherence
    WHERE rn = 1
      AND record_date >= DATE_ADD('month', -13, DATE_TRUNC('month', CURRENT_DATE))
),

ParsedActivity AS (
    -- CTE 3: Parses timestamps and filters to specific activity categories
    SELECT 
        employee_id,
        DATE_PARSE(start_time, '%m-%d-%Y %H:%i:%s.%f') AS parsed_start_time,
        activity,
        duration
    FROM 
        company_data.activity_log
    WHERE 
        activity IN ('In Service', 'Ready', 'After Call Work')
        AND CAST(DATE_PARSE(start_time, '%m-%d-%Y %H:%i:%s.%f') AS DATE) >= DATE_ADD('month', -13, DATE_TRUNC('month', CURRENT_DATE))
),

Occupancy AS (
    -- CTE 4: Aggregates daily activity duration per employee
    SELECT 
        employee_id,
        DATE(parsed_start_time) AS record_date,
        SUM(CASE WHEN activity = 'In Service' THEN duration ELSE 0 END) AS in_service,
        SUM(CASE WHEN activity = 'Ready' THEN duration ELSE 0 END) AS ready,
        SUM(CASE WHEN activity = 'After Call Work' THEN duration ELSE 0 END) AS after_call_work
    FROM 
        ParsedActivity
    GROUP BY 
        employee_id, DATE(parsed_start_time)
)

-- Final output combining adherence and occupancy data
SELECT 
    adh.first_name,
    adh.last_name,
    adh.record_date,
    adh.in_adherence,
    adh.non_adherence,
    ops.supervisor_name,
    ops.team_name,   
    ops.department,
    ops.company_name,
    ops.employment_status,
    ops.role_function,
    occ.in_service,
    occ.ready,
    occ.after_call_work
FROM 
    FilteredAdherence adh
JOIN 
    company_data.operations_info ops
    ON adh.username = ops.email
    AND adh.record_date = ops.record_date
LEFT JOIN Occupancy occ
    ON adh.employee_id = occ.employee_id
    AND adh.record_date = occ.record_date
ORDER BY 
    adh.record_date DESC;
