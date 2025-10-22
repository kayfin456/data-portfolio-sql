-- Purpose: Analyze team activity data by group and calculate total productive hours per day.
-- Demonstrates advanced SQL: CTEs, array handling, joins, and pivot-style aggregation.

WITH activity_groups AS (
    SELECT 'Support' AS group_name, 
           ARRAY['Team_A1', 'Team_A2'] AS team_names, 
           ARRAY['Ringing', 'In Service Phone', 'After Call Work', 'Outbound Call', 'Ready'] AS activities
    UNION ALL
    SELECT 'Inbound', 
           ARRAY['Team_B1','Team_B2','Team_B3'], 
           ARRAY['Ringing','In Service Phone','After Call Work','Outbound Call','Ready']
    UNION ALL
    SELECT 'Outbound', 
           ARRAY['Team_C1'], 
           ARRAY['Ringing','In Service Phone','After Call Work','Outbound Call','Ready']
    UNION ALL
    SELECT 'Messaging', 
           ARRAY['Team_D1','Team_D2','Team_D3','Team_D4'], 
           ARRAY['Ringing','In Service Phone','After Call Work','Outbound Call','Ready']
    UNION ALL
    SELECT 'Processing', 
           ARRAY['Team_E1'], 
           ARRAY['Ringing','In Service Phone','After Call Work','Outbound Call','Ready']
    UNION ALL
    SELECT 'Verification', 
           ARRAY['Team_F1','Team_F2','Team_F3'], 
           ARRAY['Ringing','In Service Phone','After Call Work','Outbound Call','Ready','Back Office']
    UNION ALL
    SELECT 'Automation', 
           ARRAY['Team_G1'], 
           ARRAY['Ringing','In Service Phone','After Call Work','Outbound Call','Ready']
    UNION ALL
    SELECT 'Credit', 
           ARRAY['Team_H1','Team_H2'], 
           ARRAY['Pipeline Management', 'Loan Review']
    UNION ALL
    SELECT 'Chat Inbound', 
           ARRAY['Team_I1'], 
           ARRAY['Written Task Completion','In Service Chat']
    UNION ALL
    SELECT 'Chat Messaging', 
           ARRAY['Team_J1'], 
           ARRAY['Written Task Completion','In Service Chat']
),

ranked_activity AS (
    -- CTE 1: Capture latest activity record per employee/activity/start_time
    SELECT
        act.employee_id,
        emp.username,
        act.activity,
        act.start_time,
        act.duration,
        act.period,
        CAST(
            (DATE_PARSE(act.start_time, '%m-%d-%Y %H:%i:%s.%f') AT TIME ZONE 'UTC' AT TIME ZONE 'America/Denver') AS DATE
        ) AS activity_date,
        ROW_NUMBER() OVER (
            PARTITION BY act.employee_id, act.activity, act.start_time
            ORDER BY act.period DESC
        ) AS row_num
    FROM company_data.activity_log act
    JOIN company_data.employee_info emp
        ON act.employee_id = emp.employee_id
    WHERE CAST(DATE_PARSE(act.start_time, '%m-%d-%Y %H:%i:%s.%f') AS DATE) >= CURRENT_DATE - INTERVAL '30' DAY
),

joined_with_roles AS (
    -- CTE 2: Attach employee role/team metadata
    SELECT
        ra.*,
        ops.team_name
    FROM ranked_activity ra
    JOIN company_data.operations_info ops
        ON ra.username = ops.email
        AND ra.activity_date = ops.record_date
    WHERE ra.row_num = 1
),

exploded_groups AS (
    -- CTE 3: Expand team/activity combinations by group
    SELECT
        group_name,
        team_name,
        activity
    FROM activity_groups
    CROSS JOIN UNNEST(team_names) AS t(team_name)
    CROSS JOIN UNNEST(activities) AS a(activity)
),

tagged_activities AS (
    -- CTE 4: Tag activities with their corresponding group
    SELECT
        jwr.*,
        eg.group_name
    FROM joined_with_roles jwr
    JOIN exploded_groups eg
        ON jwr.team_name = eg.team_name
       AND jwr.activity = eg.activity
)

-- FINAL OUTPUT: Pivot daily totals by group
SELECT
    activity_date,
    ROUND(SUM(CASE WHEN group_name = 'Credit' THEN duration ELSE 0 END) / 3600.0, 2) AS credit_hours,
    ROUND(SUM(CASE WHEN group_name = 'Support' THEN duration ELSE 0 END) / 3600.0, 2) AS support_hours,
    ROUND(SUM(CASE WHEN group_name = 'Inbound' THEN duration ELSE 0 END) / 3600.0, 2) AS inbound_hours,
    ROUND(SUM(CASE WHEN group_name = 'Outbound' THEN duration ELSE 0 END) / 3600.0, 2) AS outbound_hours,
    ROUND(SUM(CASE WHEN group_name = 'Messaging' THEN duration ELSE 0 END) / 3600.0, 2) AS messaging_hours,
    ROUND(SUM(CASE WHEN group_name = 'Processing' THEN duration ELSE 0 END) / 3600.0, 2) AS processing_hours,
    ROUND(SUM(CASE WHEN group_name = 'Verification' THEN duration ELSE 0 END) / 3600.0, 2) AS verification_hours,
    ROUND(SUM(CASE WHEN group_name = 'Automation' THEN duration ELSE 0 END) / 3600.0, 2) AS automation_hours,
    ROUND(SUM(CASE WHEN group_name = 'Chat Inbound' THEN duration ELSE 0 END) / 3600.0, 2) AS chat_inbound_hours,
    ROUND(SUM(CASE WHEN group_name = 'Chat Messaging' THEN duration ELSE 0 END) / 3600.0, 2) AS chat_messaging_hours
FROM tagged_activities
GROUP BY activity_date
ORDER BY activity_date;
