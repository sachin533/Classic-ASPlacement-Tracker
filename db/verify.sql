/* ============================================================================
   Placement Tracker  -  db\verify.sql
   ----------------------------------------------------------------------------
   Read-only checks that prove what the application actually saved.
   Nothing here inserts, updates or deletes.

       sqlcmd -S .\SQLEXPRESS -E -C -i db\verify.sql
   ============================================================================ */

USE [PlacementTracker];
GO

SET NOCOUNT ON;
PRINT '=== 1. Table shape ===';
SELECT c.COLUMN_NAME,
       c.DATA_TYPE,
       c.IS_NULLABLE,
       c.CHARACTER_MAXIMUM_LENGTH,
       COLUMNPROPERTY(OBJECT_ID('dbo.Applications'), c.COLUMN_NAME, 'IsIdentity') AS is_identity
  FROM INFORMATION_SCHEMA.COLUMNS AS c
 WHERE c.TABLE_SCHEMA = N'dbo'
   AND c.TABLE_NAME   = N'Applications'
 ORDER BY c.ORDINAL_POSITION;
GO

PRINT '=== 2. Constraints ===';
SELECT name       AS constraint_name,
       type_desc,
       definition
  FROM sys.check_constraints
 WHERE parent_object_id = OBJECT_ID(N'dbo.Applications')
 ORDER BY name;
GO

PRINT '=== 3. Total row count (expected > 0 after seeding) ===';
SELECT COUNT(*) AS total FROM dbo.Applications;
GO

PRINT '=== 3a. DEMO DATA: sample rows added by db\seed_500.sql ===';
/* Every row seeded by seed_500.sql is stamped inside the first ten minutes
   of 2024-01-01; rows created through the web app always carry the current
   UTC time, so the two sets can never overlap. */
SELECT COUNT(*) AS total_rows,
       COUNT(CASE WHEN created_at >= CAST(N'2024-01-01' AS DATE)
                   AND created_at <  DATEADD(DAY, 1, CAST(N'2024-01-01' AS DATE))
                  THEN 1 END) AS sample_rows,
       COUNT(CASE WHEN created_at <  CAST(N'2024-01-01' AS DATE)
                   OR   created_at >= DATEADD(DAY, 1, CAST(N'2024-01-01' AS DATE))
                  THEN 1 END) AS rows_outside_the_sample_set
  FROM dbo.Applications;
GO

PRINT '=== 3b. DEMO DATA: sample row count must be >= 500 ===';
SELECT COUNT(*) AS sample_rows,
       CASE WHEN COUNT(*) >= 500 THEN 'OK' ELSE 'BELOW 500 - run db\seed_500.sql' END AS verdict
  FROM dbo.Applications
 WHERE created_at >= CAST(N'2024-01-01' AS DATE)
   AND created_at <  DATEADD(DAY, 1, CAST(N'2024-01-01' AS DATE));
GO

PRINT '=== 3c. DEMO DATA: seed_500.sql is idempotent (all timestamps unique) ===';
/* The seed keys each sample row on its own created_at.  Duplicate timestamps
   would mean a second run had written the same rows again - expect 0. */
SELECT COUNT(*) AS sample_rows,
       COUNT(DISTINCT created_at) AS distinct_sample_timestamps,
       CASE WHEN COUNT(*) = COUNT(DISTINCT created_at)
            THEN 'OK - a re-run would insert 0 rows'
            ELSE 'DUPLICATED' END AS verdict
  FROM dbo.Applications
 WHERE created_at >= CAST(N'2024-01-01' AS DATE)
   AND created_at <  DATEADD(DAY, 1, CAST(N'2024-01-01' AS DATE));
GO

PRINT '=== 3d. PAGINATION: OFFSET/FETCH returns a stable 20-row page ===';
/* Mirrors the query in inc\db.asp - proves the paging syntax the dashboard
   relies on works on this server, with a tie-breaker on the unique id. */
DECLARE @offset INT, @size INT;
SET @offset = 20;   -- page 2 at 20 rows per page
SET @size   = 20;
SELECT id, company, status
  FROM dbo.Applications
 ORDER BY id DESC
 OFFSET @offset ROWS FETCH NEXT @size ROWS ONLY;
GO

PRINT '=== 4. Rows grouped by status ===';
SELECT status, COUNT(*) AS rows_in_status
  FROM dbo.Applications
 GROUP BY status
 ORDER BY status;
GO

PRINT '=== 5. INTEGRITY: non-Wishlist rows missing a date (must return 0 rows) ===';
SELECT id, company, role, status, applied_on
  FROM dbo.Applications
 WHERE status <> N'Wishlist'
   AND applied_on IS NULL;
GO

PRINT '=== 6. INTEGRITY: job_url that is not http/https (must return 0 rows) ===';
SELECT id, company, job_url
  FROM dbo.Applications
 WHERE job_url IS NOT NULL
   AND job_url NOT LIKE N'http://%'
   AND job_url NOT LIKE N'https://%';
GO

PRINT '=== 7. INTEGRITY: company/role outside 2-100 chars (must return 0 rows) ===';
SELECT id, company, role
  FROM dbo.Applications
 WHERE LEN(company) < 2 OR LEN(company) > 100
    OR LEN(role)    < 2 OR LEN(role)    > 100;
GO

PRINT '=== 8. INTEGRITY: bad status values (must return 0 rows) ===';
SELECT id, status
  FROM dbo.Applications
 WHERE status NOT IN (N'Wishlist', N'Applied', N'Interview', N'Offer', N'Rejected');
GO

PRINT '=== 9. Newest 5 rows ===';
SELECT TOP 5
       id, company, role, status, applied_on, job_url, notes, created_at
  FROM dbo.Applications
 ORDER BY id DESC;
GO

PRINT '=== verify.sql finished ===';
GO
