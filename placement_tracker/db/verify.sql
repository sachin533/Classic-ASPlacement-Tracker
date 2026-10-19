/* ==========================================================================
   Placement Application Tracker - verify.sql
   Run after a CRUD session to prove what actually reached the table.
   ========================================================================== */

-- 1. Total rows
SELECT COUNT(*) AS total FROM dbo.Applications;

-- 2. Rows per status
SELECT status, COUNT(*) AS rows_count
FROM dbo.Applications
GROUP BY status
ORDER BY status;

-- 3. Violation check: must return ZERO rows
--    (every status other than Wishlist requires a date)
SELECT *
FROM dbo.Applications
WHERE status <> N'Wishlist' AND applied_on IS NULL;

-- 4. Five most recently created rows
SELECT TOP 5 id, company, role, status, applied_on, created_at
FROM dbo.Applications
ORDER BY created_at DESC;
GO
