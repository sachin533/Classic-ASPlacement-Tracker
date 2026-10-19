/* ============================================================================
   Placement Tracker  -  db\seed_500.sql
   ----------------------------------------------------------------------------
   Adds a large, realistic, DEMONSTRATION data set to dbo.Applications so that
   search, filtering, sorting and pagination can be exercised properly.

   WHAT THIS SCRIPT DOES
     * inserts 560 sample job / internship applications,
     * is DETERMINISTIC - the same script always produces exactly the same
       560 rows (every value is derived from the row number with modular
       arithmetic, there is no RAND(), no GETDATE() and no dependency on the
       state of the table),
     * is IDEMPOTENT - each sample row is keyed on its own created_at stamp,
       and the INSERT is guarded by WHERE NOT EXISTS, so running this file
       any number of times never creates a duplicate,
     * is NON-DESTRUCTIVE - it never DELETEs, UPDATEs or TRUNCATEs anything.
       Rows you already have (including rows created through the web app and
       the five reference rows from db\seed.sql) are left exactly as they are,
     * never touches the identity seed or the schema.

   HOW SAMPLE ROWS ARE DISTINGUISHED FROM REAL ROWS
     Every sample row carries created_at = 2024-01-01T00:00:00 + (n-1) seconds,
     i.e. every sample row was "created" during the first 10 minutes of
     2024-01-01.  Rows entered through the web app always get SYSUTCDATETIME()
     (the current UTC time), so the two sets can never overlap.  Count them with:

         SELECT COUNT(*) FROM dbo.Applications
          WHERE created_at >= '2024-01-01' AND created_at < '2024-01-02';

   CONSTRAINTS RESPECTED
     company / role        2-100 characters
     status                one of Wishlist, Applied, Interview, Offer, Rejected
     applied_on            NULL only when status = 'Wishlist'
     job_url               NULL or starting with http:// or https://
     notes                 1000 characters or fewer

   RUN IT
       sqlcmd -S .\SQLEXPRESS -E -C -b -i db\seed_500.sql
   ============================================================================ */

USE [PlacementTracker];
GO

SET NOCOUNT ON;
GO

IF OBJECT_ID(N'dbo.Applications', N'U') IS NULL
BEGIN
    RAISERROR('dbo.Applications does not exist. Run db\schema.sql first.', 16, 1);
END
ELSE
BEGIN
    DECLARE @before INT, @after INT, @samples INT, @inserted INT;
    SELECT @before = COUNT(*) FROM dbo.Applications

    ;WITH
    /* ---- 560 deterministic row numbers -------------------------------- */
    seq AS (
        SELECT 1 AS n
        UNION ALL
        SELECT n + 1 FROM seq WHERE n < 560
    ),

    /* ---- lookup tables, each with its own ordinal so the join is stable
           (the VALUES lists keep the script free of RAND()/GETDATE()) ---- */
    companies AS (
        SELECT CAST(v.ord AS INT)     AS ord,
               CAST(v.company AS NVARCHAR(100)) AS company,
               CAST(v.site   AS NVARCHAR(100))  AS site,
               CAST(COUNT(*) OVER () AS INT)    AS cnt
          FROM (VALUES
            ( 1, N'Google',            N'careers.google.com'),
            ( 2, N'Microsoft',         N'careers.microsoft.com'),
            ( 3, N'Amazon',            N'amazon.jobs'),
            ( 4, N'TCS',               N'careers.tcs.com'),
            ( 5, N'Infosys',           N'careers.infosys.com'),
            ( 6, N'Wipro',             N'careers.wipro.com'),
            ( 7, N'Accenture',         N'careers.accenture.com'),
            ( 8, N'Capgemini',         N'careers.capgemini.com'),
            ( 9, N'Cognizant',         N'careers.cognizant.com'),
            (10, N'IBM',               N'careers.ibm.com'),
            (11, N'Oracle',            N'careers.oracle.com'),
            (12, N'Adobe',             N'careers.adobe.com'),
            (13, N'NVIDIA',            N'careers.nvidia.com'),
            (14, N'Deloitte',          N'careers.deloitte.com'),
            (15, N'Salesforce',        N'careers.salesforce.com'),
            (16, N'Samsung',           N'careers.samsung.com'),
            (17, N'Cisco',             N'careers.cisco.com'),
            (18, N'SAP',               N'careers.sap.com'),
            (19, N'Intel',             N'careers.intel.com'),
            (20, N'Qualcomm',          N'careers.qualcomm.com'),
            (21, N'HCLTech',           N'careers.hcltech.com'),
            (22, N'Tech Mahindra',     N'careers.techmahindra.com'),
            (23, N'Flipkart',          N'careers.flipkart.com'),
            (24, N'Zoho',              N'careers.zoho.com')
          ) AS v(ord, company, site)
    ),

    roles AS (
        SELECT CAST(v.ord AS INT) AS ord,
               CAST(v.role AS NVARCHAR(100)) AS role,
               CAST(COUNT(*) OVER () AS INT) AS cnt
          FROM (VALUES
            ( 1, N'Software Engineer'),
            ( 2, N'Java Developer'),
            ( 3, N'Full Stack Developer'),
            ( 4, N'Frontend Developer'),
            ( 5, N'Backend Developer'),
            ( 6, N'Data Analyst'),
            ( 7, N'QA Engineer'),
            ( 8, N'DevOps Engineer'),
            ( 9, N'Graduate Trainee'),
            (10, N'Software Engineer Intern'),
            (11, N'Cloud Engineer'),
            (12, N'Mobile Developer'),
            (13, N'Business Analyst'),
            (14, N'Cyber Security Analyst'),
            (15, N'Machine Learning Engineer'),
            (16, N'Site Reliability Engineer'),
            (17, N'Platform Engineer'),
            (18, N'Data Science Intern'),
            (19, N'Product Analyst'),
            (20, N'Test Automation Engineer')
          ) AS v(ord, role)
    ),

    statuses AS (
        SELECT CAST(v.ord AS INT) AS ord,
               CAST(v.status AS NVARCHAR(20)) AS status,
               CAST(COUNT(*) OVER () AS INT)  AS cnt
          FROM (VALUES
            (1, N'Wishlist'),
            (2, N'Applied'),
            (3, N'Interview'),
            (4, N'Offer'),
            (5, N'Rejected')
          ) AS v(ord, status)
    ),

    noteTexts AS (
        SELECT CAST(v.ord AS INT) AS ord,
               CAST(v.note AS NVARCHAR(1000)) AS note,
               CAST(COUNT(*) OVER () AS INT)  AS cnt
          FROM (VALUES
            ( 1, N'Applied through the official careers portal.'),
            ( 2, N'Campus drive referral from the placement cell.'),
            ( 3, N'Referred by an employee; resume shared by email.'),
            ( 4, N'Cleared the online assessment; awaiting the next round.'),
            ( 5, N'Attended the on-campus drive; group discussion round cleared.'),
            ( 6, N'Application submitted for the winter internship drive.'),
            ( 7, N'Follow-up mail sent to the recruiter after one week.'),
            ( 8, N'Shortlisted for the technical interview round.'),
            ( 9, N'Position closed by the recruiter after the first round.'),
            (10, N'Saved this opening to apply before the deadline.'),
            (11, N'Off-campus application submitted through a job board.'),
            (12, N'HR round completed; waiting for the offer letter.')
          ) AS v(ord, note)
    )

    INSERT INTO dbo.Applications
        (company, role, status, applied_on, job_url, notes, created_at)
    SELECT
        /* company / role / status: cycled through every lookup value so all
           companies, all roles and all five statuses are represented evenly */
        c.company,
        r.role,
        st.status,

        /* applied_on: NULL exactly when the status is Wishlist, otherwise a
           date spread across two years ending on 2024-01-01 */
        CASE WHEN st.status = N'Wishlist'
             THEN NULL
             ELSE DATEADD(DAY, -((s.n * 37) % 700), CAST(N'2024-01-01' AS DATE))
        END,

        /* job_url: roughly seven out of ten rows carry an HTTPS link */
        CASE WHEN (s.n % 10) < 7
             THEN N'https://' + c.site + N'/jobs/' + CAST(s.n AS NVARCHAR(10))
             ELSE NULL
        END,

        /* notes: one of twelve realistic sentences plus a unique reference,
           so a search for e.g. "PT-0270" matches exactly one row */
        nt.note + N' (reference PT-' +
            RIGHT(N'0000' + CAST(s.n AS NVARCHAR(10)), 4) + N')',

        /* created_at: the demo-data marker - every sample row is stamped
           within the first ten minutes of 2024-01-01, one second apart, so
           the value is unique per row and doubles as the idempotency key */
        DATEADD(SECOND, s.n - 1, CAST(N'2024-01-01T00:00:00' AS DATETIME2(0)))

      FROM seq AS s
      INNER JOIN companies  AS c  ON c.ord  = ((s.n - 1)      % c.cnt) + 1
      INNER JOIN roles      AS r  ON r.ord  = ((s.n * 3  - 1)  % r.cnt) + 1
      INNER JOIN statuses   AS st ON st.ord = ((s.n * 7  - 1)  % st.cnt) + 1
      INNER JOIN noteTexts  AS nt ON nt.ord = ((s.n * 11 - 1)  % nt.cnt) + 1
     WHERE NOT EXISTS (
             SELECT 1
               FROM dbo.Applications AS a
              WHERE a.created_at =
                    DATEADD(SECOND, s.n - 1, CAST(N'2024-01-01T00:00:00' AS DATETIME2(0)))
           )
    OPTION (MAXRECURSION 600);

    SELECT @after    = COUNT(*) FROM dbo.Applications;
    SELECT @samples  = COUNT(*)
      FROM dbo.Applications
     WHERE created_at >= CAST(N'2024-01-01' AS DATE)
       AND created_at <  DATEADD(DAY, 1, CAST(N'2024-01-01' AS DATE));
    SET @inserted = @after - @before;

    PRINT '-----------------------------------------------------------------';
    PRINT 'seed_500.sql';
    PRINT '  rows before this run : ' + CAST(@before    AS VARCHAR(12));
    PRINT '  rows inserted now    : ' + CAST(@inserted  AS VARCHAR(12));
    PRINT '  sample rows present  : ' + CAST(@samples   AS VARCHAR(12));
    PRINT '  rows after this run  : ' + CAST(@after     AS VARCHAR(12));
    IF @samples >= 500
        PRINT '  OK - at least 500 sample records are present.';
    ELSE
        PRINT '  WARNING - fewer than 500 sample records present.';
    PRINT '-----------------------------------------------------------------';

    /* One readable result set so sqlcmd/SSMS show the outcome. */
    SELECT @before   AS rows_before,
           @inserted AS rows_inserted,
           @samples  AS sample_rows,
           @after    AS rows_total;
END
GO

/* ============================================================================
   Read back what a re-run would do: running the whole script a second time
   must report rows_inserted = 0.  This batch is informational only.
   ============================================================================ */
SELECT COUNT(*)                       AS total_rows,
       COUNT(CASE WHEN created_at >= CAST(N'2024-01-01' AS DATE)
                   AND created_at <  DATEADD(DAY, 1, CAST(N'2024-01-01' AS DATE))
                  THEN 1 END)         AS sample_rows,
       COUNT(DISTINCT company)        AS distinct_companies,
       COUNT(DISTINCT role)           AS distinct_roles,
       COUNT(DISTINCT status)         AS distinct_statuses
  FROM dbo.Applications;
GO
