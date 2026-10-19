/* ============================================================================
   Placement Tracker  -  db\seed.sql
   ----------------------------------------------------------------------------
   Inserts the five demonstration rows used by the reference design.

   The script is NON-DESTRUCTIVE: it only inserts when dbo.Applications is
   completely empty, so running it twice never duplicates data and never
   overwrites rows you added by hand.

   Run it from SSMS or:

       sqlcmd -S .\SQLEXPRESS -E -C -b -i db\seed.sql
   ============================================================================ */

USE [PlacementTracker];
GO

SET NOCOUNT ON;
GO

IF OBJECT_ID(N'dbo.Applications', N'U') IS NULL
BEGIN
    RAISERROR('dbo.Applications does not exist. Run db\schema.sql first.', 16, 1);
END
ELSE IF EXISTS (SELECT 1 FROM dbo.Applications)
BEGIN
    PRINT 'dbo.Applications already contains rows - seed skipped.';
END
ELSE
BEGIN
    BEGIN TRANSACTION;

    INSERT INTO dbo.Applications (company, role, status, applied_on, job_url, notes, created_at)
    VALUES
        (N'Google',  N'Software Engineer Intern', N'Applied',   CAST(N'2024-08-10' AS DATE),
         N'https://careers.google.com',
         N'Applied through the university placement portal.',
         CAST(N'2024-08-10T09:15:00' AS DATETIME2(0))),

        (N'Microsoft', N'Software Engineer',      N'Interview', CAST(N'2024-07-21' AS DATE),
         N'https://careers.microsoft.com',
         N'Completed online assessment. Waiting for interview email.',
         CAST(N'2024-07-21T14:30:00' AS DATETIME2(0))),

        (N'Amazon',  N'SDE Intern',               N'Offer',     CAST(N'2024-06-15' AS DATE),
         N'https://www.amazon.jobs',
         N'Offer received. Response due by the end of the month.',
         CAST(N'2024-06-15T11:05:00' AS DATETIME2(0))),

        (N'TCS',     N'Graduate Trainee',         N'Rejected',  CAST(N'2024-06-01' AS DATE),
         NULL,
         N'Rejected after the online round.',
         CAST(N'2024-06-01T16:45:00' AS DATETIME2(0))),

        (N'Infosys', N'System Engineer',          N'Wishlist',  NULL,
         N'https://careers.infosys.com',
         N'Not applied yet - opening expected next month.',
         CAST(N'2024-05-28T08:00:00' AS DATETIME2(0)));

    COMMIT TRANSACTION;

    PRINT 'Seeded 5 demonstration rows into dbo.Applications.';
END
GO
