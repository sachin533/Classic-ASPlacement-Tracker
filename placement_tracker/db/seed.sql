/* ==========================================================================
   Placement Application Tracker - seed.sql
   Sample rows so the dashboard has something to show on first run.
   Safe to re-run: clears existing rows first (development only).
   ========================================================================== */

IF OBJECT_ID(N'dbo.Applications', N'U') IS NULL
  RAISERROR (N'dbo.Applications does not exist. Run schema.sql first.', 16, 1);

DELETE FROM dbo.Applications;
DBCC CHECKIDENT (N'dbo.Applications', RESEED, 0);
GO

INSERT INTO dbo.Applications (company, role, status, applied_on, job_url, notes)
VALUES
  (N'Google',    N'Software Engineer Intern', N'Applied',   '2024-08-10', N'https://careers.google.com',
   N'Referred by campus placement cell. Application submitted online.'),
  (N'Microsoft', N'Software Engineer',        N'Interview', '2024-07-21', N'https://careers.microsoft.com',
   N'Completed online assessment. Waiting for interview email.'),
  (N'Amazon',    N'SDE Intern',               N'Offer',     '2024-06-15', N'https://www.amazon.jobs',
   N'Offer received. Respond by end of month.'),
  (N'TCS',       N'Graduate Trainee',         N'Rejected',  '2024-06-01', NULL,
   N'Rejected after technical round.'),
  (N'Infosys',   N'System Engineer',          N'Wishlist',  NULL,         N'https://careers.infosys.com',
   N'Not applied yet - openings open in September.');
GO

SELECT COUNT(*) AS seeded_rows FROM dbo.Applications;
GO
