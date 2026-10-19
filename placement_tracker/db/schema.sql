/* ==========================================================================
   Placement Application Tracker - schema.sql
   Creates the dbo.Applications table in the PlacementTracker database.
   Run against .\SQLEXPRESS with a login that can create tables.
   ========================================================================== */

IF OBJECT_ID(N'dbo.Applications', N'U') IS NOT NULL
  DROP TABLE dbo.Applications;
GO

CREATE TABLE dbo.Applications (
  id          INT IDENTITY(1,1) CONSTRAINT PK_Applications PRIMARY KEY,
  company     NVARCHAR(100)  NOT NULL CONSTRAINT CK_App_Company
                CHECK (LEN(company) BETWEEN 2 AND 100),
  role        NVARCHAR(100)  NOT NULL CONSTRAINT CK_App_Role
                CHECK (LEN(role) BETWEEN 2 AND 100),
  status      NVARCHAR(20)   NOT NULL CONSTRAINT CK_App_Status
                CHECK (status IN (N'Wishlist', N'Applied', N'Interview', N'Offer', N'Rejected')),
  applied_on  DATE           NULL,
  job_url     NVARCHAR(500)  NULL CONSTRAINT CK_App_JobUrl
                CHECK (job_url IS NULL OR job_url LIKE N'http://%' OR job_url LIKE N'https://%'),
  notes       NVARCHAR(1000) NULL CONSTRAINT CK_App_Notes
                CHECK (notes IS NULL OR LEN(notes) <= 1000),
  created_at  DATETIME2(0)   NOT NULL CONSTRAINT DF_App_Created
                DEFAULT SYSUTCDATETIME(),
  /* An application only needs a date once it has actually been sent,
     so Wishlist rows may leave applied_on empty. */
  CONSTRAINT CK_App_AppliedOn CHECK (status = N'Wishlist' OR applied_on IS NOT NULL)
);
GO

/* Helpful for the dashboard, which lists by most recent activity. */
IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE name = N'IX_Applications_CreatedAt'
                 AND object_id = OBJECT_ID(N'dbo.Applications'))
  CREATE NONCLUSTERED INDEX IX_Applications_CreatedAt
    ON dbo.Applications (created_at DESC);
GO
