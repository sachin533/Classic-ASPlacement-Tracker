/* ============================================================================
   Placement Tracker  -  db\schema.sql
   ----------------------------------------------------------------------------
   Creates the PlacementTracker database and the dbo.Applications table.

   The script is intentionally NON-DESTRUCTIVE and idempotent:
     * the database is only created when it does not already exist,
     * the table is only created when it does not already exist,
     * no rows are ever deleted or updated.

   Run it from SSMS (Open File -> Execute) or from a command prompt:

       sqlcmd -S .\SQLEXPRESS -E -C -b -i db\schema.sql

   Tested with SQL Server 2022 Express (instance .\SQLEXPRESS).
   ============================================================================ */

SET NOCOUNT ON;
GO

/* ----------------------------------------------------------------------------
   1. Database
   ---------------------------------------------------------------------------- */
IF DB_ID(N'PlacementTracker') IS NULL
BEGIN
    CREATE DATABASE [PlacementTracker];
    PRINT 'Created database [PlacementTracker].';
END
ELSE
    PRINT 'Database [PlacementTracker] already exists - left untouched.';
GO

USE [PlacementTracker];
GO

/* ----------------------------------------------------------------------------
   2. dbo.Applications
   ---------------------------------------------------------------------------- */
IF OBJECT_ID(N'dbo.Applications', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Applications
    (
        id          INT             IDENTITY(1,1)
                                    CONSTRAINT PK_Applications PRIMARY KEY,

        company     NVARCHAR(100)   NOT NULL
                                    CONSTRAINT CK_App_Company
                                    CHECK (LEN(company) BETWEEN 2 AND 100),

        role        NVARCHAR(100)   NOT NULL
                                    CONSTRAINT CK_App_Role
                                    CHECK (LEN(role) BETWEEN 2 AND 100),

        status      NVARCHAR(20)    NOT NULL
                                    CONSTRAINT CK_App_Status
                                    CHECK (status IN (N'Wishlist',
                                                      N'Applied',
                                                      N'Interview',
                                                      N'Offer',
                                                      N'Rejected')),

        /* Required for every status except Wishlist. */
        applied_on  DATE            NULL,

        job_url     NVARCHAR(500)   NULL
                                    CONSTRAINT CK_App_JobUrl
                                    CHECK (job_url IS NULL
                                            OR job_url LIKE N'http://%'
                                            OR job_url LIKE N'https://%'),

        notes       NVARCHAR(1000)  NULL,

        created_at  DATETIME2(0)    NOT NULL
                                    CONSTRAINT DF_App_Created
                                    DEFAULT (SYSUTCDATETIME()),

        /* Application-level rule: Wishlist is the only status allowed
           without an application date. */
        CONSTRAINT CK_App_AppliedOn
            CHECK (status = N'Wishlist' OR applied_on IS NOT NULL)
    );
    PRINT 'Created table dbo.Applications.';
END
ELSE
    PRINT 'Table dbo.Applications already exists - left untouched.';
GO

/* ----------------------------------------------------------------------------
   3. Indexes (dashboard ordering + search helper)
   ---------------------------------------------------------------------------- */
IF NOT EXISTS (SELECT 1
                 FROM sys.indexes
                WHERE object_id = OBJECT_ID(N'dbo.Applications')
                  AND name      = N'IX_Applications_IdDesc')
    CREATE NONCLUSTERED INDEX IX_Applications_IdDesc
        ON dbo.Applications (id DESC);
GO

IF NOT EXISTS (SELECT 1
                 FROM sys.indexes
                WHERE object_id = OBJECT_ID(N'dbo.Applications')
                  AND name      = N'IX_Applications_Status')
    CREATE NONCLUSTERED INDEX IX_Applications_Status
        ON dbo.Applications (status);
GO

/* ----------------------------------------------------------------------------
   4. Read back what was created
   ---------------------------------------------------------------------------- */
SET NOCOUNT ON;
SELECT DB_NAME()                                   AS [database],
       OBJECT_SCHEMA_NAME(OBJECT_ID(N'dbo.Applications')) AS [schema],
       OBJECT_NAME(OBJECT_ID(N'dbo.Applications'))       AS [table];
GO
