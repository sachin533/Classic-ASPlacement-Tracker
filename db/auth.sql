/* ============================================================================
   Placement Tracker  -  db\auth.sql
   ----------------------------------------------------------------------------
   Creates the dbo.Users table that backs login protection.

   The script is NON-DESTRUCTIVE and idempotent:
     * the table is only created when it does not already exist,
     * no rows are ever inserted, updated or deleted - administrators are
       created by db\provision_admin.ps1, which is the ONLY place a password
       (or rather its hash) is ever written.

   It is safe to run repeatedly, and it never touches dbo.Applications.

   Run it after db\schema.sql, from SSMS (Open File -> Execute) or:

       sqlcmd -S .\SQLEXPRESS -E -C -b -i db\auth.sql

   Password storage
   ----------------
   No plaintext password is stored anywhere in the database or in any file.
   Each row keeps a 32-character random salt and

        password_hash = HASHBYTES('SHA2_256', salt + password)

   which is SQL Server's SHA-256 over the UTF-16 (UCS-2) bytes of the
   concatenation.  db\provision_admin.ps1 computes exactly the same digest
   with .NET, so a hash written by the provisioning script is byte-for-byte
   what the login query recomputes (verified: all three sources produce
   0x4FF19F48BAA002221DC71B4DA007EAC899108EC4E9219511BB18B8C9CCB329E2 for a
   known salt/password pair).

   ============================================================================ */

SET NOCOUNT ON;
GO

USE [PlacementTracker];
GO

IF OBJECT_ID(N'dbo.Users', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Users
    (
        id            INT             IDENTITY(1,1)
                                      CONSTRAINT PK_Users PRIMARY KEY,

        /* Login name, always lower case so matching is culture-proof. */
        username      NVARCHAR(64)    NOT NULL
                                      CONSTRAINT UQ_Users_Username UNIQUE,

        /* SHA-256 digest of salt + password (32 bytes). */
        password_hash VARBINARY(32)   NOT NULL,

        /* Random per-user salt, stored as lower-case hex. */
        salt          NVARCHAR(32)    NOT NULL,

        is_active     BIT             NOT NULL
                                      CONSTRAINT DF_Users_IsActive DEFAULT (1),

        last_login_at DATETIME2(0)    NULL,

        created_at    DATETIME2(0)    NOT NULL
                                      CONSTRAINT DF_Users_Created
                                      DEFAULT (SYSUTCDATETIME()),

        CONSTRAINT CK_Users_Username CHECK
            (username = LOWER(username)
             AND username = LTRIM(RTRIM(username))
             AND LEN(username) BETWEEN 3 AND 64),

        CONSTRAINT CK_Users_Salt CHECK
            (LEN(salt) BETWEEN 16 AND 32
             AND salt = LOWER(salt))
    );

    PRINT 'Created table dbo.Users.';
END
ELSE
    PRINT 'Table dbo.Users already exists - left untouched.';
GO

/* Nothing else is needed for lookups: the UNIQUE constraint on username
   already creates a covering index for the login query
       SELECT id, password_hash, salt FROM dbo.Users WHERE username = ?   */
GO
