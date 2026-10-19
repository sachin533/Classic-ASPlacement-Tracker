<#
.SYNOPSIS
  Creates or resets the Placement Tracker administrator account.

.DESCRIPTION
  The application has NO default account and NO default password: until this
  script has been run at least once, sign-in is refused for everyone and the
  login page explains that no administrator exists yet.

  This script writes exactly one row into dbo.Users.  It stores a random
  per-user salt and the SHA-256 digest of (salt + password) - the plaintext
  password is never written to the database, to a file, or to the screen.

  The digest is computed with .NET as SHA256(UTF-16LE(salt + password)),
  which is byte-identical to SQL Server's HASHBYTES('SHA2_256', salt + password)
  - verified on this machine - so the login query in inc\auth.asp reproduces
  it exactly.

  Run it from the project root:

      # prompts for the password (nothing is echoed, nothing is stored)
      powershell -ExecutionPolicy Bypass -File .\db\provision_admin.ps1

      # or supply it for scripted/automated use
      powershell -ExecutionPolicy Bypass -File .\db\provision_admin.ps1 -Username admin -Password '...'

.PARAMETER Username
  Login name.  Normalised to lower case; 3-64 characters from a-z 0-9 _ . -
  Defaults to 'admin'.

.PARAMETER Password
  Optional.  When omitted the script prompts securely on the console.
  Minimum 8 characters.  Only ever held in memory for the duration of the run.

.PARAMETER Remove
  Deletes the account instead of creating it.  Useful to retire a
  provisioned login.

.PARAMETER SqlServer
  SQL Server instance.  Defaults to .\SQLEXPRESS.

.PARAMETER Database
  Target database.  Defaults to PlacementTracker.  The smoke tests pass
  PlacementTracker_Test so they never touch production data.

.NOTES
  Exit code 0 = the account was created, updated or removed as requested.
  Exit code 1 = nothing was changed (bad input, or the table is missing).
#>
[CmdletBinding()]
param(
    [string]   $Username   = 'admin',
    [string]   $Password,
    [switch]   $Remove,
    [string]   $SqlServer  = '.\SQLEXPRESS',
    [string]   $Database   = 'PlacementTracker'
)

$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$AuthSqlFile = Join-Path $ProjectRoot 'db\auth.sql'

# ---------------------------------------------------------------------------
#  Input validation
# ---------------------------------------------------------------------------
$name = ([string]$Username).Trim().ToLowerInvariant()
if ($name -notmatch '^[a-z0-9_.\-]{3,64}$') {
    Write-Error ("Username '{0}' is not valid: use 3-64 characters from a-z, 0-9, _ . - (lower case)." -f $Username)
    exit 1
}

# Read the password from the console unless one was supplied.  Read-Host
# -AsSecureString keeps it off the screen and out of the shell history.
$plain = $Password
if (-not $Remove) {
    if (-not $plain) {
        $sec = Read-Host -Prompt "Password for '$name'" -AsSecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
        try   { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

        $confirm = Read-Host -Prompt 'Confirm password' -AsSecureString
        $bstr2 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirm)
        try   { $confirmPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr2) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr2) }

        if ($plain -cne $confirmPlain) {
            Write-Error 'The two passwords did not match - nothing was changed.'
            exit 1
        }
    }
    if ($plain.Length -lt 8) {
        Write-Error 'The password must be at least 8 characters - nothing was changed.'
        exit 1
    }
    if ($plain.Length -gt 128) {
        Write-Error 'The password must be 128 characters or fewer - nothing was changed.'
        exit 1
    }
}

# ---------------------------------------------------------------------------
#  SQL helpers (System.Data.SqlClient - parameterised, no string building)
# ---------------------------------------------------------------------------
function Invoke-ScalarSql {
    param([string]$Sql, [System.Data.SqlClient.SqlParameter[]]$Params = @())
    $cn = New-Object System.Data.SqlClient.SqlConnection(
        "Server=$SqlServer;Database=$Database;Integrated Security=True;Connect Timeout=10")
    try {
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 60
        foreach ($p in $Params) { [void]$cmd.Parameters.Add($p) }
        return $cmd.ExecuteScalar()
    } finally { $cn.Close() }
}

function Invoke-NonQuerySql {
    param([string]$Sql, [System.Data.SqlClient.SqlParameter[]]$Params = @())
    $cn = New-Object System.Data.SqlClient.SqlConnection(
        "Server=$SqlServer;Database=$Database;Integrated Security=True;Connect Timeout=10")
    try {
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 60
        foreach ($p in $Params) { [void]$cmd.Parameters.Add($p) }
        return $cmd.ExecuteNonQuery()
    } finally { $cn.Close() }
}

# Runs a .sql file the way sqlcmd would: split on GO, all batches on ONE
# connection so an early USE statement still applies to the later batches.
function Invoke-SqlFile {
    param([string]$Path, [string]$Sql)
    if (-not $Sql) { $Sql = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 }

    $batches = @($Sql -split '(?m)^\s*GO\s*;?\s*$' | Where-Object { $_.Trim().Length -gt 0 })
    if ($batches.Count -eq 0) { return }

    $cn = New-Object System.Data.SqlClient.SqlConnection(
        "Server=$SqlServer;Database=$Database;Integrated Security=True;Connect Timeout=10")
    try {
        $cn.Open()
        foreach ($batch in $batches) {
            $cmd = $cn.CreateCommand()
            $cmd.CommandText    = $batch
            $cmd.CommandTimeout = 60
            [void]$cmd.ExecuteNonQuery()
            $cmd.Dispose()
        }
    } finally { $cn.Close() }
}

# ---------------------------------------------------------------------------
#  1. Make sure the table exists (idempotent, and it never touches rows)
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $AuthSqlFile)) {
    Write-Error "Cannot find $AuthSqlFile"
    exit 1
}

# db\auth.sql opens with "USE [PlacementTracker]" so it lands in the right
# database when run by hand; substitute the requested database here, which is
# how the smoke tests keep their work inside PlacementTracker_Test.
$sqlText = (Get-Content -LiteralPath $AuthSqlFile -Raw -Encoding UTF8) -replace 'PlacementTracker', $Database
Invoke-SqlFile -Path $AuthSqlFile -Sql $sqlText

# ---------------------------------------------------------------------------
#  2. Remove (optional)
# ---------------------------------------------------------------------------
if ($Remove) {
    $p = New-Object System.Data.SqlClient.SqlParameter('@u', [Data.SqlDbType]::NVarChar, 64)
    $p.Value = $name
    $gone = Invoke-NonQuerySql -Sql 'DELETE FROM dbo.Users WHERE username = @u' -Params @($p)
    Write-Host ("Removed {0} account(s) named '{1}' in [{2}]." -f $gone, $name, $Database) -ForegroundColor Yellow
    exit 0
}

# ---------------------------------------------------------------------------
#  3. Salt + hash
# ---------------------------------------------------------------------------
# 16 random bytes -> 32 lower-case hex characters.
$rng   = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
$bytes = New-Object byte[] 16
$rng.GetBytes($bytes)
$salt  = ($bytes | ForEach-Object { $_.ToString('x2') }) -join ''
$rng.Dispose()

# SHA-256 over UTF-16LE(salt + password) == SQL Server HASHBYTES('SHA2_256', salt + password)
$sha = [Security.Cryptography.SHA256]::Create()
$hash = $sha.ComputeHash([Text.Encoding]::Unicode.GetBytes($salt + $plain))
$sha.Dispose()

# ---------------------------------------------------------------------------
#  4. Upsert
# ---------------------------------------------------------------------------
$pUser = New-Object System.Data.SqlClient.SqlParameter('@u',  [Data.SqlDbType]::NVarChar, 64)
$pUser.Value = $name
$pSalt = New-Object System.Data.SqlClient.SqlParameter('@s',  [Data.SqlDbType]::NVarChar, 32)
$pSalt.Value = $salt
$pHash = New-Object System.Data.SqlClient.SqlParameter('@h',  [Data.SqlDbType]::VarBinary, 32)
$pHash.Value = $hash

$updated = Invoke-NonQuerySql -Sql 'UPDATE dbo.Users SET password_hash = @h, salt = @s, is_active = 1 WHERE username = @u' `
                              -Params @($pHash, $pSalt, $pUser)

$action = 'updated'
if ($updated -eq 0) {
    Invoke-NonQuerySql -Sql 'INSERT INTO dbo.Users (username, password_hash, salt) VALUES (@u, @h, @s)' `
                       -Params @($pUser, $pHash, $pSalt) | Out-Null
    $action = 'created'
}

$exists = Invoke-ScalarSql -Sql 'SELECT COUNT(*) FROM dbo.Users' 
Write-Host ""
Write-Host ("{0} administrator '{1}' in [{2}]." -f (Get-Culture).TextInfo.ToTitleCase($action), $name, $Database) -ForegroundColor Green
Write-Host ("  rows in dbo.Users : {0}" -f $exists) -ForegroundColor DarkGray
Write-Host ("  last_login_at     : {0}" -f $(if ($action -eq 'created') { 'never' } else { 'will be set at the next sign-in' })) -ForegroundColor DarkGray
Write-Host ""
Write-Host '  The password itself was not stored or displayed anywhere.' -ForegroundColor DarkGray
Write-Host ""
exit 0
