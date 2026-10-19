<#
.SYNOPSIS
  HTTP + SQL smoke tests for the Placement Tracker Classic ASP application.

.DESCRIPTION
  The script never touches the production PlacementTracker database. It
  rebuilds an isolated PlacementTracker_Test database, starts a SECOND
  IIS Express instance on its own port with PT_DB_NAME=PlacementTracker_Test
  (inc\db.asp reads that environment variable), runs the whole CRUD /
  status-code / security suite against it, then shuts the instance down.

  Run from the project root:

      powershell -ExecutionPolicy Bypass -File .\tests\smoke.ps1

  Exit code 0 = every test passed, 1 = at least one failure, 2 = could not
  start the test server.

.PARAMETER Port
  Port for the throw-away test instance. Default 8081 so it never collides
  with the development site on 8080.

.PARAMETER SqlServer
  SQL Server instance. Defaults to .\SQLEXPRESS.

.PARAMETER SkipServerStart
  Use an already-running test instance (for example when IIS Express cannot
  be launched from this session).
#>
[CmdletBinding()]
param(
    [string] $SqlServer = '.\SQLEXPRESS',
    [int]    $Port      = 8081,
    [switch] $SkipServerStart
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$BaseUrl     = "http://localhost:$Port"
$TestDb      = 'PlacementTracker_Test'
$SchemaFile  = Join-Path $ProjectRoot 'db\schema.sql'
$SeedFile    = Join-Path $ProjectRoot 'db\seed.sql'
$Seed500File = Join-Path $ProjectRoot 'db\seed_500.sql'

$IisExpress = 'C:\Program Files\IIS Express\iisexpress.exe'
if (-not (Test-Path $IisExpress)) { $IisExpress = 'C:\Program Files (x86)\IIS Express\iisexpress.exe' }
if (-not (Test-Path $IisExpress)) { throw 'IIS Express is not installed on this machine.' }

# ===========================================================================
#  Result tracking
# ===========================================================================
$script:Passed  = 0
$script:Failed  = 0
$script:Results = New-Object System.Collections.ArrayList

function Report {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    if ($Ok) { $script:Passed++ } else { $script:Failed++ }
    [void]$script:Results.Add([pscustomobject]@{ Test = $Name; Result = $(if ($Ok) { 'PASS' } else { 'FAIL' }); Detail = $Detail })
    $suffix = if ($Detail) { "  ($Detail)" } else { '' }
    Write-Host ('{0,-5} {1}{2}' -f $(if ($Ok) { 'PASS' } else { 'FAIL' }), $Name, $suffix) `
        -ForegroundColor $(if ($Ok) { 'Green' } else { 'Red' })
}

# ===========================================================================
#  SQL helpers  (System.Data.SqlClient - no sqlcmd dependency)
# ===========================================================================
function New-Conn([string]$Database) {
    New-Object System.Data.SqlClient.SqlConnection(
        "Server=$SqlServer;Database=$Database;Integrated Security=True;Connect Timeout=10")
}

function Exec-Sql {
    param([string]$Database, [string]$Sql)
    $cn = New-Conn $Database
    try {
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 60
        [void]$cmd.ExecuteNonQuery()
    } finally {
        $cn.Close()
    }
}

# Runs a .sql file the way sqlcmd would: the file may be split into batches
# with GO separators, which ADO's ExecuteNonQuery cannot parse.  All batches
# run on ONE connection so that a USE statement in an early batch still
# applies to the ones after it.
function Exec-SqlScript {
    param([string]$Database, [string]$Path, [string]$Sql)
    if (-not $Sql) { $Sql = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 }

    $batches = @($Sql -split '(?m)^\s*GO\s*;?\s*$' | Where-Object { $_.Trim().Length -gt 0 })
    if ($batches.Count -eq 0) { return }

    $cn = New-Conn $Database
    try {
        $cn.Open()
        foreach ($batch in $batches) {
            $cmd = $cn.CreateCommand()
            $cmd.CommandText    = $batch
            $cmd.CommandTimeout = 60
            [void]$cmd.ExecuteNonQuery()
            $cmd.Dispose()
        }
    } finally {
        $cn.Close()
    }
}

function Scalar {
    param([string]$Database, [string]$Sql)
    $cn = New-Conn $Database
    try {
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 60
        return $cmd.ExecuteScalar()
    } finally {
        $cn.Close()
    }
}

function Rebuild-TestDatabase {
    Write-Host "`n=== Rebuilding $TestDb ===" -ForegroundColor Cyan

    Exec-Sql 'master' @"
IF DB_ID(N'$TestDb') IS NOT NULL
BEGIN
    ALTER DATABASE [$TestDb] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [$TestDb];
END
"@

    # Reuse the real schema script so the test database is byte-for-byte the
    # same shape as production.  Only the database name is substituted.
    $sql = (Get-Content -LiteralPath $SchemaFile -Raw -Encoding UTF8) -replace 'PlacementTracker', $TestDb
    Exec-SqlScript -Database 'master' -Path $SchemaFile -Sql $sql

    # Seed the same demonstration rows production starts with, so section 1
    # can check the dashboard against real data.  seed.sql is idempotent and
    # opens with "USE [PlacementTracker]", so the database name is substituted
    # here too - otherwise this line would redirect the inserts into the
    # production database.
    $seedSql = (Get-Content -LiteralPath $SeedFile -Raw -Encoding UTF8) -replace 'PlacementTracker', $TestDb
    Exec-SqlScript -Database $TestDb -Path $SeedFile -Sql $seedSql

    $count = Scalar $TestDb 'SELECT COUNT(*) FROM dbo.Applications'
    Write-Host "Ready - $TestDb.dbo.Applications has $count rows.`n" -ForegroundColor DarkGray
}

function Reset-TestRows {
    Exec-Sql $TestDb @'
DELETE FROM dbo.Applications;
DBCC CHECKIDENT ('dbo.Applications', RESEED, 0);
'@
}

function DbCount {
    [int](Scalar $TestDb 'SELECT COUNT(*) FROM dbo.Applications')
}

function DbCountWhere([string]$Where) {
    [int](Scalar $TestDb "SELECT COUNT(*) FROM dbo.Applications WHERE $Where")
}

# ---------------------------------------------------------------------------
#  Search / pagination helpers used by sections 13-17.
#
#  DbSearchCount mirrors exactly the WHERE clause inc\db.asp builds
#  (company / role / status / job_url / notes) so the number shown in the
#  "of N applications" summary can be compared against the database rather
#  than against itself.
# ---------------------------------------------------------------------------
function DbSearchCount([string]$Term) {
    $t = $Term -replace "'", "''"
    $sql = "SELECT COUNT(*) FROM dbo.Applications
             WHERE company LIKE '%$t%' OR role LIKE '%$t%' OR status LIKE '%$t%'
                OR job_url LIKE '%$t%' OR notes LIKE '%$t%'"
    [int](Scalar $TestDb $sql)
}

function SampleRowCount {
    [int](Scalar $TestDb "
        SELECT COUNT(*) FROM dbo.Applications
         WHERE created_at >= '2024-01-01' AND created_at < '2024-01-02'")
}

function Run-Seed500 {
    # seed_500.sql opens with "USE [PlacementTracker]" - swap it for the test
    # database so this can never write into production data.
    $sql = (Get-Content -LiteralPath $Seed500File -Raw -Encoding UTF8) -replace 'PlacementTracker', $TestDb
    Exec-SqlScript -Database $TestDb -Path $Seed500File -Sql $sql
}

# Total printed in "Showing 21-40 of N applications" (-1 when absent).
function Get-ResultTotal([string]$Body) {
    if ($Body -match 'No applications match') { return 0 }
    if ($Body -match 'of\s+(\d+)\s+application') { return [int]$Matches[1] }
    return -1
}

# The "#" column of every row on the current page.
# The leading comma stops PowerShell unwrapping a zero- or one-row result,
# so callers always get a real array back and [0] / .Count stay safe.
function Get-RowNums([string]$Body) {
    $out = @([regex]::Matches($Body, '<td class="col-num">(\d+)</td>') |
        ForEach-Object { [int]$_.Groups[1].Value })
    return , $out
}

# The record id behind every row's company link.
function Get-RowIds([string]$Body) {
    $out = @([regex]::Matches($Body, '<td class="col-company"><a href="/view\.asp\?id=(\d+)">') |
        ForEach-Object { [int]$_.Groups[1].Value })
    return , $out
}

function Build-Url([string]$Path, [hashtable]$Query) {
    $pairs = @()
    foreach ($k in $Query.Keys) {
        $pairs += ('{0}={1}' -f $k, [uri]::EscapeDataString([string]$Query[$k]))
    }
    "$BaseUrl$Path" + '?' + ($pairs -join '&')
}

function DbRow {
    param([int]$Id, [string]$Column)
    $v = Scalar $TestDb "SELECT [$Column] FROM dbo.Applications WHERE id = $Id"
    # ExecuteScalar() returns [DBNull]::Value (not $null) for a NULL column,
    # which would make every "$null -eq (DbRow ...)" assertion fail even
    # though the database really is NULL.  Normalise both cases here.
    if ($null -eq $v -or $v -is [DBNull]) { return $null }
    # DATE/DATETIME2 come back DateTime objects formatted in the local culture;
    # the app stores and renders them as yyyy-MM-dd, so compare like for like.
    if ($v -is [DateTime]) { return $v.ToString('yyyy-MM-dd') }
    return [string]$v
}

# ===========================================================================
#  HTTP helpers
#
#  These use System.Net.HttpWebRequest rather than Invoke-WebRequest.  The
#  cmdlet throws on any 4xx/5xx and its exception response stream comes back
#  empty, so the error pages this suite exists to verify would look blank.
#  HttpWebRequest gives the status, the Location header AND the body for
#  error responses, which is exactly what the assertions below need.
# ===========================================================================
$script:CookieJar = New-Object System.Net.CookieContainer

function Reset-Session {
    # Fresh cookie jar, and therefore a fresh session and CSRF token.
    $script:CookieJar = New-Object System.Net.CookieContainer
}

function Invoke-Http {
    param([string]$Method = 'GET', [string]$Url, [hashtable]$Body)

    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Method             = $Method
    $req.CookieContainer    = $script:CookieJar
    $req.AllowAutoRedirect  = $false      # PRG must be observable
    $req.Timeout            = 30000
    $req.ReadWriteTimeout   = 30000
    $req.UserAgent          = 'PlacementTracker-SmokeTest'

    if ($Body) {
        $req.ContentType = 'application/x-www-form-urlencoded'
        $pairs = foreach ($k in $Body.Keys) {
            '{0}={1}' -f [uri]::EscapeDataString([string]$k), [uri]::EscapeDataString([string]$Body[$k])
        }
        $bytes = [Text.Encoding]::UTF8.GetBytes(($pairs -join '&'))
        $req.ContentLength = $bytes.Length
        $stream = $req.GetRequestStream()
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Close()
    }

    $resp = $null
    try {
        $resp = $req.GetResponse()
    } catch [System.Net.WebException] {
        $resp = $_.Exception.Response
        if (-not $resp) {
            # Transport-level failure (server down, refused, timed out).
            return [pscustomobject]@{ Code = 0; Loc = ''; Body = ''; Error = $_.Exception.Message }
        }
    }

    $code = [int]$resp.StatusCode
    $loc  = [string]$resp.Headers['Location']

    $text = ''
    try {
        $reader = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
        $reader.Close()
    } catch { }
    $resp.Close()

    return [pscustomobject]@{ Code = $code; Loc = $loc; Body = $text; Error = '' }
}

function New-AppSession { Reset-Session }

function Get-Csrf {
    $r = Invoke-Http -Url "$BaseUrl/new.asp"
    if ($r.Body -match 'name="csrf"\s+value="([0-9a-f]+)"') { return $Matches[1] }
    throw 'Could not find the CSRF token on /new.asp'
}

function Http {
    param([ValidateSet('Get','Post')][string]$Method = 'Get',
          [string]$Url,
          [hashtable]$Body)
    Invoke-Http -Method $Method -Url $Url -Body $Body
}

function Get-Page([string]$Url) {
    Invoke-Http -Url $Url
}

# ===========================================================================
#  Main
# ===========================================================================
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host ' Placement Tracker - smoke tests' -ForegroundColor Cyan
Write-Host " SQL Server : $SqlServer" -ForegroundColor DarkGray
Write-Host " Database   : $TestDb (isolated - production is untouched)" -ForegroundColor DarkGray
Write-Host " Base URL   : $BaseUrl" -ForegroundColor DarkGray
Write-Host '==========================================================' -ForegroundColor Cyan

Rebuild-TestDatabase

$proc        = $null
$testConfig  = Join-Path $env:TEMP "placement-tracker-test-$Port.config"

if (-not $SkipServerStart) {
    Write-Host 'Generating a dedicated IIS Express configuration...' -ForegroundColor Cyan
    & (Join-Path $ProjectRoot 'iisexpress\start.ps1') -GenerateOnly -Port $Port -Config $testConfig | Out-Null

    # start.ps1 binds to "<ip>:<port>:localhost" - move it to the test port.
    $cfgText = Get-Content -LiteralPath $testConfig -Raw -Encoding UTF8
    $cfgText = $cfgText -replace (':\d+:localhost'), (':{0}:localhost' -f $Port)
    Set-Content -LiteralPath $testConfig -Value $cfgText -Encoding UTF8

    # Point this process (and the child iisexpress.exe) at the test database.
    $env:PT_DB_SERVER = $SqlServer
    $env:PT_DB_NAME   = $TestDb

    Write-Host "Starting IIS Express on $BaseUrl (database: $TestDb)..." -ForegroundColor Cyan
    $proc = Start-Process -FilePath $IisExpress `
                          -ArgumentList @("/config:`"$testConfig`"", '/site:PlacementTracker') `
                          -PassThru -WindowStyle Hidden
    $env:PT_DB_NAME = ''
}

$ready = $false
for ($i = 0; $i -lt 80; $i++) {
    if ((Invoke-Http -Url "$BaseUrl/").Code -eq 200) { $ready = $true; break }
    Start-Sleep -Milliseconds 500
}

if (-not $ready) {
    Write-Host "`nBLOCKER: the test web server never answered at $BaseUrl." -ForegroundColor Red
    Write-Host 'Check that IIS Express is installed and that ASP is enabled.'
    if ($proc) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    exit 2
}

Write-Host "`nServer is up - running tests...`n" -ForegroundColor Cyan

try {

# ===========================================================================
Write-Host '--- 1. Pages load ---' -ForegroundColor Yellow

$r = Get-Page "$BaseUrl/"
Report 'GET / returns 200' ($r.Code -eq 200) "got $($r.Code)"
Report 'Dashboard lists Google' ($r.Body -match 'Google') ''
Report 'Dashboard shows Applied badge'     ($r.Body -match 'badge-applied') ''
Report 'Dashboard shows Interview badge'   ($r.Body -match 'badge-interview') ''
Report 'Dashboard shows Offer badge'       ($r.Body -match 'badge-offer') ''
Report 'Dashboard shows Rejected badge'    ($r.Body -match 'badge-rejected') ''
Report 'Dashboard shows Wishlist badge'    ($r.Body -match 'badge-wishlist') ''
Report 'Dashboard rows link to view.asp'   ($r.Body -match 'view\.asp\?id=[0-9]+') ''
Report 'Dashboard rows link to edit.asp'   ($r.Body -match 'edit\.asp\?id=[0-9]+') ''
Report 'Dashboard has a search box'        ($r.Body -match 'name="q"') ''

$r = Get-Page "$BaseUrl/new.asp"
Report 'GET /new.asp returns 200' ($r.Code -eq 200) "got $($r.Code)"
Report 'Create form carries a CSRF token' ($r.Body -match 'name="csrf"\s+value="[0-9a-f]+"') ''
Report 'Create form uses a two-column grid' ($r.Body -match 'form-grid') ''
Report 'Create form marks required fields' ($r.Body -match 'class="req"') ''

$r = Get-Page "$BaseUrl/static/styles.css"
Report 'GET /static/styles.css returns 200' ($r.Code -eq 200) "got $($r.Code)"
$r = Get-Page "$BaseUrl/static/app.js"
Report 'GET /static/app.js returns 200' ($r.Code -eq 200) "got $($r.Code)"

# ===========================================================================
Write-Host "`n--- 2. Status codes ---" -ForegroundColor Yellow

$r = Get-Page "$BaseUrl/view.asp?id=999999"
Report 'GET /view.asp?id=999999 -> 404' ($r.Code -eq 404) "got $($r.Code)"
Report '404 page shows a friendly message' ($r.Body -match 'Page not found') ''

$r = Get-Page "$BaseUrl/view.asp?id=abc"
Report 'GET /view.asp?id=abc -> 404' ($r.Code -eq 404) "got $($r.Code)"

$r = Get-Page "$BaseUrl/view.asp?id=1%20OR%201=1"
Report 'GET /view.asp?id=1 OR 1=1 -> 404' ($r.Code -eq 404) "got $($r.Code)"

$r = Get-Page "$BaseUrl/view.asp?id=-5"
Report 'GET /view.asp?id=-5 -> 404' ($r.Code -eq 404) "got $($r.Code)"

$r = Get-Page "$BaseUrl/edit.asp?id=abc"
Report 'GET /edit.asp?id=abc -> 404' ($r.Code -eq 404) "got $($r.Code)"

$r = Get-Page "$BaseUrl/delete.asp"
Report 'GET /delete.asp -> 405' ($r.Code -eq 405) "got $($r.Code)"

$r = Get-Page "$BaseUrl/nothing-here-at-all.asp"
Report 'Unknown path -> 404 through the custom error page' ($r.Code -eq 404) "got $($r.Code)"

# ===========================================================================
Write-Host "`n--- 3. Internal folders are not web reachable ---" -ForegroundColor Yellow

$r = Get-Page "$BaseUrl/inc/db.asp"
Report 'GET /inc/db.asp is blocked'      ($r.Code -ne 200) "got $($r.Code)"
Report 'Connection string is not exposed' ($r.Body -notmatch 'MSOLEDBSQL|SQLOLEDB|Integrated Security') ''
$r = Get-Page "$BaseUrl/inc/helpers.asp"
Report 'GET /inc/helpers.asp is blocked' ($r.Code -ne 200) "got $($r.Code)"
$r = Get-Page "$BaseUrl/db/schema.sql"
Report 'GET /db/schema.sql is blocked'   ($r.Code -ne 200) "got $($r.Code)"
$r = Get-Page "$BaseUrl/db/seed.sql"
Report 'GET /db/seed.sql is blocked'     ($r.Code -ne 200) "got $($r.Code)"
$r = Get-Page "$BaseUrl/tests/smoke.ps1"
Report 'GET /tests/smoke.ps1 is blocked' ($r.Code -ne 200) "got $($r.Code)"

# ===========================================================================
Write-Host "`n--- 4. CSRF protection ---" -ForegroundColor Yellow

Reset-TestRows
New-AppSession

$r = Http Post "$BaseUrl/new.asp" @{ company = 'No Token Ltd'; role = 'Engineer'; status = 'Applied'; applied_on = '2025-01-05' }
Report 'POST without a CSRF token -> 403'    ($r.Code -eq 403) "got $($r.Code)"
Report 'POST without a token wrote nothing'  ((DbCount) -eq 0) "rows=$(DbCount)"

$r = Http Post "$BaseUrl/new.asp" @{ company = 'Bad Token Ltd'; role = 'Engineer'; status = 'Applied'; applied_on = '2025-01-05'; csrf = ('deadbeef' * 8) }
Report 'POST with an invalid token -> 403'   ($r.Code -eq 403) "got $($r.Code)"
Report 'Invalid token wrote nothing'         ((DbCount) -eq 0) "rows=$(DbCount)"

# ===========================================================================
Write-Host "`n--- 5. Create ---" -ForegroundColor Yellow

Reset-TestRows
$csrf = Get-Csrf

$r = Http Post "$BaseUrl/new.asp" @{
    company    = 'Northwind Traders'
    role       = 'Backend Engineer Intern'
    status     = 'Applied'
    applied_on = '2025-02-14'
    job_url    = 'https://careers.northwind.example/jobs/42'
    notes      = 'Referred by a friend.'
    csrf       = $csrf
}
Report 'Valid POST creates the record (302)' ($r.Code -eq 302) "got $($r.Code)"
Report 'Redirects to the new details page' ($r.Loc -match 'view\.asp\?id=1') "Location=$($r.Loc)"

$newId = 0
if ($r.Loc -match 'view\.asp\?id=(\d+)') { $newId = [int]$Matches[1] }
Report 'Database holds exactly one row' ((DbCount) -eq 1) "rows=$(DbCount)"
Report 'company persisted'    ((DbRow $newId 'company')    -eq 'Northwind Traders') ''
Report 'role persisted'       ((DbRow $newId 'role')       -eq 'Backend Engineer Intern') ''
Report 'status persisted'     ((DbRow $newId 'status')     -eq 'Applied') ''
Report 'applied_on persisted' ((DbRow $newId 'applied_on') -eq '2025-02-14') "got $(DbRow $newId 'applied_on')"
Report 'job_url persisted'    ((DbRow $newId 'job_url')    -eq 'https://careers.northwind.example/jobs/42') ''
Report 'notes persisted'      ((DbRow $newId 'notes')      -eq 'Referred by a friend.') ''

$r = Get-Page "$BaseUrl/view.asp?id=$newId"
Report 'Details page returns 200'           ($r.Code -eq 200) "got $($r.Code)"
Report 'Details page shows the company'     ($r.Body -match 'Northwind Traders') ''
Report 'Details page shows the badge class' ($r.Body -match 'badge-applied') ''
Report 'Details page renders the job URL'   ($r.Body -match 'careers\.northwind\.example') ''
Report 'Details page shows created at (UTC)' ($r.Body -match '\(UTC\)') ''

$r = Get-Page "$BaseUrl/"
Report 'New record appears on the dashboard' ($r.Body -match 'Northwind Traders') ''

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{ company = 'Wish Co'; role = 'Grad Scheme'; status = 'Wishlist'; csrf = $csrf }
Report 'Wishlist row may omit the date'        ($r.Code -eq 302) "got $($r.Code)"
Report 'Wishlist applied_on stored as NULL'    ($null -eq (DbRow 2 'applied_on')) "value=$(DbRow 2 'applied_on')"

# ===========================================================================
Write-Host "`n--- 6. Read and search ---" -ForegroundColor Yellow

$r = Get-Page "$BaseUrl/?q=Northwind"
Report 'Search matches a company name' (($r.Code -eq 200) -and ($r.Body -match 'Northwind Traders')) "code=$($r.Code)"
Report 'Search excludes non-matching rows' ($r.Body -notmatch 'Wish Co') ''

$r = Get-Page "$BaseUrl/?q=wishlist"
Report 'Search matches the status column' ($r.Body -match 'Wish Co') ''

$r = Get-Page "$BaseUrl/?q=zzzz-nothing-matches"
Report 'No results shows a search empty state' (($r.Code -eq 200) -and ($r.Body -match 'No matching applications')) "code=$($r.Code)"

# ===========================================================================
Write-Host "`n--- 7. Update ---" -ForegroundColor Yellow

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/edit.asp?id=$newId" @{
    id         = "$newId"
    company    = 'Northwind Traders International'
    role       = 'Senior Platform Engineer'
    status     = 'Interview'
    applied_on = '2025-03-01'
    job_url    = 'https://careers.northwind.example/jobs/99'
    notes      = "Updated notes with a quoted word and an apostrophe: don't miss out."
    csrf       = $csrf
}
Report 'Valid edit POST -> 302'        ($r.Code -eq 302) "got $($r.Code)"
Report 'Edit redirects to the details page' ($r.Loc -match "view\.asp\?id=$newId$") "Location=$($r.Loc)"
Report 'company updated'      ((DbRow $newId 'company')       -eq 'Northwind Traders International') ''
Report 'role updated'         ((DbRow $newId 'role')          -eq 'Senior Platform Engineer') ''
Report 'status updated'       ((DbRow $newId 'status')        -eq 'Interview') ''
Report 'applied_on updated'   ((DbRow $newId 'applied_on')    -eq '2025-03-01') ''
Report 'job_url updated'      ((DbRow $newId 'job_url')       -eq 'https://careers.northwind.example/jobs/99') ''
Report 'notes updated'        ((DbRow $newId 'notes').Contains("don't miss out")) ''

$r = Get-Page "$BaseUrl/view.asp?id=$newId"
Report 'Badge colour follows the new status' ($r.Body -match 'badge-interview') ''
Report 'Updated company is rendered' ($r.Body -match 'Northwind Traders International') ''

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/edit.asp?id=999999" @{ id = '999999'; company = 'Ghost'; role = 'Ghost'; status = 'Applied'; applied_on = '2025-01-01'; csrf = $csrf }
Report 'POST /edit.asp with an unknown id -> 404' ($r.Code -eq 404) "got $($r.Code)"

# ===========================================================================
Write-Host "`n--- 8. Server-side validation ---" -ForegroundColor Yellow

$before = DbCount
$csrf   = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{
    company = ''; role = ''; status = 'NotAStatus'; applied_on = '2025-13-45'
    job_url = 'javascript:alert(1)'; notes = ''; csrf = $csrf
}
Report 'Invalid submission returns the form (200)' ($r.Code -eq 200) "got $($r.Code)"
Report 'Nothing was inserted'  ((DbCount) -eq $before) "before=$before after=$(DbCount)"
Report 'Error summary is shown'        ($r.Body -match 'Please fix the') ''
Report 'Company error message'         ($r.Body -match 'Company is required') ''
Report 'Status error message'          ($r.Body -match 'Status must be') ''
Report 'Date error message'            ($r.Body -match 'Applied On must be a real date') ''
Report 'URL scheme error message'      ($r.Body -match 'Job URL must be a full http') ''

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{ company = 'x'; role = 'y'; status = 'Applied'; applied_on = '2025-01-01'; csrf = $csrf }
Report '1-character company/role rejected' (($r.Code -eq 200) -and ($r.Body -match 'at least 2 characters')) "code=$($r.Code)"

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{ company = 'Long Notes Ltd'; role = 'Tester'; status = 'Applied'; applied_on = '2025-01-01'; notes = ('a' * 1001); csrf = $csrf }
Report 'Notes over 1000 characters rejected' (($r.Code -eq 200) -and ($r.Body -match '1,000 characters or fewer')) "code=$($r.Code)"

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{ company = 'Long Url Ltd'; role = 'Tester'; status = 'Applied'; applied_on = '2025-01-01'; job_url = ('https://example.com/' + ('a' * 600)); csrf = $csrf }
Report 'Job URL over 500 characters rejected' (($r.Code -eq 200) -and ($r.Body -match '500 characters or fewer')) "code=$($r.Code)"

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{ company = 'No Date Ltd'; role = 'Dev'; status = 'Offer'; csrf = $csrf }
Report 'Non-Wishlist row without a date rejected' (($r.Code -eq 200) -and ($r.Body -match 'required unless the status is Wishlist')) "code=$($r.Code)"

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{ company = 'Keep Me Ltd'; role = 'Keeper'; status = 'Applied'; applied_on = '2025-04-01'; job_url = 'nope'; csrf = $csrf }
Report 'Entered company survives a failed validation' ($r.Body -match 'value="Keep Me Ltd"') ''
Report 'Entered role survives a failed validation'    ($r.Body -match 'value="Keeper"') ''
Report 'Selected status survives a failed validation' ($r.Body -match 'value="Applied" selected') ''

# ===========================================================================
Write-Host "`n--- 9. Hostile input ---" -ForegroundColor Yellow

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{
    company    = '<script>alert("xss")</script>'
    role       = "Robert'); DROP TABLE dbo.Applications;--"
    status     = 'Applied'
    applied_on = '2025-05-05'
    notes      = '<img src=x onerror=alert(1)> & "quotes" and ''apostrophes'''
    job_url    = 'https://ok.example.com/?a=1&b=2'
    csrf       = $csrf
}
Report 'Hostile submission is stored, not executed' ($r.Code -eq 302) "got $($r.Code)"
Report 'The Applications table still exists' ((DbCount) -ge 1) "rows=$(DbCount)"

$badId = 0
if ($r.Loc -match 'id=(\d+)') { $badId = [int]$Matches[1] }
Report 'Script tag stored verbatim as text'   ((DbRow $badId 'company') -eq '<script>alert("xss")</script>') ''
Report 'SQL injection string stored verbatim' ((DbRow $badId 'role') -eq "Robert'); DROP TABLE dbo.Applications;--") ''

$r = Get-Page "$BaseUrl/view.asp?id=$badId"
Report 'Details page HTML-encodes the script tag' ($r.Body -notmatch '<script>alert') ''
Report 'Details page shows the encoded text'      ($r.Body -match '&lt;script&gt;') ''

$r = Get-Page "$BaseUrl/"
Report 'Dashboard HTML-encodes hostile text' ($r.Body -notmatch '<script>alert') ''

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/new.asp" @{ company = ([char]0x00DC + 'n' + [char]0x00EF + 'cod' + [char]0x00E9 + ' Corp ' + [char]0x65E5 + [char]0x672C); role = 'Developpeur'; status = 'Applied'; applied_on = '2025-06-06'; csrf = $csrf }
Report 'Unicode submission accepted' ($r.Code -eq 302) "got $($r.Code)"
$uniId = 0
if ($r.Loc -match 'id=(\d+)') { $uniId = [int]$Matches[1] }
Report 'Unicode company stored intact' ((DbRow $uniId 'company') -eq ([char]0x00DC + 'n' + [char]0x00EF + 'cod' + [char]0x00E9 + ' Corp ' + [char]0x65E5 + [char]0x672C)) "stored=$(DbRow $uniId 'company')"

# ===========================================================================
Write-Host "`n--- 10. Delete ---" -ForegroundColor Yellow

$before = DbCount
$r = Http Post "$BaseUrl/delete.asp" @{ id = "$badId"; csrf = ('0' * 64) }
Report 'POST with an invalid token -> 403'   ($r.Code -eq 403) "got $($r.Code)"
Report 'Record survived the invalid token'   ((DbCount) -eq $before) "before=$before after=$(DbCount)"

$r = Http Post "$BaseUrl/delete.asp" @{ id = '999999'; csrf = (Get-Csrf) }
Report 'POST with an unknown id -> 404'      ($r.Code -eq 404) "got $($r.Code)"

$r = Http Post "$BaseUrl/delete.asp" @{ id = 'abc'; csrf = (Get-Csrf) }
Report 'POST with a non-numeric id -> 404'   ($r.Code -eq 404) "got $($r.Code)"

$before = DbCount
$r = Get-Page "$BaseUrl/delete.asp?id=$badId"
Report 'GET to delete.asp does not delete (405)' ($r.Code -eq 405) "got $($r.Code)"
Report 'Row count unchanged after the GET'       ((DbCount) -eq $before) "before=$before after=$(DbCount)"

$csrf = Get-Csrf
$r = Http Post "$BaseUrl/delete.asp" @{ id = "$badId"; csrf = $csrf }
Report 'Valid delete POST -> 302'        ($r.Code -eq 302) "got $($r.Code)"
Report 'Delete redirects to the dashboard' ($r.Loc -match 'default\.asp') "Location=$($r.Loc)"
Report 'Row was removed'                 ((DbCount) -eq ($before - 1)) "before=$before after=$(DbCount)"

$r = Get-Page "$BaseUrl/view.asp?id=$badId"
Report 'Deleted record now returns 404'  ($r.Code -eq 404) "got $($r.Code)"

# ===========================================================================
Write-Host "`n--- 11. Empty state ---" -ForegroundColor Yellow

Reset-TestRows
$r = Get-Page "$BaseUrl/"
Report 'Empty table returns 200'                    ($r.Code -eq 200) "got $($r.Code)"
Report 'Empty state heading shown'                  ($r.Body -match 'No applications yet') ''
Report 'Empty state support text shown'             ($r.Body -match 'Start by adding your first job') ''
Report 'Empty state keeps the Add Application button' ($r.Body -match 'Add Application') ''
Report 'Empty state keeps the navigation'           ($r.Body -match 'New Application') ''
Report 'Empty state renders no table rows'          ($r.Body -notmatch 'view\.asp\?id=') ''

# ===========================================================================
Write-Host "`n--- 12. Persistence across a server restart ---" -ForegroundColor Yellow

Reset-TestRows
New-AppSession
$csrf = Get-Csrf
[void](Http Post "$BaseUrl/new.asp" @{ company = 'Durable Ltd'; role = 'Analyst'; status = 'Offer'; applied_on = '2025-07-07'; csrf = $csrf })

if ($proc) {
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    $env:PT_DB_NAME = $TestDb
    $proc = Start-Process -FilePath $IisExpress `
                          -ArgumentList @("/config:`"$testConfig`"", '/site:PlacementTracker') `
                          -PassThru -WindowStyle Hidden
    $env:PT_DB_NAME = ''

    $ready = $false
    for ($i = 0; $i -lt 80; $i++) {
        if ((Invoke-Http -Url "$BaseUrl/").Code -eq 200) { $ready = $true; break }
        Start-Sleep -Milliseconds 500
    }

    if ($ready) {
        New-AppSession
        $r = Get-Page "$BaseUrl/"
        Report 'Data survives a server restart' (($r.Code -eq 200) -and ($r.Body -match 'Durable Ltd')) "code=$($r.Code)"
        Report 'Row count survived the restart' ((DbCount) -eq 1) "rows=$(DbCount)"
    } else {
        Report 'Data survives a server restart' $false 'server did not come back up'
    }
} else {
    Report 'Data survives a server restart' $false '-SkipServerStart was used'
}

# ===========================================================================
Write-Host "`n--- 13. Seeding 500+ demo records ---" -ForegroundColor Yellow

# A record that already existed before seeding - it must survive untouched.
$preservedId      = [int](Scalar $TestDb 'SELECT TOP 1 id FROM dbo.Applications ORDER BY id')
$preservedCompany = [string](Scalar $TestDb "SELECT company FROM dbo.Applications WHERE id = $preservedId")
$rowsBefore       = DbCount

Run-Seed500

$sampleAfterFirst = SampleRowCount
$rowsAfterFirst   = DbCount

Report 'seed_500.sql creates at least 500 sample records' ($sampleAfterFirst -ge 500) "sample=$sampleAfterFirst"
Report 'Seeding inserted the expected 560 rows' (($rowsAfterFirst - $rowsBefore) -eq 560) "before=$rowsBefore after=$rowsAfterFirst inserted=$($rowsAfterFirst - $rowsBefore)"
Report 'Sample rows in the database >= 500' (($rowsAfterFirst - $rowsBefore) -ge 500) "inserted=$($rowsAfterFirst - $rowsBefore)"
Report 'Existing records are preserved during seeding' `
    ([string](Scalar $TestDb "SELECT company FROM dbo.Applications WHERE id = $preservedId") -eq $preservedCompany) `
    "id=$preservedId company=$preservedCompany"

# Names the specification calls out by hand - all of them must exist after seeding.
$knownCompanies = @('Google','Microsoft','Amazon','TCS','Infosys','Wipro','Accenture',
                    'Capgemini','Cognizant','IBM','Oracle','Adobe','NVIDIA','Deloitte')
$missingCompanies = @($knownCompanies | Where-Object { (DbCountWhere "company = '$_'") -eq 0 })
Report 'Realistic company names were seeded' ($missingCompanies.Count -eq 0) ('missing: ' + ($missingCompanies -join ', '))

$knownRoles = @('Software Engineer','Java Developer','Full Stack Developer','Frontend Developer',
                'Backend Developer','Data Analyst','QA Engineer','DevOps Engineer',
                'Graduate Trainee','Software Engineer Intern')
$missingRoles = @($knownRoles | Where-Object { (DbCountWhere "role = '$_'") -eq 0 })
Report 'Realistic job roles were seeded' ($missingRoles.Count -eq 0) ('missing: ' + ($missingRoles -join ', '))

$statusTally = @()
foreach ($s in 'Wishlist','Applied','Interview','Offer','Rejected') {
    if ((DbCountWhere "status = '$s'") -eq 0) { $statusTally += $s }
}
Report 'All five statuses are represented' ($statusTally.Count -eq 0) ('missing: ' + ($statusTally -join ', '))

Report 'Seed rows carry realistic dates' ((DbCountWhere "applied_on IS NOT NULL") -ge 400) "dated=$(DbCountWhere 'applied_on IS NOT NULL')"
Report 'Seed rows include HTTPS job URLs' ((Scalar $TestDb "SELECT COUNT(*) FROM dbo.Applications WHERE job_url LIKE 'https://%'") -gt 0) "https=$(Scalar $TestDb "SELECT COUNT(*) FROM dbo.Applications WHERE job_url LIKE 'https://%'")"

# --- running it a second time must add nothing -----------------------------
Run-Seed500
$rowsAfterSecond  = DbCount
$sampleAfterSecond = SampleRowCount

Report 'Running the seed script twice does not duplicate its records' `
    ($rowsAfterSecond -eq $rowsAfterFirst) "first=$rowsAfterFirst second=$rowsAfterSecond"
Report 'Sample row count is unchanged after the second run' `
    ($sampleAfterSecond -eq $sampleAfterFirst) "sample=$sampleAfterSecond"
Report 'Existing records survive the second run too' `
    ([string](Scalar $TestDb "SELECT company FROM dbo.Applications WHERE id = $preservedId") -eq $preservedCompany) `
    "id=$preservedId"

Write-Host "     seeded $($rowsAfterFirst - $rowsBefore) new rows; database now holds $rowsAfterSecond records." -ForegroundColor DarkGray

# ===========================================================================
Write-Host "`n--- 14. Google-like search ---" -ForegroundColor Yellow

$r = Get-Page (Build-Url '/default.asp' @{ q = 'Microsoft' })
$exp = DbSearchCount 'Microsoft'
Report 'Search by complete company name' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $exp) -and ($r.Body -match 'Microsoft')) `
    "code=$($r.Code) shown=$((Get-ResultTotal $r.Body)) sql=$exp"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'micro' })
$exp = DbSearchCount 'micro'
Report 'Search by partial company name' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $exp) -and ((Get-ResultTotal $r.Body) -gt 0)) `
    "code=$($r.Code) shown=$((Get-ResultTotal $r.Body)) sql=$exp"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'Data Analyst' })
$exp = DbSearchCount 'Data Analyst'
Report 'Search matches the role column' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $exp) -and ($r.Body -match 'Data Analyst')) `
    "shown=$((Get-ResultTotal $r.Body)) sql=$exp"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'Offer' })
$exp = DbSearchCount 'Offer'
Report 'Search matches the status column' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $exp)) "shown=$((Get-ResultTotal $r.Body)) sql=$exp"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'amazon.jobs' })
$exp = DbSearchCount 'amazon.jobs'
Report 'Search matches the job URL column' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $exp) -and ((Get-ResultTotal $r.Body) -gt 0)) `
    "shown=$((Get-ResultTotal $r.Body)) sql=$exp"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'PT-0270' })
$exp = DbSearchCount 'PT-0270'
Report 'Search matches the notes column' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $exp) -and ((Get-ResultTotal $r.Body) -gt 0)) `
    "shown=$((Get-ResultTotal $r.Body)) sql=$exp"

$plain = DbSearchCount 'Microsoft'
$r = Get-Page (Build-Url '/default.asp' @{ q = '   Microsoft   ' })
Report 'Search ignores surrounding whitespace' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $plain)) `
    "shown=$((Get-ResultTotal $r.Body)) sql=$plain"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'zzz-nothing-matches-this' })
Report 'A no-match search shows the empty state' `
    (($r.Code -eq 200) -and ($r.Body -match 'No matching applications') -and ((Get-ResultTotal $r.Body) -eq 0)) `
    "code=$($r.Code)"

$r = Get-Page (Build-Url '/default.asp' @{ q = '' })
Report 'An empty query returns every record' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq (DbCount))) "shown=$((Get-ResultTotal $r.Body)) rows=$(DbCount)"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'group discussion' })
$exp = DbSearchCount 'group discussion'
Report 'Displayed total matches the complete filtered count' `
    ((Get-ResultTotal $r.Body) -eq $exp) "shown=$((Get-ResultTotal $r.Body)) sql=$exp"

$r = Get-Page (Build-Url '/default.asp' @{ status = 'Interview' })
$exp = [int](Scalar $TestDb "SELECT COUNT(*) FROM dbo.Applications WHERE status = 'Interview'")
Report 'Status filter returns exactly the matching rows' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $exp)) "shown=$((Get-ResultTotal $r.Body)) sql=$exp"
Report 'Status filter is preserved in the page links' ($r.Body -like '*status=Interview*')

$r = Get-Page (Build-Url '/default.asp' @{ status = 'NotAStatus' })
Report 'An unknown status value is ignored, not an error' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq (DbCount))) "code=$($r.Code)"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'Engineer'; status = 'Offer' })
$exp = [int](Scalar $TestDb "SELECT COUNT(*) FROM dbo.Applications
    WHERE (company LIKE '%Engineer%' OR role LIKE '%Engineer%' OR status LIKE '%Engineer%'
           OR job_url LIKE '%Engineer%' OR notes LIKE '%Engineer%') AND status = 'Offer'")
Report 'Search and status filter combine' `
    (($r.Code -eq 200) -and ((Get-ResultTotal $r.Body) -eq $exp)) "shown=$((Get-ResultTotal $r.Body)) sql=$exp"

$r = Get-Page "$BaseUrl/default.asp"
Report 'Dashboard has a prominent search input'      ($r.Body -like '*id="dash-search"*')
Report 'Search placeholder text is present'          ($r.Body -like '*Search company, role, status or notes*')
Report 'A clear-search control is rendered'          ($r.Body -like '*js-search-clear*')
Report 'A status filter select is rendered'          ($r.Body -like '*id="dash-status"*')
Report 'The nav search box keeps the query'          ($r.Body -like '*id="site-search"*')

# ===========================================================================
Write-Host "`n--- 15. Numbered pagination ---" -ForegroundColor Yellow

$pageSize   = 20
$totalRows  = DbCount
$totalPages = [int][math]::Ceiling($totalRows / [double]$pageSize)

$r1  = Get-Page (Build-Url '/default.asp' @{ page = '1' })
$ids1 = Get-RowIds $r1.Body
$nums1 = Get-RowNums $r1.Body

Report 'Page 1 returns 200 with a results summary' `
    (($r1.Code -eq 200) -and ($r1.Body -like '*Showing 1*')) "code=$($r1.Code)"
Report 'Page 1 renders numbered pagination' ($r1.Body -like '*<nav class="pager"*')
Report 'Page 1 holds exactly 20 records' ($ids1.Count -eq $pageSize) "rows=$($ids1.Count)"
Report 'Page 1 numbering runs 1-20' `
    (($nums1.Count -eq $pageSize) -and ($nums1[0] -eq 1) -and ($nums1[$nums1.Count - 1] -eq 20)) `
    "$($nums1[0])..$($nums1[$nums1.Count - 1])"

$r2   = Get-Page (Build-Url '/default.asp' @{ page = '2' })
$ids2 = Get-RowIds $r2.Body
$nums2 = Get-RowNums $r2.Body

Report 'Page 1 and page 2 hold different records' `
    (@($ids1 | Where-Object { $ids2 -contains $_ }).Count -eq 0) `
    "overlap=$(@($ids1 | Where-Object { $ids2 -contains $_ }).Count)"
Report 'Page 2 holds exactly 20 records' ($ids2.Count -eq $pageSize) "rows=$($ids2.Count)"
Report 'Page 2 numbering runs 21-40' `
    (($nums2.Count -eq $pageSize) -and ($nums2[0] -eq 21) -and ($nums2[$nums2.Count - 1] -eq 40)) `
    "$($nums2[0])..$($nums2[$nums2.Count - 1])"
$minP1 = ($ids1 | Measure-Object -Minimum).Minimum
$maxP2 = ($ids2 | Measure-Object -Maximum).Maximum
Report 'Rows stay in order across pages (newest first)' ($minP1 -gt $maxP2) "minP1=$minP1 maxP2=$maxP2"

$lastUrl = Build-Url '/default.asp' @{ page = "$totalPages" }
$rLast = Get-Page $lastUrl

Report 'Previous is disabled on page 1' `
    ($r1.Body -like '*is-disabled" aria-disabled="true">Previous</span>*')
Report 'Next is a working link on page 1' `
    ($r1.Body -like '*rel="next" href="/default.asp?page=2">Next</a>*')
Report 'Previous links back to page 1 on page 2' `
    ($r2.Body -like '*rel="prev" href="/default.asp?page=1">Previous</a>*')
Report 'Next links to page 3 on page 2' `
    ($r2.Body -like '*rel="next" href="/default.asp?page=3">Next</a>*')
Report 'Next is disabled on the last page' `
    ($rLast.Body -like '*is-disabled" aria-disabled="true">Next</span>*')
Report 'Previous is a working link on the last page' `
    ($rLast.Body -like '*rel="prev"*Previous</a>*')
Report 'The last page marks itself as the current page' `
    ($rLast.Body -like ('*pager-btn is-active" aria-current="page">{0}*</span>*' -f $totalPages)) `
    "pages=$totalPages"
Report 'A single page of results shows no pager noise' `
    ((Get-Page (Build-Url '/default.asp' @{ q = 'PT-0270' })).Body -notlike '*<nav class="pager"*')

# --- invalid page values ---------------------------------------------------
$expectedTotal = $totalRows
$badDetail = @()
foreach ($badPage in @('abc', '0', '-3', '1e3', '999999999999', '')) {
    $rb = Get-Page (Build-Url '/default.asp' @{ page = $badPage })
    $t  = Get-ResultTotal $rb.Body
    if ($rb.Code -ne 200 -or $t -ne $expectedTotal) {
        $badDetail += ("page='{0}' -> {1}/{2}" -f $badPage, $rb.Code, $t)
    }
}
Report 'Invalid, zero, negative and oversized page values are safe' `
    ($badDetail.Count -eq 0) ($badDetail -join '; ')

# --- search + pagination ---------------------------------------------------
$rq1 = Get-Page (Build-Url '/default.asp' @{ q = 'Engineer' })
$rq2 = Get-Page (Build-Url '/default.asp' @{ q = 'Engineer'; page = '2' })
Report 'Search stays active when navigating between pages' `
    (($rq2.Code -eq 200) -and ($rq2.Body -like '*q=Engineer*') -and `
     ((Get-ResultTotal $rq2.Body) -eq (DbSearchCount 'Engineer'))) `
    "code=$($rq2.Code) shown=$((Get-ResultTotal $rq2.Body)) sql=$(DbSearchCount 'Engineer')"
Report 'Page 2 of a search numbers rows from the filtered set' `
    ((Get-RowNums $rq2.Body)[0] -eq 21) "first=$((Get-RowNums $rq2.Body)[0])"
Report 'The search form never carries a page value (so it resets to 1)' `
    ($rq1.Body -notlike '*name="page"*')
Report 'Submitting a new search starts at page 1' `
    ((Get-RowNums $rq1.Body)[0] -eq 1) "first=$((Get-RowNums $rq1.Body)[0])"
Report 'Page links preserve the search query' `
    ($rq2.Body -like '*default.asp?page=1&amp;q=Engineer*') 'no q= in the pager hrefs'

$rq3 = Get-Page (Build-Url '/default.asp' @{ status = 'Offer'; page = '2' })
Report 'Pagination works together with the status filter' `
    (($rq3.Code -eq 200) -and ($rq3.Body -like '*status=Offer*') -and ((Get-RowNums $rq3.Body)[0] -eq 21)) `
    "code=$($rq3.Code) first=$((Get-RowNums $rq3.Body)[0])"

# ===========================================================================
Write-Host "`n--- 16. Search is injection safe ---" -ForegroundColor Yellow

$r = Get-Page (Build-Url '/default.asp' @{ q = '<script>alert(1)</script>' })
Report 'Script in the search box is encoded, never executed' `
    (($r.Code -eq 200) -and ($r.Body -notlike '*<script>alert(1)</script>*') -and ($r.Body -like '*&lt;script&gt;*')) `
    "code=$($r.Code)"

$r = Get-Page (Build-Url '/default.asp' @{ q = "'; DROP TABLE dbo.Applications;--" })
Report 'SQL injection in the search box returns a normal page' ($r.Code -eq 200) "code=$($r.Code)"
Report 'The Applications table survives the injection attempt' `
    ((Scalar $TestDb "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Applications'") -eq 1)

$r = Get-Page (Build-Url '/default.asp' @{ q = '%' })
Report 'A literal % wildcard is escaped' ($r.Code -eq 200) "code=$($r.Code) shown=$((Get-ResultTotal $r.Body))"
$r = Get-Page (Build-Url '/default.asp' @{ q = '_' })
Report 'A literal _ wildcard is escaped' ($r.Code -eq 200) "code=$($r.Code) shown=$((Get-ResultTotal $r.Body))"
$r = Get-Page (Build-Url '/default.asp' @{ q = '[' })
Report 'A literal [ wildcard is escaped' ($r.Code -eq 200) "code=$($r.Code) shown=$((Get-ResultTotal $r.Body))"

# ===========================================================================
Write-Host "`n--- 17. CRUD still works with search and pagination ---" -ForegroundColor Yellow

$rowsBeforeCrud = DbCount
$csrf = Get-Csrf
$rc = Http Post "$BaseUrl/new.asp" @{
    company    = 'Pagination Corp'
    role       = 'Software Engineer'
    status     = 'Interview'
    applied_on = '2025-09-09'
    notes      = 'Created while pagination and search were active.'
    csrf       = $csrf
}
Report 'Create returns 302 with pagination enabled' ($rc.Code -eq 302) "code=$($rc.Code)"
$newId = 0
if ($rc.Loc -match 'id=(\d+)') { $newId = [int]$Matches[1] }
Report 'Row count grew by one' ((DbCount) -eq ($rowsBeforeCrud + 1)) "before=$rowsBeforeCrud after=$(DbCount)"

$r = Get-Page (Build-Url '/default.asp' @{ q = 'Pagination Corp' })
Report 'The new record is findable through search' ((Get-ResultTotal $r.Body) -eq 1) "shown=$((Get-ResultTotal $r.Body))"
$r = Get-Page "$BaseUrl/default.asp"
Report 'The new record sits on page 1 of the dashboard' ((Get-RowIds $r.Body) -contains $newId)

$csrf = Get-Csrf
$ru = Http Post "$BaseUrl/edit.asp?id=$newId" @{
    id         = "$newId"
    company    = 'Pagination Corp International'
    role       = 'Senior Software Engineer'
    status     = 'Offer'
    applied_on = '2025-09-10'
    csrf       = $csrf
}
Report 'Edit returns 302 with pagination enabled' ($ru.Code -eq 302) "code=$($ru.Code)"
Report 'The edit was saved' ((DbRow $newId 'company') -eq 'Pagination Corp International') "got=$(DbRow $newId 'company')"
$r = Get-Page (Build-Url '/default.asp' @{ q = 'Pagination Corp International' })
Report 'Search finds the renamed record' ((Get-ResultTotal $r.Body) -eq 1) "shown=$((Get-ResultTotal $r.Body))"

$csrf = Get-Csrf
$rd = Http Post "$BaseUrl/delete.asp" @{ id = "$newId"; csrf = $csrf }
Report 'Delete returns 302 with pagination enabled' ($rd.Code -eq 302) "code=$($rd.Code)"
Report 'The record is gone again' ((DbCount) -eq $rowsBeforeCrud) "rows=$(DbCount) expected=$rowsBeforeCrud"
$r = Get-Page (Build-Url '/default.asp' @{ q = 'Pagination Corp' })
Report 'The deleted record no longer appears in search' ((Get-ResultTotal $r.Body) -eq 0) "shown=$((Get-ResultTotal $r.Body))"

} finally {
    if ($proc) {
        Write-Host "`nStopping the test IIS Express instance..." -ForegroundColor DarkGray
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
    $env:PT_DB_NAME = ''
}

# ===========================================================================
Write-Host ''
Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host (" Results: {0} passed, {1} failed" -f $script:Passed, $script:Failed) `
    -ForegroundColor $(if ($script:Failed -eq 0) { 'Green' } else { 'Red' })
Write-Host '==========================================================' -ForegroundColor Cyan
$script:Results | Format-Table -AutoSize

if ($script:Failed -gt 0) { exit 1 } else { exit 0 }
