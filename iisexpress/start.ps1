<#
.SYNOPSIS
  Starts the Placement Tracker site on IIS Express at http://localhost:8080/.

.DESCRIPTION
  Full IIS (W3SVC) is not required. This script:
    1. Copies the machine's default IIS Express applicationhost.config to
       .\iisexpress\applicationhost.config the first time it runs.
    2. Adds an application pool ("No Managed Code", integrated pipeline) and
       a site named PlacementTracker bound to port 8080.
    3. Launches iisexpress.exe with that config.

  Run it from the project root:

      powershell -ExecutionPolicy Bypass -File .\iisexpress\start.ps1

  Stop the server with Ctrl+C in the console window.

.PARAMETER Port
  TCP port to bind. Defaults to 8080.

.PARAMETER DbName
  Optional database name override (written to the process environment as
  PT_DB_NAME). inc\db.asp reads that variable, which is how the smoke tests
  point a second instance at the isolated PlacementTracker_Test database.

.PARAMETER Config
  Path of the applicationhost.config to use/create.

.PARAMETER GenerateOnly
  Create/refresh the config file and exit without launching IIS Express.
  tests\smoke.ps1 uses this so it can build its own copy on a different port.
#>
[CmdletBinding()]
param(
    [int]    $Port    = 8080,
    [string] $DbName  = '',
    [string] $Config  = '',
    [switch] $GenerateOnly
)

$ErrorActionPreference = 'Stop'

# --- locate the project root (parent of this script) -------------------------
$ProjectRoot = Split-Path -Parent $PSScriptRoot
if (-not $ProjectRoot) { $ProjectRoot = (Get-Location).Path }

$IisExpressExe = 'C:\Program Files\IIS Express\iisexpress.exe'
if (-not (Test-Path $IisExpressExe)) {
    $IisExpressExe = 'C:\Program Files (x86)\IIS Express\iisexpress.exe'
}
if (-not (Test-Path $IisExpressExe)) {
    throw 'IIS Express was not found. Install "IIS Express" or open the project in full IIS instead (see README.md).'
}

if (-not $Config) { $Config = Join-Path $PSScriptRoot 'applicationhost.config' }

# --- 1. seed the config from the machine default if needed -------------------
if (-not (Test-Path $Config)) {
    $DefaultConfig = Join-Path $env:USERPROFILE 'Documents\IISExpress\config\applicationhost.config'
    if (-not (Test-Path $DefaultConfig)) {
        $DefaultConfig = 'C:\Program Files\IISExpress\config\applicationhost.config'
    }
    if (-not (Test-Path $DefaultConfig)) {
        throw "Cannot find a default applicationhost.config (looked in $DefaultConfig)."
    }
    Copy-Item $DefaultConfig $Config -Force
    Write-Host "Created $Config from $DefaultConfig"
}

# --- 2. patch in the pool + site ---------------------------------------------
$xmlText = Get-Content -LiteralPath $Config -Raw -Encoding UTF8
[xml]$xml = $xmlText

# applicationHost.config has no XML namespace - work with plain node names.
$siteName = 'PlacementTracker'
$poolName = 'PlacementTrackerPool'
$physical = $ProjectRoot

# --- application pool ---------------------------------------------------------
$appPools = $xml.configuration.'system.applicationHost'.applicationPools
$pool = $appPools.add | Where-Object { $_.name -eq $poolName }
if (-not $pool) {
    $pool = $xml.CreateElement('add')
    $pool.SetAttribute('name', $poolName)
    $pool.SetAttribute('autoStart', 'true')
    $pool.SetAttribute('managedRuntimeVersion', '')      # No Managed Code
    $pool.SetAttribute('managedPipelineMode', 'Integrated')
    $pool.SetAttribute('startMode', 'AlwaysRunning')
    [void]$appPools.AppendChild($pool)
    Write-Host "Added application pool $poolName"
}

# --- site ---------------------------------------------------------------------
$sitesParent = $xml.configuration.'system.applicationHost'.sites
$site = $sitesParent.site | Where-Object { $_.name -eq $siteName }
if ($site) { [void]$sitesParent.RemoveChild($site) }

$site = $xml.CreateElement('site')
$site.SetAttribute('name', $siteName)
$site.SetAttribute('id', '99')
$site.SetAttribute('serverAutoStart', 'true')

$app = $xml.CreateElement('application')
$app.SetAttribute('path', '/')
$app.SetAttribute('applicationPool', $poolName)

$vdir = $xml.CreateElement('virtualDirectory')
$vdir.SetAttribute('path', '/')
$vdir.SetAttribute('physicalPath', $physical)
[void]$app.AppendChild($vdir)
[void]$site.AppendChild($app)

$bindings = $xml.CreateElement('bindings')
$binding  = $xml.CreateElement('binding')
$binding.SetAttribute('protocol', 'http')
# ip:port:host.  Binding to the specific host "localhost" needs no HTTP.sys
# URL reservation, so the site can be started from an unelevated prompt.
# Use "*:$Port:" instead if you run IIS Express as an administrator and want
# the site reachable from other machines.
$binding.SetAttribute('bindingInformation', ":${Port}:localhost")
[void]$bindings.AppendChild($binding)
[void]$site.AppendChild($bindings)

[void]$sitesParent.AppendChild($site)
Write-Host "Site $siteName -> $physical on http://localhost:$Port/"

# --- make sure Classic ASP is switched on -------------------------------------
$asp = $xml.configuration.'system.webServer'.asp
if (-not $asp) {
    $systemWebServer = $xml.configuration.'system.webServer'
    $asp = $xml.CreateElement('asp')
    [void]$systemWebServer.AppendChild($asp)
}
$asp.SetAttribute('scriptErrorSentToBrowser', 'true')
$asp.SetAttribute('enableParentPaths', 'true')

# allow web.config to override <asp> (the default declaration is Deny)
$sections = $xml.configuration.configSections.sectionGroup | Where-Object { $_.name -eq 'system.webServer' }
if ($sections) {
    $aspSection = $sections.section | Where-Object { $_.name -eq 'asp' }
    if ($aspSection) { $aspSection.SetAttribute('overrideModeDefault', 'Allow') }
}

$xml.Save($Config) | Out-Null
Write-Host "Saved $Config"

if ($GenerateOnly) {
    Write-Host "Config generated; not launching IIS Express (-GenerateOnly)."
    return
}

# --- 3. launch ----------------------------------------------------------------
$env:PT_DB_SERVER   = '.\SQLEXPRESS'
if ($DbName) { $env:PT_DB_NAME = $DbName }

$iisArgs = @(
    "/config:`"$Config`"",
    "/site:$siteName"
)

Write-Host ""
Write-Host "Starting IIS Express: http://localhost:$Port/" -ForegroundColor Green
if ($DbName) { Write-Host "Database override active: $DbName" -ForegroundColor Yellow }
Write-Host "Press Ctrl+C to stop." -ForegroundColor DarkGray
Write-Host ""

& $IisExpressExe @iisArgs
