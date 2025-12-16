<#
.SYNOPSIS
    Collects common helpdesk logs and diagnostics into a ZIP file.

.DESCRIPTION
    Gathers:
      - System & Application event logs (last N hours)
      - System info
      - Network configuration and basic connectivity tests
      - Installed programs list
      - Microsoft Teams logs (if present)

    Outputs a ZIP file under out\HelpdeskLogs in the repo root.

.PARAMETER Hours
    Lookback window for event logs. Default: 24.

.PARAMETER TicketId
    Optional helpdesk ticket ID to include in the ZIP name.

.PARAMETER OutputRoot
    Optional override for output root folder. Default is repoRoot\out\HelpdeskLogs.

.EXAMPLE
    .\Collect-HelpdeskLogs.ps1 -Verbose
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$Hours = 24,

    [Parameter(Mandatory = $false)]
    [string]$TicketId,

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# repo root (normalize to plain filesystem path, no provider prefix)
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).ProviderPath

if (-not $OutputRoot) {
    $OutputRoot = Join-Path $repoRoot 'out\HelpdeskLogs'
}

$null = New-Item -Path $OutputRoot -ItemType Directory -Force -ErrorAction SilentlyContinue

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$sessionName = "HelpdeskLogs_{0}_{1}" -f $env:COMPUTERNAME, $timestamp
$sessionRoot = Join-Path $OutputRoot $sessionName

Write-Verbose "Session folder: $sessionRoot"
$null = New-Item -Path $sessionRoot -ItemType Directory -Force -ErrorAction SilentlyContinue

$systemDir  = Join-Path $sessionRoot 'system'
$networkDir = Join-Path $sessionRoot 'network'
$appsDir    = Join-Path $sessionRoot 'apps'

$null = New-Item -Path $systemDir  -ItemType Directory -Force
$null = New-Item -Path $networkDir -ItemType Directory -Force
$null = New-Item -Path $appsDir    -ItemType Directory -Force

# --- System info ---
Write-Verbose "Collecting system information..."
systeminfo | Out-File -FilePath (Join-Path $systemDir 'systeminfo.txt') -Encoding UTF8

$since = (Get-Date).AddHours(-[math]::Abs($Hours))
Write-Verbose "Exporting event logs since $since..."

try {
    Get-EventLog -LogName System -After $since -ErrorAction Stop |
        Export-Clixml -Path (Join-Path $systemDir 'SystemEvents.xml')
}
catch {
    Write-Warning "Failed to export System event log: $($_.Exception.Message)"
}

try {
    Get-EventLog -LogName Application -After $since -ErrorAction Stop |
        Export-Clixml -Path (Join-Path $systemDir 'ApplicationEvents.xml')
}
catch {
    Write-Warning "Failed to export Application event log: $($_.Exception.Message)"
}

# --- Network info ---
Write-Verbose "Collecting network diagnostics..."

ipconfig /all | Out-File -FilePath (Join-Path $networkDir 'ipconfig_all.txt') -Encoding UTF8

try {
    ping 8.8.8.8 -n 4 | Out-File -FilePath (Join-Path $networkDir 'ping_8.8.8.8.txt')
}
catch {
    Write-Warning "Ping test failed: $($_.Exception.Message)"
}

# --- Installed applications ---
Write-Verbose "Collecting installed programs list..."

$installed = @()

$paths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

foreach ($path in $paths) {
    try {
        $installed += Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Failed to read ${path}: $($_.Exception.Message)"
    }
}

$installed |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
    Where-Object { $_.DisplayName } |
    Sort-Object DisplayName |
    Export-Csv -Path (Join-Path $appsDir 'InstalledPrograms.csv') -NoTypeInformation -Encoding UTF8

# --- Teams logs ---
$teamsLog = Join-Path $env:APPDATA 'Microsoft\Teams\logs.txt'
if (Test-Path $teamsLog) {
    Copy-Item $teamsLog -Destination (Join-Path $appsDir 'Teams_logs.txt') -ErrorAction SilentlyContinue
}

# --- Create ZIP ---
$zipName = if ($TicketId) {
    "HelpdeskLogs_{0}_{1}_{2}.zip" -f $TicketId, $env:COMPUTERNAME, $timestamp
}
else {
    "HelpdeskLogs_{0}_{1}.zip" -f $env:COMPUTERNAME, $timestamp
}

$zipPath = Join-Path $OutputRoot $zipName

Write-Verbose "Creating ZIP archive: $zipPath"

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Compress-Archive -Path (Join-Path $sessionRoot '*') -DestinationPath $zipPath -ErrorAction Stop

Write-Host "Helpdesk logs collected: $zipPath" -ForegroundColor Green
