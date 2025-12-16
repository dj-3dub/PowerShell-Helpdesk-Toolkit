<#
.SYNOPSIS
    Runs pre-flight endpoint health checks before deeper troubleshooting.

.DESCRIPTION
    Collects a quick health snapshot of the local machine:

      - OS version and uptime
      - System drive disk space
      - Memory usage summary
      - Status of key services:
          * Workstation (lanmanworkstation)
          * Netlogon (netlogon)
          * DNS Client (dnscache)
          * Network Location Awareness (nlasvc)
          * Group Policy Client (gpsvc)
          * Windows Time (w32time)
      - Basic network configuration (IP, gateway, DNS servers)
      - Optional connectivity test to a configurable host
      - Event log summary for recent Errors/Warnings in System & Application

    Optionally writes results to out\PreflightChecks and compresses into a ZIP.

.PARAMETER HoursForEvents
    How many hours back to search for event log errors/warnings.
    Default: 4 hours.

.PARAMETER TestHost
    Optional host to test connectivity against (ping).
    Example: domain controller, intranet site, or known internal IP.

.PARAMETER CollectLogs
    If specified, writes collected data to a folder under out\PreflightChecks
    in the repo root and creates a ZIP.

.EXAMPLE
    .\Test-EndpointPreflight.ps1 -Verbose

.EXAMPLE
    .\Test-EndpointPreflight.ps1 -HoursForEvents 8 -TestHost 'dc01.corp.local' -CollectLogs -Verbose
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$HoursForEvents = 4,

    [Parameter(Mandatory = $false)]
    [string]$TestHost,

    [Parameter(Mandatory = $false)]
    [switch]$CollectLogs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Verbose "Starting endpoint pre-flight checks (HoursForEvents=$HoursForEvents, TestHost='$TestHost', CollectLogs=$CollectLogs)"

# --- repo root + output folder (WSL/UNC-safe string-based) ---
$repoRoot    = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$outputRoot  = Join-Path $repoRoot 'out\PreflightChecks'
$timestamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$sessionName = "Preflight_{0}_{1}" -f $env:COMPUTERNAME, $timestamp
$sessionRoot = Join-Path $outputRoot $sessionName

if ($CollectLogs) {
    $null = New-Item -Path $outputRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
    $null = New-Item -Path $sessionRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
    Write-Verbose "Pre-flight session folder: $sessionRoot"
}

function Write-DiagFile {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][string]$Content
    )
    if (-not $CollectLogs) { return }

    $targetPath = Join-Path $sessionRoot $RelativePath
    $targetDir  = Split-Path -Path $targetPath -Parent
    $null = New-Item -Path $targetDir -ItemType Directory -Force -ErrorAction SilentlyContinue

    $Content | Out-File -FilePath $targetPath -Encoding UTF8 -Force
}

# --- 1. OS & uptime ---
Write-Host "=== OS & Uptime ===" -ForegroundColor Cyan
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$lastBoot = $os.LastBootUpTime
$uptime = (Get-Date) - $lastBoot

$osInfo = [pscustomobject]@{
    ComputerName = $env:COMPUTERNAME
    OSCaption    = $os.Caption
    OSVersion    = $os.Version
    LastBoot     = $lastBoot
    UptimeDays   = [math]::Round($uptime.TotalDays, 2)
}

$osText = $osInfo | Format-List | Out-String
Write-Host $osText
Write-DiagFile -RelativePath 'system\OS_Uptime.txt' -Content $osText

# --- 2. Disk space (system drive) ---
Write-Host "=== Disk Space (System Drive) ===" -ForegroundColor Cyan
$sysDriveLetter = $env:SystemDrive.TrimEnd('\')
$disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$sysDriveLetter'"
$diskInfo = [pscustomobject]@{
    Drive        = $disk.DeviceID
    SizeGB       = [math]::Round($disk.Size / 1GB, 2)
    FreeGB       = [math]::Round($disk.FreeSpace / 1GB, 2)
    FreePercent  = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
}
$diskText = $diskInfo | Format-List | Out-String
Write-Host $diskText
Write-DiagFile -RelativePath 'system\Disk_SystemDrive.txt' -Content $diskText

# --- 3. Memory summary ---
Write-Host "=== Memory ===" -ForegroundColor Cyan
$totalMB = [math]::Round($os.TotalVisibleMemorySize / 1024, 2)
$freeMB  = [math]::Round($os.FreePhysicalMemory / 1024, 2)
$usedMB  = $totalMB - $freeMB
$memInfo = [pscustomobject]@{
    TotalMB      = $totalMB
    UsedMB       = $usedMB
    FreeMB       = $freeMB
    UsedPercent  = if ($totalMB -gt 0) { [math]::Round(($usedMB / $totalMB) * 100, 2) } else { $null }
}
$memText = $memInfo | Format-List | Out-String
Write-Host $memText
Write-DiagFile -RelativePath 'system\Memory.txt' -Content $memText

# --- 4. Key services ---
Write-Host "=== Key Service Status ===" -ForegroundColor Cyan
$serviceNames = @(
    'lanmanworkstation',  # Workstation
    'netlogon',           # Netlogon (domain)
    'Dnscache',           # DNS Client
    'NlaSvc',             # Network Location Awareness
    'gpsvc',              # Group Policy Client
    'w32time'             # Windows Time
)

$services = Get-Service -Name $serviceNames -ErrorAction SilentlyContinue
$svcInfo = $services | Select-Object Name, DisplayName, Status, StartType
$svcText = $svcInfo | Format-Table -AutoSize | Out-String
Write-Host $svcText
Write-DiagFile -RelativePath 'system\Services_Key.txt' -Content $svcText

# --- 5. Network summary ---
Write-Host "=== Network Configuration ===" -ForegroundColor Cyan
try {
    $netConfig = Get-NetIPConfiguration -ErrorAction SilentlyContinue |
        Where-Object { $_.IPv4Address -and $_.NetAdapter.Status -eq 'Up' }

    if ($netConfig) {
        $netSummary = $netConfig | Select-Object `
            @{n='InterfaceAlias';e={$_.InterfaceAlias}},
            @{n='IPv4';e={$_.IPv4Address.IPAddress}},
            @{n='IPv4DefaultGateway';e={$_.IPv4DefaultGateway.NextHop}},
            @{n='DNSServers';e={$_.DNSServer.ServerAddresses -join ', '}}
        $netText = $netSummary | Format-Table -AutoSize | Out-String
        Write-Host $netText
        Write-DiagFile -RelativePath 'network\NetIPConfiguration.txt' -Content $netText
    }
    else {
        Write-Host "No active network interfaces with IPv4 found." -ForegroundColor Yellow
    }
}
catch {
    Write-Warning "Get-NetIPConfiguration failed: $($_.Exception.Message)"
}

# --- 6. Connectivity test (optional) ---
if ($TestHost) {
    Write-Host "=== Connectivity Test: $TestHost ===" -ForegroundColor Cyan
    $connResult = $null
    try {
        $connResult = Test-NetConnection -ComputerName $TestHost -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Test-NetConnection failed for $TestHost: $($_.Exception.Message)"
    }

    if ($connResult) {
        $connInfo = [pscustomobject]@{
            Target        = $TestHost
            PingSucceeded = $connResult.PingSucceeded
            RemoteAddress = $connResult.RemoteAddress
            RoundTripMs   = $connResult.PingReplyDetails.RoundtripTime
        }
        $connText = $connInfo | Format-List | Out-String
        Write-Host $connText
        Write-DiagFile -RelativePath 'network\Connectivity_TestHost.txt' -Content $connText
    }
    else {
        Write-Host "No result from Test-NetConnection to $TestHost." -ForegroundColor Yellow
    }
}

# --- 7. Event log summary ---
Write-Host "=== Event Log Summary (last $HoursForEvents hours) ===" -ForegroundColor Cyan
$startTime = (Get-Date).AddHours(-1 * $HoursForEvents)

$logNames = @('System','Application')

$eventsSummary = @()
foreach ($log in $logNames) {
    try {
        $filter = @{
            LogName   = $log
            Level     = @(1,2,3)  # Critical, Error, Warning
            StartTime = $startTime
        }

        $events = Get-WinEvent -FilterHashtable $filter -ErrorAction SilentlyContinue

        $grouped = $events | Group-Object -Property LevelDisplayName | Select-Object Name, Count

        $summary = [pscustomobject]@{
            LogName = $log
            CriticalCount = ($grouped | Where-Object Name -eq 'Critical').Count
            ErrorCount    = ($grouped | Where-Object Name -eq 'Error').Count
            WarningCount  = ($grouped | Where-Object Name -eq 'Warning').Count
        }

        $eventsSummary += $summary
    }
    catch {
        Write-Warning "Failed to query $log events: $($_.Exception.Message)"
    }
}

if ($eventsSummary.Count -gt 0) {
    $evText = $eventsSummary | Format-Table -AutoSize | Out-String
    Write-Host $evText
    Write-DiagFile -RelativePath 'events\EventSummary.txt' -Content $evText
}
else {
    Write-Host "No recent events or failed to query event logs." -ForegroundColor Yellow
}

# Optionally dump latest 50 error events for System & Application
foreach ($log in $logNames) {
    try {
        $filter = @{
            LogName   = $log
            Level     = @(1,2)  # Critical + Error
            StartTime = $startTime
        }
        $latestErrors = Get-WinEvent -FilterHashtable $filter -MaxEvents 50 -ErrorAction SilentlyContinue
        if ($latestErrors -and $CollectLogs) {
            $errText = $latestErrors | Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message |
                Format-List | Out-String
            Write-DiagFile -RelativePath ("events\{0}_Errors.txt" -f $log) -Content $errText
        }
    }
    catch { }
}

# --- 8. Domain join / DC info ---
Write-Host "=== Domain Join Status ===" -ForegroundColor Cyan
try {
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    $domainInfo = [pscustomobject]@{
        ComputerName = $cs.Name
        PartOfDomain = $cs.PartOfDomain
        Domain       = if ($cs.PartOfDomain) { $cs.Domain } else { $null }
    }
    $domainText = $domainInfo | Format-List | Out-String
    Write-Host $domainText
    Write-DiagFile -RelativePath 'system\DomainStatus.txt' -Content $domainText
}
catch {
    Write-Warning "Failed to query domain status: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Pre-flight checks complete." -ForegroundColor Green

if ($CollectLogs) {
    $zipName = "Preflight_{0}_{1}.zip" -f $env:COMPUTERNAME, $timestamp
    $zipPath = Join-Path $outputRoot $zipName

    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    }

    Compress-Archive -Path (Join-Path $sessionRoot '*') -DestinationPath $zipPath -ErrorAction Stop
    Write-Host "Pre-flight diagnostics bundled: $zipPath" -ForegroundColor Green
}
