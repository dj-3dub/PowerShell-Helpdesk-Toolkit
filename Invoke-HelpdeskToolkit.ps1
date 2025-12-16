<#
.SYNOPSIS
    Interactive Helpdesk Automation Toolkit launcher.

.DESCRIPTION
    Provides a simple menu to run common helpdesk automations.

    Notes:
      - Run in Windows PowerShell (or PowerShell 7) for Windows-only tasks.
      - Some tasks (Security log reads, repair operations) may require "Run as Administrator".
      - Logs are written under: out\HelpdeskLogs

.EXAMPLE
    .\Invoke-HelpdeskToolkit.ps1
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Resolve the folder that contains the helpdesk scripts
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

function Get-LogsDir {
    $logsDir = Join-Path $scriptDir '..\..\out\HelpdeskLogs'
    try { return (Resolve-Path $logsDir).Path } catch { return $logsDir }
}

function Invoke-ResetPrinter {
    $script = Join-Path $scriptDir 'Reset-PrinterSubsystem.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }
    & $script -Verbose
}

function Invoke-ResetTeams {
    $script = Join-Path $scriptDir 'Reset-TeamsCache.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $restart = Read-Host "Restart Teams after clearing cache? (Y/N)"
    if ($restart -match '^[Yy]') {
        & $script -Restart -Verbose
    } else {
        & $script -Verbose
    }
}

function Invoke-ResetOutlook {
    $script = Join-Path $scriptDir 'Reset-OutlookProfile.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $confirm = Read-Host "This will reset the Outlook profile for the current user. Continue? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
    & $script -Verbose
}

function Invoke-RepairNetwork {
    $script = Join-Path $scriptDir 'Repair-NetworkStack.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }
    & $script -Verbose
}

function Invoke-RepairWindowsUpdate {
    $script = Join-Path $scriptDir 'Repair-WindowsUpdate.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }
    & $script -Verbose
}

function Invoke-CollectHelpdeskLogs {
    $script = Join-Path $scriptDir 'Collect-HelpdeskLogs.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Enter Ticket ID (optional, press Enter to skip)"
    $hoursInput = Read-Host "How many hours of event logs to collect? (default 24)"
    $hours = 24
    if ($hoursInput -match '^\d+$') { $hours = [int]$hoursInput }

    if ([string]::IsNullOrWhiteSpace($ticket)) {
        & $script -Hours $hours -Verbose
    } else {
        & $script -TicketId $ticket -Hours $hours -Verbose
    }
}

function Invoke-ResetBrowser {
    $script = Join-Path $scriptDir 'Reset-BrowserProfile.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $browser = Read-Host "Browser to reset? (Chrome/Edge/All, default All)"
    if ([string]::IsNullOrWhiteSpace($browser)) { $browser = 'All' }

    $mode = Read-Host "Mode? (C = Clear cache only, R = Full reset, default C)"
    $clearCacheOnly = $true
    if ($mode -match '^[Rr]') { $clearCacheOnly = $false }

    $backupAns = Read-Host "Backup profile(s) before changes? (Y/N, default Y)"
    $backupSwitch = $true
    if ($backupAns -match '^[Nn]') { $backupSwitch = $false }

    $restartAns = Read-Host "Restart browser after reset? (Y/N, default N)"
    $restartSwitch = $false
    if ($restartAns -match '^[Yy]') { $restartSwitch = $true }

    $params = @{
        Browser = $browser
        Verbose = $true
    }
    if ($clearCacheOnly) { $params.ClearCacheOnly = $true }
    if ($backupSwitch)   { $params.Backup        = $true }
    if ($restartSwitch)  { $params.Restart       = $true }

    & $script @params
}

function Invoke-MailboxCapacityCheck {
    $script = Join-Path $scriptDir 'Test-MailboxCapacityAndRetention.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    Write-Host "NOTE: You must be connected to Exchange (Exchange Online or EMS) before running this option." -ForegroundColor Yellow
    $scope = Read-Host "Check Single mailbox (S) or All mailboxes (A)? (default S)"
    if ([string]::IsNullOrWhiteSpace($scope)) { $scope = 'S' }

    $thresholdInput = Read-Host "Warning threshold percent for 'NearCapacity' flag? (default 90)"
    $threshold = 90
    if ($thresholdInput -match '^\d+$') { $threshold = [int]$thresholdInput }

    $exportAns = Read-Host "Export results to CSV? (Y/N, default Y)"
    $exportSwitch = $true
    if ($exportAns -match '^[Nn]') { $exportSwitch = $false }

    $params = @{
        WarningThresholdPercent = $threshold
        Verbose                 = $true
    }

    if ($scope.ToUpper() -eq 'A') {
        $params.AllMailboxes = $true
    } else {
        $id = Read-Host "Enter mailbox identity (UPN/alias)"
        if ([string]::IsNullOrWhiteSpace($id)) {
            Write-Host "No identity entered. Cancelling mailbox check." -ForegroundColor Yellow
            return
        }
        $params.Identity = $id
    }

    if ($exportSwitch) { $params.ExportCsv = $true }

    & $script @params
}

function Invoke-OneDriveRepair {
    $script = Join-Path $scriptDir 'Repair-OneDriveSync.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $modeInput = Read-Host "Repair mode: Soft (stop/start) or Reset (/reset)? (S/R, default R)"
    $mode = 'Reset'
    if ($modeInput -match '^[Ss]') { $mode = 'Soft' }

    $backupAns = Read-Host "Backup OneDrive logs before changes? (Y/N, default Y)"
    $backupSwitch = $true
    if ($backupAns -match '^[Nn]') { $backupSwitch = $false }

    $restartAns = Read-Host "Restart OneDrive after repair? (Y/N, default Y)"
    $restartSwitch = $true
    if ($restartAns -match '^[Nn]') { $restartSwitch = $false }

    $params = @{
        Mode    = $mode
        Verbose = $true
    }
    if ($backupSwitch)  { $params.BackupLogs = $true }
    if ($restartSwitch) { $params.Restart    = $true }

    & $script @params
}

function Invoke-VpnDiagnostics {
    $script = Join-Path $scriptDir 'Test-VpnDiagnostics.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $vpnName = Read-Host "VPN connection name to focus on (optional, press Enter to skip)"
    $internal = Read-Host "Internal/VPN-only hostname or IP to test (optional, press Enter to skip)"
    $public = Read-Host "Public test host (default 8.8.8.8, press Enter to accept)"
    if ([string]::IsNullOrWhiteSpace($public)) { $public = '8.8.8.8' }

    $logsAns = Read-Host "Collect logs and bundle diagnostics into ZIP? (Y/N, default Y)"
    $logsSwitch = $true
    if ($logsAns -match '^[Nn]') { $logsSwitch = $false }

    $params = @{
        PublicTestHost = $public
        Verbose        = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($vpnName)) { $params.VpnName = $vpnName }
    if (-not [string]::IsNullOrWhiteSpace($internal)) { $params.InternalTestHost = $internal }
    if ($logsSwitch) { $params.CollectLogs = $true }

    & $script @params
}

function Invoke-BitLockerHealth {
    $script = Join-Path $scriptDir 'Test-BitLockerHealth.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $exportAns = Read-Host "Export BitLocker report to CSV? (Y/N, default Y)"
    $exportSwitch = $true
    if ($exportAns -match '^[Nn]') { $exportSwitch = $false }

    $recoveryAns = Read-Host "Include recovery keys in CSV? (Y/N, default N) (Sensitive!)"
    $includeRecovery = $false
    if ($recoveryAns -match '^[Yy]') { $includeRecovery = $true }

    $params = @{ Verbose = $true }
    if ($exportSwitch)    { $params.ExportCsv = $true }
    if ($includeRecovery) { $params.IncludeRecoveryKeys = $true }

    & $script @params
}

function Invoke-OutlookOstRepair {
    $script = Join-Path $scriptDir 'Repair-OutlookOst.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $identity = Read-Host "Filter OSTs by identity (email/UPN substring)? (optional, press Enter to process all)"
    $modeInput = Read-Host "Mode: Rebuild (R), Scan (S), or Both (B)? (default R)"
    $mode = 'Rebuild'
    switch -Regex ($modeInput) {
        '^[Ss]' { $mode = 'Scan' }
        '^[Bb]' { $mode = 'Both' }
    }

    $backupAns = Read-Host "Backup OST files before changes? (Y/N, default Y)"
    $backupSwitch = $true
    if ($backupAns -match '^[Nn]') { $backupSwitch = $false }

    $params = @{ Mode = $mode; Verbose = $true }
    if (-not [string]::IsNullOrWhiteSpace($identity)) { $params.Identity = $identity }
    if ($backupSwitch) { $params.Backup = $true }

    & $script @params
}

function Invoke-PreflightChecks {
    $script = Join-Path $scriptDir 'Test-EndpointPreflight.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $hoursInput = Read-Host "How many hours of event logs to include? (default 4)"
    $hours = 4
    if ($hoursInput -match '^\d+$') { $hours = [int]$hoursInput }

    $testHost = Read-Host "Test connectivity to a specific host? (optional hostname/IP, press Enter to skip)"
    $logsAns = Read-Host "Collect logs and bundle pre-flight diagnostics into ZIP? (Y/N, default Y)"
    $logsSwitch = $true
    if ($logsAns -match '^[Nn]') { $logsSwitch = $false }

    $params = @{ HoursForEvents = $hours; Verbose = $true }
    if (-not [string]::IsNullOrWhiteSpace($testHost)) { $params.TestHost = $testHost }
    if ($logsSwitch) { $params.CollectLogs = $true }

    & $script @params
}

function Invoke-NetworkPerformance {
    $script = Join-Path $scriptDir 'Test-NetworkPerformance.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $remedAns = Read-Host "Run safe remediation (DNS flush) if suggested? (Y/N, default N)"

    $params = @{ Verbose = $true }
    if (-not [string]::IsNullOrWhiteSpace($ticket)) { $params.TicketId = $ticket }
    if ($remedAns -match '^[Yy]') { $params.EnableSafeRemediation = $true }

    & $script @params
}

function Invoke-ViewLogsByPrefix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string]$Title
    )

    $logsDir = Get-LogsDir
    if (-not (Test-Path $logsDir)) {
        Write-Host "Log directory not found: $logsDir" -ForegroundColor Yellow
        return
    }

    Write-Host ""
    Write-Host ("=== {0} ===" -f $Title) -ForegroundColor Cyan

    $files = @(
        Get-ChildItem -Path $logsDir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like ($Prefix + '*') -and $_.Extension -in @('.txt', '.zip') } |
            Sort-Object -Property LastWriteTime -Descending
    )

    if ($files.Count -lt 1) {
        Write-Host ("No logs found matching prefix: {0}" -f $Prefix) -ForegroundColor Yellow
        return
    }

    $max = [Math]::Min($files.Count, 20)
    for ($i = 0; $i -lt $max; $i++) {
        Write-Host ("[{0}] {1}  ({2})" -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime)
    }

    Write-Host ""
    $pick = Read-Host "Open which log? Enter 1-$max (or press Enter to cancel)"
    if ([string]::IsNullOrWhiteSpace($pick)) { return }
    if ($pick -notmatch '^\d+$') { Write-Host "Invalid selection." -ForegroundColor Yellow; return }

    $idx = [int]$pick - 1
    if ($idx -lt 0 -or $idx -ge $max) { Write-Host "Invalid selection." -ForegroundColor Yellow; return }

    $selected = $files[$idx].FullName
    if ($selected.EndsWith('.txt')) {
        try {
            Start-Process notepad.exe $selected | Out-Null
            Write-Host "Opened in Notepad." -ForegroundColor Green
        } catch {
            Get-Content -Path $selected -ErrorAction Stop | Select-Object -First 400 | ForEach-Object { $_ }
        }
    } else {
        try {
            $folder = Split-Path -Path $selected -Parent
            Start-Process explorer.exe $folder | Out-Null
            Write-Host "Opened containing folder in Explorer." -ForegroundColor Green
        } catch {
            Write-Host ("ZIP is at: {0}" -f $selected) -ForegroundColor Yellow
        }
    }
}

function Invoke-ViewNetworkPerformanceLogs {
    Invoke-ViewLogsByPrefix -Prefix 'NetPerf_' -Title 'Network Performance Logs'
}

function Invoke-ProxyDiagnostics {
    $script = Join-Path $scriptDir 'Test-ProxyConfiguration.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $remedAns = Read-Host "Run safe remediation (WinHTTP reset + flush DNS) if suggested? (Y/N, default N)"

    $params = @{ Verbose = $true }
    if (-not [string]::IsNullOrWhiteSpace($ticket)) { $params.TicketId = $ticket }
    if ($remedAns -match '^[Yy]') { $params.EnableSafeRemediation = $true }

    & $script @params
}

function Invoke-ViewProxyLogs {
    Invoke-ViewLogsByPrefix -Prefix 'ProxyDiag_' -Title 'Proxy Diagnostics Logs'
}

function Invoke-IdentitySignInHealth {
    $script = Join-Path $scriptDir 'Test-UserSignInHealth.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $user = Read-Host "Enter username or UPN (default current user)"
    if ([string]::IsNullOrWhiteSpace($user)) { $user = $env:USERNAME }

    $ticket = Read-Host "Ticket ID (optional)"

    $params = @{ Identity = $user; Verbose = $true }
    if (-not [string]::IsNullOrWhiteSpace($ticket)) { $params.TicketId = $ticket }

    & $script @params
}

function Invoke-UnlockAdAccount {
    $script = Join-Path $scriptDir 'Unlock-UserAccount.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $user = Read-Host "Enter AD username/UPN to unlock"
    if ([string]::IsNullOrWhiteSpace($user)) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    $ticket = Read-Host "Ticket ID (optional)"

    $params = @{ Identity = $user; Verbose = $true }
    if (-not [string]::IsNullOrWhiteSpace($ticket)) { $params.TicketId = $ticket }

    & $script @params
}

function Invoke-ForceAdPasswordReset {
    $script = Join-Path $scriptDir 'Force-PasswordReset.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $user = Read-Host "Enter AD username/UPN to reset password"
    if ([string]::IsNullOrWhiteSpace($user)) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    $ticket = Read-Host "Ticket ID (optional)"

    $params = @{ Identity = $user; Verbose = $true }
    if (-not [string]::IsNullOrWhiteSpace($ticket)) { $params.TicketId = $ticket }

    & $script @params
}

function Invoke-FindLockoutSource {
    $script = Join-Path $scriptDir 'Find-UserLockoutSource.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $user = Read-Host "Enter username or UPN to investigate lockout"
    if ([string]::IsNullOrWhiteSpace($user)) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    $hoursInput = Read-Host "Look back how many hours? (default 24)"
    $hours = 24
    if ($hoursInput -match '^\d+$') { $hours = [int]$hoursInput }

    $ticket = Read-Host "Ticket ID (optional)"

    $params = @{ Identity = $user; Hours = $hours; Verbose = $true }
    if (-not [string]::IsNullOrWhiteSpace($ticket)) { $params.TicketId = $ticket }

    & $script @params
}



function Invoke-WiFiDiagnostics {
    $script = Join-Path $scriptDir 'Collect-WiFiDiagnostics.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $wlanAns = Read-Host "Include Windows WLAN report (HTML)? (Y/N, default Y)"
    $includeWlan = $true
    if ($wlanAns -match '^[Nn]') { $includeWlan = $false }

    $remedAns = Read-Host "Run safe remediation (flush DNS, renew, restart Wi-Fi service)? (Y/N, default N)"
    $enableRemed = $false
    if ($remedAns -match '^[Yy]') { $enableRemed = $true }

    $params = @{ Verbose = $true }
    if ($ticket)      { $params.TicketId = $ticket }
    if ($includeWlan) { $params.IncludeWlanReport = $true }
    if ($enableRemed) { $params.EnableSafeRemediation = $true }

    & $script @params
}

function Invoke-LogonGpoDiagnostics {
    $script = Join-Path $scriptDir 'Test-LogonGpoPerformance.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $hoursInput = Read-Host "Look back how many hours? (default 24)"
    $hours = 24
    if ($hoursInput -match '^\d+$') { $hours = [int]$hoursInput }

    $htmlAns = Read-Host "Generate gpresult HTML report? (Y/N, default Y)"
    $includeHtml = $true
    if ($htmlAns -match '^[Nn]') { $includeHtml = $false }

    $zipAns = Read-Host "Create ZIP bundle of artifacts? (Y/N, default Y)"
    $createZip = $true
    if ($zipAns -match '^[Nn]') { $createZip = $false }

    $params = @{ Hours = $hours; Verbose = $true }
    if ($ticket)      { $params.TicketId = $ticket }
    if ($includeHtml) { $params.IncludeGpResultHtml = $true }
    if ($createZip)   { $params.CreateZipBundle = $true }

    & $script @params
}

function Invoke-ViewLogonGpoLogs {
    Invoke-ViewLogsByPrefix -Prefix 'LogonGpoDiag_' -Title 'Logon + GPO Diagnostics Logs'
}



function Invoke-EndpointDelta {
    $script = Join-Path $scriptDir 'Get-EndpointDelta.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $hoursInput = Read-Host "Look back how many hours? (default 24)"
    $hours = 24
    if ($hoursInput -match '^\d+$') { $hours = [int]$hoursInput }

    $zipAns = Read-Host "Create ZIP bundle of artifacts? (Y/N, default Y)"
    $createZip = $true
    if ($zipAns -match '^[Nn]') { $createZip = $false }

    $params = @{ Hours = $hours; Verbose = $true }
    if ($ticket)    { $params.TicketId = $ticket }
    if ($createZip) { $params.CreateZipBundle = $true }

    & $script @params
}

function Invoke-ViewEndpointDeltaLogs {
    Invoke-ViewLogsByPrefix -Prefix 'EndpointDelta_' -Title 'Endpoint Delta (What changed) Logs'
}

function Invoke-StartupDelta {
    $script = Join-Path $scriptDir 'Get-StartupDelta.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $hoursInput = Read-Host "Look back how many hours? (default 24)"
    $hours = 24
    if ($hoursInput -match '^\d+$') { $hours = [int]$hoursInput }

    $zipAns = Read-Host "Create ZIP bundle of artifacts? (Y/N, default Y)"
    $createZip = $true
    if ($zipAns -match '^[Nn]') { $createZip = $false }

    $params = @{ Hours = $hours; Verbose = $true }
    if ($ticket)    { $params.TicketId = $ticket }
    if ($createZip) { $params.CreateZipBundle = $true }

    & $script @params
}

function Invoke-ViewStartupDeltaLogs {
    Invoke-ViewLogsByPrefix -Prefix 'StartupDelta_' -Title 'Startup Delta Logs'
}


function Invoke-FileShareDiagnostics {
    $script = Join-Path $scriptDir 'Test-FileShareAndMappedDrives.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $servers = Read-Host "File server(s) to test (comma-separated, optional)"
    $shares  = Read-Host "Share path(s) to test (comma-separated UNC, optional)"

    $remedAns = Read-Host "Enable SAFE remediation prompts (net use cleanup / klist purge)? (Y/N, default N)"
    $zipAns   = Read-Host "Create ZIP bundle of artifacts? (Y/N, default Y)"

    $params = @{ Verbose = $true }
    if ($ticket) { $params.TicketId = $ticket }

    if ($servers) {
        $params.Servers = ($servers -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if ($shares) {
        $params.Shares = ($shares -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    if ($remedAns -match '^[Yy]') { $params.EnableSafeRemediation = $true }
    if ($zipAns -match '^[Nn]')   { } else { $params.CreateZipBundle = $true }

    & $script @params
}

function Invoke-ViewFileShareDiagLogs {
    Invoke-ViewLogsByPrefix -Prefix 'FileShareDiag_' -Title 'File Share Diagnostics Logs'
}


function Invoke-PrintDiagnostics {
    $script = Join-Path $scriptDir 'Test-PrintDiagnostics.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket  = Read-Host "Ticket ID (optional)"
    $printer = Read-Host "Filter by printer name (optional substring)"
    $targets = Read-Host "Test targets (printer IPs/print servers, comma-separated, optional)"
    $zipAns  = Read-Host "Create ZIP bundle of artifacts? (Y/N, default Y)"

    $params = @{ Verbose = $true }
    if ($ticket)  { $params.TicketId = $ticket }
    if ($printer) { $params.PrinterName = $printer }

    if ($targets) {
        $params.TestTargets = ($targets -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    if ($zipAns -match '^[Nn]') { } else { $params.CreateZipBundle = $true }

    & $script @params
}

function Invoke-ViewPrintDiagLogs {
    Invoke-ViewLogsByPrefix -Prefix 'PrintDiag_' -Title 'Print Diagnostics Logs'
}


function Invoke-OutlookAuthDiagnostics {
    $script = Join-Path $scriptDir 'Test-OutlookAuthPrompts.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $tenant = Read-Host "Tenant hint (optional, e.g. contoso.com)"
    $cmdkey = Read-Host "Include cmdkey /list output? (Y/N, default N) (Review before sharing!)"

    $params = @{ Verbose = $true }
    if ($ticket) { $params.TicketId = $ticket }
    if ($tenant) { $params.TenantHint = $tenant }
    if ($cmdkey -match '^[Yy]') { $params.IncludeCmdKey = $true }

    & $script @params
}

function Invoke-ViewOutlookAuthLogs {
    Invoke-ViewLogsByPrefix -Prefix 'OutlookAuth_' -Title 'Outlook Auth Prompt Logs'
}

function Invoke-InternalDnsDiagnostics {
    $script = Join-Path $scriptDir 'Test-InternalDnsResolution.ps1'
    if (-not (Test-Path $script)) { Write-Warning "Missing: $script"; return }

    $ticket = Read-Host "Ticket ID (optional)"
    $names  = Read-Host "Internal names to test (comma-separated, e.g. fileserver01,intranet.contoso.com)"
    $dns    = Read-Host "DNS servers to query directly (comma-separated, optional)"
    $remed  = Read-Host "Run safe remediation (flush DNS cache) if suggested? (Y/N, default N)"

    $params = @{ Verbose = $true }
    if ($ticket) { $params.TicketId = $ticket }

    if ($names) {
        $params.InternalNames = ($names -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    if ($dns) {
        $params.DnsServers = ($dns -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    if ($remed -match '^[Yy]') { $params.EnableSafeRemediation = $true }

    & $script @params
}

function Invoke-ViewInternalDnsLogs {
    Invoke-ViewLogsByPrefix -Prefix 'InternalDNS_' -Title 'Internal DNS Diagnostics Logs'
}


function Show-Menu {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "   Helpdesk Automation Toolkit (PowerShell)" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " 1)  Reset printer subsystem"
    Write-Host " 2)  Reset Microsoft Teams cache"
    Write-Host " 3)  Reset Outlook profile"
    Write-Host " 4)  Repair network stack (DNS/Winsock/IP)"
    Write-Host " 5)  Repair Windows Update components"
    Write-Host " 6)  Collect helpdesk logs (ZIP)"
    Write-Host " 7)  Reset browser profiles (Chrome/Edge)"
    Write-Host " 8)  Check mailbox capacity and retention"
    Write-Host " 9)  Repair OneDrive sync (Soft/Reset + logs)"
    Write-Host "10)  VPN diagnostics (connectivity + routing)"
    Write-Host "11)  BitLocker health check (local machine)"
    Write-Host "12)  Outlook OST repair (scan/rebuild)"
    Write-Host "13)  Endpoint pre-flight checks"
    Write-Host "14)  Network performance diagnostics (slow/intermittent)"
    Write-Host "15)  View Network Performance Logs"
    Write-Host "16)  Proxy diagnostics (WinINET/WinHTTP/PAC/WPAD)"
    Write-Host "17)  View Proxy Diagnostics Logs"
    Write-Host "18)  Identity: Sign-in health triage"
    Write-Host "19)  Identity: Unlock AD account"
    Write-Host "20)  Identity: Force AD password reset"
    Write-Host "21)  Identity: Find lockout source (best effort)"
    Write-Host "22)  Wi-Fi diagnostics bundle (slow/unstable wireless)"
    Write-Host "23)  Logon + GPO performance diagnostics"
    Write-Host "24)  View Logon + GPO Diagnostics Logs"
    Write-Host "25)  Endpoint delta (What changed?) report"
    Write-Host "26)  View Endpoint Delta Logs"
    Write-Host "27)  Startup delta (Scheduled Tasks / Startup entries / Services)"
    Write-Host "28)  View Startup Delta Logs"
    Write-Host "29)  File share / mapped drive diagnostics (SMB 445 + auth + mappings)"
    Write-Host "30)  View File Share Diagnostics Logs"
    Write-Host "31)  Print diagnostics (spooler/queue/ports/drivers/events)"
    Write-Host "32)  View Print Diagnostics Logs"
    Write-Host "33)  Outlook auth prompts diagnostics (M365 sign-in loops)"
    Write-Host "34)  View Outlook Auth Prompt Logs"
    Write-Host "35)  VPN connected but nothing resolves (Internal DNS diagnostics)"
    Write-Host "36)  View Internal DNS Diagnostics Logs"
    Write-Host " Q)  Quit"
    Write-Host ""
}

do {
    Show-Menu
    $choice = Read-Host "Select an option"

    switch ($choice.ToUpper()) {
        '1'  { Invoke-ResetPrinter               ; Pause }
        '2'  { Invoke-ResetTeams                 ; Pause }
        '3'  { Invoke-ResetOutlook               ; Pause }
        '4'  { Invoke-RepairNetwork              ; Pause }
        '5'  { Invoke-RepairWindowsUpdate        ; Pause }
        '6'  { Invoke-CollectHelpdeskLogs        ; Pause }
        '7'  { Invoke-ResetBrowser               ; Pause }
        '8'  { Invoke-MailboxCapacityCheck       ; Pause }
        '9'  { Invoke-OneDriveRepair             ; Pause }
        '10' { Invoke-VpnDiagnostics             ; Pause }
        '11' { Invoke-BitLockerHealth            ; Pause }
        '12' { Invoke-OutlookOstRepair           ; Pause }
        '13' { Invoke-PreflightChecks            ; Pause }
        '14' { Invoke-NetworkPerformance         ; Pause }
        '15' { Invoke-ViewNetworkPerformanceLogs ; Pause }
        '16' { Invoke-ProxyDiagnostics           ; Pause }
        '17' { Invoke-ViewProxyLogs              ; Pause }
        '18' { Invoke-IdentitySignInHealth       ; Pause }
        '19' { Invoke-UnlockAdAccount            ; Pause }
        '20' { Invoke-ForceAdPasswordReset       ; Pause }
        '21' { Invoke-FindLockoutSource          ; Pause }
        '22' { Invoke-WiFiDiagnostics              ; Pause }
        '23' { Invoke-LogonGpoDiagnostics          ; Pause }
        '24' { Invoke-ViewLogonGpoLogs             ; Pause }
        '25' { Invoke-EndpointDelta                  ; Pause }
        '26' { Invoke-ViewEndpointDeltaLogs          ; Pause }
        '27' { Invoke-StartupDelta                   ; Pause }
        '28' { Invoke-ViewStartupDeltaLogs           ; Pause }
        '29' { Invoke-FileShareDiagnostics        ; Pause }
        '30' { Invoke-ViewFileShareDiagLogs       ; Pause }
        '31' { Invoke-PrintDiagnostics           ; Pause }
        '32' { Invoke-ViewPrintDiagLogs           ; Pause }
        '33' { Invoke-OutlookAuthDiagnostics     ; Pause }
        '34' { Invoke-ViewOutlookAuthLogs          ; Pause }
        '35' { Invoke-InternalDnsDiagnostics       ; Pause }
        '36' { Invoke-ViewInternalDnsLogs          ; Pause }
        'Q'  { Write-Host "Exiting Helpdesk Toolkit." -ForegroundColor Green }
        default {
            Write-Host "Invalid selection. Choose 1-36 or Q." -ForegroundColor Yellow
            Pause
        }
    }

} while ($choice.ToUpper() -ne 'Q')
