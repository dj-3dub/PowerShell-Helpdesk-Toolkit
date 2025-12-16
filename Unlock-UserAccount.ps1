<#
.SYNOPSIS
    Unlock an Active Directory user account (safe confirmation).

.DESCRIPTION
    Attempts to unlock an AD user account if the AD module is available.
    Includes confirmation prompt and writes a log entry.

.PARAMETER Identity
    samAccountName or UPN.

.PARAMETER TicketId
    Optional ticket identifier used in log naming.

.EXAMPLE
    .\Unlock-UserAccount.ps1 -Identity jdoe

.EXAMPLE
    .\Unlock-UserAccount.ps1 -Identity jdoe@contoso.com -TicketId INC12345
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Identity,

    [string]$TicketId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$logsRoot  = Join-Path -Path (Split-Path $scriptDir -Parent) -ChildPath 'out\HelpdeskLogs'

if (-not (Test-Path $logsRoot)) {
    New-Item -Path $logsRoot -ItemType Directory -Force | Out-Null
}

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

$base = @("UnlockAD", $hostname, $timestamp)
if ($TicketId) { $base += $TicketId }
$logPath = Join-Path $logsRoot (($base -join '_') + ".txt")

function LogLine {
    param([string]$Text)
    if ($Text -eq $null) { return }
    $Text | Out-File -FilePath $logPath -Encoding UTF8 -Append
    Write-Host $Text
}

function Section {
    param([string]$Title)
    $line = ('=' * 70)
    LogLine ""
    LogLine $line
    LogLine ("== {0}" -f $Title)
    LogLine $line
}

Section "Unlock AD account"
LogLine ("Identity: {0}" -f $Identity)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Log: {0}" -f $logPath)

try {
    Get-Command Get-ADUser -ErrorAction Stop | Out-Null
    Get-Command Unlock-ADAccount -ErrorAction Stop | Out-Null
} catch {
    LogLine "AD module not available (RSAT required). Cannot unlock."
    LogLine "Done."
    return
}

try {
    $u = Get-ADUser -Identity $Identity -Properties LockedOut,Enabled -ErrorAction Stop
    LogLine ("Found: {0}, Enabled: {1}, LockedOut: {2}" -f $u.SamAccountName, $u.Enabled, $u.LockedOut)

    if (-not $u.Enabled) {
        LogLine "Account is disabled. Unlock will not help."
        return
    }

    if (-not $u.LockedOut) {
        LogLine "Account is not locked out. No action taken."
        return
    }

    $confirm = Read-Host "Unlock this account now? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
        LogLine "Cancelled by operator."
        return
    }

    if ($PSCmdlet.ShouldProcess($Identity, "Unlock-ADAccount")) {
        Unlock-ADAccount -Identity $Identity -ErrorAction Stop
        LogLine "Unlock successful."
    }
} catch {
    LogLine ("Unlock failed: {0}" -f $_.Exception.Message)
}

Section "Complete"
LogLine "Done."
