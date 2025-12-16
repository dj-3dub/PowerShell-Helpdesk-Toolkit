<#
.SYNOPSIS
    Force an Active Directory password reset (with confirmation).

.DESCRIPTION
    Resets the user's password in AD and (optionally) sets ChangePasswordAtLogon.
    This is an impactful action - script includes confirmation prompts.

.PARAMETER Identity
    samAccountName or UPN.

.PARAMETER TicketId
    Optional ticket identifier used in log naming.

.EXAMPLE
    .\Force-PasswordReset.ps1 -Identity jdoe

.EXAMPLE
    .\Force-PasswordReset.ps1 -Identity jdoe@contoso.com -TicketId INC12345
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

$base = @("ResetPwdAD", $hostname, $timestamp)
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

Section "Force AD password reset"
LogLine ("Identity: {0}" -f $Identity)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Log: {0}" -f $logPath)

try {
    Get-Command Get-ADUser -ErrorAction Stop | Out-Null
    Get-Command Set-ADAccountPassword -ErrorAction Stop | Out-Null
    Get-Command Set-ADUser -ErrorAction Stop | Out-Null
} catch {
    LogLine "AD module not available (RSAT required). Cannot reset password."
    LogLine "Done."
    return
}

try {
    $u = Get-ADUser -Identity $Identity -Properties Enabled -ErrorAction Stop
    LogLine ("Found: {0}, Enabled: {1}" -f $u.SamAccountName, $u.Enabled)

    if (-not $u.Enabled) {
        LogLine "Account is disabled. Resetting password may not help."
    }

    LogLine ""
    LogLine "You will be prompted for a temporary password (secure)."
    LogLine "Best practice: require change at next logon."

    $confirm = Read-Host "Proceed with password reset for this user? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
        LogLine "Cancelled by operator."
        return
    }

    $tempPwd = Read-Host -AsSecureString "Enter temporary password"

    $setChange = Read-Host "Require change at next logon? (Y/N, default Y)"
    $requireChange = -not ($setChange -match '^[Nn]')

    if ($PSCmdlet.ShouldProcess($Identity, "Set-ADAccountPassword -Reset")) {
        Set-ADAccountPassword -Identity $Identity -Reset -NewPassword $tempPwd -ErrorAction Stop
        LogLine "Password reset successful."
    }

    if ($requireChange) {
        Set-ADUser -Identity $Identity -ChangePasswordAtLogon $true -ErrorAction Stop
        LogLine "Set ChangePasswordAtLogon = true"
    }

} catch {
    LogLine ("Password reset failed: {0}" -f $_.Exception.Message)
}

Section "Complete"
LogLine "Done."
