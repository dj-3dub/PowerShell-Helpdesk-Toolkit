<#
.SYNOPSIS
    Identity sign-in health triage (AD first, optional Entra/Microsoft Graph if available).

.DESCRIPTION
    Performs quick checks that commonly cause login/sign-in problems:
      - AD: Enabled, LockedOut, PasswordExpired, PasswordLastSet, LastLogonDate
      - Optional Entra: basic user lookup and (if permissions allow) recent sign-in events

    Designed to be safe: read-only by default.

.PARAMETER Identity
    Username (samAccountName) or UPN/email.

.PARAMETER TicketId
    Optional ticket identifier used in log naming.

.EXAMPLE
    .\Test-UserSignInHealth.ps1 -Identity jdoe

.EXAMPLE
    .\Test-UserSignInHealth.ps1 -Identity jdoe@contoso.com -TicketId INC12345
#>

[CmdletBinding()]
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

$base = @("IdentityHealth", $hostname, $timestamp)
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

Section "Identity sign-in health triage"
LogLine ("Identity: {0}" -f $Identity)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Log: {0}" -f $logPath)

# AD checks (preferred)
Section "Active Directory checks (if RSAT/AD module available)"

$adAvailable = $false
try {
    $cmd = Get-Command Get-ADUser -ErrorAction Stop
    $adAvailable = $true
} catch {
    LogLine "AD module not available (Get-ADUser not found)."
    LogLine "Tip: Run from a machine with RSAT AD tools or from a domain-joined admin workstation."
}

if ($adAvailable) {
    try {
        $u = Get-ADUser -Identity $Identity -Properties Enabled,LockedOut,PasswordExpired,PasswordLastSet,LastLogonDate,AccountExpirationDate,PasswordNeverExpires,msDS-UserPasswordExpiryTimeComputed -ErrorAction Stop

        LogLine ("Found AD user: {0}" -f $u.SamAccountName)
        LogLine ("Enabled: {0}" -f $u.Enabled)
        LogLine ("LockedOut: {0}" -f $u.LockedOut)
        LogLine ("PasswordExpired: {0}" -f $u.PasswordExpired)
        LogLine ("PasswordNeverExpires: {0}" -f $u.PasswordNeverExpires)
        LogLine ("PasswordLastSet: {0}" -f $u.PasswordLastSet)
        LogLine ("LastLogonDate: {0}" -f $u.LastLogonDate)
        LogLine ("AccountExpirationDate: {0}" -f $u.AccountExpirationDate)

        # Compute password expiry if attribute exists (best effort)
        $expiry = $null
        try {
            if ($u.'msDS-UserPasswordExpiryTimeComputed') {
                $expiry = [DateTime]::FromFileTime([Int64]$u.'msDS-UserPasswordExpiryTimeComputed')
            }
        } catch {
            $expiry = $null
        }
        if ($expiry) {
            LogLine ("PasswordExpiryTimeComputed: {0}" -f $expiry)
        }

        # Quick recommended actions
        Section "Quick interpretation"
        if (-not $u.Enabled) { LogLine "Likely cause: Account disabled. Action: Enable account (per policy)." }
        if ($u.LockedOut)    { LogLine "Likely cause: Account lockout. Action: Unlock account, investigate bad password source." }
        if ($u.PasswordExpired) { LogLine "Likely cause: Password expired. Action: Reset password or have user change password." }
        if ($u.AccountExpirationDate -and $u.AccountExpirationDate -lt (Get-Date)) { LogLine "Likely cause: Account expired. Action: Extend/renew expiration date (per policy)." }

        if ($u.Enabled -and -not $u.LockedOut -and -not $u.PasswordExpired) {
            LogLine "AD status looks OK. If user still cannot sign in: check MFA, Conditional Access, device compliance, network, or service health."
        }

    } catch {
        LogLine ("AD lookup failed for '{0}': {1}" -f $Identity, $_.Exception.Message)
    }
}

# Optional Entra/Microsoft Graph checks (best effort)
Section "Optional Entra checks (Microsoft Graph if available and already connected)"

$graphAvailable = $false
try {
    $cmd = Get-Command Get-MgUser -ErrorAction Stop
    $graphAvailable = $true
} catch {
    LogLine "Microsoft Graph PowerShell not available (Get-MgUser not found)."
}

if ($graphAvailable) {
    try {
        # Do not auto-connect (keeps script safe/minimal). If not connected, just tell them.
        $ctx = $null
        try { $ctx = Get-MgContext } catch { $ctx = $null }

        if (-not $ctx -or -not $ctx.Account) {
            LogLine "Graph module present, but no active context found."
            LogLine "Tip: Connect first: Connect-MgGraph -Scopes 'User.Read.All','AuditLog.Read.All' (permissions required)."
        } else {
            LogLine ("Graph context: {0}" -f $ctx.Account)

            $user = $null
            try {
                # Try direct ID first; otherwise filter.
                $user = Get-MgUser -UserId $Identity -ErrorAction Stop
            } catch {
                $safe = $Identity.Replace("'", "''")
                $user = Get-MgUser -Filter "userPrincipalName eq '$safe' or mail eq '$safe'" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
            }

            if ($user) {
                LogLine ("Entra user: {0} ({1})" -f $user.DisplayName, $user.UserPrincipalName)
                LogLine ("AccountEnabled: {0}" -f $user.AccountEnabled)

                # Recent sign-ins (requires AuditLog permissions)
                try {
                    $safeUpn = $user.UserPrincipalName.Replace("'", "''")
                    $signins = Get-MgAuditLogSignIn -Filter "userPrincipalName eq '$safeUpn'" -Top 5 -ErrorAction Stop
                    if ($signins) {
                        Section "Recent sign-ins (top 5)"
                        foreach ($s in $signins) {
                            LogLine ("{0} | Status: {1} | App: {2} | IP: {3}" -f $s.CreatedDateTime, $s.Status.ErrorCode, $s.AppDisplayName, $s.IpAddress)
                        }
                    } else {
                        LogLine "No recent sign-ins returned."
                    }
                } catch {
                    LogLine ("Could not query sign-ins (likely missing permissions): {0}" -f $_.Exception.Message)
                }

            } else {
                LogLine "Could not locate Entra user with the provided identity."
            }
        }
    } catch {
        LogLine ("Entra check failed: {0}" -f $_.Exception.Message)
    }
}

Section "Complete"
LogLine "Done."
