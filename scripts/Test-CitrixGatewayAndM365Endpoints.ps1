<#
.SYNOPSIS
    Diagnoses connectivity to Citrix Gateway, VPN, and Microsoft 365 / Entra ID endpoints.

.DESCRIPTION
    Tests resolution, port response, and latency for critical corporate network resources
    and Cloud/M365 endpoints. Works cross-platform with .NET fallbacks.

.PARAMETER CitrixGateway
    Optional Citrix Gateway hostname/IP (e.g. gateway.lawfirm.com) to test.

.PARAMETER TicketId
    Optional ticket identifier used in log naming.

.EXAMPLE
    .\Test-CitrixGatewayAndM365Endpoints.ps1 -CitrixGateway gateway.contoso.com -Verbose
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CitrixGateway,

    [Parameter(Mandatory = $false)]
    [string]$TicketId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load common utilities
. (Join-Path $PSScriptRoot 'Common.ps1')

$logsRoot = Join-Path -Path $global:RepoRoot -ChildPath 'out\HelpdeskLogs'
if (-not (Test-Path $logsRoot)) {
    New-Item -Path $logsRoot -ItemType Directory -Force | Out-Null
}

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$base = @("CitrixM365Diag", $hostname, $timestamp)
if ($TicketId) { $base += $TicketId }
$logPath = Join-Path $logsRoot (($base -join '_') + ".txt")

function LogLine {
    param([string]$Text)
    if ($Text -eq $null) { return }
    $Text | Out-File -FilePath $logPath -Encoding UTF8 -Append -ErrorAction SilentlyContinue
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

# Cross-platform helper functions
function Resolve-HostDns {
    param([string]$Hostname)
    try {
        if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
            $dns = Resolve-DnsName -Name $Hostname -ErrorAction Stop
            return $dns[0].IPAddress
        } else {
            $ips = [System.Net.Dns]::GetHostAddresses($Hostname)
            return $ips[0].IPAddressToString
        }
    } catch {
        throw $_
    }
}

function Test-Port443 {
    param([string]$Hostname)
    if (Get-Command Test-NetConnection -ErrorAction SilentlyContinue) {
        try {
            $result = Test-NetConnection -ComputerName $Hostname -Port 443 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            if ($result.TcpTestSucceeded -ne $null) {
                return [pscustomobject]@{
                    Succeeded = $result.TcpTestSucceeded
                    RTT = $result.PingReplyDetails.RoundtripTime
                }
            } else {
                return [pscustomobject]@{
                    Succeeded = ($result -eq $true)
                    RTT = $null
                }
            }
        } catch {
            return [pscustomobject]@{ Succeeded = $false; RTT = $null }
        }
    } else {
        # Fallback to .NET TcpClient
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $asyncResult = $tcpClient.BeginConnect($Hostname, 443, $null, $null)
            $wait = $asyncResult.AsyncWaitHandle.WaitOne(3000, $true) # 3s timeout
            $stopwatch.Stop()
            if ($wait -and $tcpClient.Connected) {
                $tcpClient.EndConnect($asyncResult)
                $rtt = $stopwatch.ElapsedMilliseconds
                return [pscustomobject]@{ Succeeded = $true; RTT = $rtt }
            } else {
                return [pscustomobject]@{ Succeeded = $false; RTT = $null }
            }
        } catch {
            return [pscustomobject]@{ Succeeded = $false; RTT = $null }
        } finally {
            $tcpClient.Close()
        }
    }
}

Section "Citrix Gateway & Microsoft 365 Diagnostics"
LogLine ("Operator: {0}" -f $env:USERNAME)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Log: {0}" -f $logPath)

# 1. Local Network Preflight
Section "Local Adapter Diagnostics"
try {
    if (Get-Command Get-NetAdapter -ErrorAction SilentlyContinue) {
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
        foreach ($adapter in $adapters) {
            LogLine ("Interface: {0} ({1}) - Speed: {2}" -f $adapter.Name, $adapter.InterfaceDescription, $adapter.LinkSpeed)
        }
    } else {
        # Fallback to standard network interface check from .NET which works everywhere
        $adapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | 
                    Where-Object { $_.OperationalStatus -eq 'Up' }
        foreach ($adapter in $adapters) {
            LogLine ("Interface: {0} ({1}) - Speed: {2} Mbps" -f $adapter.Name, $adapter.Description, [math]::Round(($adapter.Speed / 1MB), 2))
        }
    }
} catch {
    LogLine "Failed to retrieve local net adapters."
}

# 2. Microsoft 365 / Entra ID Endpoints
Section "Microsoft 365 & Entra ID Endpoints"
$m365Endpoints = @(
    "login.microsoftonline.com",  # Authentication
    "outlook.office365.com",      # Mail
    "graph.microsoft.com",        # Microsoft Graph API
    "portal.office.com",          # Office Portal
    "teams.microsoft.com"         # Teams
)

foreach ($endpoint in $m365Endpoints) {
    LogLine ("Testing DNS resolution for {0}..." -f $endpoint)
    try {
        $ip = Resolve-HostDns -Hostname $endpoint
        LogLine ("  [SUCCESS] Resolves to {0}" -f $ip)
        
        # Test TCP 443
        LogLine ("  Testing TCP Port 443 (HTTPS) on {0}..." -f $endpoint)
        $conn = Test-Port443 -Hostname $endpoint
        if ($conn.Succeeded) {
            if ($conn.RTT -ne $null) {
                LogLine ("  [SUCCESS] Port 443 is reachable. RTT: {0} ms" -f $conn.RTT)
            } else {
                LogLine "  [SUCCESS] Port 443 is reachable."
            }
        } else {
            LogLine "  [WARNING] Port 443 is UNREACHABLE!"
        }
    } catch {
        LogLine ("  [ERROR] DNS resolution failed: {0}" -f $_.Exception.Message)
    }
}

# 3. Citrix Gateway / VPN Endpoint Checks
if ($CitrixGateway) {
    Section ("Citrix Gateway / VPN Diagnostics: {0}" -f $CitrixGateway)
    LogLine ("Testing Citrix Gateway DNS resolution..." -f $CitrixGateway)
    try {
        $ip = Resolve-HostDns -Hostname $CitrixGateway
        LogLine ("  [SUCCESS] Gateway resolves to {0}" -f $ip)

        LogLine ("  Testing TCP Port 443 (HTTPS) on Citrix Gateway..." -f $CitrixGateway)
        $conn = Test-Port443 -Hostname $CitrixGateway
        if ($conn.Succeeded) {
            if ($conn.RTT -ne $null) {
                LogLine ("  [SUCCESS] TCP Connection to Citrix Gateway succeeded. RTT: {0} ms" -f $conn.RTT)
            } else {
                LogLine "  [SUCCESS] TCP Connection to Citrix Gateway succeeded."
            }
        } else {
            LogLine "  [ERROR] TCP Connection to Citrix Gateway failed! The gateway might be down, firewall blocked, or route is missing."
        }
    } catch {
        LogLine ("  [ERROR] Gateway resolution or test failed: {0}" -f $_.Exception.Message)
    }
} else {
    Section "Citrix Gateway / VPN Diagnostics"
    LogLine "No Citrix Gateway hostname provided for specific test. (Pass -CitrixGateway parameter to run this check)"
}

Section "Diagnostics Complete"
LogLine "Done."
Write-AuditLog -Severity INFO -Action "CitrixM365Diagnostics" -Message "Ran Citrix Gateway & M365 Endpoint diagnostics. Log: $logPath"
