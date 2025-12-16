<#
.SYNOPSIS
    Runs VPN connectivity and configuration diagnostics.

.DESCRIPTION
    Collects information useful for VPN-related tickets:

      - Lists configured Windows VPN connections (Get-VpnConnection if available)
      - Lists network adapters that look like VPN/tunnel interfaces
      - Shows current routing table and IP configuration
      - Tests connectivity to:
          * Public endpoint (e.g. 8.8.8.8 / public URL)
          * Optional internal/VPN-only endpoint (e.g. intranet.corp.local)
      - Optionally collects all outputs into a ZIP under out\VpnDiagnostics

.PARAMETER VpnName
    Optional: Name of a specific VPN connection to focus on (matches
    Get-VpnConnection -Name). If omitted, all are shown.

.PARAMETER PublicTestHost
    Host or IP to use for "internet" connectivity test.
    Default: 8.8.8.8

.PARAMETER InternalTestHost
    Optional: internal hostname or IP that should be reachable only
    over VPN (e.g. intranet.corp.local, 10.0.0.10).

.PARAMETER CollectLogs
    If specified, writes diagnostics to a folder and compresses into a ZIP
    under out\VpnDiagnostics in the repo root.

.PARAMETER OutputRoot
    Optional override path for diagnostics/logs.

.EXAMPLE
    .\Test-VpnDiagnostics.ps1 -CollectLogs -InternalTestHost intranet.corp.local -Verbose

.EXAMPLE
    .\Test-VpnDiagnostics.ps1 -VpnName 'Corp VPN' -PublicTestHost '1.1.1.1' -Verbose
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$VpnName,

    [Parameter(Mandatory = $false)]
    [string]$PublicTestHost = '8.8.8.8',

    [Parameter(Mandatory = $false)]
    [string]$InternalTestHost,

    [Parameter(Mandatory = $false)]
    [switch]$CollectLogs,

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Verbose "Starting VPN diagnostics..."

# --- repo root + output folder (WSL/UNC safe string-based) ---
$repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent

if (-not $OutputRoot) {
    $OutputRoot = Join-Path $repoRoot 'out\VpnDiagnostics'
}

$null = New-Item -Path $OutputRoot -ItemType Directory -Force -ErrorAction SilentlyContinue

$timestamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$sessionName = "VpnDiag_{0}_{1}" -f $env:COMPUTERNAME, $timestamp
$sessionRoot = Join-Path $OutputRoot $sessionName

if ($CollectLogs) {
    $null = New-Item -Path $sessionRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
    Write-Verbose "Diagnostics session folder: $sessionRoot"
}

# Helper to write text output to file when -CollectLogs is used
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

# --- 1. Configured VPN connections (built-in Windows VPN) ---
$vpnConnCmd = Get-Command -Name Get-VpnConnection -ErrorAction SilentlyContinue
if ($vpnConnCmd) {
    try {
        $vpnFilter = @{}
        if ($VpnName) {
            $vpnFilter.Name = $VpnName
        }

        $vpnConnections = Get-VpnConnection @vpnFilter -ErrorAction SilentlyContinue

        if ($vpnConnections) {
            Write-Host "=== Configured VPN Connections (Get-VpnConnection) ===" -ForegroundColor Cyan
            $vpnConnections | Format-Table -AutoSize
            $vpnText = $vpnConnections | Out-String
            Write-DiagFile -RelativePath 'vpn\VpnConnections.txt' -Content $vpnText
        }
        else {
            Write-Host "No VPN connections found (Get-VpnConnection) for filter '$VpnName'." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Warning "Failed to query Get-VpnConnection: $($_.Exception.Message)"
    }
}
else {
    Write-Verbose "Get-VpnConnection not available on this system."
}

# --- 2. Network adapters that look like VPN/tunnel ---
Write-Host "=== Network Adapters (potential VPN/tunnel) ===" -ForegroundColor Cyan

$adapterPatterns = @(
    '*VPN*',
    '*Tunnel*',
    '*Virtual*',
    '*WAN Miniport*'
)

$adapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
    foreach ($p in $adapterPatterns) {
        if ($_.InterfaceDescription -like $p -or $_.Name -like $p) { return $true }
    }
    return $false
}

if ($adapters) {
    $adapters | Format-Table Name, InterfaceDescription, Status, MacAddress, LinkSpeed -AutoSize
    $adapterText = $adapters | Out-String
    Write-DiagFile -RelativePath 'vpn\NetAdapters.txt' -Content $adapterText
}
else {
    Write-Host "No obvious VPN/tunnel adapters detected by pattern match." -ForegroundColor Yellow
}

# --- 3. Routing + IP configuration ---
Write-Host "=== Routing Table (route print) ===" -ForegroundColor Cyan
$routePrint = (& route print) | Out-String
Write-Host $routePrint
Write-DiagFile -RelativePath 'network\RoutePrint.txt' -Content $routePrint

Write-Host "=== IP Configuration (ipconfig /all) ===" -ForegroundColor Cyan
$ipconfig = (& ipconfig /all) | Out-String
Write-Host $ipconfig
Write-DiagFile -RelativePath 'network\IpconfigAll.txt' -Content $ipconfig

try {
    Write-Host "=== Get-NetIPConfiguration ===" -ForegroundColor Cyan
    $netIpConfig = Get-NetIPConfiguration -ErrorAction SilentlyContinue
    if ($netIpConfig) {
        $netIpText = $netIpConfig | Format-List | Out-String
        Write-Host $netIpText
        Write-DiagFile -RelativePath 'network\NetIPConfiguration.txt' -Content $netIpText
    }
}
catch {
    Write-Warning "Get-NetIPConfiguration failed: $($_.Exception.Message)"
}

# --- 4. Connectivity tests ---
function Test-HostConnectivity {
    param(
        [Parameter(Mandatory = $true)][string]$Target,
        [Parameter(Mandatory = $false)][string]$Label = $Target
    )

    Write-Host "=== Connectivity test: $Label ($Target) ===" -ForegroundColor Cyan

    $result = $null
    try {
        $result = Test-NetConnection -ComputerName $Target -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Test-NetConnection failed for $Target: $($_.Exception.Message)"
    }

    if ($result) {
        $ok = $result.PingSucceeded
        Write-Host ("  PingSucceeded : {0}" -f $ok)
        Write-Host ("  Address       : {0}" -f $result.RemoteAddress)
        Write-Host ("  RoundTrip(ms) : {0}" -f $result.PingReplyDetails.RoundtripTime)

        return [pscustomobject]@{
            Target           = $Target
            Label            = $Label
            PingSucceeded    = $ok
            RemoteAddress    = $result.RemoteAddress
            RoundTripTimeMs  = $result.PingReplyDetails.RoundtripTime
        }
    }
    else {
        Write-Host "  No Test-NetConnection result for $Target." -ForegroundColor Yellow
        return [pscustomobject]@{
            Target           = $Target
            Label            = $Label
            PingSucceeded    = $false
            RemoteAddress    = $null
            RoundTripTimeMs  = $null
        }
    }
}

$connectivityResults = New-Object System.Collections.Generic.List[object]

# Public endpoint
$connectivityResults.Add( (Test-HostConnectivity -Target $PublicTestHost -Label 'PublicTestHost') )

# Internal/VPN-only endpoint
if ($InternalTestHost) {
    $connectivityResults.Add( (Test-HostConnectivity -Target $InternalTestHost -Label 'InternalTestHost') )
}

if ($CollectLogs -and $connectivityResults.Count -gt 0) {
    $connectivityResults |
        Export-Csv -Path (Join-Path $sessionRoot 'network\ConnectivityTests.csv') -NoTypeInformation -Encoding UTF8
}

# --- 5. Summary on screen ---
Write-Host ""
Write-Host "=== VPN Diagnostics Summary ===" -ForegroundColor Cyan

if ($connectivityResults.Count -gt 0) {
    foreach ($r in $connectivityResults) {
        Write-Host ("[{0}] {1} -> PingSucceeded={2}, RTT={3}ms" -f $r.Label, $r.Target, $r.PingSucceeded, $r.RoundTripTimeMs)
    }
}
else {
    Write-Host "No connectivity tests were performed." -ForegroundColor Yellow
}

if ($CollectLogs) {
    $zipName = "VpnDiag_{0}_{1}.zip" -f $env:COMPUTERNAME, $timestamp
    $zipPath = Join-Path $OutputRoot $zipName

    Write-Verbose "Creating ZIP archive $zipPath from $sessionRoot"

    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    }

    Compress-Archive -Path (Join-Path $sessionRoot '*') -DestinationPath $zipPath -ErrorAction Stop
    Write-Host "VPN diagnostics collected: $zipPath" -ForegroundColor Green
}
else {
    Write-Host "VPN diagnostics complete. Use -CollectLogs to generate a ZIP for tickets." -ForegroundColor Green
}
