<#
.SYNOPSIS
    Network performance diagnostics for slow or intermittent connections.

.DESCRIPTION
    Runs a focused set of checks for common "slow network / Internet" issues:
      - Adapter status and link speed
      - IP configuration and default gateway
      - DNS server configuration
      - Latency and packet loss to local and public targets
      - Optional safe remediation (DNS cache flush)

.PARAMETER TicketId
    Optional ticket identifier to include in the log file name.

.PARAMETER EnableSafeRemediation
    If specified, runs safe remediation steps such as DNS cache flush.

.EXAMPLE
    .\Test-NetworkPerformance.ps1

.EXAMPLE
    .\Test-NetworkPerformance.ps1 -TicketId INC12345 -EnableSafeRemediation
#>

[CmdletBinding()]
param(
    [string]$TicketId,
    [switch]$EnableSafeRemediation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Resolve paths for lightweight logging
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$logsRoot  = Join-Path -Path (Split-Path $scriptDir -Parent) -ChildPath 'out\HelpdeskLogs'

if (-not (Test-Path $logsRoot)) {
    New-Item -Path $logsRoot -ItemType Directory -Force | Out-Null
}

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

$baseNameParts = @("NetPerf", $hostname, $timestamp)
if ($TicketId) {
    $baseNameParts += $TicketId
}
$logName = ($baseNameParts -join '_') + ".txt"
$logPath = Join-Path -Path $logsRoot -ChildPath $logName

function Write-Section {
    param(
        [Parameter(Mandatory)]
        [string]$Title
    )
    $line = ('=' * 70)
    $text = @(
        ''
        $line
        "== $Title"
        $line
    ) -join [Environment]::NewLine

    Write-Host $text
    $text | Out-File -FilePath $logPath -Encoding UTF8 -Append
}

function Write-Info {
    param(
        [string]$Message
    )
    if ($Message -ne $null) {
        Write-Host $Message
        $Message | Out-File -FilePath $logPath -Encoding UTF8 -Append
    }
}

function Write-WarnInfo {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )
    Write-Warning $Message
    "[WARN] $Message" | Out-File -FilePath $logPath -Encoding UTF8 -Append
}

Write-Section "Network performance diagnostics started"
Write-Info ("Host: {0}" -f $hostname)
Write-Info ("Time: {0}" -f (Get-Date))
if ($TicketId) {
    Write-Info ("TicketId: {0}" -f $TicketId)
}

# STEP 1 - Adapter and link information
Write-Section "Step 1 - Adapter and link information"

try {
    $adapters = Get-NetAdapter -Physical | Sort-Object -Property ifIndex
} catch {
    Write-WarnInfo "Get-NetAdapter failed: $($_.Exception.Message)"
    $adapters = @()
}

if (-not $adapters) {
    Write-WarnInfo "No physical adapters found or Get-NetAdapter unavailable."
} else {
    foreach ($a in $adapters) {
        $msg = ("Adapter {0} ({1}) - Status: {2}, LinkSpeed: {3}" -f `
            $a.Name, $a.InterfaceDescription, $a.Status, $a.LinkSpeed)
        Write-Info $msg
    }
}

# STEP 2 - IP configuration and default gateway
Write-Section "Step 2 - IP configuration and default gateway"

try {
    $ipConfigs = Get-NetIPConfiguration
} catch {
    Write-WarnInfo "Get-NetIPConfiguration failed: $($_.Exception.Message)"
    $ipConfigs = @()
}

if (-not $ipConfigs) {
    Write-WarnInfo "No IP configuration information available."
} else {
    foreach ($cfg in $ipConfigs) {
        if ($cfg.IPv4Address -or $cfg.IPv6Address) {
            $ifName  = $cfg.InterfaceAlias
            $ifIndex = $cfg.InterfaceIndex
            $ipv4    = $cfg.IPv4Address | Select-Object -First 1
            $gwObj   = $null
            $gwAddr  = $null

            if ($cfg.IPv4DefaultGateway) {
                $gwObj = $cfg.IPv4DefaultGateway | Select-Object -First 1
                if ($gwObj -and $gwObj.PSObject.Properties.Name -contains 'NextHop') {
                    $gwAddr = $gwObj.NextHop
                }
                elseif ($gwObj -and $gwObj.PSObject.Properties.Name -contains 'Address') {
                    $gwAddr = $gwObj.Address
                }
            }

            $ipAddr = $null
            if ($ipv4 -and $ipv4.PSObject.Properties.Name -contains 'IPAddress') {
                $ipAddr = $ipv4.IPAddress
            }

            $msg = ("Interface {0} (Index {1}) - IPv4: {2}, Gateway: {3}" -f `
                $ifName,
                $ifIndex,
                ($ipAddr | ForEach-Object { $_ } | Select-Object -First 1),
                ($gwAddr | ForEach-Object { $_ } | Select-Object -First 1)
            )
            Write-Info $msg
        }
    }
}

# STEP 3 - DNS configuration
Write-Section "Step 3 - DNS configuration"

try {
    $dnsConfigs = Get-DnsClientServerAddress -AddressFamily IPv4
} catch {
    Write-WarnInfo "Get-DnsClientServerAddress failed: $($_.Exception.Message)"
    $dnsConfigs = @()
}

if (-not $dnsConfigs) {
    Write-WarnInfo "No DNS client server address information available."
} else {
    foreach ($cfg in $dnsConfigs) {
        $servers = ($cfg.ServerAddresses -join ', ')
        $msg = ("Interface {0} (Index {1}) - DNS Servers: {2}" -f `
            $cfg.InterfaceAlias,
            $cfg.InterfaceIndex,
            $servers)
        Write-Info $msg
    }
}

# Helper to run a basic latency test using Test-NetConnection or ping fallback
function Test-Target {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [string]$Target,
        [int]$Count = 4
    )

    if ([string]::IsNullOrWhiteSpace($Target)) {
        Write-WarnInfo "Skipping test for $Name because target host is empty."
        return
    }

    Write-Info ("--- Testing {0} ({1}) ---" -f $Name, $Target)

    $success = $false

    if (Get-Command Test-NetConnection -ErrorAction SilentlyContinue) {
        try {
            $result = Test-NetConnection -ComputerName $Target -WarningAction SilentlyContinue
            if ($null -ne $result) {
                $latency = $null
                if ($result.PSObject.Properties.Name -contains 'PingReplyDetails' -and
                    $result.PingReplyDetails -and
                    $result.PingReplyDetails.PSObject.Properties.Name -contains 'RoundtripTime') {
                    $latency = $result.PingReplyDetails.RoundtripTime
                }

                $msg = ("Reachable: {0}, PingSucceeded: {1}, Latency(ms): {2}" -f `
                    $result.PingSucceeded,
                    $result.PingSucceeded,
                    $latency)
                Write-Info $msg
                $success = $result.PingSucceeded
            }
        } catch {
            Write-WarnInfo "Test-NetConnection to $Target failed: $($_.Exception.Message)"
        }
    }

    if (-not $success) {
        if (Get-Command Test-Connection -ErrorAction SilentlyContinue) {
            try {
                $pings = Test-Connection -ComputerName $Target -Count $Count -ErrorAction Stop
                $avg = ($pings | Measure-Object -Property ResponseTime -Average).Average
                $loss = 100 - ((($pings | Measure-Object).Count * 100) / $Count)
                $msg = ("Ping success via Test-Connection. Average latency: {0:N1} ms, Approx loss: {1:N1}%" -f `
                    $avg, $loss)
                Write-Info $msg
                $success = $true
            } catch {
                Write-WarnInfo "Test-Connection to $Target failed: $($_.Exception.Message)"
            }
        } else {
            Write-WarnInfo "Neither Test-NetConnection nor Test-Connection available to test $Target."
        }
    }

    if (-not $success) {
        Write-WarnInfo "Target $Target appears unreachable or degraded."
    }
}

# STEP 4 - Latency tests to key targets
Write-Section "Step 4 - Latency and reachability tests"

# Try to identify a local default gateway as a "local network" target
$defaultGateway = $null
if ($ipConfigs) {
    $gwCandidate = $ipConfigs |
        Where-Object { $_.IPv4DefaultGateway } |
        Select-Object -First 1
    if ($gwCandidate -and $gwCandidate.IPv4DefaultGateway) {
        $gwObj = $gwCandidate.IPv4DefaultGateway | Select-Object -First 1
        if ($gwObj -and $gwObj.PSObject.Properties.Name -contains 'NextHop') {
            $defaultGateway = $gwObj.NextHop
        }
        elseif ($gwObj -and $gwObj.PSObject.Properties.Name -contains 'Address') {
            $defaultGateway = $gwObj.Address
        }
    }
}

if ($defaultGateway) {
    Test-Target -Name "Default gateway" -Target $defaultGateway
} else {
    Write-WarnInfo "No default gateway detected for latency testing."
}

# Identify a primary DNS server to test
$dnsTarget = $null
if ($dnsConfigs) {
    $dnsTarget = ($dnsConfigs[0].ServerAddresses | Select-Object -First 1)
}

if ($dnsTarget) {
    Test-Target -Name "Primary DNS server" -Target $dnsTarget
} else {
    Write-WarnInfo "No DNS server found for latency testing."
}

# Public targets (hard-coded, safe to use)
Test-Target -Name "Public DNS 8.8.8.8" -Target "8.8.8.8"
Test-Target -Name "Public DNS 1.1.1.1" -Target "1.1.1.1"

# STEP 5 - Optional safe remediation (DNS cache flush)
Write-Section "Step 5 - Optional safe remediation"

if ($EnableSafeRemediation) {
    Write-Info "Safe remediation requested. Flushing DNS client cache (ipconfig /flushdns)."
    try {
        $flush = & ipconfig /flushdns 2>&1
        $flush | Out-File -FilePath $logPath -Encoding UTF8 -Append
        Write-Info "DNS cache flush completed."
    } catch {
        Write-WarnInfo "ipconfig /flushdns failed: $($_.Exception.Message)"
    }
} else {
    Write-Info "Safe remediation not enabled. No changes were made."
}

Write-Section "Diagnostics complete"

Write-Info "Log file written to:"
Write-Info "  $logPath"

Write-Host ""
Write-Host "Network performance diagnostics complete." -ForegroundColor Green
