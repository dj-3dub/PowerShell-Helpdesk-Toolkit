<#
.SYNOPSIS
  Collect Wi-Fi diagnostics for slow / unstable wireless connections (enterprise-friendly).

.DESCRIPTION
  Gathers common Wi-Fi troubleshooting artifacts and writes a ZIP bundle under:
    out\HelpdeskLogs\WiFiDiag_<COMPUTER>_<timestamp>[_Ticket].zip

  Includes:
    - netsh wlan show interfaces
    - netsh wlan show drivers
    - netsh wlan show profiles
    - ipconfig /all
    - route print
    - DNS client info (Get-DnsClientServerAddress)
    - Adapter link and power info (Get-NetAdapter, Get-NetAdapterAdvancedProperty)
    - Optional: WLAN Report HTML (netsh wlan show wlanreport)
    - Basic latency tests to gateway and public DNS (best effort)

  Safe remediation is OPTIONAL and limited to:
    - ipconfig /flushdns
    - ipconfig /renew
    - restart WLAN AutoConfig (WlanSvc)
  (No adapter disable/enable by default.)

.PARAMETER TicketId
  Optional ticket identifier used in bundle naming.

.PARAMETER IncludeWlanReport
  If set, generates the Windows WLAN report (HTML) and includes it.

.PARAMETER EnableSafeRemediation
  If set, performs the limited safe remediation actions and then re-runs key commands.

.PARAMETER PublicTestHost
  Public ping target. Default: 8.8.8.8

.EXAMPLE
  .\Collect-WiFiDiagnostics.ps1 -TicketId INC12345 -IncludeWlanReport

.EXAMPLE
  .\Collect-WiFiDiagnostics.ps1 -EnableSafeRemediation
#>

[CmdletBinding()]
param(
  [string]$TicketId,
  [switch]$IncludeWlanReport,
  [switch]$EnableSafeRemediation,
  [string]$PublicTestHost = "8.8.8.8"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function New-Dir([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}

function Try-Run([string]$Label, [scriptblock]$Block) {
  try { & $Block } catch { Write-Output ("ERROR: {0}: {1}" -f $Label, $_.Exception.Message) }
}

function Write-TextFile([string]$Path, [string[]]$Lines) {
  $Lines | Out-File -FilePath $Path -Encoding UTF8
}

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
New-Dir $logsRoot

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('WiFiDiag', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }
$workDir = Join-Path $logsRoot (($baseName -join '_') + '_WORK')
New-Dir $workDir

$meta = @()
$meta += "Host: $hostname"
$meta += "User: $($env:USERDOMAIN)\$($env:USERNAME)"
$meta += "Time: $(Get-Date)"
$meta += "Ticket: $TicketId"
$meta += "IncludeWlanReport: $IncludeWlanReport"
$meta += "EnableSafeRemediation: $EnableSafeRemediation"
Write-TextFile -Path (Join-Path $workDir 'README.txt') -Lines $meta

# Helper: detect Wi-Fi adapters
$wifiAdapters = @()
Try-Run "Get-NetAdapter Wi-Fi detection" {
  $wifiAdapters = Get-NetAdapter -Physical -ErrorAction Stop | Where-Object {
    $_.Status -ne 'Disabled' -and ($_.InterfaceDescription -match 'Wireless|Wi-Fi|WLAN|802\.11' -or $_.Name -match 'Wi-Fi|WLAN')
  }
}

# STEP 1: Core dumps
Write-Host "Collecting Wi-Fi diagnostics..." -ForegroundColor Cyan

Try-Run "netsh wlan show interfaces" {
  (netsh wlan show interfaces 2>&1) | Out-File (Join-Path $workDir 'netsh_wlan_show_interfaces.txt') -Encoding UTF8
}
Try-Run "netsh wlan show drivers" {
  (netsh wlan show drivers 2>&1) | Out-File (Join-Path $workDir 'netsh_wlan_show_drivers.txt') -Encoding UTF8
}
Try-Run "netsh wlan show profiles" {
  (netsh wlan show profiles 2>&1) | Out-File (Join-Path $workDir 'netsh_wlan_show_profiles.txt') -Encoding UTF8
}
Try-Run "ipconfig /all" {
  (ipconfig /all 2>&1) | Out-File (Join-Path $workDir 'ipconfig_all.txt') -Encoding UTF8
}
Try-Run "route print" {
  (route print 2>&1) | Out-File (Join-Path $workDir 'route_print.txt') -Encoding UTF8
}

Try-Run "Get-DnsClientServerAddress" {
  Get-DnsClientServerAddress -AddressFamily IPv4,IPv6 -ErrorAction Stop |
    Format-Table -AutoSize | Out-String |
    Out-File (Join-Path $workDir 'dns_client_server_address.txt') -Encoding UTF8
}

Try-Run "Get-NetAdapter + advanced props" {
  Get-NetAdapter -ErrorAction Stop | Sort-Object Name |
    Format-Table -AutoSize | Out-String |
    Out-File (Join-Path $workDir 'netadapter_table.txt') -Encoding UTF8

  if ($wifiAdapters -and $wifiAdapters.Count -gt 0) {
    foreach ($a in $wifiAdapters) {
      $safeName = ($a.Name -replace '[^a-zA-Z0-9\-_ ]','_')
      Try-Run "Advanced props $($a.Name)" {
        Get-NetAdapterAdvancedProperty -Name $a.Name -ErrorAction Stop |
          Sort-Object DisplayName |
          Format-Table -AutoSize | Out-String |
          Out-File (Join-Path $workDir ("netadapter_advprops_{0}.txt" -f $safeName)) -Encoding UTF8
      }
    }
  }
}

# STEP 2: Wlan report (optional)
if ($IncludeWlanReport) {
  Try-Run "netsh wlan show wlanreport" {
    (netsh wlan show wlanreport 2>&1) | Out-File (Join-Path $workDir 'netsh_wlan_show_wlanreport_output.txt') -Encoding UTF8
    $reportPath = Join-Path $env:ProgramData 'Microsoft\Windows\WlanReport\wlan-report-latest.html'
    if (Test-Path $reportPath) {
      Copy-Item -Path $reportPath -Destination (Join-Path $workDir 'wlan-report-latest.html') -Force
    } else {
      "WLAN report not found at expected path: $reportPath" | Out-File (Join-Path $workDir 'wlan_report_missing.txt') -Encoding UTF8
    }
  }
}

# STEP 3: Basic reachability tests
# Derive default gateway (best effort)
$defaultGateway = $null
Try-Run "Default gateway" {
  $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop | Sort-Object RouteMetric | Select-Object -First 1
  if ($route -and $route.NextHop) { $defaultGateway = $route.NextHop }
}

Try-Run "Ping tests" {
  $lines = @()
  $lines += "DefaultGateway: $defaultGateway"
  $lines += ""
  if ($defaultGateway) {
    $lines += "ping -n 10 $defaultGateway"
    $lines += (ping -n 10 $defaultGateway 2>&1)
    $lines += ""
  } else {
    $lines += "Default gateway not detected."
    $lines += ""
  }

  $lines += "ping -n 10 $PublicTestHost"
  $lines += (ping -n 10 $PublicTestHost 2>&1)
  $lines += ""

  $lines | Out-File (Join-Path $workDir 'ping_tests.txt') -Encoding UTF8
}

# STEP 4: Safe remediation (optional) + re-run key bits
if ($EnableSafeRemediation) {
  Write-Host "Running safe remediation..." -ForegroundColor Yellow
  Try-Run "ipconfig /flushdns" {
    (ipconfig /flushdns 2>&1) | Out-File (Join-Path $workDir 'remed_flushdns.txt') -Encoding UTF8
  }
  Try-Run "ipconfig /renew" {
    (ipconfig /renew 2>&1) | Out-File (Join-Path $workDir 'remed_renew.txt') -Encoding UTF8
  }
  Try-Run "Restart WlanSvc" {
    $svc = Get-Service -Name 'WlanSvc' -ErrorAction SilentlyContinue
    if ($svc) {
      if ($svc.Status -eq 'Running') { Restart-Service -Name 'WlanSvc' -Force -ErrorAction Stop }
      else { Start-Service -Name 'WlanSvc' -ErrorAction Stop }
      "WlanSvc restarted/started." | Out-File (Join-Path $workDir 'remed_wlansvc.txt') -Encoding UTF8
    } else {
      "WlanSvc not found." | Out-File (Join-Path $workDir 'remed_wlansvc.txt') -Encoding UTF8
    }
  }

  Try-Run "Re-run netsh wlan show interfaces" {
    (netsh wlan show interfaces 2>&1) | Out-File (Join-Path $workDir 'netsh_wlan_show_interfaces_after.txt') -Encoding UTF8
  }
}

# Create zip bundle
$zipPath = Join-Path $logsRoot (($baseName -join '_') + '.zip')
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($workDir, $zipPath)

# Cleanup workdir (leave if you want to debug)
Try { Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue } Catch {}

Write-Host ""
Write-Host "Wi-Fi diagnostics bundle created:" -ForegroundColor Green
Write-Host "  $zipPath" -ForegroundColor Green
