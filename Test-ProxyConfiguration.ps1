<#
.SYNOPSIS
  Diagnose proxy/PAC configuration issues that commonly cause slow or broken internet.

.DESCRIPTION
  Collects key proxy-related settings that often drift in enterprise environments:
    - WinINET (user) proxy + PAC (Internet Options)
    - WinHTTP proxy (system) (used by many services/agents)
    - Environment variables (HTTP_PROXY/HTTPS_PROXY/NO_PROXY)
    - Basic connectivity checks with and without proxy (best-effort)

  Writes a human-readable log to out\HelpdeskLogs for easy ticket attachment.

  Safe remediation is OPTIONAL and limited to:
    - netsh winhttp reset proxy
    - ipconfig /flushdns
    - restarting WinHTTP Web Proxy Auto-Discovery Service (WinHttpAutoProxySvc) if present
  It does NOT modify user Internet Options registry settings unless you explicitly choose to do so (not included here).

.PARAMETER TicketId
  Optional ticket identifier used in log naming.

.PARAMETER EnableSafeRemediation
  If set, performs the limited safe remediation actions above.

.PARAMETER TestUrl
  URL used for a simple web request test. Default: https://www.msftconnecttest.com/connecttest.txt

.EXAMPLE
  .\Test-ProxyConfiguration.ps1

.EXAMPLE
  .\Test-ProxyConfiguration.ps1 -TicketId INC12345 -EnableSafeRemediation
#>

[CmdletBinding()]
param(
  [string]$TicketId,
  [switch]$EnableSafeRemediation,
  [string]$TestUrl = "https://www.msftconnecttest.com/connecttest.txt"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
if (-not (Test-Path $logsRoot)) { New-Item -Path $logsRoot -ItemType Directory -Force | Out-Null }

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('ProxyDiag', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }
$logPath = Join-Path $logsRoot (($baseName -join '_') + '.txt')

function LogLine {
  param([string]$Text)
  if ($null -eq $Text) { return }
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

function Try-Run {
  param(
    [string]$Label,
    [scriptblock]$Block
  )
  try {
    & $Block
  } catch {
    LogLine ("ERROR: {0}: {1}" -f $Label, $_.Exception.Message)
  }
}

function Get-WinInetProxyInfo {
  # User-level Internet Options proxy settings
  $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
  $obj = [ordered]@{
    ProxyEnable = $null
    ProxyServer = $null
    ProxyOverride = $null
    AutoConfigURL = $null
    AutoDetect = $null
  }

  if (Test-Path $key) {
    $p = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
    $obj.ProxyEnable   = $p.ProxyEnable
    $obj.ProxyServer   = $p.ProxyServer
    $obj.ProxyOverride = $p.ProxyOverride
    $obj.AutoConfigURL = $p.AutoConfigURL
    $obj.AutoDetect    = $p.AutoDetect
  }
  [pscustomobject]$obj
}

function Get-WinHttpProxyRaw {
  (netsh winhttp show proxy 2>&1) | Out-String
}

function Test-WebRequest {
  param(
    [string]$Url,
    [switch]$NoProxy
  )
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  try {
    $params = @{
      Uri = $Url
      Method = 'GET'
      UseBasicParsing = $true
      TimeoutSec = 10
    }
    if ($NoProxy) {
      # bypass proxy by using a blank proxy; on some PS versions this is ignored, so treat as best-effort
      $params.Proxy = $null
    }

    $resp = Invoke-WebRequest @params
    $sw.Stop()
    return [pscustomobject]@{
      Success = $true
      StatusCode = $resp.StatusCode
      Milliseconds = $sw.ElapsedMilliseconds
    }
  } catch {
    $sw.Stop()
    return [pscustomobject]@{
      Success = $false
      Error = $_.Exception.Message
      Milliseconds = $sw.ElapsedMilliseconds
    }
  }
}

Section "Proxy diagnostics"
LogLine ("Host: {0}" -f $hostname)
LogLine ("User: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Log: {0}" -f $logPath)

Section "WinINET (Internet Options) proxy settings (HKCU)"
$winInet = Get-WinInetProxyInfo
$winInet | Format-List | Out-String | ForEach-Object { LogLine $_.TrimEnd() }

Section "WinHTTP proxy settings (system)"
Try-Run "netsh winhttp show proxy" {
  $raw = Get-WinHttpProxyRaw
  $raw.Split("`n") | ForEach-Object { LogLine $_.TrimEnd() }
}

Section "Environment proxy variables"
$vars = @('HTTP_PROXY','HTTPS_PROXY','NO_PROXY','http_proxy','https_proxy','no_proxy')
foreach ($v in $vars) {
  $val = [Environment]::GetEnvironmentVariable($v, 'Process')
  if ($val) { LogLine ("{0} (Process): {1}" -f $v, $val) }
  $val = [Environment]::GetEnvironmentVariable($v, 'User')
  if ($val) { LogLine ("{0} (User): {1}" -f $v, $val) }
  $val = [Environment]::GetEnvironmentVariable($v, 'Machine')
  if ($val) { LogLine ("{0} (Machine): {1}" -f $v, $val) }
}

Section "Quick connectivity tests (best effort)"
LogLine ("Test URL: {0}" -f $TestUrl)
$r1 = Test-WebRequest -Url $TestUrl
if ($r1.Success) {
  LogLine ("With default settings: OK (HTTP {0}) in {1} ms" -f $r1.StatusCode, $r1.Milliseconds)
} else {
  LogLine ("With default settings: FAIL in {0} ms - {1}" -f $r1.Milliseconds, $r1.Error)
}

$r2 = Test-WebRequest -Url $TestUrl -NoProxy
if ($r2.Success) {
  LogLine ("Attempt no-proxy: OK (HTTP {0}) in {1} ms" -f $r2.StatusCode, $r2.Milliseconds)
} else {
  LogLine ("Attempt no-proxy: FAIL in {0} ms - {1}" -f $r2.Milliseconds, $r2.Error)
}

Section "Heuristics / likely causes"
# Simple heuristics
if ($winInet.ProxyEnable -eq 1 -and $winInet.ProxyServer) {
  LogLine ("User proxy is enabled with ProxyServer: {0}" -f $winInet.ProxyServer)
}
if ($winInet.AutoConfigURL) {
  LogLine ("PAC file configured (AutoConfigURL): {0}" -f $winInet.AutoConfigURL)
}
if ($winInet.AutoDetect -eq 1) {
  LogLine "Auto-detect is enabled (WPAD)."
}

$winHttpRaw = ""
try { $winHttpRaw = Get-WinHttpProxyRaw } catch { }
if ($winHttpRaw -match 'Direct access \(no proxy server\)') {
  LogLine "WinHTTP is set to Direct (no proxy)."
} else {
  LogLine "WinHTTP appears to have a proxy configured (or cannot be read)."
}

LogLine ""
LogLine "Common enterprise symptoms tied to proxy drift:"
LogLine "  - WinINET has proxy/PAC but WinHTTP differs (agents/services fail, browsing may work)"
LogLine "  - Stale PAC URL or WPAD delays cause slow initial page loads"
LogLine "  - Leftover VPN/Zscaler settings after disconnect"
LogLine "  - Environment variables forcing proxy for CLI tools"

if ($EnableSafeRemediation) {
  Section "Safe remediation (optional)"
  LogLine "Running: netsh winhttp reset proxy"
  Try-Run "netsh winhttp reset proxy" { (netsh winhttp reset proxy 2>&1) | ForEach-Object { LogLine $_ } }

  LogLine ""
  LogLine "Running: ipconfig /flushdns"
  Try-Run "ipconfig /flushdns" { (ipconfig /flushdns 2>&1) | ForEach-Object { LogLine $_ } }

  LogLine ""
  LogLine "Attempting: restart WinHttpAutoProxySvc (if present)"
  Try-Run "restart WinHttpAutoProxySvc" {
    $svc = Get-Service -Name 'WinHttpAutoProxySvc' -ErrorAction SilentlyContinue
    if ($svc) {
      if ($svc.Status -eq 'Running') {
        Restart-Service -Name 'WinHttpAutoProxySvc' -Force -ErrorAction Stop
        LogLine "WinHttpAutoProxySvc restarted."
      } else {
        Start-Service -Name 'WinHttpAutoProxySvc' -ErrorAction Stop
        LogLine "WinHttpAutoProxySvc started."
      }
    } else {
      LogLine "WinHttpAutoProxySvc not found on this system."
    }
  }

  Section "Retest after remediation"
  $r3 = Test-WebRequest -Url $TestUrl
  if ($r3.Success) {
    LogLine ("After remediation: OK (HTTP {0}) in {1} ms" -f $r3.StatusCode, $r3.Milliseconds)
  } else {
    LogLine ("After remediation: FAIL in {0} ms - {1}" -f $r3.Milliseconds, $r3.Error)
  }
}

Section "Complete"
LogLine "Done."
LogLine ("Log written: {0}" -f $logPath)
