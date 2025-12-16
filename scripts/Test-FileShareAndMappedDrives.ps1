<#
.SYNOPSIS
  SMB / File Share / Mapped Drive diagnostics (read-only by default).

.DESCRIPTION
  Designed for common enterprise issues:
    - \\server\share unreachable
    - mapped drives missing / not reconnecting
    - shares "slow" while Internet is fine
    - VPN connected but internal resources fail
    - credential / Kerberos / multiple-connection issues

  Produces a ticket-friendly log and optional ZIP bundle with raw artifacts.

  Output:
    out\HelpdeskLogs\FileShareDiag_<COMPUTER>_<timestamp>[_Ticket].txt
  Optional ZIP:
    out\HelpdeskLogs\FileShareDiag_<COMPUTER>_<timestamp>[_Ticket].zip

.PARAMETER TicketId
  Optional ticket identifier used in output naming.

.PARAMETER Servers
  One or more file servers to test (name or IP). If omitted, the script will attempt
  to discover servers from existing SMB mappings.

.PARAMETER Shares
  Optional UNC shares to validate, e.g. \\fileserver\dept

.PARAMETER EnableSafeRemediation
  When set, enables prompts for safe remediations (no changes occur unless you confirm):
    - remove a selected stale SMB mapping
    - delete selected net use mapping
    - klist purge (Kerberos refresh)

.PARAMETER CreateZipBundle
  If set, zips raw artifacts and command outputs.

.EXAMPLE
  .\Test-FileShareAndMappedDrives.ps1 -Servers filesrv01,filesrv02 -Shares \\filesrv01\dept -TicketId INC123 -CreateZipBundle

.EXAMPLE
  .\Test-FileShareAndMappedDrives.ps1 -EnableSafeRemediation
#>

[CmdletBinding()]
param(
  [string]$TicketId,
  [string[]]$Servers,
  [string[]]$Shares,
  [switch]$EnableSafeRemediation,
  [switch]$CreateZipBundle
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function New-Dir([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}

function LogLine([string]$Text) {
  if ($null -eq $Text) { return }
  $Text | Out-File -FilePath $script:LogPath -Encoding UTF8 -Append
  Write-Host $Text
}

function Section([string]$Title) {
  $line = ('=' * 70)
  LogLine ""
  LogLine $line
  LogLine ("== {0}" -f $Title)
  LogLine $line
}

function Try-Run([string]$Label, [scriptblock]$Block) {
  try { & $Block } catch { LogLine ("ERROR: {0}: {1}" -f $Label, $_.Exception.Message) }
}

function Save-Text([string]$Name, [string[]]$Lines) {
  $p = Join-Path $script:WorkDir $Name
  $Lines | Out-File -FilePath $p -Encoding UTF8
  return $p
}

function Save-Cmd([string]$Name, [string]$CommandLine) {
  $p = Join-Path $script:WorkDir $Name
  Try-Run $Name { cmd /c $CommandLine 2>&1 | Out-File -FilePath $p -Encoding UTF8 }
  return $p
}

function Resolve-HostIp([string]$Host) {
  try {
    $res = Resolve-DnsName -Name $Host -ErrorAction Stop | Where-Object { $_.IPAddress } | Select-Object -First 1
    return $res.IPAddress
  } catch { return $null }
}

function Test-TcpPort([string]$Host, [int]$Port, [int]$TimeoutMs = 2000) {
  try {
    $client = New-Object System.Net.Sockets.TcpClient
    $iar = $client.BeginConnect($Host, $Port, $null, $null)
    $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
    if (-not $ok) { $client.Close(); return $false }
    $client.EndConnect($iar)
    $client.Close()
    return $true
  } catch { return $false }
}

function Get-DefaultGateway {
  try {
    $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
      Sort-Object -Property RouteMetric, InterfaceMetric | Select-Object -First 1
    if ($route -and $route.NextHop) { return $route.NextHop }
  } catch {}
  return $null
}

# Paths
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
New-Dir $logsRoot

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('FileShareDiag', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }

$script:LogPath = Join-Path $logsRoot (($baseName -join '_') + '.txt')
$script:WorkDir = Join-Path $logsRoot (($baseName -join '_') + '_WORK')
New-Dir $script:WorkDir

Section "SMB / File Share / Mapped Drive Diagnostics"
LogLine ("Host: {0}" -f $hostname)
LogLine ("User: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Ticket: {0}" -f $TicketId)
LogLine ("Log: {0}" -f $script:LogPath)
LogLine ("Artifacts: {0}" -f $script:WorkDir)

Section "Network context"
Try-Run "Adapter/IP" {
  $adapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Sort-Object -Property Status, LinkSpeed -Descending
  foreach ($a in $adapters) {
    LogLine ("Adapter: {0}  Status={1}  Link={2}  MAC={3}" -f $a.Name, $a.Status, $a.LinkSpeed, $a.MacAddress)
  }

  $ip = Get-NetIPConfiguration -ErrorAction SilentlyContinue
  foreach ($i in $ip) {
    if ($i.IPv4Address) {
      $dns = @($i.DnsServer.ServerAddresses) -join ', '
      $gw  = $null
      if ($i.IPv4DefaultGateway) { $gw = @($i.IPv4DefaultGateway.NextHop) -join ',' }
      LogLine ("If={0}  IPv4={1}  GW={2}  DNS={3}" -f $i.InterfaceAlias, $i.IPv4Address.IPAddress, ($gw ?? ''), $dns)
    }
  }
}
Save-Cmd -Name 'ipconfig_all.txt' -CommandLine 'ipconfig /all' | Out-Null
Save-Cmd -Name 'route_print.txt'  -CommandLine 'route print'   | Out-Null

$gateway = Get-DefaultGateway
Section "Default gateway"
if ($gateway) {
  LogLine ("Gateway: {0}" -f $gateway)
  $ping = $false
  Try-Run "Ping gateway" { $ping = Test-Connection -ComputerName $gateway -Count 2 -Quiet -ErrorAction Stop }
  LogLine ("Ping gateway: {0}" -f $ping)
} else {
  LogLine "Gateway: (not detected)"
}

Section "Current SMB mappings"
Try-Run "Get-SmbMapping" {
  $maps = @(Get-SmbMapping -ErrorAction SilentlyContinue)
  if ($maps.Count -lt 1) {
    LogLine "No SMB mappings returned by Get-SmbMapping."
  } else {
    for ($i=0; $i -lt $maps.Count; $i++) {
      $m = $maps[$i]
      LogLine ("[{0}] Local={1} Remote={2} User={3} Status={4}" -f ($i+1), $m.LocalPath, $m.RemotePath, $m.UserName, $m.Status)
    }
  }
}
Save-Cmd -Name 'net_use.txt' -CommandLine 'net use' | Out-Null
Save-Cmd -Name 'klist.txt'   -CommandLine 'klist'   | Out-Null

# Auto-discover servers from mappings if not provided
if (-not $Servers -or $Servers.Count -lt 1) {
  Try-Run "Discover servers from SMB mappings" {
    $maps = @(Get-SmbMapping -ErrorAction SilentlyContinue)
    $hosts = @()
    foreach ($m in $maps) {
      if ($m.RemotePath -match '^\\\\([^\\]+)\\') {
        $hosts += $Matches[1]
      }
    }
    $hosts = $hosts | Where-Object { $_ } | Select-Object -Unique
    if ($hosts.Count -gt 0) {
      $Servers = $hosts
      LogLine ("Auto-discovered file servers from mappings: {0}" -f ($Servers -join ', '))
    }
  }
}

Section "Connectivity tests (DNS + TCP 445 SMB)"
if (-not $Servers -or $Servers.Count -lt 1) {
  LogLine "No servers specified or discovered. (You can pass -Servers filesrv01,filesrv02)"
} else {
  foreach ($s in $Servers) {
    $ip = Resolve-HostIp $s
    LogLine ("Server: {0}  IP={1}" -f $s, ($ip ?? '(DNS failed)'))

    $ping = $false
    Try-Run "Ping $s" { $ping = Test-Connection -ComputerName $s -Count 2 -Quiet -ErrorAction Stop }
    LogLine ("  Ping: {0}" -f $ping)

    $tcp445 = Test-TcpPort -Host $s -Port 445
    LogLine ("  TCP 445 (SMB): {0}" -f $tcp445)

    $tcp135 = Test-TcpPort -Host $s -Port 135
    LogLine ("  TCP 135 (RPC): {0}" -f $tcp135)
  }
}

Section "Share path tests"
if ($Shares -and $Shares.Count -gt 0) {
  foreach ($unc in $Shares) {
    LogLine ("Share: {0}" -f $unc)
    $ok = $false
    Try-Run "Test-Path $unc" { $ok = Test-Path -Path $unc -ErrorAction Stop }
    LogLine ("  Test-Path: {0}" -f $ok)
  }
} else {
  LogLine "No -Shares provided. (Optional) Example: -Shares \\filesrv01\dept"
}

Section "Common causes / quick hints"
LogLine "If TCP 445 is False but ping is True: firewall or network policy blocking SMB."
LogLine "If DNS fails for internal names while VPN is connected: check VPN DNS / split tunnel DNS settings."
LogLine "If net use shows 'multiple connections' errors: remove stale mappings for that server."
LogLine "If share prompts for credentials repeatedly: check Kerberos tickets (klist) + time skew."

if ($EnableSafeRemediation) {
  Section "Safe remediation (optional prompts)"
  LogLine "No changes will be made unless you confirm."

  Try-Run "Remove SMB mapping" {
    $maps = @(Get-SmbMapping -ErrorAction SilentlyContinue)
    if ($maps.Count -gt 0) {
      $ans = Read-Host "Remove a specific SMB mapping? (Y/N)"
      if ($ans -match '^[Yy]') {
        for ($i=0; $i -lt $maps.Count; $i++) {
          LogLine ("[{0}] {1} -> {2}  Status={3}" -f ($i+1), $maps[$i].LocalPath, $maps[$i].RemotePath, $maps[$i].Status)
        }
        $pick = Read-Host "Enter number to remove (or Enter to skip)"
        if ($pick -match '^\d+$') {
          $idx = [int]$pick - 1
          if ($idx -ge 0 -and $idx -lt $maps.Count) {
            $target = $maps[$idx].LocalPath
            $confirm = Read-Host ("Confirm remove SMB mapping {0}? (Y/N)" -f $target)
            if ($confirm -match '^[Yy]') {
              Remove-SmbMapping -LocalPath $target -Force -UpdateProfile -ErrorAction Stop
              LogLine ("Removed SMB mapping: {0}" -f $target)
            }
          }
        }
      }
    }
  }

  Try-Run "net use cleanup" {
    $ans = Read-Host "Run 'net use' cleanup for a specific server? (Y/N)"
    if ($ans -match '^[Yy]') {
      $server = Read-Host "Enter server name (e.g. filesrv01)"
      if ($server) {
        $cmd = 'net use \\\\{0} /delete /y' -f $server
        LogLine ("Running: {0}" -f $cmd)
        cmd /c $cmd 2>&1 | ForEach-Object { LogLine ("  {0}" -f $_) }
      }
    }
  }

  Try-Run "klist purge" {
    $ans = Read-Host "Purge Kerberos tickets (klist purge)? (Y/N)"
    if ($ans -match '^[Yy]') {
      cmd /c 'klist purge' 2>&1 | ForEach-Object { LogLine ("  {0}" -f $_) }
      LogLine "Kerberos tickets purged. Re-test share access."
    }
  }
}

Section "Complete"
if ($CreateZipBundle) {
  $zipPath = Join-Path $logsRoot (($baseName -join '_') + '.zip')
  Try-Run "Create ZIP bundle" {
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($script:WorkDir, $zipPath)
    LogLine ("ZIP bundle created: {0}" -f $zipPath)
  }
}
LogLine ("Done. Log written: {0}" -f $script:LogPath)
