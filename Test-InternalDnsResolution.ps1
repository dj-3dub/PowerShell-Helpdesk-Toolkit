<#
.SYNOPSIS
  VPN connected but nothing resolves - internal DNS diagnostics.

.DESCRIPTION
  Collects evidence for "VPN connected but internal names don't resolve":
    - Adapter list and DNS servers per interface
    - DNS suffix/search list
    - DNS cache summary
    - Optional tests for internal hostnames and/or forced DNS server tests
    - Optional safe remediation: ipconfig /flushdns

  Writes a timestamped log to: out\HelpdeskLogs\InternalDNS_<COMPUTER>_<timestamp>[_Ticket].txt

.PARAMETER TicketId
  Optional ticket identifier used in output naming.

.PARAMETER InternalNames
  One or more internal hostnames to resolve (e.g., fileserver01, intranet.contoso.com).

.PARAMETER DnsServers
  One or more DNS servers to query directly (bypasses normal resolver path).

.PARAMETER EnableSafeRemediation
  If set, performs ipconfig /flushdns.

.EXAMPLE
  .\Test-InternalDnsResolution.ps1 -InternalNames fileserver01,intranet.contoso.com -EnableSafeRemediation

.EXAMPLE
  .\Test-InternalDnsResolution.ps1 -DnsServers 10.0.0.10,10.0.0.11 -InternalNames fileserver01
#>

[CmdletBinding()]
param(
  [string]$TicketId,
  [string[]]$InternalNames,
  [string[]]$DnsServers,
  [switch]$EnableSafeRemediation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function New-Dir([string]$Path){ if(-not (Test-Path $Path)){ New-Item -Path $Path -ItemType Directory -Force | Out-Null } }
function Log([string]$Text){ if($null -ne $Text){ $Text | Out-File -FilePath $script:LogPath -Encoding UTF8 -Append; Write-Host $Text } }
function Section([string]$Title){
  $line = ('='*70)
  Log ""
  Log $line
  Log ("== {0}" -f $Title)
  Log $line
}
function TryRun([string]$Label,[scriptblock]$Block){
  try { & $Block } catch { Log ("ERROR: {0}: {1}" -f $Label, $_.Exception.Message) }
}

# Paths
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
New-Dir $logsRoot

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('InternalDNS', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }

$script:LogPath = Join-Path $logsRoot (($baseName -join '_') + '.txt')

Section "Internal DNS diagnostics (VPN connected but nothing resolves)"
Log ("Host: {0}" -f $hostname)
Log ("User: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
Log ("Time: {0}" -f (Get-Date))
Log ("Ticket: {0}" -f $TicketId)
Log ("Log: {0}" -f $script:LogPath)

Section "Adapters and DNS servers"
TryRun "Adapter inventory" {
  $adapters = @(Get-NetAdapter -ErrorAction SilentlyContinue | Sort-Object -Property Status,Name)
  foreach($a in $adapters){
    Log ("Adapter: {0}  Status={1}  Link={2}  IfIndex={3}  Desc={4}" -f $a.Name, $a.Status, $a.LinkSpeed, $a.ifIndex, $a.InterfaceDescription)
  }
}
TryRun "DNS per interface" {
  $dns = @(Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue)
  foreach($d in $dns){
    $servers = ($d.ServerAddresses -join ', ')
    Log ("IfIndex {0}  InterfaceAlias={1}  DNSServers={2}" -f $d.InterfaceIndex, $d.InterfaceAlias, $servers)
  }
}

Section "Suffix search list / NRPT"
TryRun "Suffix search list" {
  $g = Get-DnsClientGlobalSetting -ErrorAction SilentlyContinue
  if ($g) {
    Log ("SuffixSearchList: {0}" -f ($g.SuffixSearchList -join ', '))
    Log ("UseSuffixSearchList: {0}" -f $g.UseSuffixSearchList)
  } else {
    Log "Get-DnsClientGlobalSetting not available."
  }
}
TryRun "NRPT rules" {
  if (Get-Command Get-DnsClientNrptPolicy -ErrorAction SilentlyContinue) {
    $rules = @(Get-DnsClientNrptPolicy -ErrorAction SilentlyContinue)
    Log ("NRPT rules: {0}" -f $rules.Count)
    foreach($r in ($rules | Select-Object -First 30)){
      Log ("Rule: Namespace={0}  NameServers={1}" -f $r.Namespace, ($r.NameServers -join ', '))
    }
  } else {
    Log "Get-DnsClientNrptPolicy not available on this system."
  }
}

Section "DNS cache"
TryRun "Cache count" {
  if (Get-Command Get-DnsClientCache -ErrorAction SilentlyContinue) {
    $c = @(Get-DnsClientCache -ErrorAction SilentlyContinue)
    Log ("DNS cache entries: {0}" -f $c.Count)
  } else {
    Log "Get-DnsClientCache not available."
  }
}

Section "Resolution tests"
if (-not $InternalNames -or $InternalNames.Count -lt 1) {
  Log "No internal names provided. Re-run with -InternalNames fileserver01,intranet.contoso.com"
} else {
  foreach($n in $InternalNames){
    if ([string]::IsNullOrWhiteSpace($n)) { continue }
    $name = $n.Trim()

    TryRun ("Resolve-DnsName {0}" -f $name) {
      $r = Resolve-DnsName -Name $name -ErrorAction Stop
      $ips = @($r | Where-Object { $_.IPAddress } | Select-Object -ExpandProperty IPAddress)
      Log ("Name: {0}  IPs: {1}" -f $name, ($ips -join ', '))
    }

    if ($DnsServers -and $DnsServers.Count -gt 0) {
      foreach($s in $DnsServers){
        $server = $s.Trim()
        if (-not $server) { continue }
        TryRun ("Resolve via server {0}" -f $server) {
          $r2 = Resolve-DnsName -Name $name -Server $server -ErrorAction Stop
          $ips2 = @($r2 | Where-Object { $_.IPAddress } | Select-Object -ExpandProperty IPAddress)
          Log ("Name: {0}  Server: {1}  IPs: {2}" -f $name, $server, ($ips2 -join ', '))
        }
      }
    }
  }
}

Section "Safe remediation"
if ($EnableSafeRemediation) {
  TryRun "ipconfig /flushdns" { ipconfig /flushdns | ForEach-Object { Log $_ } }
  Log "Remediation complete: flushed DNS cache."
} else {
  Log "Skipped. Use -EnableSafeRemediation to flush DNS cache."
}

Section "Useful command outputs"
TryRun "ipconfig /all" { ipconfig /all | Out-File -FilePath (Join-Path $logsRoot (($baseName -join '_') + '_ipconfig.txt')) -Encoding UTF8 }
TryRun "route print" { route print | Out-File -FilePath (Join-Path $logsRoot (($baseName -join '_') + '_route.txt')) -Encoding UTF8 }

Log ""
Log ("Done. Log written: {0}" -f $script:LogPath)
