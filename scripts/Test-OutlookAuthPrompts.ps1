<#
.SYNOPSIS
  Outlook auth prompt diagnostics (M365) - gathers evidence for repeated sign-in prompts.

.DESCRIPTION
  Collects common signal areas without making changes:
    - AAD / device registration status (dsregcmd)
    - Time sync summary (w32tm)
    - Credential Manager (cmdkey list - redacted-ish)
    - Office / Identity registry hints (WAM/ADAL keys existence)
    - Network reachability + DNS resolution to common M365 endpoints
    - Proxy and WinHTTP proxy settings

  Writes a timestamped log to: out\HelpdeskLogs\OutlookAuth_<COMPUTER>_<timestamp>[_Ticket].txt

.PARAMETER TicketId
  Optional ticket identifier used in output naming.

.PARAMETER TenantHint
  Optional tenant hint (e.g., contoso.com) to include in report notes.

.PARAMETER IncludeCmdKey
  Include `cmdkey /list` output (can contain target names; review before sharing). Default: off.

.EXAMPLE
  .\Test-OutlookAuthPrompts.ps1 -TicketId INC123 -IncludeCmdKey

#>

[CmdletBinding()]
param(
  [string]$TicketId,
  [string]$TenantHint,
  [switch]$IncludeCmdKey
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
function SaveCmd([string]$Name,[string]$CmdLine){
  $p = Join-Path $script:WorkDir $Name
  TryRun $Name { cmd /c $CmdLine 2>&1 | Out-File -FilePath $p -Encoding UTF8 }
  return $p
}
function TestEndpoint([string]$TargetHost,[int]$Port=443){
  $dns = $false
  $tcp = $false
  try { Resolve-DnsName -Name $TargetHost -ErrorAction Stop | Out-Null; $dns = $true } catch { $dns = $false }
  try { $tcp = [bool](Test-NetConnection -ComputerName $TargetHost -Port $Port -WarningAction SilentlyContinue).TcpTestSucceeded } catch { $tcp = $false }
  [pscustomobject]@{ Host=$TargetHost; Port=$Port; Dns=$dns; Tcp=$tcp }
}

# Paths
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
New-Dir $logsRoot

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('OutlookAuth', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }

$script:LogPath = Join-Path $logsRoot (($baseName -join '_') + '.txt')
$script:WorkDir = Join-Path $logsRoot (($baseName -join '_') + '_WORK')
New-Dir $script:WorkDir

Section "Outlook auth prompts - diagnostics"
Log ("Host: {0}" -f $hostname)
Log ("User: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
TryRun "UPN" { Log ("UPN: {0}" -f (whoami /upn)) }
Log ("Time: {0}" -f (Get-Date))
Log ("Ticket: {0}" -f $TicketId)
Log ("TenantHint: {0}" -f $TenantHint)
Log ("Log: {0}" -f $script:LogPath)

Section "Quick checks (proxy / time)"
TryRun "Proxy (WinINet)" { SaveCmd "proxy_netsh_winhttp.txt" "netsh winhttp show proxy" | Out-Null }
TryRun "Proxy (Internet Settings)" { 
  $p = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue
  if ($p) {
    Log ("ProxyEnable: {0}" -f $p.ProxyEnable)
    Log ("ProxyServer: {0}" -f $p.ProxyServer)
    Log ("AutoConfigURL: {0}" -f $p.AutoConfigURL)
  } else { Log "No Internet Settings key found (HKCU)." }
}
TryRun "Time status" { SaveCmd "w32tm_status.txt" "w32tm /query /status" | Out-Null }

Section "Device registration (dsregcmd)"
TryRun "dsregcmd status" { SaveCmd "dsregcmd_status.txt" "dsregcmd /status" | Out-Null }

Section "Credential artifacts"
if ($IncludeCmdKey) {
  TryRun "cmdkey list" { SaveCmd "cmdkey_list.txt" "cmdkey /list" | Out-Null }
} else {
  Log "cmdkey /list skipped (use -IncludeCmdKey to include)."
}

Section "Office / Identity registry hints"
TryRun "WAM/ADAL keys presence" {
  $paths = @(
    'HKCU:\Software\Microsoft\Office\16.0\Common\Identity',
    'HKCU:\Software\Microsoft\Office\16.0\Common\Internet',
    'HKCU:\Software\Microsoft\Office\16.0\Common\SignIn',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\AAD',
    'HKCU:\Software\Microsoft\IdentityCRL',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI'
  )
  foreach($pp in $paths){
    $exists = Test-Path $pp
    Log ("{0} : {1}" -f $pp, $exists)
    if($exists){
      try{
        $val = Get-ItemProperty -Path $pp -ErrorAction Stop
        foreach($n in ($val.PSObject.Properties | Select-Object -ExpandProperty Name)){
          if($n -match 'Token|Refresh|Access|Secret|Password'){ continue } # avoid obvious sensitive names
        }
      } catch {}
    }
  }
}

Section "M365 endpoint reachability"
$endpoints = @(
  'outlook.office365.com',
  'autodiscover-s.outlook.com',
  'login.microsoftonline.com',
  'device.login.microsoftonline.com',
  'graph.microsoft.com',
  'officeclient.microsoft.com'
)
$results = foreach($h in $endpoints){ TestEndpoint -TargetHost $h -Port 443 }
foreach($r in $results){
  Log ("{0}: DNS={1} TCP443={2}" -f $r.Host, $r.Dns, $r.Tcp)
}

Section "Network basics"
TryRun "IP config" { SaveCmd "ipconfig_all.txt" "ipconfig /all" | Out-Null }
TryRun "Routes" { SaveCmd "route_print.txt" "route print" | Out-Null }
TryRun "NCSI" { SaveCmd "ncsi_registry.txt" "reg query ""HKLM\SOFTWARE\Policies\Microsoft\Windows\NetworkConnectivityStatusIndicator"" /s" | Out-Null }

Section "Suggested next actions (human-readable)"
Log "- If VPN is in use, re-test auth prompts OFF VPN and ON VPN."
Log "- If DNS to login.microsoftonline.com or autodiscover fails, investigate proxy/DNS filtering."
Log "- If dsregcmd shows 'AzureAdJoined: NO' but should be joined, check device registration policy."
Log "- If time skew is large, fix time sync first (Kerberos and token validation can break)."
Log ""
Log ("Done. Log written: {0}" -f $script:LogPath)
