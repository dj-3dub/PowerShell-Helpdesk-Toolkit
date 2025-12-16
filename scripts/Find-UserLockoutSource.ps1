<#
.SYNOPSIS
  Identify likely account lockout source(s) for a user (best-effort).

.DESCRIPTION
  Queries local Windows Security event logs for lockout-related events and summarizes
  the most likely workstation(s)/IP(s) involved.

  Primary signal:
    - Event ID 4740 (A user account was locked out) [Domain Controller logs]

  Secondary signals (best-effort):
    - Event ID 4625 (failed logon) filtered by the target username [Client/DC logs]

  Notes:
    - For best results, run this on a Domain Controller (4740) or a log collector that
      receives DC Security logs.
    - Reading the Security log often requires an elevated PowerShell session.
    - This script is READ-ONLY by default (no changes made).

.PARAMETER Identity
  Username (samAccountName) or UPN/email. (UPN will be reduced to the left side for matching.)

.PARAMETER Hours
  How far back to search event logs. Default: 24.

.PARAMETER MaxEvents
  Maximum number of events to retrieve per event ID. Default: 2000.

.PARAMETER TicketId
  Optional ticket identifier used in log naming.

.EXAMPLE
  .\Find-UserLockoutSource.ps1 -Identity jdoe -Hours 12

.EXAMPLE
  .\Find-UserLockoutSource.ps1 -Identity jdoe@contoso.com -TicketId INC12345
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$Identity,

  [int]$Hours = 24,

  [int]$MaxEvents = 2000,

  [string]$TicketId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ShortUser {
  param([string]$Id)
  if ([string]::IsNullOrWhiteSpace($Id)) { return $Id }
  if ($Id -match '\\') { return ($Id.Split('\\')[-1]) }
  if ($Id -match '@')  { return ($Id.Split('@')[0]) }
  return $Id
}

$shortUser = Get-ShortUser -Id $Identity

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
if (-not (Test-Path $logsRoot)) { New-Item -Path $logsRoot -ItemType Directory -Force | Out-Null }

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('LockoutSource', $hostname, $timestamp)
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

function Try-GetWinEvent {
  param(
    [hashtable]$Filter,
    [int]$Take = 2000
  )
  try {
    return Get-WinEvent -FilterHashtable $Filter -ErrorAction Stop | Select-Object -First $Take
  } catch {
    LogLine ("ERROR reading Security log (EventId {0}): {1}" -f $Filter.Id, $_.Exception.Message)
    LogLine "Tip: Run PowerShell as Administrator and/or run on a DC that has the Security events."
    return @()
  }
}

function Parse-EventXml {
  param([System.Diagnostics.Eventing.Reader.EventRecord]$Event)
  $xml = [xml]$Event.ToXml()
  $data = @{}
  foreach ($d in $xml.Event.EventData.Data) {
    $name = $d.Name
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    $data[$name] = [string]$d.'#text'
  }
  return $data
}

Section "Account lockout source detection"
LogLine ("Identity: {0} (match key: {1})" -f $Identity, $shortUser)
LogLine ("Host: {0}" -f $hostname)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Lookback: {0} hour(s)" -f $Hours)
LogLine ("Log: {0}" -f $logPath)

$start = (Get-Date).AddHours(-1 * [Math]::Abs($Hours))

Section "Event ID 4740 (lockout) - strongest signal when present"
$events4740 = Try-GetWinEvent -Filter @{ LogName='Security'; Id=4740; StartTime=$start } -Take $MaxEvents

$lockoutHits = @()
foreach ($e in $events4740) {
  $data = Parse-EventXml -Event $e
  $target = $data['TargetUserName']
  if ([string]::IsNullOrWhiteSpace($target)) { continue }

  if ($target -ieq $shortUser) {
    $lockoutHits += [pscustomobject]@{
      TimeCreated        = $e.TimeCreated
      TargetUserName     = $target
      TargetDomainName   = $data['TargetDomainName']
      CallerComputerName = $data['CallerComputerName']
      EventRecordId      = $e.RecordId
    }
  }
}

if (-not $lockoutHits -or $lockoutHits.Count -eq 0) {
  LogLine "No 4740 lockout events found for this user in the selected window."
  LogLine "If you're not running on a DC (or DC logs aren't forwarded here), that's expected."
} else {
  LogLine ("Found {0} matching 4740 event(s)." -f $lockoutHits.Count)

  $grouped = $lockoutHits | Group-Object -Property CallerComputerName | Sort-Object -Property Count -Descending
  LogLine ""
  LogLine "Top CallerComputerName values:"
  foreach ($g in $grouped) {
    $name = if ($g.Name) { $g.Name } else { "<blank>" }
    LogLine ("  {0,-35}  Count: {1}" -f $name, $g.Count)
  }

  LogLine ""
  LogLine "Recent lockouts:"
  $lockoutHits |
    Sort-Object TimeCreated -Descending |
    Select-Object -First 10 |
    ForEach-Object {
      LogLine ("  {0} | CallerComputerName: {1} | RecordId: {2}" -f $_.TimeCreated, $_.CallerComputerName, $_.EventRecordId)
    }

  LogLine ""
  LogLine "Recommendation:"
  LogLine "  - Check the top CallerComputerName system(s) for cached credentials:"
  LogLine "      * old mapped drives, scheduled tasks, services, Outlook/Teams re-auth loops, mobile mail, VPN clients"
  LogLine "  - If the caller is blank, use 4625 failed logon events below for more clues."
}

Section "Event ID 4625 (failed logon) - best effort (may be noisy)"
$events4625 = Try-GetWinEvent -Filter @{ LogName='Security'; Id=4625; StartTime=$start } -Take $MaxEvents

$failedHits = @()
foreach ($e in $events4625) {
  $data = Parse-EventXml -Event $e
  $target = $data['TargetUserName']
  if ([string]::IsNullOrWhiteSpace($target)) { continue }

  if ($target -ieq $shortUser) {
    $failedHits += [pscustomobject]@{
      TimeCreated      = $e.TimeCreated
      TargetUserName   = $target
      TargetDomainName = $data['TargetDomainName']
      WorkstationName  = $data['WorkstationName']
      IpAddress        = $data['IpAddress']
      LogonType        = $data['LogonType']
      Status           = $data['Status']
      SubStatus        = $data['SubStatus']
      AuthPackage      = $data['AuthenticationPackageName']
      EventRecordId    = $e.RecordId
    }
  }
}

if (-not $failedHits -or $failedHits.Count -eq 0) {
  LogLine "No 4625 failed logon events found for this user in the selected window (on this host)."
} else {
  LogLine ("Found {0} matching 4625 event(s)." -f $failedHits.Count)

  LogLine ""
  LogLine "Top WorkstationName values:"
  $failedHits |
    Group-Object -Property WorkstationName |
    Sort-Object Count -Descending |
    Select-Object -First 10 |
    ForEach-Object {
      $name = if ($_.Name) { $_.Name } else { "<blank>" }
      LogLine ("  {0,-35}  Count: {1}" -f $name, $_.Count)
    }

  LogLine ""
  LogLine "Top IpAddress values:"
  $failedHits |
    Group-Object -Property IpAddress |
    Sort-Object Count -Descending |
    Select-Object -First 10 |
    ForEach-Object {
      $name = if ($_.Name) { $_.Name } else { "<blank>" }
      LogLine ("  {0,-35}  Count: {1}" -f $name, $_.Count)
    }

  LogLine ""
  LogLine "Recent failures (top 10):"
  $failedHits |
    Sort-Object TimeCreated -Descending |
    Select-Object -First 10 |
    ForEach-Object {
      LogLine ("  {0} | WS: {1} | IP: {2} | LogonType: {3} | Status: {4}/{5}" -f $_.TimeCreated, $_.WorkstationName, $_.IpAddress, $_.LogonType, $_.Status, $_.SubStatus)
    }

  LogLine ""
  LogLine "Interpretation hints:"
  LogLine "  - LogonType 3 often indicates network auth (mapped drives/services)."
  LogLine "  - LogonType 10 often indicates RemoteInteractive (RDP)."
  LogLine "  - If IP is present, it can point to VPN, Wi-Fi, or a specific device."
}

Section "Complete"
LogLine "Done."
LogLine ("Log written: {0}" -f $logPath)
