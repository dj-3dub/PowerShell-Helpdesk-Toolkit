<#
.SYNOPSIS
  Logon + Group Policy performance diagnostics (read-only, ticket-friendly).

.DESCRIPTION
  Collects the most common signals used to troubleshoot slow logons in enterprise Windows:
    - Group Policy Operational events (last N hours) to identify slow/failed CSE extensions
    - DC/Site discovery (LOGONSERVER, nltest /dsgetsite, nltest /dsgetdc)
    - Current GPO result snapshot (gpresult /r + optional /h)
    - Network/proxy quick context (optional)
    - User profile size hint + redirected folders (best effort)

  Outputs:
    - Human-readable log: out\HelpdeskLogs\LogonGpoDiag_<COMPUTER>_<timestamp>[_Ticket].txt
    - Optional ZIP bundle containing raw outputs + gpresult HTML.

  This script does NOT run gpupdate, does NOT change policy, and does NOT modify registry settings.

.PARAMETER TicketId
  Optional ticket identifier used in log naming.

.PARAMETER Hours
  How far back to pull event logs. Default: 24.

.PARAMETER MaxEvents
  Max events to pull from each log. Default: 1200.

.PARAMETER IncludeGpResultHtml
  If set, generates gpresult HTML (can take ~10-30s). Saved in bundle and referenced from the log.

.PARAMETER CreateZipBundle
  If set, creates a ZIP bundle with collected artifacts.

.EXAMPLE
  .\Test-LogonGpoPerformance.ps1 -TicketId INC12345 -IncludeGpResultHtml -CreateZipBundle

.EXAMPLE
  .\Test-LogonGpoPerformance.ps1 -Hours 6
#>

[CmdletBinding()]
param(
  [string]$TicketId,
  [int]$Hours = 24,
  [int]$MaxEvents = 1200,
  [switch]$IncludeGpResultHtml,
  [switch]$CreateZipBundle
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function New-Dir([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}

function LogLine {
  param([string]$Text)
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

function Save-Cmd([string]$FileName, [scriptblock]$Cmd) {
  $p = Join-Path $script:WorkDir $FileName
  Try-Run $FileName {
    & $Cmd | Out-File -FilePath $p -Encoding UTF8
  }
}

function Try-GetWinEventText([string]$LogName, [int[]]$Ids) {
  $start = (Get-Date).AddHours(-1 * [Math]::Abs($Hours))
  $filter = @{
    LogName   = $LogName
    StartTime = $start
  }
  if ($Ids -and $Ids.Count -gt 0) { $filter.Id = $Ids }

  try {
    $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop | Select-Object -First $MaxEvents
    $lines = @()
    foreach ($e in $events) {
      $lines += ("[{0}] ID={1} Level={2} Provider={3}" -f $e.TimeCreated, $e.Id, $e.LevelDisplayName, $e.ProviderName)
      $msg = ($e.Message -replace "`r","") -split "`n"
      foreach ($m in $msg) { $lines += ("  {0}" -f $m.TrimEnd()) }
      $lines += ""
    }
    return $lines
  } catch {
    return @("ERROR reading $LogName : $($_.Exception.Message)")
  }
}

# Paths
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
New-Dir $logsRoot

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('LogonGpoDiag', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }

$script:LogPath = Join-Path $logsRoot (($baseName -join '_') + '.txt')
$script:WorkDir = Join-Path $logsRoot (($baseName -join '_') + '_WORK')
New-Dir $script:WorkDir

Section "Logon + Group Policy performance diagnostics"
LogLine ("Host: {0}" -f $hostname)
LogLine ("User: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Ticket: {0}" -f $TicketId)
LogLine ("Lookback: {0} hour(s)" -f $Hours)
LogLine ("Log: {0}" -f $script:LogPath)
LogLine ("Artifacts: {0}" -f $script:WorkDir)

Section "System context"
Try-Run "OS and uptime" {
  $os = Get-CimInstance Win32_OperatingSystem
  $uptime = (Get-Date) - $os.LastBootUpTime
  LogLine ("OS: {0} ({1})" -f $os.Caption, $os.Version)
  LogLine ("Build: {0}" -f $os.BuildNumber)
  LogLine ("Boot: {0}" -f $os.LastBootUpTime)
  LogLine ("Uptime: {0:dd\.hh\:mm\:ss}" -f $uptime)
}

Try-Run "Domain join" {
  $cs = Get-CimInstance Win32_ComputerSystem
  LogLine ("Domain: {0}" -f $cs.Domain)
  LogLine ("PartOfDomain: {0}" -f $cs.PartOfDomain)
}

Section "DC / Site discovery (common slow logon root cause)"
Try-Run "LOGONSERVER" {
  LogLine ("LOGONSERVER: {0}" -f $env:LOGONSERVER)
}
Save-Cmd "nltest_dsgetsite.txt" { cmd /c "nltest /dsgetsite 2>&1" }
Save-Cmd "nltest_dsgetdc.txt"   { cmd /c "nltest /dsgetdc:%USERDOMAIN% 2>&1" }
Save-Cmd "ipconfig_all.txt"     { cmd /c "ipconfig /all 2>&1" }
Save-Cmd "route_print.txt"      { cmd /c "route print 2>&1" }

Section "Group Policy results snapshot"
Save-Cmd "gpresult_r.txt" { cmd /c "gpresult /r 2>&1" }

if ($IncludeGpResultHtml) {
  $htmlPath = Join-Path $script:WorkDir "gpresult.html"
  Try-Run "gpresult /h" {
    cmd /c "gpresult /h ""$htmlPath"" /f 2>&1" | Out-File (Join-Path $script:WorkDir "gpresult_h_output.txt") -Encoding UTF8
    if (Test-Path $htmlPath) {
      LogLine ("gpresult HTML generated: {0}" -f $htmlPath)
    } else {
      LogLine "gpresult HTML was requested but not found (see gpresult_h_output.txt)."
    }
  }
}

Section "Group Policy Operational log (look for slow CSE/extensions)"
$gpLines = Try-GetWinEventText -LogName "Microsoft-Windows-GroupPolicy/Operational" -Ids @()
$gpOut = Join-Path $script:WorkDir "eventlog_grouppolicy_operational.txt"
$gpLines | Out-File -FilePath $gpOut -Encoding UTF8
LogLine ("Saved: {0}" -f $gpOut)

LogLine ""
LogLine "Quick guidance when reviewing GroupPolicy/Operational:"
LogLine "  - Look for repeated retries, timeouts, or long processing durations."
LogLine "  - Slow logons are often caused by: drive mapping, scripts, printers, folder redirection, security filtering, or unreachable DC/VPN."

Section "User Profile hints"
Try-Run "Profile path + size hint" {
  $p = [Environment]::GetFolderPath('UserProfile')
  LogLine ("UserProfile: {0}" -f $p)
  try {
    $size = (Get-ChildItem -Path $p -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    if ($size) { LogLine ("Approx profile size: {0:N2} GB" -f ($size / 1GB)) }
  } catch {
    LogLine "Could not compute profile size (access denied or long paths)."
  }
}

Section "Mapped drives / SMB mappings"
Try-Run "Get-SmbMapping" {
  $m = Get-SmbMapping -ErrorAction SilentlyContinue
  if ($m) { $m | Format-Table -AutoSize | Out-String | ForEach-Object { LogLine $_.TrimEnd() } }
  else { LogLine "No SMB mappings detected (or cmdlet not available)." }
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

LogLine "Done."
LogLine ("Log written: {0}" -f $script:LogPath)
