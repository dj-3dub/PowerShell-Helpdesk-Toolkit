<#
.SYNOPSIS
  Startup / Scheduled Task / Service "delta" report (read-only).

.DESCRIPTION
  Answers: "What is auto-starting now?" / "Why is logon suddenly slow?"
  Collects high-signal change indicators over the last N hours:
    - Scheduled Tasks modified/created (based on task file LastWriteTime)
    - Scheduled Task failures (TaskScheduler/Operational)
    - Startup registry keys (HKLM/HKCU Run/RunOnce)
    - Startup folders (common + per-user) recent file changes
    - Service install/change events (SCM/System) + current Auto services snapshot

  Output:
    out\HelpdeskLogs\StartupDelta_<COMPUTER>_<timestamp>[_Ticket].txt

  Optional ZIP:
    out\HelpdeskLogs\StartupDelta_<COMPUTER>_<timestamp>[_Ticket].zip

  This script does NOT modify system configuration.

.PARAMETER TicketId
  Optional ticket identifier used in file naming.

.PARAMETER Hours
  Lookback window in hours. Default: 24.

.PARAMETER MaxEvents
  Max events to pull per log section. Default: 800.

.PARAMETER CreateZipBundle
  If set, writes raw artifacts to a work folder and zips them.

.EXAMPLE
  .\Get-StartupDelta.ps1 -TicketId INC12345 -Hours 24 -CreateZipBundle

.EXAMPLE
  .\Get-StartupDelta.ps1 -Hours 6
#>

[CmdletBinding()]
param(
  [string]$TicketId,
  [int]$Hours = 24,
  [int]$MaxEvents = 800,
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
  Try-Run $Name {
    cmd /c $CommandLine 2>&1 | Out-File -FilePath $p -Encoding UTF8
  }
  return $p
}

function Get-EventsText {
  param(
    [Parameter(Mandatory)][hashtable]$Filter,
    [Parameter(Mandatory)][string]$Label
  )
  try {
    $events = Get-WinEvent -FilterHashtable $Filter -ErrorAction Stop |
      Sort-Object TimeCreated -Descending |
      Select-Object -First $MaxEvents

    $lines = @()
    foreach ($e in $events) {
      $lines += ("[{0}] ID={1} Level={2} Provider={3}" -f $e.TimeCreated, $e.Id, $e.LevelDisplayName, $e.ProviderName)
      $msg = ($e.Message -replace "`r","") -split "`n"
      foreach ($m in $msg) { $lines += ("  {0}" -f $m.TrimEnd()) }
      $lines += ""
    }
    if (-not $lines) { $lines = @("No matching events found for: $Label") }
    return $lines
  } catch {
    return @("ERROR reading events for $Label : $($_.Exception.Message)")
  }
}

function Dump-RunKey([string]$Label, [string]$Path) {
  LogLine ("{0}: {1}" -f $Label, $Path)
  try {
    $props = Get-ItemProperty -Path $Path -ErrorAction Stop
    $names = $props.PSObject.Properties |
      Where-Object { $_.Name -notin @('PSPath','PSParentPath','PSChildName','PSDrive','PSProvider') } |
      Select-Object -ExpandProperty Name

    if (-not $names) {
      LogLine "  (No values)"
      return
    }

    foreach ($n in $names | Sort-Object) {
      $v = $props.$n
      LogLine ("  {0} = {1}" -f $n, $v)
    }
  } catch {
    LogLine ("  (Not found / access denied): {0}" -f $_.Exception.Message)
  }
}

# Paths
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
New-Dir $logsRoot

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('StartupDelta', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }

$script:LogPath = Join-Path $logsRoot (($baseName -join '_') + '.txt')
$script:WorkDir = Join-Path $logsRoot (($baseName -join '_') + '_WORK')
New-Dir $script:WorkDir

$start = (Get-Date).AddHours(-1 * [Math]::Abs($Hours))

Section "Startup delta report"
LogLine ("Host: {0}" -f $hostname)
LogLine ("User: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Ticket: {0}" -f $TicketId)
LogLine ("Lookback: {0} hour(s) (since {1})" -f $Hours, $start)
LogLine ("Log: {0}" -f $script:LogPath)
LogLine ("Artifacts: {0}" -f $script:WorkDir)

Section "Scheduled Tasks changed recently (file LastWriteTime)"
Try-Run "Task file inventory" {
  $tasksRoot = Join-Path $env:WINDIR 'System32\Tasks'
  if (-not (Test-Path $tasksRoot)) {
    LogLine ("Tasks folder not found: {0}" -f $tasksRoot)
  } else {
    $changed = @(
      Get-ChildItem -Path $tasksRoot -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $start } |
        Sort-Object LastWriteTime -Descending
    )

    $out = @()
    foreach ($f in $changed) {
      $rel = $f.FullName.Substring($tasksRoot.Length).TrimStart('\')
      $out += ("{0}  LastWriteTime={1}" -f $rel, $f.LastWriteTime)
    }

    if ($out.Count -lt 1) { $out = @("No task files modified within the lookback window.") }
    $p = Save-Text -Name 'scheduledtasks_changed_files.txt' -Lines $out
    LogLine ("Saved: {0}" -f $p)
  }
}

Section "Task Scheduler failures (Operational)"
$tsLines = Get-EventsText -Filter @{
  LogName   = 'Microsoft-Windows-TaskScheduler/Operational'
  StartTime = $start
} -Label 'TaskScheduler/Operational'
$tsFile = Save-Text -Name 'eventlog_taskscheduler_operational.txt' -Lines $tsLines
LogLine ("Saved: {0}" -f $tsFile)

Section "Startup registry keys (Run / RunOnce)"
Dump-RunKey -Label 'HKLM Run'     -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'
Dump-RunKey -Label 'HKLM RunOnce' -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce'
Dump-RunKey -Label 'HKCU Run'     -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
Dump-RunKey -Label 'HKCU RunOnce' -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce'

Section "Startup folders - recent file changes"
Try-Run "Startup folders" {
  $commonStartup = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup'
  $userStartup   = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'

  foreach ($p in @($commonStartup, $userStartup)) {
    LogLine ("Folder: {0}" -f $p)
    if (-not (Test-Path $p)) {
      LogLine "  (Not found)"
      continue
    }

    $recent = @(
      Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $start } |
        Sort-Object LastWriteTime -Descending
    )

    if ($recent.Count -lt 1) {
      LogLine "  No recent changes in lookback window."
    } else {
      foreach ($f in $recent) {
        LogLine ("  {0}  LastWriteTime={1}" -f $f.Name, $f.LastWriteTime)
      }
    }
  }
}

Section "Service installs/changes (SCM/System) + Auto service snapshot"
$scmLines = Get-EventsText -Filter @{
  LogName      = 'System'
  ProviderName = 'Service Control Manager'
  StartTime    = $start
} -Label 'SCM/System'
$scmFile = Save-Text -Name 'eventlog_scm_system.txt' -Lines $scmLines
LogLine ("Saved: {0}" -f $scmFile)
LogLine ""
LogLine "Tip: Event ID 7045 indicates a new service was installed."

Try-Run "Auto-start services snapshot" {
  $svc = Get-CimInstance Win32_Service |
    Where-Object { $_.StartMode -eq 'Auto' } |
    Select-Object Name, DisplayName, State, StartMode, PathName |
    Sort-Object DisplayName

  $p = Join-Path $script:WorkDir 'services_autostart_snapshot.txt'
  $svc | Format-Table -AutoSize | Out-String | Out-File -FilePath $p -Encoding UTF8
  LogLine ("Saved: {0}" -f $p)
}

Section "Summary pointers"
LogLine "Start here:"
LogLine "  1) scheduledtasks_changed_files.txt (new/changed scheduled tasks)"
LogLine "  2) TaskScheduler/Operational errors in eventlog_taskscheduler_operational.txt"
LogLine "  3) HKLM/HKCU Run and RunOnce values (unexpected executables)"
LogLine "  4) Startup folder recent changes"
LogLine "  5) SCM/System event ID 7045 and new services"
LogLine ""
LogLine "If you want, we can extend this to include: Startup Approved entries, browser extension changes, WMI permanent event subscriptions."

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
