<#
.SYNOPSIS
  Endpoint "What changed?" delta report (best effort, read-only).

.DESCRIPTION
  Helps answer: "It was fine yesterday - what changed?"
  Collects common change signals over the last N hours, including:
    - Windows Update events + hotfixes
    - MSI installs/uninstalls
    - AppX deployment events (Store / MSIX)
    - Driver / device install events
    - Service installs (SCM 7045)
    - Network profile changes
    - Current proxy snapshot (WinINET + WinHTTP)

  Output:
    out\HelpdeskLogs\EndpointDelta_<COMPUTER>_<timestamp>[_Ticket].txt

  Optional ZIP:
    out\HelpdeskLogs\EndpointDelta_<COMPUTER>_<timestamp>[_Ticket].zip

  This script does NOT modify system configuration.

.PARAMETER TicketId
  Optional ticket identifier used in file naming.

.PARAMETER Hours
  Lookback window in hours. Default: 24.

.PARAMETER MaxEvents
  Max events to pull per section. Default: 800.

.PARAMETER CreateZipBundle
  If set, writes raw artifacts to a work folder and zips them.

.EXAMPLE
  .\Get-EndpointDelta.ps1 -TicketId INC12345 -Hours 24 -CreateZipBundle
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

# Paths
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
New-Dir $logsRoot

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('EndpointDelta', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }

$script:LogPath = Join-Path $logsRoot (($baseName -join '_') + '.txt')
$script:WorkDir = Join-Path $logsRoot (($baseName -join '_') + '_WORK')
New-Dir $script:WorkDir

$start = (Get-Date).AddHours(-1 * [Math]::Abs($Hours))

Section "Endpoint delta report (what changed)"
LogLine ("Host: {0}" -f $hostname)
LogLine ("User: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Ticket: {0}" -f $TicketId)
LogLine ("Lookback: {0} hour(s) (since {1})" -f $Hours, $start)
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

Section "Windows Update (events + hotfixes)"
$wuLines = Get-EventsText -Filter @{
  LogName   = 'System'
  ProviderName = 'Microsoft-Windows-WindowsUpdateClient'
  StartTime = $start
} -Label 'WindowsUpdateClient/System'
$wuFile = Save-Text -Name 'eventlog_windowsupdateclient_system.txt' -Lines $wuLines
LogLine ("Saved: {0}" -f $wuFile)

Try-Run "Hotfix list (best effort)" {
  $hotfix = Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending
  if ($hotfix) {
    $recent = $hotfix | Where-Object { $_.InstalledOn -ge $start } | Select-Object -First 200
    if ($recent) {
      $recent | Format-Table -AutoSize | Out-String | ForEach-Object { LogLine $_.TrimEnd() }
    } else {
      LogLine "No hotfixes show InstalledOn within the lookback window (InstalledOn can be blank/unreliable)."
    }
  } else {
    LogLine "Get-HotFix returned no results - rely on WindowsUpdateClient events above."
  }
}

Section "MSI installs/uninstalls (Application/MsiInstaller)"
$msiLines = Get-EventsText -Filter @{
  LogName      = 'Application'
  ProviderName = 'MsiInstaller'
  StartTime    = $start
} -Label 'MsiInstaller/Application'
$msiFile = Save-Text -Name 'eventlog_msiinstaller_application.txt' -Lines $msiLines
LogLine ("Saved: {0}" -f $msiFile)

Section "AppX / MSIX deployments (AppXDeploymentServer/Operational)"
$appxLines = Get-EventsText -Filter @{
  LogName   = 'Microsoft-Windows-AppXDeploymentServer/Operational'
  StartTime = $start
} -Label 'AppXDeploymentServer/Operational'
$appxFile = Save-Text -Name 'eventlog_appxdeployment_operational.txt' -Lines $appxLines
LogLine ("Saved: {0}" -f $appxFile)

Section "Driver / device installs (UserPnp + DriverFrameworks - best effort)"
$userPnpLines = Get-EventsText -Filter @{
  LogName   = 'Microsoft-Windows-UserPnp/DeviceInstall'
  StartTime = $start
} -Label 'UserPnp/DeviceInstall'
$userPnpFile = Save-Text -Name 'eventlog_userpnp_deviceinstall.txt' -Lines $userPnpLines
LogLine ("Saved: {0}" -f $userPnpFile)

$dfLines = Get-EventsText -Filter @{
  LogName   = 'Microsoft-Windows-DriverFrameworks-UserMode/Operational'
  StartTime = $start
} -Label 'DriverFrameworks-UserMode/Operational'
$dfFile = Save-Text -Name 'eventlog_driverframeworks_usermode_operational.txt' -Lines $dfLines
LogLine ("Saved: {0}" -f $dfFile)

Section "Service installs / changes (System/Service Control Manager)"
$scmLines = Get-EventsText -Filter @{
  LogName      = 'System'
  ProviderName = 'Service Control Manager'
  StartTime    = $start
} -Label 'SCM/System'
$scmFile = Save-Text -Name 'eventlog_scm_system.txt' -Lines $scmLines
LogLine ("Saved: {0}" -f $scmFile)
LogLine ""
LogLine "Tip: Event ID 7045 indicates a new service was installed."

Section "Network profile changes (NetworkProfile/Operational)"
$netLines = Get-EventsText -Filter @{
  LogName   = 'Microsoft-Windows-NetworkProfile/Operational'
  StartTime = $start
} -Label 'NetworkProfile/Operational'
$netFile = Save-Text -Name 'eventlog_networkprofile_operational.txt' -Lines $netLines
LogLine ("Saved: {0}" -f $netFile)

Section "Current proxy snapshot (often causes 'only some sites slow')"
Save-Cmd -Name 'winhttp_proxy.txt' -CommandLine 'netsh winhttp show proxy'
Save-Cmd -Name 'ipconfig_all.txt'  -CommandLine 'ipconfig /all'
Save-Cmd -Name 'route_print.txt'   -CommandLine 'route print'

Try-Run "WinINET proxy registry snapshot" {
  $p = Join-Path $script:WorkDir 'wininet_proxy_reg.txt'
  reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyEnable 2>&1 | Out-File $p -Encoding UTF8
  reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyServer 2>&1 | Out-File $p -Encoding UTF8 -Append
  reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v AutoConfigURL 2>&1 | Out-File $p -Encoding UTF8 -Append
  LogLine ("Saved: {0}" -f $p)
}

Section "Summary pointers"
LogLine "Start with these sections when troubleshooting:"
LogLine "  1) WindowsUpdateClient/System events"
LogLine "  2) MsiInstaller/Application events"
LogLine "  3) UserPnp/DeviceInstall + DriverFrameworks events"
LogLine "  4) SCM/System event ID 7045 (new service installed)"
LogLine "  5) NetworkProfile/Operational (network changed)"
LogLine "  6) Proxy snapshot (WinINET vs WinHTTP mismatch)"
LogLine ""
LogLine "If you want more depth later, we can add: scheduled tasks, startup items, Defender updates, Intune/MDM policy arrival timing."

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
