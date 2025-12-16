<#
.SYNOPSIS
  Printer / Spooler diagnostics (ticket-friendly) with optional ZIP artifacts.

.DESCRIPTION
  Complements Reset-PrinterSubsystem.ps1 by collecting evidence for common issues:
    - stuck queue / corrupt spool files
    - spooler crashes / restarts
    - driver problems after updates
    - network printer port reachability
    - PrintService operational events

  Output:
    out\HelpdeskLogs\PrintDiag_<COMPUTER>_<timestamp>[_Ticket].txt
  Optional ZIP:
    out\HelpdeskLogs\PrintDiag_<COMPUTER>_<timestamp>[_Ticket].zip

.PARAMETER TicketId
  Optional ticket identifier used in output naming.

.PARAMETER PrinterName
  Optional printer name filter (substring). If omitted, collects for all printers.

.PARAMETER TestTargets
  Optional hostnames/IPs to test connectivity against (printer IPs, print servers).

.PARAMETER CreateZipBundle
  If set, zips raw artifacts and command outputs.

.EXAMPLE
  .\Test-PrintDiagnostics.ps1 -TicketId INC123 -PrinterName "HP" -CreateZipBundle

.EXAMPLE
  .\Test-PrintDiagnostics.ps1 -TestTargets printsvr01,10.10.20.55
#>

[CmdletBinding()]
param(
  [string]$TicketId,
  [string]$PrinterName,
  [string[]]$TestTargets,
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
function Save-Cmd([string]$Name, [string]$CommandLine) {
  $p = Join-Path $script:WorkDir $Name
  Try-Run $Name { cmd /c $CommandLine 2>&1 | Out-File -FilePath $p -Encoding UTF8 }
  return $p
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

# Paths
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$repoRoot  = Split-Path -Path (Split-Path -Path $scriptDir -Parent) -Parent
$logsRoot  = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'
New-Dir $logsRoot

$hostname  = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$baseName  = @('PrintDiag', $hostname, $timestamp)
if ($TicketId) { $baseName += $TicketId }

$script:LogPath = Join-Path $logsRoot (($baseName -join '_') + '.txt')
$script:WorkDir = Join-Path $logsRoot (($baseName -join '_') + '_WORK')
New-Dir $script:WorkDir

Section "Print Diagnostics"
LogLine ("Host: {0}" -f $hostname)
LogLine ("User: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
LogLine ("Time: {0}" -f (Get-Date))
LogLine ("Ticket: {0}" -f $TicketId)
LogLine ("Log: {0}" -f $script:LogPath)
LogLine ("Artifacts: {0}" -f $script:WorkDir)

Section "Spooler service"
Try-Run "Spooler status" {
  $svc = Get-Service Spooler -ErrorAction Stop
  $mode = (Get-CimInstance Win32_Service -Filter "Name='Spooler'" -ErrorAction Stop).StartMode
  LogLine ("Spooler: Status={0}  StartMode={1}" -f $svc.Status, $mode)
}
Try-Run "Spool folder" {
  $spoolPath = Join-Path $env:SystemRoot 'System32\spool\PRINTERS'
  LogLine ("Spool folder: {0}" -f $spoolPath)
  if (Test-Path $spoolPath) {
    $files = @(Get-ChildItem $spoolPath -ErrorAction SilentlyContinue)
    LogLine ("Spool files: {0}" -f $files.Count)
    foreach ($f in ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 10)) {
      LogLine ("  {0}  {1} bytes  {2}" -f $f.Name, $f.Length, $f.LastWriteTime)
    }
  } else {
    LogLine "Spool folder not found."
  }
}

Section "Printers / ports / drivers"
Try-Run "Get-Printer" {
  $printers = @(Get-Printer -ErrorAction SilentlyContinue)
  if ($PrinterName) { $printers = @($printers | Where-Object { $_.Name -like "*$PrinterName*" }) }
  if ($printers.Count -lt 1) {
    LogLine "No printers found (or no match for filter)."
  } else {
    foreach ($p in $printers) {
      LogLine ("Printer: {0}  Default={1}  Shared={2}  Driver={3}  Port={4}" -f $p.Name, $p.Default, $p.Shared, $p.DriverName, $p.PortName)
    }
  }
}
Try-Run "Get-PrinterPort" {
  $ports = @(Get-PrinterPort -ErrorAction SilentlyContinue)
  foreach ($pp in ($ports | Select-Object -First 50)) {
    $hn = $pp.PrinterHostAddress
    if (-not $hn) { $hn = $pp.Name }
    LogLine ("Port: {0}  Host={1}  Protocol={2}  SNMP={3}" -f $pp.Name, $hn, $pp.Protocol, $pp.SNMPEnabled)
  }
}
Try-Run "Driver inventory" {
  $drivers = @(Get-PrinterDriver -ErrorAction SilentlyContinue)
  if ($PrinterName) { $drivers = @($drivers | Where-Object { $_.Name -like "*$PrinterName*" }) }
  foreach ($d in ($drivers | Sort-Object Name | Select-Object -First 80)) {
    LogLine ("Driver: {0}  Version={1}  Manufacturer={2}" -f $d.Name, $d.DriverVersion, $d.Manufacturer)
  }
}

Section "Queue / jobs"
Try-Run "Print jobs" {
  $jobs = @(Get-PrintJob -PrinterName * -ErrorAction SilentlyContinue)
  if ($PrinterName) { $jobs = @($jobs | Where-Object { $_.PrinterName -like "*$PrinterName*" }) }
  if ($jobs.Count -lt 1) {
    LogLine "No print jobs found."
  } else {
    foreach ($j in ($jobs | Sort-Object SubmittedTime -Descending | Select-Object -First 80)) {
      $status = ($j.JobStatus -join ',')
      LogLine ("Job: Printer={0}  Document={1}  User={2}  Size={3}  Status={4}" -f $j.PrinterName, $j.DocumentName, $j.Submitter, $j.Size, $status)
    }
  }
}

Section "Reachability tests"
Try-Run "Collect test targets" {
  $targets = @()
  if ($TestTargets) { $targets += $TestTargets }

  $ports = @(Get-PrinterPort -ErrorAction SilentlyContinue)
  foreach ($pp in $ports) {
    if ($pp.PrinterHostAddress) { $targets += $pp.PrinterHostAddress }
  }

  $targets = $targets | Where-Object { $_ } | Select-Object -Unique
  if ($targets.Count -lt 1) {
    LogLine "No targets to test. (Use -TestTargets printsvr01,10.0.0.5)"
  } else {
    foreach ($t in $targets) {
      LogLine ("Target: {0}" -f $t)
      LogLine ("  TCP 9100 (RAW): {0}" -f (Test-TcpPort -Host $t -Port 9100))
      LogLine ("  TCP 515 (LPD):  {0}" -f (Test-TcpPort -Host $t -Port 515))
      LogLine ("  TCP 631 (IPP):  {0}" -f (Test-TcpPort -Host $t -Port 631))
      LogLine ("  TCP 445 (SMB):  {0}" -f (Test-TcpPort -Host $t -Port 445))
    }
  }
}

Section "Event logs (PrintService)"
Try-Run "Export PrintService Operational" {
  $logName = 'Microsoft-Windows-PrintService/Operational'
  $p = Join-Path $script:WorkDir 'PrintService_Operational_Last200.evtx'
  wevtutil epl $logName $p /ow:true | Out-Null
  LogLine ("Exported: {0}" -f $p)
}
Try-Run "PrintService summary (last 48h)" {
  $events = @(Get-WinEvent -LogName 'Microsoft-Windows-PrintService/Operational' -ErrorAction SilentlyContinue |
      Where-Object { $_.TimeCreated -gt (Get-Date).AddHours(-48) } |
      Select-Object -First 100)
  LogLine ("Events (last 48h): {0}" -f $events.Count)
  foreach ($e in $events) {
    $msg = ($e.Message -replace '\s+',' ')
    if ($msg.Length -gt 220) { $msg = $msg.Substring(0,220) + '...' }
    LogLine ("[{0}] Id={1} Level={2} Msg={3}" -f $e.TimeCreated, $e.Id, $e.LevelDisplayName, $msg)
  }
}

Section "Raw artifacts"
Save-Cmd -Name 'ipconfig_all.txt' -CommandLine 'ipconfig /all' | Out-Null
Save-Cmd -Name 'net_print_jobs.txt' -CommandLine 'wmic printjob list full' | Out-Null
Save-Cmd -Name 'pnputil_enum_drivers.txt' -CommandLine 'pnputil /enum-drivers' | Out-Null

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
