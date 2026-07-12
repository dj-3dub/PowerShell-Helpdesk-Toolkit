<#
.SYNOPSIS
  Collect Wi-Fi diagnostics for slow or unstable wireless connections.

.DESCRIPTION
  Gathers common Wi-Fi troubleshooting artifacts and writes a ZIP bundle under:

    out\HelpdeskLogs\WiFiDiag_<COMPUTER>_<timestamp>[_Ticket].zip

  Includes:
    - netsh wlan show interfaces
    - netsh wlan show drivers
    - netsh wlan show profiles
    - ipconfig /all
    - route print
    - DNS client information
    - Network adapter and advanced adapter properties
    - Optional Windows WLAN report
    - WLAN AutoConfig operational events
    - Basic latency tests to the default gateway and a public target

  Optional safe remediation is limited to:
    - ipconfig /flushdns
    - ipconfig /renew
    - Restarting the WLAN AutoConfig service

  Adapter disable and enable operations are not performed.

.PARAMETER TicketId
  Optional ticket identifier used in the bundle name.

.PARAMETER IncludeWlanReport
  Generates the Windows WLAN report and includes the HTML report in the bundle.

.PARAMETER EnableSafeRemediation
  Performs limited remediation actions and then collects updated interface
  information. Administrative privileges are required.

.PARAMETER PublicTestHost
  Public ping target. The default is 8.8.8.8.

.PARAMETER EventLogHours
  Number of hours of WLAN AutoConfig events to collect.

  The default is 24 hours. Valid values are 1 through 168.

.EXAMPLE
  .\Collect-WiFiDiagnostics.ps1 `
    -TicketId INC12345 `
    -IncludeWlanReport

.EXAMPLE
  .\Collect-WiFiDiagnostics.ps1 `
    -EnableSafeRemediation

.EXAMPLE
  .\Collect-WiFiDiagnostics.ps1 `
    -TicketId INC12345 `
    -IncludeWlanReport `
    -EventLogHours 48

.EXAMPLE
  .\Collect-WiFiDiagnostics.ps1 `
    -TicketId INC12345 `
    -IncludeWlanReport `
    -EnableSafeRemediation `
    -PublicTestHost 1.1.1.1 `
    -EventLogHours 72
#>

[CmdletBinding()]
param(
  [string]$TicketId,

  [switch]$IncludeWlanReport,

  [switch]$EnableSafeRemediation,

  [ValidateNotNullOrEmpty()]
  [string]$PublicTestHost = '8.8.8.8',

  [ValidateRange(1, 168)]
  [int]$EventLogHours = 24
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($env:OS -ne 'Windows_NT') {
  throw 'Collect-WiFiDiagnostics.ps1 is supported only on Windows.'
}

$isAdministrator = $false

try {
  $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = [Security.Principal.WindowsPrincipal]::new($currentIdentity)

  $isAdministrator = $principal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
  )
}
catch {
  Write-Warning "Unable to determine administrative status: $($_.Exception.Message)"
}

function New-Directory {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [string]$Path
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item `
      -Path $Path `
      -ItemType Directory `
      -Force `
      -ErrorAction Stop |
      Out-Null
  }
}

function Write-TextFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [string]$Path,

    [Parameter(Mandatory)]
    [AllowEmptyCollection()]
    [string[]]$Lines
  )

  $Lines |
    Out-File `
      -FilePath $Path `
      -Encoding UTF8 `
      -Force
}

function Invoke-BestEffort {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [string]$Label,

    [Parameter(Mandatory)]
    [scriptblock]$ScriptBlock,

    [string]$ErrorOutputPath
  )

  try {
    & $ScriptBlock
    return $true
  }
  catch {
    $message = "ERROR: ${Label}: $($_.Exception.Message)"
    Write-Warning $message

    if ($ErrorOutputPath) {
      $message |
        Out-File `
          -FilePath $ErrorOutputPath `
          -Encoding UTF8 `
          -Append
    }

    return $false
  }
}

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$logsRoot = Join-Path -Path $repoRoot -ChildPath 'out\HelpdeskLogs'

New-Directory -Path $logsRoot

$computerName = if ($env:COMPUTERNAME) {
  $env:COMPUTERNAME
}
else {
  'UNKNOWN'
}

$domainName = if ($env:USERDOMAIN) {
  $env:USERDOMAIN
}
else {
  $computerName
}

$userName = if ($env:USERNAME) {
  $env:USERNAME
}
else {
  'UNKNOWN'
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

$baseName = @(
  'WiFiDiag'
  $computerName
  $timestamp
)

if (-not [string]::IsNullOrWhiteSpace($TicketId)) {
  $safeTicketId = $TicketId -replace '[^a-zA-Z0-9._-]', '_'
  $baseName += $safeTicketId
}

$bundleName = $baseName -join '_'
$workDir = Join-Path $logsRoot "${bundleName}_WORK"
$zipPath = Join-Path $logsRoot "${bundleName}.zip"
$errorLogPath = Join-Path $workDir 'collection_errors.txt'

New-Directory -Path $workDir

$metadata = @(
  "Host: $computerName"
  "User: $domainName\$userName"
  "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
  "Ticket: $TicketId"
  "Administrator: $isAdministrator"
  "IncludeWlanReport: $IncludeWlanReport"
  "EnableSafeRemediation: $EnableSafeRemediation"
  "PublicTestHost: $PublicTestHost"
  "EventLogHours: $EventLogHours"
)

Write-TextFile `
  -Path (Join-Path $workDir 'README.txt') `
  -Lines $metadata

$wifiAdapters = @()

Invoke-BestEffort `
  -Label 'Get-NetAdapter Wi-Fi detection' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    $wifiAdapters = @(
      Get-NetAdapter -Physical -ErrorAction Stop |
        Where-Object {
          $_.Status -ne 'Disabled' -and (
            $_.InterfaceDescription -match 'Wireless|Wi-Fi|WLAN|802\.11' -or
            $_.Name -match 'Wi-Fi|WLAN'
          )
        }
    )
  } |
  Out-Null

Write-Host 'Collecting Wi-Fi diagnostics...' -ForegroundColor Cyan

# STEP 1: Core diagnostic collection

Invoke-BestEffort `
  -Label 'netsh wlan show interfaces' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    netsh.exe wlan show interfaces 2>&1 |
      Out-File `
        -FilePath (Join-Path $workDir 'netsh_wlan_show_interfaces.txt') `
        -Encoding UTF8
  } |
  Out-Null

Invoke-BestEffort `
  -Label 'netsh wlan show drivers' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    netsh.exe wlan show drivers 2>&1 |
      Out-File `
        -FilePath (Join-Path $workDir 'netsh_wlan_show_drivers.txt') `
        -Encoding UTF8
  } |
  Out-Null

Invoke-BestEffort `
  -Label 'netsh wlan show profiles' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    netsh.exe wlan show profiles 2>&1 |
      Out-File `
        -FilePath (Join-Path $workDir 'netsh_wlan_show_profiles.txt') `
        -Encoding UTF8
  } |
  Out-Null

Invoke-BestEffort `
  -Label 'ipconfig /all' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    ipconfig.exe /all 2>&1 |
      Out-File `
        -FilePath (Join-Path $workDir 'ipconfig_all.txt') `
        -Encoding UTF8
  } |
  Out-Null

Invoke-BestEffort `
  -Label 'route print' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    route.exe print 2>&1 |
      Out-File `
        -FilePath (Join-Path $workDir 'route_print.txt') `
        -Encoding UTF8
  } |
  Out-Null

Invoke-BestEffort `
  -Label 'Get-DnsClientServerAddress' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    Get-DnsClientServerAddress `
      -AddressFamily IPv4, IPv6 `
      -ErrorAction Stop |
      Format-Table -AutoSize |
      Out-String |
      Out-File `
        -FilePath (Join-Path $workDir 'dns_client_server_address.txt') `
        -Encoding UTF8
  } |
  Out-Null

Invoke-BestEffort `
  -Label 'Get-NetAdapter' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    Get-NetAdapter -ErrorAction Stop |
      Sort-Object Name |
      Format-Table -AutoSize |
      Out-String |
      Out-File `
        -FilePath (Join-Path $workDir 'netadapter_table.txt') `
        -Encoding UTF8
  } |
  Out-Null

if ($wifiAdapters.Count -gt 0) {
  foreach ($adapter in $wifiAdapters) {
    $safeAdapterName = $adapter.Name -replace '[^a-zA-Z0-9._ -]', '_'

    Invoke-BestEffort `
      -Label "Advanced properties for $($adapter.Name)" `
      -ErrorOutputPath $errorLogPath `
      -ScriptBlock {
        Get-NetAdapterAdvancedProperty `
          -Name $adapter.Name `
          -ErrorAction Stop |
          Sort-Object DisplayName |
          Format-Table -AutoSize |
          Out-String |
          Out-File `
            -FilePath (
              Join-Path `
                $workDir `
                "netadapter_advprops_${safeAdapterName}.txt"
            ) `
            -Encoding UTF8
      } |
      Out-Null
  }
}
else {
  Write-TextFile `
    -Path (Join-Path $workDir 'wifi_adapter_status.txt') `
    -Lines @(
      'No enabled physical Wi-Fi adapter was detected.'
      'The remaining diagnostic collection was still attempted.'
    )
}

# STEP 2: WLAN report collection

if ($IncludeWlanReport) {
  Invoke-BestEffort `
    -Label 'netsh wlan show wlanreport' `
    -ErrorOutputPath $errorLogPath `
    -ScriptBlock {
      $netshCommand = Get-Command netsh.exe -ErrorAction SilentlyContinue

      if (-not $netshCommand) {
        throw 'netsh.exe was not found.'
      }

      if ([string]::IsNullOrWhiteSpace($env:ProgramData)) {
        throw 'The ProgramData environment variable is unavailable.'
      }

      $commandOutputPath = Join-Path `
        $workDir `
        'netsh_wlan_show_wlanreport_output.txt'

      $reportPath = Join-Path `
        $env:ProgramData `
        'Microsoft\Windows\WlanReport\wlan-report-latest.html'

      $destinationPath = Join-Path `
        $workDir `
        'wlan-report-latest.html'

      $commandOutput = & $netshCommand.Source wlan show wlanreport 2>&1
      $exitCode = $LASTEXITCODE

      @(
        'Command: netsh wlan show wlanreport'
        "ExitCode: $exitCode"
        ''
        $commandOutput
      ) |
        Out-File `
          -FilePath $commandOutputPath `
          -Encoding UTF8

      if ($exitCode -ne 0) {
        throw "WLAN report generation failed with exit code $exitCode."
      }

      if (-not (Test-Path -LiteralPath $reportPath)) {
        throw "The WLAN report was not found at: $reportPath"
      }

      Copy-Item `
        -LiteralPath $reportPath `
        -Destination $destinationPath `
        -Force `
        -ErrorAction Stop

      Write-TextFile `
        -Path (Join-Path $workDir 'wlan_report_status.txt') `
        -Lines @(
          'Status: Success'
          "Source: $reportPath"
          "Destination: $destinationPath"
          "Collected: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
          ''
          'Warning: The WLAN report may contain SSIDs, device details,'
          'connection history, and other potentially sensitive information.'
        )
    } |
    Out-Null
}

# STEP 3: WLAN AutoConfig event-log collection

Invoke-BestEffort `
  -Label 'WLAN AutoConfig event-log collection' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    $wlanLogName = 'Microsoft-Windows-WLAN-AutoConfig/Operational'
    $startTime = (Get-Date).AddHours(-$EventLogHours)

    $events = @(
      Get-WinEvent `
        -FilterHashtable @{
          LogName   = $wlanLogName
          StartTime = $startTime
        } `
        -ErrorAction Stop |
        Sort-Object TimeCreated
    )

    if ($events.Count -eq 0) {
      Write-TextFile `
        -Path (Join-Path $workDir 'wlan_autoconfig_events.txt') `
        -Lines @(
          "No WLAN AutoConfig events were found during the previous $EventLogHours hours."
          "Log: $wlanLogName"
          "StartTime: $startTime"
        )

      return
    }

    $events |
      Select-Object `
        TimeCreated,
        Id,
        LevelDisplayName,
        ProviderName,
        MachineName,
        Message |
      Format-List |
      Out-String -Width 240 |
      Out-File `
        -FilePath (Join-Path $workDir 'wlan_autoconfig_events.txt') `
        -Encoding UTF8

    $events |
      Select-Object `
        TimeCreated,
        Id,
        LevelDisplayName,
        ProviderName,
        MachineName,
        Message |
      Export-Csv `
        -Path (Join-Path $workDir 'wlan_autoconfig_events.csv') `
        -NoTypeInformation `
        -Encoding UTF8
  } |
  Out-Null

# STEP 4: Basic reachability tests

$defaultGateway = $null

Invoke-BestEffort `
  -Label 'Default gateway detection' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    $defaultRoute = Get-NetRoute `
      -DestinationPrefix '0.0.0.0/0' `
      -ErrorAction Stop |
      Where-Object {
        $_.NextHop -and $_.NextHop -ne '0.0.0.0'
      } |
      Sort-Object RouteMetric, InterfaceMetric |
      Select-Object -First 1

    if ($defaultRoute) {
      $defaultGateway = $defaultRoute.NextHop
    }
  } |
  Out-Null

Invoke-BestEffort `
  -Label 'Ping tests' `
  -ErrorOutputPath $errorLogPath `
  -ScriptBlock {
    $pingResults = @(
      "DefaultGateway: $defaultGateway"
      "PublicTestHost: $PublicTestHost"
      ''
    )

    if ($defaultGateway) {
      $pingResults += "ping -n 10 $defaultGateway"
      $pingResults += ping.exe -n 10 $defaultGateway 2>&1
      $pingResults += ''
    }
    else {
      $pingResults += 'Default gateway was not detected.'
      $pingResults += ''
    }

    $pingResults += "ping -n 10 $PublicTestHost"
    $pingResults += ping.exe -n 10 $PublicTestHost 2>&1
    $pingResults += ''

    Write-TextFile `
      -Path (Join-Path $workDir 'ping_tests.txt') `
      -Lines $pingResults
  } |
  Out-Null

# STEP 5: Optional safe remediation

if ($EnableSafeRemediation -and -not $isAdministrator) {
  Write-Warning 'Safe remediation was requested, but administrative privileges are required.'

  Write-TextFile `
    -Path (Join-Path $workDir 'remediation_skipped.txt') `
    -Lines @(
      'Status: Skipped'
      'Reason: Administrative privileges are required.'
      "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
    )
}
elseif ($EnableSafeRemediation) {
  Write-Host 'Running safe remediation...' -ForegroundColor Yellow

  Invoke-BestEffort `
    -Label 'ipconfig /flushdns' `
    -ErrorOutputPath $errorLogPath `
    -ScriptBlock {
      ipconfig.exe /flushdns 2>&1 |
        Out-File `
          -FilePath (Join-Path $workDir 'remed_flushdns.txt') `
          -Encoding UTF8
    } |
    Out-Null

  Invoke-BestEffort `
    -Label 'ipconfig /renew' `
    -ErrorOutputPath $errorLogPath `
    -ScriptBlock {
      ipconfig.exe /renew 2>&1 |
        Out-File `
          -FilePath (Join-Path $workDir 'remed_renew.txt') `
          -Encoding UTF8
    } |
    Out-Null

  Invoke-BestEffort `
    -Label 'Restart WlanSvc' `
    -ErrorOutputPath $errorLogPath `
    -ScriptBlock {
      $service = Get-Service `
        -Name 'WlanSvc' `
        -ErrorAction SilentlyContinue

      if (-not $service) {
        throw 'The WLAN AutoConfig service was not found.'
      }

      if ($service.Status -eq 'Running') {
        Restart-Service `
          -Name 'WlanSvc' `
          -Force `
          -ErrorAction Stop

        $serviceAction = 'Restarted'
      }
      else {
        Start-Service `
          -Name 'WlanSvc' `
          -ErrorAction Stop

        $serviceAction = 'Started'
      }

      Write-TextFile `
        -Path (Join-Path $workDir 'remed_wlansvc.txt') `
        -Lines @(
          'Service: WlanSvc'
          "Action: $serviceAction"
          "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
        )
    } |
    Out-Null

  Invoke-BestEffort `
    -Label 'Re-run netsh wlan show interfaces' `
    -ErrorOutputPath $errorLogPath `
    -ScriptBlock {
      netsh.exe wlan show interfaces 2>&1 |
        Out-File `
          -FilePath (
            Join-Path `
              $workDir `
              'netsh_wlan_show_interfaces_after.txt'
          ) `
          -Encoding UTF8
    } |
    Out-Null
}

# STEP 6: Create the diagnostic bundle

try {
  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item `
      -LiteralPath $zipPath `
      -Force `
      -ErrorAction Stop
  }

  Add-Type `
    -AssemblyName System.IO.Compression.FileSystem `
    -ErrorAction Stop

  [System.IO.Compression.ZipFile]::CreateFromDirectory(
    $workDir,
    $zipPath
  )

  if (-not (Test-Path -LiteralPath $zipPath)) {
    throw "The diagnostic ZIP was not created at: $zipPath"
  }
}
catch {
  Write-Error "Failed to create the diagnostic ZIP bundle: $($_.Exception.Message)"
  Write-Warning "The working directory was preserved at: $workDir"
  throw
}

try {
  Remove-Item `
    -LiteralPath $workDir `
    -Recurse `
    -Force `
    -ErrorAction Stop
}
catch {
  Write-Warning "The ZIP was created, but the working directory could not be removed: $workDir"
}

Write-Host ''
Write-Host 'Wi-Fi diagnostics bundle created:' -ForegroundColor Green
Write-Host "  $zipPath" -ForegroundColor Green

Get-Item -LiteralPath $zipPath
