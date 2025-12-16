<#
.SYNOPSIS
    Checks BitLocker health and protection status on the local machine.

.DESCRIPTION
    Enumerates BitLocker volumes and reports:

      - Volume type and mount point
      - Protection status (On/Off)
      - Volume status (FullyEncrypted, EncryptionInProgress, etc.)
      - Encryption percentage and capacity (GB)
      - Encryption method
      - Auto-unlock status
      - Key protector types (TPM, RecoveryPassword, etc.)

    Optionally exports results to CSV and (optionally) includes recovery
    passwords in the CSV (off by default; use with care).

.PARAMETER ExportCsv
    If specified, exports a CSV report under exports\BitLocker in the repo root.

.PARAMETER IncludeRecoveryKeys
    If specified, includes recovery passwords in the CSV. Use cautiously
    and protect the resulting file. By default, recovery keys are omitted.

.PARAMETER OutputPath
    Optional explicit CSV path. If not supplied and -ExportCsv is used,
    the report is written to exports\BitLocker\Test-BitLockerHealth-<timestamp>.csv.

.EXAMPLE
    .\Test-BitLockerHealth.ps1 -Verbose

.EXAMPLE
    .\Test-BitLockerHealth.ps1 -ExportCsv -Verbose

.EXAMPLE
    .\Test-BitLockerHealth.ps1 -ExportCsv -IncludeRecoveryKeys -Verbose
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeRecoveryKeys,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Ensure BitLocker cmdlets exist
$blCmd = Get-Command -Name Get-BitLockerVolume -ErrorAction SilentlyContinue
if (-not $blCmd) {
    throw "Get-BitLockerVolume not found. This script requires the BitLocker module (Windows 8+/Server 2012+)."
}

# repo root + export folder (string-based, WSL/UNC safe)
$repoRoot   = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$exportRoot = Join-Path $repoRoot 'exports\BitLocker'
$null = New-Item -Path $exportRoot -ItemType Directory -Force -ErrorAction SilentlyContinue

Write-Verbose "Repo root  : $repoRoot"
Write-Verbose "ExportRoot : $exportRoot"

$results = New-Object System.Collections.Generic.List[object]

# Try to gather TPM info (if available)
$tpmInfo = $null
try {
    $tpmCmd = Get-Command -Name Get-Tpm -ErrorAction SilentlyContinue
    if ($tpmCmd) {
        $tpmInfo = Get-Tpm -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Warning "Get-Tpm failed: $($_.Exception.Message)"
}

if ($tpmInfo) {
    Write-Host "=== TPM Info ===" -ForegroundColor Cyan
    Write-Host ("  TPM Present     : {0}" -f $tpmInfo.TpmPresent)
    Write-Host ("  TPM Ready       : {0}" -f $tpmInfo.TpmReady)
    Write-Host ("  AutoProvisioning: {0}" -f $tpmInfo.AutoProvisioning)
    Write-Host ""
}

# Enumerate BitLocker volumes
$volumes = Get-BitLockerVolume
if (-not $volumes) {
    Write-Warning "No BitLocker volumes found on this system."
    return
}

Write-Host "=== BitLocker Volumes ===" -ForegroundColor Cyan

foreach ($vol in $volumes) {
    $mountPoint   = ($vol.MountPoint -join ', ')
    $volumeType   = $vol.VolumeType
    $protStatus   = $vol.ProtectionStatus
    $volStatus    = $vol.VolumeStatus
    $encPercent   = $vol.EncryptionPercentage
    $encMethod    = $vol.EncryptionMethod
    $capacityGB   = $null

    # Capacity from CapacityGB (Win 11/Server 2022+ sometimes exposes) or via WMI as fallback
    if ($vol | Get-Member -Name CapacityGB -ErrorAction SilentlyContinue) {
        $capacityGB = [math]::Round($vol.CapacityGB, 2)
    }
    else {
        try {
            if ($mountPoint) {
                $driveLetter = $mountPoint.Split(':')[0]
                $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID = '$driveLetter:`\'" -ErrorAction SilentlyContinue
                if ($disk.Size) {
                    $capacityGB = [math]::Round(($disk.Size / 1GB), 2)
                }
            }
        }
        catch { }
    }

    $autoUnlock   = $vol.AutoUnlockEnabled

    # Key protector summary
    $kpTypes = @()
    $recoveryKeys = @()

    foreach ($kp in $vol.KeyProtector) {
        if ($kp.KeyProtectorType) {
            $kpTypes += $kp.KeyProtectorType
        }
        if ($kp.KeyProtectorType -eq 'RecoveryPassword' -and $kp.RecoveryPassword) {
            $recoveryKeys += $kp.RecoveryPassword
        }
    }

    $kpTypesString = ($kpTypes | Sort-Object -Unique) -join '; '

    $obj = [pscustomobject]@{
        ComputerName         = $env:COMPUTERNAME
        MountPoint           = $mountPoint
        VolumeType           = $volumeType
        ProtectionStatus     = $protStatus
        VolumeStatus         = $volStatus
        EncryptionPercentage = $encPercent
        CapacityGB           = $capacityGB
        EncryptionMethod     = $encMethod
        AutoUnlockEnabled    = $autoUnlock
        KeyProtectorTypes    = $kpTypesString
        RecoveryPasswords    = if ($IncludeRecoveryKeys) { $recoveryKeys -join '; ' } else { $null }
    }

    $results.Add($obj)

    # On-screen summary
    Write-Host ("Volume: {0} [{1}]" -f $mountPoint, $volumeType) -ForegroundColor Cyan
    Write-Host ("  ProtectionStatus     : {0}" -f $protStatus)
    Write-Host ("  VolumeStatus         : {0}" -f $volStatus)
    Write-Host ("  Encryption           : {0}% ({1})" -f $encPercent, $encMethod)
    Write-Host ("  Capacity             : {0} GB" -f $capacityGB)
    Write-Host ("  AutoUnlockEnabled    : {0}" -f $autoUnlock)
    Write-Host ("  KeyProtectorTypes    : {0}" -f $kpTypesString)

    if ($IncludeRecoveryKeys -and $recoveryKeys.Count -gt 0) {
        Write-Host ("  RecoveryPasswords    : {0}" -f ($recoveryKeys -join '; '))
    }

    Write-Host ""
}

if ($ExportCsv -and $results.Count -gt 0) {
    if (-not $OutputPath) {
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $OutputPath = Join-Path $exportRoot "Test-BitLockerHealth-$timestamp.csv"
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "BitLocker health report exported to: $OutputPath" -ForegroundColor Green
}
else {
    Write-Host "BitLocker health check completed." -ForegroundColor Green
}
