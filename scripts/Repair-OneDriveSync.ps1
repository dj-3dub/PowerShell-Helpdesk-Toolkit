<#
.SYNOPSIS
    Repairs or resets the Microsoft OneDrive sync client for the current user.

.DESCRIPTION
    Automates common OneDrive troubleshooting steps:

      - Stops OneDrive processes
      - Optionally backs up OneDrive logs
      - Runs onedrive.exe /reset to rebuild configuration
      - Optionally restarts OneDrive after reset

    This is useful for tickets like:
      - "My OneDrive isn't syncing"
      - "Files stuck on 'Processing changes'"
      - "Sync icons not updating"

.PARAMETER Mode
    Repair mode:
      - Soft  : stop/start OneDrive only
      - Reset : full onedrive.exe /reset (default)

.PARAMETER BackupLogs
    If specified, copies OneDrive logs to out\OneDriveLogs in the repo root.

.PARAMETER Restart
    If specified, attempts to restart OneDrive after repair/reset.

.EXAMPLE
    # Full reset + logs backup, no auto-restart
    .\Repair-OneDriveSync.ps1 -Mode Reset -BackupLogs -Verbose

.EXAMPLE
    # Soft repair (stop/start only)
    .\Repair-OneDriveSync.ps1 -Mode Soft -Restart -Verbose
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Soft','Reset')]
    [string]$Mode = 'Reset',

    [Parameter(Mandatory = $false)]
    [switch]$BackupLogs,

    [Parameter(Mandatory = $false)]
    [switch]$Restart
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Resolve repo root + output folders (WSL/UNC safe) ---
$repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$logsRoot = Join-Path $repoRoot 'out\OneDriveLogs'
$null = New-Item -Path $logsRoot -ItemType Directory -Force -ErrorAction SilentlyContinue

Write-Verbose "Repo root : $repoRoot"
Write-Verbose "Logs root : $logsRoot"

# --- OneDrive paths & process names ---
$procNames = @('OneDrive', 'OneDriveStandaloneUpdater')

# Per-user OneDrive install paths (user context)
$oneDrivePaths = @(
    "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
    "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe",
    "$env:ProgramFiles(x86)\Microsoft OneDrive\OneDrive.exe"
)

# --- Helper functions ---

function Stop-OneDriveProcess {
    [CmdletBinding()]
    param()

    foreach ($name in $procNames) {
        Write-Verbose "Stopping process '$name' (if running)..."
        Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

function Backup-OneDriveLogs {
    [CmdletBinding()]
    param()

    $logDir = Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\logs'
    if (-not (Test-Path $logDir)) {
        Write-Verbose "No OneDrive logs found at '$logDir'. Skipping log backup."
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $zipName   = "OneDriveLogs_{0}_{1}_{2}.zip" -f $env:USERNAME, $env:COMPUTERNAME, $timestamp
    $zipPath   = Join-Path $logsRoot $zipName

    Write-Verbose "Backing up OneDrive logs from '$logDir' to '$zipPath'"

    if ($PSCmdlet.ShouldProcess($logDir, "Backup OneDrive logs to ZIP")) {
        if (Test-Path $zipPath) {
            Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        }

        Compress-Archive -Path (Join-Path $logDir '*') -DestinationPath $zipPath -ErrorAction Stop
        Write-Host "OneDrive logs backup created: $zipPath" -ForegroundColor Green
    }
}

function Get-OneDriveExecutable {
    [CmdletBinding()]
    param()

    foreach ($p in $oneDrivePaths) {
        if (Test-Path $p) {
            Write-Verbose "Found OneDrive executable at '$p'"
            return $p
        }
    }

    Write-Warning "Could not find OneDrive.exe in standard locations."
    return $null
}

function Reset-OneDriveClient {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath
    )

    Write-Verbose "Running OneDrive reset: '$ExePath /reset'"

    if ($PSCmdlet.ShouldProcess($ExePath, "Reset OneDrive client (/reset)")) {
        Start-Process -FilePath $ExePath -ArgumentList '/reset' -ErrorAction SilentlyContinue
        Write-Host "OneDrive reset command issued. It may take up to a few minutes for the client to fully reset." -ForegroundColor Yellow
    }
}

function Start-OneDriveClient {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath
    )

    Write-Verbose "Starting OneDrive from '$ExePath'"
    Start-Process -FilePath $ExePath -ErrorAction SilentlyContinue
}

# --- Main logic ---

Write-Host "=== OneDrive Repair ($Mode) ===" -ForegroundColor Cyan

Stop-OneDriveProcess

if ($BackupLogs) {
    Backup-OneDriveLogs
}

$exe = Get-OneDriveExecutable

if ($Mode -eq 'Reset') {
    if (-not $exe) {
        Write-Warning "Skipping reset: OneDrive executable not found."
    }
    else {
        Reset-OneDriveClient -ExePath $exe
    }
}
else {
    Write-Verbose "Mode 'Soft' selected: not running /reset, just stop/start."
}

if ($Restart) {
    if ($exe) {
        Start-OneDriveClient -ExePath $exe
        Write-Host "OneDrive client restarted." -ForegroundColor Green
    }
    else {
        Write-Warning "Could not restart OneDrive: executable not found."
    }
}
else {
    Write-Host "OneDrive repair completed. You may start OneDrive manually if needed." -ForegroundColor Green
}
