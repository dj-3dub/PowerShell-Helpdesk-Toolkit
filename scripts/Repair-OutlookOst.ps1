<#
.SYNOPSIS
    Repairs Outlook OST issues by scanning and/or rebuilding OST files.

.DESCRIPTION
    Common Outlook fixes for slowness, sync issues, or corruption:

      - Locates OST files under the current user's profile
      - Optionally filters OSTs by Identity (e.g. UPN/email substring)
      - Optionally backs up OST files before changes
      - Runs scanpst.exe (Inbox Repair Tool) against OST files (Scan mode)
      - Renames/deletes OST files to force Outlook to rebuild them (Rebuild mode)

    Modes:
      - Rebuild (default) : backup (optional) then rename OST to .bak-<timestamp>
      - Scan              : run scanpst.exe only, no deletion
      - Both              : run scanpst.exe, then rebuild (rename)

.PARAMETER Identity
    Optional string to filter OST file names. For example, pass part of the user's
    UPN or email address. If omitted, all OST files under the Outlook data path
    are processed.

.PARAMETER Mode
    Repair mode:
      - Rebuild : rename/delete OST so Outlook rebuilds it (default)
      - Scan    : run scanpst.exe against OST (no rebuild)
      - Both    : scan first, then rebuild

.PARAMETER Backup
    If specified, backs up OST files into out\OutlookOstRepair\backups before
    scanning/rebuilding.

.EXAMPLE
    # Default: rebuild all OSTs (with backup)
    .\Repair-OutlookOst.ps1 -Backup -Verbose

.EXAMPLE
    # Rebuild OSTs containing 'user@contoso.com' in the name
    .\Repair-OutlookOst.ps1 -Identity 'user@contoso.com' -Backup -Verbose

.EXAMPLE
    # Scan only, no rebuild
    .\Repair-OutlookOst.ps1 -Mode Scan -Verbose

.EXAMPLE
    # Scan, then rebuild, for OSTs matching identity
    .\Repair-OutlookOst.ps1 -Identity 'user@contoso.com' -Mode Both -Backup -Verbose
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$Identity,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Rebuild','Scan','Both')]
    [string]$Mode = 'Rebuild',

    [Parameter(Mandatory = $false)]
    [switch]$Backup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Verbose "Starting Outlook OST repair. Mode=$Mode, Identity='$Identity', Backup=$Backup"

# --- Repo root + output folders (WSL/UNC safe) ---
$repoRoot   = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$outputRoot = Join-Path $repoRoot 'out\OutlookOstRepair'
$sessionTs  = Get-Date -Format 'yyyyMMdd-HHmmss'
$sessionRoot = Join-Path $outputRoot ("Session_{0}_{1}" -f $env:COMPUTERNAME, $sessionTs)

$backupsRoot = Join-Path $sessionRoot 'backups'
$logsRoot    = Join-Path $sessionRoot 'logs'

$null = New-Item -Path $outputRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
$null = New-Item -Path $sessionRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
$null = New-Item -Path $backupsRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
$null = New-Item -Path $logsRoot    -ItemType Directory -Force -ErrorAction SilentlyContinue

Write-Verbose "Repo root    : $repoRoot"
Write-Verbose "Session root : $sessionRoot"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message
    )
    $logPath = Join-Path $logsRoot 'Repair-OutlookOst.log'
    $entry = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    $entry | Out-File -FilePath $logPath -Encoding UTF8 -Append
}

Write-Log "=== Outlook OST repair started (Mode=$Mode, Identity='$Identity', Backup=$Backup) ==="

# --- Helper: find scanpst.exe (Inbox Repair Tool) ---
function Get-ScanPstPath {
    [CmdletBinding()]
    param()

    $candidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\scanpst.exe",
        "$env:ProgramFiles(x86)\Microsoft Office\root\Office16\scanpst.exe",
        "$env:ProgramFiles\Microsoft Office\Office16\scanpst.exe",
        "$env:ProgramFiles(x86)\Microsoft Office\Office16\scanpst.exe",
        "$env:ProgramFiles\Microsoft Office\root\Office15\scanpst.exe",
        "$env:ProgramFiles(x86)\Microsoft Office\root\Office15\scanpst.exe",
        "$env:ProgramFiles\Microsoft Office\Office15\scanpst.exe",
        "$env:ProgramFiles(x86)\Microsoft Office\Office15\scanpst.exe"
    )

    foreach ($c in $candidates) {
        if (Test-Path $c) {
            Write-Verbose "Found scanpst.exe at '$c'"
            return $c
        }
    }

    # Fallback: search under ProgramFiles for scanpst.exe (can be slow but thorough)
    $roots = @(
        "$env:ProgramFiles\Microsoft Office",
        "$env:ProgramFiles(x86)\Microsoft Office"
    ) | Where-Object { $_ -and (Test-Path $_) }

    foreach ($r in $roots) {
        try {
            $found = Get-ChildItem -Path $r -Recurse -Filter 'scanpst.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($found) {
                Write-Verbose "Found scanpst.exe at '$($found.FullName)' via search."
                return $found.FullName
            }
        }
        catch {
            Write-Warning "Error searching for scanpst.exe under '$r': $($_.Exception.Message)"
        }
    }

    return $null
}

$doScan    = $Mode -in @('Scan','Both')
$doRebuild = $Mode -in @('Rebuild','Both')

$scanPstPath = $null
if ($doScan) {
    $scanPstPath = Get-ScanPstPath
    if (-not $scanPstPath) {
        Write-Warning "scanpst.exe not found. Scan mode not available."
        Write-Log "scanpst.exe not found; Scan mode unavailable."

        if ($Mode -eq 'Scan') {
            throw "scanpst.exe not found and Mode=Scan. Cannot continue."
        }
        elseif ($Mode -eq 'Both') {
            Write-Warning "Falling back to Rebuild only."
            Write-Log "Falling back to Rebuild mode only."
            $doScan    = $false
            $doRebuild = $true
        }
    }
}

# --- Locate OST files ---
$outlookDataPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
Write-Verbose "Outlook data path: $outlookDataPath"

if (-not (Test-Path $outlookDataPath)) {
    Write-Warning "Outlook data path not found at '$outlookDataPath'. No OST files to process."
    Write-Log "Outlook data path not found. Exiting."
    return
}

$allOsts = Get-ChildItem -Path $outlookDataPath -Filter '*.ost' -File -Recurse -ErrorAction SilentlyContinue

if (-not $allOsts) {
    Write-Warning "No OST files found under '$outlookDataPath'."
    Write-Log "No OST files found. Exiting."
    return
}

$targetOsts = $allOsts
if ($Identity) {
    $targetOsts = $allOsts | Where-Object {
        $_.Name -like "*$Identity*" -or $_.FullName -like "*$Identity*"
    }
}

if (-not $targetOsts -or $targetOsts.Count -eq 0) {
    Write-Warning "No OST files matched Identity filter '$Identity'."
    Write-Log "No OST files matched Identity '$Identity'. Exiting."
    return
}

Write-Host "=== OST files to process ===" -ForegroundColor Cyan
$targetOsts | ForEach-Object {
    Write-Host (" - {0}" -f $_.FullName)
}
Write-Log ("OST files to process:`n" + ($targetOsts.FullName -join "`n"))

# --- Process each OST ---
foreach ($ost in $targetOsts) {
    Write-Host ""
    Write-Host ("Processing OST: {0}" -f $ost.FullName) -ForegroundColor Cyan
    Write-Log ("Processing OST: {0}" -f $ost.FullName)

    $currentPath = $ost.FullName

    # Backup
    if ($Backup) {
        $backupName = "{0}_{1}.bak" -f $ost.Name, $sessionTs
        $backupPath = Join-Path $backupsRoot $backupName

        Write-Verbose "Backing up OST to '$backupPath'"
        Write-Log ("Backing up to: {0}" -f $backupPath)

        if ($PSCmdlet.ShouldProcess($currentPath, "Backup OST to $backupPath")) {
            Copy-Item -Path $currentPath -Destination $backupPath -Force -ErrorAction SilentlyContinue
        }
    }

    # Scan via scanpst.exe
    if ($doScan -and $scanPstPath) {
        Write-Verbose "Running scanpst.exe against '$currentPath'"
        Write-Log ("Running scanpst.exe: '{0}'" -f $currentPath)

        if ($PSCmdlet.ShouldProcess($currentPath, "Scan OST using scanpst.exe")) {
            try {
                $scanArgs = '"' + $currentPath + '"'
                $proc = Start-Process -FilePath $scanPstPath -ArgumentList $scanArgs -PassThru -ErrorAction SilentlyContinue
                if ($proc) {
                    $proc.WaitForExit()
                    Write-Host ("Scan completed with exit code {0}." -f $proc.ExitCode)
                    Write-Log ("scanpst.exe exit code: {0}" -f $proc.ExitCode)
                }
                else {
                    Write-Warning "Failed to start scanpst.exe."
                    Write-Log "Failed to start scanpst.exe."
                }
            }
            catch {
                Write-Warning "Error running scanpst.exe: $($_.Exception.Message)"
                Write-Log ("Error running scanpst.exe: {0}" -f $_.Exception.Message)
            }
        }
    }

    # Rebuild (rename OST so Outlook rebuilds it)
    if ($doRebuild) {
        $newName  = "{0}.rebuilt-{1}" -f $ost.Name, $sessionTs
        $newPath  = Join-Path $ost.DirectoryName $newName

        Write-Verbose "Renaming OST to '$newPath' to force rebuild."
        Write-Log ("Renaming OST to: {0}" -f $newPath)

        if ($PSCmdlet.ShouldProcess($currentPath, "Rename OST to $newPath (force rebuild)")) {
            try {
                Rename-Item -Path $currentPath -NewName $newName -ErrorAction Stop
                Write-Host ("OST renamed to '{0}'. Outlook will rebuild it on next launch." -f $newPath) -ForegroundColor Green
                Write-Log ("OST renamed successfully to: {0}" -f $newPath)
            }
            catch {
                Write-Warning "Failed to rename OST: $($_.Exception.Message)"
                Write-Log ("Failed to rename OST: {0}" -f $_.Exception.Message)
            }
        }
    }
}

Write-Host ""
Write-Host "Outlook OST repair complete. Logs at: $logsRoot" -ForegroundColor Green
Write-Log "Outlook OST repair completed."

