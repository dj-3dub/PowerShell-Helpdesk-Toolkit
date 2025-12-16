<#
.SYNOPSIS
    Resets or cleans browser profiles for Chrome and Edge.

.DESCRIPTION
    Automates common browser troubleshooting actions:

      - Stops running browser processes
      - Optionally backs up browser profile data
      - Either:
          * Clears cache-only (less disruptive), or
          * Renames the entire "User Data" folder to force a fresh profile

    Currently supports:
      - Google Chrome
      - Microsoft Edge

.PARAMETER Browser
    Browser to target: Chrome, Edge, or All (default: All).

.PARAMETER ClearCacheOnly
    If specified, only cache-related folders are cleared. Profile data
    (bookmarks, extensions, passwords) is left intact.

.PARAMETER Backup
    If specified, a ZIP backup of the browser's User Data folder is
    created under out\BrowserBackups in the repo root.

.PARAMETER Restart
    If specified, restarts the browser after reset/cleanup.

.EXAMPLE
    .\Reset-BrowserProfile.ps1 -Browser Chrome -ClearCacheOnly -Verbose

.EXAMPLE
    .\Reset-BrowserProfile.ps1 -Browser All -Backup -Verbose

.EXAMPLE
    .\Reset-BrowserProfile.ps1 -Browser Edge -Restart -Verbose

.NOTES
    Run as the affected user. For best results, close all browser windows
    before running, or allow the script to stop browser processes.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Chrome','Edge','All')]
    [string]$Browser = 'All',

    [Parameter(Mandatory = $false)]
    [switch]$ClearCacheOnly,

    [Parameter(Mandatory = $false)]
    [switch]$Backup,

    [Parameter(Mandatory = $false)]
    [switch]$Restart
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Repo root + output folder (string-based, WSL/UNC-safe) ---
$repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$outputRoot = Join-Path $repoRoot 'out\BrowserBackups'
$null = New-Item -Path $outputRoot -ItemType Directory -Force -ErrorAction SilentlyContinue

Write-Verbose "Repo root    : $repoRoot"
Write-Verbose "Backup folder: $outputRoot"

# --- Helper functions ---

function Stop-BrowserProcess {
    param(
        [Parameter(Mandatory = $true)][string[]]$ProcessNames
    )

    foreach ($name in $ProcessNames) {
        Write-Verbose "Stopping process '$name' if running..."
        Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

function Backup-BrowserProfile {
    param(
        [Parameter(Mandatory = $true)][string]$BrowserName,
        [Parameter(Mandatory = $true)][string]$UserDataPath
    )

    if (-not (Test-Path $UserDataPath)) {
        Write-Verbose "No user data folder for $BrowserName at '$UserDataPath'; skipping backup."
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $zipName   = "{0}_{1}_{2}_{3}.zip" -f $BrowserName, $env:USERNAME, $env:COMPUTERNAME, $timestamp
    $zipPath   = Join-Path $outputRoot $zipName

    Write-Verbose "Backing up $BrowserName profile from '$UserDataPath' to '$zipPath'"

    if ($PSCmdlet.ShouldProcess($UserDataPath, "Backup $BrowserName profile to ZIP")) {
        if (Test-Path $zipPath) {
            Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        }

        Compress-Archive -Path (Join-Path $UserDataPath '*') -DestinationPath $zipPath -ErrorAction Stop
        Write-Host "Backup created: $zipPath" -ForegroundColor Green
    }
}

function Clear-BrowserCacheFolders {
    param(
        [Parameter(Mandatory = $true)][string]$BrowserName,
        [Parameter(Mandatory = $true)][string]$UserDataPath
    )

    if (-not (Test-Path $UserDataPath)) {
        Write-Verbose "User data path for $BrowserName does not exist: '$UserDataPath'"
        return
    }

    # For Chrome/Edge, cache folders exist under profile folders like "Default" and "Profile *"
    $profilePaths = Get-ChildItem -Path $UserDataPath -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -eq 'Default' -or $_.Name -like 'Profile *' }

    if (-not $profilePaths) {
        Write-Verbose "No profile folders found under '$UserDataPath' for $BrowserName."
        return
    }

    $cacheRelative = @(
        'Cache',
        'Code Cache',
        'GPUCache',
        'Service Worker\CacheStorage',
        'Service Worker\ScriptCache'
    )

    foreach ($profile in $profilePaths) {
        Write-Verbose "Processing profile '$($profile.Name)' for $BrowserName..."

        foreach ($rel in $cacheRelative) {
            $target = Join-Path $profile.FullName $rel
            if (Test-Path $target) {
                Write-Verbose "Clearing cache path: '$target'"
                if ($PSCmdlet.ShouldProcess($target, "Remove cache folder")) {
                    Remove-Item $target -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    Write-Host "$BrowserName cache cleared (profiles: $($profilePaths.Name -join ', '))." -ForegroundColor Green
}

function Reset-BrowserUserData {
    param(
        [Parameter(Mandatory = $true)][string]$BrowserName,
        [Parameter(Mandatory = $true)][string]$UserDataPath
    )

    if (-not (Test-Path $UserDataPath)) {
        Write-Verbose "User data path for $BrowserName does not exist: '$UserDataPath'"
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $newName   = "User Data.bak-$timestamp"
    $parent    = Split-Path -Path $UserDataPath -Parent
    $backupPath = Join-Path $parent $newName

    Write-Verbose "Renaming '$UserDataPath' to '$backupPath' to reset $BrowserName profile."

    if ($PSCmdlet.ShouldProcess($UserDataPath, "Reset $BrowserName profile (rename to $newName)")) {
        Rename-Item -Path $UserDataPath -NewName $newName -ErrorAction Stop
        Write-Host "$BrowserName profile reset. Original data preserved at '$backupPath'." -ForegroundColor Green
    }
}

function Restart-Browser {
    param(
        [Parameter(Mandatory = $true)][string]$BrowserName
    )

    switch ($BrowserName) {
        'Chrome' {
            $exe = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
            if (-not (Test-Path $exe)) {
                $exe = "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
            }
        }
        'Edge' {
            $exe = "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"
            if (-not (Test-Path $exe)) {
                $exe = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
            }
        }
        default { return }
    }

    if (Test-Path $exe) {
        Write-Verbose "Starting $BrowserName from '$exe'"
        Start-Process -FilePath $exe -ErrorAction SilentlyContinue
    }
    else {
        Write-Warning "Could not find $BrowserName executable to restart."
    }
}

# --- Main logic ---

$targetBrowsers = switch ($Browser) {
    'Chrome' { @('Chrome') }
    'Edge'   { @('Edge') }
    'All'    { @('Chrome','Edge') }
}

foreach ($b in $targetBrowsers) {
    switch ($b) {
        'Chrome' {
            $name          = 'Chrome'
            $processNames  = @('chrome')
            $userDataPath  = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
        }
        'Edge' {
            $name          = 'Edge'
            $processNames  = @('msedge')
            $userDataPath  = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data'
        }
        default { continue }
    }

    Write-Host "=== Processing $name ===" -ForegroundColor Cyan
    Write-Verbose "User data path for $name: $userDataPath"

    Stop-BrowserProcess -ProcessNames $processNames

    if ($Backup) {
        Backup-BrowserProfile -BrowserName $name -UserDataPath $userDataPath
    }

    if ($ClearCacheOnly) {
        Clear-BrowserCacheFolders -BrowserName $name -UserDataPath $userDataPath
    }
    else {
        Reset-BrowserUserData -BrowserName $name -UserDataPath $userDataPath
    }

    if ($Restart) {
        Restart-Browser -BrowserName $name
    }

    Write-Host ""
}

Write-Host "Browser reset/cleanup operations completed." -ForegroundColor Green
