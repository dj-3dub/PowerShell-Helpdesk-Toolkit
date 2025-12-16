<#
.SYNOPSIS
    Clears Microsoft Teams cache for the current user.

.DESCRIPTION
    Closes Teams, removes cached files, and restarts Teams if desired.

.EXAMPLE
    .\Reset-TeamsCache.ps1 -Verbose

.NOTES
    Safe to run. Teams will recreate all cache folders on launch.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Restart
)

$teamsPaths = @(
    "$env:APPDATA\Microsoft\Teams",
    "$env:LOCALAPPDATA\Microsoft\Teams",
    "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"
)

Write-Verbose "Closing Teams..."
Get-Process Teams -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

foreach ($path in $teamsPaths) {
    if (Test-Path $path) {
        Write-Verbose "Removing Teams cache folder '$path'"
        if ($PSCmdlet.ShouldProcess($path, "Remove folder")) {
            Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

if ($Restart) {
    Write-Verbose "Restarting Teams..."
    Start-Process "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe" "--processStart 'Teams.exe'"
}

Write-Host "Teams cache reset complete." -ForegroundColor Green
