[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Restart
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load common utilities
. (Join-Path $PSScriptRoot 'Common.ps1')

$teamsPaths = @(
    "$env:APPDATA\Microsoft\Teams",
    "$env:LOCALAPPDATA\Microsoft\Teams",
    "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"
)

Write-Verbose "Closing Teams..."
try {
    Get-Process Teams -ErrorAction SilentlyContinue | Stop-Process -Force
    
    $clearedPaths = @()
    foreach ($path in $teamsPaths) {
        if (Test-Path $path) {
            Write-Verbose "Removing Teams cache folder '$path'"
            if ($PSCmdlet.ShouldProcess($path, "Remove folder")) {
                Remove-Item $path -Recurse -Force
                $clearedPaths += $path
            }
        }
    }

    if ($Restart) {
        Write-Verbose "Restarting Teams..."
        $launcher = "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe"
        if (Test-Path $launcher) {
            Start-Process $launcher "--processStart 'Teams.exe'"
        } else {
            Write-Warning "Could not find Teams launcher at $launcher to restart."
        }
    }

    $msg = "Teams cache reset complete."
    if ($clearedPaths) {
        $msg += " Cleared paths: " + ($clearedPaths -join ", ")
    }
    Write-AuditLog -Severity INFO -Action "ResetTeamsCache" -Message $msg
    Write-Host "Teams cache reset complete." -ForegroundColor Green
}
catch {
    Write-AuditLog -Severity ERROR -Action "ResetTeamsCache" -Message "Failed to reset Microsoft Teams cache." -Details $_.Exception.Message
    Write-Error "Failed to reset Microsoft Teams cache: $($_.Exception.Message)"
}
