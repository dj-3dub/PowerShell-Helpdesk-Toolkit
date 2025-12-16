<#
.SYNOPSIS
    Resets the Outlook profile for the current user.

.DESCRIPTION
    Closes Outlook, backs up existing Outlook profile registry keys,
    removes them, and launches Outlook to rebuild the profile.

.EXAMPLE
    .\Reset-OutlookProfile.ps1 -Verbose

.NOTES
    Safe to run. Outlook will recreate a clean profile on next launch.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

Write-Verbose "Closing Outlook..."
Get-Process outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
$backupRoot = "$env:USERPROFILE\OutlookProfileBackup"
$timestamp = (Get-Date -Format 'yyyyMMdd-HHmmss')
$backupPath = "$backupRoot\OutlookProfile-$timestamp.reg"

if (Test-Path $regPath) {
    Write-Verbose "Backing up Outlook profile registry key to $backupPath"

    if ($PSCmdlet.ShouldProcess("Outlook Profile", "Backup + Remove")) {
        New-Item -Path $backupRoot -ItemType Directory -Force | Out-Null
        reg export "HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles" $backupPath /y | Out-Null

        Write-Verbose "Deleting Outlook profile registry key..."
        Remove-Item -Path $regPath -Recurse -Force
    }
}
else {
    Write-Verbose "No Outlook profile found to reset."
}

Write-Host "Outlook profile reset complete. Outlook will rebuild the profile on next launch." -ForegroundColor Green
