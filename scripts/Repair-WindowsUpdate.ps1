<#
.SYNOPSIS
    Resets Windows Update components.

.DESCRIPTION
    Stops Windows Update services, clears the SoftwareDistribution folder,
    and restarts services.

.EXAMPLE
    .\Repair-WindowsUpdate.ps1 -Verbose
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

$services = @('wuauserv','bits','cryptsvc')

foreach ($svc in $services) {
    Write-Verbose "Stopping service '$svc'"
    Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
}

$dist = "C:\Windows\SoftwareDistribution"

if (Test-Path $dist) {
    Write-Verbose "Removing SoftwareDistribution folder..."
    if ($PSCmdlet.ShouldProcess($dist, "Clear Windows Update cache")) {
        Remove-Item $dist -Recurse -Force -ErrorAction SilentlyContinue
    }
}

foreach ($svc in $services) {
    Write-Verbose "Starting service '$svc'"
    Start-Service -Name $svc -ErrorAction SilentlyContinue
}

Write-Host "Windows Update repair completed." -ForegroundColor Green
