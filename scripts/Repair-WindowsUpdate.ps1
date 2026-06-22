[CmdletBinding(SupportsShouldProcess = $true)]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load common utilities
. (Join-Path $PSScriptRoot 'Common.ps1')

$services = @('wuauserv','bits','cryptsvc')
$dist = "C:\Windows\SoftwareDistribution"

try {
    if ($PSCmdlet.ShouldProcess("Windows Update Components", "Reset cache and restart services")) {
        foreach ($svc in $services) {
            Write-Verbose "Stopping service '$svc'..."
            if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
                Stop-Service -Name $svc -Force
            }
        }

        if (Test-Path $dist) {
            Write-Verbose "Removing SoftwareDistribution folder '$dist'..."
            Remove-Item $dist -Recurse -Force
        }

        foreach ($svc in $services) {
            Write-Verbose "Starting service '$svc'..."
            if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
                Start-Service -Name $svc
            }
        }

        Write-AuditLog -Severity INFO -Action "RepairWindowsUpdate" -Message "Successfully reset Windows Update components (stopped services, cleared SoftwareDistribution, restarted services)."
    }
    Write-Host "Windows Update repair completed." -ForegroundColor Green
}
catch {
    # Attempt to restart services in case of failure
    foreach ($svc in $services) {
        try {
            if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
                Start-Service -Name $svc -ErrorAction SilentlyContinue
            }
        } catch {}
    }
    
    Write-AuditLog -Severity ERROR -Action "RepairWindowsUpdate" -Message "Failed to repair Windows Update components." -Details $_.Exception.Message
    Write-Error "Failed to repair Windows Update components: $($_.Exception.Message)"
}
