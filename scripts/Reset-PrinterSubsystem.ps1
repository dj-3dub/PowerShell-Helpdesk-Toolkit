[CmdletBinding(SupportsShouldProcess = $true)]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load common utilities
. (Join-Path $PSScriptRoot 'Common.ps1')

$spoolService = 'Spooler'

try {
    # SystemRoot is Windows specific; default to system path if null to prevent binding errors
    $systemRoot = $env:SystemRoot
    if ([string]::IsNullOrEmpty($systemRoot)) { $systemRoot = 'C:\Windows' }
    
    $spoolPath = Join-Path $systemRoot 'System32\spool\PRINTERS'
    Write-Verbose "Target spool folder: $spoolPath"

    if (-not (Test-Path $spoolPath)) {
        Write-Verbose "Spool folder '$spoolPath' does not exist. Nothing to clear."
        Write-AuditLog -Severity WARNING -Action "ResetPrinterSubsystem" -Message "Spool folder '$spoolPath' does not exist. No reset performed."
    }
    else {
        if ($PSCmdlet.ShouldProcess($spoolPath, "Clear print queue")) {
            Write-Verbose "Stopping Print Spooler service..."
            Stop-Service -Name $spoolService -Force

            Write-Verbose "Removing files from '$spoolPath'..."
            $files = Get-ChildItem -Path $spoolPath
            if ($files) {
                $files | Remove-Item -Force
            }

            Write-Verbose "Starting Print Spooler service..."
            Start-Service -Name $spoolService
            
            Write-AuditLog -Severity INFO -Action "ResetPrinterSubsystem" -Message "Successfully reset printer subsystem and cleared print queue."
        }
    }
    Write-Host "Printer subsystem reset complete. Try printing again." -ForegroundColor Green
}
catch {
    # Ensure Spooler service is left started in case of failures during folder clearing
    try { 
        if (Get-Service -Name $spoolService -ErrorAction SilentlyContinue) {
            Start-Service -Name $spoolService -ErrorAction SilentlyContinue 
        }
    } catch {}
    
    Write-AuditLog -Severity ERROR -Action "ResetPrinterSubsystem" -Message "Failed to reset printer subsystem." -Details $_.Exception.Message
    Write-Error "Failed to reset printer subsystem: $($_.Exception.Message)"
}
