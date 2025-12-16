[CmdletBinding(SupportsShouldProcess = $true)]
param()

$spoolService = 'Spooler'
$spoolPath    = Join-Path $env:SystemRoot 'System32\spool\PRINTERS'

Write-Verbose "Target spool folder: $spoolPath"

if (-not (Test-Path $spoolPath)) {
    Write-Verbose "Spool folder '$spoolPath' does not exist. Nothing to clear."
}
else {
    if ($PSCmdlet.ShouldProcess($spoolPath, "Clear print queue")) {
        Write-Verbose "Stopping Print Spooler service..."
        Stop-Service -Name $spoolService -Force -ErrorAction SilentlyContinue

        Write-Verbose "Removing files from '$spoolPath'..."
        Get-ChildItem -Path $spoolPath -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

        Write-Verbose "Starting Print Spooler service..."
        Start-Service -Name $spoolService -ErrorAction SilentlyContinue
    }
}

Write-Host "Printer subsystem reset complete. Try printing again." -ForegroundColor Green
