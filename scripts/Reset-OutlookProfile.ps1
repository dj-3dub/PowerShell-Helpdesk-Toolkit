[CmdletBinding(SupportsShouldProcess = $true)]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load common utilities
. (Join-Path $PSScriptRoot 'Common.ps1')

Write-Verbose "Closing Outlook..."
try {
    Get-Process outlook -ErrorAction SilentlyContinue | Stop-Process -Force

    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    $backupRoot = "$env:USERPROFILE\OutlookProfileBackup"
    $timestamp = (Get-Date -Format 'yyyyMMdd-HHmmss')
    $backupPath = "$backupRoot\OutlookProfile-$timestamp.reg"

    if (Test-Path $regPath) {
        Write-Verbose "Backing up Outlook profile registry key to $backupPath"

        if ($PSCmdlet.ShouldProcess("Outlook Profile", "Backup + Remove")) {
            $null = New-Item -Path $backupRoot -ItemType Directory -Force
            # Reg export returns non-zero exit code if it fails
            reg export "HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles" $backupPath /y
            if ($LASTEXITCODE -ne 0) {
                throw "Registry export failed with exit code $LASTEXITCODE"
            }

            Write-Verbose "Deleting Outlook profile registry key..."
            Remove-Item -Path $regPath -Recurse -Force
            
            Write-AuditLog -Severity INFO -Action "ResetOutlookProfile" -Message "Successfully reset Outlook profile and backed up registry to $backupPath."
        }
    }
    else {
        Write-Verbose "No Outlook profile found to reset."
        Write-AuditLog -Severity WARNING -Action "ResetOutlookProfile" -Message "No Outlook registry profile key found at $regPath; no action taken."
    }

    Write-Host "Outlook profile reset complete. Outlook will rebuild the profile on next launch." -ForegroundColor Green
}
catch {
    Write-AuditLog -Severity ERROR -Action "ResetOutlookProfile" -Message "Failed to reset Outlook profile." -Details $_.Exception.Message
    Write-Error "Failed to reset Outlook profile: $($_.Exception.Message)"
}
