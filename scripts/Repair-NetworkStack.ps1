<#
.SYNOPSIS
    Repairs the Windows network stack (DNS flush, Winsock reset, IP renew).

.DESCRIPTION
    Runs a set of safe network troubleshooting commands used by helpdesk teams.

.EXAMPLE
    .\Repair-NetworkStack.ps1 -Verbose
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load common utilities
. (Join-Path $PSScriptRoot 'Common.ps1')

try {
    if ($PSCmdlet.ShouldProcess("Network Interfaces", "Flush DNS, Reset Winsock, Renew IP")) {
        Write-Verbose "Flushing DNS..."
        ipconfig /flushdns | Out-Null

        Write-Verbose "Resetting Winsock..."
        netsh winsock reset | Out-Null

        Write-Verbose "Resetting TCP/IP..."
        netsh int ip reset | Out-Null

        Write-Verbose "Releasing IP..."
        ipconfig /release | Out-Null

        Write-Verbose "Renewing IP..."
        ipconfig /renew | Out-Null
        
        Write-AuditLog -Severity INFO -Action "RepairNetworkStack" -Message "Successfully repaired the local network stack (DNS flushed, Winsock/TCP reset, IP renewed)."
    }
    Write-Host "Network stack repair complete. A reboot may be required." -ForegroundColor Green
}
catch {
    Write-AuditLog -Severity ERROR -Action "RepairNetworkStack" -Message "Failed to repair network stack." -Details $_.Exception.Message
    Write-Error "Failed to repair network stack: $($_.Exception.Message)"
}
