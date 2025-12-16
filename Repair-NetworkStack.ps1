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

Write-Host "Network stack repair complete. A reboot may be required." -ForegroundColor Green
