# Common.ps1 - Central utilities and logging helper for PowerShell Helpdesk Toolkit
# Ensures secure, auditable, and governance-compliant operations.

# Correct calculation of repository root (PSScriptRoot is scripts/)
if ($PSScriptRoot) {
    $global:RepoRoot = Split-Path -Path $PSScriptRoot -Parent
} else {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
    $global:RepoRoot = Split-Path -Path $scriptDir -Parent
}

$global:AuditLogPath = Join-Path $global:RepoRoot 'out\HelpdeskLogs\audit.log'

# Ensure audit log directory exists
$null = New-Item -Path (Split-Path -Path $global:AuditLogPath -Parent) -ItemType Directory -Force -ErrorAction SilentlyContinue

function Write-AuditLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]$Severity,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Details = ''
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $username = $env:USERNAME
    if (-not $username) { $username = "SYSTEM" }
    $computer = $env:COMPUTERNAME
    if (-not $computer) { $computer = "localhost" }

    $scriptName = $MyInvocation.ScriptName
    if (-not $scriptName) {
        $scriptName = "Interactive/CLI"
    } else {
        $scriptName = Split-Path $scriptName -Leaf
    }

    # Format standard entry
    $logEntry = "[{0}] [{1}] [{2}\{3}] [{4} / {5}] {6}" -f $timestamp, $Severity, $computer, $username, $scriptName, $Action, $Message
    if (-not [string]::IsNullOrWhiteSpace($Details)) {
        $logEntry += " | Details: $Details"
    }

    # 1. Log to local file
    try {
        $logEntry | Out-File -FilePath $global:AuditLogPath -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}

    # 2. Log to Windows Event Log if available
    if ($env:OS -eq "Windows_NT" -and (Get-Command Write-EventLog -ErrorAction SilentlyContinue)) {
        try {
            $source = "PowerShellHelpdeskToolkit"
            # Attempt to register event source if not present
            if (-not [System.Diagnostics.EventLog]::SourceExists($source)) {
                # New-EventLog requires administrative privileges
                New-EventLog -LogName Application -Source $source -ErrorAction SilentlyContinue
            }
            
            if ([System.Diagnostics.EventLog]::SourceExists($source)) {
                $entryType = 'Information'
                if ($Severity -eq 'WARNING') { $entryType = 'Warning' }
                elseif ($Severity -eq 'ERROR') { $entryType = 'Error' }

                Write-EventLog -LogName Application -Source $source -EventId 1000 -EntryType $entryType -Message $logEntry -ErrorAction SilentlyContinue
            }
        } catch {}
    }

    # Output to console
    switch ($Severity) {
        'INFO' {
            Write-Verbose $logEntry
        }
        'WARNING' {
            Write-Warning "${Action}: ${Message}"
        }
        'ERROR' {
            Write-Error "${Action}: ${Message}"
        }
    }
}
