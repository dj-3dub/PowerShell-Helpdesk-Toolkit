<#
.SYNOPSIS
    Checks mailbox capacity and retention configuration.

.DESCRIPTION
    For one or more mailboxes, this script:
      - Retrieves mailbox size, quotas, and storage limit status
      - Calculates % used of ProhibitSendQuota (or IssueWarningQuota if ProhibitSend not set)
      - Inspects the assigned retention policy and its tags
      - Flags whether a "MoveToArchive" tag with AgeLimitForRetention <= 365 days exists
      - Optionally exports a CSV report

    Useful for tickets like:
      - "Mailbox at 100GB capacity"
      - "Is this user on the correct retention policy?"
      - "Should email older than 1 year be auto-archived?"

.PARAMETER Identity
    One or more mailbox identities (UPN, alias, or display name).
    If omitted and -AllMailboxes is not used, you will be prompted.

.PARAMETER AllMailboxes
    If specified, runs against all user mailboxes.

.PARAMETER WarningThresholdPercent
    Threshold (in %) of quota at which to flag a mailbox as "NearCapacity".
    Default is 90.

.PARAMETER ExportCsv
    If specified, exports results to a CSV under exports\Mailboxes.

.PARAMETER OutputPath
    Optional explicit CSV path. If not supplied and -ExportCsv is used,
    the report is written to exports\Mailboxes\Test-MailboxCapacity-<timestamp>.csv
    under the repo root.

.EXAMPLE
    # Single mailbox
    .\Test-MailboxCapacityAndRetention.ps1 -Identity user@contoso.com -Verbose

.EXAMPLE
    # All mailboxes, export CSV
    .\Test-MailboxCapacityAndRetention.ps1 -AllMailboxes -ExportCsv

.NOTES
    Requires Exchange cmdlets:
      - Preferably: Get-EXOMailbox, Get-EXOMailboxStatistics (EXO V2)
      - Or:        Get-Mailbox, Get-MailboxStatistics (on-prem / older EXO)

    Run Connect-ExchangeOnline (or open Exchange Management Shell) before using.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [string[]]$Identity,

    [Parameter(Mandatory = $false)]
    [switch]$AllMailboxes,

    [Parameter(Mandatory = $false)]
    [int]$WarningThresholdPercent = 90,

    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # Detect Exchange cmdlets
    $exoMailboxCmd  = Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue
    $exoStatsCmd    = Get-Command -Name Get-EXOMailboxStatistics -ErrorAction SilentlyContinue
    $legacyMailbox  = Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue
    $legacyStats    = Get-Command -Name Get-MailboxStatistics -ErrorAction SilentlyContinue

    if (-not $exoMailboxCmd -and -not $legacyMailbox) {
        throw "No Exchange mailbox cmdlets found. Connect to Exchange Online or open an Exchange Management Shell first."
    }

    function Get-MailboxWrapper {
        param(
            [Parameter(Mandatory = $false)][string[]]$Id,
            [Parameter(Mandatory = $false)][switch]$All
        )

        if ($exoMailboxCmd) {
            if ($All) {
                return Get-EXOMailbox -ResultSize Unlimited
            }
            elseif ($Id) {
                $all = @()
                foreach ($i in $Id) {
                    $all += Get-EXOMailbox -Identity $i
                }
                return $all
            }
        }
        else {
            if ($All) {
                return Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
            }
            elseif ($Id) {
                $all = @()
                foreach ($i in $Id) {
                    $all += Get-Mailbox -Identity $i
                }
                return $all
            }
        }
    }

    function Get-MailboxStatsWrapper {
        param(
            [Parameter(Mandatory = $true)][string]$Identity
        )

        if ($exoStatsCmd) {
            return Get-EXOMailboxStatistics -Identity $Identity
        }
        elseif ($legacyStats) {
            return Get-MailboxStatistics -Identity $Identity
        }
        else {
            throw "No mailbox statistics cmdlets available."
        }
    }

    function Convert-ToGB {
        param(
            [Parameter(Mandatory = $true)]
            $Value
        )

        if ($null -eq $Value) { return $null }

        # Try ByteQuantifiedSize (Value.ToBytes())
        if ($Value -is [Microsoft.Exchange.Data.ByteQuantifiedSize]) {
            return [Math]::Round(($Value.ToBytes() / 1GB), 2)
        }

        # Try string like "99.13 GB (106,512,345,678 bytes)"
        if ($Value -is [string]) {
            if ($Value -match '([\d\.,]+)\s*GB') {
                [double]$num = ($matches[1] -replace ',', '.')
                return [Math]::Round($num, 2)
            }
        }

        # Fallback: try to cast to double
        try {
            return [Math]::Round(([double]$Value / 1GB), 2)
        }
        catch {
            return $null
        }
    }

    function Get-RetentionDetails {
        param(
            [Parameter(Mandatory = $true)]
            [string]$PolicyName
        )

        $result = [pscustomobject]@{
            RetentionPolicy         = $PolicyName
            RetentionTags           = $null
            HasOneYearArchivePolicy = $false
        }

        if (-not $PolicyName) {
            return $result
        }

        $policyCmd = Get-Command -Name Get-RetentionPolicy -ErrorAction SilentlyContinue
        $tagCmd    = Get-Command -Name Get-RetentionPolicyTag -ErrorAction SilentlyContinue

        if (-not $policyCmd -or -not $tagCmd) {
            return $result
        }

        try {
            $policy = Get-RetentionPolicy -Identity $PolicyName -ErrorAction Stop
        }
        catch {
            return $result
        }

        $tags = @()

        foreach ($link in $policy.RetentionPolicyTagLinks) {
            try {
                $tags += Get-RetentionPolicyTag -Identity $link -ErrorAction SilentlyContinue
            }
            catch { }
        }

        if ($tags.Count -gt 0) {
            $result.RetentionTags = ($tags | Select-Object -ExpandProperty Name | Sort-Object -Unique) -join '; '

            # Check for MoveToArchive tag with AgeLimitForRetention <= 365 days
            $moveTags = $tags | Where-Object {
                $_.RetentionAction -eq 'MoveToArchive' -and $_.AgeLimitForRetention
            }

            foreach ($t in $moveTags) {
                # AgeLimitForRetention is a TimeSpan
                try {
                    if ($t.AgeLimitForRetention.Days -le 365) {
                        $result.HasOneYearArchivePolicy = $true
                        break
                    }
                }
                catch { }
            }
        }

        return $result
    }

    # repo root + export folder (for CSV)
    $repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
    $exportRoot = Join-Path $repoRoot 'exports\Mailboxes'
    $null = New-Item -Path $exportRoot -ItemType Directory -Force -ErrorAction SilentlyContinue

    $script:results = New-Object System.Collections.Generic.List[object]
}

process {
    if (-not $AllMailboxes -and -not $Identity) {
        $inputId = Read-Host "Enter mailbox identity (UPN/alias). For multiple, separate with commas"
        if (-not [string]::IsNullOrWhiteSpace($inputId)) {
            $Identity = $inputId.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }
    }

    $mailboxes = Get-MailboxWrapper -Id $Identity -All:$AllMailboxes

    if (-not $mailboxes) {
        Write-Warning "No mailboxes found for the given parameters."
        return
    }

    foreach ($mbx in $mailboxes) {
        Write-Verbose "Processing mailbox: $($mbx.UserPrincipalName ?? $mbx.PrimarySmtpAddress)"

        $stats = $null
        try {
            $stats = Get-MailboxStatsWrapper -Identity ($mbx.UserPrincipalName ?? $mbx.PrimarySmtpAddress)
        }
        catch {
            Write-Warning "Failed to get statistics for $($mbx.UserPrincipalName ?? $mbx.PrimarySmtpAddress): $_"
        }

        # Size / quota calculations
        $totalSizeGB = $null
        $quotaGB     = $null
        $percentUsed = $null

        if ($stats) {
            $totalSizeGB = Convert-ToGB -Value $stats.TotalItemSize
        }

        # Use ProhibitSendQuota as primary limit, fallback to IssueWarningQuota
        $quotaSource = $null
        if ($mbx.ProhibitSendQuota -and $mbx.ProhibitSendQuota -ne 'Unlimited') {
            $quotaGB     = Convert-ToGB -Value $mbx.ProhibitSendQuota
            $quotaSource = 'ProhibitSendQuota'
        }
        elseif ($mbx.IssueWarningQuota -and $mbx.IssueWarningQuota -ne 'Unlimited') {
            $quotaGB     = Convert-ToGB -Value $mbx.IssueWarningQuota
            $quotaSource = 'IssueWarningQuota'
        }

        if ($quotaGB -and $totalSizeGB -ne $null -and $quotaGB -gt 0) {
            $percentUsed = [Math]::Round(($totalSizeGB / $quotaGB) * 100, 2)
        }

        $nearCapacity = $false
        if ($percentUsed -ne $null -and $percentUsed -ge $WarningThresholdPercent) {
            $nearCapacity = $true
        }

        # Retention
        $retentionInfo = Get-RetentionDetails -PolicyName $mbx.RetentionPolicy

        # Archive
        $archiveEnabl

