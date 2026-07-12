BeforeAll {
    $ScriptPath = Join-Path `
        $PSScriptRoot `
        '..\scripts\Collect-WiFiDiagnostics.ps1'

    $ScriptPath = [System.IO.Path]::GetFullPath($ScriptPath)
    $ScriptContent = Get-Content -LiteralPath $ScriptPath -Raw
    $ParsedTokens = $null
    $ParseErrors = $null

    $null = [System.Management.Automation.Language.Parser]::ParseFile(
        $ScriptPath,
        [ref]$ParsedTokens,
        [ref]$ParseErrors
    )
}

Describe 'Collect-WiFiDiagnostics script structure' {
    It 'exists in the expected location' {
        Test-Path -LiteralPath $ScriptPath | Should -BeTrue
    }

    It 'contains no PowerShell parser errors' {
        $ParseErrors | Should -BeNullOrEmpty
    }

    It 'uses advanced function behavior' {
        $ScriptContent | Should -Match '\[CmdletBinding\(\)\]'
    }

    It 'enables strict mode' {
        $ScriptContent |
            Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }

    It 'uses the repository root derived from PSScriptRoot' {
        $ScriptContent |
            Should -Match '\$repoRoot\s*=\s*Split-Path\s+-Path\s+\$PSScriptRoot\s+-Parent'
    }

    It 'writes diagnostic output under out\HelpdeskLogs' {
        $ScriptContent |
            Should -Match "out\\HelpdeskLogs"
    }
}

Describe 'Collect-WiFiDiagnostics parameters' {
    It 'supports an optional ticket ID' {
        $ScriptContent |
            Should -Match '\[string\]\$TicketId'
    }

    It 'supports WLAN report collection' {
        $ScriptContent |
            Should -Match '\[switch\]\$IncludeWlanReport'
    }

    It 'supports safe remediation' {
        $ScriptContent |
            Should -Match '\[switch\]\$EnableSafeRemediation'
    }

    It 'supports a configurable public test host' {
        $ScriptContent |
            Should -Match '\[string\]\$PublicTestHost'
    }

    It 'supports configurable WLAN event-log history' {
        $ScriptContent |
            Should -Match '\[int\]\$EventLogHours'
    }

    It 'limits event-log history to 1 through 168 hours' {
        $ScriptContent |
            Should -Match '\[ValidateRange\(1,\s*168\)\]'
    }
}

Describe 'Windows platform and safety controls' {
    It 'rejects unsupported operating systems' {
    $ScriptContent |
        Should -Match '\$env:OS\s+-ne\s+''Windows_NT'''
    }

    It 'detects administrator membership' {
        $ScriptContent |
            Should -Match 'WindowsBuiltInRole\]::Administrator'
    }

    It 'prevents remediation when not elevated' {
        $ScriptContent |
            Should -Match '\$EnableSafeRemediation\s+-and\s+-not\s+\$isAdministrator'
    }

    It 'sanitizes ticket IDs used in filenames' {
        $ScriptContent |
            Should -Match '\$safeTicketId\s*=\s*\$TicketId\s+-replace'
    }
}

Describe 'WLAN report collection' {
    It 'runs the Windows WLAN report command' {
        $ScriptContent |
            Should -Match 'wlan\s+show\s+wlanreport'
    }

    It 'checks the native command exit code' {
        $ScriptContent |
            Should -Match '\$LASTEXITCODE'
    }

    It 'uses the expected Windows WLAN report path' {
        $ScriptContent |
            Should -Match 'Microsoft\\Windows\\WlanReport\\wlan-report-latest\.html'
    }

    It 'copies the WLAN report into the diagnostic bundle' {
        $ScriptContent |
            Should -Match "wlan-report-latest\.html"
    }

    It 'includes a WLAN report sensitivity warning' {
        $ScriptContent |
            Should -Match 'potentially sensitive information'
    }
}

Describe 'WLAN AutoConfig event collection' {
    It 'queries the WLAN AutoConfig operational log' {
        $ScriptContent |
            Should -Match 'Microsoft-Windows-WLAN-AutoConfig/Operational'
    }

    It 'uses Get-WinEvent' {
        $ScriptContent |
            Should -Match 'Get-WinEvent'
    }

    It 'exports a readable text event report' {
        $ScriptContent |
            Should -Match 'wlan_autoconfig_events\.txt'
    }

    It 'exports structured event data to CSV' {
        $ScriptContent |
            Should -Match 'wlan_autoconfig_events\.csv'
    }
}

Describe 'Diagnostic bundle creation' {
    It 'creates a ZIP file from the working directory' {
        $ScriptContent |
            Should -Match 'CreateFromDirectory'
    }

    It 'preserves the working directory when ZIP creation fails' {
        $ScriptContent |
            Should -Match 'working directory was preserved'
    }

    It 'returns the resulting ZIP file object' {
        $ScriptContent |
            Should -Match 'Get-Item\s+-LiteralPath\s+\$zipPath'
    }
}
