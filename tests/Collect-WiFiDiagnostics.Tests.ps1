BeforeAll {
    $ScriptPath = Join-Path `
        $PSScriptRoot `
        '..\scripts\Collect-WiFiDiagnostics.ps1'
}

Describe 'Collect-WiFiDiagnostics' {
    It 'has a Wi-Fi diagnostic script' {
        Test-Path -LiteralPath $ScriptPath | Should -BeTrue
    }

    It 'contains WLAN report support' {
        $content = Get-Content -LiteralPath $ScriptPath -Raw

        $content | Should -Match 'IncludeWlanReport'
        $content | Should -Match 'netsh wlan show wlanreport'
        $content | Should -Match 'wlan-report-latest\.html'
    }
}
