# PowerShell Helpdesk Toolkit Roadmap

## Phase 1 — Project Foundation

- Establish a clean development branch
- Document the existing Wi-Fi diagnostic workflow
- Correct repository path handling
- Add basic validation commands
- Preserve current functionality

## Phase 2 — WLAN Report Improvements

- Validate `netsh wlan show wlanreport`
- Include the WLAN report in diagnostic bundles
- Verify the generated HTML report location
- Capture command output and exit status
- Improve failure reporting
- Add sensitive-data warnings

## Phase 3 — Wi-Fi Diagnostic Expansion

- Export WLAN AutoConfig event logs
- Capture Wi-Fi adapter driver information
- Capture signal strength and connection details
- Capture available wireless networks
- Add before-and-after remediation results
- Generate a diagnostic summary

## Phase 4 — Automated Testing

- Add Pester
- Add tests for path handling
- Mock Windows commands
- Test successful WLAN report generation
- Test missing-report behavior
- Test ZIP bundle contents

## Phase 5 — Code Quality

- Add PSScriptAnalyzer
- Standardize error handling
- Move shared functions into Common.ps1
- Add operating-system and elevation checks
- Add GitHub Actions validation

## Phase 6 — Toolkit Architecture

- Replace repetitive launcher code with an action registry
- Add noninteractive command execution
- Return structured PowerShell objects
- Add JSON and CSV output options
- Introduce toolkit configuration files

## Phase 7 — Documentation and Release

- Improve the main README
- Add architecture documentation
- Add security and privacy guidance
- Add sanitized screenshots and sample reports
- Publish a versioned release
