# Wi-Fi Diagnostics Project

## Purpose

The Wi-Fi diagnostics workflow collects information needed to troubleshoot slow,
unstable, or unavailable wireless connections on Windows endpoints.

Primary implementation:

`./scripts/Collect-WiFiDiagnostics.ps1`

## Existing Capabilities

The current script collects:

- Wireless interface information
- Wireless driver information
- Saved wireless profiles
- IP configuration
- Routing information
- DNS server information
- Network adapter information
- Gateway and public-host latency tests
- Optional Windows WLAN report
- Optional safe remediation
- A compressed diagnostic bundle

## WLAN Report

Windows generates a wireless diagnostic report with:

    netsh wlan show wlanreport

The expected report location is:

    C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html

The existing script exposes this functionality through:

    -IncludeWlanReport

Example:

    .\scripts\Collect-WiFiDiagnostics.ps1 `
        -TicketId INC12345 `
        -IncludeWlanReport

## Development Goals

The WLAN report workflow should:

1. Verify that `netsh.exe` is available.
2. Run the WLAN report command.
3. Capture standard output and errors.
4. Verify the command exit code.
5. Confirm that the expected HTML file exists.
6. Copy the report into the diagnostic working directory.
7. Include the report in the final ZIP bundle.
8. Report success or failure clearly.
9. Avoid logging sensitive report contents.
10. Support automated tests through mocked commands.

## Security Considerations

The WLAN report can contain:

- Computer and adapter details
- Wireless network names
- Connection history
- Driver information
- WLAN event information
- Failure and disconnection reasons

Diagnostic bundles should be stored securely and handled according to the
organization's support-data and retention policies.

## Windows 11 Test Plan

Testing will be performed on a dedicated Windows 11 virtual machine.

The initial test should confirm:

- The script runs in Windows PowerShell 5.1.
- The script runs in PowerShell 7.
- The report is generated successfully.
- The HTML report is included in the ZIP.
- The output directory is inside the repository.
- Missing or disabled Wi-Fi hardware is handled gracefully.
- Non-administrator execution produces a clear result.
