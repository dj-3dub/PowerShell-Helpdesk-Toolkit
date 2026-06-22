# PowerShell Helpdesk Toolkit (Standard Operating Procedure)

### Standard Operating Procedure (SOP) for Tier 2/3 Helpdesk Operations and Endpoint Diagnostics

---

## 1. Document Overview & Governance

This Standard Operating Procedure (SOP) governs the use of the **PowerShell Helpdesk Toolkit**, an interactive utility designed for desktop support engineers, enterprise analysts, and system administrators. It standardizes diagnostics and remediation paths for user accounts, network stacks, local client caches, and VPN/Citrix connectivity.

To ensure compliance with corporate IT auditing and security standards (including ISO 27001, SOC 2, and law firm governance):
* **Audit Logging**: Every diagnostic run and remediation action is written to a centralized, local audit log at `out\HelpdeskLogs\audit.log`.
* **Windows Event Logging**: Script execution logs are duplicated in the **Windows Event Log** under `Application` (Source: `PowerShellHelpdeskToolkit`) when executed from an administrative PowerShell session.
* **Scope-Safe Execution**: State modifications (e.g. registry deletions, service resets) are wrapped in try/catch blocks with safety checks.

---

## 2. Prerequisites & Requirements

Before running the toolkit, ensure the following requirements are met on the target workstation:

* **Windows OS**: Windows 10 (21H2 or later) and Windows 11 are recommended. Windows Server 2019/2022 supported for server-side diagnostics.
* **PowerShell Version**: Windows PowerShell 5.1 (Standard) or PowerShell Core 7.0+.
* **Execution Policy**: Set to `RemoteSigned` or `Bypass` in the scoping session:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
  ```
* **Privilege Level**:
  * **Standard User**: Most diagnostic tasks and browser profile cache clears can be run under standard user context.
  * **Administrator**: Required for restarting print spoolers, resetting Windows Update services, flushing network stacks, and reading security logs. Run PowerShell as an Administrator.
* **Remote Access Modules**: The Active Directory (RSAT) modules must be installed on the analyst's machine to run password resets and account unlocking tasks.

---

## 3. Quick Start

For analysts new to the toolkit, follow this operational sequence:

1. **Verify Network Connectivity** (Option 37): Run Citrix Gateway & M365 Endpoint Diagnostics first to confirm VPN and DNS are functioning. If this fails, network-layer issues are blocking all subsequent repairs.
2. **Address Identity Issues** (Options 18-21): If sign-in loops or account lockouts are reported, diagnose and resolve before application-specific repairs.
3. **Execute Application Repairs** (Options 1-12): Once connectivity and identity are verified, proceed with Microsoft Teams, Outlook, and printer subsystem resets.

---

## 4. Usage Instructions

Follow these steps to execute the toolkit:

1. Open an elevated (Administrator) PowerShell session.
2. Navigate to the repository root.
3. Launch the interactive menu:
   ```powershell
   .\scripts\Invoke-HelpdeskToolkit.ps1
   ```
4. Enter the option number corresponding to the troubleshooting task.
5. Provide requested inputs (such as Ticket ID or Target Username) when prompted.
6. Review terminal feedback and check generated logs under `out\HelpdeskLogs\`.

---

## 5. Console Interface Layout

The toolkit runs in a clean, interactive text-based CLI. A representation of the interface is shown below:

```text
==========================================
   Helpdesk Automation Toolkit (PowerShell)
==========================================

 1)  Reset printer subsystem
 2)  Reset Microsoft Teams cache
 3)  Reset Outlook profile
 4)  Repair network stack (DNS/Winsock/IP)
 5)  Repair Windows Update components
 6)  Collect helpdesk logs (ZIP)
 7)  Reset browser profiles (Chrome/Edge)
 8)  Check mailbox capacity and retention
 9)  Repair OneDrive sync (Soft/Reset + logs)
10)  VPN diagnostics (connectivity + routing)
11)  BitLocker health check (local machine)
12)  Outlook OST repair (scan/rebuild)
13)  Endpoint pre-flight checks
14)  Network performance diagnostics (slow/intermittent)
15)  View Network Performance Logs
16)  Proxy diagnostics (WinINET/WinHTTP/PAC/WPAD)
17)  View Proxy Diagnostics Logs
18)  Identity: Sign-in health triage
19)  Identity: Unlock AD account
20)  Identity: Force AD password reset
21)  Identity: Find lockout source (best effort)
22)  Wi-Fi diagnostics bundle (slow/unstable wireless)
23)  Logon + GPO performance diagnostics
24)  View Logon + GPO Diagnostics Logs
25)  Endpoint delta (What changed?) report
26)  View Endpoint Delta Logs
27)  Startup delta (Scheduled Tasks / Startup entries / Services)
28)  View Startup Delta Logs
29)  File share / mapped drive diagnostics (SMB 445 + auth + mappings)
30)  View File Share Diagnostics Logs
31)  Print diagnostics (spooler/queue/ports/drivers/events)
32)  View Print Diagnostics Logs
33)  Outlook auth prompts diagnostics (M365 sign-in loops)
34)  View Outlook Auth Prompt Logs
35)  VPN connected but nothing resolves (Internal DNS diagnostics)
36)  View Internal DNS Diagnostics Logs
37)  Citrix Gateway & M365 Endpoint Diagnostics
 Q)  Quit

Select an option: 
```

---

## 6. Features Breakdown & Operational Flow

The toolkit separates operations into two primary types: **Diagnostic Checks** and **Remediation Actions**.

### Diagnostic Checks
* **Citrix Gateway & M365 Endpoints (Option 37)**: Performs DNS lookup, route trace, and TCP port 443 reachability checks targeting Citrix gateways and core Microsoft 365 services (e.g. `login.microsoftonline.com`, `outlook.office365.com`). Isolates local DNS failures from remote gateway drops.
* **VPN & DNS Diagnostics (Options 10, 35)**: Identifies if a remote worker's connection drop is caused by split-tunnel routing issues or internal corporate DNS misconfigurations.
* **Identity Sign-In Health (Options 18, 21)**: Triages user lockout reasons, AD group policies, and credentials loops.

### Remediation Actions (Audited)
* **Printer Subsystem Reset (Option 1)**: Stops the Windows print spooler service, clears the local spool directory, and restarts the service.
* **Microsoft Teams / Browser Cache Reset (Options 2, 7)**: Safely closes target processes, backs up profile data, and removes cached files to fix token corruption.
* **Outlook OST/Profile Repair (Options 3, 12)**: Exports registry backups, runs scanpst, or renames corruption-prone OST files to trigger clean builds.
* **AD Management (Options 19, 20)**: Unlocks user accounts and resets domain passwords directly from the analyst's workstation using secure credential inputs.

**Destructive Operations Warning**: Options 1, 3, 7, 9, 12, and 14 modify or delete cached files and service configurations. Ensure end users have saved all work and understand the scope of the repair before executing. For sensitive accounts (executives, legal staff), confirm with their manager before proceeding.

---

## 8. Structure & Log Locations

```
├── out/
│   └── HelpdeskLogs/
│       ├── audit.log                       # Centralized, secure audit trail
│       ├── CitrixM365Diag_<Host>_<TS>.txt  # Connection diagnostic reports
│       └── ...                             # Specific ticket logs and zip bundles
└── scripts/
    ├── Common.ps1                          # Shared security logging and path utilities
    ├── Invoke-HelpdeskToolkit.ps1          # Main interactive menu launcher
    ├── Test-*.ps1                          # Read-only diagnostics
    └── Repair-*.ps1 / Reset-*.ps1          # Audited remediation and reset tasks
```

---

## 9. Known Limitations

* **PowerShell Version**: Requires PowerShell 5.1 or later. PowerShell 4.0 and earlier are not supported.
* **RSAT Dependency**: Active Directory password resets and account unlocking require the Remote Server Administration Tools (RSAT) to be installed on the analyst's machine.
* **Citrix Gateway Diagnostics**: Assumes standard TCP 443 routes. Non-standard port configurations or custom gateway architectures are not tested and may require manual diagnostics.
* **Concurrent Execution**: Running multiple toolkit instances simultaneously on the same machine may cause log contention. Execute one operation per session when possible.

---

## 10. Compliance and Auditing Summary

All actions executed through this toolkit append a record to `out\HelpdeskLogs\audit.log` containing:
1. **Timestamp** (ISO-8601 standard format).
2. **Operator Identity** (Windows username of the executing analyst).
3. **Target Action** (The script name and internal operation performed).
4. **Machine Info** (The hostname of the system running the action).
5. **Execution Outcome** (SUCCESS/WARNING/ERROR details with caught exceptions).

This fulfills strict law-firm security constraints by providing complete forensic traceability for helpdesk operators executing administrative tasks.

---

## 11. Support & Escalation

If the toolkit fails to resolve an end-user issue:

1. **Collect Logs**: Gather all relevant diagnostic logs from `out\HelpdeskLogs\` (especially `audit.log` and any option-specific diagnostic files).
2. **Review Error Messages**: Check for ERROR entries in the audit log that indicate permission issues, locked files, or service failures.
3. **Escalate to L3 Engineering**: Forward the ticket number, affected username, and complete diagnostic logs to the IT Engineering team for deeper investigation.

For toolkit bugs, feature requests, or documentation questions, contact the toolkit maintainer with the relevant logs and reproduction steps.

---

## 12. Sample Audit Log Output

The following entries demonstrate how the toolkit logs successful operations and errors:

```
[2026-06-22 14:32:15] [INFO] [DESKTOP-ABC123\analyst_user] [Reset-PrinterSubsystem.ps1 / ResetPrinterSubsystem] Successfully reset printer subsystem and cleared print queue.
[2026-06-22 14:35:42] [INFO] [DESKTOP-ABC123\analyst_user] [Reset-TeamsCache.ps1 / ResetTeamsCache] Microsoft Teams cache cleared successfully. | Details: Teams application closed, %AppData%\Microsoft\Teams\Cache cleared
[2026-06-22 14:38:18] [ERROR] [DESKTOP-ABC123\analyst_user] [Force-PasswordReset.ps1 / ForceADPasswordReset] Failed to reset AD password for jsmith. | Details: Access denied. Ensure your user account has AD management permissions.
[2026-06-22 14:41:05] [WARNING] [DESKTOP-ABC123\analyst_user] [Test-CitrixGatewayAndM365Endpoints.ps1 / CitrixConnectivity] Citrix gateway (192.168.1.50:443) unreachable. Check VPN connectivity or firewall rules.
```

Analysts can inspect `out\HelpdeskLogs\audit.log` to verify that operations completed successfully and to diagnose failures before escalating to L3 engineering.

---

## Author

**Tim Heverin**  
LinkedIn: https://www.linkedin.com/in/tim-heverin/

---

## License

MIT
