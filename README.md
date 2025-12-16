# 📘 PowerShell-Helpdesk-Toolkit

### Interactive PowerShell toolkit for Tier 2/3 helpdesk and endpoint diagnostics

---

## Overview

**PowerShell-Helpdesk-Toolkit** is a production-oriented, interactive PowerShell toolkit designed to support **Tier 2 and Tier 3 endpoint troubleshooting** in enterprise environments.

It provides a standardized, menu-driven workflow for diagnosing and resolving common but complex helpdesk issues involving identity, networking, VPN connectivity, endpoint state, and user productivity services. The toolkit emphasizes **diagnostics first, safe remediation second**, with structured logging to support accurate escalation and root-cause analysis.

This project is intended for real-world support operations where issues are often intermittent, environment-dependent, and difficult to reproduce without consistent tooling.

---

## 🎫 Supported Ticket Scenarios

The toolkit is designed around common, high-friction helpdesk tickets seen in enterprise environments, including but not limited to:

### Identity & Authentication
- Outlook or Microsoft 365 repeatedly prompting for authentication
- VPN connected but authentication failures persist
- User account lockouts with unclear source
- Sign-in failures after password changes or device restarts
- Stale credentials or token-related authentication issues

### Network, DNS & Proxy
- VPN connected but internal resources do not resolve
- Slow or intermittent network performance
- Internal DNS resolution failures
- Proxy or PAC file configuration mismatches
- Split-tunnel routing issues impacting application connectivity

### Endpoint Health & Performance
- Slow logon or extended Group Policy processing
- Endpoint behaving differently “since yesterday”
- Startup items or services impacting performance
- BitLocker health or encryption state concerns

### File Access & Productivity
- Mapped drives missing or inaccessible
- SMB file share latency or authentication failures
- OneDrive sync issues or stalled clients
- Outlook profile corruption or OST-related errors
- Teams or browser cache and profile issues

### Hardware & Peripheral Services
- Print spooler hangs or failed print jobs
- Wi-Fi connectivity instability or roaming issues

Each scenario is addressed through **guided diagnostics**, producing clear output and logs that help determine whether the issue is endpoint-specific, identity-related, network-related, or requires escalation to another team.

---

## Key Features

- Interactive, menu-driven PowerShell launcher
- Diagnostics-first workflows with optional safe remediation
- Coverage across identity, networking, VPN, endpoint, and productivity tooling
- Centralized, timestamped log output for escalation and auditability
- Designed for use during live ticket handling

---

## Project Structure

```
scripts/helpdesk/
├── Invoke-HelpdeskToolkit.ps1   # Main interactive launcher
├── Test-*                       # Diagnostic scripts
├── Repair-*                     # Safe remediation scripts
├── Reset-*                      # Targeted reset actions
├── Collect-*                    # Log and evidence collection
└── Get-*                        # Change and delta analysis
```

Logs are written to:

```
out/HelpdeskLogs/
```

---

## Usage

Launch the toolkit from an elevated PowerShell session if required:

```powershell
.\Invoke-HelpdeskToolkit.ps1
```

Some actions may require administrative privileges depending on the task being performed.

---

## Design Principles

- Diagnostics before remediation
- Safe, explicit operator actions
- Repeatable and standardized troubleshooting
- Enterprise-aligned assumptions
- Escalation-ready evidence and logging

---

## Intended Audience

- Tier 2 / Tier 3 Helpdesk Engineers
- Endpoint and Desktop Engineers
- IT Support Leads
- Enterprise Support Teams

---

## Author

**Tim Heverin**  
GitHub: https://github.com/dj-3dub

---

## License

MIT
