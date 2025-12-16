# PowerShell-Playbook — New Automation Scripts (20251028-170120)

Drop these into your repo under `Scripts/Automation/` (or adjust to your structure). Each script includes `-?` help.

## Included
- `Diagnose-SlowPC.ps1` — One-shot triage for slow PCs; HTML + TXT report and zipped bundle.
- `Repair-WindowsUpdate.ps1` — Reset WU components; optional scan/install via COM.
- `Invoke-ChocoAppSync.ps1` — Chocolatey bootstrap + upgrade-all; ensure from `packages.config`.
- `Export-LocalAdmins.ps1` — Export local Administrators group members (single/multiple hosts) to CSV.
- `Collect-SupportBundle.ps1` — msinfo32, dxdiag, ipconfig, services, processes, event errors → zipped bundle.
- `Reset-Proxy.ps1` — Show/clear/set WinHTTP & WinINET proxy settings.
- `Repair-WindowsSearch.ps1` — Reset index, run SFC/DISM.
- `Fix-DnsStack.ps1` — Show + repair DNS/Winsock; set DNS servers per interface.

## Quick Usage
```powershell
# 1) Slow PC triage
.\Scripts\Automation\Diagnose-SlowPC.ps1 -Hours 3 -Verbose

# 2) Windows Update reset + scan
.\Scripts\Automation\Repair-WindowsUpdate.ps1 -Scan -Verbose

# 3) Chocolatey sync
.\Scripts\Automation\Invoke-ChocoAppSync.ps1 -PackagesConfig .\packages.config

# 4) Export local admins from multiple machines
.\Scripts\Automation\Export-LocalAdmins.ps1 -ComputerName PC01,PC02 -OutputCsv .\local-admins.csv

# 5) Collect support bundle
.\Scripts\Automation\Collect-SupportBundle.ps1

# 6) Proxy quick-fix
.\Scripts\Automation\Reset-Proxy.ps1 -Show
.\Scripts\Automation\Reset-Proxy.ps1 -Clear

# 7) Repair Windows Search
.\Scripts\Automation\Repair-WindowsSearch.ps1 -ResetIndex -RunSfc

# 8) DNS/Winsock fix
.\Scripts\Automation\Fix-DnsStack.ps1 -Repair
```

## Notes
- Prefer running elevated (`Run as administrator`) for full fidelity.
- Scripts avoid external dependencies unless noted (Chocolatey installer uses its official bootstrapper).
- Use `-Verbose` to see progress and `-WhatIf` in scripts that support it.