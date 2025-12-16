# 📑 Changelog

All notable changes to **PowerShell Playbook** will be documented here.

---

## [Unreleased]

- GitHub Actions CI/CD workflow for automated lint + test
- New automation modules (deployment, compliance, reporting)
 - `scripts/Export-ADObjects-NoRSAT.ps1`: added `-Mock` mode to simulate AD objects for offline testing and CI (use `-MockCount` to control rows).

---

## [2025-09-24] – 🚀 Repo Refresh & VMware Mock Deploy

### ✨ Added
- **Mock VMware Deployment Script**  
  `scripts/dev/Invoke-WinServerVmBuild-Mock.ps1`  
  - Simulates a Windows Server deployment in VMware (no vCenter required).  
  - Attaches 4 NICs and configures 2 NIC teams.  
  - Handles sizing (CPU, memory, disk) and IP/DNS config.  
  - Ideal for demonstrating automation workflows in interviews or labs.

### 📝 Changed
- **README.md**
  - Updated repo name to **PowerShell Playbook**
  - Improved layout and section flow
  - Added VMware mock deployment usage example
  - Cleaned up repo structure section
  - Fixed formatting in code snippets
  - Refined Topics list for discoverability

### ✅ Improvements
- Project docs now align with current scope of the repo
- README is recruiter-ready and portfolio-friendly
