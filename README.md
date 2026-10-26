<div align="center">

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/banner-dark.gif">
  <img alt="PowerShell Scripts" src="docs/banner-light.gif" width="100%">
</picture>

Everyday reporting scripts for Windows servers, file shares and Active Directory, with every report saved as an Excel workbook or HTML page and optionally emailed.

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-2F2C28?style=flat-square&labelColor=121110)](#requirements)
[![Licence](https://img.shields.io/github/license/nafthan/powershell-scripts?style=flat-square&label=licence&labelColor=121110&color=2F2C28)](LICENSE)

[nafthan.co.uk](https://nafthan.co.uk)

</div>

## About

These started as one script. I needed to find every path on a file share that was over the Windows 260-character limit, and it grew from there into the reports I actually reach for during the week: who has access to what, which accounts haven't logged on in months, whose password is about to expire, and whether the servers are healthy.

Two decisions shaped all of it. First, nothing needs installing: the scripts write real `.xlsx` workbooks themselves, so the machine running them doesn't need Excel or any extra modules. Second, it's meant to live in a public repo, so there are no server names, domains or email addresses in the scripts at all. Everything specific to an environment lives in one config file that git never sees.

## Scripts

| Area | Script | What it does |
|---|---|---|
| File server | `FileServer/Find-LongPaths.ps1` | Paths longer than the 260-character Windows limit, with how far over each one is. Standalone, no config needed |
| File server | `FileServer/Get-FolderSizeReport.ps1` | Size of every folder down to N levels, with file counts and last-changed dates |
| File server | `FileServer/Get-NtfsPermissionsAudit.ps1` | NTFS permissions, flagging broken inheritance, orphaned SIDs, Everyone or Domain Users with write access, and more |
| Active Directory | `ActiveDirectory/Get-InactiveAccountsReport.ps1` | Users and computers with no logon for N days, with an optional disable-and-move mode |
| Active Directory | `ActiveDirectory/Get-PasswordExpiryReport.ps1` | Passwords expiring soon, and can email users a reminder |
| Servers | `ServerHealth/Invoke-DailyHealthCheck.ps1` | Uptime, pending reboot, disk space, stopped services, last update and event log errors, as one HTML report |
| Servers | `ServerHealth/Get-CertificateExpiryReport.ps1` | Expiring certificates in server stores and on live HTTPS endpoints |
| Network | `Network/Get-OpenPortReport.ps1` | Which TCP ports are open on each server, compared with what should be |
| Security | `Security/Get-RiskyAccountReport.ps1` | AD accounts with risky settings, rated High, Medium or Low, each with a fix |

## Getting started

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- For the Active Directory scripts: the ActiveDirectory module (RSAT), or run them on a domain controller
- For remote server checks: WinRM on the target servers, and an account with local admin rights on them. The health check falls back to DCOM if WinRM fails

### Setup

```powershell
git clone https://github.com/nafthan/powershell-scripts.git
cd powershell-scripts
Get-ChildItem -Recurse -Filter *.ps1 | Unblock-File      # only needed if you downloaded a zip

Copy-Item .\config.example.psd1 .\config.psd1
notepad .\config.psd1
```

Replace every placeholder in `config.psd1` (`example.com`, `SERVER01`, `<...>`) with your own values. The scripts warn you if a setting they use still contains a placeholder.

If your mail relay needs a login, save it as an encrypted credential file while logged on as the account that will run the scripts, then set `CredentialFile = 'smtp.cred.xml'` in the Email section:

```powershell
Get-Credential | Export-Clixml -Path .\smtp.cred.xml
```

Reports go to `Reports\` in the toolkit folder unless you change `General.OutputFolder`.

## Usage

Every setting in `config.psd1` can be overridden with the matching parameter, and `Get-Help .\<script>.ps1 -Full` lists them all.

```powershell
.\FileServer\Find-LongPaths.ps1 -Path '\\FILESERVER01\Shares' -OutputFile 'C:\Temp\LongPaths.xlsx'
.\FileServer\Get-FolderSizeReport.ps1 -Path 'D:\Data' -Depth 3 -Top 100
.\FileServer\Get-NtfsPermissionsAudit.ps1 -Path '\\FILESERVER01\Shares\Finance' -MaxDepth 3

.\ActiveDirectory\Get-InactiveAccountsReport.ps1                            # report only
.\ActiveDirectory\Get-InactiveAccountsReport.ps1 -DisableAccounts -WhatIf   # preview changes
.\ActiveDirectory\Get-PasswordExpiryReport.ps1 -NotifyUsers -TestRecipient 'you@yourdomain'

.\ServerHealth\Invoke-DailyHealthCheck.ps1 -SendEmail -OnlyEmailOnIssues
.\ServerHealth\Get-CertificateExpiryReport.ps1 -Endpoint 'https://intranet.yourdomain' -SendEmail

.\Network\Get-OpenPortReport.ps1 -ComputerName SERVER01, SERVER02
.\Security\Get-RiskyAccountReport.ps1 -SendEmail
```

| Script | Worth knowing |
|---|---|
| Find-LongPaths | Flags anything over 259 characters (`-MaxLength` to change it). Folders it can't read go on their own sheet, so you know what the scan missed |
| NTFS audit | The top folder is listed in full; subfolders only show entries that differ from their parent. `-AllEntries` lists everything |
| Inactive accounts | Report-only unless you add `-DisableAccounts`, and even then it asks before each account. Domain controllers are never touched |
| Password expiry | Users are only emailed on the days in `PasswordExpiry.NotifyOnDays`, so a daily run doesn't nag them every day |
| Certificate expiry | Endpoint checks read the certificate the site is actually serving, which catches a site still serving the old one after a renewal |
| Open ports | Results show what's reachable from wherever you run it. Run it from a user VLAN as well as the server VLAN to compare |
| Risky accounts | Read-only. Known exceptions go in `RiskyAccounts.AcceptedRisks`; they stay in the report but don't trigger alerts |

### Scheduling

Run them as scheduled tasks under a service account with the rights each script needs, for example the health check every morning:

```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
           -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\powershell-scripts\ServerHealth\Invoke-DailyHealthCheck.ps1" -SendEmail'
$trigger = New-ScheduledTaskTrigger -Daily -At 07:30
Register-ScheduledTask -TaskName 'Daily health check' -Action $action -Trigger $trigger `
           -User 'DOMAIN\svc_reports' -Password (Read-Host 'Password') -RunLevel Highest
```

If you use an SMTP credential file, create it while logged on as the task account, because only that account can decrypt it.

## How it works

Every script loads `Shared/Common.ps1`, which holds the parts they all need.

| Part | What it does |
|---|---|
| `Shared/Common.ps1` | Loads `config.psd1` and warns about leftover placeholders, writes `.xlsx` workbooks, builds email-friendly HTML tables and sends mail |
| `config.example.psd1` | Every environment-specific setting, as placeholders. Copied to `config.psd1`, which is git-ignored |
| `FileServer/`, `ActiveDirectory/`, `ServerHealth/`, `Network/`, `Security/` | One script per report |

An `.xlsx` file is a zip of XML files, so the workbook writer builds one directly. Numbers are stored as real numbers, so you can sort and filter them, and text is escaped properly so paths containing `&`, `<` or quotes don't break the file.

The file system scans read paths with the `\\?\` prefix, so they can still get into folders that are already too deep for Explorer. Find-LongPaths walks the tree with its own stack rather than recursion, so one unreadable folder is logged and skipped instead of stopping the run, and junctions are reported but never followed, so a loop can't trap it.

## Security and privacy

- No real environment details are in the repo. `config.psd1`, `*.cred.xml` and generated reports are all in `.gitignore`.
- Passwords never go in files. SMTP logins use `Export-Clixml`, which Windows encrypts so only that user on that machine can read it.
- Anything that changes something is opt-in. The inactive accounts script needs `-DisableAccounts`, supports `-WhatIf` and confirms each account.
- The port check only makes a TCP connection and sends no data. Only scan systems you're responsible for.
- Before pushing a fork, run `git status` to check no `config.psd1`, credential or report files are staged. A secret scanner such as [gitleaks](https://github.com/gitleaks/gitleaks) is a good extra check. New settings get a placeholder in `config.example.psd1`, using the reserved `example.com` domain or `<ANGLE BRACKETS>` for free text.

## Testing

I checked every script with PowerShell's own parser and PSScriptAnalyzer, including its compatibility check for both 5.1 and 7, and tested the logic that doesn't need a real domain. The workbook writer was tested with awkward data and opened in two spreadsheet libraries. The NTFS rights logic was tested against real permission values, which caught Windows storing the Synchronize bit on some rights and not others. The port check ran against live test listeners and a public host, and the risky account audit against a fake domain with one account for each risky setting.

Testing under PowerShell 7 also caught a difference in how the scripts created lists, which Windows PowerShell 5.1 didn't have. Every script was fixed and the earlier tests re-run under both.

## Limitations

- Last logon comes from `lastLogonTimestamp`, which can be up to 14 days behind the real last logon.
- Port results only show what's reachable from the machine running the script.
- The remote checks need WinRM (or DCOM for the health check) and admin rights on the target servers.
- Run each script by hand once in your environment before scheduling it.

## Licence

MIT. See [LICENSE](LICENSE).

<div align="center">
<sub><a href="https://nafthan.co.uk">nafthan.co.uk</a></sub>
</div>
