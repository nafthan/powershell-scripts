# PowerShell Scripts

By **nafthan** - [nafthan.co.uk](https://nafthan.co.uk)

Day-to-day reporting scripts for Windows servers, file shares, Active Directory and Microsoft 365.

Every report is saved as an Excel workbook (`.xlsx`) or HTML page, and can be emailed. Excel and extra PowerShell modules are not needed on the machine running the scripts.

| Area | Script | What it does |
|---|---|---|
| File server | `FileServer/Get-FolderSizeReport.ps1` | Size of every folder down to N levels, with file counts and last-changed dates |
| File server | `FileServer/Get-NtfsPermissionsAudit.ps1` | NTFS permissions, flagging broken inheritance, orphaned SIDs, Everyone/Domain Users with write access, and more |
| File server | `FileServer/Find-LongPaths.ps1` | Paths longer than the 260-character Windows limit |
| Active Directory | `ActiveDirectory/Get-InactiveAccountsReport.ps1` | Users and computers with no logon for N days; can optionally disable and move them |
| Active Directory | `ActiveDirectory/Get-PasswordExpiryReport.ps1` | Passwords expiring soon; can email users reminders |
| Servers | `ServerHealth/Invoke-DailyHealthCheck.ps1` | Uptime, pending reboot, disk space, stopped services, last update and event log errors, as an HTML report |
| Servers | `ServerHealth/Get-CertificateExpiryReport.ps1` | Expiring certificates in server stores and on live HTTPS endpoints |
| Network | `Network/Get-OpenPortReport.ps1` | Which TCP ports are reachable on your servers, compared with the ports that should be open |
| Microsoft 365 | `M365/Get-MfaStatusReport.ps1` | MFA status for every user, with sign-in status, last sign-in, creation date, methods, admin roles and risk flags |
| Security | `Security/Get-RiskyAccountReport.ps1` | AD accounts with settings attackers abuse (Kerberoastable, AS-REP roastable, password not required, unconstrained delegation, and more) |

## Two ways to use these scripts

| | Standalone | Toolkit |
|---|---|---|
| Where | `Standalone/` | The other folders, with `Shared/` and `config.psd1` |
| Setup | None - download one `.ps1` and run it | Copy the whole repo and fill in `config.psd1` once |
| Settings | Parameters on each run | One config file shared by every script, with parameters to override |
| Best for | Trying a script, one-off reports | Scheduled tasks and regular use across several scripts |

The scripts work the same either way.

### Standalone scripts

Download a script from `Standalone/`, unblock it, and run it:

```powershell
Unblock-File .\Get-FolderSizeReport.ps1
.\Get-FolderSizeReport.ps1 -Path 'D:\Shares' -Depth 2
Get-Help .\Get-FolderSizeReport.ps1 -Full      # every option
```

- Reports are saved in the current folder.
- Settings that live in the toolkit's config file use built-in defaults, for example the health check thresholds and the certificate stores.
- A few scripts get extra parameters instead:
  - `Get-OpenPortReport.ps1 -AllowedPort 135,445,3389` flags any other open port.
  - `Get-PasswordExpiryReport.ps1 -HelpdeskContact 'the IT Service Desk on 0123 456789'` sets the contact shown in user reminders.
  - `Get-InactiveAccountsReport.ps1` has no default exclusions. Use `-Exclude 'svc_*','admin_*'` to skip service and admin accounts.
- To email a report, add the mail settings:
  ```powershell
  .\Get-MfaStatusReport.ps1 -SendEmail -SmtpServer smtp.yourdomain -MailFrom reports@yourdomain -MailTo it@yourdomain
  ```
  Add `-SmtpPort`, `-SmtpUseSsl` and `-SmtpCredential (Get-Credential)` if your mail server needs them.

The standalone files are generated from the toolkit scripts by `build/Build-Standalone.ps1`. Don't edit them directly: change the toolkit script or `Shared/Common.ps1`, then run `.\build\Build-Standalone.ps1` to regenerate them.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+.
- **AD scripts:** the ActiveDirectory module (RSAT: Active Directory Domain Services tools), or run them on a domain controller.
- **Microsoft 365 script:** the Microsoft Graph module (`Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`), and an account with Global Reader, or with both Reports Reader and Security Reader.
- **Remote server checks:** WinRM enabled on the target servers, and an account with local admin rights on them.
  - The health check falls back to DCOM if WinRM fails.

## Toolkit setup

1. **Clone** the repository (`git clone https://github.com/nafthan/powershell-scripts.git`) or download it to the server or admin PC that will run the scripts.

2. **Unblock the files** if you downloaded a zip:
   ```powershell
   Get-ChildItem -Recurse -Filter *.ps1 | Unblock-File
   ```

3. **Create your config** from the template:
   ```powershell
   Copy-Item .\config.example.psd1 .\config.psd1
   notepad .\config.psd1
   ```
   - Replace every placeholder (`example.com`, `SERVER01`, `<...>`) with your real values.
   - `config.psd1` is in `.gitignore`, so it is never committed.
   - The scripts warn you if a setting they use still contains a placeholder.

4. **Optional: SMTP authentication.** If your mail relay needs a login, create an encrypted credential file. Run this as the account that will run the scripts:
   ```powershell
   Get-Credential | Export-Clixml -Path .\smtp.cred.xml
   ```
   Then set `CredentialFile = 'smtp.cred.xml'` in the Email section.
   - The file can only be decrypted by that user on that machine.
   - `*.cred.xml` is in `.gitignore`.

Reports are saved to `Reports\` in the toolkit folder by default. Change this with `General.OutputFolder`.

## Usage

Every setting in `config.psd1` can be overridden with the matching parameter. Run `Get-Help .\<script>.ps1 -Full` for every option.

### Folder size report
```powershell
.\FileServer\Get-FolderSizeReport.ps1 -Path '\\FILESERVER01\Shares' -Depth 2
.\FileServer\Get-FolderSizeReport.ps1 -Path 'D:\Data' -Depth 3 -Top 100
```
Sheets: **Folder Sizes**, **Access Errors**, **Summary**. Each folder's size includes everything below it.

### NTFS permissions audit
```powershell
.\FileServer\Get-NtfsPermissionsAudit.ps1 -Path '\\FILESERVER01\Shares\Finance' -MaxDepth 3
.\FileServer\Get-NtfsPermissionsAudit.ps1 -Path 'D:\Shares\HR' -MaxDepth -1 -AllEntries
```
Sheets: **Permissions**, **Flagged**, **Access Errors**, **Summary**.
- The top folder is listed in full.
- Subfolders only show entries that differ from their parent, which keeps the report readable.
- Use `-AllEntries` to list everything.

### Inactive accounts
```powershell
.\ActiveDirectory\Get-InactiveAccountsReport.ps1                                   # report only
.\ActiveDirectory\Get-InactiveAccountsReport.ps1 -DisableAccounts -WhatIf          # preview changes
.\ActiveDirectory\Get-InactiveAccountsReport.ps1 -DisableAccounts                  # asks before each account
```
- It is report-only unless you add `-DisableAccounts`.
- Disabled accounts get a description noting the date and reason, and are moved to `ActiveDirectory.DisabledOU` if that is set.
- Domain controllers are never touched.
- Last logon comes from `lastLogonTimestamp`, which can be up to 14 days behind the real last logon.

### Password expiry
```powershell
.\ActiveDirectory\Get-PasswordExpiryReport.ps1 -WarningDays 21
.\ActiveDirectory\Get-PasswordExpiryReport.ps1 -NotifyUsers -TestRecipient 'you@yourdomain'   # all reminders to you
.\ActiveDirectory\Get-PasswordExpiryReport.ps1 -NotifyUsers -SendReport                       # daily task
```
- Users are only emailed on the days listed in `PasswordExpiry.NotifyOnDays`, so a daily run doesn't email them every day.
- Edit the help-desk contact and instructions in the config.

### Daily health check
```powershell
.\ServerHealth\Invoke-DailyHealthCheck.ps1 -ComputerName SERVER01, SERVER02
.\ServerHealth\Invoke-DailyHealthCheck.ps1 -SendEmail -OnlyEmailOnIssues
```
- Each server is marked OK, Warning or Critical, using the thresholds in the `HealthCheck` section.
- Automatic services that are normally stopped are listed in `HealthCheck.ExcludeServices`. Add any of your own that cause false alerts.

### Certificate expiry
```powershell
.\ServerHealth\Get-CertificateExpiryReport.ps1 -ComputerName WEB01 -Endpoint 'https://intranet.yourdomain'
.\ServerHealth\Get-CertificateExpiryReport.ps1 -SendEmail          # emails only if something is expiring
```
- Endpoint checks read the certificate the site is actually serving. They catch sites still serving an old certificate after a renewal.

### Open port check
```powershell
.\Network\Get-OpenPortReport.ps1 -ComputerName SERVER01, SERVER02
.\Network\Get-OpenPortReport.ps1 -ComputerName SERVER01 -Port 80, 443, 3389
.\Network\Get-OpenPortReport.ps1 -SendEmail          # emails only if there are findings
```
Each port is reported as **Open**, **Closed** (refused) or **Filtered** (no reply, usually a firewall). Findings:
- **High:** a risky service is open (`PortCheck.RiskyPorts`, e.g. Telnet, FTP, VNC).
- **Medium:** a port is open but not in the server's `PortCheck.Allowed` list.
- **Low:** an allowed port isn't reachable.

Results show what's reachable from the machine running the script. Run it from different network segments to compare what users and servers can reach. Only scan systems you're responsible for.

### Microsoft 365 MFA report
```powershell
.\M365\Get-MfaStatusReport.ps1
.\M365\Get-MfaStatusReport.ps1 -IncludeGuests -IncludePerUserMfaState
.\M365\Get-MfaStatusReport.ps1 -UseDeviceCode          # sign in from a server without a browser
.\M365\Get-MfaStatusReport.ps1 -SendEmail
```
Read-only. Sheets: **MFA Status**, **Needs Attention**, **Admins**, **Methods**, **Conditional Access**, **Errors**, **Summary**.

- **The first run** asks a Global Administrator to consent to these read-only Graph permissions: User.Read.All, AuditLog.Read.All, UserAuthenticationMethod.Read.All, Policy.Read.All and RoleManagement.Read.Directory.
- **Licensing:** the MFA registration report and last sign-in dates need **Entra ID P1 or P2**, which is included in Microsoft 365 Business Premium and E3/E5.
  - Without P1/P2, the script reads each user's methods one by one instead. This is slower, and last sign-in dates won't be available.
  - Eligible (PIM) admin roles need P2.
  - Anything unavailable is listed on the **Errors** sheet.
- **Risk ratings:**
  - **Critical:** an admin with no MFA.
  - **High:** an enabled user with no MFA.
  - **Medium:** SMS/voice-only MFA, or a method not allowed by policy.
  - **Low:** an admin without a phishing-resistant method, a guest with no MFA, or a user who has never signed in or has no sign-in for `M365.StaleDays`.
- **Registered isn't enforced.** "Registered" means the user has set up MFA. Whether it's required at sign-in depends on security defaults or Conditional Access, which the report also checks.
- **Scheduled runs:** use app-only sign-in with a certificate. Create an Entra app registration with the permissions above as *application* permissions, then set `M365.TenantId`, `ClientId` and `CertificateThumbprint`.

### Risky account settings
```powershell
.\Security\Get-RiskyAccountReport.ps1
.\Security\Get-RiskyAccountReport.ps1 -SearchBase 'OU=Service Accounts,DC=corp,DC=example,DC=com' -SendEmail
```
- Read-only; any domain user can run it.
- Findings are rated High, Medium or Low, each with a recommended fix.
- Members of Domain Admins, Administrators and the other built-in admin groups, including through nested groups, are treated as privileged, and their findings rated higher.
- Known, accepted risks can be listed in `RiskyAccounts.AcceptedRisks`. They stay in the report marked **Accepted**, but don't count towards totals or emails.

## Scheduling

Create a scheduled task that runs as a service account with the rights each script needs. For example, the daily health check at 07:30:

```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
           -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\powershell-scripts\ServerHealth\Invoke-DailyHealthCheck.ps1" -SendEmail'
$trigger = New-ScheduledTaskTrigger -Daily -At 07:30
Register-ScheduledTask -TaskName 'Daily health check' -Action $action -Trigger $trigger `
           -User 'DOMAIN\svc_reports' -Password (Read-Host 'Password') -RunLevel Highest
```

- When run with `-File`, list parameters can be passed as one comma-separated value, e.g. `-ComputerName "SERVER01,SERVER02"`.
- If you use `Email.CredentialFile`, create it while logged on as the task account, because only that account can decrypt it.

## Publishing safely

- Commit `config.example.psd1` and never `config.psd1`.
- Before pushing, run `git status` and check that no `config.psd1`, `*.cred.xml` or report files are staged.
- If you add a setting, put a placeholder for it in `config.example.psd1`. Use the reserved `example.com` domain, or `<ANGLE BRACKETS>` for free text.
- For extra safety, run a secret scanner such as [gitleaks](https://github.com/gitleaks/gitleaks) before pushing.

## Layout

```
├── Standalone/             single-file versions of every script (generated)
├── build/                  Build-Standalone.ps1 - regenerates Standalone/
├── config.example.psd1     settings template (placeholders only)
├── Shared/Common.ps1       config loading, Excel writer, HTML and email helpers
├── FileServer/
├── ActiveDirectory/
├── ServerHealth/
├── Network/
├── M365/
└── Security/
```

## Licence

MIT - see [LICENSE](LICENSE).
