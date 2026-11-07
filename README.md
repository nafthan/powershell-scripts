<div align="center">

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/banner-dark.svg">
  <img alt="PowerShell Scripts" src="docs/banner-light.svg" width="100%">
</picture>

[nafthan.co.uk](https://nafthan.co.uk)

</div>

## About

Day to day reporting scripts for Windows servers, file shares, Active Directory and Microsoft 365.

Two decisions shaped all of it. First, nothing needs installing: the scripts write real `.xlsx` workbooks themselves, so the machine running them doesn't need Excel or any extra modules. Second, it's meant to live in a public repo, so there are no server names, domains or email addresses in the scripts at all. Everything specific to an environment lives in one config file that git never sees.

Every report is saved as an Excel workbook (`.xlsx`) or HTML page, and can be emailed. Excel and extra PowerShell modules are not needed on the machine running the scripts.

## Scripts

| Area | Script | What it does |
|---|---|---|
| File server | `FileServer/Get-FolderSizeReport.ps1` | Size of every folder down to N levels, with file counts and last-changed dates |
| File server | `FileServer/Get-NtfsPermissionsAudit.ps1` | NTFS permissions, flagging broken inheritance, orphaned SIDs, Everyone/Domain Users with write access, and more |
| File server | `FileServer/Find-LongPaths.ps1` | Paths longer than the 260-character Windows limit |
| Active Directory | `ActiveDirectory/Get-InactiveAccountsReport.ps1` | Users and computers with no logon for N days; can optionally disable and move them |
| Active Directory | `ActiveDirectory/Get-PasswordExpiryReport.ps1` | Passwords expiring soon; can email users reminders |
| Active Directory | `ActiveDirectory/Get-BitLockerKeyAudit.ps1` | Computers with no BitLocker recovery key backed up to AD (never reads the key itself) |
| Active Directory | `ActiveDirectory/Get-GpoHealthReport.ps1` | GPO problems: conflicting settings, unlinked/empty/disabled GPOs, missing Apply or read permissions (MS16-072), blocked inheritance |
| Servers | `ServerHealth/Invoke-DailyHealthCheck.ps1` | Uptime, pending reboot, disk space, stopped services, last update and event log errors, as an HTML report |
| Servers | `ServerHealth/Get-CertificateExpiryReport.ps1` | Expiring certificates in server stores and on live HTTPS endpoints |
| Network | `Network/Get-OpenPortReport.ps1` | Which TCP ports are reachable on your servers, compared with the ports that should be open |
| Microsoft 365 | `M365/Get-MfaStatusReport.ps1` | MFA status for every user, with sign-in status, last sign-in, creation date, methods, admin roles and risk flags |
| Microsoft 365 | `M365/Get-AdminRoleReport.ps1` | Everyone with an Entra admin role (active, PIM-eligible, via groups, apps), with MFA, last sign-in, risk flags and changes since the last run |
| Microsoft 365 | `M365/Get-AppCredentialExpiryReport.ps1` | App registration secrets and certificates (including SAML signing certificates) that are expired or expiring soon |
| Microsoft 365 | `M365/Get-RiskyAppPermissionReport.ps1` | Enterprise apps with risky permissions, unverified publishers, user consent and possible consent phishing |
| Microsoft 365 | `M365/Get-MailboxForwardingAudit.ps1` | Exchange Online forwarding to external addresses, inbox rules that forward or hide mail, transport rules and forwarding settings |
| Microsoft 365 | `M365/Get-MailboxPermissionReport.ps1` | Full Access, Send As, Send on Behalf and calendar permissions on every mailbox, flagging orphaned accounts and anonymous calendar sharing |
| Microsoft 365 | `M365/Get-GroupMembershipReport.ps1` | Every Team, Microsoft 365 group, distribution list and security group with members and owners; flags ownerless, empty and guest groups |
| Microsoft 365 | `M365/Get-SharePointSiteReport.ps1` | SharePoint sites with sharing settings, storage use, last activity and external users |
| Security | `Security/Get-NonCompliantDeviceReport.ps1` | Non-compliant Intune devices with the policies and settings they fail, plus Defender onboarding, sensor health, exposure, active alerts and antivirus state |
| Security | `Security/Invoke-WindowsTriage.ps1` | Collects incident response evidence from a Windows machine (local or remote) into a hashed zip, and highlights suspicious findings |
| Security | `Security/Invoke-AccountContainment.ps1` | Contains a compromised account: saves evidence, then disables it, resets the password, revokes sessions and cleans up mail forwarding and rules |
| Security | `Security/Search-MailboxActivity.ps1` | Pulls a mailbox's audit log (rule changes, deletes, sends, access by others, IPs) into a hashed evidence package |
| Security | `Security/Get-RiskyAccountReport.ps1` | AD accounts with settings attackers abuse (Kerberoastable, AS-REP roastable, password not required, unconstrained delegation, and more) |
| Intune | `Intune/Detect-` / `Remediate-WingetAppUpdates.ps1` | Remediation pair: keeps a list of apps updated with winget |
| Intune | `Intune/Detect-` / `Remediate-OneStart.ps1` | Remediation pair: finds and removes the OneStart.ai browser for every user |
| Intune | `Intune/Detect-` / `Remediate-OemBloatware.ps1` | Remediation pair: removes Dell, HP and Lenovo bloatware |
| Intune | `Intune/Detect-` / `Remediate-LowDiskSpace.ps1` | Remediation pair: clears temp files, Recycle Bin and caches when the system drive is low |

Can find a cheatsheet for general PowerShell usage on my site at [Cheatsheets](https://nafthan.co.uk/cheatsheets)

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
- **Microsoft 365 scripts:** the Microsoft Graph module (`Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`), and an account with Global Reader.
  - The mailbox forwarding audit, mailbox permissions and mailbox activity scripts use the Exchange Online module instead (`Install-Module ExchangeOnlineManagement -Scope CurrentUser`). Run it as an Exchange Administrator, because Global Reader may not be able to read every mailbox's inbox rules.
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
   - The scripts warn you if a setting they use still contains a placeholder.

4. **Optional: SMTP authentication.** If your mail relay needs a login, create an encrypted credential file. Run this as the account that will run the scripts:
   ```powershell
   Get-Credential | Export-Clixml -Path .\smtp.cred.xml
   ```
   Then set `CredentialFile = 'smtp.cred.xml'` in the Email section.
   - The file can only be decrypted by that user on that machine.

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

- **The first run** asks a Global Administrator to consent to these read-only Graph permissions: User.Read.All, AuditLog.Read.All, UserAuthenticationMethod.Read.All, Policy.Read.All, RoleManagement.Read.Directory and Organization.Read.All (for licence names).
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
- **Blank sign-in window:** Microsoft Graph module 2.34 and later signs in through the Windows sign-in broker, whose window can go blank or hide behind the terminal. Use `-UseDeviceCode`, or set `M365.UseDeviceCode = $true` in the config, to sign in with a code in any browser instead.
- **Scheduled runs:** use app-only sign-in with a certificate. Create an Entra app registration with the permissions above as *application* permissions, then set `M365.TenantId`, `ClientId` and `CertificateThumbprint`.

### Entra admin roles
```powershell
.\M365\Get-AdminRoleReport.ps1
.\M365\Get-AdminRoleReport.ps1 -SendEmail          # schedule it weekly to be told about new admins
```
Sheets: **Admins** (one row per person or app), **Assignments**, **Changes**, **Errors**, **Summary**.

- **What it covers:** active, time-bound and PIM-eligible assignments, roles given through role-assignable groups (each member is listed), apps with roles, and roles limited to an administrative unit.
- **Flags:**
  - **Critical:** an admin with no MFA.
  - **High:** a guest with an admin role.
  - **Medium:** a privileged account synced from on-premises AD, or an app with a privileged role.
  - **Low:** no phishing-resistant method, a licensed admin account, a permanent assignment where PIM is available, a stale or disabled admin, or an assignment to a deleted object.
- **Global Administrator count:** flagged if there are more than `AdminRoles.MaxGlobalAdmins`, or fewer than 2.
- **Changes since the last run:** each run saves a snapshot next to the report, and the next run lists the roles added and removed.
- **Licensing:** PIM details need Entra ID P2, and last sign-in dates need P1. Without them, active assignments are still listed.

### App secrets and certificates expiry
```powershell
.\M365\Get-AppCredentialExpiryReport.ps1
.\M365\Get-AppCredentialExpiryReport.ps1 -WarningDays 60 -SendEmail      # emails only if something is expiring
```
- Covers client secrets and certificates on app registrations, plus enterprise app certificates such as SAML signing certificates.
- Each app's owners are listed, so you know who to chase. Apps with no owners are flagged.
- Also flags client secrets valid for more than two years (`AppCredentials.MaxSecretLifetimeDays`).

### Risky app permissions
```powershell
.\M365\Get-RiskyAppPermissionReport.ps1
.\M365\Get-RiskyAppPermissionReport.ps1 -MinimumRisk High -SendEmail
```
- Lists application permissions (granted to the app itself) and delegated permissions (acting as a signed-in user) for every non-Microsoft enterprise app. Add `-IncludeMicrosoftApps` to include Microsoft's own apps.
- **Ratings:**
  - **Critical:** permissions that can take over the tenant, such as `Directory.ReadWrite.All` or `RoleManagement.ReadWrite.Directory`.
  - **High:** broad access to mail, files, chats or the directory.
  - **Adjustments:** delegated permissions rate one level lower. Apps from unverified external publishers rate one level higher.
- **Also flags:** apps consented by individual users, and high-risk apps added in the last 30 days, which can be a sign of consent phishing.

### Mailbox forwarding audit
```powershell
.\M365\Get-MailboxForwardingAudit.ps1
.\M365\Get-MailboxForwardingAudit.ps1 -Mailbox user@yourdomain        # one or more mailboxes, e.g. during an incident
.\M365\Get-MailboxForwardingAudit.ps1 -SkipInboxRules                 # much faster on large tenants
```
- **What it checks:**
  - mailbox forwarding (SMTP or to a contact)
  - inbox rules that forward, redirect or hide mail
  - transport rules that copy mail externally
  - the outbound spam policy and remote domain forwarding settings
- **Inbox rules:** the script flags rules that delete messages, or move them to folders like RSS Feeds, while marking them read or filtering on words like "invoice" or "password". This is a common pattern after an account takeover.
- **Speed:** inbox rules are read one mailbox at a time, so allow roughly a second per mailbox.
- **Sign-in:** uses the same `M365` config section. App-only sign-in needs the tenant's `.onmicrosoft.com` domain as `M365.TenantId`, plus the Exchange.ManageAsApp permission and an Exchange role for the app.

### BitLocker key audit
```powershell
.\ActiveDirectory\Get-BitLockerKeyAudit.ps1
.\ActiveDirectory\Get-BitLockerKeyAudit.ps1 -SearchBase 'OU=Laptops,DC=corp,DC=example,DC=com' -SendEmail
```
- Lists enabled workstations with no recovery key in AD, or whose newest key is older than `BitLocker.StaleDays`. Add `-IncludeServers` or `-IncludeDisabled` to widen it.
- Only reads when each key was saved, never the recovery password itself. Reading recovery objects needs Domain Admin or delegated read access to `msFVE-RecoveryInformation`.

### GPO health report
```powershell
.\ActiveDirectory\Get-GpoHealthReport.ps1
```
- Read-only; needs the GroupPolicy module (RSAT: Group Policy Management tools).
- **Conflicts:** the same setting set to different values by GPOs that apply to the same OU, showing which one wins.
- **Findings:** unlinked, empty, disabled or partly disabled GPOs, disabled links, blocked inheritance, GPOs nobody can apply, and GPOs that computers can't read (MS16-072).

### Mailbox permissions
```powershell
.\M365\Get-MailboxPermissionReport.ps1
.\M365\Get-MailboxPermissionReport.ps1 -Mailbox shared@yourdomain -IncludeCalendars
.\M365\Get-MailboxPermissionReport.ps1 -Trustee user@yourdomain        # what can this user get into?
```
- Exchange Online module. Sheets: **Permissions**, **By Trustee**, **Errors**, **Summary**.
- Flags permissions held by deleted accounts (orphaned SIDs), and calendars shared with Default or Anonymous.

### Group membership
```powershell
.\M365\Get-GroupMembershipReport.ps1
.\M365\Get-GroupMembershipReport.ps1 -GroupType Team -IncludeTeamsChannels
```
- Microsoft Graph (Group.Read.All, GroupMember.Read.All; Channel and ChannelMember read permissions for `-IncludeTeamsChannels`).
- Flags groups with no owner or a single owner, empty groups, guests and disabled members.

### SharePoint sites
```powershell
.\M365\Get-SharePointSiteReport.ps1 -AdminUrl https://contoso-admin.sharepoint.com
.\M365\Get-SharePointSiteReport.ps1 -IncludeExternalUsers -IncludeOneDrive
```
- SharePoint Online module (`Install-Module Microsoft.Online.SharePoint.PowerShell`). In PowerShell 7 it's loaded through Windows PowerShell compatibility. Run as a SharePoint Administrator.
- Flags sites that allow Anyone links, sites at 90% or more of their quota, locked sites and sites with no activity for `SharePoint.InactiveDays`.

### Non-compliant devices (Intune and Defender)
```powershell
.\Security\Get-NonCompliantDeviceReport.ps1
.\Security\Get-NonCompliantDeviceReport.ps1 -Platform Windows -SendEmail
.\Security\Get-NonCompliantDeviceReport.ps1 -SkipDefender          # Intune only
.\Security\Get-NonCompliantDeviceReport.ps1 -SkipSettingDetail     # faster on large tenants - no per-setting reasons
.\Security\Get-NonCompliantDeviceReport.ps1 -AntivirusForAllDevices
```
- Read-only. Sheets: **Devices** (needs attention), **All Devices**, **Failing Settings**, **By Setting**, **Defender Only**, **Active Alerts**, **Errors**, **Summary**.
- **Devices** lists every Intune device that is not compliant, in its grace period, in conflict or in error (add `-IncludeUnknown` for devices not yet evaluated), with the failing policies and settings. Devices Intune says are compliant also appear if Defender has a problem with them.
- **All Devices** lists every Intune device with the same columns, compliant or not.
- Each device is matched to Defender for Endpoint by Entra device ID, then by name, to add onboarding status, sensor health, exposure level and active alerts. If the Defender data can't be read, the report still includes every device, with the Defender columns blank and the reason on the **Errors** sheet.
- Windows devices on the **Devices** sheet also get their antivirus state (real-time protection, signatures, tamper protection). Add `-AntivirusForAllDevices` to read it for every Windows device. This is one request per device, so it's slower on large tenants.
- **Defender Only** lists devices Defender sees with problems, such as devices that can be onboarded but aren't, that Intune doesn't manage.
- **Risk ratings:**
  - **High:** not compliant and not encrypted, antivirus or real-time protection off, a Defender sensor that isn't active, an unmanaged device that isn't onboarded, a high-severity alert, or a jailbroken device.
  - **Medium:** not compliant, out-of-date signatures, tamper protection off, High exposure, other active alerts, or a Windows device Defender hasn't seen.
  - **Low:** in grace period, policy conflict or error, or no Intune check-in for `Devices.StaleSyncDays`.
- **Permissions:** DeviceManagementManagedDevices.Read.All and DeviceManagementConfiguration.Read.All, plus ThreatHunting.Read.All and SecurityAlert.Read.All for Defender. Roles: Global Reader (or Intune Administrator) and Security Reader.
- Defender data needs Defender for Endpoint (Plan 2, or Microsoft 365 E5 / Business Premium with Defender for Business). Without it, use `-SkipDefender`; anything unavailable is listed on the **Errors** sheet.

### Mailbox activity search
```powershell
.\Security\Search-MailboxActivity.ps1 -Mailbox user@yourdomain -Days 14 -CaseId INC0001
.\Security\Search-MailboxActivity.ps1 -Mailbox user@yourdomain -StartDate '2026-01-01' -EndDate '2026-01-08' -IncludeAccessByOthers
```
- Uses `Search-UnifiedAuditLog`, which replaces `Search-MailboxAuditLog` (retired by Microsoft in 2025). Needs auditing turned on in the tenant and the Audit Logs or View-Only Audit Logs role.
- Saves an evidence package like the other incident response scripts: CSVs, an HTML summary of indicators (new inbox rules, forwarding, mass deletes, sends, access from several IPs), a SHA-256 manifest and a zip.

### Intune remediations
Each pair goes into **Intune > Devices > Scripts and remediations** as one remediation: `Detect-*.ps1` as the detection script, `Remediate-*.ps1` as the remediation script. Run them as SYSTEM in 64-bit PowerShell, and leave signature checking off unless you sign them.

| Pair | Settings at the top of both scripts |
|---|---|
| WingetAppUpdates | `$AppIds` - winget IDs to keep updated (find with `winget search`). Only machine-wide installs can be updated as SYSTEM. |
| OneStart | None. Checks every user profile, loaded user registry hive, scheduled tasks and startup entries. |
| OemBloatware | `$Bloatware` - app names per manufacturer. Uses each app's silent uninstall command, not Win32_Product. Test on a pilot group first. |
| LowDiskSpace | `$MinFreePercent`, `$MinFreeGB`, and `$TempDays` (remediation). Never deletes user documents or downloads. |

- Logs go to `C:\ProgramData\IntuneScripts\`, which Intune can collect with **Collect diagnostics**.
- Keep the lists the same in both scripts of a pair.
- These don't use `Shared/` or the config file, so they're the same in `Standalone/`.

### Windows triage collector
```powershell
.\Security\Invoke-WindowsTriage.ps1 -CaseId IR-2026-014                        # this machine (run elevated)
.\Security\Invoke-WindowsTriage.ps1 -ComputerName PC042 -Hours 168 -OutputPath E:\Evidence
```
- **Read-only on the target.** Run it as an administrator. `-ComputerName` uses PowerShell remoting, and the results are written on your machine, not the suspect one.
- **What it collects:**
  - processes with their SHA-256 hashes and signatures
  - network connections, DNS cache and ARP table
  - persistence: services, scheduled tasks, Run keys, Winlogon, IFEO, startup folders and WMI subscriptions
  - local users and admins
  - Security/System/PowerShell/RDP/Defender event logs
  - Defender status and exclusions
  - recently dropped executables and scripts, prefetch and PowerShell history
- **Findings:** `00_Findings.csv` and `00_Summary.html` rate everything from Low to Critical. For example:
  - unsigned programs running from AppData
  - encoded PowerShell
  - persistence pointing at user folders
  - cleared logs
  - password spray
  - RDP from the internet
  - Defender turned off or with exclusions added
  - known attacker tools such as Mimikatz

  Findings are leads, not proof. Review the CSV files as well.
- **Evidence handling:**
  - Every file is listed in `manifest.csv` with its SHA-256.
  - The folder is zipped, and the zip's hash is written next to it. Record that hash in your case notes.
  - All times are UTC.

### Account compromise playbook
```powershell
.\Security\Invoke-AccountContainment.ps1 -User jane.doe@yourdomain -CaseId IR-2026-014 -WhatIf     # preview
.\Security\Invoke-AccountContainment.ps1 -User jane.doe@yourdomain -CaseId IR-2026-014             # asks before each step
.\Security\Invoke-AccountContainment.ps1 -User jane.doe@yourdomain -EvidenceOnly                   # change nothing
```
1. **Saves evidence first:** account details, groups, MFA methods with their dates, sign-ins and sign-in locations, audit events, devices, app consents, mailbox forwarding and inbox rules.
2. **Then contains the account,** asking before each step. Use `-Confirm:$false` to skip the prompts, or `-Skip Disable,Mailbox` to leave out steps.
   - Disables the account. Synced users are disabled in AD.
   - Resets the password to a random value. It is never shown or saved.
   - Revokes all sessions and refresh tokens.
   - Removes mailbox forwarding.
   - *Disables* (doesn't delete) suspicious inbox rules, so they're kept as evidence.
3. **MFA methods** are only removed if you ask: `-RemoveMfaMethodsAddedAfter '2026-10-28 09:00'` removes methods registered after that time, i.e. after the suspected compromise.
4. **Summary page:** lists indicators (sign-ins from several countries, legacy-protocol sign-ins, recently added MFA methods, suspicious rules, app consents) and next steps, including how to restore access.

- **Account types:** works for cloud-only, AD-synced (needs the ActiveDirectory module) and AD-only accounts (`-SkipCloud`).
- **Sign-in:** password reset needs a signed-in admin, not app-only sign-in.
- **Roles:** User Administrator + Authentication Administrator (Privileged Authentication Administrator if the user is an admin), and Exchange Administrator for the mailbox steps.

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
├── Security/              AD risk checks, device compliance and incident response
├── Intune/                Intune detection and remediation pairs
└── docs/                  PowerShell-One-Liners.md and the README banner
```

## Licence

MIT - see [LICENSE](LICENSE).
