# =============================================================================
#  PowerShell scripts - example configuration
#  https://github.com/nafthan/powershell-scripts
#
#  1. Copy this file to config.psd1 in the same folder.
#  2. Replace every placeholder (example.com, SERVER01, <...>) with your values.
#
#  config.psd1 is listed in .gitignore, so your real values are never committed.
#  Only edit THIS file to add new settings or change placeholder text.
#
#  Never put passwords in this file. For SMTP authentication, see
#  Email.CredentialFile below.
#
#  Any setting here can be overridden by the matching script parameter.
# =============================================================================
@{
    General = @{
        # Where reports are saved. Relative paths are relative to the toolkit folder.
        OutputFolder = 'Reports'
    }

    Email = @{
        SmtpServer = 'smtp.example.com'
        Port       = 25
        UseSsl     = $false
        From       = 'it-reports@example.com'
        To         = @('it-team@example.com')

        # Leave empty for an anonymous internal relay. For an authenticated relay,
        # create an encrypted credential file once, as the account that runs the scripts:
        #     Get-Credential | Export-Clixml -Path .\smtp.cred.xml
        # then set:  CredentialFile = 'smtp.cred.xml'
        # (*.cred.xml is ignored by git; the file only decrypts for that user on that machine.)
        CredentialFile = ''
    }

    # Servers checked by Invoke-DailyHealthCheck and Get-CertificateExpiryReport.
    Servers = @{
        ComputerName = @('SERVER01', 'SERVER02')
    }

    FolderSize = @{
        Path  = '\\FILESERVER01\Shares'
        Depth = 2        # how many folder levels below Path to report on
        Top   = 0        # 0 = all folders, otherwise only the N largest
    }

    NtfsAudit = @{
        Path     = '\\FILESERVER01\Shares\Department'
        MaxDepth = 3     # folder levels below Path to check; -1 = no limit
    }

    ActiveDirectory = @{
        # Optional - leave empty to use the whole domain / nearest domain controller.
        SearchBase = 'OU=Staff,DC=corp,DC=example,DC=com'
        Server     = ''

        InactiveDays = 90

        # Where -DisableAccounts moves accounts. Leave empty to disable without moving.
        DisabledOU = 'OU=Disabled Accounts,DC=corp,DC=example,DC=com'

        # Accounts never reported or disabled (wildcards allowed, matched on SamAccountName).
        Exclude = @('krbtgt', 'Guest', 'svc_*', 'admin_*')
    }

    PasswordExpiry = @{
        WarningDays = 14

        # When run daily with -NotifyUsers, users are only emailed on these days-left values
        # so they get a few reminders rather than one every day.
        NotifyOnDays = @(14, 7, 3, 2, 1, 0)

        # Shown in the email sent to users.
        HelpdeskContact     = 'the IT Service Desk on <PHONE NUMBER> or <SERVICE DESK EMAIL>'
        ChangeInstructions  = 'Press Ctrl+Alt+Delete on a domain-joined PC and choose "Change a password", or use <SELF-SERVICE PASSWORD URL>.'
    }

    HealthCheck = @{
        DiskWarningPercent  = 15     # free space below this % = Warning
        DiskCriticalPercent = 5      # free space below this % = Critical
        UptimeWarningDays   = 45     # not rebooted for longer than this = Warning
        PatchWarningDays    = 35     # no update installed for longer than this = Warning
        ErrorEventWarning   = 50     # more System/Application errors than this in 24h = Warning

        # Pending file rename operations often come from antivirus/backup agents and
        # cause false "pending reboot" alerts. Set to $true to include them anyway.
        CheckPendingFileRename = $false

        # Automatic services that are normally stopped (trigger-start or delayed). Wildcards allowed.
        ExcludeServices = @(
            'gupdate', 'gupdatem', 'edgeupdate', 'edgeupdatem', 'MapsBroker', 'RemoteRegistry',
            'sppsvc', 'CDPSvc', 'OneSyncSvc*', 'tiledatamodelsvc', 'WbioSrvc', 'DoSvc', 'UsoSvc',
            'BITS', 'ShellHWDetection', 'TrustedInstaller', 'wuauserv', 'GoogleUpdater*', 'clr_optimization_*'
        )
    }

    Certificates = @{
        WarningDays  = 30
        CriticalDays = 7

        # Certificate stores checked on each server in Servers.ComputerName.
        Stores = @('Cert:\LocalMachine\My', 'Cert:\LocalMachine\WebHosting')

        # Certificates that expired more than this many days ago are ignored (old leftovers).
        IgnoreExpiredOlderThanDays = 30

        # HTTPS / TLS endpoints to check. Formats: 'https://host', 'host', 'host:port'.
        Endpoints = @('https://intranet.example.com', 'mail.example.com:443')
    }

    PortCheck = @{
        # Servers to check. Leave empty to use Servers.ComputerName.
        ComputerName = @()

        # How long to wait for each server's ports (milliseconds).
        TimeoutMs = 3000

        # Ports to check and their names (keys must be quoted).
        Ports = @{
            '21' = 'FTP';  '22' = 'SSH';  '23' = 'Telnet';  '25' = 'SMTP';  '53' = 'DNS';  '80' = 'HTTP'
            '88' = 'Kerberos';  '135' = 'RPC';  '139' = 'NetBIOS';  '389' = 'LDAP';  '443' = 'HTTPS'
            '445' = 'SMB';  '636' = 'LDAPS';  '1433' = 'SQL Server';  '3306' = 'MySQL';  '3389' = 'RDP'
            '5900' = 'VNC';  '5985' = 'WinRM HTTP';  '5986' = 'WinRM HTTPS';  '8080' = 'HTTP alt'
        }

        # Always flagged as High when open (cleartext or commonly abused remote access).
        RiskyPorts = @(21, 23, 69, 512, 513, 514, 5800, 5900)

        # Which ports SHOULD be open. Anything else open is flagged. Keys are server names
        # (wildcards allowed); 'Default' applies to servers not listed. Remove this whole
        # setting to only flag RiskyPorts.
        Allowed = @{
            'Default'  = @(135, 445, 3389, 5985)
            'SERVER01' = @(80, 135, 443, 445, 3389, 5985)
            'DC*'      = @(53, 88, 135, 389, 445, 636, 3389, 5985)
        }
    }

    RiskyAccounts = @{
        # Optional - leave empty to check the whole domain.
        SearchBase = ''

        # Privileged accounts whose password is older than this are flagged.
        PrivilegedPasswordMaxDays = 365

        # krbtgt password older than this is flagged.
        KrbtgtMaxDays = 180

        # Known risks you have accepted. They stay in the report, marked "Accepted", but are
        # left out of the totals and email. Use the exact finding names from the report,
        # or '*' for all findings on that account.
        AcceptedRisks = @{
            '<service account name>' = @('Password never expires', 'Kerberoastable account')
        }
    }
}
