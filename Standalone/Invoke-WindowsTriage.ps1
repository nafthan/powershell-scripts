<#
.SYNOPSIS
    Collects incident response evidence from a Windows machine into a hashed, zipped evidence package.
.DESCRIPTION
    Gathers processes (with hashes and signatures), network connections, persistence (services, tasks,
    run keys, WMI), accounts, key event logs, Defender status and recently dropped files, then highlights
    suspicious findings. Read-only on the target. Run as administrator.
.EXAMPLE
    .\Invoke-WindowsTriage.ps1 -CaseId IR-2026-014
.EXAMPLE
    .\Invoke-WindowsTriage.ps1 -ComputerName PC042 -Hours 168 -OutputPath E:\Evidence
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: local administrator on the target; PowerShell remoting (WinRM) for -ComputerName
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$ComputerName,
    [pscredential]$Credential,

    [string]$CaseId,

    [ValidateRange(1, 8760)] [int]$Hours,
    [ValidateRange(1, 365)]  [int]$FileDays,
    [ValidateRange(100, 100000)] [int]$MaxEvents,
    [switch]$SkipHashes,

    [string]$OutputPath
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('Hours'))     { $Hours     = [int](Get-Setting $config 'SOC' 'TriageHours' 72) }
if (-not $PSBoundParameters.ContainsKey('FileDays'))  { $FileDays  = [int](Get-Setting $config 'SOC' 'TriageFileDays' 14) }
if (-not $PSBoundParameters.ContainsKey('MaxEvents')) { $MaxEvents = [int](Get-Setting $config 'SOC' 'TriageMaxEvents' 5000) }
$ComputerName = Split-ListValue $ComputerName

# --- Collector ------------------------------------------------------------------
# Runs on the target (locally, or through Invoke-Command). Must be self-contained.

$collector = {
    param([int]$Hours, [int]$FileDays, [bool]$HashFiles, [int]$MaxEvents)

    $sections = [System.Collections.Generic.List[object]]::new()
    $log      = [System.Collections.Generic.List[object]]::new()
    $since    = (Get-Date).AddHours(-$Hours)
    $fileSince = (Get-Date).AddDays(-$FileDays)
    $hashCache = @{}

    function Utc($d) {
        if ($null -eq $d -or "$d" -eq '') { return '' }
        try { return ([datetime]$d).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') + 'Z' } catch { return "$d" }
    }

    function Add-Section([string]$Name, [scriptblock]$Block) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $rows = @(& $Block)
            $sections.Add([pscustomobject]@{ Name = $Name; Rows = $rows })
            $log.Add([pscustomobject]@{ Section = $Name; Status = 'OK'; Rows = $rows.Count; Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1); Error = '' })
        }
        catch {
            $sections.Add([pscustomobject]@{ Name = $Name; Rows = @() })
            $log.Add([pscustomobject]@{ Section = $Name; Status = 'Failed'; Rows = 0; Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1); Error = $_.Exception.Message })
        }
    }

    function Get-FileDetail([string]$Path) {
        $empty = [pscustomobject]@{ SHA256 = ''; Signature = ''; Signer = '' }
        if (-not $Path -or -not $HashFiles) { return $empty }
        $Path = [Environment]::ExpandEnvironmentVariables($Path.Trim('"'))
        if (-not $hashCache.ContainsKey($Path)) {
            $d = $empty
            if (Test-Path -LiteralPath $Path -PathType Leaf -ErrorAction SilentlyContinue) {
                $hash = ''; $status = ''; $signer = ''
                try { $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash } catch { $hash = 'Unreadable' }
                try {
                    $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
                    $status = "$($sig.Status)"
                    if ($sig.SignerCertificate) { $signer = ($sig.SignerCertificate.Subject -replace '^.*?CN=("?)([^,"]+)\1.*$', '$2') }
                } catch { $status = 'Unknown' }
                $d = [pscustomobject]@{ SHA256 = $hash; Signature = $status; Signer = $signer }
            }
            $hashCache[$Path] = $d
        }
        return $hashCache[$Path]
    }

    function Get-ExePath([string]$CommandLine) {
        # Executable path from a command line / service ImagePath
        if (-not $CommandLine) { return '' }
        $c = [Environment]::ExpandEnvironmentVariables($CommandLine.Trim())
        if ($c -match '^"([^"]+)"') { return $Matches[1] }
        if ($c -match '^(.+?\.(exe|dll|sys|com|scr|bat|cmd|ps1|vbs|js))(\s|$)') { return $Matches[1] }
        return ($c -split '\s+')[0]
    }

    function Get-EventFields($WinEvent) {
        $map = @{}
        try {
            $xml = [xml]$WinEvent.ToXml()
            foreach ($d in @($xml.Event.EventData.Data)) { if ($d.Name) { $map[$d.Name] = $d.'#text' } }
            if ($xml.Event.UserData) {
                foreach ($n in @($xml.Event.UserData.FirstChild.ChildNodes)) { $map[$n.LocalName] = $n.InnerText }
            }
        } catch { Write-Verbose "Could not parse event XML: $($_.Exception.Message)" }
        return $map
    }

    # --- System -------------------------------------------------------------------
    Add-Section 'SystemInfo' {
        $os = Get-CimInstance Win32_OperatingSystem
        $cs = Get-CimInstance Win32_ComputerSystem
        [pscustomobject]@{
            ComputerName = $env:COMPUTERNAME; Domain = $cs.Domain; OS = $os.Caption; Version = $os.Version; Build = $os.BuildNumber
            InstallDate = Utc $os.InstallDate; LastBoot = Utc $os.LastBootUpTime; TimeZone = (Get-TimeZone).Id
            LocalTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); UtcTime = Utc (Get-Date)
            CollectedBy = "$env:USERDOMAIN\$env:USERNAME"; Manufacturer = $cs.Manufacturer; Model = $cs.Model
            IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        }
    }

    Add-Section 'LoggedOnUsers' {
        $q = @(quser.exe 2>$null)
        foreach ($line in ($q | Select-Object -Skip 1)) { [pscustomobject]@{ Session = "$line".Trim() } }
    }

    # --- Processes and network ------------------------------------------------------
    $procNames = @{}
    Add-Section 'Processes' {
        foreach ($p in (Get-CimInstance Win32_Process)) {
            $procNames[[int]$p.ProcessId] = $p.Name
            $owner = ''
            try { $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction Stop; if ($o.User) { $owner = "$($o.Domain)\$($o.User)" } } catch { $owner = '(unknown)' }
            $f = Get-FileDetail $p.ExecutablePath
            [pscustomobject]@{
                ProcessId = $p.ProcessId; ParentProcessId = $p.ParentProcessId; Name = $p.Name; Path = $p.ExecutablePath
                CommandLine = $p.CommandLine; Owner = $owner; Started = Utc $p.CreationDate
                SHA256 = $f.SHA256; Signature = $f.Signature; Signer = $f.Signer
            }
        }
    }
    Add-Section 'NetworkConnections' {
        foreach ($c in (Get-NetTCPConnection)) {
            [pscustomobject]@{
                Protocol = 'TCP'; State = "$($c.State)"; LocalAddress = $c.LocalAddress; LocalPort = $c.LocalPort
                RemoteAddress = $c.RemoteAddress; RemotePort = $c.RemotePort; ProcessId = $c.OwningProcess
                ProcessName = $procNames[[int]$c.OwningProcess]; Created = Utc $c.CreationTime
            }
        }
        foreach ($u in (Get-NetUDPEndpoint)) {
            [pscustomobject]@{
                Protocol = 'UDP'; State = 'Listen'; LocalAddress = $u.LocalAddress; LocalPort = $u.LocalPort
                RemoteAddress = ''; RemotePort = ''; ProcessId = $u.OwningProcess; ProcessName = $procNames[[int]$u.OwningProcess]; Created = Utc $u.CreationTime
            }
        }
    }
    Add-Section 'DnsCache'  { Get-DnsClientCache | ForEach-Object { [pscustomobject]@{ Name = $_.Entry; RecordName = $_.Name; Type = "$($_.Type)"; Data = $_.Data; TTL = $_.TimeToLive } } }
    Add-Section 'ArpCache'  { Get-NetNeighbor -AddressFamily IPv4 | Where-Object { "$($_.State)" -notin 'Unreachable', 'Permanent' } | ForEach-Object { [pscustomobject]@{ IPAddress = $_.IPAddress; MAC = $_.LinkLayerAddress; State = "$($_.State)"; Interface = $_.InterfaceAlias } } }
    Add-Section 'HostsFile' {
        $h = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
        Get-Content -LiteralPath $h | Where-Object { $_.Trim() -and -not $_.Trim().StartsWith('#') } | ForEach-Object { [pscustomobject]@{ Entry = $_.Trim() } }
    }
    Add-Section 'Shares'   { Get-SmbShare | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Description = $_.Description; Special = $_.Special } } }
    Add-Section 'SmbSessions' { Get-SmbSession | ForEach-Object { [pscustomobject]@{ Client = $_.ClientComputerName; User = $_.ClientUserName; Opens = $_.NumOpens; SecondsExists = $_.SecondsExists } } }
    Add-Section 'Firewall' { Get-NetFirewallProfile | ForEach-Object { [pscustomobject]@{ Profile = $_.Name; Enabled = "$($_.Enabled)"; DefaultInbound = "$($_.DefaultInboundAction)"; DefaultOutbound = "$($_.DefaultOutboundAction)" } } }

    # --- Persistence ------------------------------------------------------------------
    Add-Section 'Services' {
        foreach ($s in (Get-CimInstance Win32_Service)) {
            $exe = Get-ExePath $s.PathName
            $f = Get-FileDetail $exe
            [pscustomobject]@{
                Name = $s.Name; DisplayName = $s.DisplayName; State = $s.State; StartMode = $s.StartMode; Account = $s.StartName
                PathName = $s.PathName; Executable = $exe; SHA256 = $f.SHA256; Signature = $f.Signature; Signer = $f.Signer
            }
        }
    }
    Add-Section 'ScheduledTasks' {
        foreach ($t in (Get-ScheduledTask)) {
            $info = $null
            try { $info = $t | Get-ScheduledTaskInfo -ErrorAction Stop } catch { $info = $null }
            $actions = @($t.Actions | ForEach-Object { if ($_.Execute) { "$($_.Execute) $($_.Arguments)".Trim() } elseif ($_.ClassId) { "COM $($_.ClassId)" } })
            $exe = if (@($t.Actions)[0].Execute) { Get-ExePath "$(@($t.Actions)[0].Execute)" } else { '' }
            $f = Get-FileDetail $exe
            [pscustomobject]@{
                TaskPath = $t.TaskPath; TaskName = $t.TaskName; State = "$($t.State)"; Author = $t.Author; Created = $t.Date
                RunAs = $t.Principal.UserId; Actions = ($actions -join ' | '); Triggers = (@($t.Triggers | ForEach-Object { ($_.CimClass.CimClassName -replace '^MSFT_Task|Trigger$', '') }) -join ', ')
                LastRun = if ($info) { Utc $info.LastRunTime } else { '' }; NextRun = if ($info) { Utc $info.NextRunTime } else { '' }
                Microsoft = ($t.TaskPath -like '\Microsoft\*'); SHA256 = $f.SHA256; Signature = $f.Signature; Signer = $f.Signer
            }
        }
    }
    Add-Section 'Autoruns' {
        $keys = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run', 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run'
        )
        foreach ($sid in @(Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^S-1-5-21-[\d-]+$' } | ForEach-Object { $_.PSChildName })) {
            $keys += "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Run"
            $keys += "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\RunOnce"
        }
        foreach ($k in $keys) {
            if (-not (Test-Path -Path $k)) { continue }
            $item = Get-ItemProperty -Path $k
            foreach ($p in $item.PSObject.Properties) {
                if ($p.Name -like 'PS*') { continue }
                $exe = Get-ExePath "$($p.Value)"; $f = Get-FileDetail $exe
                [pscustomobject]@{ Location = ($k -replace '^Registry::', ''); Name = $p.Name; Command = "$($p.Value)"; Executable = $exe; SHA256 = $f.SHA256; Signature = $f.Signature; Signer = $f.Signer }
            }
        }
        $winlogon = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -ErrorAction SilentlyContinue
        foreach ($n in 'Shell', 'Userinit', 'Taskman') {
            if ($winlogon.$n) { [pscustomobject]@{ Location = 'HKLM\...\Winlogon'; Name = $n; Command = "$($winlogon.$n)"; Executable = ''; SHA256 = ''; Signature = ''; Signer = '' } }
        }
        foreach ($ifeo in @(Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options' -ErrorAction SilentlyContinue)) {
            $dbg = (Get-ItemProperty -Path $ifeo.PSPath -ErrorAction SilentlyContinue).Debugger
            if ($dbg) { [pscustomobject]@{ Location = 'HKLM\...\Image File Execution Options'; Name = $ifeo.PSChildName; Command = "Debugger = $dbg"; Executable = (Get-ExePath $dbg); SHA256 = ''; Signature = ''; Signer = '' } }
        }
        $startup = @("$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp") + @(Get-ChildItem -Path "$env:SystemDrive\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' })
        foreach ($folder in $startup) {
            foreach ($file in @(Get-ChildItem -LiteralPath $folder -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'desktop.ini' })) {
                $f = Get-FileDetail $file.FullName
                [pscustomobject]@{ Location = $folder; Name = $file.Name; Command = $file.FullName; Executable = $file.FullName; SHA256 = $f.SHA256; Signature = $f.Signature; Signer = $f.Signer }
            }
        }
    }
    Add-Section 'WmiPersistence' {
        foreach ($c in @(Get-CimInstance -Namespace root/subscription -ClassName __EventConsumer -ErrorAction Stop)) {
            [pscustomobject]@{ Type = $c.CimClass.CimClassName; Name = $c.Name; Detail = "$($c.CommandLineTemplate)$($c.ScriptText)$($c.ExecutablePath)" }
        }
        foreach ($f in @(Get-CimInstance -Namespace root/subscription -ClassName __EventFilter -ErrorAction Stop)) {
            [pscustomobject]@{ Type = '__EventFilter'; Name = $f.Name; Detail = $f.Query }
        }
        foreach ($b in @(Get-CimInstance -Namespace root/subscription -ClassName __FilterToConsumerBinding -ErrorAction Stop)) {
            [pscustomobject]@{ Type = '__FilterToConsumerBinding'; Name = "$($b.Filter) -> $($b.Consumer)"; Detail = '' }
        }
    }

    # --- Accounts -----------------------------------------------------------------------
    Add-Section 'LocalUsers' {
        Get-LocalUser | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Enabled = $_.Enabled; LastLogon = Utc $_.LastLogon; PasswordLastSet = Utc $_.PasswordLastSet; Description = $_.Description; SID = "$($_.SID)" } }
    }
    Add-Section 'LocalAdministrators' {
        try {
            Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Type = "$($_.ObjectClass)"; Source = "$($_.PrincipalSource)"; SID = "$($_.SID)" } }
        }
        catch {
            # Get-LocalGroupMember fails when the group contains orphaned SIDs - fall back to ADSI
            $group = [ADSI]"WinNT://$env:COMPUTERNAME/$((New-Object Security.Principal.SecurityIdentifier 'S-1-5-32-544').Translate([Security.Principal.NTAccount]).Value.Split('\')[-1]),group"
            @($group.psbase.Invoke('Members')) | ForEach-Object { [pscustomobject]@{ Name = $_.GetType().InvokeMember('AdsPath', 'GetProperty', $null, $_, $null) -replace '^WinNT://', ''; Type = ''; Source = 'ADSI'; SID = '' } }
        }
    }

    # --- Event logs ------------------------------------------------------------------------
    $eventNames = @{
        4624 = 'Logon'; 4625 = 'Failed logon'; 4648 = 'Logon with explicit credentials'; 4672 = 'Admin logon (special privileges)'
        4688 = 'Process created'; 4697 = 'Service installed'; 4698 = 'Scheduled task created'; 4702 = 'Scheduled task updated'
        4720 = 'User account created'; 4722 = 'User account enabled'; 4724 = 'Password reset'; 4728 = 'Added to global group'
        4732 = 'Added to local group'; 4738 = 'User account changed'; 4756 = 'Added to universal group'; 1102 = 'Security log cleared'
        7045 = 'Service installed'; 104 = 'Event log cleared'; 6008 = 'Unexpected shutdown'
    }
    $noise = '^(SYSTEM|LOCAL SERVICE|NETWORK SERVICE|ANONYMOUS LOGON|DWM-\d+|UMFD-\d+|-)$|\$$'
    function Convert-SecurityEvent($e) {
        $f = Get-EventFields $e
        [pscustomobject]@{
            Time = Utc $e.TimeCreated; EventId = $e.Id; Event = $eventNames[[int]$e.Id]
            Account = if ($f['TargetUserName']) { $f['TargetUserName'] } else { $f['SubjectUserName'] }
            Domain = if ($f['TargetDomainName']) { $f['TargetDomainName'] } else { $f['SubjectDomainName'] }
            SubjectAccount = $f['SubjectUserName']; LogonType = $f['LogonType']; SourceIp = $f['IpAddress']; Workstation = $f['WorkstationName']
            Process = if ($f['NewProcessName']) { $f['NewProcessName'] } else { $f['ProcessName'] }; CommandLine = $f['CommandLine']
            ServiceOrTask = (@($f['ServiceName'], $f['TaskName']) | Where-Object { $_ }) -join ''; ServiceFile = $f['ServiceFileName']
            Member = $f['MemberName']; Status = (@($f['Status'], $f['SubStatus']) | Where-Object { $_ }) -join '/'
        }
    }
    Add-Section 'SecurityEvents' {
        $logons = @()
        try {
            $logons = @(Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4624, 4625, 4648; StartTime = $since } -MaxEvents $MaxEvents -ErrorAction Stop |
                        ForEach-Object { Convert-SecurityEvent $_ } | Where-Object { $_.Account -notmatch $noise -and $_.LogonType -notin '0', '5' })
        } catch { if ($_.Exception.Message -notmatch 'No events were found') { throw } }
        $other = @()
        try {
            $other = @(Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4688, 4697, 4698, 4702, 4720, 4722, 4724, 4728, 4732, 4738, 4756, 1102; StartTime = $since } -MaxEvents $MaxEvents -ErrorAction Stop |
                       ForEach-Object { Convert-SecurityEvent $_ })
        } catch { if ($_.Exception.Message -notmatch 'No events were found') { throw } }
        @($logons) + @($other) | Sort-Object Time
    }
    Add-Section 'SystemEvents' {
        try {
            Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 7045, 104, 6008; StartTime = $since } -MaxEvents $MaxEvents -ErrorAction Stop | ForEach-Object {
                $f = Get-EventFields $_
                [pscustomobject]@{ Time = Utc $_.TimeCreated; EventId = $_.Id; Event = $eventNames[[int]$_.Id]; Service = $f['ServiceName']; ImagePath = $f['ImagePath']; Account = $f['AccountName']; StartType = $f['StartType']; Log = $f['Channel'] }
            }
        } catch { if ($_.Exception.Message -notmatch 'No events were found') { throw } }
    }
    Add-Section 'PowerShellScriptBlocks' {
        try {
            Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-PowerShell/Operational'; Id = 4104; StartTime = $since } -MaxEvents ([math]::Min($MaxEvents, 2000)) -ErrorAction Stop | ForEach-Object {
                $f = Get-EventFields $_
                $text = "$($f['ScriptBlockText'])"
                [pscustomobject]@{ Time = Utc $_.TimeCreated; Level = "$($_.LevelDisplayName)"; Path = $f['Path']; ScriptBlockId = $f['ScriptBlockId']; Script = $text.Substring(0, [math]::Min(3000, $text.Length)) }
            }
        } catch { if ($_.Exception.Message -notmatch 'No events were found') { throw } }
    }
    Add-Section 'RdpEvents' {
        $rdp = @()
        foreach ($q in @(
            @{ LogName = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'; Id = 1149 },
            @{ LogName = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'; Id = 21, 24, 25 })) {
            try {
                $q.StartTime = $since
                $rdp += @(Get-WinEvent -FilterHashtable $q -MaxEvents $MaxEvents -ErrorAction Stop | ForEach-Object {
                    $f = Get-EventFields $_
                    [pscustomobject]@{ Time = Utc $_.TimeCreated; EventId = $_.Id; User = (@($f['Param1'], $f['User']) | Where-Object { $_ }) -join ''; Domain = $f['Param2']; SourceIp = (@($f['Param3'], $f['Address']) | Where-Object { $_ }) -join ''; SessionId = $f['SessionID'] }
                })
            } catch { if ($_.Exception.Message -notmatch 'No events were found') { throw } }
        }
        $rdp | Sort-Object Time
    }
    Add-Section 'DefenderEvents' {
        try {
            Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-Windows Defender/Operational'; Id = 1116, 1117, 5001, 5007, 5010, 5012; StartTime = $since } -MaxEvents $MaxEvents -ErrorAction Stop | ForEach-Object {
                $f = Get-EventFields $_
                [pscustomobject]@{ Time = Utc $_.TimeCreated; EventId = $_.Id; Threat = $f['Threat Name']; Severity = $f['Severity Name']; Path = $f['Path']; Action = $f['Action Name']; User = $f['Detection User']; Change = (@($f['Old Value'], $f['New Value']) | Where-Object { $_ }) -join ' -> ' }
            }
        } catch { if ($_.Exception.Message -notmatch 'No events were found') { throw } }
    }

    # --- Defender ---------------------------------------------------------------------------
    Add-Section 'DefenderStatus' {
        $s = Get-MpComputerStatus
        $p = Get-MpPreference
        [pscustomobject]@{
            RealTimeProtection = $s.RealTimeProtectionEnabled; AntivirusEnabled = $s.AntivirusEnabled; BehaviorMonitor = $s.BehaviorMonitorEnabled
            TamperProtection = $s.IsTamperProtected; SignatureAge = $s.AntivirusSignatureAge; SignatureVersion = $s.AntivirusSignatureVersion
            LastQuickScan = Utc $s.QuickScanEndTime; ExclusionPath = (@($p.ExclusionPath) -join '; '); ExclusionProcess = (@($p.ExclusionProcess) -join '; ')
            ExclusionExtension = (@($p.ExclusionExtension) -join '; ')
        }
    }
    Add-Section 'DefenderDetections' {
        Get-MpThreatDetection | ForEach-Object { [pscustomobject]@{ Time = Utc $_.InitialDetectionTime; ThreatId = $_.ThreatID; Resources = (@($_.Resources) -join '; '); Action = $_.ActionSuccess; Process = $_.ProcessName; User = $_.DomainUser } }
    }

    # --- Files and execution history --------------------------------------------------------
    Add-Section 'RecentFiles' {
        $exts = '\.(exe|dll|scr|com|ps1|psm1|bat|cmd|vbs|vbe|js|jse|wsf|hta|msi|lnk|iso|img|vhd|vhdx|jar)$'
        $roots = @("$env:SystemRoot\Temp", "$env:ProgramData", "$env:SystemDrive\Users\Public")
        foreach ($u in @(Get-ChildItem -Path "$env:SystemDrive\Users" -Directory -ErrorAction SilentlyContinue)) {
            $roots += @('AppData\Local\Temp', 'AppData\Roaming', 'AppData\Local', 'Downloads', 'Desktop', 'Documents') | ForEach-Object { Join-Path $u.FullName $_ }
        }
        $found = 0
        foreach ($r in ($roots | Select-Object -Unique)) {
            if (-not (Test-Path -LiteralPath $r)) { continue }
            foreach ($file in (Get-ChildItem -LiteralPath $r -File -Recurse -Force -ErrorAction SilentlyContinue)) {
                if ($file.LastWriteTime -lt $fileSince -and $file.CreationTime -lt $fileSince) { continue }
                if ($file.Name -notmatch $exts) { continue }
                $found++
                if ($found -gt 3000) { return }
                $f = Get-FileDetail $file.FullName
                [pscustomobject]@{ Path = $file.FullName; Bytes = $file.Length; Created = Utc $file.CreationTime; Modified = Utc $file.LastWriteTime; SHA256 = $f.SHA256; Signature = $f.Signature; Signer = $f.Signer }
            }
        }
    }
    Add-Section 'Prefetch' {
        Get-ChildItem -LiteralPath "$env:SystemRoot\Prefetch" -Filter *.pf -ErrorAction Stop | ForEach-Object {
            [pscustomobject]@{ Program = ($_.BaseName -replace '-[0-9A-F]{8}$', ''); File = $_.Name; FirstRun = Utc $_.CreationTime; LastRun = Utc $_.LastWriteTime }
        }
    }
    Add-Section 'PowerShellHistory' {
        foreach ($u in @(Get-ChildItem -Path "$env:SystemDrive\Users" -Directory -ErrorAction SilentlyContinue)) {
            $h = Join-Path $u.FullName 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt'
            if (Test-Path -LiteralPath $h) {
                $lines = @(Get-Content -LiteralPath $h -Tail 500 -ErrorAction SilentlyContinue)
                $n = 0
                foreach ($l in $lines) { $n++; [pscustomobject]@{ User = $u.Name; Line = $n; Command = $l } }
            }
        }
    }
    Add-Section 'InstalledSoftware' {
        foreach ($k in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*') {
            Get-ItemProperty -Path $k -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } | ForEach-Object {
                [pscustomobject]@{ Name = $_.DisplayName; Version = $_.DisplayVersion; Publisher = $_.Publisher; InstallDate = $_.InstallDate; InstallLocation = $_.InstallLocation }
            }
        }
    }

    [pscustomobject]@{ ComputerName = $env:COMPUTERNAME; Sections = $sections.ToArray(); Log = $log.ToArray() }
}

# --- Findings (run on the analyst's machine) ---------------------------------------

$userWritable = '\\Users\\|\\ProgramData\\|\\Windows\\Temp\\|\\AppData\\|\\Temp\\|\\Downloads\\|\\Public\\|\\PerfLogs\\'
$lolbinArgs   = '(?i)(-enc(odedcommand)?\s|frombase64string|\biex\b|invoke-expression|downloadstring|downloadfile|net\.webclient|invoke-webrequest.+\|\s*iex|bitsadmin.+/transfer|certutil.+(-urlcache|-decode)|mshta\s+(http|javascript|vbscript)|rundll32.+javascript|regsvr32.+/i:http|-w(indowstyle)?\s+hidden|-nop\b.+-w)'
$scriptHosts  = '(?i)\\(powershell|pwsh|cmd|mshta|wscript|cscript|rundll32|regsvr32)\.exe'
$psSuspicious = '(?i)(frombase64string|invoke-expression|\biex\b|downloadstring|net\.webclient|invoke-mimikatz|add-mppreference.+exclusion|set-mppreference.+disable|amsiutils|amsiinitfailed|reflection\.assembly::load|virtualalloc|-bxor|invoke-shellcode|sekurlsa)'
$toolsHigh    = '(?i)^(mimikatz|rubeus|sharphound|bloodhound|adfind|lazagne|nanodump|ngrok|kerbrute|safetykatz|seatbelt|certify|whisker|procdump64?)\b'
$toolsMedium  = '(?i)^(psexec|psexesvc|paexec|rclone|anydesk|screenconnect|plink|netscan|advanced_ip_scanner|angryip|winscp|megasync|7z|rar|nltest|teamviewer|atera|splashtop)\b'

function Test-PublicIp([string]$Ip) {
    if (-not $Ip -or $Ip -in '0.0.0.0', '::', '127.0.0.1', '::1', '*') { return $false }
    $addr = $null
    if (-not [System.Net.IPAddress]::TryParse($Ip, [ref]$addr)) { return $false }
    if ($addr.AddressFamily -eq 'InterNetworkV6') { return -not ($Ip -match '^(?i)(fe80|fc|fd|ff|::ffff:(10|127|192\.168|172\.(1[6-9]|2\d|3[01]))\.)') }
    $b = $addr.GetAddressBytes()
    if ($b[0] -eq 10 -or $b[0] -eq 127 -or $b[0] -ge 224 -or $b[0] -eq 0) { return $false }
    if ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) { return $false }
    if ($b[0] -eq 192 -and $b[1] -eq 168) { return $false }
    if ($b[0] -eq 169 -and $b[1] -eq 254) { return $false }
    if ($b[0] -eq 100 -and $b[1] -ge 64 -and $b[1] -le 127) { return $false }   # carrier-grade NAT
    return $true
}

function Get-TriageFindings($Sections) {
    $s = @{}; foreach ($x in $Sections) { $s[$x.Name] = @($x.Rows) }
    $findings = [System.Collections.Generic.List[object]]::new()
    $add = { param($Severity, $Category, $Finding, $Evidence, $Section) $findings.Add([pscustomobject]@{ Severity = $Severity; Category = $Category; Finding = $Finding; Evidence = $Evidence; Section = $Section }) }

    # Processes
    $procById = @{}
    foreach ($p in $s['Processes']) {
        $procById["$($p.ProcessId)"] = $p
        $unsigned = $p.Signature -and $p.Signature -ne 'Valid'
        if ($p.Path -match $userWritable) {
            & $add $(if ($unsigned) { 'High' } else { 'Medium' }) 'Process' "Running from a user-writable folder$(if ($unsigned) { ' (not signed)' })" "$($p.Name) [PID $($p.ProcessId)] $($p.Path)" 'Processes'
        }
        elseif ($unsigned -and $p.Path -and $p.Path -notmatch '(?i)\\Windows\\') {
            & $add 'Low' 'Process' "Running process is not signed ($($p.Signature))" "$($p.Name) [PID $($p.ProcessId)] $($p.Path)" 'Processes'
        }
        if ($p.CommandLine -match $lolbinArgs) { & $add 'High' 'Process' 'Suspicious command line (encoded/download/LOLBin pattern)' "$($p.Name) [PID $($p.ProcessId)]: $($p.CommandLine)" 'Processes' }
        if ($p.Name -match $toolsHigh)   { & $add 'High' 'Process' 'Known attacker tool running' "$($p.Name) $($p.Path)" 'Processes' }
        elseif ($p.Name -match $toolsMedium) { & $add 'Medium' 'Process' 'Remote access / admin tool often misused by attackers is running' "$($p.Name) $($p.Path)" 'Processes' }
    }

    # Network
    foreach ($c in @($s['NetworkConnections'] | Where-Object { $_.State -eq 'Established' -and (Test-PublicIp "$($_.RemoteAddress)") })) {
        $p = $procById["$($c.ProcessId)"]
        $risky = $p -and (($p.Path -match $userWritable) -or ($p.Signature -and $p.Signature -ne 'Valid') -or ($p.Path -match $scriptHosts))
        if ($risky) { & $add 'High' 'Network' 'Internet connection from a suspicious process' "$($c.ProcessName) [PID $($c.ProcessId)] -> $($c.RemoteAddress):$($c.RemotePort)" 'NetworkConnections' }
    }
    foreach ($c in @($s['NetworkConnections'] | Where-Object { $_.State -eq 'Listen' -and $_.Protocol -eq 'TCP' -and "$($_.LocalAddress)" -notin '127.0.0.1', '::1' })) {
        $p = $procById["$($c.ProcessId)"]
        if ($p -and $p.Path -match $userWritable) { & $add 'High' 'Network' 'Listening port opened by a process in a user-writable folder' "$($c.ProcessName) listening on $($c.LocalAddress):$($c.LocalPort)" 'NetworkConnections' }
    }

    # Services
    foreach ($svc in $s['Services']) {
        if ($svc.Executable -match $userWritable) { & $add 'High' 'Persistence' 'Service runs from a user-writable folder' "$($svc.Name): $($svc.PathName)" 'Services' }
        elseif ($svc.PathName -match $scriptHosts -and $svc.PathName -match '(?i)(powershell|cmd\.exe\s*/c|mshta|wscript|cscript)') { & $add 'High' 'Persistence' 'Service runs a script host' "$($svc.Name): $($svc.PathName)" 'Services' }
        elseif ($svc.Signature -and $svc.Signature -ne 'Valid' -and $svc.Executable -notmatch '(?i)\\Windows\\') { & $add 'Medium' 'Persistence' "Service binary not signed ($($svc.Signature))" "$($svc.Name): $($svc.Executable)" 'Services' }
        if ($svc.PathName -match '^[A-Za-z]:\\[^"]*\s[^"]*\.exe' -and $svc.PathName -notmatch '^"' -and $svc.StartMode -eq 'Auto') { & $add 'Low' 'Persistence' 'Unquoted service path with spaces (privilege escalation risk)' "$($svc.Name): $($svc.PathName)" 'Services' }
    }

    # Scheduled tasks
    foreach ($t in @($s['ScheduledTasks'] | Where-Object { -not $_.Microsoft -or $_.Actions -match $userWritable })) {
        if ($t.Actions -match $userWritable -or $t.Actions -match $lolbinArgs) { & $add 'High' 'Persistence' 'Scheduled task runs from a user-writable folder or with a suspicious command' "$($t.TaskPath)$($t.TaskName): $($t.Actions)" 'ScheduledTasks' }
        elseif ($t.Actions -match $scriptHosts) { & $add 'Medium' 'Persistence' 'Non-Microsoft scheduled task runs a script host' "$($t.TaskPath)$($t.TaskName): $($t.Actions)" 'ScheduledTasks' }
    }

    # Autoruns
    foreach ($a in $s['Autoruns']) {
        if ($a.Location -like '*Image File Execution Options*') { & $add 'High' 'Persistence' 'Image File Execution Options debugger set (can hijack programs)' "$($a.Name): $($a.Command)" 'Autoruns'; continue }
        if ($a.Location -like '*Winlogon*') {
            $expected = @{ Shell = '^explorer\.exe,?$'; Userinit = '(?i)^C:\\Windows\\system32\\userinit\.exe,?$'; Taskman = '^$' }
            if ($a.Command -notmatch $expected[$a.Name]) { & $add 'High' 'Persistence' "Winlogon $($a.Name) has been changed" $a.Command 'Autoruns' }
            continue
        }
        if ($a.Command -match $userWritable -or $a.Command -match $lolbinArgs -or $a.Command -match $scriptHosts) { & $add 'High' 'Persistence' 'Startup entry runs from a user-writable folder or a script host' "$($a.Location) $($a.Name): $($a.Command)" 'Autoruns' }
        elseif ($a.Signature -and $a.Signature -ne 'Valid') { & $add 'Medium' 'Persistence' "Startup entry not signed ($($a.Signature))" "$($a.Location) $($a.Name): $($a.Command)" 'Autoruns' }
    }

    # WMI
    foreach ($w in @($s['WmiPersistence'] | Where-Object { $_.Type -match 'CommandLineEventConsumer|ActiveScriptEventConsumer' })) {
        & $add 'High' 'Persistence' "WMI event consumer ($($w.Type)) - a common fileless persistence method" "$($w.Name): $($w.Detail)" 'WmiPersistence'
    }

    # Accounts
    foreach ($m in @($s['LocalAdministrators'] | Where-Object { "$($_.SID)" -notmatch '-500$|-512$|-519$' -and $_.Name -notmatch '(?i)\\(Administrator|Domain Admins|Enterprise Admins)$' })) {
        & $add 'Low' 'Accounts' 'Additional member of the local Administrators group - check it is expected' "$($m.Name) ($($m.Type), $($m.Source))" 'LocalAdministrators'
    }

    # Events
    $ev = $s['SecurityEvents']
    foreach ($e in @($ev | Where-Object { $_.EventId -eq 1102 })) { & $add 'Critical' 'Anti-forensics' 'Security log was cleared' "$($e.Time) by $($e.SubjectAccount)" 'SecurityEvents' }
    foreach ($e in @($s['SystemEvents'] | Where-Object { $_.EventId -eq 104 })) { & $add 'Critical' 'Anti-forensics' 'An event log was cleared' "$($e.Time) $($e.Log) by $($e.Account)" 'SystemEvents' }
    foreach ($e in @($ev | Where-Object { $_.EventId -eq 4720 })) { & $add 'Medium' 'Accounts' 'User account created' "$($e.Time) $($e.Account) by $($e.SubjectAccount)" 'SecurityEvents' }
    foreach ($e in @($ev | Where-Object { $_.EventId -in 4728, 4732, 4756 })) { & $add $(if ($e.Account -match '(?i)admin') { 'High' } else { 'Medium' }) 'Accounts' "Member added to group '$($e.Account)'" "$($e.Time) $($e.Member) by $($e.SubjectAccount)" 'SecurityEvents' }
    foreach ($e in @(@($ev | Where-Object { $_.EventId -eq 4697 }) + @($s['SystemEvents'] | Where-Object { $_.EventId -eq 7045 }))) {
        $path = "$($e.ServiceFile)$($e.ImagePath)"
        & $add $(if ($path -match $userWritable -or $path -match $scriptHosts -or $path -match '(?i)psexesvc|\\\\') { 'High' } else { 'Medium' }) 'Persistence' 'Service installed' "$($e.Time) $($e.ServiceOrTask)$($e.Service): $path" $(if ($e.EventId -eq 7045) { 'SystemEvents' } else { 'SecurityEvents' })
    }
    foreach ($e in @($ev | Where-Object { $_.EventId -eq 4698 })) { & $add 'Medium' 'Persistence' 'Scheduled task created' "$($e.Time) $($e.ServiceOrTask) by $($e.SubjectAccount)" 'SecurityEvents' }
    foreach ($e in @($ev | Where-Object { $_.EventId -eq 4688 -and $_.CommandLine -match $lolbinArgs })) { & $add 'High' 'Execution' 'Suspicious process command line in the event log' "$($e.Time) $($e.Process): $($e.CommandLine)" 'SecurityEvents' }
    foreach ($g in @($ev | Where-Object { $_.EventId -eq 4625 } | Group-Object SourceIp | Where-Object { $_.Count -ge 20 -and $_.Name -and $_.Name -ne '-' })) {
        $accounts = @($g.Group | Select-Object -ExpandProperty Account -Unique)
        & $add $(if ($accounts.Count -ge 5) { 'High' } else { 'Medium' }) 'Credential access' $(if ($accounts.Count -ge 5) { 'Possible password spray (many accounts from one source)' } else { 'Possible brute force (many failed logons from one source)' }) "$($g.Name): $($g.Count) failures, $($accounts.Count) account(s): $(($accounts | Select-Object -First 5) -join ', ')" 'SecurityEvents'
    }
    foreach ($e in @($ev | Where-Object { $_.EventId -eq 4624 -and $_.LogonType -eq '10' -and (Test-PublicIp "$($_.SourceIp)") })) { & $add 'High' 'Initial access' 'Remote Desktop logon from a public IP address' "$($e.Time) $($e.Domain)\$($e.Account) from $($e.SourceIp)" 'SecurityEvents' }
    foreach ($e in @($s['RdpEvents'] | Where-Object { $_.EventId -eq 1149 -and (Test-PublicIp "$($_.SourceIp)") })) { & $add 'High' 'Initial access' 'Remote Desktop connection from a public IP address' "$($e.Time) $($e.User) from $($e.SourceIp)" 'RdpEvents' }
    foreach ($b in @($s['PowerShellScriptBlocks'] | Where-Object { $_.Script -match $psSuspicious } | Select-Object -First 25)) {
        $snippet = "$($b.Script)" -replace '\s+', ' '
        & $add 'High' 'Execution' 'Suspicious PowerShell script block' "$($b.Time): $($snippet.Substring(0, [math]::Min(200, $snippet.Length)))" 'PowerShellScriptBlocks'
    }
    foreach ($h in @($s['PowerShellHistory'] | Where-Object { $_.Command -match $lolbinArgs -or $_.Command -match $psSuspicious } | Select-Object -First 25)) {
        & $add 'High' 'Execution' 'Suspicious command in PowerShell history' "$($h.User) line $($h.Line): $($h.Command)" 'PowerShellHistory'
    }

    # Defender
    foreach ($d in $s['DefenderStatus']) {
        if ("$($d.RealTimeProtection)" -eq 'False') { & $add 'High' 'Defence evasion' 'Defender real-time protection is off' '' 'DefenderStatus' }
        if ("$($d.TamperProtection)" -eq 'False')   { & $add 'Low' 'Defence evasion' 'Defender tamper protection is off' '' 'DefenderStatus' }
        $ex = (@($d.ExclusionPath, $d.ExclusionProcess, $d.ExclusionExtension) | Where-Object { $_ }) -join '; '
        if ($ex) { & $add 'Medium' 'Defence evasion' 'Defender exclusions are set - check they are expected' $ex 'DefenderStatus' }
        if ([int]"0$($d.SignatureAge)" -gt 7) { & $add 'Low' 'Defence evasion' "Defender signatures are $($d.SignatureAge) days old" '' 'DefenderStatus' }
    }
    foreach ($d in @($s['DefenderEvents'] | Where-Object { $_.EventId -in 1116, 1117 })) { & $add 'High' 'Malware' "Defender detection: $($d.Threat)" "$($d.Time) $($d.Path) (action: $($d.Action))" 'DefenderEvents' }
    foreach ($d in @($s['DefenderEvents'] | Where-Object { $_.EventId -in 5001, 5010, 5012 })) { & $add 'High' 'Defence evasion' 'Defender protection was turned off' "$($d.Time) event $($d.EventId)" 'DefenderEvents' }
    foreach ($d in @($s['DefenderEvents'] | Where-Object { $_.EventId -eq 5007 -and $_.Change -match '(?i)exclusion' })) { & $add 'High' 'Defence evasion' 'Defender exclusion added' "$($d.Time) $($d.Change)" 'DefenderEvents' }

    # Files and execution history
    foreach ($f in @($s['RecentFiles'] | Where-Object { (Split-Path -Leaf "$($_.Path)") -match $toolsHigh })) { & $add 'High' 'Tools' 'Known attacker tool on disk' "$($f.Path) $($f.SHA256)" 'RecentFiles' }
    foreach ($p in @($s['Prefetch'] | Where-Object { $_.Program -match $toolsHigh })) { & $add 'High' 'Tools' 'Known attacker tool has been run (prefetch)' "$($p.Program) last run $($p.LastRun)" 'Prefetch' }
    foreach ($p in @($s['Prefetch'] | Where-Object { $_.Program -match $toolsMedium })) { & $add 'Medium' 'Tools' 'Remote access / admin tool has been run (prefetch) - check it is expected' "$($p.Program) last run $($p.LastRun)" 'Prefetch' }
    $unsignedNew = @($s['RecentFiles'] | Where-Object { $_.Path -match '(?i)\.(exe|dll|scr)$' -and $_.Signature -and $_.Signature -ne 'Valid' })
    if ($unsignedNew.Count) { & $add 'Medium' 'Files' "$($unsignedNew.Count) recently created/modified unsigned executable(s) in user-writable folders" (($unsignedNew | Select-Object -First 5 | ForEach-Object { $_.Path }) -join '; ') 'RecentFiles' }

    foreach ($h in $s['HostsFile']) {
        if ($h.Entry -match '(?i)microsoft|windowsupdate|office|outlook|live\.com|okta|google|apple|bank|paypal|defender|sophos|crowdstrike|sentinelone|eset|kaspersky|virustotal') {
            & $add 'High' 'Network' 'Hosts file redirects a sign-in, update or security site' $h.Entry 'HostsFile'
        }
        else { & $add 'Low' 'Network' 'Entry in the hosts file - check it is expected' $h.Entry 'HostsFile' }
    }
    foreach ($fw in @($s['Firewall'] | Where-Object { "$($_.Enabled)" -eq 'False' })) { & $add 'Medium' 'Defence evasion' "Windows Firewall is off for the $($fw.Profile) profile" '' 'Firewall' }

    $order = @{ Critical = 0; High = 1; Medium = 2; Low = 3 }
    return @($findings | Sort-Object @{ Expression = { $order[$_.Severity] } }, Category)
}

# --- Main ---------------------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    $targets = if ($ComputerName.Count) { $ComputerName } else { @([System.Environment]::MachineName) }
    $results = @()

    foreach ($target in $targets) {
        $isLocal = $target -in '.', 'localhost', [System.Environment]::MachineName
        $start   = Get-Date
        Write-Host "Collecting from $target (last $Hours hours of events, files changed in the last $FileDays days)..."

        try {
            if ($isLocal) {
                if ($script:IsWindowsHost) {
                    $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                    if (-not $admin) { Write-Warning 'Not running as administrator - the Security log and some other sections will be missing. Re-run from an elevated PowerShell.' }
                }
                else { Write-Warning 'This is not a Windows machine - most sections will be empty.' }
                $data = & $collector $Hours $FileDays (-not $SkipHashes) $MaxEvents
            }
            else {
                $icm = @{ ComputerName = $target; ScriptBlock = $collector; ArgumentList = @($Hours, $FileDays, (-not $SkipHashes), $MaxEvents); ErrorAction = 'Stop' }
                if ($Credential) { $icm.Credential = $Credential }
                $data = Invoke-Command @icm
            }
        }
        catch {
            Write-Error "Could not collect from ${target}: $($_.Exception.Message)"
            continue
        }

        $hostName = if ($data.ComputerName) { $data.ComputerName } else { $target }
        $name     = (@($CaseId, 'Triage', $hostName, (Get-Date -Format 'yyyyMMdd_HHmmss')) | Where-Object { $_ }) -join '_'
        $folder   = New-EvidenceFolder -Config $config -OutputPath $OutputPath -Name $name
        Write-Host "  Writing evidence to $folder"

        $n = 0
        $counts = [ordered]@{}
        foreach ($sec in $data.Sections) {
            $n++
            $counts[$sec.Name] = Export-EvidenceCsv -Folder $folder -Name ('{0:D2}_{1}' -f $n, $sec.Name) -Rows $sec.Rows
        }
        $findings = Get-TriageFindings $data.Sections
        [void](Export-EvidenceCsv -Folder $folder -Name '00_Findings' -Rows $findings)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_CollectionLog' -Rows $data.Log)

        # Summary page
        $sys = @(($data.Sections | Where-Object { $_.Name -eq 'SystemInfo' }).Rows)[0]
        $sevCounts = [ordered]@{}; foreach ($sv in 'Critical', 'High', 'Medium', 'Low') { $sevCounts[$sv] = @($findings | Where-Object { $_.Severity -eq $sv }).Count }
        $intro = @(
            "Host: $hostName$(if ($sys) { " - $($sys.OS) $($sys.Version), domain $($sys.Domain), time zone $($sys.TimeZone)" })",
            "Collected: $(ConvertTo-UtcText $start) by $env:USERDOMAIN\$env:USERNAME$(if ($CaseId) { " - case $CaseId" }). All times are UTC.",
            "Findings: $($sevCounts.Critical) critical, $($sevCounts.High) high, $($sevCounts.Medium) medium, $($sevCounts.Low) low. Findings are leads to investigate, not proof of compromise."
        )
        $sectionRows = @($data.Log | ForEach-Object { [pscustomobject]@{ Section = $_.Section; Rows = $_.Rows; Status = $(if ($_.Status -eq 'OK') { 'OK' } else { 'Failed' }); Seconds = $_.Seconds; Error = $_.Error } })
        $html = New-HtmlReport -Title "Windows triage - $hostName" -IntroLines $intro -Sections @(
            @{ Heading = "Findings ($($findings.Count))"; Html = ConvertTo-HtmlTable -Rows $findings -Columns @('Severity', 'Category', 'Finding', 'Evidence', 'Section') -StatusColumn 'Severity' -EmptyText 'No automatic findings. Review the CSV files - absence of findings does not mean the machine is clean.' }
            @{ Heading = 'What was collected'; Html = ConvertTo-HtmlTable -Rows $sectionRows -Columns @('Section', 'Rows', 'Status', 'Seconds', 'Error') -StatusColumn 'Status' }
        )
        [System.IO.File]::WriteAllText((Join-Path $folder '00_Summary.html'), $html, $script:Utf8NoBom)

        $package = Complete-EvidencePackage -Folder $folder
        $results += $package

        Write-Host ""
        Write-Host ("{0}: {1} critical, {2} high, {3} medium, {4} low finding(s)" -f $hostName, $sevCounts.Critical, $sevCounts.High, $sevCounts.Medium, $sevCounts.Low) -ForegroundColor $(if ($sevCounts.Critical + $sevCounts.High) { 'Red' } elseif ($sevCounts.Medium) { 'Yellow' } else { 'Green' })
        $findings | Where-Object { $_.Severity -in 'Critical', 'High' } | Select-Object -First 10 | ForEach-Object { Write-Host ("  {0,-8} {1}: {2}" -f $_.Severity, $_.Finding, $_.Evidence) }
        $failed = @($data.Log | Where-Object { $_.Status -ne 'OK' })
        if ($failed.Count) { Write-Host ("  {0} section(s) could not be collected - see 00_CollectionLog.csv: {1}" -f $failed.Count, (($failed | Select-Object -ExpandProperty Section) -join ', ')) -ForegroundColor Yellow }
        Write-Host "  Evidence : $($package.Zip)"
        Write-Host "  SHA-256  : $($package.Sha256)  (record this in your case notes)"
        Write-Host ""
    }
}
