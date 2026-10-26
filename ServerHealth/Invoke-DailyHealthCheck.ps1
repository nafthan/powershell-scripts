<#
.SYNOPSIS
    Daily health check for Windows servers, saved as an HTML report.
.DESCRIPTION
    Checks uptime, pending reboots, disk space, stopped services, last update and event log errors.
.EXAMPLE
    .\Invoke-DailyHealthCheck.ps1 -ComputerName SERVER01, SERVER02
.EXAMPLE
    .\Invoke-DailyHealthCheck.ps1 -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: admin rights and WinRM (or DCOM) on the servers
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$ComputerName,

    [switch]$SendEmail,
    [switch]$OnlyEmailOnIssues,

    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('ComputerName')) {
    $ComputerName = Get-Setting $config 'Servers' 'ComputerName' @([System.Environment]::MachineName)
}

$ComputerName = Split-ListValue $ComputerName

$thresholds = @{
    DiskWarning     = [double](Get-Setting $config 'HealthCheck' 'DiskWarningPercent' 15)
    DiskCritical    = [double](Get-Setting $config 'HealthCheck' 'DiskCriticalPercent' 5)
    UptimeDays      = [int](Get-Setting $config 'HealthCheck' 'UptimeWarningDays' 45)
    PatchDays       = [int](Get-Setting $config 'HealthCheck' 'PatchWarningDays' 35)
    ErrorEvents     = [int](Get-Setting $config 'HealthCheck' 'ErrorEventWarning' 50)
    CheckFileRename = [bool](Get-Setting $config 'HealthCheck' 'CheckPendingFileRename' $false)
    ExcludeServices = @(Get-Setting $config 'HealthCheck' 'ExcludeServices' @('gupdate', 'gupdatem', 'edgeupdate', 'edgeupdatem', 'MapsBroker', 'RemoteRegistry', 'sppsvc', 'CDPSvc', 'OneSyncSvc*', 'tiledatamodelsvc', 'WbioSrvc', 'DoSvc', 'UsoSvc', 'BITS', 'ShellHWDetection', 'TrustedInstaller', 'wuauserv'))
}

# --- Checks ---------------------------------------------------------------

function Test-IsLocalComputer([string]$Name) {
    $n = $Name.Split('.')[0]
    return ($n -eq 'localhost' -or $n -eq '.' -or $n -eq [System.Environment]::MachineName)
}

function New-HealthCimSession([string]$Name) {
    try {
        return New-CimSession -ComputerName $Name -OperationTimeoutSec 30 -ErrorAction Stop
    }
    catch {
        $dcom = New-CimSessionOption -Protocol Dcom
        return New-CimSession -ComputerName $Name -SessionOption $dcom -OperationTimeoutSec 30 -ErrorAction Stop
    }
}

function Test-RegistryKey($Session, [string]$SubKey) {
    $r = Invoke-CimMethod -CimSession $Session -Namespace 'root/cimv2' -ClassName 'StdRegProv' -MethodName 'EnumKey' `
         -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $SubKey } -ErrorAction Stop   # HKLM
    return ($r.ReturnValue -eq 0)
}

function Get-PendingRebootReasons($Session) {
    $reasons = @()
    if (Test-RegistryKey $Session 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons += 'Component servicing' }
    if (Test-RegistryKey $Session 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')  { $reasons += 'Windows Update' }
    if ($thresholds.CheckFileRename) {
        $r = Invoke-CimMethod -CimSession $Session -Namespace 'root/cimv2' -ClassName 'StdRegProv' -MethodName 'GetMultiStringValue' `
             -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = 'SYSTEM\CurrentControlSet\Control\Session Manager'; sValueName = 'PendingFileRenameOperations' } -ErrorAction Stop
        if ($r.ReturnValue -eq 0 -and $r.sValue) { $reasons += 'File rename operations' }
    }
    return $reasons
}

function ConvertTo-HotfixDate([string]$Value) {
    if (-not $Value) { return $null }
    $formats = @('M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd', 'd/M/yyyy', 'dd/MM/yyyy')
    foreach ($f in $formats) {
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($Value, $f, [System.Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$d)) { return $d }
    }
    return $null
}

function Get-ServerHealth([string]$Name) {
    $h = [pscustomobject]@{
        Server          = $Name
        Status          = 'OK'
        OS              = ''
        LastBoot        = $null
        UptimeDays      = $null
        PendingReboot   = ''
        LastPatch       = $null
        DaysSincePatch  = $null
        LowDisks        = ''
        StoppedServices = ''
        Errors24h       = ''
        TopErrorSources = ''
        Issues          = [System.Collections.Generic.List[string]]::new()
        Disks           = @()
        Services        = @()
    }
    $raise = {
        param([string]$Level, [string]$Text)
        $h.Issues.Add($Text)
        if ($Level -eq 'Critical') { $h.Status = 'Critical' }
        elseif ($Level -eq 'Warning' -and $h.Status -eq 'OK') { $h.Status = 'Warning' }
    }

    $isLocal = Test-IsLocalComputer $Name
    if (-not $isLocal -and -not (Test-Connection -ComputerName $Name -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        # Ping can be blocked - only give up if CIM fails too
        $pingFailed = $true
    } else { $pingFailed = $false }

    try {
        $session = New-HealthCimSession $Name
    }
    catch {
        $h.Status = 'Critical'
        $h.Issues.Add($(if ($pingFailed) { 'Unreachable (no ping, no remote management)' } else { "Remote management failed: $($_.Exception.Message)" }))
        return $h
    }

    try {
        # OS and uptime
        try {
            $os = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem -ErrorAction Stop
            $h.OS         = $os.Caption -replace 'Microsoft ', ''
            $h.LastBoot   = $os.LastBootUpTime
            $h.UptimeDays = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)
            if ($h.UptimeDays -gt $thresholds.UptimeDays) { & $raise 'Warning' "Not rebooted for $([int]$h.UptimeDays) days" }
        } catch { & $raise 'Warning' "OS check failed: $($_.Exception.Message)" }

        # Pending reboot
        try {
            $reasons = @(Get-PendingRebootReasons $session)
            $h.PendingReboot = if ($reasons.Count) { 'Yes (' + ($reasons -join ', ') + ')' } else { 'No' }
            if ($reasons.Count) { & $raise 'Warning' "Reboot pending: $($reasons -join ', ')" }
        } catch { $h.PendingReboot = 'Unknown' }

        # Disks
        try {
            $disks = @(Get-CimInstance -CimSession $session -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)
            $h.Disks = @(foreach ($d in $disks) {
                if (-not $d.Size) { continue }
                $pct = [math]::Round(($d.FreeSpace / $d.Size) * 100, 1)
                $st  = if ($pct -lt $thresholds.DiskCritical) { 'Critical' } elseif ($pct -lt $thresholds.DiskWarning) { 'Warning' } else { 'OK' }
                [pscustomobject]@{
                    Server = $Name; Drive = $d.DeviceID; Label = $d.VolumeName
                    SizeGB = [math]::Round($d.Size / 1GB, 1); FreeGB = [math]::Round($d.FreeSpace / 1GB, 1)
                    FreePercent = $pct; Status = $st
                }
            })
            $low = @($h.Disks | Where-Object { $_.Status -ne 'OK' })
            $h.LowDisks = if ($low.Count) { ($low | ForEach-Object { "$($_.Drive) $($_.FreePercent)%" }) -join ', ' } else { 'None' }
            foreach ($d in $low) { & $raise $d.Status "Disk $($d.Drive) has $($d.FreePercent)% free ($($d.FreeGB) GB)" }
        } catch { & $raise 'Warning' "Disk check failed: $($_.Exception.Message)" }

        # Automatic services not running
        try {
            $svc = @(Get-CimInstance -CimSession $session -ClassName Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" -ErrorAction Stop)
            $svc = @($svc | Where-Object {
                $n = $_.Name
                -not ($thresholds.ExcludeServices | Where-Object { $n -like $_ })
            })
            $h.Services = @($svc | ForEach-Object {
                [pscustomobject]@{ Server = $Name; Name = $_.Name; DisplayName = $_.DisplayName; State = $_.State; ExitCode = $_.ExitCode }
            })
            $h.StoppedServices = if ($svc.Count) { ($svc.Name -join ', ') } else { 'None' }
            if ($svc.Count) { & $raise 'Warning' "$($svc.Count) automatic service(s) not running: $($svc.Name -join ', ')" }
        } catch { & $raise 'Warning' "Service check failed: $($_.Exception.Message)" }

        # Last update installed
        try {
            $qfe = @(Get-CimInstance -CimSession $session -ClassName Win32_QuickFixEngineering -ErrorAction Stop)
            $dates = @($qfe | ForEach-Object { ConvertTo-HotfixDate "$($_.InstalledOn)" } | Where-Object { $_ })
            if ($dates.Count) {
                $h.LastPatch      = ($dates | Sort-Object -Descending | Select-Object -First 1)
                $h.DaysSincePatch = [int]((Get-Date).Date - $h.LastPatch.Date).TotalDays
                if ($h.DaysSincePatch -gt $thresholds.PatchDays) { & $raise 'Warning' "No update installed for $($h.DaysSincePatch) days" }
            }
        } catch { Write-Verbose "Update history unavailable on ${Name}: $($_.Exception.Message)" }
    }
    finally {
        if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
    }

    # Event log errors (critical + error) in the last 24 hours
    try {
        $filter = @{ LogName = 'System', 'Application'; Level = 1, 2; StartTime = (Get-Date).AddHours(-24) }
        $evParams = @{ FilterHashtable = $filter; ErrorAction = 'Stop'; MaxEvents = 5000 }
        if (-not $isLocal) { $evParams.ComputerName = $Name }
        $events = @(Get-WinEvent @evParams)
        $h.Errors24h = $events.Count
        $h.TopErrorSources = ($events | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 3 |
                              ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
        if ($events.Count -gt $thresholds.ErrorEvents) { & $raise 'Warning' "$($events.Count) errors in the System/Application logs in 24h" }
        if (@($events | Where-Object { $_.Level -eq 1 }).Count) { & $raise 'Warning' "$(@($events | Where-Object { $_.Level -eq 1 }).Count) critical event(s) in 24h" }
    }
    catch {
        if ($_.Exception.Message -match 'No events were found') { $h.Errors24h = 0; $h.TopErrorSources = '' }
        else { $h.Errors24h = 'Unavailable' }
    }

    return $h
}

# --- Report ---------------------------------------------------------------

function New-HealthReportHtml([object[]]$Results) {
    $critical = @($Results | Where-Object { $_.Status -eq 'Critical' }).Count
    $warning  = @($Results | Where-Object { $_.Status -eq 'Warning' }).Count
    $ok       = @($Results | Where-Object { $_.Status -eq 'OK' }).Count

    $summaryRows = $Results | Sort-Object @{ Expression = { switch ($_.Status) { 'Critical' { 0 } 'Warning' { 1 } default { 2 } } } }, Server |
        ForEach-Object {
            [pscustomobject]@{
                Server            = $_.Server
                Status            = $_.Status
                OS                = $_.OS
                'Uptime (days)'   = $_.UptimeDays
                'Pending reboot'  = $_.PendingReboot
                'Last update'     = if ($_.LastPatch) { $_.LastPatch.ToString('yyyy-MM-dd') } else { '' }
                'Low disks'       = $_.LowDisks
                'Errors (24h)'    = $_.Errors24h
            }
        }

    $issueRows = foreach ($r in $Results) {
        foreach ($i in $r.Issues) { [pscustomobject]@{ Server = $r.Server; Status = $r.Status; Issue = $i } }
    }
    $diskRows = @($Results | ForEach-Object { $_.Disks }) | Sort-Object FreePercent
    $svcRows  = @($Results | ForEach-Object { $_.Services })
    $evRows   = $Results | Where-Object { $_.TopErrorSources } |
                ForEach-Object { [pscustomobject]@{ Server = $_.Server; 'Errors (24h)' = $_.Errors24h; 'Top sources' = $_.TopErrorSources } }

    $intro = @(
        "$($Results.Count) server(s) checked: $critical critical, $warning warning, $ok OK.",
        ("Thresholds: disk warning below {0}% free, critical below {1}%; uptime over {2} days; no update for over {3} days." -f
            $thresholds.DiskWarning, $thresholds.DiskCritical, $thresholds.UptimeDays, $thresholds.PatchDays)
    )
    $sections = @(
        @{ Heading = 'Overview'; Html = ConvertTo-HtmlTable -Rows @($summaryRows) -Columns @('Server', 'Status', 'OS', 'Uptime (days)', 'Pending reboot', 'Last update', 'Low disks', 'Errors (24h)') -StatusColumn 'Status' }
        @{ Heading = 'Issues'; Html = ConvertTo-HtmlTable -Rows @($issueRows) -Columns @('Server', 'Status', 'Issue') -StatusColumn 'Status' -EmptyText 'No issues found.' }
        @{ Heading = 'Disks'; Html = ConvertTo-HtmlTable -Rows @($diskRows) -Columns @('Server', 'Drive', 'Label', 'SizeGB', 'FreeGB', 'FreePercent', 'Status') -StatusColumn 'Status' }
        @{ Heading = 'Automatic services not running'; Html = ConvertTo-HtmlTable -Rows @($svcRows) -Columns @('Server', 'Name', 'DisplayName', 'State', 'ExitCode') -EmptyText 'All automatic services are running.' }
        @{ Heading = 'Event log errors (last 24 hours)'; Html = ConvertTo-HtmlTable -Rows @($evRows) -Columns @('Server', 'Errors (24h)', 'Top sources') -EmptyText 'No errors logged.' }
    )
    $title = "Daily health check - $(Get-Date -Format 'dddd d MMMM yyyy')"
    return @{
        Html     = New-HtmlReport -Title $title -IntroLines $intro -Sections $sections
        Critical = $critical
        Warning  = $warning
        OK       = $ok
    }
}

# --- Main -----------------------------------------------------------------
# (Skipped when the file is dot-sourced, so the functions above can be reused or tested.)

if ($MyInvocation.InvocationName -ne '.') {
    $reportPath = New-ReportPath -Config $config -Prefix 'HealthCheck' -Extension '.html' -OutputFile $OutputFile

    $results = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($server in $ComputerName) {
        $i++
        Write-Progress -Activity 'Daily health check' -Status "$server ($i of $($ComputerName.Count))" -PercentComplete (($i - 1) / $ComputerName.Count * 100)
        $r = Get-ServerHealth $server
        $results.Add($r)
        $colour = switch ($r.Status) { 'Critical' { 'Red' } 'Warning' { 'Yellow' } default { 'Green' } }
        Write-Host ("{0,-25} {1}" -f $server, $r.Status) -ForegroundColor $colour
        foreach ($issue in $r.Issues) { Write-Host "    - $issue" -ForegroundColor $colour }
    }
    Write-Progress -Activity 'Daily health check' -Completed

    $report = New-HealthReportHtml $results.ToArray()
    [System.IO.File]::WriteAllText($reportPath, $report.Html, $script:Utf8NoBom)

    Write-Host ""
    Write-Host ("Critical: {0}  Warning: {1}  OK: {2}" -f $report.Critical, $report.Warning, $report.OK)
    Write-Host "Report  : $reportPath"

    if ($SendEmail) {
        if ($OnlyEmailOnIssues -and ($report.Critical + $report.Warning) -eq 0) {
            Write-Host 'All servers OK - no email sent (-OnlyEmailOnIssues).'
        }
        else {
            $state = if ($report.Critical) { 'CRITICAL' } elseif ($report.Warning) { 'Warning' } else { 'All OK' }
            $subject = "[$state] Daily health check - $($report.Critical) critical, $($report.Warning) warning, $($report.OK) OK"
            if (Send-ToolkitMail -Config $config -Subject $subject -HtmlBody $report.Html) { Write-Host 'Email sent.' }
        }
    }
}
