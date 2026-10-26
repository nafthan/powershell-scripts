<#
.SYNOPSIS
    Checks which TCP ports are open on your servers.
.DESCRIPTION
    Compares open ports with the allowed list in config.psd1 and flags risky services. Only scan systems you're responsible for.
.EXAMPLE
    .\Get-OpenPortReport.ps1 -ComputerName SERVER01, SERVER02
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$ComputerName,

    [string[]]$Port,

    [ValidateRange(200, 60000)]
    [int]$TimeoutMs,

    [switch]$SendEmail,
    [switch]$AlwaysEmail,

    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

# --- Settings -------------------------------------------------------------

$defaultPorts = @{
    '21' = 'FTP'; '22' = 'SSH'; '23' = 'Telnet'; '25' = 'SMTP'; '53' = 'DNS'; '80' = 'HTTP'
    '88' = 'Kerberos'; '135' = 'RPC'; '139' = 'NetBIOS'; '389' = 'LDAP'; '443' = 'HTTPS'; '445' = 'SMB'
    '636' = 'LDAPS'; '1433' = 'SQL Server'; '3306' = 'MySQL'; '3389' = 'RDP'; '5900' = 'VNC'
    '5985' = 'WinRM HTTP'; '5986' = 'WinRM HTTPS'; '8080' = 'HTTP alt'
}

if (-not $PSBoundParameters.ContainsKey('ComputerName')) {
    $ComputerName = Get-Setting $config 'PortCheck' 'ComputerName' $null
    if (-not $ComputerName) { $ComputerName = Get-Setting $config 'Servers' 'ComputerName' @([System.Environment]::MachineName) }
}
$ComputerName = Split-ListValue $ComputerName
if (-not $PSBoundParameters.ContainsKey('TimeoutMs')) { $TimeoutMs = Get-Setting $config 'PortCheck' 'TimeoutMs' 3000 }

$portNames = @{}
$cfgPorts  = Get-Setting $config 'PortCheck' 'Ports' $defaultPorts
foreach ($k in $cfgPorts.Keys) { $portNames[[int]$k] = "$($cfgPorts[$k])" }

if ($PSBoundParameters.ContainsKey('Port')) {
    $ports = @(Split-ListValue $Port | ForEach-Object { [int]$_ })
} else {
    $ports = @($portNames.Keys)
}
$ports = @($ports | Where-Object { $_ -ge 1 -and $_ -le 65535 } | Sort-Object -Unique)
if (-not $ports.Count)        { throw 'No ports to check. Use -Port or set PortCheck.Ports in config.psd1.' }
if (-not $ComputerName.Count) { throw 'No servers to check. Use -ComputerName or set PortCheck.ComputerName / Servers.ComputerName.' }

$riskyPorts = @(Get-Setting $config 'PortCheck' 'RiskyPorts' @(21, 23, 69, 512, 513, 514, 5800, 5900) | ForEach-Object { [int]$_ })
$allowedCfg = Get-Setting $config 'PortCheck' 'Allowed' $null

function Get-PortName([int]$Number) {
    if ($portNames.ContainsKey($Number)) { return $portNames[$Number] }
    return ''
}

function Get-AllowedPorts([string]$Server) {
    <# Returns the allowed port list for a server, or $null when no policy is configured. #>
    if (-not $allowedCfg) { return $null }
    $short = $Server.Split('.')[0]
    foreach ($key in $allowedCfg.Keys) {
        if ($key -eq 'Default') { continue }
        if ($Server -like $key -or $short -like $key) { return ,@($allowedCfg[$key] | ForEach-Object { [int]$_ }) }
    }
    if ($allowedCfg.ContainsKey('Default')) { return ,@($allowedCfg['Default'] | ForEach-Object { [int]$_ }) }
    return ,@()
}

# --- Port test ------------------------------------------------------------

function Test-TcpPorts([string]$Target, [int[]]$Ports, [int]$Timeout) {
    <# Tries all ports on one server at the same time and waits up to $Timeout ms. #>
    $addresses = [System.Net.Dns]::GetHostAddresses($Target)
    $ip = $addresses | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1
    if (-not $ip) { $ip = $addresses | Select-Object -First 1 }

    $attempts = foreach ($p in $Ports) {
        $client = New-Object System.Net.Sockets.TcpClient($ip.AddressFamily)
        [pscustomobject]@{ Port = $p; Client = $client; Async = $client.BeginConnect($ip, $p, $null, $null); ElapsedMs = $null }
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $Timeout) {
        $pending = 0
        foreach ($a in $attempts) {
            if ($null -eq $a.ElapsedMs) {
                if ($a.Async.IsCompleted) { $a.ElapsedMs = [int]$sw.ElapsedMilliseconds } else { $pending++ }
            }
        }
        if ($pending -eq 0) { break }
        Start-Sleep -Milliseconds 20
    }

    foreach ($a in $attempts) {
        $state = 'Filtered'
        if ($a.Async.IsCompleted) {
            try {
                $a.Client.EndConnect($a.Async)
                $state = 'Open'
            }
            catch {
                $ex = $_.Exception
                while ($ex.InnerException) { $ex = $ex.InnerException }
                $code = if ($ex -is [System.Net.Sockets.SocketException]) { "$($ex.SocketErrorCode)" } else { $ex.GetType().Name }
                $state = switch ($code) {
                    'ConnectionRefused'  { 'Closed' }
                    'TimedOut'           { 'Filtered' }
                    'HostUnreachable'    { 'Unreachable' }
                    'NetworkUnreachable' { 'Unreachable' }
                    default              { "Error ($code)" }
                }
            }
        }
        $a.Client.Close()
        [pscustomobject]@{ Port = $a.Port; State = $state; ResponseMs = $a.ElapsedMs; IP = "$ip" }
    }
}

# --- Scan -----------------------------------------------------------------

$reportPath = New-ReportPath -Config $config -Prefix 'OpenPorts' -OutputFile $OutputFile
$details    = [System.Collections.Generic.List[object]]::new()
$matrix     = [System.Collections.Generic.List[object]]::new()
$startTime  = Get-Date
$portLabels = @{}
foreach ($p in $ports) { $name = Get-PortName $p; $portLabels[$p] = if ($name) { "$p $name" } else { "$p" } }

Write-Host ("Checking {0} port(s) on {1} server(s) from {2} (timeout {3} ms)" -f $ports.Count, $ComputerName.Count, [System.Environment]::MachineName, $TimeoutMs)
Write-Host ""

$n = 0
foreach ($server in $ComputerName) {
    $n++
    Write-Progress -Activity 'Checking ports' -Status "$server ($n of $($ComputerName.Count))" -PercentComplete (($n - 1) / $ComputerName.Count * 100)

    $allowed = Get-AllowedPorts $server
    $row = [ordered]@{ Server = $server; IP = ''; OpenPorts = '' }

    try {
        $results = @(Test-TcpPorts $server $ports $TimeoutMs)
    }
    catch {
        $err = Get-ErrorMessage $_
        $row.IP = 'Not resolved'
        $row.OpenPorts = "Error: $($err.Message)"
        foreach ($p in $ports) { $row[$portLabels[$p]] = '' }
        $matrix.Add([pscustomobject]$row)
        $details.Add([pscustomobject]@{ Server = $server; IP = ''; Port = $null; Service = ''; State = 'Error'; ResponseMs = $null
                                        Allowed = ''; Severity = 'High'; Finding = "Could not resolve or connect: $($err.Message)" })
        Write-Host ("{0,-25} could not be resolved: {1}" -f $server, $err.Message) -ForegroundColor Red
        continue
    }

    $row.IP = $results[0].IP
    foreach ($r in $results) {
        $isAllowed = if ($null -eq $allowed) { $null } else { $allowed -contains $r.Port }
        $finding = ''; $severity = ''
        if ($r.State -eq 'Open' -and $riskyPorts -contains $r.Port) {
            $finding = 'Risky service open'; $severity = 'High'
        }
        elseif ($r.State -eq 'Open' -and $isAllowed -eq $false) {
            $finding = 'Open but not allowed'; $severity = 'Medium'
        }
        elseif ($r.State -ne 'Open' -and $isAllowed -eq $true) {
            $finding = 'Allowed port not reachable'; $severity = 'Low'
        }

        $details.Add([pscustomobject]@{
            Server     = $server
            IP         = $r.IP
            Port       = $r.Port
            Service    = Get-PortName $r.Port
            State      = $r.State
            ResponseMs = if ($r.State -eq 'Open' -or $r.State -eq 'Closed') { $r.ResponseMs } else { $null }
            Allowed    = if ($null -eq $isAllowed) { '(no policy)' } elseif ($isAllowed) { 'Yes' } else { 'No' }
            Severity   = $severity
            Finding    = $finding
        })
        $row[$portLabels[$r.Port]] = $r.State
    }
    $open = @($results | Where-Object { $_.State -eq 'Open' } | ForEach-Object { $_.Port })
    $row.OpenPorts = $open -join ', '
    $matrix.Add([pscustomobject]$row)

    $serverFindings = @($details | Where-Object { $_.Server -eq $server -and $_.Finding })
    $colour = if ($serverFindings | Where-Object { $_.Severity -eq 'High' }) { 'Red' } elseif ($serverFindings.Count) { 'Yellow' } else { 'Green' }
    Write-Host ("{0,-25} open: {1}" -f $server, $(if ($open.Count) { $open -join ', ' } else { 'none' })) -ForegroundColor $colour
    foreach ($f in $serverFindings) { Write-Host ("    - {0}: {1} {2}" -f $f.Finding, $f.Port, $f.Service) -ForegroundColor $colour }
}
Write-Progress -Activity 'Checking ports' -Completed

# --- Export ---------------------------------------------------------------

$sevOrder = @{ High = 0; Medium = 1; Low = 2 }
$findings = @($details | Where-Object { $_.Finding } | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } }, Server, Port)
$counts   = [ordered]@{}
foreach ($s in 'High', 'Medium', 'Low') { $counts[$s] = @($findings | Where-Object { $_.Severity -eq $s }).Count }

$summary = [ordered]@{
    'Run at'              = $startTime.ToString('yyyy-MM-dd HH:mm:ss')
    'Scanned from'        = "$([System.Environment]::MachineName) - results show what is reachable from this machine only"
    'Servers'             = $ComputerName.Count
    'Ports checked'       = ($ports -join ', ')
    'Timeout (ms)'        = $TimeoutMs
    'Allowed-port policy' = if ($allowedCfg) { 'PortCheck.Allowed in config.psd1' } else { 'Not configured - only risky ports are flagged' }
    'Risky ports'         = ($riskyPorts -join ', ')
    'High findings'       = $counts.High
    'Medium findings'     = $counts.Medium
    'Low findings'        = $counts.Low
    'State meanings'      = 'Open = accepted; Closed = refused (nothing listening); Filtered = no reply (firewall)'
    'Run by'              = "$env:USERDOMAIN\$env:USERNAME"
}

$matrixHeaders = @('Server', 'IP', 'OpenPorts') + @($ports | ForEach-Object { $portLabels[$_] })
$matrixWidths  = @(25, 16, 30) + @($ports | ForEach-Object { [math]::Max(10, $portLabels[$_].Length + 2) })
$detailHeaders = @('Server', 'IP', 'Port', 'Service', 'State', 'ResponseMs', 'Allowed', 'Severity', 'Finding')
$detailWidths  = @(25, 16, 8, 16, 14, 12, 12, 10, 30)

$saved = Save-Workbook -FilePath $reportPath -Sheets @(
    @{ Name = 'Port Matrix'; Rows = $matrix;   Headers = $matrixHeaders; Widths = $matrixWidths }
    @{ Name = 'Findings';    Rows = $findings; Headers = $detailHeaders; Widths = $detailWidths }
    @{ Name = 'Details';     Rows = $details;  Headers = $detailHeaders; Widths = $detailWidths }
    @{ Name = 'Summary';     Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(22, 90) }
)
if (-not $saved) { exit 1 }

Write-Host ""
Write-Host ("Findings: {0} high, {1} medium, {2} low" -f $counts.High, $counts.Medium, $counts.Low) -ForegroundColor $(if ($counts.High) { 'Red' } elseif ($counts.Medium + $counts.Low) { 'Yellow' } else { 'Green' })
Write-Host "Report  : $reportPath"

if ($SendEmail) {
    if (-not $findings.Count -and -not $AlwaysEmail) {
        Write-Host 'No findings - no email sent (use -AlwaysEmail to send anyway).'
    }
    else {
        $intro = @("$($ComputerName.Count) server(s) checked from $([System.Environment]::MachineName): $($counts.High) high, $($counts.Medium) medium, $($counts.Low) low.")
        $html  = New-HtmlReport -Title 'Open port check' -IntroLines $intro -Sections @(
            @{ Heading = 'Findings'; Html = ConvertTo-HtmlTable -Rows $findings -Columns @('Severity', 'Server', 'Port', 'Service', 'State', 'Finding') -StatusColumn 'Severity' -EmptyText 'No findings.' }
            @{ Heading = 'Open ports by server'; Html = ConvertTo-HtmlTable -Rows $matrix -Columns @('Server', 'IP', 'OpenPorts') }
        )
        $state = if ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Medium' } elseif ($counts.Low) { 'Low' } else { 'OK' }
        if (Send-ToolkitMail -Config $config -Subject "[$state] Open port check - $($findings.Count) finding(s)" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
    }
}
