<#
.SYNOPSIS
    Checks which TCP ports are open on your servers.
.DESCRIPTION
    Flags risky services (Telnet, FTP, VNC...) and any open port not listed in -AllowedPort. Only scan systems you're responsible for.
.EXAMPLE
    .\Get-OpenPortReport.ps1 -ComputerName SERVER01, SERVER02 -AllowedPort 135, 445, 3389, 5985
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
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

    [string]$OutputFile,

    # Email settings (for -SendEmail / -NotifyUsers)
    [string]$SmtpServer,
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [pscredential]$SmtpCredential,
    [string]$MailFrom,
    [string[]]$MailTo,
    [string[]]$AllowedPort
)

#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)

$script:ToolkitRoot   = $PSScriptRoot

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)

$script:IsWindowsHost = ($PSVersionTable.PSEdition -eq 'Desktop') -or $IsWindows

$script:HtmlFont = 'font-family:Segoe UI,Arial,sans-serif;font-size:13px;'

function Get-ToolkitConfig {
    # Standalone: no config file. Settings come from parameters and built-in defaults.
    $cfg = @{ General = @{ OutputFolder = (Get-Location).Path } }
    if ($SmtpServer) {
        $cfg.Email = @{
            SmtpServer = $SmtpServer; Port = $SmtpPort; UseSsl = [bool]$SmtpUseSsl; From = $MailFrom
            To = @($MailTo | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
    }
    if ($AllowedPort) { $cfg.PortCheck = @{ Allowed = @{ Default = @($AllowedPort | ForEach-Object { "$_" -split "," } | Where-Object { $_.Trim() } | ForEach-Object { [int]$_.Trim() }) } } }
    return $cfg
}

function Get-Setting {
    param([hashtable]$Config, [string]$Section, [string]$Key, $Default = $null)
    $value = $null
    if ($Config -and $Config.ContainsKey($Section) -and $Config[$Section] -is [hashtable] -and $Config[$Section].ContainsKey($Key)) {
        $value = $Config[$Section][$Key]
    }
    if ($null -eq $value -or ($value -is [string] -and $value -eq '') -or ($value -is [array] -and $value.Count -eq 0)) { $value = $Default }
    return $value
}

function Split-ListValue([string[]]$Value) {
    <#
        Accepts both -Param A, B and -Param 'A,B'. The second form is what you get when a
        script is run with powershell.exe -File (for example from Task Scheduler).
    #>
    return @($Value | ForEach-Object { "$_" -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Resolve-ToolkitPath([string]$PathValue) {
    <# Turns a relative path from the config into a full path under the toolkit root. #>
    if (-not $PathValue) { return $PathValue }
    if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
    return [System.IO.Path]::GetFullPath((Join-Path $script:ToolkitRoot $PathValue))
}

function New-ReportPath {
    <#
        Returns the output file path. Uses -OutputFile if given (fixing the extension),
        otherwise <OutputFolder>\<Prefix>_<timestamp><Extension>. Creates the folder.
    #>
    param(
        [hashtable]$Config,
        [string]$Prefix,
        [string]$Extension = '.xlsx',
        [string]$OutputFile
    )
    if ($OutputFile) {
        if ([System.IO.Path]::GetExtension($OutputFile) -ne $Extension) {
            $OutputFile = [System.IO.Path]::ChangeExtension($OutputFile, $Extension)
        }
        $folder = Split-Path -Path $OutputFile -Parent
        $file   = $OutputFile
    }
    else {
        $folder = Resolve-ToolkitPath (Get-Setting $Config 'General' 'OutputFolder' 'Reports')
        $file   = Join-Path $folder ("{0}_{1}{2}" -f $Prefix, (Get-Date -Format 'yyyyMMdd_HHmmss'), $Extension)
    }
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force -WhatIf:$false -Confirm:$false | Out-Null
    }
    return $file
}

function Get-ErrorMessage($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    while ($ex.InnerException) { $ex = $ex.InnerException }
    return [pscustomobject]@{ Type = $ex.GetType().Name; Message = $ex.Message }
}

function ConvertTo-XmlText([string]$s) {
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[^\u0009\u000A\u000D\u0020-\uFFFD]', '')
    return [System.Security.SecurityElement]::Escape($s)
}

function Get-ColumnLetter([int]$Index) {
    # 0 -> A, 25 -> Z, 26 -> AA
    $name = ''
    $n = $Index + 1
    while ($n -gt 0) {
        $rem  = ($n - 1) % 26
        $name = [char](65 + $rem) + $name
        $n    = [math]::Floor(($n - 1) / 26)
    }
    return $name
}

function Add-ZipText($Zip, [string]$Name, [string]$Text) {
    $entry  = $Zip.CreateEntry($Name)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Add-XlsxSheet {
    param($Zip, [int]$Index, [string[]]$Headers, $Rows, [int[]]$Widths)

    $inv     = [System.Globalization.CultureInfo]::InvariantCulture
    $letters = for ($i = 0; $i -lt $Headers.Count; $i++) { Get-ColumnLetter $i }
    $entry   = $Zip.CreateEntry("xl/worksheets/sheet$Index.xml")
    $w       = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try {
        $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        $w.Write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        $w.Write('<cols>')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $width = if ($Widths -and $i -lt $Widths.Count) { $Widths[$i] } else { 20 }
            $w.Write(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $width))
        }
        $w.Write('</cols><sheetData>')

        $w.Write('<row r="1">')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $w.Write(('<c r="{0}1" s="1" t="inlineStr"><is><t>{1}</t></is></c>' -f $letters[$i], (ConvertTo-XmlText $Headers[$i])))
        }
        $w.Write('</row>')

        $r = 1
        foreach ($row in $Rows) {
            $r++
            $w.Write("<row r=`"$r`">")
            for ($i = 0; $i -lt $Headers.Count; $i++) {
                $ref = "$($letters[$i])$r"
                $val = $row.($Headers[$i])
                if ($null -eq $val) {
                    continue
                }
                elseif ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal] -or $val -is [single]) {
                    $w.Write("<c r=`"$ref`"><v>$(([double]$val).ToString('R', $inv))</v></c>")
                }
                elseif ($val -is [datetime]) {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t>$($val.ToString('yyyy-MM-dd HH:mm', $inv))</t></is></c>")
                }
                elseif ($val -is [bool]) {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t>$(if ($val) { 'Yes' } else { 'No' })</t></is></c>")
                }
                else {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText "$val")</t></is></c>")
                }
            }
            $w.Write('</row>')
        }
        $w.Write('</sheetData>')
        $w.Write("<autoFilter ref=`"A1:$($letters[-1])$r`"/>")
        $w.Write('</worksheet>')
    }
    finally { $w.Dispose() }
}

function Export-Xlsx {
    <#
        Writes a workbook. -Sheets is an array of hashtables:
            @{ Name = 'Sheet name'; Headers = @('Col1','Col2'); Rows = $objects; Widths = @(40, 12) }
        Row values are read by header name. Numbers stay numeric; dates are written as yyyy-MM-dd HH:mm.
    #>
    param([string]$FilePath, [object[]]$Sheets)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -LiteralPath $FilePath) { Remove-Item -LiteralPath $FilePath -Force -ErrorAction Stop -WhatIf:$false -Confirm:$false }
    $zip = [System.IO.Compression.ZipFile]::Open($FilePath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $overrides = ''; $entries = ''; $rels = ''
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $n    = $i + 1
            $name = $Sheets[$i].Name -replace '[\[\]\*\?/\\:]', '-'
            if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }
            $overrides += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
            $entries   += "<sheet name=`"$(ConvertTo-XmlText $name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
            $rels      += "<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
        }
        $stylesId = $Sheets.Count + 1

        Add-ZipText $zip '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
            $overrides + '</Types>')

        Add-ZipText $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>')

        Add-ZipText $zip 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheets>' + $entries + '</sheets></workbook>')

        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $rels +
            "<Relationship Id=`"rId$stylesId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>" +
            '</Relationships>')

        Add-ZipText $zip 'xl/styles.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
            '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor indexed="64"/></patternFill></fill></fills>' +
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>' +
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
            '</styleSheet>')

        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $s = $Sheets[$i]
            Add-XlsxSheet -Zip $zip -Index ($i + 1) -Headers $s.Headers -Rows $s.Rows -Widths $s.Widths
        }
    }
    finally { $zip.Dispose() }
}

function Save-Workbook {
    <# Export-Xlsx with a friendly error if the file is open in Excel. #>
    param([string]$FilePath, [object[]]$Sheets)
    try {
        Export-Xlsx -FilePath $FilePath -Sheets $Sheets
        return $true
    }
    catch {
        Write-Error ("Could not save {0}. If it is open in Excel, close it and run again. Details: {1}" -f $FilePath, $_.Exception.Message)
        return $false
    }
}

function New-SummaryRows([System.Collections.IDictionary]$Items) {
    <# Turns an ordered hashtable into Item/Value rows for a Summary sheet. #>
    foreach ($k in $Items.Keys) { [pscustomobject]@{ Item = $k; Value = $Items[$k] } }
}

function Get-StatusStyle([string]$Status) {
    switch -Regex ($Status) {
        '^(Critical|Expired|Error|Unreachable|Failed|High)' { return 'background-color:#f8d7da;color:#842029;font-weight:600;' }
        '^(Warning|Expiring|Medium)'                   { return 'background-color:#fff3cd;color:#664d03;font-weight:600;' }
        '^Low'                                         { return 'background-color:#e7f1ff;color:#084298;font-weight:600;' }
        '^OK'                                          { return 'background-color:#d1e7dd;color:#0f5132;font-weight:600;' }
        default                                        { return '' }
    }
}

function ConvertTo-HtmlTable {
    <#
        Builds a table with inline styles (works in Outlook).
        -StatusColumn colours that cell by its value (OK / Warning / Critical ...).
    #>
    param(
        $Rows,
        [string[]]$Columns,
        [string]$StatusColumn,
        [string]$EmptyText = 'None.'
    )
    if ($null -eq $Rows -or @($Rows).Count -eq 0) {
        return "<p style=`"$($script:HtmlFont)color:#555;`">$([System.Net.WebUtility]::HtmlEncode($EmptyText))</p>"
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table cellpadding=`"6`" cellspacing=`"0`" style=`"border-collapse:collapse;$($script:HtmlFont)border:1px solid #d0d7de;margin-bottom:16px;`">")
    [void]$sb.Append('<tr>')
    foreach ($c in $Columns) {
        [void]$sb.Append("<th style=`"background-color:#305496;color:#ffffff;text-align:left;border:1px solid #d0d7de;`">$([System.Net.WebUtility]::HtmlEncode($c))</th>")
    }
    [void]$sb.Append('</tr>')
    $odd = $false
    foreach ($row in $Rows) {
        $odd = -not $odd
        $bg  = if ($odd) { '#ffffff' } else { '#f6f8fa' }
        [void]$sb.Append("<tr style=`"background-color:$bg;`">")
        foreach ($c in $Columns) {
            $val = $row.$c
            if ($val -is [datetime]) { $val = $val.ToString('yyyy-MM-dd HH:mm') }
            elseif ($val -is [bool]) { $val = if ($val) { 'Yes' } else { 'No' } }
            $style = 'border:1px solid #d0d7de;vertical-align:top;'
            if ($c -eq $StatusColumn) { $style += Get-StatusStyle "$val" }
            [void]$sb.Append("<td style=`"$style`">$([System.Net.WebUtility]::HtmlEncode("$val"))</td>")
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

function New-HtmlReport {
    <#
        Wraps sections into a full HTML page.
        -Sections: array of @{ Heading = '...'; Html = '...' }
    #>
    param(
        [string]$Title,
        [string[]]$IntroLines,
        [object[]]$Sections
    )
    $enc = { param($t) [System.Net.WebUtility]::HtmlEncode($t) }
    $sb  = New-Object System.Text.StringBuilder
    [void]$sb.Append("<!DOCTYPE html><html><head><meta charset=`"utf-8`"><title>$(& $enc $Title)</title></head>")
    [void]$sb.Append("<body style=`"$($script:HtmlFont)color:#1f2328;background-color:#ffffff;margin:16px;`">")
    [void]$sb.Append("<h2 style=`"font-family:Segoe UI,Arial,sans-serif;color:#305496;margin:0 0 8px 0;`">$(& $enc $Title)</h2>")
    foreach ($line in $IntroLines) {
        [void]$sb.Append("<p style=`"$($script:HtmlFont)margin:2px 0;`">$(& $enc $line)</p>")
    }
    foreach ($s in $Sections) {
        [void]$sb.Append("<h3 style=`"font-family:Segoe UI,Arial,sans-serif;color:#1f2328;margin:20px 0 8px 0;`">$(& $enc $s.Heading)</h3>")
        [void]$sb.Append($s.Html)
    }
    [void]$sb.Append("<p style=`"font-family:Segoe UI,Arial,sans-serif;font-size:11px;color:#6e7781;margin-top:24px;`">Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') on $([System.Environment]::MachineName) by $(& $enc (Split-Path -Leaf "$($MyInvocation.PSCommandPath)")).</p>")
    [void]$sb.Append('</body></html>')
    return $sb.ToString()
}

function Send-ToolkitMail {
    <# Sends an HTML email using -SmtpServer, -SmtpPort, -SmtpUseSsl, -SmtpCredential, -MailFrom and -MailTo. #>
    param(
        [hashtable]$Config,
        [string]$Subject,
        [string]$HtmlBody,
        [string[]]$To,
        [string[]]$Attachments
    )
    $smtp = Get-Setting $Config 'Email' 'SmtpServer'
    $from = Get-Setting $Config 'Email' 'From'
    if (-not $To) { $To = Get-Setting $Config 'Email' 'To' }

    if (-not $smtp -or -not $from -or -not $To) {
        Write-Warning 'Email not sent: use -SmtpServer, -MailFrom and -MailTo.'
        return $false
    }

    $params = @{
        SmtpServer = $smtp
        Port       = [int](Get-Setting $Config 'Email' 'Port' 25)
        From       = $from
        To         = $To
        Subject    = $Subject
        Body       = $HtmlBody
        BodyAsHtml = $true
        Encoding   = [System.Text.Encoding]::UTF8
    }
    if (Get-Setting $Config 'Email' 'UseSsl' $false) { $params.UseSsl = $true }

    $credFile = Get-Setting $Config 'Email' 'CredentialFile'
    if ($credFile) {
        $credPath = Resolve-ToolkitPath $credFile
        if (-not (Test-Path -LiteralPath $credPath)) {
            Write-Warning "Email not sent: credential file not found at $credPath."
            return $false
        }
        $params.Credential = Import-Clixml -LiteralPath $credPath
    }
    if ($SmtpCredential) { $params.Credential = $SmtpCredential }
    if ($Attachments) { $params.Attachments = $Attachments }

    try {
        # Send-MailMessage is marked obsolete in PowerShell 7 but still works with most internal relays.
        Send-MailMessage @params -ErrorAction Stop -WarningAction SilentlyContinue
        return $true
    }
    catch {
        Write-Warning "Email could not be sent: $($_.Exception.Message)"
        return $false
    }
}

#endregion
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
if (-not $ports.Count)        { throw 'No ports to check. Use -Port.' }
if (-not $ComputerName.Count) { throw 'No servers to check. Use -ComputerName.' }

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
    'Allowed-port policy' = if ($allowedCfg) { ('-AllowedPort ' + ($AllowedPort -join ', ')) } else { 'Not configured - only risky ports are flagged' }
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
