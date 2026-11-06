<#
.SYNOPSIS
    Finds Active Directory computers with no BitLocker recovery key backed up to AD.
.DESCRIPTION
    Matches every computer account with its msFVE-RecoveryInformation objects (the keys themselves are
    never read) and reports coverage. Read-only.
.EXAMPLE
    .\Get-BitLockerKeyAudit.ps1
.EXAMPLE
    .\Get-BitLockerKeyAudit.ps1 -SearchBase 'OU=Workstations,DC=corp,DC=example,DC=com' -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: ActiveDirectory module (RSAT); an account that can see BitLocker recovery objects (e.g. Domain Admins)
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$Server,
    [switch]$IncludeDisabled,
    [switch]$IncludeServers,
    [ValidateRange(1, 3650)] [int]$StaleDays,

    [switch]$SendEmail,
    [string]$OutputFile,

    # Email settings (for -SendEmail / -NotifyUsers)
    [string]$SmtpServer,
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [pscredential]$SmtpCredential,
    [string]$MailFrom,
    [string[]]$MailTo
)

#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)

$script:ToolkitRoot   = $PSScriptRoot

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)

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
        $name = [string][char](65 + [int]$rem) + $name
        $n    = [int][math]::Floor(($n - 1) / 26)
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
        '^(Critical|Expired|Error|Unreachable|Failed|Not done|High)' { return 'background-color:#f8d7da;color:#842029;font-weight:600;' }
        '^(Warning|Expiring|Medium)'                   { return 'background-color:#fff3cd;color:#664d03;font-weight:600;' }
        '^Low'                                         { return 'background-color:#e7f1ff;color:#084298;font-weight:600;' }
        '^(OK|Done|Requested)'                         { return 'background-color:#d1e7dd;color:#0f5132;font-weight:600;' }
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

if (-not $PSBoundParameters.ContainsKey('SearchBase')) { $SearchBase = Get-Setting $config 'BitLocker' 'SearchBase' '' }
if (-not $PSBoundParameters.ContainsKey('Server'))     { $Server     = Get-Setting $config 'ActiveDirectory' 'Server' '' }
if (-not $PSBoundParameters.ContainsKey('StaleDays'))  { $StaleDays  = [int](Get-Setting $config 'BitLocker' 'StaleDays' 90) }

if ($MyInvocation.InvocationName -ne '.') {
    try { Import-Module ActiveDirectory -ErrorAction Stop }
    catch { throw 'The ActiveDirectory module is not installed. Install RSAT: Active Directory tools.' }

    $ad = @{}
    if ($Server) { $ad.Server = $Server }
    if (-not $SearchBase) { $SearchBase = (Get-ADDomain @ad).DistinguishedName }
    $now        = Get-Date
    $reportPath = New-ReportPath -Config $config -Prefix 'BitLockerKeys' -OutputFile $OutputFile

    Write-Host "Search base: $SearchBase"
    Write-Host 'Getting computers...'
    $filter = if ($IncludeDisabled) { '*' } else { 'Enabled -eq $true' }
    $computers = @(Get-ADComputer -Filter $filter -SearchBase $SearchBase -Properties OperatingSystem, LastLogonDate, whenCreated, Enabled @ad |
                   Where-Object { "$($_.OperatingSystem)" -like 'Windows*' -and ($IncludeServers -or "$($_.OperatingSystem)" -notlike '*Server*') })

    # One query for every recovery object; they sit under the computer they belong to
    Write-Host 'Getting BitLocker recovery objects...'
    $keysByComputer = @{}
    foreach ($k in @(Get-ADObject -LDAPFilter '(objectClass=msFVE-RecoveryInformation)' -SearchBase $SearchBase -Properties whenCreated @ad)) {
        $parent = $k.DistinguishedName -replace '^.+?(?<!\\),', ''
        if (-not $keysByComputer.ContainsKey($parent)) { $keysByComputer[$parent] = [System.Collections.Generic.List[datetime]]::new() }
        $keysByComputer[$parent].Add($k.whenCreated)
    }

    $rows = foreach ($c in $computers) {
        $keys     = $keysByComputer[$c.DistinguishedName]
        $count    = if ($keys) { $keys.Count } else { 0 }
        $latest   = if ($count) { ($keys | Sort-Object -Descending | Select-Object -First 1) } else { $null }
        $lastSeen = if ($c.LastLogonDate) { [int]($now - $c.LastLogonDate).TotalDays } else { $null }
        $stale    = ($null -eq $lastSeen -and ($now - $c.whenCreated).TotalDays -gt $StaleDays) -or ($lastSeen -gt $StaleDays)

        $status = if ($count) { 'Key backed up' } elseif ($stale) { 'No key (stale computer)' } else { 'No key' }
        [pscustomobject]@{
            Status          = $status
            Computer        = $c.Name
            OperatingSystem = $c.OperatingSystem
            Enabled         = [bool]$c.Enabled
            LastLogon       = $c.LastLogonDate
            DaysSinceLogon  = $lastSeen
            Keys            = $count
            LatestKey       = $latest
            OU              = ($c.DistinguishedName -replace '^.+?(?<!\\),', '')
        }
    }
    $rows = @($rows)

    $order    = @{ 'No key' = 0; 'No key (stale computer)' = 1; 'Key backed up' = 2 }
    $sorted   = @($rows | Sort-Object @{ Expression = { $order[$_.Status] } }, Computer)
    $missing  = @($rows | Where-Object { $_.Status -eq 'No key' })
    $active   = @($rows | Where-Object { $_.Status -ne 'No key (stale computer)' })
    $coverage = if ($active.Count) { [math]::Round(@($active | Where-Object { $_.Keys -gt 0 }).Count / $active.Count * 100, 1) } else { 0 }

    $summary = [ordered]@{
        'Run at'                       = $now.ToString('yyyy-MM-dd HH:mm:ss')
        'Search base'                  = $SearchBase
        'Computers checked'            = "$($rows.Count) Windows $(if ($IncludeServers) { 'computers' } else { 'clients (servers excluded - use -IncludeServers)' })"
        'Key backed up'                = @($rows | Where-Object Keys -gt 0).Count
        'No key (active)'              = $missing.Count
        "No key (no logon for $StaleDays+ days)" = @($rows | Where-Object Status -eq 'No key (stale computer)').Count
        'Coverage of active computers' = "$coverage%"
        'Note'                         = 'Devices joined only to Entra ID (not AD) back up their keys to Entra ID and are not covered here. A computer may be encrypted without a backed-up key, or have a key from an old installation.'
    }

    $saved = Save-Workbook -FilePath $reportPath -Sheets @(
        @{ Name = 'Computers'; Rows = $sorted; Headers = @('Status', 'Computer', 'OperatingSystem', 'Enabled', 'LastLogon', 'DaysSinceLogon', 'Keys', 'LatestKey', 'OU'); Widths = @(24, 22, 34, 9, 18, 15, 7, 18, 60) }
        @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(36, 100) }
    )
    if (-not $saved) { exit 1 }

    Write-Host ""
    Write-Host ("Coverage: {0}% of active computers have a key backed up. {1} active computer(s) have none." -f $coverage, $missing.Count) -ForegroundColor $(if ($missing.Count) { 'Yellow' } else { 'Green' })
    $missing | Select-Object -First 15 | ForEach-Object { Write-Host ("  {0,-20} {1}" -f $_.Computer, $_.OperatingSystem) }
    Write-Host "Report  : $reportPath"

    if ($SendEmail) {
        $html = New-HtmlReport -Title 'BitLocker recovery key audit' -IntroLines @("$coverage% of active Windows computers have a recovery key backed up to AD. $($missing.Count) have none.") -Sections @(
            @{ Heading = "Active computers with no key ($($missing.Count))"; Html = ConvertTo-HtmlTable -Rows $missing -Columns @('Computer', 'OperatingSystem', 'LastLogon', 'OU') -EmptyText 'None - every active computer has a key backed up.' }
        )
        if (Send-ToolkitMail -Config $config -Subject "BitLocker keys: $coverage% coverage, $($missing.Count) computer(s) without a key" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
    }
}
