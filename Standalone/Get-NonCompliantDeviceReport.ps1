<#
.SYNOPSIS
    Exports non-compliant devices from Intune and Microsoft Defender for Endpoint, with the reasons why.
.DESCRIPTION
    Lists Intune devices that are not compliant, with the policies and settings they fail, then adds what
    Defender knows about each one: onboarding, sensor health, exposure level, active alerts and antivirus
    state. Every device is also listed on an All Devices sheet. If Defender data is unavailable, those columns
    are left blank. Read-only.
.EXAMPLE
    .\Get-NonCompliantDeviceReport.ps1
.EXAMPLE
    .\Get-NonCompliantDeviceReport.ps1 -Platform Windows -SendEmail
.EXAMPLE
    .\Get-NonCompliantDeviceReport.ps1 -SkipDefender      # Intune only, e.g. without Defender for Endpoint
.EXAMPLE
    .\Get-NonCompliantDeviceReport.ps1 -AntivirusForAllDevices      # antivirus state for every Windows device, not just flagged ones
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: Microsoft.Graph.Authentication module, Global Reader (or Intune Administrator) plus Security Reader
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [string[]]$Platform,   # Windows, macOS, iOS, Android, Linux
    [switch]$IncludeUnknown,
    [switch]$SkipDefender,
    [switch]$SkipSettingDetail,
    [switch]$AntivirusForAllDevices,
    [ValidateRange(1, 3650)] [int]$StaleDays,

    [switch]$SendEmail,
    [string]$OutputFile,

    # Email settings (for -SendEmail / -NotifyUsers)
    [string]$SmtpServer,
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [pscredential]$SmtpCredential,
    [string]$MailFrom,
    [string[]]$MailTo,

    # Report labels (HTML reports and emails)
    [string]$Organisation,
    [string]$Classification = 'Internal'
)

#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)

$script:ToolkitRoot   = $PSScriptRoot

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)

$script:ReportStyle = @{
    Sans   = "font-family:'IBM Plex Sans','Segoe UI',Helvetica,Arial,sans-serif;"
    Mono   = "font-family:'IBM Plex Mono',Consolas,'Courier New',monospace;"
    Page   = '#f4f2ee'; Card = '#ffffff'; Ink = '#1c1b19'; Text = '#45423d'; Muted = '#78736a'
    Line   = '#e6e2da'; Rule = '#d6d1c8'; Head = '#faf9f7'; OnInk = '#e8e4dc'
}

function Get-ToolkitConfig {
    # Standalone: no config file. Settings come from parameters and built-in defaults.
    $cfg = @{ General = @{ OutputFolder = (Get-Location).Path } }
    if ($SmtpServer) {
        $cfg.Email = @{
            SmtpServer = $SmtpServer; Port = $SmtpPort; UseSsl = [bool]$SmtpUseSsl; From = $MailFrom
            To = @($MailTo | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
    }
    $cfg.General.Organisation = $Organisation; $cfg.General.Classification = $Classification
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

function Get-XlsxStatusStyle([string]$Value) {
    <# Cell style for a status value, matching the HTML badges: 2 red, 3 amber, 4 blue, 5 green, 0 plain. #>
    switch -Regex ($Value) {
        '^(Critical|Expired|Error|Unreachable|Failed|Not done|High)' { return 2 }
        '^(Warning|Expiring|Medium)'                                  { return 3 }
        '^Low'                                                        { return 4 }
        '^(OK|Done|Requested)'                                        { return 5 }
        default                                                       { return 0 }
    }
}

function Add-XlsxSheet {
    param($Zip, [int]$Index, [string[]]$Headers, $Rows, [int[]]$Widths)

    $inv     = [System.Globalization.CultureInfo]::InvariantCulture
    $letters = for ($i = 0; $i -lt $Headers.Count; $i++) { Get-ColumnLetter $i }
    # Columns whose values get status colours, and the label column of an Item/Value summary sheet
    $statusCols = @(for ($i = 0; $i -lt $Headers.Count; $i++) { if ($Headers[$i] -match '^(Status|Risk|Severity|Rating|Result|Overall)$') { $i } })
    $labelCol   = if ($Headers.Count -eq 2 -and $Headers[0] -eq 'Item' -and $Headers[1] -eq 'Value') { 0 } else { -1 }
    $entry   = $Zip.CreateEntry("xl/worksheets/sheet$Index.xml")
    $w       = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try {
        $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        $w.Write('<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        $w.Write('<sheetFormatPr defaultRowHeight="18"/>')
        $w.Write('<cols>')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $width = if ($Widths -and $i -lt $Widths.Count) { $Widths[$i] } else { 20 }
            $w.Write(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $width))
        }
        $w.Write('</cols><sheetData>')

        $w.Write('<row r="1" ht="24" customHeight="1">')
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
                $st  = if ($i -eq $labelCol) { 6 } elseif ($statusCols -contains $i -and $null -ne $val) { Get-XlsxStatusStyle "$val" } else { 0 }
                if ($null -eq $val) {
                    $w.Write("<c r=`"$ref`" s=`"0`"/>")   # empty but styled, so the row line is unbroken
                }
                elseif ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal] -or $val -is [single]) {
                    $w.Write("<c r=`"$ref`" s=`"$st`"><v>$(([double]$val).ToString('R', $inv))</v></c>")
                }
                elseif ($val -is [datetime]) {
                    $w.Write("<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t>$($val.ToString('yyyy-MM-dd HH:mm', $inv))</t></is></c>")
                }
                elseif ($val -is [bool]) {
                    $w.Write("<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t>$(if ($val) { 'Yes' } else { 'No' })</t></is></c>")
                }
                else {
                    $w.Write("<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText "$val")</t></is></c>")
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

        # Styles match the HTML reports: warm ink and greys, thin row lines, status colours like the badges.
        #   0 body  1 header  2 red  3 amber  4 blue  5 green  6 summary label
        $font = { param($Size, $Colour, $Name, [switch]$Bold) "<font>$(if ($Bold) { '<b/>' })<sz val=`"$Size`"/><color rgb=`"FF$Colour`"/><name val=`"$Name`"/></font>" }
        $fill = { param($Colour) "<fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF$Colour`"/><bgColor indexed=`"64`"/></patternFill></fill>" }
        $xf   = { param($Font, $Fill, $Border) "<xf numFmtId=`"0`" fontId=`"$Font`" fillId=`"$Fill`" borderId=`"$Border`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"left`" vertical=`"center`" indent=`"1`"/></xf>" }
        Add-ZipText $zip 'xl/styles.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<fonts count="7">' +
                (& $font 10 '1C1B19' 'Segoe UI') + (& $font 9 'E8E4DC' 'Consolas' -Bold) +
                (& $font 10 '9B2C2C' 'Segoe UI' -Bold) + (& $font 10 '875A0B' 'Segoe UI' -Bold) +
                (& $font 10 '2D5A86' 'Segoe UI' -Bold) + (& $font 10 '2F6A47' 'Segoe UI' -Bold) +
                (& $font 9 '78736A' 'Consolas') + '</fonts>' +
            '<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
                (& $fill '1C1B19') + (& $fill 'F9EBE9') + (& $fill 'FBF3E1') + (& $fill 'EAF1F8') + (& $fill 'E9F3EC') + '</fills>' +
            '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
                '<border><left/><right/><top/><bottom style="thin"><color rgb="FFE6E2DA"/></bottom><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="7">' +
                (& $xf 0 0 1) + (& $xf 1 2 0) + (& $xf 2 3 1) + (& $xf 3 4 1) + (& $xf 4 5 1) + (& $xf 5 6 1) + (& $xf 6 0 1) + '</cellXfs>' +
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
    <# Text, background and border colours for a status badge. #>
    switch -Regex ($Status) {
        '^(Critical|Expired|Error|Unreachable|Failed|Not done|High)' { return @{ Fg = '#9b2c2c'; Bg = '#f9ebe9'; Bd = '#ecc9c4' } }
        '^(Warning|Expiring|Medium)'                                  { return @{ Fg = '#875a0b'; Bg = '#fbf3e1'; Bd = '#ecd9ab' } }
        '^Low'                                                        { return @{ Fg = '#2d5a86'; Bg = '#eaf1f8'; Bd = '#c7d7e8' } }
        '^(OK|Done|Requested)'                                        { return @{ Fg = '#2f6a47'; Bg = '#e9f3ec'; Bd = '#c4ddcc' } }
        default                                                       { return $null }
    }
}

function ConvertTo-HtmlStatusBadge([string]$Status) {
    $enc = [System.Net.WebUtility]::HtmlEncode($Status)
    $c = Get-StatusStyle $Status
    if (-not $c) { return $enc }
    return "<span style=`"$($script:ReportStyle.Mono)font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;white-space:nowrap;color:$($c.Fg);background-color:$($c.Bg);border:1px solid $($c.Bd);border-radius:999px;padding:2px 9px;`">$enc</span>"
}

function ConvertTo-HtmlTable {
    <#
        Builds a table with inline styles (works in Outlook).
        -StatusColumn shows that column's value (OK / Warning / Critical ...) as a coloured badge.
    #>
    param(
        $Rows,
        [string[]]$Columns,
        [string]$StatusColumn,
        [string]$EmptyText = 'None.'
    )
    $h = $script:ReportStyle
    if ($null -eq $Rows -or @($Rows).Count -eq 0) {
        return "<p style=`"$($h.Sans)font-size:13px;color:$($h.Muted);margin:0 0 8px 0;`">$([System.Net.WebUtility]::HtmlEncode($EmptyText))</p>"
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"border-collapse:collapse;width:100%;border:1px solid $($h.Line);margin:0 0 8px 0;`">")
    [void]$sb.Append('<tr>')
    foreach ($c in $Columns) {
        $label = [regex]::Replace($c, '(?<=[a-z])(?=[A-Z])', ' ')   # SizeGB -> Size GB, DisplayName -> Display Name
        [void]$sb.Append("<th align=`"left`" style=`"$($h.Mono)font-size:10.5px;font-weight:500;letter-spacing:0.1em;text-transform:uppercase;color:$($h.OnInk);background-color:$($h.Ink);text-align:left;padding:9px 12px;white-space:nowrap;`">$([System.Net.WebUtility]::HtmlEncode($label))</th>")
    }
    [void]$sb.Append('</tr>')
    $rowList = @($Rows); $last = $rowList.Count - 1
    for ($i = 0; $i -le $last; $i++) {
        $row = $rowList[$i]
        $border = if ($i -lt $last) { "border-bottom:1px solid $($h.Line);" } else { '' }
        [void]$sb.Append('<tr>')
        foreach ($c in $Columns) {
            $val = $row.$c
            if ($val -is [datetime]) { $val = $val.ToString('yyyy-MM-dd HH:mm') }
            elseif ($val -is [bool]) { $val = if ($val) { 'Yes' } else { 'No' } }
            $cell = if ($c -eq $StatusColumn) { ConvertTo-HtmlStatusBadge "$val" } else { [System.Net.WebUtility]::HtmlEncode("$val") }
            [void]$sb.Append("<td style=`"$($h.Sans)font-size:13px;line-height:1.45;color:$($h.Ink);padding:9px 12px;vertical-align:top;$border`">$cell</td>")
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

function New-HtmlReport {
    <#
        Wraps sections into a full HTML page (also used as the email body).
        -Sections: array of @{ Heading = '...'; Html = '...' }
        -Stats:    optional summary tiles, array of @{ Label = '...'; Value = '...'; Status = 'Critical|Warning|OK|...' }
        -Meta:     optional extra rows for the details block, e.g. [ordered]@{ Tenant = 'Contoso' }
        Organisation and classification come from General.Organisation and General.Classification in the config.
    #>
    param(
        [string]$Title,
        [string[]]$IntroLines,
        [object[]]$Sections,
        [object[]]$Stats,
        [System.Collections.IDictionary]$Meta
    )
    $h   = $script:ReportStyle
    $enc = { param($t) [System.Net.WebUtility]::HtmlEncode("$t") }
    $now = Get-Date
    $scriptName = Split-Path -Leaf "$($MyInvocation.PSCommandPath)"
    $cfg = Get-Variable -Name config -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    # Read these directly so an empty Classification hides the label instead of falling back to the default
    $general        = if ($cfg -is [hashtable] -and $cfg['General'] -is [hashtable]) { $cfg['General'] } else { @{} }
    $classification = if ($general.ContainsKey('Classification')) { "$($general['Classification'])".Trim() } else { 'Internal' }
    $organisation   = "$($general['Organisation'])".Trim()
    $runBy = if ($env:USERNAME) { if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { $env:USERNAME } } elseif ($env:USER) { $env:USER } else { '' }

    $details = [ordered]@{}
    if ($Meta) { foreach ($k in $Meta.Keys) { $details[$k] = $Meta[$k] } }
    $details['Generated'] = $now.ToString('dd MMM yyyy HH:mm')
    if ($runBy) { $details['Run by'] = $runBy }
    $details['Computer'] = [System.Environment]::MachineName
    $details['Script']   = $scriptName

    $sb  = New-Object System.Text.StringBuilder
    [void]$sb.Append("<!DOCTYPE html><html lang=`"en`"><head><meta charset=`"utf-8`"><meta name=`"viewport`" content=`"width=device-width,initial-scale=1`"><title>$(& $enc $Title)</title></head>")
    [void]$sb.Append("<body style=`"margin:0;padding:0;background-color:$($h.Page);`">")
    [void]$sb.Append("<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"background-color:$($h.Page);`"><tr><td align=`"center`" style=`"padding:28px 16px;`">")
    [void]$sb.Append("<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"max-width:1040px;width:100%;background-color:$($h.Card);border:1px solid $($h.Line);`">")

    # Top bar: organisation / report type on the left, classification on the right
    $left = if ($organisation) { "$(& $enc $organisation) &nbsp;/&nbsp; Report" } else { 'Report' }
    [void]$sb.Append("<tr><td style=`"background-color:$($h.Ink);padding:0;`"><table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`"><tr>")
    [void]$sb.Append("<td style=`"$($h.Mono)font-size:10.5px;letter-spacing:0.14em;text-transform:uppercase;color:$($h.OnInk);padding:10px 32px;`">$left</td>")
    if ($classification) {
        [void]$sb.Append("<td align=`"right`" style=`"padding:6px 32px 6px 0;`"><span style=`"$($h.Mono)font-size:10.5px;font-weight:600;letter-spacing:0.14em;text-transform:uppercase;color:$($h.Ink);background-color:$($h.OnInk);padding:4px 10px;`">$(& $enc $classification)</span></td>")
    }
    [void]$sb.Append('</tr></table></td></tr>')

    # Title and intro
    [void]$sb.Append("<tr><td style=`"padding:26px 32px 18px 32px;`">")
    [void]$sb.Append("<div style=`"$($h.Sans)font-size:24px;line-height:1.25;font-weight:600;color:$($h.Ink);margin:0 0 10px 0;`">$(& $enc $Title)</div>")
    foreach ($line in $IntroLines) {
        [void]$sb.Append("<div style=`"$($h.Sans)font-size:13.5px;line-height:1.6;color:$($h.Text);margin:0;`">$(& $enc $line)</div>")
    }
    [void]$sb.Append('</td></tr>')

    # Details block (document control)
    [void]$sb.Append("<tr><td style=`"padding:0 32px 4px 32px;`"><table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"border-collapse:collapse;border-top:1px solid $($h.Rule);border-bottom:1px solid $($h.Rule);`"><tr>")
    $keys = @($details.Keys); $w = [math]::Floor(100 / $keys.Count)
    for ($i = 0; $i -lt $keys.Count; $i++) {
        $sep = if ($i -gt 0) { "border-left:1px solid $($h.Line);" } else { '' }
        [void]$sb.Append("<td width=`"$w%`" valign=`"top`" style=`"padding:10px 12px 10px $(if ($i) { 12 } else { 0 })px;$sep`">")
        [void]$sb.Append("<div style=`"$($h.Mono)font-size:10px;letter-spacing:0.12em;text-transform:uppercase;color:$($h.Muted);margin:0 0 3px 0;`">$(& $enc $keys[$i])</div>")
        [void]$sb.Append("<div style=`"$($h.Sans)font-size:13px;color:$($h.Ink);word-break:break-word;`">$(& $enc $details[$keys[$i]])</div></td>")
    }
    [void]$sb.Append('</tr></table></td></tr>')

    # Summary tiles
    if ($Stats) {
        [void]$sb.Append("<tr><td style=`"padding:20px 32px 4px 32px;`"><table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`"><tr>")
        $n = @($Stats).Count; $w = [math]::Floor(100 / $n)
        for ($i = 0; $i -lt $n; $i++) {
            $t = @($Stats)[$i]
            $c = if ($t.Status) { Get-StatusStyle "$($t.Status)" } else { $null }
            $valColour = if ($c -and "$($t.Value)" -ne '0') { $c.Fg } else { $h.Ink }
            $bar = if ($c) { $c.Fg } else { $h.Ink }
            $gap = if ($i -lt $n - 1) { 'padding-right:10px;' } else { '' }
            [void]$sb.Append("<td width=`"$w%`" valign=`"top`" style=`"$gap`"><div style=`"border:1px solid $($h.Line);border-top:3px solid $bar;padding:12px 14px;`">")
            [void]$sb.Append("<div style=`"$($h.Mono)font-size:10.5px;letter-spacing:0.12em;text-transform:uppercase;color:$($h.Muted);margin:0 0 4px 0;`">$(& $enc $t.Label)</div>")
            [void]$sb.Append("<div style=`"$($h.Sans)font-size:22px;font-weight:600;line-height:1.2;color:$valColour;`">$(& $enc "$($t.Value)")</div></div></td>")
        }
        [void]$sb.Append('</tr></table></td></tr>')
    }

    # Numbered sections
    $num = 0
    foreach ($s in $Sections) {
        $num++
        [void]$sb.Append("<tr><td style=`"padding:24px 32px 6px 32px;`">")
        [void]$sb.Append("<div style=`"border-bottom:1px solid $($h.Ink);padding:0 0 7px 0;margin:0 0 12px 0;`"><span style=`"$($h.Mono)font-size:12px;color:$($h.Muted);letter-spacing:0.08em;`">$($num.ToString('00'))</span>&nbsp;&nbsp;<span style=`"$($h.Sans)font-size:15px;font-weight:600;color:$($h.Ink);`">$(& $enc $s.Heading)</span></div>")
        [void]$sb.Append($s.Html)
        [void]$sb.Append('</td></tr>')
    }

    # Footer: organisation and classification / title / generated
    $footLeft = (@($organisation, $classification) | Where-Object { $_ }) -join ' - '
    [void]$sb.Append("<tr><td style=`"padding:20px 32px 22px 32px;`"><table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"border-top:1px solid $($h.Line);`"><tr>")
    $foot = "$($h.Mono)font-size:10.5px;letter-spacing:0.06em;color:$($h.Muted);padding-top:12px;"
    [void]$sb.Append("<td style=`"$foot`">$(& $enc $footLeft)</td>")
    [void]$sb.Append("<td align=`"center`" style=`"$foot`">$(& $enc $Title)</td>")
    [void]$sb.Append("<td align=`"right`" style=`"$foot`">$(& $enc $now.ToString('yyyy-MM-dd HH:mm'))</td>")
    [void]$sb.Append('</tr></table></td></tr>')
    [void]$sb.Append('</table></td></tr></table></body></html>')
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

function Connect-ToolkitGraph {
    <#
        Connects to Microsoft Graph, or reuses an existing connection.
        Sign-in settings come from the calling script's parameters (-Bound $PSBoundParameters) first,
        then built-in defaults. Returns $true if this call made the connection,
        so the caller knows to run Disconnect-MgGraph at the end.
    #>
    param(
        [hashtable]$Config,
        [string[]]$Scopes,
        $Bound,
        [string]$RoleHint = 'Global Reader'
    )
    $boundKeys = if ($Bound) { @($Bound.Keys) } else { @() }
    $setting = { param($Name, $Default) if ($boundKeys -contains $Name) { $Bound[$Name] } else { Get-Setting $Config 'M365' $Name $Default } }
    $tenantId   = & $setting 'TenantId' ''
    $clientId   = & $setting 'ClientId' ''
    $thumbprint = & $setting 'CertificateThumbprint' ''
    $deviceCode = [bool](& $setting 'UseDeviceCode' $false)

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'The Microsoft Graph PowerShell module is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $context = Get-MgContext
    if ($context -and (-not $tenantId -or $context.TenantId -eq $tenantId -or $tenantId -like '*.*')) {
        Write-Host "Using existing Microsoft Graph connection ($($context.Account)$($context.AppName))."
        if ($context.AuthType -eq 'Delegated') {
            $missing = @($Scopes | Where-Object { $context.Scopes -notcontains $_ })
            if ($missing.Count) { Write-Warning "Current connection is missing scope(s): $($missing -join ', '). Some sections may be empty. Run Disconnect-MgGraph and try again." }
        }
        return $false
    }

    $connect = @{ NoWelcome = $true; ErrorAction = 'Stop' }
    if ($tenantId) { $connect.TenantId = $tenantId }
    if ($clientId -and $thumbprint) {
        if (-not $tenantId) { throw 'App-only sign-in needs a tenant. Use -TenantId.' }
        $connect.ClientId = $clientId
        $connect.CertificateThumbprint = $thumbprint
        Write-Host 'Connecting to Microsoft Graph (app-only, certificate)...'
    }
    else {
        $connect.Scopes = $Scopes
        Write-Host "Connecting to Microsoft Graph - sign in with an account that has $RoleHint..."
        if ($deviceCode) {
            $connect.UseDeviceCode = $true
            Write-Host 'Open the link below in any browser and enter the code shown.' -ForegroundColor Cyan
        }
        else {
            Write-Host 'If the sign-in window goes blank or hangs, press Ctrl+C and run again with -UseDeviceCode.' -ForegroundColor Cyan
        }
    }
    Connect-MgGraph @connect | Out-Null
    return $true
}

function Get-GraphRunBy {
    $c = Get-MgContext
    if ($c.Account) { return $c.Account }
    return "App: $($c.AppName)"
}

function Get-GraphTenantName {
    try {
        $org = @(Get-GraphAll 'v1.0/organization?$select=displayName,verifiedDomains') | Select-Object -First 1
        $defaultDomain = @($org.verifiedDomains | Where-Object { $_.isDefault }) | Select-Object -First 1
        return "$($org.displayName) ($($defaultDomain.name))"
    }
    catch { return "$((Get-MgContext).TenantId)" }
}

function Invoke-Graph([string]$Uri) {
    # GET with a simple retry for throttling and transient errors
    for ($attempt = 1; ; $attempt++) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            if ($attempt -lt 5 -and $_.Exception.Message -match '429|TooManyRequests|503|ServiceUnavailable|504|GatewayTimeout') {
                Start-Sleep -Seconds ([math]::Min(30, [math]::Pow(2, $attempt)))
                continue
            }
            throw
        }
    }
}

function Get-GraphAll([string]$Uri) {
    # Follows @odata.nextLink and returns every item
    $next = $Uri
    while ($next) {
        $page = Invoke-Graph $next
        foreach ($item in @($page.value)) { if ($null -ne $item) { $item } }
        $next = $page.'@odata.nextLink'
    }
}

function ConvertTo-LocalTime($Value) {
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Utc) { return $Value.ToLocalTime() }
        return $Value
    }
    if ($Value -is [datetimeoffset]) { return $Value.LocalDateTime }
    $dto = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse("$Value", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) {
        return $dto.LocalDateTime
    }
    return $null
}

#endregion
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('StaleDays')) { $StaleDays = [int](Get-Setting $config 'Devices' 'StaleSyncDays' 30) }
$Platform = @(Split-ListValue $Platform)
$bad = @($Platform | Where-Object { $_ -notin 'Windows', 'macOS', 'iOS', 'Android', 'Linux' })
if ($bad.Count) { throw "Unknown platform: $($bad -join ', '). Use Windows, macOS, iOS, Android or Linux." }

$scopes = @('DeviceManagementManagedDevices.Read.All', 'DeviceManagementConfiguration.Read.All')
if (-not $SkipDefender) { $scopes += 'ThreatHunting.Read.All', 'SecurityAlert.Read.All' }

# Intune compliance states that count as "not compliant"
$badStates = @('noncompliant', 'inGracePeriod', 'conflict', 'error')
$stateNames = @{
    compliant = 'Compliant'; noncompliant = 'Not compliant'; inGracePeriod = 'In grace period'; conflict = 'Conflict'
    error = 'Error'; unknown = 'Not evaluated'; configManager = 'Managed by ConfigMgr'
}

# Friendly names for common compliance settings (others are split from their internal name)
$settingNames = @{
    # Built-in checks from the default compliance policy
    RequireRemainContact                  = 'Device has not checked in within the compliance validity period'
    RequireDeviceCompliancePolicyAssigned = 'No compliance policy assigned'
    RequireUserExistence                  = 'Enrolled user no longer exists'
    # Windows
    BitLockerEnabled                      = 'BitLocker'
    SecureBootEnabled                     = 'Secure Boot'
    CodeIntegrityEnabled                  = 'Code integrity'
    StorageRequireEncryption              = 'Encryption of data storage'
    TpmRequired                           = 'TPM'
    ActiveFirewallRequired                = 'Firewall'
    AntivirusRequired                     = 'Antivirus'
    AntiSpywareRequired                   = 'Antispyware'
    DefenderEnabled                       = 'Microsoft Defender Antimalware'
    RtpEnabled                            = 'Real-time protection'
    SignatureOutOfDate                    = 'Defender security intelligence up to date'
    DeviceThreatProtectionEnabled         = 'Defender for Endpoint machine risk score'
    DeviceThreatProtectionRequiredSecurityLevel = 'Defender for Endpoint machine risk score'
    OsMinimumVersion                      = 'Minimum OS version'
    OsMaximumVersion                      = 'Maximum OS version'
    PasswordRequired                      = 'Password required'
    # Mobile
    PasscodeRequired                      = 'Passcode required'
    SecurityBlockJailbrokenDevices        = 'Jailbroken / rooted devices blocked'
}

# Defender sensor states that mean it isn't reporting properly
$healthySensor = @('Active', '')

# --- Helpers ---------------------------------------------------------------------

function Get-SettingLabel([string]$SettingName) {
    <# "Windows10CompliancePolicy.BitLockerEnabled" -> "BitLocker Enabled" #>
    $leaf = ($SettingName -split '\.')[-1]
    if ($settingNames.ContainsKey($leaf)) { return $settingNames[$leaf] }
    if (-not $leaf) { return $SettingName }
    return ([regex]::Replace($leaf, '(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])', ' ')).Trim()
}

function Get-PlatformName([string]$Os) {
    switch -Wildcard ($Os) {
        'Windows*' { 'Windows' } 'mac*' { 'macOS' } 'iOS*' { 'iOS' } 'iPadOS*' { 'iOS' } 'Android*' { 'Android' } 'Linux*' { 'Linux' } default { $Os }
    }
}

function Get-ShortName([string]$Name) {
    <# Defender device names are FQDNs; Intune names aren't. #>
    return (($Name -split '\.')[0]).ToLowerInvariant()
}

function Invoke-HuntingQuery([string]$Query) {
    <# Runs an advanced hunting query through Graph and returns the result rows. #>
    $body = @{ Query = $Query } | ConvertTo-Json -Compress
    for ($attempt = 1; ; $attempt++) {
        try {
            $res = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/security/runHuntingQuery' -Body $body -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
            return @($res.results)
        }
        catch {
            if ($attempt -lt 4 -and $_.Exception.Message -match '429|TooManyRequests|503|504') { Start-Sleep -Seconds (10 * $attempt); continue }
            # The useful part of a BadRequest (e.g. an unknown column) is in the response body
            $detail = "$($_.ErrorDetails.Message)"
            if ($detail -match '"message"\s*:\s*"([^"]+)"') { $detail = $Matches[1] }
            throw "$($_.Exception.Message)$(if ($detail) { " - $detail" })"
        }
    }
}

function Get-FirstValue($Object, [string[]]$Names) {
    <# Hunting results can come back with different key casing - read whichever exists. #>
    foreach ($n in $Names) { $p = $Object.PSObject.Properties[$n]; if ($p -and "$($p.Value)" -ne '') { return $p.Value } }
    return $null
}

# --- Main ------------------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'Global Reader (or Intune Administrator) plus Security Reader'
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'NonCompliantDevices' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        function Add-ReportError([string]$Section, [string]$Message) { $errors.Add([pscustomobject]@{ Section = $Section; Error = $Message }) }
        $tenantName = Get-GraphTenantName
        Write-Host "Tenant : $tenantName"

        # --- 1. Intune devices -------------------------------------------------------

        Write-Host 'Getting Intune devices...'
        $select = 'id,deviceName,userPrincipalName,userDisplayName,operatingSystem,osVersion,complianceState,lastSyncDateTime,enrolledDateTime,' +
                  'model,manufacturer,serialNumber,managedDeviceOwnerType,isEncrypted,azureADDeviceId,complianceGracePeriodExpirationDateTime,managementAgent,jailBroken'
        $allDevices = @()
        try { $allDevices = @(Get-GraphAll "v1.0/deviceManagement/managedDevices?`$select=$select&`$top=999") }
        catch { throw "Could not read Intune devices (needs an Intune licence and DeviceManagementManagedDevices.Read.All): $($_.Exception.Message)" }
        $everyDevice = $allDevices   # unfiltered, to tell which Defender devices Intune manages
        if ($Platform.Count) { $allDevices = @($allDevices | Where-Object { $Platform -contains (Get-PlatformName $_.operatingSystem) }) }
        $inScope = if ($IncludeUnknown) { $badStates + 'unknown' } else { $badStates }
        $flagged = @($allDevices | Where-Object { $inScope -contains "$($_.complianceState)" })
        Write-Host ("  {0} device(s), {1} not compliant" -f $allDevices.Count, $flagged.Count)

        # --- 2. Why each device fails ------------------------------------------------

        $settingRows = [System.Collections.Generic.List[object]]::new()
        $reasons = @{}   # device id -> @{ Policies = list; Settings = list }
        if (-not $SkipSettingDetail) {
            $i = 0
            foreach ($d in $flagged) {
                $i++
                Write-Progress -Activity 'Reading failing compliance policies' -Status "$i of $($flagged.Count): $($d.deviceName)" -PercentComplete ($i / [math]::Max(1, $flagged.Count) * 100)
                $entry = @{ Policies = [System.Collections.Generic.List[string]]::new(); Settings = [System.Collections.Generic.List[string]]::new() }
                $reasons[$d.id] = $entry
                try {
                    $policies = @(Get-GraphAll "v1.0/deviceManagement/managedDevices/$($d.id)/deviceCompliancePolicyStates" |
                                  Where-Object { "$($_.state)" -in 'nonCompliant', 'error', 'conflict' })
                    foreach ($p in $policies) {
                        $policyName = if ($p.displayName) { $p.displayName } else { 'Default device compliance policy' }
                        if (-not $entry.Policies.Contains($policyName)) { $entry.Policies.Add($policyName) }
                        $settings = @(Get-GraphAll "v1.0/deviceManagement/managedDevices/$($d.id)/deviceCompliancePolicyStates/$($p.id)/settingStates" |
                                      Where-Object { "$($_.state)" -in 'nonCompliant', 'error', 'conflict' })
                        foreach ($s in $settings) {
                            $label = Get-SettingLabel "$($s.settingName)$(if (-not $s.settingName) { $s.setting })"
                            if (-not $entry.Settings.Contains($label)) { $entry.Settings.Add($label) }
                            $settingRows.Add([pscustomobject]@{
                                Device = $d.deviceName; User = $d.userPrincipalName; Platform = Get-PlatformName $d.operatingSystem
                                Policy = $policyName; Setting = $label; State = "$($s.state)"; SettingId = "$($s.settingName)"
                            })
                        }
                    }
                }
                catch { Add-ReportError 'Compliance detail' "$($d.deviceName): $($_.Exception.Message)" }
            }
            Write-Progress -Activity 'Reading failing compliance policies' -Completed
        }

        # --- 3. Defender -------------------------------------------------------------

        $defender = @()
        $defByAad = @{}; $defByName = @{}
        $alertsByDevice = @{}
        $alertRows = [System.Collections.Generic.List[object]]::new()
        $defenderAvailable = $false
        if (-not $SkipDefender) {
            Write-Host 'Getting Defender for Endpoint devices...'
            # The full query first; if the tenant's schema doesn't have every column, a simpler one
            $queries = @(
                @'
DeviceInfo
| where Timestamp > ago(30d) and (isempty(DeviceCategory) or DeviceCategory == "Endpoint")
| summarize arg_max(Timestamp, *) by DeviceId
| where isempty(MergedToDeviceId)
| project DeviceId, DeviceName, AadDeviceId, OSPlatform, OSVersion, OnboardingStatus, SensorHealthState, ExposureLevel, ClientVersion, MachineGroup, LastSeen = Timestamp
'@
                @'
DeviceInfo
| where Timestamp > ago(30d)
| summarize arg_max(Timestamp, DeviceName, AadDeviceId, OSPlatform, OSVersion, OnboardingStatus, SensorHealthState, ExposureLevel) by DeviceId
| project DeviceId, DeviceName, AadDeviceId, OSPlatform, OSVersion, OnboardingStatus, SensorHealthState, ExposureLevel, LastSeen = Timestamp
'@
                @'
DeviceInfo
| where Timestamp > ago(30d)
| summarize arg_max(Timestamp, DeviceName, OSPlatform, OnboardingStatus) by DeviceId
| project DeviceId, DeviceName, OSPlatform, OnboardingStatus, LastSeen = Timestamp
'@
            )
            try {
                $huntRows = $null
                $failures = @()
                foreach ($q in $queries) {
                    try { $huntRows = @(Invoke-HuntingQuery $q); break }
                    catch { $failures += $_.Exception.Message }
                }
                if ($null -eq $huntRows) { throw ($failures | Select-Object -Last 1) }
                if ($failures.Count) { Add-ReportError 'Defender devices' "Full Defender query failed, so a simpler one was used (some Defender columns may be blank): $($failures[0])" }
                $defender = @($huntRows | ForEach-Object {
                    [pscustomobject]@{
                        DeviceId      = Get-FirstValue $_ 'DeviceId', 'deviceId'
                        DeviceName    = Get-FirstValue $_ 'DeviceName', 'deviceName'
                        AadDeviceId   = "$(Get-FirstValue $_ 'AadDeviceId', 'aadDeviceId')".ToLowerInvariant()
                        OSPlatform    = Get-FirstValue $_ 'OSPlatform', 'osPlatform'
                        OSVersion     = Get-FirstValue $_ 'OSVersion', 'osVersion'
                        Onboarding    = "$(Get-FirstValue $_ 'OnboardingStatus', 'onboardingStatus')"
                        SensorHealth  = "$(Get-FirstValue $_ 'SensorHealthState', 'sensorHealthState')"
                        Exposure      = "$(Get-FirstValue $_ 'ExposureLevel', 'exposureLevel')"
                        ClientVersion = Get-FirstValue $_ 'ClientVersion', 'clientVersion'
                        MachineGroup  = Get-FirstValue $_ 'MachineGroup', 'machineGroup'
                        LastSeen      = ConvertTo-LocalTime (Get-FirstValue $_ 'LastSeen', 'lastSeen')
                    }
                })
                $defenderAvailable = $true
                foreach ($m in $defender) {
                    if ($m.AadDeviceId -and $m.AadDeviceId -ne '00000000-0000-0000-0000-000000000000') { $defByAad[$m.AadDeviceId] = $m }
                    if ($m.DeviceName) { $defByName[(Get-ShortName $m.DeviceName)] = $m }
                }
                Write-Host "  $($defender.Count) device(s) seen in the last 30 days"
            }
            catch {
                Add-ReportError 'Defender devices' "Advanced hunting unavailable, so Defender columns are blank (needs Defender for Endpoint, ThreatHunting.Read.All and Security Reader): $($_.Exception.Message)"
                Write-Warning 'Defender device data is unavailable - see the Errors sheet. Continuing without it.'
            }

            Write-Host 'Getting active Defender alerts...'
            try {
                $alerts = @(Get-GraphAll "v1.0/security/alerts_v2?`$filter=serviceSource eq 'microsoftDefenderForEndpoint' and status ne 'resolved'&`$top=2000")
                foreach ($a in $alerts) {
                    $devices = @($a.evidence | Where-Object { "$($_.'@odata.type')" -like '*deviceEvidence' })
                    foreach ($ev in $devices) {
                        $keys = @("$($ev.mdeDeviceId)", "aad:$("$($ev.azureAdDeviceId)".ToLowerInvariant())", "name:$(Get-ShortName "$($ev.deviceDnsName)")") | Where-Object { $_ -and $_ -notmatch ':$' }
                        foreach ($k in $keys) {
                            if (-not $alertsByDevice.ContainsKey($k)) { $alertsByDevice[$k] = [System.Collections.Generic.List[object]]::new() }
                            if (-not ($alertsByDevice[$k] | Where-Object { $_.id -eq $a.id })) { $alertsByDevice[$k].Add($a) }
                        }
                        $alertRows.Add([pscustomobject]@{
                            Severity = (Get-Culture).TextInfo.ToTitleCase("$($a.severity)"); Device = $ev.deviceDnsName; Alert = $a.title; Category = $a.category
                            Status = "$($a.status)"; Created = ConvertTo-LocalTime $a.createdDateTime; Link = $a.alertWebUrl
                        })
                    }
                }
                Write-Host "  $($alerts.Count) active alert(s)"
            }
            catch { Add-ReportError 'Defender alerts' "Alerts unavailable (needs SecurityAlert.Read.All and Security Reader): $($_.Exception.Message)" }
        }

        function Find-DefenderDevice($IntuneDevice) {
            $aad = "$($IntuneDevice.azureADDeviceId)".ToLowerInvariant()
            if ($aad -and $defByAad.ContainsKey($aad)) { return $defByAad[$aad] }
            $short = Get-ShortName $IntuneDevice.deviceName
            if ($defByName.ContainsKey($short)) { return $defByName[$short] }
            return $null
        }
        function Get-DeviceAlerts($Machine, $IntuneDevice) {
            $found = @()
            $keys = @()
            if ($Machine) { $keys += "$($Machine.DeviceId)"; $keys += "name:$(Get-ShortName $Machine.DeviceName)" }
            if ($IntuneDevice) { $keys += "aad:$("$($IntuneDevice.azureADDeviceId)".ToLowerInvariant())"; $keys += "name:$(Get-ShortName $IntuneDevice.deviceName)" }
            foreach ($k in $keys) { if ($alertsByDevice.ContainsKey($k)) { $found += @($alertsByDevice[$k]) } }
            return @($found | Sort-Object id -Unique)
        }
        function Get-DefenderIssues($Machine, [object[]]$Alerts) {
            $out = @()
            if ($Machine) {
                # "Unsupported" and "Insufficient info" are left out - they are usually phones, printers and the like
                if ($Machine.Onboarding -eq 'Can be onboarded') { $out += 'Defender: not onboarded (can be onboarded)' }
                elseif ($Machine.Onboarding -eq 'Onboarded' -and $healthySensor -notcontains $Machine.SensorHealth) { $out += "Defender sensor: $($Machine.SensorHealth)" }
                if ($Machine.Exposure -eq 'High') { $out += 'Defender exposure level High' }
            }
            $hi = @($Alerts | Where-Object { "$($_.severity)" -eq 'high' }).Count
            if ($Alerts.Count) { $out += "$($Alerts.Count) active Defender alert(s)$(if ($hi) { " ($hi high)" })" }
            return $out
        }

        # --- 4. Devices with Defender problems that Intune says are fine -----------

        $extraIntune = @()
        if ($defenderAvailable -or $alertsByDevice.Count) {
            $flaggedIds = @{}; foreach ($d in $flagged) { $flaggedIds[$d.id] = $true }
            $extraIntune = @($allDevices | Where-Object { -not $flaggedIds.ContainsKey($_.id) } | Where-Object {
                $m = Find-DefenderDevice $_
                @(Get-DefenderIssues $m (Get-DeviceAlerts $m $_)).Count -gt 0
            })
        }

        # --- 5. Antivirus state for Windows devices in the report -----------------

        $avById = @{}
        $avScope = if ($AntivirusForAllDevices) { $allDevices } else { @($flagged) + @($extraIntune) }
        $windowsInReport = @($avScope | Where-Object { (Get-PlatformName $_.operatingSystem) -eq 'Windows' })
        if ($windowsInReport.Count) {
            $i = 0
            foreach ($d in $windowsInReport) {
                $i++
                Write-Progress -Activity 'Reading antivirus state' -Status "$i of $($windowsInReport.Count): $($d.deviceName)" -PercentComplete ($i / $windowsInReport.Count * 100)
                try { $avById[$d.id] = Invoke-Graph "v1.0/deviceManagement/managedDevices/$($d.id)/windowsProtectionState" }
                catch { if ($_.Exception.Message -notmatch '404|NotFound') { Add-ReportError 'Antivirus state' "$($d.deviceName): $($_.Exception.Message)" } }
            }
            Write-Progress -Activity 'Reading antivirus state' -Completed
        }

        # --- 6. Build rows -----------------------------------------------------------

        Write-Host 'Building report...'
        $rank = @{ OK = 0; Low = 1; Medium = 2; High = 3 }
        $matchedDefender = @{}
        $attentionIds = @{}; foreach ($d in @(@($flagged) + @($extraIntune))) { $attentionIds[$d.id] = $true }
        $allRows = foreach ($d in $allDevices) {
            $m       = Find-DefenderDevice $d
            if ($m) { $matchedDefender[$m.DeviceId] = $true }
            $alerts  = @(Get-DeviceAlerts $m $d)
            $av      = $avById[$d.id]
            $why     = $reasons[$d.id]
            $state   = "$($d.complianceState)"
            $lastSync = ConvertTo-LocalTime $d.lastSyncDateTime
            $syncDays = if ($lastSync) { [int][math]::Floor(($now - $lastSync).TotalDays) } else { $null }
            $devPlatform = Get-PlatformName $d.operatingSystem

            $issues = [System.Collections.Generic.List[string]]::new()
            $current = @{ Risk = 'OK' }
            $raise = { param($Level, $Text) $issues.Add($Text); if ($rank[$Level] -gt $rank[$current.Risk]) { $current.Risk = $Level } }

            switch ($state) {
                'noncompliant'  { & $raise 'Medium' 'Not compliant in Intune' }
                'inGracePeriod' { & $raise 'Low' "In grace period until $((ConvertTo-LocalTime $d.complianceGracePeriodExpirationDateTime).ToString('yyyy-MM-dd HH:mm'))" }
                'conflict'      { & $raise 'Low' 'Compliance policy conflict' }
                'error'         { & $raise 'Low' 'Compliance evaluation error' }
                'unknown'       { & $raise 'Low' 'Compliance not evaluated' }
            }
            if ($state -ne 'compliant' -and $d.isEncrypted -eq $false -and $devPlatform -in 'Windows', 'macOS') { & $raise 'High' 'Not compliant and not encrypted' }
            if ("$($d.jailBroken)" -eq 'True') { & $raise 'High' 'Jailbroken / rooted' }
            if ($null -ne $syncDays -and $syncDays -gt $StaleDays) { & $raise 'Low' "No Intune check-in for $syncDays days" }

            if ($av) {
                if ($av.malwareProtectionEnabled -eq $false) { & $raise 'High' 'Antivirus off' }
                elseif ($av.realTimeProtectionEnabled -eq $false) { & $raise 'High' 'Real-time protection off' }
                if ($av.signatureUpdateOverdue -eq $true) { & $raise 'Medium' 'Antivirus signatures out of date' }
                if ($av.tamperProtectionEnabled -eq $false) { & $raise 'Medium' 'Tamper protection off' }
                if ($av.quickScanOverdue -eq $true -or $av.fullScanOverdue -eq $true) { & $raise 'Low' 'Antivirus scan overdue' }
            }

            if ($defenderAvailable -and $devPlatform -eq 'Windows' -and -not $m) { & $raise 'Medium' 'Not seen in Defender in the last 30 days' }
            foreach ($t in (Get-DefenderIssues $m $alerts)) {
                & $raise $(if ($t -match 'high\)|sensor|not onboarded') { 'High' } else { 'Medium' }) $t
            }

            [pscustomobject]@{
                Risk              = $current.Risk
                Device            = $d.deviceName
                User              = $d.userPrincipalName
                UserName          = $d.userDisplayName
                Platform          = $devPlatform
                OSVersion         = $d.osVersion
                IntuneCompliance  = if ($stateNames.ContainsKey($state)) { $stateNames[$state] } else { $state }
                FailingPolicies   = if ($why) { ($why.Policies -join '; ') } elseif ($SkipSettingDetail) { '(not read)' } else { '' }
                FailingSettings   = if ($why) { ($why.Settings -join '; ') } else { '' }
                Issues            = ($issues -join '; ')
                LastSync          = $lastSync
                DaysSinceSync     = $syncDays
                GraceExpires      = if ($state -eq 'inGracePeriod') { ConvertTo-LocalTime $d.complianceGracePeriodExpirationDateTime } else { $null }
                Encrypted         = $d.isEncrypted
                DefenderStatus    = if (-not $defenderAvailable) { 'Not checked' } elseif ($m) { $m.Onboarding } else { 'Not found' }
                SensorHealth      = if ($m) { $m.SensorHealth } else { '' }
                ExposureLevel     = if ($m) { $m.Exposure } else { '' }
                ActiveAlerts      = $alerts.Count
                RealTimeProtection = if ($av) { [bool]$av.realTimeProtectionEnabled } else { '' }
                SignaturesOverdue = if ($av) { [bool]$av.signatureUpdateOverdue } else { '' }
                TamperProtection  = if ($av) { [bool]$av.tamperProtectionEnabled } else { '' }
                SignatureVersion  = if ($av) { $av.signatureVersion } else { '' }
                Ownership         = "$($d.managedDeviceOwnerType)"
                Manufacturer      = $d.manufacturer
                Model             = $d.model
                SerialNumber      = $d.serialNumber
                Enrolled          = ConvertTo-LocalTime $d.enrolledDateTime
                IntuneDeviceId    = $d.id
                EntraDeviceId     = $d.azureADDeviceId
                DefenderDeviceId  = if ($m) { $m.DeviceId } else { '' }
            }
        }
        $allRows = @($allRows | Sort-Object @{ Expression = { $rank[$_.Risk] }; Descending = $true }, Device)
        $rows    = @($allRows | Where-Object { $attentionIds.ContainsKey($_.IntuneDeviceId) })

        # Defender devices with problems that aren't in Intune at all
        $defOnly = @()
        if ($defenderAvailable) {
            $intuneAad = @{}; $intuneNames = @{}
            foreach ($d in $everyDevice) { if ($d.azureADDeviceId) { $intuneAad["$($d.azureADDeviceId)".ToLowerInvariant()] = $true }; $intuneNames[(Get-ShortName $d.deviceName)] = $true }
            $defOnly = @($defender | Where-Object { -not $Platform.Count -or $Platform -contains (Get-PlatformName $_.OSPlatform) } | Where-Object { -not $matchedDefender.ContainsKey($_.DeviceId) -and -not $intuneAad.ContainsKey($_.AadDeviceId) -and -not $intuneNames.ContainsKey((Get-ShortName $_.DeviceName)) } |
                ForEach-Object {
                    $alerts = @(Get-DeviceAlerts $_ $null)
                    $iss = @(Get-DefenderIssues $_ $alerts)
                    if (-not $iss.Count) { return }
                    [pscustomobject]@{
                        Device = $_.DeviceName; Platform = $_.OSPlatform; OSVersion = $_.OSVersion; Onboarding = $_.Onboarding; SensorHealth = $_.SensorHealth
                        ExposureLevel = $_.Exposure; ActiveAlerts = $alerts.Count; Issues = ($iss -join '; '); LastSeen = $_.LastSeen
                        DeviceGroup = $_.MachineGroup; SensorVersion = $_.ClientVersion; DefenderDeviceId = $_.DeviceId
                    }
                } | Sort-Object Device)
        }

        $bySetting = @($settingRows | Group-Object Platform, Policy, Setting | ForEach-Object {
            $g = $_.Group
            [pscustomobject]@{ Setting = $g[0].Setting; Policy = $g[0].Policy; Platform = $g[0].Platform; Devices = @($g.Device | Select-Object -Unique).Count }
        } | Sort-Object Devices -Descending)

        # --- 7. Summary and export ---------------------------------------------------

        $counts = @{}; foreach ($k in 'High', 'Medium', 'Low', 'OK') { $counts[$k] = @($rows | Where-Object Risk -eq $k).Count }
        $stateCounts = $allDevices | Group-Object complianceState | ForEach-Object { "$(if ($stateNames.ContainsKey($_.Name)) { $stateNames[$_.Name] } else { $_.Name }): $($_.Count)" }
        $compliantPct = if ($allDevices.Count) { [math]::Round(@($allDevices | Where-Object complianceState -eq 'compliant').Count / $allDevices.Count * 100, 1) } else { 0 }
        $summary = [ordered]@{
            'Tenant'                         = $tenantName
            'Run at'                         = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Platforms'                      = if ($Platform.Count) { $Platform -join ', ' } else { 'All' }
            'Intune devices'                 = $allDevices.Count
            'Compliant'                      = "$compliantPct%"
            'By compliance state'            = ($stateCounts -join ', ')
            'Not compliant in Intune'        = $flagged.Count
            'Compliant, but Defender issues' = $extraIntune.Count
            'Defender-only devices with issues' = if ($defenderAvailable) { $defOnly.Count } else { 'Not checked' }
            'Risk'                           = "$($counts.High) high, $($counts.Medium) medium, $($counts.Low) low"
            'Most common failing setting'    = if ($bySetting.Count) { "$($bySetting[0].Setting) ($($bySetting[0].Devices) device(s))" } else { '' }
            'Note'                           = "Defender data covers devices seen in the last 30 days. Devices are matched by Entra device ID, then by name. Stale = no Intune check-in for $StaleDays+ days."
            'Run by'                         = Get-GraphRunBy
        }

        $deviceHeaders = @('Risk', 'Device', 'User', 'Platform', 'OSVersion', 'IntuneCompliance', 'FailingPolicies', 'FailingSettings', 'Issues', 'LastSync', 'DaysSinceSync', 'GraceExpires',
                           'Encrypted', 'DefenderStatus', 'SensorHealth', 'ExposureLevel', 'ActiveAlerts', 'RealTimeProtection', 'SignaturesOverdue', 'TamperProtection', 'SignatureVersion',
                           'Ownership', 'Manufacturer', 'Model', 'SerialNumber', 'Enrolled', 'UserName', 'IntuneDeviceId', 'EntraDeviceId', 'DefenderDeviceId')
        $deviceWidths  = @(9, 24, 32, 10, 16, 16, 40, 55, 70, 18, 14, 18, 10, 15, 16, 14, 12, 18, 17, 16, 16, 11, 16, 24, 18, 18, 24, 38, 38, 42)
        $sheets = @(
            @{ Name = 'Devices'; Rows = $rows; Headers = $deviceHeaders; Widths = $deviceWidths }
            @{ Name = 'All Devices'; Rows = $allRows; Headers = $deviceHeaders; Widths = $deviceWidths }
            @{ Name = 'Failing Settings'; Rows = @($settingRows | Sort-Object Device, Policy, Setting)
               Headers = @('Device', 'User', 'Platform', 'Policy', 'Setting', 'State', 'SettingId'); Widths = @(24, 32, 10, 40, 50, 13, 60) }
            @{ Name = 'By Setting'; Rows = $bySetting; Headers = @('Setting', 'Policy', 'Platform', 'Devices'); Widths = @(55, 40, 10, 9) }
        )
        if (-not $SkipDefender) {
            $sheets += @{ Name = 'Defender Only'; Rows = $defOnly
                          Headers = @('Device', 'Platform', 'OSVersion', 'Onboarding', 'SensorHealth', 'ExposureLevel', 'ActiveAlerts', 'Issues', 'LastSeen', 'DeviceGroup', 'SensorVersion', 'DefenderDeviceId')
                          Widths  = @(34, 12, 14, 16, 22, 14, 12, 60, 18, 20, 16, 42) }
            $sheets += @{ Name = 'Active Alerts'; Rows = @($alertRows | Sort-Object @{ Expression = { @{ High = 0; Medium = 1; Low = 2 }["$($_.Severity)"] } }, Created)
                          Headers = @('Severity', 'Device', 'Alert', 'Category', 'Status', 'Created', 'Link'); Widths = @(10, 30, 55, 20, 12, 18, 60) }
        }
        $sheets += @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(22, 120) }
        $sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(34, 110) }
        if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

        # --- Console and email -------------------------------------------------------

        Write-Host ""
        Write-Host ("Compliant : {0}% of {1} Intune device(s)" -f $compliantPct, $allDevices.Count) -ForegroundColor $(if ($compliantPct -ge 95) { 'Green' } elseif ($compliantPct -ge 85) { 'Yellow' } else { 'Red' })
        Write-Host ("In report : {0} device(s) - {1} high, {2} medium, {3} low" -f $rows.Count, $counts.High, $counts.Medium, $counts.Low)
        if ($defenderAvailable) { Write-Host ("Defender-only devices with issues: {0}" -f $defOnly.Count) }
        foreach ($s in $bySetting | Select-Object -First 5) { Write-Host ("  {0,4}  {1} ({2})" -f $s.Devices, $s.Setting, $s.Platform) }
        if ($errors.Count) { Write-Host "Some data was unavailable - see the 'Errors' sheet ($($errors.Count))." -ForegroundColor Yellow }
        Write-Host "Report    : $reportPath"

        if ($SendEmail) {
            $top = @($rows | Where-Object { $_.Risk -in 'High', 'Medium' } | Select-Object -First 50)
            $html = New-HtmlReport -Title 'Non-compliant devices' -IntroLines @(
                "$tenantName - $compliantPct% of $($allDevices.Count) Intune devices are compliant. $($rows.Count) device(s) need attention: $($counts.High) high, $($counts.Medium) medium, $($counts.Low) low."
            ) -Meta ([ordered]@{ Tenant = $tenantName }) -Stats @(
                @{ Label = 'Compliant'; Value = "$compliantPct%" }
                @{ Label = 'Need attention'; Value = $rows.Count }
                @{ Label = 'High risk'; Value = $counts.High; Status = 'High' }
                @{ Label = 'Medium risk'; Value = $counts.Medium; Status = 'Medium' }
            ) -Sections @(
                @{ Heading = "High and medium risk ($(@($rows | Where-Object { $_.Risk -in 'High', 'Medium' }).Count))"
                   Html    = ConvertTo-HtmlTable -Rows $top -Columns @('Risk', 'Device', 'User', 'Platform', 'IntuneCompliance', 'Issues', 'LastSync') -StatusColumn 'Risk' -EmptyText 'None.' }
                @{ Heading = 'Most common failing settings'
                   Html    = ConvertTo-HtmlTable -Rows @($bySetting | Select-Object -First 10) -Columns @('Setting', 'Policy', 'Platform', 'Devices') -EmptyText 'None recorded.' }
            )
            $tag = if ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Warning' } else { 'OK' }
            if (Send-ToolkitMail -Config $config -Subject "[$tag] Devices: $compliantPct% compliant, $($rows.Count) need attention" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
