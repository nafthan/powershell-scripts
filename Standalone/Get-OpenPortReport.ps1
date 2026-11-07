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

    # Report labels (HTML reports and emails)
    [string]$Organisation,
    [string]$Classification = 'Internal',
    [string[]]$AllowedPort
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
    if ($AllowedPort) { $cfg.PortCheck = @{ Allowed = @{ Default = @($AllowedPort | ForEach-Object { "$_" -split "," } | Where-Object { $_.Trim() } | ForEach-Object { [int]$_.Trim() }) } } }
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
