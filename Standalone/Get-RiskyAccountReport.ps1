<#
.SYNOPSIS
    Audits AD accounts for settings that attackers commonly abuse.
.DESCRIPTION
    Read-only. Rates each finding High, Medium or Low, with a recommended fix.
.EXAMPLE
    .\Get-RiskyAccountReport.ps1
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: ActiveDirectory module (RSAT)
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$Server,
    [switch]$IncludeDisabled,
    [switch]$SendEmail,
    [string]$OutputFile,

    # Email settings (for -SendEmail / -NotifyUsers)
    [string]$SmtpServer,
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [pscredential]$SmtpCredential,
    [string]$MailFrom,
    [string[]]$MailTo,

    # Report look: a theme file (see themes\ in the repo) and the labels on HTML reports
    [string]$ThemeFile,
    [string]$Organisation,
    [string]$Classification
)

#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)

$script:ToolkitRoot   = $PSScriptRoot

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)

$script:DefaultTheme = @{
    Organisation    = ''
    Classification  = 'Internal'
    LogoUrl         = ''

    Font            = "'IBM Plex Sans', 'Segoe UI', Helvetica, Arial, sans-serif"
    LabelFont       = "'IBM Plex Mono', Consolas, 'Courier New', monospace"
    UppercaseLabels = $true
    ExcelFont       = 'Segoe UI'
    ExcelHeaderFont = 'Consolas'
    ExcelFontSize   = 10

    Primary         = '#1c1b19'
    OnPrimary       = '#e8e4dc'
    Accent          = ''
    Heading         = '#1c1b19'
    Subtitle        = '#45423d'
    Text            = '#1c1b19'
    Muted           = '#78736a'
    Page            = '#f4f2ee'
    Card            = '#ffffff'
    Line            = '#e6e2da'
    Rule            = '#d6d1c8'
    RowStripe       = ''
    LabelFill       = ''

    StatusBad       = @('#9b2c2c', '#f9ebe9', '#ecc9c4')
    StatusWarning   = @('#875a0b', '#fbf3e1', '#ecd9ab')
    StatusInfo      = @('#2d5a86', '#eaf1f8', '#c7d7e8')
    StatusGood      = @('#2f6a47', '#e9f3ec', '#c4ddcc')
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
    if ($ThemeFile) { $cfg.Theme = @{ File = $ThemeFile } }
    if ($script:PSBoundParameters.ContainsKey('Organisation')) { $cfg.General.Organisation = $Organisation }; if ($script:PSBoundParameters.ContainsKey('Classification')) { $cfg.General.Classification = $Classification }
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
    $stripe     = [bool](Get-ReportTheme).RowStripe
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
                if ($st -eq 0 -and $stripe -and $r % 2 -eq 1) { $st = 7 }
                if ($null -eq $val) {
                    $w.Write("<c r=`"$ref`" s=`"$st`"/>")   # empty but styled, so the row line is unbroken
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

        # Styles come from the report theme, so the workbook matches the HTML reports.
        #   0 body  1 header  2 red  3 amber  4 blue  5 green  6 summary label  7 body on stripe
        $t    = Get-ReportTheme
        $hex  = { param($c) "FF$($c.TrimStart('#').ToUpperInvariant())" }
        $esc  = { param($n) [System.Security.SecurityElement]::Escape($n) }
        $font = { param($Size, $Colour, $Name, [switch]$Bold) "<font>$(if ($Bold) { '<b/>' })<sz val=`"$Size`"/><color rgb=`"$(& $hex $Colour)`"/><name val=`"$(& $esc $Name)`"/></font>" }
        $fill = { param($Colour) "<fill><patternFill patternType=`"solid`"><fgColor rgb=`"$(& $hex $Colour)`"/><bgColor indexed=`"64`"/></patternFill></fill>" }
        $xf   = { param($Font, $Fill, $Border) "<xf numFmtId=`"0`" fontId=`"$Font`" fillId=`"$Fill`" borderId=`"$Border`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"left`" vertical=`"center`" indent=`"1`"/></xf>" }
        $body = ($t.ExcelFont -split ',')[0].Trim(" '")
        $head = ($t.ExcelHeaderFont -split ',')[0].Trim(" '")
        $sz   = $t.ExcelFontSize
        $stripeFill = if ($t.RowStripe) { $t.RowStripe } else { $t.Card }
        $labelFill  = if ($t.LabelFill) { $t.LabelFill } else { $t.Card }
        $accentBorder = if ($t.Accent) { "<border><left/><right/><top/><bottom style=`"medium`"><color rgb=`"$(& $hex $t.Accent)`"/></bottom><diagonal/></border>" } else { '<border><left/><right/><top/><bottom/><diagonal/></border>' }
        Add-ZipText $zip 'xl/styles.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<fonts count="7">' +
                (& $font $sz $t.Text $body) + (& $font ($sz - 1) $t.OnPrimary $head -Bold) +
                (& $font $sz $t.StatusBad[0] $body -Bold) + (& $font $sz $t.StatusWarning[0] $body -Bold) +
                (& $font $sz $t.StatusInfo[0] $body -Bold) + (& $font $sz $t.StatusGood[0] $body -Bold) +
                (& $font ($sz - 1) $t.Muted $head) + '</fonts>' +
            '<fills count="9"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
                (& $fill $t.Primary) + (& $fill $t.StatusBad[1]) + (& $fill $t.StatusWarning[1]) + (& $fill $t.StatusInfo[1]) + (& $fill $t.StatusGood[1]) +
                (& $fill $stripeFill) + (& $fill $labelFill) + '</fills>' +
            '<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>' +
                "<border><left/><right/><top/><bottom style=`"thin`"><color rgb=`"$(& $hex $t.Line)`"/></bottom><diagonal/></border>" +
                $accentBorder + '</borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="8">' +
                (& $xf 0 0 1) + (& $xf 1 2 2) + (& $xf 2 3 1) + (& $xf 3 4 1) + (& $xf 4 5 1) + (& $xf 5 6 1) + (& $xf 6 8 1) + (& $xf 0 7 1) + '</cellXfs>' +
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

function ConvertTo-ThemeColour($Value, [string]$Name, [switch]$AllowEmpty) {
    <# '#1f3a5f', '1F3A5F' or '#1f3a5f'-style values -> '#1f3a5f'. Returns $null if invalid. #>
    $v = "$Value".Trim()
    if ($v -eq '' -and $AllowEmpty) { return '' }
    if ($v -match '^#?([0-9a-fA-F]{6})$') { return "#$($Matches[1].ToLowerInvariant())" }
    Write-Warning "Theme setting $Name = '$Value' is not a colour like '#1f3a5f'. Using the default."
    return $null
}

function Get-ReportTheme {
    <#
        Returns the report theme: built-in defaults, then the theme file (Theme.File in the config,
        or -ThemeFile on a standalone script), then any other settings in the Theme section.
    #>
    if ($script:ReportThemeCache) { return $script:ReportThemeCache }
    $theme = @{}
    foreach ($k in $script:DefaultTheme.Keys) { $theme[$k] = $script:DefaultTheme[$k] }

    $cfg     = Get-Variable -Name config -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    $section = if ($cfg -is [hashtable] -and $cfg['Theme'] -is [hashtable]) { $cfg['Theme'] } else { @{} }
    $layers  = @()
    if ($section['File']) {
        $file = "$($section['File'])"
        $path = if ([System.IO.Path]::IsPathRooted($file)) { $file }
                elseif (Test-Path -LiteralPath (Join-Path (Get-Location).Path $file)) { Join-Path (Get-Location).Path $file }
                else { Resolve-ToolkitPath $file }
        try { $layers += , (Import-PowerShellDataFile -LiteralPath $path -ErrorAction Stop) }
        catch { Write-Warning "Theme file '$file' could not be read, so the default look is used. Details: $($_.Exception.Message)" }
    }
    $layers += , $section
    foreach ($layer in $layers) {
        foreach ($k in $layer.Keys) {
            if ($k -eq 'File') { continue }
            if ($k -eq 'Organisation' -and "$($layer[$k])".Trim() -eq '') { continue }   # empty means "not set here"
            if (-not $script:DefaultTheme.ContainsKey($k)) { Write-Warning "Unknown theme setting '$k' ignored."; continue }
            $theme[$k] = $layer[$k]
        }
    }

    # Validate, so a typo can't break the HTML or the workbook
    foreach ($k in 'Primary', 'OnPrimary', 'Heading', 'Subtitle', 'Text', 'Muted', 'Page', 'Card', 'Line', 'Rule') {
        $c = ConvertTo-ThemeColour $theme[$k] $k; $theme[$k] = if ($c) { $c } else { $script:DefaultTheme[$k] }
    }
    foreach ($k in 'Accent', 'RowStripe', 'LabelFill') {
        $c = ConvertTo-ThemeColour $theme[$k] $k -AllowEmpty; $theme[$k] = if ($null -ne $c) { $c } else { '' }
    }
    foreach ($k in 'StatusBad', 'StatusWarning', 'StatusInfo', 'StatusGood') {
        $parts = @($theme[$k])
        $ok = $parts.Count -eq 3 -and -not (@($parts | ForEach-Object { ConvertTo-ThemeColour $_ $k }) -contains $null)
        $theme[$k] = if ($ok) { @($parts | ForEach-Object { ConvertTo-ThemeColour $_ $k }) } else { $script:DefaultTheme[$k] }
    }
    foreach ($k in 'Font', 'LabelFont', 'ExcelFont', 'ExcelHeaderFont', 'Organisation', 'Classification') {
        $theme[$k] = ("$($theme[$k])" -replace '["<>;{}\\]', '').Trim()
    }
    if (-not $theme.Font)            { $theme.Font = $script:DefaultTheme.Font }
    if (-not $theme.LabelFont)       { $theme.LabelFont = $theme.Font }
    if (-not $theme.ExcelFont)       { $theme.ExcelFont = $script:DefaultTheme.ExcelFont }
    if (-not $theme.ExcelHeaderFont) { $theme.ExcelHeaderFont = $theme.ExcelFont }
    $size = 0; if (-not [int]::TryParse("$($theme.ExcelFontSize)", [ref]$size) -or $size -lt 6 -or $size -gt 20) { $size = 10 }
    $theme.ExcelFontSize   = $size
    $theme.UppercaseLabels = [bool]$theme.UppercaseLabels
    $logo = "$($theme.LogoUrl)".Trim()
    $theme.LogoUrl = if ($logo -match '^https://[^\s"<>]+$') { $logo } else { if ($logo) { Write-Warning 'Theme LogoUrl must be an https:// address. Ignored.' }; '' }

    # Ready-made CSS pieces used by the HTML functions
    $theme.Css = @{
        Sans  = "font-family:$($theme.Font);"
        Label = "font-family:$($theme.LabelFont);" + $(if ($theme.UppercaseLabels) { 'text-transform:uppercase;letter-spacing:0.12em;' } else { 'letter-spacing:0.02em;' })
    }
    $script:ReportThemeCache = $theme
    return $theme
}

function Get-StatusStyle([string]$Status) {
    <# Text, background and border colours for a status value, from the theme. #>
    $t = Get-ReportTheme
    $c = switch -Regex ($Status) {
        '^(Critical|Expired|Error|Unreachable|Failed|Not done|High)' { $t.StatusBad; break }
        '^(Warning|Expiring|Medium)'                                  { $t.StatusWarning; break }
        '^Low'                                                        { $t.StatusInfo; break }
        '^(OK|Done|Requested)'                                        { $t.StatusGood; break }
        default                                                       { $null }
    }
    if (-not $c) { return $null }
    return @{ Fg = $c[0]; Bg = $c[1]; Bd = $c[2] }
}

function ConvertTo-HtmlStatusBadge([string]$Status) {
    $enc = [System.Net.WebUtility]::HtmlEncode($Status)
    $c = Get-StatusStyle $Status
    if (-not $c) { return $enc }
    $t = Get-ReportTheme
    return "<span style=`"$($t.Css.Label)font-size:11px;font-weight:600;letter-spacing:0.06em;white-space:nowrap;color:$($c.Fg);background-color:$($c.Bg);border:1px solid $($c.Bd);border-radius:999px;padding:2px 9px;`">$enc</span>"
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
    $t = Get-ReportTheme
    if ($null -eq $Rows -or @($Rows).Count -eq 0) {
        return "<p style=`"$($t.Css.Sans)font-size:13px;color:$($t.Muted);margin:0 0 8px 0;`">$([System.Net.WebUtility]::HtmlEncode($EmptyText))</p>"
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"border-collapse:collapse;width:100%;border:1px solid $($t.Line);margin:0 0 8px 0;`">")
    [void]$sb.Append('<tr>')
    foreach ($c in $Columns) {
        $label = [regex]::Replace($c, '(?<=[a-z])(?=[A-Z])', ' ')   # SizeGB -> Size GB, DisplayName -> Display Name
        [void]$sb.Append("<th align=`"left`" style=`"$($t.Css.Label)font-size:10.5px;font-weight:600;color:$($t.OnPrimary);background-color:$($t.Primary);text-align:left;padding:9px 12px;white-space:nowrap;`">$([System.Net.WebUtility]::HtmlEncode($label))</th>")
    }
    [void]$sb.Append('</tr>')
    $rowList = @($Rows); $last = $rowList.Count - 1
    for ($i = 0; $i -le $last; $i++) {
        $row = $rowList[$i]
        $border = if ($i -lt $last) { "border-bottom:1px solid $($t.Line);" } else { '' }
        $stripe = if ($t.RowStripe -and $i % 2 -eq 1) { "background-color:$($t.RowStripe);" } else { '' }
        [void]$sb.Append('<tr>')
        foreach ($c in $Columns) {
            $val = $row.$c
            if ($val -is [datetime]) { $val = $val.ToString('yyyy-MM-dd HH:mm') }
            elseif ($val -is [bool]) { $val = if ($val) { 'Yes' } else { 'No' } }
            $cell = if ($c -eq $StatusColumn) { ConvertTo-HtmlStatusBadge "$val" } else { [System.Net.WebUtility]::HtmlEncode("$val") }
            [void]$sb.Append("<td style=`"$($t.Css.Sans)font-size:13px;line-height:1.45;color:$($t.Text);padding:9px 12px;vertical-align:top;$border$stripe`">$cell</td>")
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

function Get-ReportLabels {
    <# Organisation and classification: General section / script parameters first, then the theme. #>
    $t   = Get-ReportTheme
    $cfg = Get-Variable -Name config -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    $general = if ($cfg -is [hashtable] -and $cfg['General'] -is [hashtable]) { $cfg['General'] } else { @{} }
    $org = "$($general['Organisation'])".Trim()
    if (-not $org) { $org = $t.Organisation }
    # An empty Classification hides the label, so only fall back when it isn't set at all
    $class = if ($general.ContainsKey('Classification') -and $null -ne $general['Classification']) { "$($general['Classification'])".Trim() } else { $t.Classification }
    return @{ Organisation = $org; Classification = $class }
}

function New-HtmlReport {
    <#
        Wraps sections into a full HTML page (also used as the email body). Colours and fonts come from the theme.
        -Sections: array of @{ Heading = '...'; Html = '...' }
        -Stats:    optional summary tiles, array of @{ Label = '...'; Value = '...'; Status = 'Critical|Warning|OK|...' }
        -Meta:     optional extra rows for the details block, e.g. [ordered]@{ Tenant = 'Contoso' }
    #>
    param(
        [string]$Title,
        [string[]]$IntroLines,
        [object[]]$Sections,
        [object[]]$Stats,
        [System.Collections.IDictionary]$Meta
    )
    $t   = Get-ReportTheme
    $enc = { param($x) [System.Net.WebUtility]::HtmlEncode("$x") }
    $now = Get-Date
    $scriptName = Split-Path -Leaf "$($MyInvocation.PSCommandPath)"
    $labels = Get-ReportLabels
    $runBy = if ($env:USERNAME) { if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { $env:USERNAME } } elseif ($env:USER) { $env:USER } else { '' }

    $details = [ordered]@{}
    if ($Meta) { foreach ($k in $Meta.Keys) { $details[$k] = $Meta[$k] } }
    $details['Generated'] = $now.ToString('dd MMM yyyy HH:mm')
    if ($runBy) { $details['Run by'] = $runBy }
    $details['Computer'] = [System.Environment]::MachineName
    $details['Script']   = $scriptName

    $label = $t.Css.Label; $sans = $t.Css.Sans
    $sb  = New-Object System.Text.StringBuilder
    [void]$sb.Append("<!DOCTYPE html><html lang=`"en`"><head><meta charset=`"utf-8`"><meta name=`"viewport`" content=`"width=device-width,initial-scale=1`"><title>$(& $enc $Title)</title></head>")
    [void]$sb.Append("<body style=`"margin:0;padding:0;background-color:$($t.Page);`">")
    [void]$sb.Append("<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"background-color:$($t.Page);`"><tr><td align=`"center`" style=`"padding:28px 16px;`">")
    [void]$sb.Append("<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"max-width:1040px;width:100%;background-color:$($t.Card);border:1px solid $($t.Line);`">")

    # Top bar: logo / organisation on the left, classification on the right, optional accent stripe below
    $left = if ($labels.Organisation) { "$(& $enc $labels.Organisation) &nbsp;/&nbsp; Report" } else { 'Report' }
    if ($t.LogoUrl) { $left = "<img src=`"$(& $enc $t.LogoUrl)`" alt=`"$(& $enc $labels.Organisation)`" height=`"22`" style=`"height:22px;vertical-align:middle;border:0;margin-right:12px;`">$left" }
    $stripe = if ($t.Accent) { "border-bottom:4px solid $($t.Accent);" } else { '' }
    [void]$sb.Append("<tr><td style=`"background-color:$($t.Primary);padding:0;$stripe`"><table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`"><tr>")
    [void]$sb.Append("<td style=`"$($label)font-size:10.5px;color:$($t.OnPrimary);padding:10px 32px;`">$left</td>")
    if ($labels.Classification) {
        [void]$sb.Append("<td align=`"right`" style=`"padding:6px 32px 6px 0;`"><span style=`"$($label)font-size:10.5px;font-weight:600;color:$($t.Primary);background-color:$($t.OnPrimary);padding:4px 10px;`">$(& $enc $labels.Classification)</span></td>")
    }
    [void]$sb.Append('</tr></table></td></tr>')

    # Title and intro
    [void]$sb.Append("<tr><td style=`"padding:26px 32px 18px 32px;`">")
    [void]$sb.Append("<div style=`"$($sans)font-size:24px;line-height:1.25;font-weight:600;color:$($t.Heading);margin:0 0 10px 0;`">$(& $enc $Title)</div>")
    foreach ($line in $IntroLines) {
        [void]$sb.Append("<div style=`"$($sans)font-size:13.5px;line-height:1.6;color:$($t.Subtitle);margin:0;`">$(& $enc $line)</div>")
    }
    [void]$sb.Append('</td></tr>')

    # Details block (document control)
    [void]$sb.Append("<tr><td style=`"padding:0 32px 4px 32px;`"><table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"border-collapse:collapse;border-top:1px solid $($t.Rule);border-bottom:1px solid $($t.Rule);`"><tr>")
    $keys = @($details.Keys); $w = [math]::Floor(100 / $keys.Count)
    for ($i = 0; $i -lt $keys.Count; $i++) {
        $sep  = if ($i -gt 0) { "border-left:1px solid $($t.Line);" } else { '' }
        $fill = if ($t.LabelFill) { "background-color:$($t.LabelFill);" } else { '' }
        [void]$sb.Append("<td width=`"$w%`" valign=`"top`" style=`"padding:10px 12px 10px $(if ($i -or $fill) { 12 } else { 0 })px;$sep$fill`">")
        [void]$sb.Append("<div style=`"$($label)font-size:10px;color:$($t.Muted);margin:0 0 3px 0;`">$(& $enc $keys[$i])</div>")
        [void]$sb.Append("<div style=`"$($sans)font-size:13px;color:$($t.Text);word-break:break-word;`">$(& $enc $details[$keys[$i]])</div></td>")
    }
    [void]$sb.Append('</tr></table></td></tr>')

    # Summary tiles
    if ($Stats) {
        [void]$sb.Append("<tr><td style=`"padding:20px 32px 4px 32px;`"><table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`"><tr>")
        $n = @($Stats).Count; $w = [math]::Floor(100 / $n)
        for ($i = 0; $i -lt $n; $i++) {
            $s = @($Stats)[$i]
            $c = if ($s.Status) { Get-StatusStyle "$($s.Status)" } else { $null }
            $valColour = if ($c -and "$($s.Value)" -ne '0') { $c.Fg } else { $t.Text }
            $bar = if ($c) { $c.Fg } else { $t.Primary }
            $gap = if ($i -lt $n - 1) { 'padding-right:10px;' } else { '' }
            [void]$sb.Append("<td width=`"$w%`" valign=`"top`" style=`"$gap`"><div style=`"border:1px solid $($t.Line);border-top:3px solid $bar;padding:12px 14px;`">")
            [void]$sb.Append("<div style=`"$($label)font-size:10.5px;color:$($t.Muted);margin:0 0 4px 0;`">$(& $enc $s.Label)</div>")
            [void]$sb.Append("<div style=`"$($sans)font-size:22px;font-weight:600;line-height:1.2;color:$valColour;`">$(& $enc "$($s.Value)")</div></div></td>")
        }
        [void]$sb.Append('</tr></table></td></tr>')
    }

    # Numbered sections, with the accent colour (or primary) under each heading
    $ruleColour = if ($t.Accent) { $t.Accent } else { $t.Primary }
    $num = 0
    foreach ($s in $Sections) {
        $num++
        [void]$sb.Append("<tr><td style=`"padding:24px 32px 6px 32px;`">")
        [void]$sb.Append("<div style=`"border-bottom:2px solid $ruleColour;padding:0 0 7px 0;margin:0 0 12px 0;`"><span style=`"$($label)font-size:12px;color:$($t.Muted);letter-spacing:0.08em;`">$($num.ToString('00'))</span>&nbsp;&nbsp;<span style=`"$($sans)font-size:15px;font-weight:600;color:$($t.Heading);`">$(& $enc $s.Heading)</span></div>")
        [void]$sb.Append($s.Html)
        [void]$sb.Append('</td></tr>')
    }

    # Footer: organisation and classification / title / generated
    $footLeft = (@($labels.Organisation, $labels.Classification) | Where-Object { $_ }) -join ' - '
    [void]$sb.Append("<tr><td style=`"padding:20px 32px 22px 32px;`"><table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"border-top:1px solid $($t.Line);`"><tr>")
    $foot = "$($label)font-size:10.5px;letter-spacing:0.06em;text-transform:none;color:$($t.Muted);padding-top:12px;"
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

if (-not $PSBoundParameters.ContainsKey('SearchBase')) { $SearchBase = Get-Setting $config 'RiskyAccounts' 'SearchBase' '' }
if (-not $PSBoundParameters.ContainsKey('Server'))     { $Server     = Get-Setting $config 'ActiveDirectory' 'Server' '' }

$privMaxDays   = [int](Get-Setting $config 'RiskyAccounts' 'PrivilegedPasswordMaxDays' 365)
$krbtgtMaxDays = [int](Get-Setting $config 'RiskyAccounts' 'KrbtgtMaxDays' 180)
$accepted      = Get-Setting $config 'RiskyAccounts' 'AcceptedRisks' @{}
if (-not ($accepted -is [hashtable])) { $accepted = @{} }

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { throw 'The ActiveDirectory module is not installed. Install RSAT: Active Directory tools (or run on a domain controller).' }

$ad = @{}
if ($Server) { $ad.Server = $Server }
$search = @{} + $ad
if ($SearchBase) { $search.SearchBase = $SearchBase }

$now        = Get-Date
$reportPath = New-ReportPath -Config $config -Prefix 'RiskyAccounts' -OutputFile $OutputFile
$domain     = Get-ADDomain @ad
$domainSid  = $domain.DomainSID.Value

Write-Host "Domain : $($domain.DNSRoot)"
if ($SearchBase) { Write-Host "Search : $SearchBase" }
Write-Host ""

# --- Privileged accounts --------------------------------------------------

Write-Host 'Finding privileged accounts...'
$privGroups = [ordered]@{
    'Domain Admins'     = "$domainSid-512"
    'Enterprise Admins' = "$domainSid-519"
    'Schema Admins'     = "$domainSid-518"
    'Administrators'    = 'S-1-5-32-544'
    'Account Operators' = 'S-1-5-32-548'
    'Server Operators'  = 'S-1-5-32-549'
    'Backup Operators'  = 'S-1-5-32-551'
    'Print Operators'   = 'S-1-5-32-550'
}
$privileged = @{}   # DistinguishedName -> list of group names
foreach ($name in $privGroups.Keys) {
    try {
        $group = Get-ADGroup -Identity $privGroups[$name] @ad -ErrorAction Stop
    }
    catch { continue }   # Enterprise/Schema Admins only exist in the forest root domain
    # LDAP_MATCHING_RULE_IN_CHAIN: all users in the group, including through nested groups
    $members = @(Get-ADUser -LDAPFilter "(memberOf:1.2.840.113556.1.4.1941:=$($group.DistinguishedName))" @ad)
    foreach ($m in $members) {
        if (-not $privileged.ContainsKey($m.DistinguishedName)) { $privileged[$m.DistinguishedName] = [System.Collections.Generic.List[string]]::new() }
        $privileged[$m.DistinguishedName].Add($group.Name)
    }
}

# --- Users ----------------------------------------------------------------

Write-Host 'Checking user accounts...'
$UF = @{
    PasswordNotRequired   = 0x0020
    ReversibleEncryption  = 0x0080
    PasswordNeverExpires  = 0x10000
    TrustedForDelegation  = 0x80000
    NotDelegated          = 0x100000
    UseDesKeyOnly         = 0x200000
    DontRequirePreauth    = 0x400000
}

$props  = 'userAccountControl', 'servicePrincipalName', 'adminCount', 'PasswordLastSet', 'lastLogonTimestamp',
          'SIDHistory', 'Description', 'DisplayName', 'Enabled', 'objectSid', 'whenCreated'
$filter = if ($IncludeDisabled) { '*' } else { 'Enabled -eq $true' }
$users  = @(Get-ADUser -Filter $filter -Properties $props @search)

$findings = [System.Collections.Generic.List[object]]::new()

function Add-Finding($Account, [string]$Type, [string]$Severity, [string]$Finding, [string]$Detail, [string]$Recommendation) {
    $sam      = $Account.SamAccountName
    $isAccepted = $false
    if ($accepted.ContainsKey($sam)) { $isAccepted = @($accepted[$sam]) -contains $Finding -or @($accepted[$sam]) -contains '*' }
    $groups   = if ($privileged.ContainsKey($Account.DistinguishedName)) { ($privileged[$Account.DistinguishedName] | Sort-Object -Unique) -join ', ' } else { '' }
    $lastLogon = if ($Account.lastLogonTimestamp) { [datetime]::FromFileTime($Account.lastLogonTimestamp) } else { $null }
    $findings.Add([pscustomobject]@{
        Severity        = $Severity
        Status          = if ($isAccepted) { 'Accepted' } else { 'Open' }
        Finding         = $Finding
        ObjectType      = $Type
        SamAccountName  = $sam
        Name            = if ($Account.DisplayName) { $Account.DisplayName } else { $Account.Name }
        Enabled         = [bool]$Account.Enabled
        PrivilegedVia   = $groups
        Detail          = $Detail
        Recommendation  = $Recommendation
        PasswordLastSet = $Account.PasswordLastSet
        LastLogon       = $lastLogon
        OU              = ($Account.DistinguishedName -replace '^.+?(?<!\\),', '')
    })
}

foreach ($u in $users) {
    $uac      = [int]$u.userAccountControl
    $isPriv   = $privileged.ContainsKey($u.DistinguishedName)
    $spns     = @($u.servicePrincipalName | Where-Object { $_ })
    $pwdAge   = if ($u.PasswordLastSet) { [int]($now - $u.PasswordLastSet).TotalDays } else { $null }
    $rid      = ($u.objectSid.Value -split '-')[-1]

    if ($rid -eq '502') { continue }   # krbtgt is checked separately below
    if ($rid -eq '501' -and $u.Enabled) {
        Add-Finding $u 'User' 'High' 'Guest account enabled' 'The built-in Guest account is enabled.' 'Disable the Guest account.'
    }

    if ($uac -band $UF.PasswordNotRequired) {
        Add-Finding $u 'User' 'High' 'Password not required' 'PASSWD_NOTREQD is set, so the account can have an empty password.' `
            'Clear the flag: Set-ADUser <name> -PasswordNotRequired $false, then set a strong password.'
    }
    if ($uac -band $UF.ReversibleEncryption) {
        Add-Finding $u 'User' 'High' 'Reversible password encryption' 'The password is stored in a form that can be decrypted.' `
            'Turn off "Store password using reversible encryption" and reset the password.'
    }
    if ($uac -band $UF.DontRequirePreauth) {
        Add-Finding $u 'User' 'High' 'Kerberos pre-auth not required' 'Anyone can request data encrypted with this password and crack it offline (AS-REP roasting).' `
            'Turn off "Do not require Kerberos preauthentication" unless an application truly needs it.'
    }
    if ($uac -band $UF.TrustedForDelegation) {
        Add-Finding $u 'User' 'High' 'Unconstrained delegation' 'This account can impersonate any user that authenticates to it, anywhere.' `
            'Switch to constrained or resource-based constrained delegation.'
    }
    if ($spns.Count) {
        $spnText = ($spns | Select-Object -First 3) -join '; '
        if ($spns.Count -gt 3) { $spnText += " (+$($spns.Count - 3) more)" }
        if ($isPriv) {
            Add-Finding $u 'User' 'High' 'Kerberoastable privileged account' "Privileged account with SPN(s): $spnText. Its password hash can be requested by any user and cracked offline." `
                'Remove the SPN from the privileged account, or move the service to a gMSA. Otherwise use a 25+ character random password.'
        } else {
            Add-Finding $u 'User' 'Medium' 'Kerberoastable account' "Has SPN(s): $spnText. Any domain user can request a ticket and try to crack the password offline." `
                'Use a group Managed Service Account (gMSA), or a 25+ character random password, rotated regularly.'
        }
    }
    if ($uac -band $UF.UseDesKeyOnly) {
        Add-Finding $u 'User' 'Medium' 'DES-only Kerberos encryption' '"Use only Kerberos DES encryption types" is set; DES is broken.' 'Clear the flag and reset the password.'
    }
    if (@($u.SIDHistory).Count -and $u.SIDHistory) {
        Add-Finding $u 'User' 'Medium' 'SID history present' "$(@($u.SIDHistory).Count) SID(s) in sIDHistory - grants the rights of old accounts." `
            'Remove SID history once migrations are finished.'
    }

    if ($uac -band $UF.PasswordNeverExpires) {
        if ($isPriv) {
            Add-Finding $u 'User' 'Medium' 'Password never expires (privileged)' 'Privileged account whose password never has to change.' `
                'Require password changes, or use a vault/rotation process for this account.'
        } else {
            Add-Finding $u 'User' 'Low' 'Password never expires' 'The password never has to change.' `
                'Turn off "Password never expires", or use a gMSA for service accounts.'
        }
    }

    if ($isPriv) {
        if ($null -ne $pwdAge -and $pwdAge -gt $privMaxDays) {
            Add-Finding $u 'User' 'Medium' 'Stale privileged password' "Password last changed $pwdAge days ago." 'Change the password.'
        }
        if (-not ($uac -band $UF.NotDelegated)) {
            Add-Finding $u 'User' 'Low' 'Privileged account can be delegated' 'Not marked "Account is sensitive and cannot be delegated".' `
                'Tick "Account is sensitive and cannot be delegated", or add the account to Protected Users (test first).'
        }
    }
    elseif ($u.adminCount -eq 1) {
        Add-Finding $u 'User' 'Low' 'Orphaned adminCount' 'adminCount=1 but the account is no longer in a privileged group, so it keeps locked-down permissions.' `
            'Clear adminCount and re-enable permission inheritance on the account.'
    }
}

# --- krbtgt (always disabled, so checked on its own) ----------------------

try {
    $krbtgt = Get-ADUser -Identity "$domainSid-502" -Properties $props @ad -ErrorAction Stop
    $age = if ($krbtgt.PasswordLastSet) { [int]($now - $krbtgt.PasswordLastSet).TotalDays } else { $null }
    if ($null -ne $age -and $age -gt $krbtgtMaxDays) {
        Add-Finding $krbtgt 'User' 'Medium' 'krbtgt password old' "Last changed $age days ago (limit $krbtgtMaxDays)." `
            "Reset the krbtgt password twice, allowing replication in between, using Microsoft's documented procedure."
    }
} catch { Write-Warning "Could not read the krbtgt account: $($_.Exception.Message)" }

# --- Computers with unconstrained delegation ------------------------------

Write-Host 'Checking computers for unconstrained delegation...'
$computers = @(Get-ADComputer -Filter 'TrustedForDelegation -eq $true' -Properties 'TrustedForDelegation', 'PrimaryGroupID', 'OperatingSystem', 'lastLogonTimestamp', 'PasswordLastSet', 'Enabled' @search)
foreach ($c in $computers) {
    if ($c.PrimaryGroupID -in 516, 521) { continue }   # domain controllers / read-only DCs
    if (-not $IncludeDisabled -and -not $c.Enabled) { continue }
    Add-Finding $c 'Computer' 'High' 'Unconstrained delegation' "$($c.OperatingSystem) - can impersonate any user that connects to it." `
        'Switch to constrained or resource-based constrained delegation.'
}

# --- Export ---------------------------------------------------------------

$sevOrder = @{ High = 0; Medium = 1; Low = 2 }
$sorted   = @($findings | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } }, Finding, SamAccountName)
$open     = @($sorted | Where-Object { $_.Status -eq 'Open' })

# One row per flagged account
$accounts = @($sorted | Group-Object ObjectType, SamAccountName | ForEach-Object {
    $first = $_.Group[0]
    $worst = ($_.Group | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } } | Select-Object -First 1).Severity
    [pscustomobject]@{
        HighestSeverity = $worst
        SamAccountName  = $first.SamAccountName
        ObjectType      = $first.ObjectType
        Name            = $first.Name
        Enabled         = $first.Enabled
        PrivilegedVia   = $first.PrivilegedVia
        OpenFindings    = @($_.Group | Where-Object { $_.Status -eq 'Open' }).Count
        Findings        = ($_.Group | ForEach-Object { $_.Finding + $(if ($_.Status -eq 'Accepted') { ' (accepted)' } else { '' }) }) -join '; '
        PasswordLastSet = $first.PasswordLastSet
        LastLogon       = $first.LastLogon
        OU              = $first.OU
    }
} | Sort-Object @{ Expression = { $sevOrder[$_.HighestSeverity] } }, SamAccountName)

# All privileged users, for review even when nothing is flagged
$userByDn = @{}
foreach ($u in $users) { $userByDn[$u.DistinguishedName] = $u }
$privRows = @(foreach ($dn in $privileged.Keys) {
    $u = $userByDn[$dn]
    if (-not $u) { try { $u = Get-ADUser -Identity $dn -Properties $props @ad } catch { continue } }
    [pscustomobject]@{
        SamAccountName  = $u.SamAccountName
        Name            = if ($u.DisplayName) { $u.DisplayName } else { $u.Name }
        Enabled         = [bool]$u.Enabled
        Groups          = ($privileged[$dn] | Sort-Object -Unique) -join ', '
        PasswordLastSet = $u.PasswordLastSet
        PasswordAgeDays = if ($u.PasswordLastSet) { [int]($now - $u.PasswordLastSet).TotalDays } else { $null }
        LastLogon       = if ($u.lastLogonTimestamp) { [datetime]::FromFileTime($u.lastLogonTimestamp) } else { $null }
        OU              = ($dn -replace '^.+?(?<!\\),', '')
    }
}) | Sort-Object SamAccountName

$counts = [ordered]@{}
foreach ($s in 'High', 'Medium', 'Low') { $counts[$s] = @($open | Where-Object { $_.Severity -eq $s }).Count }

$summaryItems = [ordered]@{
    'Run at'              = $now.ToString('yyyy-MM-dd HH:mm:ss')
    'Domain'              = $domain.DNSRoot
    'Search base'         = if ($SearchBase) { $SearchBase } else { '(whole domain)' }
    'Accounts checked'    = "$($users.Count) user(s)" + $(if ($IncludeDisabled) { ' incl. disabled' } else { ' (enabled only)' })
    'Privileged users'    = $privileged.Count
    'Open findings'       = "$($open.Count) ($($counts.High) high, $($counts.Medium) medium, $($counts.Low) low)"
    'Accepted findings'   = @($sorted | Where-Object { $_.Status -eq 'Accepted' }).Count
}
foreach ($f in ($open | Group-Object Finding | Sort-Object Count -Descending)) { $summaryItems["Finding: $($f.Name)"] = $f.Count }
$summaryItems['Run by'] = "$env:USERDOMAIN\$env:USERNAME"

$saved = Save-Workbook -FilePath $reportPath -Sheets @(
    @{ Name = 'Findings'; Rows = $sorted
       Headers = @('Severity', 'Status', 'Finding', 'ObjectType', 'SamAccountName', 'Name', 'Enabled', 'PrivilegedVia', 'Detail', 'Recommendation', 'PasswordLastSet', 'LastLogon', 'OU')
       Widths  = @(10, 10, 34, 11, 22, 26, 9, 30, 60, 60, 18, 18, 50) }
    @{ Name = 'Accounts'; Rows = $accounts
       Headers = @('HighestSeverity', 'SamAccountName', 'ObjectType', 'Name', 'Enabled', 'PrivilegedVia', 'OpenFindings', 'Findings', 'PasswordLastSet', 'LastLogon', 'OU')
       Widths  = @(15, 22, 11, 26, 9, 30, 13, 80, 18, 18, 50) }
    @{ Name = 'Privileged Accounts'; Rows = $privRows
       Headers = @('SamAccountName', 'Name', 'Enabled', 'Groups', 'PasswordLastSet', 'PasswordAgeDays', 'LastLogon', 'OU')
       Widths  = @(22, 26, 9, 45, 18, 16, 18, 50) }
    @{ Name = 'Summary'; Rows = (New-SummaryRows $summaryItems); Headers = @('Item', 'Value'); Widths = @(44, 70) }
)
if (-not $saved) { exit 1 }

Write-Host ""
Write-Host ("Open findings: {0} high, {1} medium, {2} low" -f $counts.High, $counts.Medium, $counts.Low) -ForegroundColor $(if ($counts.High) { 'Red' } elseif ($counts.Medium) { 'Yellow' } else { 'Green' })
foreach ($f in ($open | Group-Object Finding | Sort-Object Count -Descending)) { Write-Host ("  {0,-38} {1}" -f $f.Name, $f.Count) }
Write-Host "Privileged users: $($privileged.Count)"
Write-Host "Report : $reportPath"

if ($SendEmail) {
    $intro = @("Domain $($domain.DNSRoot): $($counts.High) high, $($counts.Medium) medium and $($counts.Low) low open findings across $(@($accounts | Where-Object { $_.OpenFindings }).Count) account(s).",
               'Accepted risks are not listed here. The attached workbook has full details and recommendations.')
    $html = New-HtmlReport -Title 'Risky AD account settings' -IntroLines $intro -Sections @(
        @{ Heading = 'High and medium findings'; Html = ConvertTo-HtmlTable -Rows @($open | Where-Object { $_.Severity -ne 'Low' } | Select-Object -First 100) -Columns @('Severity', 'Finding', 'SamAccountName', 'PrivilegedVia', 'Recommendation') -StatusColumn 'Severity' -EmptyText 'None.' }
        @{ Heading = 'Low findings (count)'; Html = ConvertTo-HtmlTable -Rows @($open | Where-Object { $_.Severity -eq 'Low' } | Group-Object Finding | ForEach-Object { [pscustomobject]@{ Finding = $_.Name; Accounts = $_.Count } }) -Columns @('Finding', 'Accounts') -EmptyText 'None.' }
    )
    $state = if ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Medium' } else { 'Low' }
    if (Send-ToolkitMail -Config $config -Subject "[$state] Risky AD accounts - $($open.Count) open finding(s)" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
}
