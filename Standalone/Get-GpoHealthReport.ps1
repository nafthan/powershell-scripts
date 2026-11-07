<#
.SYNOPSIS
    Checks Group Policy for conflicting settings, unlinked, empty or broken GPOs, and risky link settings.
.DESCRIPTION
    Uses each OU's effective GPO list (including inherited, enforced and blocked links) to find settings
    configured differently by more than one GPO, and shows which GPO wins. Also flags unlinked, empty and
    disabled GPOs, disabled links, block inheritance and GPOs that no computer or user can apply. Read-only.
.EXAMPLE
    .\Get-GpoHealthReport.ps1
.EXAMPLE
    .\Get-GpoHealthReport.ps1 -SearchBase 'OU=Workstations,DC=corp,DC=example,DC=com' -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: GroupPolicy and ActiveDirectory modules (RSAT)
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$Server,

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
if (-not $PSBoundParameters.ContainsKey('Server')) { $Server = Get-Setting $config 'ActiveDirectory' 'Server' '' }

# --- Reading GPO XML ----------------------------------------------------------------
# GPO reports use XML namespaces, so nodes are matched by local name.

function Select-Local($Node, [string]$Path) {
    $xpath = ($Path -split '/' | ForEach-Object { if ($_ -eq '') { '' } elseif ($_ -eq '..') { '..' } else { "*[local-name()='$_']" } }) -join '/'
    if ($Node) { return @($Node.SelectNodes($xpath)) }
    return @()
}
function Get-LocalText($Node, [string]$Name) {
    $n = @(Select-Local $Node $Name)
    if ($n.Count) { return $n[0].InnerText }
    return ''
}

function Get-GpoSettings($GpoNode) {
    <# Returns @{ Settings = list of Scope/Category/Name/Value; ScopesWithData = @('Computer','User') } #>
    $settings = [System.Collections.Generic.List[object]]::new()
    $withData = @()
    foreach ($scope in 'Computer', 'User') {
        $scopeNode = @(Select-Local $GpoNode $scope)[0]
        if (-not $scopeNode) { continue }
        $extensions = @(Select-Local $scopeNode 'ExtensionData/Extension')
        if ($extensions.Count) { $withData += $scope }
        if ((Get-LocalText $scopeNode 'Enabled') -eq 'false') { continue }   # this half of the GPO is switched off
        $add = { param($Category, $Name, $Value) if ($Name) { $settings.Add([pscustomobject]@{ Scope = $scope; Category = $Category; Name = "$Name"; Value = "$Value" }) } }

        foreach ($ext in $extensions) {
            foreach ($p in (Select-Local $ext 'Policy')) {
                $vals = @($p.ChildNodes | Where-Object { $_.LocalName -in 'EditText', 'Numeric', 'DropDownList', 'CheckBox', 'ListBox' } |
                          ForEach-Object { "$(Get-LocalText $_ 'Name')=$(Get-LocalText $_ 'Value')$(Get-LocalText $_ 'State')" })
                & $add 'Administrative Templates' "$(Get-LocalText $p 'Category')\$(Get-LocalText $p 'Name')" ((@(Get-LocalText $p 'State') + $vals) -join '; ')
            }
            foreach ($s in (Select-Local $ext 'SecurityOptions')) {
                $name = Get-LocalText $s 'KeyName'; if (-not $name) { $name = Get-LocalText $s 'SystemAccessPolicyName' }
                & $add 'Security Options' $name ((@('SettingNumber', 'SettingString', 'SettingBoolean') | ForEach-Object { Get-LocalText $s $_ } | Where-Object { $_ }) -join '; ')
            }
            foreach ($a in (Select-Local $ext 'Account')) {
                & $add "Account Policy ($(Get-LocalText $a 'Type'))" (Get-LocalText $a 'Name') ((@('SettingNumber', 'SettingBoolean', 'SettingString') | ForEach-Object { Get-LocalText $a $_ } | Where-Object { $_ }) -join '; ')
            }
            foreach ($a in (Select-Local $ext 'Audit')) {
                & $add 'Audit Policy' (Get-LocalText $a 'Name') "Success=$(Get-LocalText $a 'SuccessAttempts'); Failure=$(Get-LocalText $a 'FailureAttempts')"
            }
            foreach ($a in (Select-Local $ext 'AuditSetting')) { & $add 'Advanced Audit Policy' (Get-LocalText $a 'SubcategoryName') (Get-LocalText $a 'SettingValue') }
            foreach ($u in (Select-Local $ext 'UserRightsAssignment')) {
                & $add 'User Rights' (Get-LocalText $u 'Name') ((@(Select-Local $u 'Member') | ForEach-Object { Get-LocalText $_ 'Name' } | Sort-Object) -join ', ')
            }
            foreach ($r in (Select-Local $ext 'RegistrySettings/Registry')) {
                $props = @(Select-Local $r 'Properties')[0]
                if ($props) { & $add 'Registry Preference' "$($props.hive)\$($props.key)\$($props.name)" "$($props.action) $($props.type)=$($props.value)" }
            }
        }
    }
    return @{ Settings = $settings; ScopesWithData = $withData }
}

function Get-GpoPermissionIssue($GpoNode) {
    $perms = @(Select-Local $GpoNode 'SecurityDescriptor/Permissions/TrusteePermissions')
    if (-not $perms.Count) { return '' }
    $allow   = @($perms | Where-Object { (Get-LocalText $_ 'Type/PermissionType') -ne 'Deny' })
    $apply   = @($allow | Where-Object { (Get-LocalText $_ 'Standard/GPOGroupedAccessEnum') -eq 'Apply Group Policy' })
    $readers = @($allow | ForEach-Object { Get-LocalText $_ 'Trustee/SID' })
    if (-not $apply.Count) { return 'No one has Apply permission - the GPO applies to nobody' }
    $canRead = $readers | Where-Object { $_ -eq 'S-1-5-11' -or $_ -match '-515$' }
    if (-not $canRead) { return 'Neither Authenticated Users nor Domain Computers can read it - user settings may not apply (MS16-072)' }
    return ''
}

# --- Main ------------------------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    foreach ($m in 'GroupPolicy', 'ActiveDirectory') {
        try { Import-Module $m -ErrorAction Stop } catch { throw "The $m module is not installed. Install RSAT: Group Policy Management and Active Directory tools." }
    }
    $ad = @{}; $gp = @{}
    if ($Server) { $ad.Server = $Server; $gp.Server = $Server }
    $domain = Get-ADDomain @ad
    $gp.Domain = $domain.DNSRoot
    if (-not $SearchBase) { $SearchBase = $domain.DistinguishedName }
    $now        = Get-Date
    $reportPath = New-ReportPath -Config $config -Prefix 'GpoHealth' -OutputFile $OutputFile
    $findings   = [System.Collections.Generic.List[object]]::new()
    function Add-Finding([string]$Severity, [string]$Finding, [string]$Gpo, [string]$Where, [string]$Detail) {
        $findings.Add([pscustomobject]@{ Severity = $Severity; Finding = $Finding; GPO = $Gpo; Where = $Where; Detail = $Detail })
    }

    Write-Host "Domain      : $($domain.DNSRoot)"
    Write-Host "Search base : $SearchBase"

    # 1. GPOs and their settings - one XML report for all GPOs
    Write-Host 'Reading GPOs...'
    $gpos = @(Get-GPO -All @gp)
    $gpoById = @{}; foreach ($g in $gpos) { $gpoById[$g.Id.ToString().ToLower()] = $g }
    $info = @{}
    $xmlNodes = @()
    try {
        [xml]$all = Get-GPOReport -All -ReportType Xml @gp -ErrorAction Stop
        $xmlNodes = @(Select-Local $all '/report/GPO')
        if (-not $xmlNodes.Count) { $xmlNodes = @(Select-Local $all '//GPO') }
    }
    catch {
        Write-Warning "Reading all GPO reports at once failed ($($_.Exception.Message)) - reading them one by one."
        $n = 0
        foreach ($g in $gpos) {
            $n++; Write-Progress -Activity 'Reading GPO reports' -Status "$n of $($gpos.Count)" -PercentComplete ($n / [math]::Max(1, $gpos.Count) * 100)
            try { [xml]$one = Get-GPOReport -Guid $g.Id -ReportType Xml @gp -ErrorAction Stop; $xmlNodes += @(Select-Local $one '/GPO') }
            catch { Add-Finding 'Low' 'GPO report could not be read' $g.DisplayName '' $_.Exception.Message }
        }
        Write-Progress -Activity 'Reading GPO reports' -Completed
    }
    foreach ($node in $xmlNodes) {
        $id = (Get-LocalText $node 'Identifier/Identifier').Trim('{', '}').ToLower()
        if (-not $id) { continue }
        $s = Get-GpoSettings $node
        $info[$id] = [pscustomobject]@{ Settings = $s.Settings; ScopesWithData = $s.ScopesWithData; PermissionIssue = (Get-GpoPermissionIssue $node) }
    }

    # 2. Effective GPO links for the search base and every OU under it
    Write-Host 'Reading OUs and links...'
    $targets = @($SearchBase) + @(Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase @ad | Select-Object -ExpandProperty DistinguishedName | Sort-Object { ($_ -split '(?<!\\),').Count }, { $_ })
    $linkRows = [System.Collections.Generic.List[object]]::new()
    $effective = @{}
    $linkedIds = @{}
    $n = 0
    foreach ($t in $targets) {
        $n++; Write-Progress -Activity 'Reading GPO inheritance' -Status "$n of $($targets.Count)" -PercentComplete ($n / [math]::Max(1, $targets.Count) * 100)
        try { $inh = Get-GPInheritance -Target $t @gp -ErrorAction Stop }
        catch { Add-Finding 'Low' 'Could not read GPO inheritance' '' $t $_.Exception.Message; continue }
        if ($inh.GpoInheritanceBlocked) { Add-Finding 'Low' 'Block inheritance is set - GPOs from parent OUs do not apply here unless enforced' '' $t '' }
        foreach ($l in @($inh.GpoLinks)) {
            $id = "$($l.GpoId)".ToLower()
            $linkedIds[$id] = $true
            $linkRows.Add([pscustomobject]@{ OU = $t; GPO = $l.DisplayName; LinkOrder = $l.Order; Enabled = [bool]$l.Enabled; Enforced = [bool]$l.Enforced })
            if (-not $l.Enabled) { Add-Finding 'Low' 'Link is disabled - the GPO does not apply here' $l.DisplayName $t '' }
        }
        # InheritedGpoLinks is the effective list in precedence order (1 = wins)
        $effective[$t] = @($inh.InheritedGpoLinks | Where-Object { $_.Enabled } | Sort-Object Order)
    }
    Write-Progress -Activity 'Reading GPO inheritance' -Completed
    # GPOs linked above the search base (e.g. at the domain root) also count as linked
    foreach ($t in $effective.Keys) { foreach ($l in $effective[$t]) { $linkedIds["$($l.GpoId)".ToLower()] = $true } }

    # 3. Conflicts - the same setting with different values from more than one effective GPO
    Write-Host 'Looking for conflicts...'
    $conflicts = @{}
    foreach ($t in $effective.Keys) {
        $bySetting = @{}
        foreach ($l in $effective[$t]) {
            $i = $info["$($l.GpoId)".ToLower()]
            if (-not $i) { continue }
            foreach ($s in $i.Settings) {
                $key = "$($s.Scope)|$($s.Category)|$($s.Name)"
                if (-not $bySetting.ContainsKey($key)) { $bySetting[$key] = [System.Collections.Generic.List[object]]::new() }
                $bySetting[$key].Add([pscustomobject]@{ Gpo = $l.DisplayName; Order = $l.Order; Value = $s.Value; Setting = $s })
            }
        }
        foreach ($key in $bySetting.Keys) {
            $entries = @($bySetting[$key] | Sort-Object Order)
            if (@($entries.Value | Select-Object -Unique).Count -lt 2) { continue }
            $sig = "$key|" + (($entries | ForEach-Object { "$($_.Gpo)=$($_.Value)" }) -join '|')
            if (-not $conflicts.ContainsKey($sig)) {
                $conflicts[$sig] = [pscustomobject]@{
                    Scope = $entries[0].Setting.Scope; Category = $entries[0].Setting.Category; Setting = $entries[0].Setting.Name
                    WinningGPO = $entries[0].Gpo; WinningValue = $entries[0].Value
                    OverriddenBy = (($entries | Select-Object -Skip 1 | ForEach-Object { "$($_.Gpo) [$($_.Value)]" }) -join '; ')
                    OUs = [System.Collections.Generic.List[string]]::new()
                }
            }
            $conflicts[$sig].OUs.Add($t)
        }
    }
    $conflictRows = @($conflicts.Values | ForEach-Object {
        [pscustomobject]@{
            Scope = $_.Scope; Category = $_.Category; Setting = $_.Setting; WinningGPO = $_.WinningGPO; WinningValue = $_.WinningValue
            LosingGPOs = $_.OverriddenBy; AffectedOUs = $_.OUs.Count; ExampleOU = $_.OUs[0]
        }
    } | Sort-Object Category, Setting)
    foreach ($c in $conflictRows) { Add-Finding 'Medium' "Conflicting setting: $($c.Setting)" $c.WinningGPO $c.ExampleOU "Wins with [$($c.WinningValue)] over $($c.LosingGPOs) - $($c.AffectedOUs) OU(s)" }

    # 4. GPO-level checks
    $gpoRows = foreach ($g in $gpos) {
        $id  = $g.Id.ToString().ToLower()
        $i   = $info[$id]
        $linked = $linkedIds.ContainsKey($id)
        $status = "$($g.GpoStatus)"
        $count  = if ($i) { $i.Settings.Count } else { $null }
        $issues = @()
        if (-not $linked) { $issues += 'Not linked within the search base' }
        if ($i -and -not $i.ScopesWithData.Count) { $issues += 'Empty - no settings' }
        if ($status -eq 'AllSettingsDisabled') { $issues += 'All settings disabled' }
        elseif ($status -eq 'ComputerSettingsDisabled' -and $i -and $i.ScopesWithData -contains 'Computer') { $issues += 'Computer settings exist but are disabled' }
        elseif ($status -eq 'UserSettingsDisabled' -and $i -and $i.ScopesWithData -contains 'User') { $issues += 'User settings exist but are disabled' }
        if ($i -and $i.PermissionIssue) { $issues += $i.PermissionIssue }

        if ($linked -and $i -and $i.PermissionIssue) { Add-Finding 'High' $i.PermissionIssue $g.DisplayName '' 'Check the Delegation tab in Group Policy Management' }
        if ($linked -and $status -eq 'AllSettingsDisabled') { Add-Finding 'Medium' 'Linked GPO has all settings disabled - it does nothing' $g.DisplayName '' '' }
        if ($issues -match 'exist but are disabled') { Add-Finding 'Low' ($issues -match 'exist but are disabled')[0] $g.DisplayName '' 'Often left behind after troubleshooting' }
        if (-not $linked) { Add-Finding 'Low' 'GPO is not linked within the search base' $g.DisplayName '' "Last changed $($g.ModificationTime.ToString('yyyy-MM-dd'))" }
        elseif ($i -and -not $i.ScopesWithData.Count) { Add-Finding 'Low' 'Linked GPO is empty' $g.DisplayName '' '' }

        [pscustomobject]@{
            GPO = $g.DisplayName; Status = $status; Linked = $linked; Settings = $count; WmiFilter = if ($g.WmiFilter) { $g.WmiFilter.Name } else { '' }
            Created = $g.CreationTime; Modified = $g.ModificationTime; Issues = ($issues -join '; '); Id = $g.Id
        }
    }

    # --- Export --------------------------------------------------------------------
    $order = @{ High = 0; Medium = 1; Low = 2 }
    $sortedFindings = @($findings | Sort-Object @{ Expression = { $order[$_.Severity] } }, Finding, GPO)
    $counts = [ordered]@{}; foreach ($s in 'High', 'Medium', 'Low') { $counts[$s] = @($sortedFindings | Where-Object Severity -eq $s).Count }
    $summary = [ordered]@{
        'Domain'                 = $domain.DNSRoot
        'Search base'            = $SearchBase
        'Run at'                 = $now.ToString('yyyy-MM-dd HH:mm:ss')
        'GPOs'                   = $gpos.Count
        'OUs checked'            = $targets.Count
        'Conflicting settings'   = $conflictRows.Count
        'Unlinked GPOs'          = @($gpoRows | Where-Object { -not $_.Linked }).Count
        'Findings'               = "$($counts.High) high, $($counts.Medium) medium, $($counts.Low) low"
        'Settings compared'      = 'Administrative Templates, Security Options, Account/Audit policy, User Rights and Registry preferences. Other preference types and scripts are not compared.'
        'Note'                   = 'Site-linked GPOs and WMI/security filtering are not taken into account when working out which GPO wins.'
    }

    $saved = Save-Workbook -FilePath $reportPath -Sheets @(
        @{ Name = 'Findings'; Rows = $sortedFindings; Headers = @('Severity', 'Finding', 'GPO', 'Where', 'Detail'); Widths = @(10, 60, 36, 50, 80) }
        @{ Name = 'Conflicts'; Rows = $conflictRows; Headers = @('Scope', 'Category', 'Setting', 'WinningGPO', 'WinningValue', 'LosingGPOs', 'AffectedOUs', 'ExampleOU'); Widths = @(10, 26, 60, 30, 30, 60, 12, 50) }
        @{ Name = 'GPOs'; Rows = @($gpoRows | Sort-Object GPO); Headers = @('GPO', 'Status', 'Linked', 'Settings', 'WmiFilter', 'Created', 'Modified', 'Issues', 'Id'); Widths = @(36, 24, 8, 10, 20, 18, 18, 70, 38) }
        @{ Name = 'Links'; Rows = @($linkRows | Sort-Object OU, LinkOrder); Headers = @('OU', 'GPO', 'LinkOrder', 'Enabled', 'Enforced'); Widths = @(60, 36, 10, 9, 10) }
        @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(22, 110) }
    )
    if (-not $saved) { exit 1 }

    Write-Host ""
    Write-Host ("GPOs: {0}. Conflicts: {1}. Findings: {2}" -f $gpos.Count, $conflictRows.Count, $summary.Findings) -ForegroundColor $(if ($counts.High) { 'Red' } elseif ($counts.Medium) { 'Yellow' } else { 'Green' })
    $sortedFindings | Where-Object Severity -eq 'High' | Select-Object -First 10 | ForEach-Object { Write-Host "  HIGH  $($_.GPO): $($_.Finding)" -ForegroundColor Red }
    Write-Host "Report: $reportPath"

    if ($SendEmail) {
        $html = New-HtmlReport -Title 'Group Policy health report' -IntroLines @("$($domain.DNSRoot): $($gpos.Count) GPOs, $($conflictRows.Count) conflicting setting(s). Findings: $($summary.Findings).") -Sections @(
            @{ Heading = 'High and medium findings'; Html = ConvertTo-HtmlTable -Rows @($sortedFindings | Where-Object Severity -ne 'Low' | Select-Object -First 100) -Columns @('Severity', 'Finding', 'GPO', 'Where', 'Detail') -StatusColumn 'Severity' -EmptyText 'None.' }
        )
        if (Send-ToolkitMail -Config $config -Subject "Group Policy health: $($summary.Findings)" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
    }
}
