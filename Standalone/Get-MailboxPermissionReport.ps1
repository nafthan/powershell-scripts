<#
.SYNOPSIS
    Reports who has access to which Exchange Online mailboxes and calendars.
.DESCRIPTION
    Lists Full Access, Send As and Send on Behalf permissions, and optionally calendar permissions,
    for user, shared, room and equipment mailboxes. Flags permissions held by deleted users and
    calendars shared with the whole organisation. Read-only.
.EXAMPLE
    .\Get-MailboxPermissionReport.ps1
.EXAMPLE
    .\Get-MailboxPermissionReport.ps1 -Trustee jane.doe@contoso.com -IncludeCalendars
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: ExchangeOnlineManagement module, Exchange Administrator (or Global Reader)
#>
[CmdletBinding()]
param(
    [string[]]$Mailbox,
    [string]$Trustee,
    [switch]$IncludeCalendars,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [string]$OutputFile,

    # Report look: a theme file (see themes\ in the repo) and the labels on HTML reports
    [string]$ThemeFile
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

function Connect-ToolkitExchange {
    <#
        Connects to Exchange Online, or reuses an existing connection. Sign-in settings come from the
        calling script's parameters (-Bound $PSBoundParameters) first, then built-in defaults.
        Returns $true if this call made the connection.
    #>
    param([hashtable]$Config, $Bound, [string]$RoleHint = 'Exchange Administrator (or Global Reader)')
    $boundKeys = if ($Bound) { @($Bound.Keys) } else { @() }
    $setting = { param($Name, $Default) if ($boundKeys -contains $Name) { $Bound[$Name] } else { Get-Setting $Config 'M365' $Name $Default } }
    $TenantId              = & $setting 'TenantId' ''
    $ClientId              = & $setting 'ClientId' ''
    $CertificateThumbprint = & $setting 'CertificateThumbprint' ''
    $UseDeviceCode         = [bool](& $setting 'UseDeviceCode' $false)

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw 'The Exchange Online PowerShell module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    $existing = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' })
    if ($existing.Count) {
        Write-Host "Using existing Exchange Online connection ($($existing[0].UserPrincipalName)$($existing[0].AppId))."
        return $false
    }
    $connect = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($ClientId -and $CertificateThumbprint) {
        if ($TenantId -notlike '*.*') { throw 'App-only sign-in to Exchange needs the tenant domain (e.g. contoso.onmicrosoft.com). Use -TenantId.' }
        $connect.AppId = $ClientId; $connect.CertificateThumbprint = $CertificateThumbprint; $connect.Organization = $TenantId
        Write-Host 'Connecting to Exchange Online (app-only, certificate)...'
    }
    else {
        Write-Host "Connecting to Exchange Online - sign in with an account that has $RoleHint..."
        if ($UseDeviceCode) {
            $connect.Device = $true
            Write-Host 'Open the link below in any browser and enter the code shown.' -ForegroundColor Cyan
        }
        else {
            Write-Host 'If the sign-in window goes blank or hangs, press Ctrl+C and run again with -UseDeviceCode.' -ForegroundColor Cyan
        }
    }
    Connect-ExchangeOnline @connect | Out-Null
    return $true
}

#endregion
$config  = Get-ToolkitConfig
$Mailbox = Split-ListValue $Mailbox

# Calendar access levels that let others see more than free/busy
$broadCalendarRights = 'Reviewer|Author|Editor|Owner|PublishingEditor|PublishingAuthor|NonEditingAuthor|Contributor'

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitExchange -Config $config -Bound $PSBoundParameters
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'MailboxPermissions' -OutputFile $OutputFile
        $rows       = [System.Collections.Generic.List[object]]::new()
        $errors     = [System.Collections.Generic.List[object]]::new()

        Write-Host 'Getting mailboxes...'
        $types = 'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'
        if ($Mailbox.Count) { $mailboxes = @($Mailbox | ForEach-Object { Get-EXOMailbox -Identity $_ -Properties GrantSendOnBehalfTo -ErrorAction Stop }) }
        else { $mailboxes = @(Get-EXOMailbox -RecipientTypeDetails $types -ResultSize Unlimited -Properties GrantSendOnBehalfTo) }
        Write-Host "  $($mailboxes.Count) mailbox(es)"

        function Add-Row($Mbx, [string]$Permission, [string]$Who, [string]$Rights, [string]$Note) {
            if ($Trustee -and $Who -notlike "*$Trustee*") { return }
            $flag = ''
            if ($Who -match '^S-1-5-21-') { $flag = 'Held by a deleted user (orphaned SID) - remove it' }
            elseif ($Note) { $flag = $Note }
            $rows.Add([pscustomobject]@{
                Mailbox = "$($Mbx.PrimarySmtpAddress)"; DisplayName = $Mbx.DisplayName; MailboxType = "$($Mbx.RecipientTypeDetails)"
                Permission = $Permission; Trustee = $Who; AccessRights = $Rights; Flag = $flag
            })
        }

        $i = 0
        foreach ($m in $mailboxes) {
            $i++
            Write-Progress -Activity 'Reading mailbox permissions' -Status "$i of $($mailboxes.Count): $($m.PrimarySmtpAddress)" -PercentComplete ($i / [math]::Max(1, $mailboxes.Count) * 100)
            $id = "$($m.PrimarySmtpAddress)"

            try {
                foreach ($p in @(Get-EXOMailboxPermission -Identity $id -ErrorAction Stop)) {
                    if ($p.IsInherited -or "$($p.User)" -like 'NT AUTHORITY\*' -or "$($p.User)" -match '^S-1-5-(18|19|20|32)') { continue }
                    if (@($p.AccessRights) -notcontains 'FullAccess') { continue }
                    Add-Row $m 'Full Access' "$($p.User)" (@($p.AccessRights) -join ', ') $(if ($p.Deny) { 'Deny entry' })
                }
            }
            catch { $errors.Add([pscustomobject]@{ Mailbox = $id; Check = 'Full Access'; Error = $_.Exception.Message }) }

            try {
                foreach ($p in @(Get-EXORecipientPermission -Identity $id -ErrorAction Stop)) {
                    if ($p.IsInherited -or "$($p.Trustee)" -like 'NT AUTHORITY\*') { continue }
                    if (@($p.AccessRights) -notcontains 'SendAs') { continue }
                    Add-Row $m 'Send As' "$($p.Trustee)" 'SendAs' ''
                }
            }
            catch { $errors.Add([pscustomobject]@{ Mailbox = $id; Check = 'Send As'; Error = $_.Exception.Message }) }

            foreach ($d in @($m.GrantSendOnBehalfTo | Where-Object { $_ })) { Add-Row $m 'Send on Behalf' "$d" 'SendOnBehalf' '' }

            if ($IncludeCalendars) {
                try {
                    $calendar = "${id}:\Calendar"
                    try { $perms = @(Get-EXOMailboxFolderPermission -Identity $calendar -ErrorAction Stop) }
                    catch {
                        # Calendar folder has a different (localised) name
                        $folder = @(Get-EXOMailboxFolderStatistics -Identity $id -FolderScope Calendar -ErrorAction Stop | Where-Object { $_.FolderType -eq 'Calendar' })[0]
                        $calendar = "${id}:\$($folder.Name)"
                        $perms = @(Get-EXOMailboxFolderPermission -Identity $calendar -ErrorAction Stop)
                    }
                    foreach ($p in $perms) {
                        $who    = "$($p.User)"
                        $rights = (@($p.AccessRights) -join ', ')
                        if ($rights -eq 'None') { continue }
                        $note = ''
                        if ($who -eq 'Default' -and $rights -match $broadCalendarRights) { $note = 'Everyone in the organisation can see calendar details' }
                        elseif ($who -eq 'Default') { continue }   # Default AvailabilityOnly/LimitedDetails is normal
                        elseif ($who -eq 'Anonymous') { $note = 'Calendar published to anyone (anonymous access)' }
                        Add-Row $m 'Calendar' $who $rights $note
                    }
                }
                catch { $errors.Add([pscustomobject]@{ Mailbox = $id; Check = 'Calendar'; Error = $_.Exception.Message }) }
            }
        }
        Write-Progress -Activity 'Reading mailbox permissions' -Completed

        # --- Export ----------------------------------------------------------

        $sorted = @($rows | Sort-Object @{ Expression = { -not $_.Flag } }, Mailbox, Permission, Trustee)
        $byTrustee = @($rows | Where-Object { $_.Trustee -notin 'Default', 'Anonymous' } | Group-Object Trustee | ForEach-Object {
            $g = $_.Group
            [pscustomobject]@{
                Trustee      = $_.Name
                FullAccess   = @($g | Where-Object Permission -eq 'Full Access').Count
                SendAs       = @($g | Where-Object Permission -eq 'Send As').Count
                SendOnBehalf = @($g | Where-Object Permission -eq 'Send on Behalf').Count
                Calendars    = @($g | Where-Object Permission -eq 'Calendar').Count
                Mailboxes    = (@($g.Mailbox | Select-Object -Unique) -join ', ')
            }
        } | Sort-Object @{ Expression = { $_.FullAccess + $_.SendAs + $_.SendOnBehalf + $_.Calendars }; Descending = $true })

        $summary = [ordered]@{
            'Run at'                 = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Mailboxes checked'      = $mailboxes.Count
            'Filtered to trustee'    = if ($Trustee) { $Trustee } else { '(all)' }
            'Full Access'            = @($rows | Where-Object Permission -eq 'Full Access').Count
            'Send As'                = @($rows | Where-Object Permission -eq 'Send As').Count
            'Send on Behalf'         = @($rows | Where-Object Permission -eq 'Send on Behalf').Count
            'Calendar permissions'   = if ($IncludeCalendars) { @($rows | Where-Object Permission -eq 'Calendar').Count } else { 'Not checked (use -IncludeCalendars)' }
            'Flagged'                = @($rows | Where-Object Flag).Count
            'Note'                   = 'Inherited and system permissions are not listed. Full Access granted through groups shows the group name.'
        }

        $saved = Save-Workbook -FilePath $reportPath -Sheets @(
            @{ Name = 'Permissions'; Rows = $sorted; Headers = @('Mailbox', 'DisplayName', 'MailboxType', 'Permission', 'Trustee', 'AccessRights', 'Flag'); Widths = @(36, 26, 16, 15, 36, 24, 50) }
            @{ Name = 'By Trustee'; Rows = $byTrustee; Headers = @('Trustee', 'FullAccess', 'SendAs', 'SendOnBehalf', 'Calendars', 'Mailboxes'); Widths = @(36, 11, 9, 13, 10, 100) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Mailbox', 'Check', 'Error'); Widths = @(36, 14, 100) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(24, 90) }
        )
        if (-not $saved) { exit 1 }

        Write-Host ""
        Write-Host ("Permissions: {0} Full Access, {1} Send As, {2} Send on Behalf{3}" -f $summary['Full Access'], $summary['Send As'], $summary['Send on Behalf'], $(if ($IncludeCalendars) { ", $($summary['Calendar permissions']) calendar" }))
        if ($summary.Flagged) { Write-Host ("Flagged    : {0} (orphaned SIDs or calendars shared widely)" -f $summary.Flagged) -ForegroundColor Yellow }
        if ($errors.Count) { Write-Host ("Errors     : {0} - see the 'Errors' sheet" -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report     : $reportPath"
    }
    finally {
        if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    }
}
