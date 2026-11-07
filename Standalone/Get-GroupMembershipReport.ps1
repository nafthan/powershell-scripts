<#
.SYNOPSIS
    Exports every Microsoft 365 group, Team, distribution list and security group with its members and owners.
.DESCRIPTION
    One row per member, plus a summary row per group that flags groups with no owner, guests and
    empty groups. Optionally includes private and shared Teams channel members. Read-only.
.EXAMPLE
    .\Get-GroupMembershipReport.ps1
.EXAMPLE
    .\Get-GroupMembershipReport.ps1 -GroupType Team -IncludeTeamsChannels
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: Microsoft.Graph.Authentication module, Global Reader
#>
[CmdletBinding()]
param(
    [string[]]$GroupType,
    [switch]$IncludeTeamsChannels,

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
$config    = Get-ToolkitConfig
$GroupType = Split-ListValue $GroupType
$validTypes = 'Team', 'Microsoft 365 group', 'Distribution list', 'Mail-enabled security', 'Security'
foreach ($t in $GroupType) { if ($validTypes -notcontains $t) { throw "Unknown -GroupType '$t'. Valid: $($validTypes -join ', ')." } }

$scopes = @('Group.Read.All', 'User.Read.All')
if ($IncludeTeamsChannels) { $scopes += 'Channel.ReadBasic.All', 'ChannelMember.Read.All' }

function Get-GroupKind($G) {
    if (@($G.resourceProvisioningOptions) -contains 'Team') { return 'Team' }
    if (@($G.groupTypes) -contains 'Unified') { return 'Microsoft 365 group' }
    if ($G.mailEnabled -and -not $G.securityEnabled) { return 'Distribution list' }
    if ($G.mailEnabled -and $G.securityEnabled) { return 'Mail-enabled security' }
    return 'Security'
}

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'Global Reader'
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'GroupMembership' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        $tenantName = Get-GraphTenantName
        Write-Host "Tenant : $tenantName"

        Write-Host 'Getting groups...'
        $select = 'id,displayName,mail,description,createdDateTime,groupTypes,mailEnabled,securityEnabled,resourceProvisioningOptions,membershipRule,visibility,onPremisesSyncEnabled'
        $groups = @(Get-GraphAll "v1.0/groups?`$select=$select&`$top=999" | ForEach-Object {
            $_ | Add-Member -NotePropertyName Kind -NotePropertyValue (Get-GroupKind $_) -PassThru
        })
        if ($GroupType.Count) { $groups = @($groups | Where-Object { $GroupType -contains $_.Kind }) }
        Write-Host "  $($groups.Count) group(s)"

        $memberRows  = [System.Collections.Generic.List[object]]::new()
        $groupRows   = [System.Collections.Generic.List[object]]::new()
        $channelRows = [System.Collections.Generic.List[object]]::new()
        $i = 0
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity 'Reading members' -Status "$i of $($groups.Count): $($g.displayName)" -PercentComplete ($i / [math]::Max(1, $groups.Count) * 100)
            $members = @(); $owners = @(); $membersRead = $true
            try {
                # One request per page of members - no per-member lookups
                $members = @(Get-GraphAll "v1.0/groups/$($g.id)/members?`$select=id,displayName,userPrincipalName,mail,userType,accountEnabled&`$top=999")
            }
            catch { $membersRead = $false; $errors.Add([pscustomobject]@{ Group = $g.displayName; Check = 'Members'; Error = $_.Exception.Message }) }
            if ($g.Kind -in 'Team', 'Microsoft 365 group', 'Security', 'Mail-enabled security') {
                try { $owners = @(Get-GraphAll "v1.0/groups/$($g.id)/owners?`$select=displayName,userPrincipalName") }
                catch { $errors.Add([pscustomobject]@{ Group = $g.displayName; Check = 'Owners'; Error = $_.Exception.Message }) }
            }

            foreach ($m in $members) {
                $type = "$($m.'@odata.type')" -replace '^#microsoft\.graph\.', ''
                $memberRows.Add([pscustomobject]@{
                    GroupType   = $g.Kind
                    Group       = $g.displayName
                    GroupEmail  = $g.mail
                    Member      = $m.displayName
                    MemberEmail = if ($m.userPrincipalName) { $m.userPrincipalName } else { $m.mail }
                    MemberType  = switch ($type) { 'user' { if ($m.userType -eq 'Guest') { 'Guest' } else { 'User' } } 'group' { 'Nested group' } 'orgContact' { 'Contact' } 'device' { 'Device' } 'servicePrincipal' { 'App' } default { $type } }
                    Enabled     = if ($type -eq 'user') { [bool]$m.accountEnabled } else { '' }
                })
            }

            $guests  = @($members | Where-Object { $_.userType -eq 'Guest' }).Count
            $issues  = @()
            if ($g.Kind -in 'Team', 'Microsoft 365 group' -and -not $owners.Count) { $issues += 'No owner - nobody manages membership or renews it' }
            elseif ($g.Kind -in 'Team', 'Microsoft 365 group' -and $owners.Count -eq 1) { $issues += 'Only one owner' }
            if (-not $membersRead) { $issues += 'Members could not be read' }
            elseif (-not $members.Count) { $issues += 'No members' }
            if ($guests) { $issues += "$guests guest(s)" }
            $disabled = @($members | Where-Object { "$($_.'@odata.type')" -like '*user' -and $_.accountEnabled -eq $false }).Count
            if ($disabled) { $issues += "$disabled disabled account(s) still members" }

            $groupRows.Add([pscustomobject]@{
                GroupType  = $g.Kind
                Group      = $g.displayName
                Email      = $g.mail
                Members    = if ($membersRead) { $members.Count } else { 'Error' }
                Guests     = $guests
                Owners     = (@($owners | ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }) -join ', ')
                Dynamic    = if (@($g.groupTypes) -contains 'DynamicMembership') { 'Yes' } else { 'No' }
                Visibility = $g.visibility
                Source     = if ($g.onPremisesSyncEnabled) { 'Synced (AD)' } else { 'Cloud' }
                Created    = ConvertTo-LocalTime $g.createdDateTime
                Issues     = ($issues -join '; ')
                Description = $g.description
            })

            if ($IncludeTeamsChannels -and $g.Kind -eq 'Team') {
                try {
                    foreach ($c in @(Get-GraphAll "v1.0/teams/$($g.id)/channels?`$select=id,displayName,membershipType" | Where-Object { $_.membershipType -ne 'standard' })) {
                        foreach ($cm in @(Get-GraphAll "v1.0/teams/$($g.id)/channels/$($c.id)/members")) {
                            $channelRows.Add([pscustomobject]@{
                                Team = $g.displayName; Channel = $c.displayName; ChannelType = (Get-Culture).TextInfo.ToTitleCase("$($c.membershipType)")
                                Member = $cm.displayName; MemberEmail = $cm.email; Role = if (@($cm.roles).Count) { (@($cm.roles) -join ', ') } else { 'member' }
                            })
                        }
                    }
                }
                catch { $errors.Add([pscustomobject]@{ Group = $g.displayName; Check = 'Channels'; Error = $_.Exception.Message }) }
            }
        }
        Write-Progress -Activity 'Reading members' -Completed

        # --- Export ----------------------------------------------------------

        $summaryItems = [ordered]@{
            'Tenant'          = $tenantName
            'Run at'          = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Groups'          = $groupRows.Count
        }
        foreach ($k in $validTypes) { $summaryItems["  $k"] = @($groupRows | Where-Object GroupType -eq $k).Count }
        $summaryItems['Membership rows']          = $memberRows.Count
        $summaryItems['Teams/M365 groups with no owner'] = @($groupRows | Where-Object { $_.Issues -like 'No owner*' }).Count
        $summaryItems['Groups with guests']       = @($groupRows | Where-Object { $_.Guests -gt 0 }).Count
        $summaryItems['Empty groups']             = @($groupRows | Where-Object { $_.Members -is [int] -and $_.Members -eq 0 }).Count
        $summaryItems['Note']                     = 'Members are direct members (nested groups are listed as one row). Exchange dynamic distribution groups are not in Microsoft Graph and are not included.'

        $sheets = @(
            @{ Name = 'Groups'; Rows = @($groupRows | Sort-Object GroupType, Group)
               Headers = @('GroupType', 'Group', 'Email', 'Members', 'Guests', 'Owners', 'Dynamic', 'Visibility', 'Source', 'Created', 'Issues', 'Description')
               Widths  = @(20, 34, 34, 10, 8, 45, 9, 11, 12, 18, 50, 50) }
            @{ Name = 'Members'; Rows = @($memberRows | Sort-Object GroupType, Group, Member)
               Headers = @('GroupType', 'Group', 'GroupEmail', 'Member', 'MemberEmail', 'MemberType', 'Enabled'); Widths = @(20, 34, 34, 28, 36, 13, 9) }
        )
        if ($IncludeTeamsChannels) {
            $sheets += @{ Name = 'Private-Shared Channels'; Rows = @($channelRows | Sort-Object Team, Channel, Member)
                          Headers = @('Team', 'Channel', 'ChannelType', 'Member', 'MemberEmail', 'Role'); Widths = @(30, 30, 12, 28, 36, 12) }
        }
        $sheets += @{ Name = 'Errors'; Rows = $errors; Headers = @('Group', 'Check', 'Error'); Widths = @(34, 12, 100) }
        $sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summaryItems); Headers = @('Item', 'Value'); Widths = @(32, 100) }
        if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

        Write-Host ""
        Write-Host ("Groups: {0}  ({1})" -f $groupRows.Count, (($validTypes | ForEach-Object { "$(@($groupRows | Where-Object GroupType -eq $_).Count) $_" }) -join ', '))
        Write-Host ("No owner: {0}   With guests: {1}   Empty: {2}" -f $summaryItems['Teams/M365 groups with no owner'], $summaryItems['Groups with guests'], $summaryItems['Empty groups'])
        if ($errors.Count) { Write-Host ("Errors: {0} - see the 'Errors' sheet" -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report: $reportPath"
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
