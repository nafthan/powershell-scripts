<#
.SYNOPSIS
    Lists everyone with an Entra ID admin role, with MFA, last sign-in and risk flags.
.DESCRIPTION
    Includes active, PIM-eligible and group-based assignments, users, groups and apps.
    Shows what changed since the last run. Read-only.
.EXAMPLE
    .\Get-AdminRoleReport.ps1
.EXAMPLE
    .\Get-AdminRoleReport.ps1 -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: Microsoft.Graph.Authentication module, Global Reader
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

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

$methodNames = @{
    email                            = 'Email (SSPR only)'
    mobilePhone                      = 'Mobile phone (SMS/voice)'
    alternateMobilePhone             = 'Alternate mobile phone'
    officePhone                      = 'Office phone'
    microsoftAuthenticatorPush       = 'Microsoft Authenticator'
    microsoftAuthenticatorPasswordless = 'Authenticator passwordless'
    softwareOneTimePasscode          = 'Software OATH token'
    hardwareOneTimePasscode          = 'Hardware OATH token'
    windowsHelloForBusiness          = 'Windows Hello for Business'
    fido2SecurityKey                 = 'FIDO2 security key'
    passKeyDeviceBound               = 'Passkey'
    passKeyDeviceBoundAuthenticator  = 'Passkey (Authenticator)'
    passKeyDeviceBoundWindowsHello   = 'Passkey (Windows Hello)'
    temporaryAccessPass              = 'Temporary Access Pass'
    securityQuestion                 = 'Security questions'
    macOsSecureEnclaveKey            = 'macOS Platform SSO'
    x509Certificate                  = 'Certificate'
}

$phoneMethods     = @('mobilePhone', 'alternateMobilePhone', 'officePhone')

$appMethods       = @('microsoftAuthenticatorPush', 'microsoftAuthenticatorPasswordless', 'softwareOneTimePasscode', 'hardwareOneTimePasscode')

$phishResistant   = '^(fido2|windowsHello|passKey|macOsSecureEnclave|x509|certificate)'

$passwordlessRx   = '^(fido2|windowsHello|passKey|macOsSecureEnclave|microsoftAuthenticatorPasswordless)'

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

function Get-Latest([object[]]$Dates) {
    $valid = @($Dates | Where-Object { $_ -is [datetime] })
    if (-not $valid.Count) { return $null }
    return ($valid | Sort-Object -Descending | Select-Object -First 1)
}

function Test-StrongMethod([string]$Method) {
    return ($phoneMethods -contains $Method) -or ($appMethods -contains $Method) -or ($Method -match $phishResistant) -or ($Method -match $passwordlessRx)
}

function ConvertFrom-MethodObject($Method) {
    # /users/{id}/authentication/methods item -> registration-report style name
    $type = "$($Method.'@odata.type')" -replace '^#microsoft\.graph\.', ''
    switch ($type) {
        'microsoftAuthenticatorAuthenticationMethod'   { 'microsoftAuthenticatorPush' }
        'phoneAuthenticationMethod' {
            switch ("$($Method.phoneType)") { 'alternateMobile' { 'alternateMobilePhone' } 'office' { 'officePhone' } default { 'mobilePhone' } }
        }
        'fido2AuthenticationMethod'                    { 'fido2SecurityKey' }
        'windowsHelloForBusinessAuthenticationMethod'  { 'windowsHelloForBusiness' }
        'softwareOathAuthenticationMethod'             { 'softwareOneTimePasscode' }
        'hardwareOathAuthenticationMethod'             { 'hardwareOneTimePasscode' }
        'emailAuthenticationMethod'                    { 'email' }
        'temporaryAccessPassAuthenticationMethod'      { 'temporaryAccessPass' }
        'platformCredentialAuthenticationMethod'       { 'macOsSecureEnclaveKey' }
        'x509CertificateAuthenticationMethod'          { 'x509Certificate' }
        'passwordAuthenticationMethod'                 { }
        default                                        { if ($type) { $type } }
    }
}

function Get-GraphUserMfa([string]$UserId) {
    <#
        MFA registration for one user. Uses the registration report (Entra ID P1/P2) when available,
        otherwise reads the user's authentication methods.
    #>
    $methods = $null; $registered = $null; $source = ''
    if ($script:MfaReportAvailable -ne $false) {
        try {
            $r = Invoke-Graph "v1.0/reports/authenticationMethods/userRegistrationDetails/$UserId"
            $script:MfaReportAvailable = $true
            $methods = @($r.methodsRegistered | Where-Object { $_ }); $registered = [bool]$r.isMfaRegistered; $source = 'Registration report'
        }
        catch {
            if ($_.Exception.Message -notmatch '404|NotFound|ResourceNotFound') { $script:MfaReportAvailable = $false }
        }
    }
    if ($null -eq $registered) {
        try {
            $methods = @(Get-GraphAll "v1.0/users/$UserId/authentication/methods" | ForEach-Object { ConvertFrom-MethodObject $_ } | Where-Object { $_ } | Sort-Object -Unique)
            $registered = [bool](@($methods | Where-Object { Test-StrongMethod $_ }).Count)
            $source = 'Methods API'
        }
        catch { $methods = @(); $source = "Unavailable: $($_.Exception.Message)" }
    }
    [pscustomobject]@{
        Registered        = $registered
        Methods           = (@($methods | ForEach-Object { if ($methodNames.ContainsKey($_)) { $methodNames[$_] } else { $_ } }) -join ', ')
        PhishingResistant = [bool](@($methods | Where-Object { $_ -match $phishResistant }).Count)
        Source            = $source
    }
}

#endregion
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('StaleDays')) { $StaleDays = [int](Get-Setting $config 'AdminRoles' 'StaleDays' 90) }
$maxGlobalAdmins = [int](Get-Setting $config 'AdminRoles' 'MaxGlobalAdmins' 4)

$scopes = @('RoleManagement.Read.Directory', 'Directory.Read.All', 'AuditLog.Read.All', 'UserAuthenticationMethod.Read.All')

# Roles that can change security settings or other admins (used when Graph doesn't say)
$privilegedRoles = @(
    'Global Administrator', 'Privileged Role Administrator', 'Privileged Authentication Administrator', 'Security Administrator',
    'Conditional Access Administrator', 'Exchange Administrator', 'SharePoint Administrator', 'User Administrator',
    'Authentication Administrator', 'Application Administrator', 'Cloud Application Administrator', 'Hybrid Identity Administrator',
    'Intune Administrator', 'Helpdesk Administrator', 'Password Administrator', 'Groups Administrator', 'Domain Name Administrator',
    'External Identity Provider Administrator', 'Partner Tier2 Support', 'Directory Synchronization Accounts', 'Teams Administrator'
)

function Test-NotFound($ErrorRecord) { return "$($ErrorRecord.Exception.Message)" -match '404|NotFound|Request_ResourceNotFound|does not exist' }

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'Global Reader'
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'AdminRoles' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        function Add-ReportError([string]$Section, [string]$Message) { $errors.Add([pscustomobject]@{ Section = $Section; Error = $Message }) }
        $tenantName = Get-GraphTenantName
        Write-Host "Tenant : $tenantName"

        # --- Roles and assignments ----------------------------------------

        Write-Host 'Getting roles and assignments...'
        $roles = @{}
        foreach ($d in (Get-GraphAll 'v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName,isBuiltIn,isPrivileged')) {
            $isPriv = if ($null -ne $d.isPrivileged) { [bool]$d.isPrivileged } else { $privilegedRoles -contains $d.displayName }
            $roles[$d.id] = [pscustomobject]@{ Name = $d.displayName; Privileged = $isPriv -or ($privilegedRoles -contains $d.displayName) }
        }

        $assignments = [System.Collections.Generic.List[object]]::new()
        $pimAvailable = $true
        try {
            foreach ($a in (Get-GraphAll 'v1.0/roleManagement/directory/roleAssignmentSchedules')) {
                $end = ConvertTo-LocalTime $a.scheduleInfo.expiration.endDateTime
                $assignments.Add([pscustomobject]@{
                    PrincipalId = $a.principalId; RoleId = $a.roleDefinitionId; Scope = $a.directoryScopeId
                    Assignment  = if ($a.assignmentType -eq 'Activated') { 'Active (PIM activated)' } elseif ($end) { 'Active (time-bound)' } else { 'Active (permanent)' }
                    Ends = $end
                })
            }
        }
        catch {
            $pimAvailable = $false
            Add-ReportError 'Assignment schedules' "PIM details unavailable (needs Entra ID P2), so active assignments are shown without start/end: $($_.Exception.Message)"
            foreach ($a in (Get-GraphAll 'v1.0/roleManagement/directory/roleAssignments?$select=principalId,roleDefinitionId,directoryScopeId')) {
                $assignments.Add([pscustomobject]@{ PrincipalId = $a.principalId; RoleId = $a.roleDefinitionId; Scope = $a.directoryScopeId; Assignment = 'Active'; Ends = $null })
            }
        }
        $eligibleKnown = $false
        if ($pimAvailable) {
            try {
                foreach ($a in (Get-GraphAll 'v1.0/roleManagement/directory/roleEligibilitySchedules')) {
                    $assignments.Add([pscustomobject]@{
                        PrincipalId = $a.principalId; RoleId = $a.roleDefinitionId; Scope = $a.directoryScopeId
                        Assignment = 'Eligible (PIM)'; Ends = ConvertTo-LocalTime $a.scheduleInfo.expiration.endDateTime
                    })
                }
                $eligibleKnown = $true
            }
            catch { Add-ReportError 'Eligible roles (PIM)' $_.Exception.Message }
        }
        Write-Host "  $($assignments.Count) assignment(s)"

        # --- Principals ---------------------------------------------------

        $userSelect = 'id,displayName,userPrincipalName,accountEnabled,userType,createdDateTime,onPremisesSyncEnabled,assignedLicenses'
        $signInAvailable = $true
        $principals = @{}
        $scopeNames = @{ '/' = 'Tenant' }

        function Get-UserPrincipal([string]$Id) {
            $u = $null
            if ($signInAvailable) {
                try { $u = Invoke-Graph "v1.0/users/$Id`?`$select=$userSelect,signInActivity" }
                catch {
                    if (Test-NotFound $_) { return $null }
                    $script:signInAvailable = $false
                    Add-ReportError 'Sign-in activity' "Last sign-in dates unavailable (needs Entra ID P1/P2 and AuditLog.Read.All): $($_.Exception.Message)"
                }
            }
            if (-not $u) {
                try { $u = Invoke-Graph "v1.0/users/$Id`?`$select=$userSelect" }
                catch { if (Test-NotFound $_) { return $null } else { throw } }
            }
            $mfa = Get-GraphUserMfa $u.id
            $sia = $u.signInActivity
            $last = ConvertTo-LocalTime $sia.lastSuccessfulSignInDateTime
            if (-not $last) { $last = Get-Latest @((ConvertTo-LocalTime $sia.lastSignInDateTime), (ConvertTo-LocalTime $sia.lastNonInteractiveSignInDateTime)) }
            [pscustomobject]@{
                Id = $u.id; Type = $(if ($u.userType -eq 'Guest') { 'Guest' } else { 'User' }); Name = $u.displayName; Upn = $u.userPrincipalName
                Enabled = [bool]$u.accountEnabled; Created = ConvertTo-LocalTime $u.createdDateTime; Synced = [bool]$u.onPremisesSyncEnabled
                Licensed = @($u.assignedLicenses).Count -gt 0; LastSignIn = $last; Mfa = $mfa
            }
        }

        function Get-Principal([string]$Id) {
            if ($principals.ContainsKey($Id)) { return $principals[$Id] }
            $p = $null
            try {
                $p = Get-UserPrincipal $Id
                if (-not $p) {
                    try {
                        $g = Invoke-Graph "v1.0/groups/$Id`?`$select=id,displayName,isAssignableToRole"
                        $members = @(Get-GraphAll "v1.0/groups/$Id/transitiveMembers/microsoft.graph.user?`$select=id" | ForEach-Object { $_.id })
                        $p = [pscustomobject]@{ Id = $g.id; Type = 'Group'; Name = $g.displayName; Upn = ''; Members = $members }
                    }
                    catch { if (-not (Test-NotFound $_)) { throw } }
                }
                if (-not $p) {
                    try {
                        $sp = Invoke-Graph "v1.0/servicePrincipals/$Id`?`$select=id,displayName,appId,accountEnabled"
                        $p = [pscustomobject]@{ Id = $sp.id; Type = 'App'; Name = $sp.displayName; Upn = $sp.appId; Enabled = [bool]$sp.accountEnabled }
                    }
                    catch { if (-not (Test-NotFound $_)) { throw } }
                }
            }
            catch { Add-ReportError 'Principal lookup' "${Id}: $($_.Exception.Message)" }
            if (-not $p) { $p = [pscustomobject]@{ Id = $Id; Type = 'Deleted / unknown'; Name = "(object $Id)"; Upn = '' } }
            $principals[$Id] = $p
            return $p
        }

        function Get-ScopeName([string]$Scope) {
            if (-not $Scope) { return 'Tenant' }
            if (-not $scopeNames.ContainsKey($Scope)) {
                $name = $Scope
                if ($Scope -match '^/administrativeUnits/(.+)$') {
                    try { $name = "Admin unit: $((Invoke-Graph "v1.0/directory/administrativeUnits/$($Matches[1])?`$select=displayName").displayName)" } catch { $name = "Admin unit $($Matches[1])" }
                }
                elseif ($Scope -match '^/([0-9a-f-]{36})$') { $name = "Object $($Matches[1])" }
                $scopeNames[$Scope] = $name
            }
            return $scopeNames[$Scope]
        }

        # --- Rows ---------------------------------------------------------

        Write-Host 'Checking each admin (MFA, sign-in)...'
        $riskRank = @{ OK = 0; Low = 1; Medium = 2; High = 3; Critical = 4 }
        $rows = [System.Collections.Generic.List[object]]::new()

        function New-AssignmentRow($Assignment, $Principal, [string]$Via) {
            $role = $roles[$Assignment.RoleId]
            $roleName = if ($role) { $role.Name } else { $Assignment.RoleId }
            $isPriv = $role -and $role.Privileged
            $issues = [System.Collections.Generic.List[string]]::new()
            $script:rowRisk = 'OK'
            $raise = { param($Level, $Text) $issues.Add($Text); if ($riskRank[$Level] -gt $riskRank[$script:rowRisk]) { $script:rowRisk = $Level } }

            switch ($Principal.Type) {
                { $_ -in 'User', 'Guest' } {
                    if (-not $Principal.Enabled) { & $raise 'Low' 'Account disabled but still holds the role - remove it' }
                    else {
                        if ($Principal.Type -eq 'Guest')                { & $raise 'High' 'Guest account with an admin role' }
                        if ($Principal.Mfa.Registered -eq $false)       { & $raise 'Critical' 'No MFA registered' }
                        elseif ($null -eq $Principal.Mfa.Registered)    { & $raise 'Medium' 'MFA status could not be read' }
                        elseif ($isPriv -and -not $Principal.Mfa.PhishingResistant) { & $raise 'Low' 'No phishing-resistant MFA method' }
                        if ($isPriv -and $Principal.Synced)             { & $raise 'Medium' 'Synced from on-premises AD - use a cloud-only admin account' }
                        if ($isPriv -and $Principal.Licensed)           { & $raise 'Low' 'Admin account is licensed - use a separate unlicensed admin account' }
                        if ($signInAvailable) {
                            if (-not $Principal.LastSignIn -and $Principal.Created -and ($now - $Principal.Created).TotalDays -gt 14) { & $raise 'Low' 'Never signed in - remove the role if not needed' }
                            elseif ($Principal.LastSignIn -and ($now - $Principal.LastSignIn).TotalDays -gt $StaleDays) { & $raise 'Low' "No sign-in for $([int]($now - $Principal.LastSignIn).TotalDays) days - remove the role if not needed" }
                        }
                    }
                }
                'App'               { if ($isPriv) { & $raise 'Medium' 'App holds a privileged role - check it is still needed' } }
                'Deleted / unknown' { & $raise 'Low' 'Assigned to a deleted or unknown object - remove the assignment' }
            }
            if ($pimAvailable -and $isPriv -and $Assignment.Assignment -eq 'Active (permanent)' -and $Principal.Type -in 'User', 'Guest') {
                & $raise 'Low' 'Permanent assignment - consider making it PIM-eligible'
            }

            [pscustomobject]@{
                Risk            = $script:rowRisk
                Role            = $roleName
                Privileged      = [bool]$isPriv
                Name            = $Principal.Name
                UserPrincipalName = $Principal.Upn
                PrincipalType   = $Principal.Type
                Assignment      = $Assignment.Assignment
                Via             = $Via
                Scope           = Get-ScopeName $Assignment.Scope
                Ends            = if ($Assignment.Ends) { $Assignment.Ends } else { '' }
                Enabled         = if ($null -ne $Principal.Enabled) { [bool]$Principal.Enabled } else { '' }
                MfaRegistered   = if ($Principal.Mfa) { if ($null -eq $Principal.Mfa.Registered) { 'Unknown' } elseif ($Principal.Mfa.Registered) { 'Yes' } else { 'No' } } else { '' }
                MfaMethods      = if ($Principal.Mfa) { $Principal.Mfa.Methods } else { '' }
                PhishingResistant = if ($Principal.Mfa) { [bool]$Principal.Mfa.PhishingResistant } else { '' }
                LastSignIn      = if ($Principal.Type -in 'User', 'Guest') { if ($signInAvailable) { $Principal.LastSignIn } else { 'Unavailable' } } else { '' }
                Synced          = if ($Principal.Type -in 'User', 'Guest') { [bool]$Principal.Synced } else { '' }
                Licensed        = if ($Principal.Type -in 'User', 'Guest') { [bool]$Principal.Licensed } else { '' }
                Issues          = ($issues -join '; ')
                PrincipalId     = $Principal.Id
            }
        }

        $n = 0
        foreach ($a in $assignments) {
            $n++
            Write-Progress -Activity 'Checking admins' -Status "$n of $($assignments.Count)" -PercentComplete ($n / [math]::Max(1, $assignments.Count) * 100)
            $p = Get-Principal $a.PrincipalId
            $rows.Add((New-AssignmentRow $a $p ''))
            if ($p.Type -eq 'Group') {
                foreach ($memberId in $p.Members) {
                    $m = Get-Principal $memberId
                    if ($m.Type -in 'User', 'Guest') { $rows.Add((New-AssignmentRow $a $m "Group: $($p.Name)")) }
                }
            }
        }
        Write-Progress -Activity 'Checking admins' -Completed

        $order = @{ Critical = 0; High = 1; Medium = 2; Low = 3; OK = 4 }
        $rowsSorted = @($rows | Sort-Object @{ Expression = { $order[$_.Risk] } }, @{ Expression = { -not $_.Privileged } }, Role, Name)

        # One row per admin
        $admins = @($rows | Group-Object PrincipalId | ForEach-Object {
            $g = $_.Group; $first = $g[0]
            $worst = ($g | Sort-Object @{ Expression = { $order[$_.Risk] } } | Select-Object -First 1).Risk
            [pscustomobject]@{
                Risk              = $worst
                Name              = $first.Name
                UserPrincipalName = $first.UserPrincipalName
                PrincipalType     = $first.PrincipalType
                Roles             = (@($g | ForEach-Object { $label = $_.Role; if ($_.Assignment -ne 'Active (permanent)' -and $_.Assignment -ne 'Active') { $label += " [$($_.Assignment)]" }; if ($_.Via) { $label += " (via $($_.Via -replace '^Group: ', ''))" }; if ($_.Scope -ne 'Tenant') { $label += " - $($_.Scope)" }; $label } | Sort-Object -Unique) -join '; ')
                Enabled           = $first.Enabled
                MfaRegistered     = $first.MfaRegistered
                MfaMethods        = $first.MfaMethods
                PhishingResistant = $first.PhishingResistant
                LastSignIn        = $first.LastSignIn
                Synced            = $first.Synced
                Licensed          = $first.Licensed
                Issues            = (@($g | Sort-Object @{ Expression = { $order[$_.Risk] } } | ForEach-Object { $_.Issues -split '; ' } | Where-Object { $_ } | Select-Object -Unique) -join '; ')
            }
        } | Sort-Object @{ Expression = { $order[$_.Risk] } }, Name)

        # --- Changes since last run ---------------------------------------

        $tenantKey    = "$((Get-MgContext).TenantId)" -replace '[^\w-]', ''
        $snapshotPath = Join-Path (Split-Path -Parent $reportPath) "AdminRoles_LastRun_$tenantKey.json"
        # Compare as Active / Eligible only, and skip temporary PIM activations, so changes in
        # licensing or someone activating an eligible role don't show up as changes
        $current = @($rows | Where-Object { $_.Assignment -ne 'Active (PIM activated)' } | ForEach-Object { [pscustomobject]@{
            Key = "$($_.Role)|$($_.PrincipalId)|$(if ($_.Assignment -like 'Eligible*') { 'Eligible' } else { 'Active' })|$($_.Via)|$($_.Scope)"
            Role = $_.Role; Name = $_.Name; UserPrincipalName = $_.UserPrincipalName; Assignment = $_.Assignment; Via = $_.Via; Scope = $_.Scope } })
        $changes = @()
        $previousRun = ''
        if (Test-Path -LiteralPath $snapshotPath) {
            try {
                $prev = Get-Content -LiteralPath $snapshotPath -Raw | ConvertFrom-Json
                $previousRun = "$($prev.RunAt)"
                $prevItems = @($prev.Items)
                if (-not $eligibleKnown) {
                    # Eligible roles couldn't be read this time - keep last run's instead of reporting them as removed
                    $current += @($prevItems | Where-Object { $_.Key -like '*|Eligible|*' })
                }
                $prevKeys = @{}; foreach ($p in $prevItems) { $prevKeys[$p.Key] = $p }
                $curKeys  = @{}; foreach ($c in $current)   { $curKeys[$c.Key]  = $c }
                $changes += @($current   | Where-Object { -not $prevKeys.ContainsKey($_.Key) } | ForEach-Object { [pscustomobject]@{ Change = 'Added';   Role = $_.Role; Name = $_.Name; UserPrincipalName = $_.UserPrincipalName; Assignment = $_.Assignment; Via = $_.Via; Scope = $_.Scope } })
                $changes += @($prevItems | Where-Object { -not $curKeys.ContainsKey($_.Key) }  | ForEach-Object { [pscustomobject]@{ Change = 'Removed'; Role = $_.Role; Name = $_.Name; UserPrincipalName = $_.UserPrincipalName; Assignment = $_.Assignment; Via = $_.Via; Scope = $_.Scope } })
            }
            catch { Add-ReportError 'Changes' "Could not read the previous run ($snapshotPath): $($_.Exception.Message)" }
        }
        @{ RunAt = $now.ToString('yyyy-MM-dd HH:mm'); Items = $current } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $snapshotPath -Encoding UTF8

        # --- Summary ------------------------------------------------------

        $gaUsers = @($rows | Where-Object { $_.Role -eq 'Global Administrator' -and $_.PrincipalType -in 'User', 'Guest' -and $_.Enabled -eq $true } | Select-Object -ExpandProperty PrincipalId -Unique)
        $gaNote = if ($gaUsers.Count -gt $maxGlobalAdmins) { "Too many - Microsoft recommends fewer than 5" } elseif ($gaUsers.Count -lt 2) { 'Too few - keep at least 2, including an emergency access account' } else { 'OK' }
        $counts = [ordered]@{}
        foreach ($r in 'Critical', 'High', 'Medium', 'Low') { $counts[$r] = @($admins | Where-Object { $_.Risk -eq $r }).Count }

        $summary = [ordered]@{
            'Tenant'                 = $tenantName
            'Run at'                 = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Role assignments'       = $assignments.Count
            'Admins (people/apps)'   = $admins.Count
            'Global Administrators'  = "$($gaUsers.Count) enabled user(s) - $gaNote"
            'Admins without MFA'     = @($admins | Where-Object { $_.MfaRegistered -eq 'No' -and $_.Enabled -eq $true }).Count
            'Risk: Critical'         = $counts.Critical
            'Risk: High'             = $counts.High
            'Risk: Medium'           = $counts.Medium
            'Risk: Low'              = $counts.Low
            'PIM'                    = if ($pimAvailable) { 'Available - active, time-bound and eligible assignments shown' } else { 'Not available (needs Entra ID P2) - active assignments only' }
            'Changes since last run' = if ($previousRun) { "$(@($changes | Where-Object Change -eq 'Added').Count) added, $(@($changes | Where-Object Change -eq 'Removed').Count) removed (since $previousRun)" } else { 'First run - changes will show next time' }
            'Run by'                 = Get-GraphRunBy
        }

        $assignHeaders = @('Risk', 'Role', 'Privileged', 'Name', 'UserPrincipalName', 'PrincipalType', 'Assignment', 'Via', 'Scope', 'Ends', 'Enabled', 'MfaRegistered', 'MfaMethods', 'PhishingResistant', 'LastSignIn', 'Synced', 'Licensed', 'Issues')
        $assignWidths  = @(10, 34, 11, 26, 34, 16, 22, 28, 18, 18, 9, 15, 45, 17, 18, 8, 9, 70)
        $saved = Save-Workbook -FilePath $reportPath -Sheets @(
            @{ Name = 'Admins'; Rows = $admins
               Headers = @('Risk', 'Name', 'UserPrincipalName', 'PrincipalType', 'Roles', 'Enabled', 'MfaRegistered', 'MfaMethods', 'PhishingResistant', 'LastSignIn', 'Synced', 'Licensed', 'Issues')
               Widths  = @(10, 26, 34, 16, 60, 9, 15, 45, 17, 18, 8, 9, 80) }
            @{ Name = 'Assignments'; Rows = $rowsSorted; Headers = $assignHeaders; Widths = $assignWidths }
            @{ Name = 'Changes'; Rows = $changes; Headers = @('Change', 'Role', 'Name', 'UserPrincipalName', 'Assignment', 'Via', 'Scope'); Widths = @(10, 34, 26, 34, 22, 28, 18) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(22, 120) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(24, 100) }
        )
        if (-not $saved) { exit 1 }

        Write-Host ""
        Write-Host ("Admins             : {0}  ({1} critical, {2} high, {3} medium, {4} low)" -f $admins.Count, $counts.Critical, $counts.High, $counts.Medium, $counts.Low) -ForegroundColor $(if ($counts.Critical + $counts.High) { 'Red' } elseif ($counts.Medium) { 'Yellow' } else { 'Green' })
        Write-Host ("Global Admins      : {0} - {1}" -f $gaUsers.Count, $gaNote) -ForegroundColor $(if ($gaNote -eq 'OK') { 'Green' } else { 'Yellow' })
        Write-Host ("Changes            : {0}" -f $summary['Changes since last run'])
        foreach ($c in $changes) { Write-Host ("  {0,-8} {1} - {2} ({3})" -f $c.Change, $c.Role, $c.Name, $c.Assignment) -ForegroundColor Cyan }
        if ($errors.Count) { Write-Host ("Some data was unavailable - see the 'Errors' sheet ({0})." -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report             : $reportPath"

        if ($SendEmail) {
            $intro = @("$tenantName - $($admins.Count) admin(s). Global Administrators: $($gaUsers.Count) ($gaNote).", "Changes since last run: $($summary['Changes since last run']).")
            $html = New-HtmlReport -Title 'Entra ID admin role report' -IntroLines $intro -Sections @(
                @{ Heading = "Changes since last run ($($changes.Count))"; Html = ConvertTo-HtmlTable -Rows $changes -Columns @('Change', 'Role', 'Name', 'UserPrincipalName', 'Assignment', 'Via') -EmptyText 'No changes.' }
                @{ Heading = 'Admins needing attention'; Html = ConvertTo-HtmlTable -Rows @($admins | Where-Object { $_.Risk -in 'Critical', 'High', 'Medium' }) -Columns @('Risk', 'Name', 'Roles', 'MfaRegistered', 'Issues') -StatusColumn 'Risk' -EmptyText 'No critical, high or medium issues.' }
            )
            $state = if ($counts.Critical) { 'CRITICAL' } elseif ($counts.High) { 'HIGH' } elseif ($changes.Count) { 'Changes' } else { 'OK' }
            if (Send-ToolkitMail -Config $config -Subject "[$state] Entra admin roles - $($admins.Count) admins, $($changes.Count) change(s)" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
