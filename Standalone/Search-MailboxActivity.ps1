<#
.SYNOPSIS
    Searches the Microsoft 365 unified audit log for activity on a mailbox and saves it as evidence.
.DESCRIPTION
    Finds what was done by (and optionally to) a mailbox: mail read, sent and deleted, inbox rules and
    forwarding changed, and from which IP addresses. Replaces the retired Search-MailboxAuditLog. Read-only.
.EXAMPLE
    .\Search-MailboxActivity.ps1 -Mailbox jane.doe@contoso.com -Days 7 -CaseId IR-2026-014
.EXAMPLE
    .\Search-MailboxActivity.ps1 -Mailbox shared@contoso.com -StartDate '2026-10-20' -EndDate '2026-10-23' -IncludeAccessByOthers
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: ExchangeOnlineManagement module; Audit Logs or View-Only Audit Logs role (e.g. Compliance Administrator)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Mailbox,

    [ValidateRange(1, 180)] [int]$Days = 7,
    [datetime]$StartDate,
    [datetime]$EndDate,

    [switch]$IncludeAccessByOthers,
    [switch]$AllWorkloads,
    [ValidateRange(100, 50000)] [int]$MaxRecords = 50000,

    [string]$CaseId,
    [string]$OutputPath,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

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

function New-EvidenceFolder {
    <# Creates <EvidenceFolder>\<Name>. Default location: SOC.EvidenceFolder, otherwise <OutputFolder>\Evidence. #>
    param([hashtable]$Config, [string]$OutputPath, [string]$Name)
    $root = $OutputPath
    if (-not $root) {
        $configured = Get-Setting $Config 'SOC' 'EvidenceFolder' ''
        $root = if ($configured) { Resolve-ToolkitPath $configured } else { Join-Path (Resolve-ToolkitPath (Get-Setting $Config 'General' 'OutputFolder' 'Reports')) 'Evidence' }
    }
    $path = Join-Path $root ($Name -replace '[^\w\.-]', '_')
    New-Item -ItemType Directory -Path $path -Force -WhatIf:$false -Confirm:$false | Out-Null
    return (Resolve-Path -LiteralPath $path).Path
}

function ConvertTo-UtcText($Value) {
    <# Dates as ISO 8601 UTC text, so evidence from different time zones lines up. #>
    if ($null -eq $Value -or "$Value" -eq '') { return '' }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime.ToString('yyyy-MM-dd HH:mm:ss') + 'Z' }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') + 'Z' }
    $dto = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse("$Value", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) {
        return $dto.UtcDateTime.ToString('yyyy-MM-dd HH:mm:ss') + 'Z'
    }
    return "$Value"
}

function Export-EvidenceCsv {
    <# Writes rows to <Folder>\<Name>.csv. Lists are joined with '; ' and dates converted to UTC. Returns the row count. #>
    param([string]$Folder, [string]$Name, $Rows)
    $file = Join-Path $Folder "$Name.csv"
    $list = @($Rows | Where-Object { $null -ne $_ })
    if (-not $list.Count) {
        [System.IO.File]::WriteAllText($file, '', $script:Utf8NoBom)
        return 0
    }
    $flat = foreach ($r in $list) {
        $o = [ordered]@{}
        foreach ($p in $r.PSObject.Properties) {
            if ($p.Name -like 'PS*' -and $p.Name -in 'PSComputerName', 'PSShowComputerName', 'PSSourceJobInstanceId', 'RunspaceId') { continue }
            $v = $p.Value
            if ($v -is [datetime] -or $v -is [datetimeoffset]) { $v = ConvertTo-UtcText $v }
            elseif ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $v = (@($v) | ForEach-Object { "$_" }) -join '; ' }
            $o[$p.Name] = $v
        }
        [pscustomobject]$o
    }
    $flat | Export-Csv -LiteralPath $file -NoTypeInformation -Encoding UTF8 -WhatIf:$false -Confirm:$false
    return $list.Count
}

function Complete-EvidencePackage {
    <#
        Writes manifest.csv (SHA-256 of every file), zips the folder to <Folder>.zip and writes
        <Folder>.zip.sha256. Record the zip hash in your case notes so the evidence can be shown to be unchanged.
    #>
    param([string]$Folder)
    $files = @(Get-ChildItem -LiteralPath $Folder -File -Recurse | Where-Object { $_.Name -ne 'manifest.csv' })
    $manifest = foreach ($f in $files) {
        [pscustomobject]@{
            File   = $f.FullName.Substring($Folder.Length).TrimStart('\', '/')
            Bytes  = $f.Length
            SHA256 = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash
        }
    }
    @($manifest) | Export-Csv -LiteralPath (Join-Path $Folder 'manifest.csv') -NoTypeInformation -Encoding UTF8 -WhatIf:$false -Confirm:$false

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = "$Folder.zip"
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force -WhatIf:$false -Confirm:$false }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($Folder, $zip, [System.IO.Compression.CompressionLevel]::Optimal, $true)
    $hash = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
    [System.IO.File]::WriteAllText("$zip.sha256", "$hash  $(Split-Path -Leaf $zip)`r`n", $script:Utf8NoBom)
    return [pscustomobject]@{ Folder = $Folder; Zip = $zip; Sha256 = $hash; Files = $files.Count + 1 }
}

#endregion
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('EndDate'))   { $EndDate = Get-Date }
if (-not $PSBoundParameters.ContainsKey('StartDate')) { $StartDate = $EndDate.AddDays(-$Days) }
if ($StartDate -ge $EndDate) { throw '-StartDate must be before -EndDate.' }

# Exchange record types that cover mailbox activity
$mailboxRecordTypes = 'ExchangeItem', 'ExchangeItemGroup', 'ExchangeItemAggregated', 'ExchangeAdmin'
$ruleOperations     = 'New-InboxRule', 'Set-InboxRule', 'Enable-InboxRule', 'UpdateInboxRules', 'Set-Mailbox', 'Set-CASMailbox', 'Add-MailboxPermission', 'Add-RecipientPermission'
$deleteOperations   = 'SoftDelete', 'HardDelete', 'MoveToDeletedItems'
$sendOperations     = 'Send', 'SendAs', 'SendOnBehalf'

function Invoke-AuditSearch([hashtable]$Filter) {
    <# Pages through Search-UnifiedAuditLog with a session, up to $MaxRecords results. #>
    $session = "MbxActivity_$([guid]::NewGuid())"
    $found = [System.Collections.Generic.List[object]]::new()
    do {
        $batch = @(Search-UnifiedAuditLog @Filter -StartDate $StartDate -EndDate $EndDate -SessionId $session -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
        foreach ($b in $batch) { $found.Add($b) }
        $total = if ($batch.Count) { [int]$batch[0].ResultCount } else { 0 }
    } while ($batch.Count -gt 0 -and $found.Count -lt $total -and $found.Count -lt $MaxRecords)
    return $found
}

function ConvertFrom-AuditRecord($Record) {
    $d = $null
    try { $d = $Record.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $d = [pscustomobject]@{} }
    $folders  = @($d.Folders | Where-Object { $_ })
    $items    = @($d.AffectedItems | Where-Object { $_ })
    $subjects = @()
    if ($d.Item.Subject) { $subjects += $d.Item.Subject }
    $subjects += @($items | ForEach-Object { $_.Subject } | Where-Object { $_ })
    $folder   = if ($d.Item.ParentFolder.Path) { $d.Item.ParentFolder.Path } elseif ($folders.Count) { (@($folders | ForEach-Object { $_.Path }) | Select-Object -Unique) -join '; ' } elseif ($items.Count) { (@($items | ForEach-Object { $_.ParentFolder.Path }) | Select-Object -Unique) -join '; ' } else { $d.DestFolder.Path }
    $count    = if ($folders.Count) { (@($folders | ForEach-Object { @($_.FolderItems).Count }) | Measure-Object -Sum).Sum } elseif ($items.Count) { $items.Count } elseif ($d.OperationCount) { $d.OperationCount } else { 1 }
    $params   = if ($d.Parameters) { (@($d.Parameters | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; ') } else { '' }
    $ip       = if ($d.ClientIPAddress) { $d.ClientIPAddress } else { $d.ClientIP }
    if ($ip -match '^\[?([0-9a-fA-F:.]+?)\]?(:\d+)?$' -and $ip -notmatch '^[0-9a-fA-F:]+$') { $ip = $Matches[1] }   # strip port from IPv4:port / [IPv6]:port

    [pscustomobject]@{
        Time          = ConvertTo-UtcText $(if ($d.CreationTime) { $d.CreationTime } else { $Record.CreationDate })
        Operation     = if ($d.Operation) { $d.Operation } else { $Record.Operations }
        User          = if ($d.UserId) { $d.UserId } else { $Record.UserIds }
        MailboxOwner  = $d.MailboxOwnerUPN
        LogonType     = switch ("$($d.LogonType)") { '0' { 'Owner' } '1' { 'Admin' } '2' { 'Delegate' } default { "$($d.LogonType)" } }
        ClientIP      = $ip
        Client        = (@($d.ClientInfoString, $d.ClientAppId, $d.AppId) | Where-Object { $_ } | Select-Object -First 1)
        Folder        = $folder
        Items         = $count
        Subjects      = (@($subjects | Select-Object -First 5) -join ' | ')
        Parameters    = $params
        Result        = $d.ResultStatus
        Workload      = $d.Workload
        RecordType    = "$($Record.RecordType)"
        RecordId      = $Record.Identity
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitExchange -Config $config -Bound $PSBoundParameters -RoleHint 'an audit log role (e.g. Compliance Administrator)'
    try {
        Write-Host ("Searching the audit log for {0}, {1} to {2} (UTC)..." -f $Mailbox, $StartDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm'), $EndDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm'))
        $log = [System.Collections.Generic.List[object]]::new()
        $raw = [System.Collections.Generic.List[object]]::new()

        # 1. Everything done by the account itself
        if ($AllWorkloads) { $searches = @(@{ Label = 'All workloads'; Filter = @{ UserIds = $Mailbox } }) }
        else { $searches = @($mailboxRecordTypes | ForEach-Object { @{ Label = $_; Filter = @{ UserIds = $Mailbox; RecordType = $_ } } }) }
        # 2. Optionally, mailbox activity by anyone, filtered to this mailbox afterwards (slower)
        if ($IncludeAccessByOthers) { $searches += @('ExchangeItem', 'ExchangeItemGroup', 'ExchangeItemAggregated' | ForEach-Object { @{ Label = "$_ (all users)"; Filter = @{ RecordType = $_ }; OthersOnly = $true } }) }

        foreach ($s in $searches) {
            Write-Host "  $($s.Label)..."
            try {
                $found = @(Invoke-AuditSearch $s.Filter)
                if ($s.OthersOnly) { $found = @($found | Where-Object { "$($_.AuditData)" -match ('"MailboxOwnerUPN"\s*:\s*"' + [regex]::Escape($Mailbox) + '"') -and "$($_.UserIds)" -ne $Mailbox }) }
                foreach ($f in $found) { $raw.Add($f) }
                $log.Add([pscustomobject]@{ Search = $s.Label; Status = 'OK'; Records = $found.Count; Detail = $(if ($found.Count -ge $MaxRecords) { "Stopped at $MaxRecords records - narrow the dates" } else { '' }) })
            }
            catch { $log.Add([pscustomobject]@{ Search = $s.Label; Status = 'Failed'; Records = 0; Detail = $_.Exception.Message }) }
        }

        $seen = @{}
        $rows = @($raw | Where-Object { if ($seen.ContainsKey($_.Identity)) { $false } else { $seen[$_.Identity] = $true; $true } } |
                  ForEach-Object { ConvertFrom-AuditRecord $_ } | Sort-Object Time)
        Write-Host "  $($rows.Count) event(s)"

        # --- Analysis --------------------------------------------------------
        $ips = @($rows | Where-Object ClientIP | Group-Object ClientIP | ForEach-Object {
            $g = @($_.Group)
            [pscustomobject]@{
                ClientIP = $_.Name; Events = $g.Count; FirstSeen = $g[0].Time; LastSeen = $g[-1].Time
                Users = (@($g.User | Select-Object -Unique) -join ', '); Operations = (@($g | Group-Object Operation | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ')
                Clients = (@($g.Client | Where-Object { $_ } | Select-Object -Unique | Select-Object -First 3) -join ' | ')
            }
        } | Sort-Object Events -Descending)
        $ops = @($rows | Group-Object Operation | ForEach-Object { [pscustomobject]@{ Operation = $_.Name; Events = $_.Count; Items = ($_.Group | Measure-Object Items -Sum).Sum } } | Sort-Object Events -Descending)

        $indicators = [System.Collections.Generic.List[object]]::new()
        foreach ($r in @($rows | Where-Object { $ruleOperations -contains $_.Operation })) {
            $sev = if ($r.Parameters -match '(?i)ForwardTo|RedirectTo|ForwardingSmtpAddress|DeleteMessage|MoveToFolder') { 'High' } else { 'Medium' }
            $indicators.Add([pscustomobject]@{ Severity = $sev; Indicator = "$($r.Operation) by $($r.User)"; Time = $r.Time; Detail = "$($r.Parameters) (IP $($r.ClientIP))" })
        }
        $deletes = @($rows | Where-Object { $deleteOperations -contains $_.Operation })
        $deleted = ($deletes | Measure-Object Items -Sum).Sum
        if ($deleted -ge 50) { $indicators.Add([pscustomobject]@{ Severity = 'Medium'; Indicator = "$deleted item(s) deleted"; Time = "$($deletes[0].Time) - $($deletes[-1].Time)"; Detail = (@($deletes.Operation | Select-Object -Unique) -join ', ') }) }
        $sends = @($rows | Where-Object { $sendOperations -contains $_.Operation })
        if ($sends.Count) { $indicators.Add([pscustomobject]@{ Severity = 'Low'; Indicator = "$($sends.Count) message(s) sent"; Time = "$($sends[0].Time) - $($sends[-1].Time)"; Detail = "From IP(s): $(@($sends.ClientIP | Where-Object { $_ } | Select-Object -Unique) -join ', ')" }) }
        $others = @($rows | Where-Object { $_.User -and $_.User -ne $Mailbox -and $_.MailboxOwner -eq $Mailbox })
        foreach ($g in ($others | Group-Object User)) { $indicators.Add([pscustomobject]@{ Severity = 'Low'; Indicator = "Mailbox accessed by $($g.Name)"; Time = "$($g.Group[0].Time) - $($g.Group[-1].Time)"; Detail = "$($g.Count) event(s): $(@($g.Group.Operation | Select-Object -Unique) -join ', ')" }) }
        if ($ips.Count -gt 1) { $indicators.Add([pscustomobject]@{ Severity = 'Low'; Indicator = "Activity from $($ips.Count) IP addresses"; Time = ''; Detail = (@($ips | Select-Object -First 10 | ForEach-Object { "$($_.ClientIP) ($($_.Events))" }) -join ', ') }) }
        $sevOrder = @{ High = 0; Medium = 1; Low = 2 }
        $indicators = @($indicators | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } }, Time)

        # --- Evidence package ------------------------------------------------
        $name   = (@($CaseId, 'MailboxActivity', ($Mailbox -split '@')[0], (Get-Date -Format 'yyyyMMdd_HHmmss')) | Where-Object { $_ }) -join '_'
        $folder = New-EvidenceFolder -Config $config -OutputPath $OutputPath -Name $name
        [void](Export-EvidenceCsv -Folder $folder -Name '01_Activity' -Rows $rows)
        [void](Export-EvidenceCsv -Folder $folder -Name '02_IPAddresses' -Rows $ips)
        [void](Export-EvidenceCsv -Folder $folder -Name '03_Operations' -Rows $ops)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_Indicators' -Rows $indicators)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_SearchLog' -Rows $log)

        $html = New-HtmlReport -Title "Mailbox activity - $Mailbox" -IntroLines @(
            "Period: $(ConvertTo-UtcText $StartDate) to $(ConvertTo-UtcText $EndDate). $($rows.Count) event(s). All times are UTC.",
            "Searched: activity by the account$(if ($IncludeAccessByOthers) { ', and by anyone else on this mailbox' }). Run by $env:USERDOMAIN\$env:USERNAME$(if ($CaseId) { " - case $CaseId" }).",
            'Some events (e.g. MailItemsAccessed) are only recorded with the right audit licensing and settings; absence of an event is not proof it did not happen.'
        ) -Sections @(
            @{ Heading = "Indicators ($($indicators.Count))"; Html = ConvertTo-HtmlTable -Rows $indicators -Columns @('Severity', 'Indicator', 'Time', 'Detail') -StatusColumn 'Severity' -EmptyText 'No automatic indicators.' }
            @{ Heading = 'IP addresses'; Html = ConvertTo-HtmlTable -Rows @($ips | Select-Object -First 25) -Columns @('ClientIP', 'Events', 'FirstSeen', 'LastSeen', 'Operations') -EmptyText 'No IP addresses recorded.' }
            @{ Heading = 'Operations'; Html = ConvertTo-HtmlTable -Rows $ops -Columns @('Operation', 'Events', 'Items') -EmptyText 'No events found.' }
            @{ Heading = 'Searches'; Html = ConvertTo-HtmlTable -Rows $log -Columns @('Search', 'Status', 'Records', 'Detail') -StatusColumn 'Status' }
        )
        [System.IO.File]::WriteAllText((Join-Path $folder '00_Summary.html'), $html, $script:Utf8NoBom)
        $package = Complete-EvidencePackage -Folder $folder

        Write-Host ""
        foreach ($i in $indicators | Select-Object -First 10) { Write-Host ("  {0,-6} {1} - {2}" -f $i.Severity, $i.Indicator, $i.Detail) -ForegroundColor $(switch ($i.Severity) { 'High' { 'Red' } 'Medium' { 'Yellow' } default { 'Gray' } }) }
        $failed = @($log | Where-Object Status -ne 'OK')
        if ($failed.Count) { Write-Host ("Some searches failed: {0}" -f (($failed | ForEach-Object { "$($_.Search): $($_.Detail)" }) -join '; ')) -ForegroundColor Yellow }
        Write-Host "Evidence : $($package.Zip)"
        Write-Host "SHA-256  : $($package.Sha256)  (record this in your case notes)"
    }
    finally {
        if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    }
}
