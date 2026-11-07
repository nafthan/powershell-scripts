<#
.SYNOPSIS
    Contains a compromised user account in Entra ID / Microsoft 365 and Active Directory, after saving evidence.
.DESCRIPTION
    1. Collects evidence first: account details, groups, MFA methods, sign-ins, audit events, devices,
       app consents, mailbox forwarding and inbox rules.
    2. Then (asking before each step): disables the account, resets the password to a random value,
       revokes sessions, removes mail forwarding and disables suspicious inbox rules.
    Works for cloud-only and AD-synced users. Use -WhatIf to preview, or -EvidenceOnly to change nothing.
.EXAMPLE
    .\Invoke-AccountContainment.ps1 -User jane.doe@contoso.com -CaseId IR-2026-014 -WhatIf
.EXAMPLE
    .\Invoke-AccountContainment.ps1 -User jane.doe@contoso.com -CaseId IR-2026-014
.EXAMPLE
    .\Invoke-AccountContainment.ps1 -User jane.doe@contoso.com -EvidenceOnly
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: Microsoft.Graph.Authentication module (+ ExchangeOnlineManagement for mailbox steps, ActiveDirectory for synced users).
              Roles: User Administrator + Authentication Administrator (Privileged Authentication Administrator if the user is an admin),
              Exchange Administrator, and rights to reset passwords in AD.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$User,
    [string]$CaseId,

    [switch]$EvidenceOnly,
    [string[]]$Skip,
    [datetime]$RemoveMfaMethodsAddedAfter,

    [ValidateRange(1, 30)] [int]$SignInDays,
    [switch]$SkipExchange,
    [switch]$SkipAD,
    [switch]$SkipCloud,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [string]$OutputPath,

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

$hidingFolders   = 'RSS|Conversation History|Archive|Junk|Deleted Items|Notes|Sync Issues|Outbox'

$sensitiveWords  = 'invoice|payment|bank|wire|transfer|remit|password|credential|security|phish|hack|compromis|suspicious|helpdesk|it support|mfa|verify'

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

function Get-AddressInfo([string]$Entry, [string[]]$InternalDomains) {
    # Rule targets look like: "Name" [SMTP:user@domain.com]  or  "Name" [EX:/o=ExchangeLabs/...]
    $address = $Entry
    if ($Entry -match 'SMTP:([^\]\s]+)') { $address = $Matches[1] }
    elseif ($Entry -match '^smtp:(.+)$') { $address = $Matches[1] }
    elseif ($Entry -match '\[EX:') { return [pscustomobject]@{ Address = ($Entry -replace '\s*\[EX:.*$', '').Trim('"', ' '); External = $false } }
    $domain = if ($address -match '@(.+)$') { $Matches[1].ToLower() } else { '' }
    $external = $domain -and -not ($InternalDomains | Where-Object { $domain -eq $_ -or $domain.EndsWith(".$_") })
    return [pscustomobject]@{ Address = $address; External = [bool]$external }
}

function Get-InboxRuleAssessment {
    <#
        Rates an inbox rule (from Get-InboxRule). Risk is '' when there is nothing to report.
        High: forwards externally, or hides mail (deletes/moves it and marks it read, or filters on sensitive words).
        Critical: both. Medium: unusual name (only symbols). Low: forwards internally.
    #>
    param($Rule, [string[]]$InternalDomains)
    $forwards = @()
    foreach ($field in 'ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo') {
        foreach ($entry in @($Rule.$field | Where-Object { $_ })) { $forwards += [pscustomobject]@{ Action = $field; Info = Get-AddressInfo "$entry" $InternalDomains } }
    }
    $external = @($forwards | Where-Object { $_.Info.External })
    $words    = (@($Rule.SubjectContainsWords) + @($Rule.BodyContainsWords) + @($Rule.SubjectOrBodyContainsWords) | Where-Object { $_ }) -join ', '
    $hidesMail = $Rule.DeleteMessage -or $Rule.SoftDeleteMessage -or ("$($Rule.MoveToFolder)" -match $hidingFolders)
    $hides    = [bool]($hidesMail -and ($Rule.MarkAsRead -or $words -match $sensitiveWords))
    $unusual  = ("$($Rule.Name)" -match '^[\W_]{1,4}$' -or "$($Rule.Name)".Trim().Length -le 1)
    $issues = @(); $short = @(); $risk = ''

    if ($external.Count) { $issues += 'Forwards to an external address'; $short += 'Forwards externally'; $risk = 'High' }
    elseif ($forwards.Count) { $issues += 'Forwards internally'; $short += 'Forwards internally'; $risk = 'Low' }
    if ($hides) {
        $issues += 'Hides messages (deletes or moves them and marks them read, or filters on sensitive words) - common after an account takeover'
        $short  += 'Hides messages'
        $risk = if ($risk -eq 'High') { 'Critical' } else { 'High' }
    }
    if ($unusual) {
        $issues += 'Unusual rule name (only symbols or one character)'
        $short  += 'Unusual rule name'
        if (-not $risk -or $risk -eq 'Low') { $risk = 'Medium' }
    }

    $actions = @()
    if ($forwards.Count)      { $actions += (@($forwards.Action | Sort-Object -Unique) -join '/') }
    if ($Rule.DeleteMessage)  { $actions += 'Delete' }
    if ($Rule.MoveToFolder)   { $actions += "Move to $($Rule.MoveToFolder)" }
    if ($Rule.MarkAsRead)     { $actions += 'Mark as read' }

    [pscustomobject]@{
        Risk               = $risk
        Issues             = $issues
        Short              = $short
        ForwardsExternally = [bool]$external.Count
        Hides              = $hides
        Targets            = (@($forwards | ForEach-Object { "$($_.Info.Address)$(if ($_.Info.External) { ' (external)' })" }) -join ', ')
        Actions            = ($actions -join '; ')
        Conditions         = if ($words) { "Words: $words" } elseif ($Rule.From) { "From: $(@($Rule.From) -join ', ')" } else { '' }
    }
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

if (-not $PSBoundParameters.ContainsKey('SignInDays')) { $SignInDays = [int](Get-Setting $config 'SOC' 'SignInDays' 7) }
$Skip = @(Split-ListValue $Skip)
$validSteps = 'Disable', 'ResetPassword', 'RevokeSessions', 'Mailbox'
foreach ($s in $Skip) { if ($validSteps -notcontains $s) { throw "Unknown step '$s' in -Skip. Valid steps: $($validSteps -join ', ')." } }

$readScopes  = @('User.Read.All', 'UserAuthenticationMethod.Read.All', 'AuditLog.Read.All', 'Directory.Read.All')
$writeScopes = @('User.ReadWrite.All', 'User.RevokeSessions.All', 'UserAuthenticationMethod.ReadWrite.All', 'AuditLog.Read.All', 'Directory.Read.All')
$passwordMethodId = '28c10230-6103-485e-b985-444c60001490'   # fixed ID of the password authentication method

$methodSegments = @{
    microsoftAuthenticatorAuthenticationMethod  = 'microsoftAuthenticatorMethods'
    fido2AuthenticationMethod                   = 'fido2Methods'
    softwareOathAuthenticationMethod            = 'softwareOathMethods'
    windowsHelloForBusinessAuthenticationMethod = 'windowsHelloForBusinessMethods'
    phoneAuthenticationMethod                   = 'phoneMethods'
    emailAuthenticationMethod                   = 'emailMethods'
    temporaryAccessPassAuthenticationMethod     = 'temporaryAccessPassMethods'
    platformCredentialAuthenticationMethod      = 'platformCredentialMethods'
}

function Invoke-GraphChange([string]$Method, [string]$Uri, $Body) {
    $p = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $p.Body = ($Body | ConvertTo-Json -Depth 5 -Compress); $p.ContentType = 'application/json' }
    Invoke-MgGraphRequest @p
}

function New-RandomPassword([int]$Length = 24) {
    $sets = @('ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnopqrstuvwxyz', '23456789', '!#%^*-_=+?')
    $all  = -join $sets
    $rng  = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = New-Object byte[] $Length
    $rng.GetBytes($bytes)
    $chars = for ($i = 0; $i -lt $Length; $i++) { $set = if ($i -lt $sets.Count) { $sets[$i] } else { $all }; $set[$bytes[$i] % $set.Length] }
    return -join ($chars | Sort-Object { Get-Random })
}

if ($MyInvocation.InvocationName -ne '.') {
    $start   = Get-Date
    $since   = $start.AddDays(-$SignInDays)
    $sinceIso = $since.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $log     = [System.Collections.Generic.List[object]]::new()
    $actions = [System.Collections.Generic.List[object]]::new()
    $notes   = [System.Collections.Generic.List[string]]::new()
    function Add-Log([string]$Section, [string]$Status, [string]$Detail) { $log.Add([pscustomobject]@{ Section = $Section; Status = $Status; Detail = $Detail }) }
    function Add-Action([string]$Step, [string]$Result, [string]$Detail) {
        $actions.Add([pscustomobject]@{ Time = ConvertTo-UtcText (Get-Date); Step = $Step; Result = $Result; Detail = $Detail })
        $colour = switch -Regex ($Result) { '^(Done|Requested)' { 'Green' } '^(Failed|Not done)' { 'Red' } default { 'Yellow' } }
        Write-Host ("  {0,-16} {1}{2}" -f $Step, $Result, $(if ($Detail) { " - $Detail" })) -ForegroundColor $colour
    }
    $mode = if ($EvidenceOnly) { 'Evidence only (no changes)' } elseif ($WhatIfPreference) { 'WhatIf (no changes)' } else { 'Containment' }
    Write-Host "Mode: $mode"

    $graphConnected = $false; $exoConnected = $false
    $cloudUser = $null; $adUser = $null; $mailbox = $null
    try {
        # --- Connect and find the user -------------------------------------------

        $adAvailable = $false
        $needAd = -not $SkipAD -and ($SkipCloud -or $User -notmatch '@')
        if ($needAd -or -not $SkipAD) {
            if (Get-Module -ListAvailable -Name ActiveDirectory) {
                try { Import-Module ActiveDirectory -ErrorAction Stop; $adAvailable = $true } catch { Add-Log 'Active Directory' 'Unavailable' $_.Exception.Message }
            }
            elseif ($needAd) { Add-Log 'Active Directory' 'Unavailable' 'ActiveDirectory module not installed' }
        }

        $upn = $User
        if ($User -notmatch '@' -and $adAvailable) {
            try { $adUser = Get-ADUser -Identity $User -Properties * -ErrorAction Stop; $upn = $adUser.UserPrincipalName } catch { throw "User '$User' not found in Active Directory: $($_.Exception.Message)" }
        }

        if (-not $SkipCloud) {
            $scopes = if ($EvidenceOnly) { $readScopes } else { $writeScopes }
            $graphConnected = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'User Administrator + Authentication Administrator'
            $select = 'id,displayName,userPrincipalName,mail,accountEnabled,userType,onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesDomainName,createdDateTime,lastPasswordChangeDateTime,signInSessionsValidFromDateTime,assignedLicenses,jobTitle,department'
            try { $cloudUser = Invoke-Graph "v1.0/users/$([uri]::EscapeDataString($upn))?`$select=$select" }
            catch {
                if ($_.Exception.Message -match '404|NotFound|ResourceNotFound|does not exist') { Add-Log 'Entra ID' 'Not found' "No cloud account for $upn" }
                else { throw }
            }
        }
        if (-not $cloudUser -and -not $adUser) { throw "User '$User' was not found." }

        $isSynced = $cloudUser -and $cloudUser.onPremisesSyncEnabled
        if (($isSynced -or -not $cloudUser) -and -not $adUser -and -not $SkipAD) {
            if ($adAvailable) {
                $sam = if ($cloudUser.onPremisesSamAccountName) { $cloudUser.onPremisesSamAccountName } else { ($upn -split '@')[0] }
                try { $adUser = Get-ADUser -Identity $sam -Properties * -ErrorAction Stop } catch { Add-Log 'Active Directory' 'Not found' "${sam}: $($_.Exception.Message)" }
            }
            else {
                Add-Log 'Active Directory' 'Unavailable' 'The user is synced from AD but the ActiveDirectory module is not available - AD steps will be skipped'
            }
        }
        $display = if ($cloudUser) { "$($cloudUser.displayName) <$($cloudUser.userPrincipalName)>" } else { "$($adUser.Name) <$($adUser.UserPrincipalName)>" }
        if ($cloudUser) { $upn = $cloudUser.userPrincipalName }
        Write-Host "User: $display  ($(if ($isSynced) { 'synced from AD' } elseif ($cloudUser) { 'cloud-only' } else { 'AD only' }))"

        if (-not $SkipExchange -and $cloudUser) {
            try {
                $exoConnected = Connect-ToolkitExchange -Config $config -Bound $PSBoundParameters -RoleHint 'Exchange Administrator'
                $mailbox = Get-EXOMailbox -Identity $upn -Properties ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward -ErrorAction Stop
            }
            catch { Add-Log 'Exchange Online' 'Unavailable' $_.Exception.Message }
        }

        $folderName = (@($CaseId, 'Containment', (($upn -split '@')[0]), (Get-Date -Format 'yyyyMMdd_HHmmss')) | Where-Object { $_ }) -join '_'
        $folder = New-EvidenceFolder -Config $config -OutputPath $OutputPath -Name $folderName

        # --- 1. Evidence (before any change) --------------------------------------

        Write-Host 'Collecting evidence...'
        $evidence = [ordered]@{}
        function Save-Evidence([string]$Name, [scriptblock]$Block) {
            try {
                $rows = @(& $Block)
                $evidence[$Name] = $rows
                $count = Export-EvidenceCsv -Folder $folder -Name $Name -Rows $rows
                Add-Log $Name 'OK' "$count row(s)"
            }
            catch {
                $evidence[$Name] = @()
                Add-Log $Name 'Failed' $_.Exception.Message
            }
        }

        Save-Evidence '01_Account' {
            [pscustomobject]@{
                DisplayName = if ($cloudUser) { $cloudUser.displayName } else { $adUser.Name }
                UserPrincipalName = $upn; Mail = $cloudUser.mail; CloudObjectId = $cloudUser.id; UserType = $cloudUser.userType
                CloudEnabled = $cloudUser.accountEnabled; SyncedFromAD = [bool]$isSynced; Created = ConvertTo-UtcText $cloudUser.createdDateTime
                CloudPasswordLastChanged = ConvertTo-UtcText $cloudUser.lastPasswordChangeDateTime; SessionsValidFrom = ConvertTo-UtcText $cloudUser.signInSessionsValidFromDateTime
                Licensed = @($cloudUser.assignedLicenses).Count -gt 0; JobTitle = $cloudUser.jobTitle; Department = $cloudUser.department
                AdSamAccountName = $adUser.SamAccountName; AdDistinguishedName = $adUser.DistinguishedName; AdEnabled = $adUser.Enabled
                AdLockedOut = $adUser.LockedOut; AdPasswordLastSet = ConvertTo-UtcText $adUser.PasswordLastSet; AdLastLogon = ConvertTo-UtcText $adUser.LastLogonDate
                AdBadLogonCount = $adUser.badPwdCount; AdLastBadPassword = ConvertTo-UtcText $adUser.LastBadPasswordAttempt
            }
        }

        if ($cloudUser) {
            $id = $cloudUser.id
            Save-Evidence '02_Groups' {
                Get-GraphAll "v1.0/users/$id/memberOf?`$select=id,displayName" | ForEach-Object {
                    [pscustomobject]@{ Source = 'Entra ID'; Type = ("$($_.'@odata.type')" -replace '^#microsoft\.graph\.', ''); Name = $_.displayName; Id = $_.id }
                }
                foreach ($g in @($adUser.MemberOf | Where-Object { $_ })) { [pscustomobject]@{ Source = 'Active Directory'; Type = 'group'; Name = ($g -replace '^CN=([^,]+),.*$', '$1'); Id = $g } }
            }
            Save-Evidence '03_MfaMethods' {
                Get-GraphAll "v1.0/users/$id/authentication/methods" | ForEach-Object {
                    $type = "$($_.'@odata.type')" -replace '^#microsoft\.graph\.', ''
                    $name = ConvertFrom-MethodObject $_
                    if ($type -eq 'passwordAuthenticationMethod') { $name = 'Password' }
                    [pscustomobject]@{
                        Method  = if ($methodNames.ContainsKey("$name")) { $methodNames["$name"] } else { $name }
                        Detail  = (@($_.phoneNumber, $_.emailAddress, $_.displayName, $_.deviceTag, $_.model) | Where-Object { $_ }) -join ' / '
                        Created = ConvertTo-UtcText $_.createdDateTime
                        Type    = $type
                        Id      = $_.id
                    }
                }
            }
            Save-Evidence '04_SignIns' {
                $uri = "v1.0/auditLogs/signIns?`$filter=userId eq '$id' and createdDateTime ge $sinceIso&`$top=500"
                $n = 0
                while ($uri -and $n -lt 5000) {
                    $page = Invoke-Graph $uri
                    foreach ($s in @($page.value)) {
                        $n++
                        $ok = [int]"0$($s.status.errorCode)" -eq 0
                        [pscustomobject]@{
                            Time = ConvertTo-UtcText $s.createdDateTime; Result = if ($ok) { 'Success' } else { "Failed: $($s.status.failureReason) ($($s.status.errorCode))" }
                            App = $s.appDisplayName; Resource = $s.resourceDisplayName; IP = $s.ipAddress
                            City = $s.location.city; Country = $s.location.countryOrRegion; ClientApp = $s.clientAppUsed
                            Device = (@($s.deviceDetail.operatingSystem, $s.deviceDetail.browser) | Where-Object { $_ }) -join ' / '
                            MfaRequirement = $s.authenticationRequirement; ConditionalAccess = $s.conditionalAccessStatus
                            Risk = $s.riskLevelDuringSignIn; CorrelationId = $s.correlationId
                        }
                    }
                    $uri = $page.'@odata.nextLink'
                }
            }
            Save-Evidence '05_SignInLocations' {
                $evidence['04_SignIns'] | Group-Object IP | ForEach-Object {
                    $g = @($_.Group | Sort-Object Time)
                    [pscustomobject]@{
                        IP = $_.Name; Country = $g[0].Country; City = $g[0].City; SignIns = $g.Count
                        Successes = @($g | Where-Object Result -eq 'Success').Count; Failures = @($g | Where-Object Result -ne 'Success').Count
                        FirstSeen = $g[0].Time; LastSeen = $g[-1].Time; Apps = (@($g.App | Select-Object -Unique) -join ', ')
                    }
                } | Sort-Object SignIns -Descending
            }
            Save-Evidence '06_DirectoryAudits' {
                $rows = @()
                foreach ($q in @(
                    @{ Direction = 'Changes to this user'; Filter = "activityDateTime ge $sinceIso and targetResources/any(t:t/id eq '$id')" },
                    @{ Direction = 'Actions by this user'; Filter = "activityDateTime ge $sinceIso and initiatedBy/user/id eq '$id'" })) {
                    foreach ($a in (Get-GraphAll "v1.0/auditLogs/directoryAudits?`$filter=$($q.Filter)")) {
                        $rows += [pscustomobject]@{
                            Time = ConvertTo-UtcText $a.activityDateTime; Direction = $q.Direction; Activity = $a.activityDisplayName; Category = $a.category; Result = $a.result
                            InitiatedBy = (@($a.initiatedBy.user.userPrincipalName, $a.initiatedBy.app.displayName) | Where-Object { $_ }) -join ''
                            Targets = (@($a.targetResources | ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }) -join '; ')
                            Changes = (@($a.targetResources | ForEach-Object { $_.modifiedProperties } | Where-Object { $_ } | ForEach-Object { $_.displayName }) | Select-Object -Unique) -join ', '
                        }
                    }
                }
                $rows | Sort-Object Time
            }
            Save-Evidence '07_Devices' {
                $seen = @{}
                foreach ($kind in 'registeredDevices', 'ownedDevices') {
                    foreach ($d in (Get-GraphAll "v1.0/users/$id/$kind")) {
                        if ($seen.ContainsKey($d.id)) { continue }; $seen[$d.id] = $true
                        [pscustomobject]@{
                            Name = $d.displayName; OS = (@($d.operatingSystem, $d.operatingSystemVersion) | Where-Object { $_ }) -join ' '; TrustType = $d.trustType
                            Registered = ConvertTo-UtcText $d.registrationDateTime; LastSignIn = ConvertTo-UtcText $d.approximateLastSignInDateTime
                            Compliant = $d.isCompliant; Managed = $d.isManaged; Enabled = $d.accountEnabled; DeviceId = $d.deviceId
                        }
                    }
                }
            }
            Save-Evidence '08_AppConsents' {
                $names = @{}
                Get-GraphAll "v1.0/oauth2PermissionGrants?`$filter=principalId eq '$id'" | Where-Object { $_.principalId -eq $id } | ForEach-Object {
                    if (-not $names.ContainsKey($_.clientId)) {
                        try { $names[$_.clientId] = (Invoke-Graph "v1.0/servicePrincipals/$($_.clientId)?`$select=displayName,appOwnerOrganizationId").displayName } catch { $names[$_.clientId] = $_.clientId }
                    }
                    [pscustomobject]@{ App = $names[$_.clientId]; Permissions = $_.scope; ConsentType = $_.consentType; ClientId = $_.clientId }
                }
            }
        }

        $internal = @()
        if ($mailbox) {
            try { $internal = @(Get-AcceptedDomain | ForEach-Object { "$($_.DomainName)".ToLower() }) } catch { Add-Log 'Accepted domains' 'Failed' $_.Exception.Message }
            Save-Evidence '09_MailboxForwarding' {
                [pscustomobject]@{ Mailbox = "$($mailbox.PrimarySmtpAddress)"; ForwardingSmtpAddress = "$($mailbox.ForwardingSmtpAddress)"; ForwardingAddress = "$($mailbox.ForwardingAddress)"; DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward }
            }
            Save-Evidence '10_InboxRules' {
                Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction Stop -WarningAction SilentlyContinue | ForEach-Object {
                    $a = Get-InboxRuleAssessment $_ $internal
                    [pscustomobject]@{
                        Risk = $a.Risk; Rule = "$($_.Name)"; Enabled = [bool]$_.Enabled; Actions = $a.Actions; ForwardsTo = $a.Targets; Conditions = $a.Conditions
                        Issues = ($a.Issues -join '; '); Description = "$($_.Description)" -replace '\s+', ' '; Identity = "$($_.Identity)"
                    }
                }
            }
        }

        # --- 2. Containment ----------------------------------------------------------

        $isAppOnly = (Get-Command Get-MgContext -ErrorAction SilentlyContinue) -and "$((Get-MgContext).AuthType)" -eq 'AppOnly'
        function Invoke-Step {
            # Uses the script's own -WhatIf/-Confirm handling ($PSCmdlet of the script)
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
            param([string]$Step, [string]$Target, [string]$What, [scriptblock]$Do)
            if ($PSCmdlet.ShouldProcess($Target, $What)) {
                try { $detail = & $Do; Add-Action $Step $(if ($detail -is [array]) { $detail[0] } else { 'Done' }) $(if ($detail -is [array]) { $detail[1] } else { "$detail" }) }
                catch { Add-Action $Step 'Failed' $_.Exception.Message }
            }
            elseif ($WhatIfPreference) { Add-Action $Step 'WhatIf' "Would: $What" }
            else { Add-Action $Step 'Skipped' 'Not confirmed' }
        }

        if ($EvidenceOnly) {
            Write-Host 'Evidence only - no changes made.'
        }
        else {
            Write-Host 'Containing account...'

            # Disable
            if ($Skip -contains 'Disable') { Add-Action 'Disable' 'Skipped' '-Skip' }
            elseif ($isSynced -or -not $cloudUser) {
                if ($adUser) {
                    Invoke-Step 'Disable' "$($adUser.SamAccountName) (Active Directory)" 'Disable the account in Active Directory' {
                        Disable-ADAccount -Identity $adUser.DistinguishedName -Confirm:$false -ErrorAction Stop
                        if ($isSynced) { 'Disabled in AD. Entra ID follows at the next sync - run Start-ADSyncSyncCycle -PolicyType Delta on the Entra Connect server to apply it now.' }
                        elseif ($SkipCloud) { 'Disabled in AD. If the account syncs to Entra ID, run a delta sync so the change applies there now.' }
                        else { 'Disabled in AD.' }
                    }
                }
                else { Add-Action 'Disable' 'Not done' 'The account is synced from AD but AD is not available here - disable it in AD now.' }
            }
            else {
                Invoke-Step 'Disable' $upn 'Block sign-in (disable the account) in Entra ID' {
                    Invoke-GraphChange PATCH "v1.0/users/$($cloudUser.id)" @{ accountEnabled = $false } | Out-Null
                    'Sign-in blocked in Entra ID.'
                }
            }

            # Reset password
            if ($Skip -contains 'ResetPassword') { Add-Action 'ResetPassword' 'Skipped' '-Skip' }
            elseif (($isSynced -or -not $cloudUser) -and $adUser) {
                Invoke-Step 'ResetPassword' "$($adUser.SamAccountName) (Active Directory)" 'Reset the password to a random value in Active Directory' {
                    $secure = New-Object System.Security.SecureString
                    foreach ($ch in (New-RandomPassword).ToCharArray()) { $secure.AppendChar($ch) }
                    $secure.MakeReadOnly()
                    Set-ADAccountPassword -Identity $adUser.DistinguishedName -Reset -NewPassword $secure -Confirm:$false -ErrorAction Stop
                    if ($isSynced -or $SkipCloud) { 'Random password set in AD (not shown or saved). It syncs to Entra ID if password hash sync is on.' } else { 'Random password set in AD (not shown or saved).' }
                }
            }
            elseif ($isSynced) { Add-Action 'ResetPassword' 'Not done' 'The account is synced from AD but AD is not available here - reset the password in AD now.' }
            elseif ($isAppOnly) { Add-Action 'ResetPassword' 'Not done' 'Password reset needs a signed-in admin (it is not available with app-only sign-in).' }
            else {
                Invoke-Step 'ResetPassword' $upn 'Reset the password to a random value in Entra ID' {
                    Invoke-GraphChange POST "v1.0/users/$($cloudUser.id)/authentication/methods/$passwordMethodId/resetPassword" @{ newPassword = (New-RandomPassword) } | Out-Null
                    , @('Requested', 'Random password set (not shown or saved). Microsoft applies it within a minute or so.')
                }
            }

            # Revoke sessions
            if ($Skip -contains 'RevokeSessions') { Add-Action 'RevokeSessions' 'Skipped' '-Skip' }
            elseif (-not $cloudUser) { Add-Action 'RevokeSessions' 'Skipped' 'No cloud account' }
            else {
                Invoke-Step 'RevokeSessions' $upn 'Revoke all sign-in sessions and refresh tokens' {
                    Invoke-GraphChange POST "v1.0/users/$($cloudUser.id)/revokeSignInSessions" $null | Out-Null
                    'All sessions and refresh tokens revoked (can take a few minutes to take effect).'
                }
            }

            # Mailbox
            if ($Skip -contains 'Mailbox') { Add-Action 'Mailbox' 'Skipped' '-Skip' }
            elseif (-not $mailbox) { Add-Action 'Mailbox' 'Skipped' $(if ($SkipExchange) { '-SkipExchange' } else { 'No mailbox found, or Exchange Online not connected (see 00_CollectionLog.csv)' }) }
            else {
                $fwd = (@(("$($mailbox.ForwardingSmtpAddress)" -replace '^(?i)smtp:', ''), "$($mailbox.ForwardingAddress)") | Where-Object { $_ }) -join ', '
                if ($fwd) {
                    Invoke-Step 'Mailbox' $upn "Remove mailbox forwarding to $fwd" {
                        Set-Mailbox -Identity $upn -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false -Confirm:$false -ErrorAction Stop
                        "Removed forwarding to $fwd"
                    }
                }
                $badRules = @($evidence['10_InboxRules'] | Where-Object { $_.Risk -in 'Critical', 'High', 'Medium' -and $_.Enabled })
                foreach ($r in $badRules) {
                    Invoke-Step 'Mailbox' "$upn rule '$($r.Rule)'" "Disable inbox rule '$($r.Rule)' ($($r.Issues))" {
                        Disable-InboxRule -Identity $r.Identity -Confirm:$false -ErrorAction Stop
                        "Disabled inbox rule '$($r.Rule)' (kept, not deleted, for evidence)"
                    }
                }
                if (-not $fwd -and -not $badRules.Count) { Add-Action 'Mailbox' 'Done' 'No forwarding or suspicious enabled inbox rules found' }
            }

            # MFA methods (only when asked)
            if ($PSBoundParameters.ContainsKey('RemoveMfaMethodsAddedAfter') -and $cloudUser) {
                $cutoffText = ConvertTo-UtcText $RemoveMfaMethodsAddedAfter
                $recent = @($evidence['03_MfaMethods'] | Where-Object { $_.Created -and $_.Created -gt $cutoffText -and $methodSegments.ContainsKey($_.Type) })
                if (-not $recent.Count) { Add-Action 'MfaMethods' 'Done' "No MFA methods added after $cutoffText" }
                foreach ($m in $recent) {
                    Invoke-Step 'MfaMethods' $upn "Remove MFA method $($m.Method) ($($m.Detail)) added $($m.Created)" {
                        Invoke-GraphChange DELETE "v1.0/users/$($cloudUser.id)/authentication/$($methodSegments[$m.Type])/$($m.Id)" $null | Out-Null
                        "Removed $($m.Method) ($($m.Detail)) added $($m.Created)"
                    }
                }
                $undated = @($evidence['03_MfaMethods'] | Where-Object { -not $_.Created -and $_.Method -ne 'Password' })
                if ($undated.Count) { $notes.Add("These MFA methods have no creation date, so they were not checked: $(($undated | ForEach-Object { "$($_.Method) $($_.Detail)" }) -join '; '). Confirm them with the user.") }
            }
        }

        # --- 3. Indicators and next steps -------------------------------------------

        $indicators = [System.Collections.Generic.List[object]]::new()
        $si = @($evidence['04_SignIns'])
        $countries = @($si | Where-Object { $_.Result -eq 'Success' -and $_.Country } | Select-Object -ExpandProperty Country -Unique)
        if ($countries.Count -gt 1) { $indicators.Add([pscustomobject]@{ Indicator = 'Successful sign-ins from more than one country'; Detail = ($countries -join ', ') }) }
        $failed = @($si | Where-Object { $_.Result -ne 'Success' })
        if ($failed.Count -ge 10) { $indicators.Add([pscustomobject]@{ Indicator = "$($failed.Count) failed sign-ins in $SignInDays days"; Detail = (@($failed | Select-Object -ExpandProperty IP -Unique | Select-Object -First 10) -join ', ') }) }
        $legacy = @($si | Where-Object { $_.ClientApp -match 'IMAP|POP|SMTP|Exchange ActiveSync|Other clients|Authenticated SMTP' -and $_.Result -eq 'Success' })
        if ($legacy.Count) { $indicators.Add([pscustomobject]@{ Indicator = 'Successful legacy-protocol sign-ins (no MFA possible)'; Detail = (@($legacy.ClientApp | Select-Object -Unique) -join ', ') }) }
        $recentMfa = @($evidence['03_MfaMethods'] | Where-Object { $_.Created -and $_.Created -ge (ConvertTo-UtcText $since) })
        if ($recentMfa.Count) { $indicators.Add([pscustomobject]@{ Indicator = "MFA method(s) added in the last $SignInDays days"; Detail = (($recentMfa | ForEach-Object { "$($_.Method) $($_.Detail) ($($_.Created))" }) -join '; ') }) }
        $badRuleRows = @($evidence['10_InboxRules'] | Where-Object { $_.Risk -in 'Critical', 'High', 'Medium' })
        if ($badRuleRows.Count) { $indicators.Add([pscustomobject]@{ Indicator = 'Suspicious inbox rules'; Detail = (($badRuleRows | ForEach-Object { "'$($_.Rule)': $($_.Issues -replace ' - common.*$', '')" }) -join '; ') }) }
        $fwdRow = @($evidence['09_MailboxForwarding'])[0]
        if ($fwdRow -and ($fwdRow.ForwardingSmtpAddress -or $fwdRow.ForwardingAddress)) { $indicators.Add([pscustomobject]@{ Indicator = 'Mailbox forwarding was set'; Detail = ("$($fwdRow.ForwardingSmtpAddress) $($fwdRow.ForwardingAddress)" -replace '(?i)smtp:', '').Trim() }) }
        $userConsents = @($evidence['08_AppConsents'] | Where-Object { $_.ConsentType -eq 'Principal' })
        if ($userConsents.Count) { $indicators.Add([pscustomobject]@{ Indicator = 'Apps the user consented to'; Detail = (@($userConsents.App | Select-Object -Unique) -join ', ') }) }
        $byUser = @($evidence['06_DirectoryAudits'] | Where-Object { $_.Direction -eq 'Actions by this user' -and $_.Activity -match '(?i)security info|consent|add member|add owner|add app|reset password|update application' })
        if ($byUser.Count) { $indicators.Add([pscustomobject]@{ Indicator = 'Sensitive changes made by this account'; Detail = (@($byUser | ForEach-Object { "$($_.Time) $($_.Activity)" }) | Select-Object -First 10) -join '; ' }) }

        $nextSteps = @(
            'Review 04_SignIns.csv and 05_SignInLocations.csv for attacker IPs. Block them (Conditional Access named location) and search other accounts for the same IPs.'
            'Check 03_MfaMethods.csv with the user - remove any method they did not add (or re-run with -RemoveMfaMethodsAddedAfter <date>).'
            'Review 08_AppConsents.csv and revoke any app the user does not recognise.'
            'Check whether the account sent phishing or read sensitive mail: message trace and the unified audit log (Purview) for the same period.'
            'Review 07_Devices.csv for devices the user does not recognise and disable them.'
        )
        if ($isSynced) { $nextSteps += 'The account is synced from AD: run a delta sync (Start-ADSyncSyncCycle -PolicyType Delta) so the disable reaches Entra ID now, and check the AD account for other signs of compromise.' }
        $nextSteps += 'To restore access: verify the user out-of-band, have them set a new password, require MFA re-registration, re-enable the account and re-enable any inbox rules that were legitimate.'
        $nextSteps += @($notes)

        [void](Export-EvidenceCsv -Folder $folder -Name '00_Actions' -Rows $actions)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_Indicators' -Rows $indicators)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_CollectionLog' -Rows $log)

        $runBy = if ($graphConnected -or $cloudUser) { Get-GraphRunBy } else { "$env:USERDOMAIN\$env:USERNAME" }
        $html = New-HtmlReport -Title "Account containment - $upn" -IntroLines @(
            "User: $display - $(if ($isSynced) { 'synced from AD' } elseif ($cloudUser) { 'cloud-only' } else { 'AD only' })",
            "Mode: $mode. Run $(ConvertTo-UtcText $start) by $runBy$(if ($CaseId) { " - case $CaseId" }). All times are UTC.",
            "Evidence covers the last $SignInDays day(s)."
        ) -Sections @(
            @{ Heading = 'Actions'; Html = ConvertTo-HtmlTable -Rows $actions -Columns @('Time', 'Step', 'Result', 'Detail') -StatusColumn 'Result' -EmptyText 'No changes made.' }
            @{ Heading = "Indicators ($($indicators.Count))"; Html = ConvertTo-HtmlTable -Rows $indicators -Columns @('Indicator', 'Detail') -EmptyText 'No automatic indicators. Review the evidence files.' }
            @{ Heading = 'Sign-in locations'; Html = ConvertTo-HtmlTable -Rows @($evidence['05_SignInLocations'] | Select-Object -First 20) -Columns @('IP', 'Country', 'City', 'SignIns', 'Successes', 'Failures', 'LastSeen', 'Apps') -EmptyText 'No sign-in data (needs Entra ID P1, and AuditLog.Read.All).' }
            @{ Heading = 'Next steps'; Html = ConvertTo-HtmlTable -Rows @($nextSteps | ForEach-Object { [pscustomobject]@{ Step = $_ } }) -Columns @('Step') }
            @{ Heading = 'Evidence collected'; Html = ConvertTo-HtmlTable -Rows $log -Columns @('Section', 'Status', 'Detail') -StatusColumn 'Status' }
        )
        [System.IO.File]::WriteAllText((Join-Path $folder '00_Summary.html'), $html, $script:Utf8NoBom)
        $package = Complete-EvidencePackage -Folder $folder

        Write-Host ""
        if ($indicators.Count) {
            Write-Host 'Indicators:' -ForegroundColor Yellow
            foreach ($i in $indicators) { Write-Host "  - $($i.Indicator): $($i.Detail)" -ForegroundColor Yellow }
        }
        $problems = @($log | Where-Object { $_.Status -ne 'OK' })
        if ($problems.Count) { Write-Host ("Some evidence could not be collected: {0}" -f (($problems | ForEach-Object { "$($_.Section) ($($_.Status))" }) -join ', ')) -ForegroundColor Yellow }
        Write-Host "Evidence : $($package.Zip)"
        Write-Host "SHA-256  : $($package.Sha256)  (record this in your case notes)"
        Write-Host "Summary  : $(Join-Path $folder '00_Summary.html')"
    }
    finally {
        if ($exoConnected)   { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
        if ($graphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
