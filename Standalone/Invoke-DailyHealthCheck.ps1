<#
.SYNOPSIS
    Daily health check for Windows servers, saved as an HTML report.
.DESCRIPTION
    Checks uptime, pending reboots, disk space, stopped services, last update and event log errors.
.EXAMPLE
    .\Invoke-DailyHealthCheck.ps1 -ComputerName SERVER01, SERVER02
.EXAMPLE
    .\Invoke-DailyHealthCheck.ps1 -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: admin rights and WinRM (or DCOM) on the servers
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$ComputerName,

    [switch]$SendEmail,
    [switch]$OnlyEmailOnIssues,

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

if (-not $PSBoundParameters.ContainsKey('ComputerName')) {
    $ComputerName = Get-Setting $config 'Servers' 'ComputerName' @([System.Environment]::MachineName)
}

$ComputerName = Split-ListValue $ComputerName

$thresholds = @{
    DiskWarning     = [double](Get-Setting $config 'HealthCheck' 'DiskWarningPercent' 15)
    DiskCritical    = [double](Get-Setting $config 'HealthCheck' 'DiskCriticalPercent' 5)
    UptimeDays      = [int](Get-Setting $config 'HealthCheck' 'UptimeWarningDays' 45)
    PatchDays       = [int](Get-Setting $config 'HealthCheck' 'PatchWarningDays' 35)
    ErrorEvents     = [int](Get-Setting $config 'HealthCheck' 'ErrorEventWarning' 50)
    CheckFileRename = [bool](Get-Setting $config 'HealthCheck' 'CheckPendingFileRename' $false)
    ExcludeServices = @(Get-Setting $config 'HealthCheck' 'ExcludeServices' @('gupdate', 'gupdatem', 'edgeupdate', 'edgeupdatem', 'MapsBroker', 'RemoteRegistry', 'sppsvc', 'CDPSvc', 'OneSyncSvc*', 'tiledatamodelsvc', 'WbioSrvc', 'DoSvc', 'UsoSvc', 'BITS', 'ShellHWDetection', 'TrustedInstaller', 'wuauserv'))
}

# --- Checks ---------------------------------------------------------------

function Test-IsLocalComputer([string]$Name) {
    $n = $Name.Split('.')[0]
    return ($n -eq 'localhost' -or $n -eq '.' -or $n -eq [System.Environment]::MachineName)
}

function New-HealthCimSession([string]$Name) {
    try {
        return New-CimSession -ComputerName $Name -OperationTimeoutSec 30 -ErrorAction Stop
    }
    catch {
        $dcom = New-CimSessionOption -Protocol Dcom
        return New-CimSession -ComputerName $Name -SessionOption $dcom -OperationTimeoutSec 30 -ErrorAction Stop
    }
}

function Test-RegistryKey($Session, [string]$SubKey) {
    $r = Invoke-CimMethod -CimSession $Session -Namespace 'root/cimv2' -ClassName 'StdRegProv' -MethodName 'EnumKey' `
         -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $SubKey } -ErrorAction Stop   # HKLM
    return ($r.ReturnValue -eq 0)
}

function Get-PendingRebootReasons($Session) {
    $reasons = @()
    if (Test-RegistryKey $Session 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons += 'Component servicing' }
    if (Test-RegistryKey $Session 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')  { $reasons += 'Windows Update' }
    if ($thresholds.CheckFileRename) {
        $r = Invoke-CimMethod -CimSession $Session -Namespace 'root/cimv2' -ClassName 'StdRegProv' -MethodName 'GetMultiStringValue' `
             -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = 'SYSTEM\CurrentControlSet\Control\Session Manager'; sValueName = 'PendingFileRenameOperations' } -ErrorAction Stop
        if ($r.ReturnValue -eq 0 -and $r.sValue) { $reasons += 'File rename operations' }
    }
    return $reasons
}

function ConvertTo-HotfixDate([string]$Value) {
    if (-not $Value) { return $null }
    $formats = @('M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd', 'd/M/yyyy', 'dd/MM/yyyy')
    foreach ($f in $formats) {
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($Value, $f, [System.Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$d)) { return $d }
    }
    return $null
}

function Get-ServerHealth([string]$Name) {
    $h = [pscustomobject]@{
        Server          = $Name
        Status          = 'OK'
        OS              = ''
        LastBoot        = $null
        UptimeDays      = $null
        PendingReboot   = ''
        LastPatch       = $null
        DaysSincePatch  = $null
        LowDisks        = ''
        StoppedServices = ''
        Errors24h       = ''
        TopErrorSources = ''
        Issues          = [System.Collections.Generic.List[string]]::new()
        Disks           = @()
        Services        = @()
    }
    $raise = {
        param([string]$Level, [string]$Text)
        $h.Issues.Add($Text)
        if ($Level -eq 'Critical') { $h.Status = 'Critical' }
        elseif ($Level -eq 'Warning' -and $h.Status -eq 'OK') { $h.Status = 'Warning' }
    }

    $isLocal = Test-IsLocalComputer $Name
    if (-not $isLocal -and -not (Test-Connection -ComputerName $Name -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        # Ping can be blocked - only give up if CIM fails too
        $pingFailed = $true
    } else { $pingFailed = $false }

    try {
        $session = New-HealthCimSession $Name
    }
    catch {
        $h.Status = 'Critical'
        $h.Issues.Add($(if ($pingFailed) { 'Unreachable (no ping, no remote management)' } else { "Remote management failed: $($_.Exception.Message)" }))
        return $h
    }

    try {
        # OS and uptime
        try {
            $os = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem -ErrorAction Stop
            $h.OS         = $os.Caption -replace 'Microsoft ', ''
            $h.LastBoot   = $os.LastBootUpTime
            $h.UptimeDays = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)
            if ($h.UptimeDays -gt $thresholds.UptimeDays) { & $raise 'Warning' "Not rebooted for $([int]$h.UptimeDays) days" }
        } catch { & $raise 'Warning' "OS check failed: $($_.Exception.Message)" }

        # Pending reboot
        try {
            $reasons = @(Get-PendingRebootReasons $session)
            $h.PendingReboot = if ($reasons.Count) { 'Yes (' + ($reasons -join ', ') + ')' } else { 'No' }
            if ($reasons.Count) { & $raise 'Warning' "Reboot pending: $($reasons -join ', ')" }
        } catch { $h.PendingReboot = 'Unknown' }

        # Disks
        try {
            $disks = @(Get-CimInstance -CimSession $session -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)
            $h.Disks = @(foreach ($d in $disks) {
                if (-not $d.Size) { continue }
                $pct = [math]::Round(($d.FreeSpace / $d.Size) * 100, 1)
                $st  = if ($pct -lt $thresholds.DiskCritical) { 'Critical' } elseif ($pct -lt $thresholds.DiskWarning) { 'Warning' } else { 'OK' }
                [pscustomobject]@{
                    Server = $Name; Drive = $d.DeviceID; Label = $d.VolumeName
                    SizeGB = [math]::Round($d.Size / 1GB, 1); FreeGB = [math]::Round($d.FreeSpace / 1GB, 1)
                    FreePercent = $pct; Status = $st
                }
            })
            $low = @($h.Disks | Where-Object { $_.Status -ne 'OK' })
            $h.LowDisks = if ($low.Count) { ($low | ForEach-Object { "$($_.Drive) $($_.FreePercent)%" }) -join ', ' } else { 'None' }
            foreach ($d in $low) { & $raise $d.Status "Disk $($d.Drive) has $($d.FreePercent)% free ($($d.FreeGB) GB)" }
        } catch { & $raise 'Warning' "Disk check failed: $($_.Exception.Message)" }

        # Automatic services not running
        try {
            $svc = @(Get-CimInstance -CimSession $session -ClassName Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" -ErrorAction Stop)
            $svc = @($svc | Where-Object {
                $n = $_.Name
                -not ($thresholds.ExcludeServices | Where-Object { $n -like $_ })
            })
            $h.Services = @($svc | ForEach-Object {
                [pscustomobject]@{ Server = $Name; Name = $_.Name; DisplayName = $_.DisplayName; State = $_.State; ExitCode = $_.ExitCode }
            })
            $h.StoppedServices = if ($svc.Count) { ($svc.Name -join ', ') } else { 'None' }
            if ($svc.Count) { & $raise 'Warning' "$($svc.Count) automatic service(s) not running: $($svc.Name -join ', ')" }
        } catch { & $raise 'Warning' "Service check failed: $($_.Exception.Message)" }

        # Last update installed
        try {
            $qfe = @(Get-CimInstance -CimSession $session -ClassName Win32_QuickFixEngineering -ErrorAction Stop)
            $dates = @($qfe | ForEach-Object { ConvertTo-HotfixDate "$($_.InstalledOn)" } | Where-Object { $_ })
            if ($dates.Count) {
                $h.LastPatch      = ($dates | Sort-Object -Descending | Select-Object -First 1)
                $h.DaysSincePatch = [int]((Get-Date).Date - $h.LastPatch.Date).TotalDays
                if ($h.DaysSincePatch -gt $thresholds.PatchDays) { & $raise 'Warning' "No update installed for $($h.DaysSincePatch) days" }
            }
        } catch { Write-Verbose "Update history unavailable on ${Name}: $($_.Exception.Message)" }
    }
    finally {
        if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
    }

    # Event log errors (critical + error) in the last 24 hours
    try {
        $filter = @{ LogName = 'System', 'Application'; Level = 1, 2; StartTime = (Get-Date).AddHours(-24) }
        $evParams = @{ FilterHashtable = $filter; ErrorAction = 'Stop'; MaxEvents = 5000 }
        if (-not $isLocal) { $evParams.ComputerName = $Name }
        $events = @(Get-WinEvent @evParams)
        $h.Errors24h = $events.Count
        $h.TopErrorSources = ($events | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 3 |
                              ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
        if ($events.Count -gt $thresholds.ErrorEvents) { & $raise 'Warning' "$($events.Count) errors in the System/Application logs in 24h" }
        if (@($events | Where-Object { $_.Level -eq 1 }).Count) { & $raise 'Warning' "$(@($events | Where-Object { $_.Level -eq 1 }).Count) critical event(s) in 24h" }
    }
    catch {
        if ($_.Exception.Message -match 'No events were found') { $h.Errors24h = 0; $h.TopErrorSources = '' }
        else { $h.Errors24h = 'Unavailable' }
    }

    return $h
}

# --- Report ---------------------------------------------------------------

function New-HealthReportHtml([object[]]$Results) {
    $critical = @($Results | Where-Object { $_.Status -eq 'Critical' }).Count
    $warning  = @($Results | Where-Object { $_.Status -eq 'Warning' }).Count
    $ok       = @($Results | Where-Object { $_.Status -eq 'OK' }).Count

    $summaryRows = $Results | Sort-Object @{ Expression = { switch ($_.Status) { 'Critical' { 0 } 'Warning' { 1 } default { 2 } } } }, Server |
        ForEach-Object {
            [pscustomobject]@{
                Server            = $_.Server
                Status            = $_.Status
                OS                = $_.OS
                'Uptime (days)'   = $_.UptimeDays
                'Pending reboot'  = $_.PendingReboot
                'Last update'     = if ($_.LastPatch) { $_.LastPatch.ToString('yyyy-MM-dd') } else { '' }
                'Low disks'       = $_.LowDisks
                'Errors (24h)'    = $_.Errors24h
            }
        }

    $issueRows = foreach ($r in $Results) {
        foreach ($i in $r.Issues) { [pscustomobject]@{ Server = $r.Server; Status = $r.Status; Issue = $i } }
    }
    $diskRows = @($Results | ForEach-Object { $_.Disks }) | Sort-Object FreePercent
    $svcRows  = @($Results | ForEach-Object { $_.Services })
    $evRows   = $Results | Where-Object { $_.TopErrorSources } |
                ForEach-Object { [pscustomobject]@{ Server = $_.Server; 'Errors (24h)' = $_.Errors24h; 'Top sources' = $_.TopErrorSources } }

    $intro = @(
        ("Thresholds: disk warning below {0}% free, critical below {1}%; uptime over {2} days; no update for over {3} days." -f
            $thresholds.DiskWarning, $thresholds.DiskCritical, $thresholds.UptimeDays, $thresholds.PatchDays)
    )
    $sections = @(
        @{ Heading = 'Overview'; Html = ConvertTo-HtmlTable -Rows @($summaryRows) -Columns @('Server', 'Status', 'OS', 'Uptime (days)', 'Pending reboot', 'Last update', 'Low disks', 'Errors (24h)') -StatusColumn 'Status' }
        @{ Heading = 'Issues'; Html = ConvertTo-HtmlTable -Rows @($issueRows) -Columns @('Server', 'Status', 'Issue') -StatusColumn 'Status' -EmptyText 'No issues found.' }
        @{ Heading = 'Disks'; Html = ConvertTo-HtmlTable -Rows @($diskRows) -Columns @('Server', 'Drive', 'Label', 'SizeGB', 'FreeGB', 'FreePercent', 'Status') -StatusColumn 'Status' }
        @{ Heading = 'Automatic services not running'; Html = ConvertTo-HtmlTable -Rows @($svcRows) -Columns @('Server', 'Name', 'DisplayName', 'State', 'ExitCode') -EmptyText 'All automatic services are running.' }
        @{ Heading = 'Event log errors (last 24 hours)'; Html = ConvertTo-HtmlTable -Rows @($evRows) -Columns @('Server', 'Errors (24h)', 'Top sources') -EmptyText 'No errors logged.' }
    )
    $title = "Daily health check - $(Get-Date -Format 'dddd d MMMM yyyy')"
    return @{
        Html     = New-HtmlReport -Title $title -IntroLines $intro -Sections $sections -Stats @(
                       @{ Label = 'Servers checked'; Value = $Results.Count }
                       @{ Label = 'Critical'; Value = $critical; Status = 'Critical' }
                       @{ Label = 'Warning'; Value = $warning; Status = 'Warning' }
                       @{ Label = 'OK'; Value = $ok; Status = 'OK' }
                   )
        Critical = $critical
        Warning  = $warning
        OK       = $ok
    }
}

# --- Main -----------------------------------------------------------------
# (Skipped when the file is dot-sourced, so the functions above can be reused or tested.)

if ($MyInvocation.InvocationName -ne '.') {
    $reportPath = New-ReportPath -Config $config -Prefix 'HealthCheck' -Extension '.html' -OutputFile $OutputFile

    $results = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($server in $ComputerName) {
        $i++
        Write-Progress -Activity 'Daily health check' -Status "$server ($i of $($ComputerName.Count))" -PercentComplete (($i - 1) / $ComputerName.Count * 100)
        $r = Get-ServerHealth $server
        $results.Add($r)
        $colour = switch ($r.Status) { 'Critical' { 'Red' } 'Warning' { 'Yellow' } default { 'Green' } }
        Write-Host ("{0,-25} {1}" -f $server, $r.Status) -ForegroundColor $colour
        foreach ($issue in $r.Issues) { Write-Host "    - $issue" -ForegroundColor $colour }
    }
    Write-Progress -Activity 'Daily health check' -Completed

    $report = New-HealthReportHtml $results.ToArray()
    [System.IO.File]::WriteAllText($reportPath, $report.Html, $script:Utf8NoBom)

    Write-Host ""
    Write-Host ("Critical: {0}  Warning: {1}  OK: {2}" -f $report.Critical, $report.Warning, $report.OK)
    Write-Host "Report  : $reportPath"

    if ($SendEmail) {
        if ($OnlyEmailOnIssues -and ($report.Critical + $report.Warning) -eq 0) {
            Write-Host 'All servers OK - no email sent (-OnlyEmailOnIssues).'
        }
        else {
            $state = if ($report.Critical) { 'CRITICAL' } elseif ($report.Warning) { 'Warning' } else { 'All OK' }
            $subject = "[$state] Daily health check - $($report.Critical) critical, $($report.Warning) warning, $($report.OK) OK"
            if (Send-ToolkitMail -Config $config -Subject $subject -HtmlBody $report.Html) { Write-Host 'Email sent.' }
        }
    }
}
