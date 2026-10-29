<#
.SYNOPSIS
    Lists AD passwords that are expiring soon and can email users a reminder.
.DESCRIPTION
    Reminders are only sent at 14, 7, 3, 2, 1 and 0 days left.
.EXAMPLE
    .\Get-PasswordExpiryReport.ps1 -WarningDays 21
.EXAMPLE
    .\Get-PasswordExpiryReport.ps1 -NotifyUsers -SendReport
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: ActiveDirectory module (RSAT)
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateRange(0, 365)]
    [int]$WarningDays,

    [switch]$SkipExpired,

    [string]$SearchBase,
    [string]$Server,
    [string[]]$Exclude,

    [switch]$NotifyUsers,
    [switch]$NotifyAll,
    [string]$TestRecipient,

    [switch]$SendReport,
    [string]$OutputFile,

    # Email settings (for -SendEmail / -NotifyUsers)
    [string]$SmtpServer,
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [pscredential]$SmtpCredential,
    [string]$MailFrom,
    [string[]]$MailTo,
    [string]$HelpdeskContact
)

#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)

$script:ToolkitRoot   = $PSScriptRoot

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)

$script:IsWindowsHost = ($PSVersionTable.PSEdition -eq 'Desktop') -or $IsWindows

$script:HtmlFont = 'font-family:Segoe UI,Arial,sans-serif;font-size:13px;'

function Get-ToolkitConfig {
    # Standalone: no config file. Settings come from parameters and built-in defaults.
    $cfg = @{ General = @{ OutputFolder = (Get-Location).Path } }
    if ($SmtpServer) {
        $cfg.Email = @{
            SmtpServer = $SmtpServer; Port = $SmtpPort; UseSsl = [bool]$SmtpUseSsl; From = $MailFrom
            To = @($MailTo | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
    }
    if ($HelpdeskContact) { $cfg.PasswordExpiry = @{ HelpdeskContact = $HelpdeskContact } }
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
        $name = [char](65 + $rem) + $name
        $n    = [math]::Floor(($n - 1) / 26)
    }
    return $name
}

function Add-ZipText($Zip, [string]$Name, [string]$Text) {
    $entry  = $Zip.CreateEntry($Name)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Add-XlsxSheet {
    param($Zip, [int]$Index, [string[]]$Headers, $Rows, [int[]]$Widths)

    $inv     = [System.Globalization.CultureInfo]::InvariantCulture
    $letters = for ($i = 0; $i -lt $Headers.Count; $i++) { Get-ColumnLetter $i }
    $entry   = $Zip.CreateEntry("xl/worksheets/sheet$Index.xml")
    $w       = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try {
        $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        $w.Write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        $w.Write('<cols>')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $width = if ($Widths -and $i -lt $Widths.Count) { $Widths[$i] } else { 20 }
            $w.Write(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $width))
        }
        $w.Write('</cols><sheetData>')

        $w.Write('<row r="1">')
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
                if ($null -eq $val) {
                    continue
                }
                elseif ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal] -or $val -is [single]) {
                    $w.Write("<c r=`"$ref`"><v>$(([double]$val).ToString('R', $inv))</v></c>")
                }
                elseif ($val -is [datetime]) {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t>$($val.ToString('yyyy-MM-dd HH:mm', $inv))</t></is></c>")
                }
                elseif ($val -is [bool]) {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t>$(if ($val) { 'Yes' } else { 'No' })</t></is></c>")
                }
                else {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText "$val")</t></is></c>")
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

        Add-ZipText $zip 'xl/styles.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
            '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor indexed="64"/></patternFill></fill></fills>' +
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>' +
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
    switch -Regex ($Status) {
        '^(Critical|Expired|Error|Unreachable|Failed|High)' { return 'background-color:#f8d7da;color:#842029;font-weight:600;' }
        '^(Warning|Expiring|Medium)'                   { return 'background-color:#fff3cd;color:#664d03;font-weight:600;' }
        '^Low'                                         { return 'background-color:#e7f1ff;color:#084298;font-weight:600;' }
        '^OK'                                          { return 'background-color:#d1e7dd;color:#0f5132;font-weight:600;' }
        default                                        { return '' }
    }
}

function ConvertTo-HtmlTable {
    <#
        Builds a table with inline styles (works in Outlook).
        -StatusColumn colours that cell by its value (OK / Warning / Critical ...).
    #>
    param(
        $Rows,
        [string[]]$Columns,
        [string]$StatusColumn,
        [string]$EmptyText = 'None.'
    )
    if ($null -eq $Rows -or @($Rows).Count -eq 0) {
        return "<p style=`"$($script:HtmlFont)color:#555;`">$([System.Net.WebUtility]::HtmlEncode($EmptyText))</p>"
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table cellpadding=`"6`" cellspacing=`"0`" style=`"border-collapse:collapse;$($script:HtmlFont)border:1px solid #d0d7de;margin-bottom:16px;`">")
    [void]$sb.Append('<tr>')
    foreach ($c in $Columns) {
        [void]$sb.Append("<th style=`"background-color:#305496;color:#ffffff;text-align:left;border:1px solid #d0d7de;`">$([System.Net.WebUtility]::HtmlEncode($c))</th>")
    }
    [void]$sb.Append('</tr>')
    $odd = $false
    foreach ($row in $Rows) {
        $odd = -not $odd
        $bg  = if ($odd) { '#ffffff' } else { '#f6f8fa' }
        [void]$sb.Append("<tr style=`"background-color:$bg;`">")
        foreach ($c in $Columns) {
            $val = $row.$c
            if ($val -is [datetime]) { $val = $val.ToString('yyyy-MM-dd HH:mm') }
            elseif ($val -is [bool]) { $val = if ($val) { 'Yes' } else { 'No' } }
            $style = 'border:1px solid #d0d7de;vertical-align:top;'
            if ($c -eq $StatusColumn) { $style += Get-StatusStyle "$val" }
            [void]$sb.Append("<td style=`"$style`">$([System.Net.WebUtility]::HtmlEncode("$val"))</td>")
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

function New-HtmlReport {
    <#
        Wraps sections into a full HTML page.
        -Sections: array of @{ Heading = '...'; Html = '...' }
    #>
    param(
        [string]$Title,
        [string[]]$IntroLines,
        [object[]]$Sections
    )
    $enc = { param($t) [System.Net.WebUtility]::HtmlEncode($t) }
    $sb  = New-Object System.Text.StringBuilder
    [void]$sb.Append("<!DOCTYPE html><html><head><meta charset=`"utf-8`"><title>$(& $enc $Title)</title></head>")
    [void]$sb.Append("<body style=`"$($script:HtmlFont)color:#1f2328;background-color:#ffffff;margin:16px;`">")
    [void]$sb.Append("<h2 style=`"font-family:Segoe UI,Arial,sans-serif;color:#305496;margin:0 0 8px 0;`">$(& $enc $Title)</h2>")
    foreach ($line in $IntroLines) {
        [void]$sb.Append("<p style=`"$($script:HtmlFont)margin:2px 0;`">$(& $enc $line)</p>")
    }
    foreach ($s in $Sections) {
        [void]$sb.Append("<h3 style=`"font-family:Segoe UI,Arial,sans-serif;color:#1f2328;margin:20px 0 8px 0;`">$(& $enc $s.Heading)</h3>")
        [void]$sb.Append($s.Html)
    }
    [void]$sb.Append("<p style=`"font-family:Segoe UI,Arial,sans-serif;font-size:11px;color:#6e7781;margin-top:24px;`">Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') on $([System.Environment]::MachineName) by $(& $enc (Split-Path -Leaf "$($MyInvocation.PSCommandPath)")).</p>")
    [void]$sb.Append('</body></html>')
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

if (-not $PSBoundParameters.ContainsKey('WarningDays')) { $WarningDays = Get-Setting $config 'PasswordExpiry' 'WarningDays' 14 }
if (-not $PSBoundParameters.ContainsKey('SearchBase'))  { $SearchBase  = Get-Setting $config 'ActiveDirectory' 'SearchBase' '' }
if (-not $PSBoundParameters.ContainsKey('Server'))      { $Server      = Get-Setting $config 'ActiveDirectory' 'Server' '' }
if (-not $PSBoundParameters.ContainsKey('Exclude'))     { $Exclude     = Get-Setting $config 'ActiveDirectory' 'Exclude' @() }

$Exclude      = Split-ListValue $Exclude
$notifyOnDays = @(Get-Setting $config 'PasswordExpiry' 'NotifyOnDays' @(14, 7, 3, 2, 1, 0))

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { throw 'The ActiveDirectory module is not installed. Install RSAT: Active Directory tools (or run on a domain controller).' }

$search = @{}
if ($Server)     { $search.Server = $Server }
if ($SearchBase) { $search.SearchBase = $SearchBase }

$now        = Get-Date
$today      = $now.Date
$reportPath = New-ReportPath -Config $config -Prefix 'PasswordExpiry' -OutputFile $OutputFile

Write-Host "Checking passwords that expire within $WarningDays day(s)..."

# --- Query ----------------------------------------------------------------

$props = 'msDS-UserPasswordExpiryTimeComputed', 'PasswordLastSet', 'EmailAddress', 'DisplayName', 'GivenName', 'UserPrincipalName'
$filter = 'Enabled -eq $true -and PasswordNeverExpires -eq $false'

$users = @(Get-ADUser -Filter $filter -Properties $props @search | ForEach-Object {
    $u = $_
    foreach ($pattern in $Exclude) { if ($u.SamAccountName -like $pattern) { return } }
    if (-not $u.PasswordLastSet) { return }                      # must change at next logon

    $raw = $u.'msDS-UserPasswordExpiryTimeComputed'
    if (-not $raw -or $raw -eq 0 -or $raw -ge 9223372036854775807) { return }   # never expires

    $expires  = [datetime]::FromFileTime($raw)
    $daysLeft = [int][math]::Floor(($expires.Date - $today).TotalDays)
    if ($daysLeft -gt $WarningDays) { return }
    if ($expires -lt $now -and $SkipExpired) { return }

    [pscustomobject]@{
        SamAccountName  = $u.SamAccountName
        DisplayName     = $u.DisplayName
        GivenName       = $u.GivenName
        Email           = $u.EmailAddress
        PasswordLastSet = $u.PasswordLastSet
        ExpiresOn       = $expires
        DaysLeft        = $daysLeft
        Status          = if ($expires -lt $now) { 'Expired' } elseif ($daysLeft -le 3) { 'Expiring - urgent' } else { 'Expiring' }
        Notification    = ''
        OU              = ($u.DistinguishedName -replace '^.+?(?<!\\),', '')
    }
} | Sort-Object -Property ExpiresOn)

$expired  = @($users | Where-Object { $_.Status -eq 'Expired' })
$expiring = @($users | Where-Object { $_.Status -ne 'Expired' })

# --- Notify users ---------------------------------------------------------

$notifications = [System.Collections.Generic.List[object]]::new()

function New-UserNotificationHtml($User) {
    $helpdesk     = Get-Setting $config 'PasswordExpiry' 'HelpdeskContact' 'the IT Service Desk'
    $instructions = Get-Setting $config 'PasswordExpiry' 'ChangeInstructions' 'Press Ctrl+Alt+Delete on a domain-joined PC and choose "Change a password".'
    $enc  = { param($t) [System.Net.WebUtility]::HtmlEncode($t) }
    $name = if ($User.GivenName) { $User.GivenName } elseif ($User.DisplayName) { $User.DisplayName } else { $User.SamAccountName }
    $when = switch ($User.DaysLeft) {
        { $_ -le 0 } { 'today'; break }
        1            { 'tomorrow'; break }
        default      { "in $($User.DaysLeft) days" }
    }
    $font = 'font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1f2328;'
    return @"
<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="$font">
<p style="$font">Hello $(& $enc $name),</p>
<p style="$font">The password for your account <b>$(& $enc $User.SamAccountName)</b> expires <b>$when</b> ($($User.ExpiresOn.ToString('dddd d MMMM yyyy, HH:mm'))).</p>
<p style="$font">To change it: $(& $enc $instructions)</p>
<p style="$font">If you are working remotely, connect to the VPN before changing your password so your laptop picks up the new one.</p>
<p style="$font">If you need help, please contact $(& $enc $helpdesk).</p>
<p style="$font">This is an automated message - please do not reply.</p>
</body></html>
"@
}

if ($NotifyUsers) {
    foreach ($u in $expiring) {
        if (-not $NotifyAll -and $notifyOnDays -notcontains $u.DaysLeft) {
            $u.Notification = "Not due (notifies at $($notifyOnDays -join ', ') days left)"
            continue
        }
        if (-not $u.Email) { $u.Notification = 'No email address in AD'; continue }

        $to      = if ($TestRecipient) { $TestRecipient } else { $u.Email }
        $subject = if ($u.DaysLeft -le 0) { 'Your password expires today' } elseif ($u.DaysLeft -eq 1) { 'Your password expires tomorrow' } else { "Your password expires in $($u.DaysLeft) days" }

        if ($PSCmdlet.ShouldProcess($to, "Send password expiry reminder for $($u.SamAccountName)")) {
            $sent = Send-ToolkitMail -Config $config -To $to -Subject $subject -HtmlBody (New-UserNotificationHtml $u)
            $u.Notification = if ($sent) { "Sent to $to" } else { 'Send failed - see console warning' }
        }
        else {
            $u.Notification = "WhatIf - would send to $to"
        }
        $notifications.Add([pscustomobject]@{ SamAccountName = $u.SamAccountName; DaysLeft = $u.DaysLeft; SentTo = $to; Result = $u.Notification })
    }
    foreach ($u in $expired) { $u.Notification = 'Already expired - not emailed' }
}

# --- Export ---------------------------------------------------------------

$summary = [ordered]@{
    'Run at'                   = $now.ToString('yyyy-MM-dd HH:mm:ss')
    'Warning window'           = "$WarningDays days"
    'Search base'              = if ($SearchBase) { $SearchBase } else { '(whole domain)' }
    'Expiring within window'   = $expiring.Count
    'Already expired'          = if ($SkipExpired) { 'Not included (-SkipExpired)' } else { $expired.Count }
    'Expiring with no email'   = @($expiring | Where-Object { -not $_.Email }).Count
    'Notifications'            = if (-not $NotifyUsers) { 'Not sent (use -NotifyUsers)' } elseif ($TestRecipient) { "Test mode - all sent to $TestRecipient" } else { "Sent on days left: $($notifyOnDays -join ', ')" }
    'Notifications sent'       = @($notifications | Where-Object { $_.Result -like 'Sent*' }).Count
    'Run by'                   = "$env:USERDOMAIN\$env:USERNAME"
}

$sheets = @(
    @{ Name = 'Password Expiry'; Rows = $users
       Headers = @('SamAccountName', 'DisplayName', 'Email', 'PasswordLastSet', 'ExpiresOn', 'DaysLeft', 'Status', 'Notification', 'OU')
       Widths  = @(22, 28, 34, 18, 18, 10, 18, 50, 60) }
)
if ($NotifyUsers) {
    $sheets += @{ Name = 'Notifications'; Rows = $notifications; Headers = @('SamAccountName', 'DaysLeft', 'SentTo', 'Result'); Widths = @(22, 10, 34, 50) }
}
$sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(26, 80) }

if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

Write-Host ""
Write-Host ("Expiring within {0} days : {1:N0}" -f $WarningDays, $expiring.Count) -ForegroundColor $(if ($expiring.Count) { 'Yellow' } else { 'Green' })
if (-not $SkipExpired) { Write-Host ("Already expired          : {0:N0}" -f $expired.Count) -ForegroundColor $(if ($expired.Count) { 'Red' } else { 'Green' }) }
if ($NotifyUsers) { Write-Host ("Notifications sent       : {0:N0}" -f $summary['Notifications sent']) }
Write-Host "Report                   : $reportPath"

# --- Email the IT team ----------------------------------------------------

if ($SendReport) {
    $cols = @('SamAccountName', 'DisplayName', 'ExpiresOn', 'DaysLeft', 'Status', 'Notification')
    $sections = @(
        @{ Heading = "Expiring within $WarningDays days ($($expiring.Count))"; Html = ConvertTo-HtmlTable -Rows $expiring -Columns $cols -StatusColumn 'Status' -EmptyText 'None.' }
    )
    if (-not $SkipExpired) {
        $sections += @{ Heading = "Already expired ($($expired.Count))"; Html = ConvertTo-HtmlTable -Rows @($expired | Select-Object -First 100) -Columns $cols -StatusColumn 'Status' -EmptyText 'None.' }
    }
    $html = New-HtmlReport -Title 'Password expiry report' -IntroLines @("Enabled users whose passwords expire within $WarningDays days.") -Sections $sections
    if (Send-ToolkitMail -Config $config -Subject "Password expiry: $($expiring.Count) expiring, $($expired.Count) expired" -HtmlBody $html -Attachments $reportPath) {
        Write-Host 'Report emailed.'
    }
}
