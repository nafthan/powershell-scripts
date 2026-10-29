<#
.SYNOPSIS
    Finds AD users and computers that haven't logged on for N days.
.DESCRIPTION
    Report only, unless -DisableAccounts is used. Then it disables the accounts and moves them, asking before each one.
.EXAMPLE
    .\Get-InactiveAccountsReport.ps1
.EXAMPLE
    .\Get-InactiveAccountsReport.ps1 -DisableAccounts -WhatIf
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
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [ValidateRange(1, 3650)]
    [int]$InactiveDays,

    [ValidateSet('Users', 'Computers', 'Both')]
    [string]$AccountType = 'Both',

    [string]$SearchBase,
    [string]$Server,
    [string[]]$Exclude,

    [switch]$DisableAccounts,
    [string]$MoveToOU,

    [switch]$SendEmail,
    [string]$OutputFile,

    # Email settings (for -SendEmail / -NotifyUsers)
    [string]$SmtpServer,
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [pscredential]$SmtpCredential,
    [string]$MailFrom,
    [string[]]$MailTo
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

if (-not $PSBoundParameters.ContainsKey('InactiveDays')) { $InactiveDays = Get-Setting $config 'ActiveDirectory' 'InactiveDays' 90 }
if (-not $PSBoundParameters.ContainsKey('SearchBase'))   { $SearchBase   = Get-Setting $config 'ActiveDirectory' 'SearchBase' '' }
if (-not $PSBoundParameters.ContainsKey('Server'))       { $Server       = Get-Setting $config 'ActiveDirectory' 'Server' '' }
if (-not $PSBoundParameters.ContainsKey('Exclude'))      { $Exclude      = Get-Setting $config 'ActiveDirectory' 'Exclude' @() }
if (-not $PSBoundParameters.ContainsKey('MoveToOU') -and $DisableAccounts) { $MoveToOU = Get-Setting $config 'ActiveDirectory' 'DisabledOU' '' }

$Exclude = Split-ListValue $Exclude

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { throw 'The ActiveDirectory module is not installed. Install RSAT: Active Directory tools (or run on a domain controller).' }

$ad = @{}
if ($Server) { $ad.Server = $Server }
$search = @{} + $ad
if ($SearchBase) { $search.SearchBase = $SearchBase }

$now    = Get-Date
$cutoff = $now.AddDays(-$InactiveDays)
$reportPath = New-ReportPath -Config $config -Prefix 'InactiveAccounts' -OutputFile $OutputFile

Write-Host ("Looking for enabled accounts with no logon since {0:yyyy-MM-dd} ({1} days)" -f $cutoff, $InactiveDays)
if ($SearchBase) { Write-Host "Search base : $SearchBase" }
Write-Host ""

function Test-Excluded([string]$Name) {
    foreach ($pattern in $Exclude) { if ($Name -like $pattern) { return $true } }
    return $false
}

function Get-ParentOU([string]$DistinguishedName) {
    return ($DistinguishedName -replace '^.+?(?<!\\),', '')
}

function Get-InactiveInfo($Account) {
    $lastLogon = if ($Account.lastLogonTimestamp) { [datetime]::FromFileTime($Account.lastLogonTimestamp) } else { $null }
    $inactive  = if ($lastLogon) { $lastLogon -lt $cutoff } else { $Account.whenCreated -lt $cutoff }
    $days      = if ($lastLogon) { [int][math]::Floor(($now - $lastLogon).TotalDays) } else { $null }
    return @{ LastLogon = $lastLogon; Inactive = $inactive; Days = $days }
}

# --- Users ----------------------------------------------------------------

$users = @()
if ($AccountType -in 'Users', 'Both') {
    Write-Host 'Querying users...'
    $props = 'lastLogonTimestamp', 'whenCreated', 'PasswordLastSet', 'Description', 'UserPrincipalName', 'EmailAddress', 'DisplayName'
    $users = @(Get-ADUser -Filter 'Enabled -eq $true' -Properties $props @search | ForEach-Object {
        if (Test-Excluded $_.SamAccountName) { return }
        $info = Get-InactiveInfo $_
        if (-not $info.Inactive) { return }
        [pscustomobject]@{
            SamAccountName    = $_.SamAccountName
            DisplayName       = $_.DisplayName
            UserPrincipalName = $_.UserPrincipalName
            Email             = $_.EmailAddress
            LastLogon         = if ($info.LastLogon) { $info.LastLogon } else { 'Never' }
            DaysInactive      = $info.Days
            Created           = $_.whenCreated
            PasswordLastSet   = $_.PasswordLastSet
            Description       = $_.Description
            OU                = Get-ParentOU $_.DistinguishedName
            DistinguishedName = $_.DistinguishedName
            ObjectType        = 'User'
        }
    } | Sort-Object -Property @{ Expression = { if ($null -eq $_.DaysInactive) { [int]::MaxValue } else { $_.DaysInactive } }; Descending = $true })
}

# --- Computers ------------------------------------------------------------

$computers = @()
if ($AccountType -in 'Computers', 'Both') {
    Write-Host 'Querying computers...'
    $props = 'lastLogonTimestamp', 'whenCreated', 'OperatingSystem', 'OperatingSystemVersion', 'Description', 'DNSHostName'
    $computers = @(Get-ADComputer -Filter 'Enabled -eq $true' -Properties $props @search | ForEach-Object {
        if (Test-Excluded $_.SamAccountName) { return }
        # Never touch domain controllers
        if ($_.DistinguishedName -match 'OU=Domain Controllers,') { return }
        $info = Get-InactiveInfo $_
        if (-not $info.Inactive) { return }
        [pscustomobject]@{
            Name              = $_.Name
            DNSHostName       = $_.DNSHostName
            OperatingSystem   = $_.OperatingSystem
            OSVersion         = $_.OperatingSystemVersion
            LastLogon         = if ($info.LastLogon) { $info.LastLogon } else { 'Never' }
            DaysInactive      = $info.Days
            Created           = $_.whenCreated
            Description       = $_.Description
            OU                = Get-ParentOU $_.DistinguishedName
            DistinguishedName = $_.DistinguishedName
            SamAccountName    = $_.SamAccountName
            ObjectType        = 'Computer'
        }
    } | Sort-Object -Property @{ Expression = { if ($null -eq $_.DaysInactive) { [int]::MaxValue } else { $_.DaysInactive } }; Descending = $true })
}

# --- Optional: disable and move -------------------------------------------

$actions = [System.Collections.Generic.List[object]]::new()
if ($DisableAccounts) {
    $targets = @($users) + @($computers)
    $stamp   = $now.ToString('yyyy-MM-dd')
    $verb    = if ($MoveToOU) { "Disable and move to $MoveToOU" } else { 'Disable' }

    foreach ($acct in $targets) {
        $label  = "$($acct.ObjectType) $($acct.SamAccountName)"
        $result = ''
        if ($PSCmdlet.ShouldProcess($label, $verb)) {
            try {
                # -Confirm:$false: the confirmation was already given above
                Disable-ADAccount -Identity $acct.DistinguishedName @ad -Confirm:$false -ErrorAction Stop
                $result = 'Disabled'
                $note = "Disabled $stamp - no logon for $InactiveDays+ days."
                if ($acct.Description) { $note += " Was: $($acct.Description)" }
                if ($note.Length -gt 1024) { $note = $note.Substring(0, 1024) }
                Set-ADObject -Identity $acct.DistinguishedName -Replace @{ description = $note } @ad -Confirm:$false -ErrorAction Stop
                if ($MoveToOU) {
                    Move-ADObject -Identity $acct.DistinguishedName -TargetPath $MoveToOU @ad -Confirm:$false -ErrorAction Stop
                    $result = 'Disabled and moved'
                }
            }
            catch {
                $msg = $_.Exception.Message
                $result = if ($result) { "${result}, then failed: $msg" } else { "Failed: $msg" }
            }
        }
        elseif ($WhatIfPreference) { $result = 'WhatIf - no change made' }
        else                       { $result = 'Skipped (not confirmed)' }

        $actions.Add([pscustomobject]@{
            ObjectType        = $acct.ObjectType
            SamAccountName    = $acct.SamAccountName
            DaysInactive      = $acct.DaysInactive
            Action            = $verb
            Result            = $result
            OriginalOU        = $acct.OU
            DistinguishedName = $acct.DistinguishedName
        })
    }
}

# --- Export ---------------------------------------------------------------

$summary = [ordered]@{
    'Run at'                = $now.ToString('yyyy-MM-dd HH:mm:ss')
    'Inactive threshold'    = "$InactiveDays days (no logon since $($cutoff.ToString('yyyy-MM-dd')))"
    'Search base'           = if ($SearchBase) { $SearchBase } else { '(whole domain)' }
    'Excluded patterns'     = ($Exclude -join ', ')
    'Inactive users'        = $users.Count
    'Inactive computers'    = $computers.Count
    'Mode'                  = if (-not $DisableAccounts) { 'Report only' } elseif ($WhatIfPreference) { 'WhatIf (no changes)' } else { 'Disable' + $(if ($MoveToOU) { ' and move' } else { '' }) }
    'Accounts changed'      = @($actions | Where-Object { $_.Result -like 'Disabled*' }).Count
    'Note'                  = 'Last logon comes from lastLogonTimestamp, which can lag the real last logon by up to 14 days.'
    'Run by'                = "$env:USERDOMAIN\$env:USERNAME"
}

$sheets = @(
    @{ Name = 'Inactive Users'; Rows = $users
       Headers = @('SamAccountName', 'DisplayName', 'UserPrincipalName', 'Email', 'LastLogon', 'DaysInactive', 'Created', 'PasswordLastSet', 'Description', 'OU')
       Widths  = @(22, 28, 34, 32, 18, 13, 18, 18, 40, 60) }
    @{ Name = 'Inactive Computers'; Rows = $computers
       Headers = @('Name', 'DNSHostName', 'OperatingSystem', 'OSVersion', 'LastLogon', 'DaysInactive', 'Created', 'Description', 'OU')
       Widths  = @(20, 34, 32, 16, 18, 13, 18, 40, 60) }
)
if ($DisableAccounts) {
    $sheets += @{ Name = 'Actions'; Rows = $actions
                  Headers = @('ObjectType', 'SamAccountName', 'DaysInactive', 'Action', 'Result', 'OriginalOU', 'DistinguishedName')
                  Widths  = @(11, 22, 13, 40, 40, 60, 80) }
}
$sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(24, 90) }

if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

Write-Host ""
Write-Host ("Inactive users     : {0:N0}" -f $users.Count)     -ForegroundColor $(if ($users.Count) { 'Yellow' } else { 'Green' })
Write-Host ("Inactive computers : {0:N0}" -f $computers.Count) -ForegroundColor $(if ($computers.Count) { 'Yellow' } else { 'Green' })
if ($DisableAccounts) {
    $actions | Group-Object Result | ForEach-Object { Write-Host ("  {0,-40} {1:N0}" -f $_.Name, $_.Count) }
}
Write-Host "Report             : $reportPath"

# --- Email ----------------------------------------------------------------

if ($SendEmail) {
    $sections = @(
        @{ Heading = "Inactive users ($($users.Count))"
           Html    = ConvertTo-HtmlTable -Rows @($users | Select-Object -First 50) -Columns @('SamAccountName', 'DisplayName', 'LastLogon', 'DaysInactive', 'OU') -EmptyText 'No inactive users.' }
        @{ Heading = "Inactive computers ($($computers.Count))"
           Html    = ConvertTo-HtmlTable -Rows @($computers | Select-Object -First 50) -Columns @('Name', 'OperatingSystem', 'LastLogon', 'DaysInactive', 'OU') -EmptyText 'No inactive computers.' }
    )
    if ($DisableAccounts) {
        $sections += @{ Heading = 'Actions taken'; Html = ConvertTo-HtmlTable -Rows $actions -Columns @('ObjectType', 'SamAccountName', 'Result') }
    }
    $intro = @("Enabled accounts with no logon for $InactiveDays days or more.", 'Tables show up to 50 rows each - the attached workbook has the full list.')
    $html  = New-HtmlReport -Title 'Inactive AD accounts' -IntroLines $intro -Sections $sections
    $subject = "Inactive AD accounts: $($users.Count) users, $($computers.Count) computers"
    if (Send-ToolkitMail -Config $config -Subject $subject -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
}
