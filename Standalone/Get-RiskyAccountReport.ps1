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
