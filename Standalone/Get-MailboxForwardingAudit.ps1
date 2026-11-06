<#
.SYNOPSIS
    Finds Exchange Online mail forwarding to external addresses and suspicious inbox rules.
.DESCRIPTION
    Checks mailbox forwarding, inbox rules (forward, redirect, and rules that hide mail - a common
    sign of a compromised account), transport rules and the tenant's external forwarding settings. Read-only.
.EXAMPLE
    .\Get-MailboxForwardingAudit.ps1
.EXAMPLE
    .\Get-MailboxForwardingAudit.ps1 -Mailbox user@contoso.com, shared@contoso.com
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
    [switch]$SkipInboxRules,
    [switch]$IncludeHiddenRules,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [switch]$SendEmail,
    [switch]$AlwaysEmail,
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

$script:HtmlFont = 'font-family:Segoe UI,Arial,sans-serif;font-size:13px;'

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
        '^(Critical|Expired|Error|Unreachable|Failed|Not done|High)' { return 'background-color:#f8d7da;color:#842029;font-weight:600;' }
        '^(Warning|Expiring|Medium)'                   { return 'background-color:#fff3cd;color:#664d03;font-weight:600;' }
        '^Low'                                         { return 'background-color:#e7f1ff;color:#084298;font-weight:600;' }
        '^(OK|Done|Requested)'                         { return 'background-color:#d1e7dd;color:#0f5132;font-weight:600;' }
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

#endregion
$config = Get-ToolkitConfig

$Mailbox = Split-ListValue $Mailbox

# --- Main -----------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitExchange -Config $config -Bound $PSBoundParameters
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'MailForwarding' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        $findings   = [System.Collections.Generic.List[object]]::new()
        function Add-Finding([string]$Risk, [string]$Mailbox, [string]$Source, [string]$Finding, [string]$Target, [string]$Detail) {
            $findings.Add([pscustomobject]@{ Risk = $Risk; Mailbox = $Mailbox; Source = $Source; Finding = $Finding; Target = $Target; Detail = $Detail })
        }

        $internal = @(Get-AcceptedDomain | ForEach-Object { "$($_.DomainName)".ToLower() })
        Write-Host "Internal domains: $($internal -join ', ')"

        # --- Mailboxes and mailbox-level forwarding ----------------------

        Write-Host 'Getting mailboxes...'
        $props = 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward', 'WhenCreated'
        if ($Mailbox.Count) { $mailboxes = @($Mailbox | ForEach-Object { Get-EXOMailbox -Identity $_ -Properties $props -ErrorAction Stop }) }
        else { $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties $props) }
        Write-Host "  $($mailboxes.Count) mailbox(es)"

        $mailboxRows = [System.Collections.Generic.List[object]]::new()
        foreach ($m in $mailboxes) {
            $targets = @()
            if ($m.ForwardingSmtpAddress) { $targets += [pscustomobject]@{ Kind = 'SMTP forwarding'; Info = Get-AddressInfo "$($m.ForwardingSmtpAddress)" $internal } }
            if ($m.ForwardingAddress) {
                $info = [pscustomobject]@{ Address = "$($m.ForwardingAddress)"; External = $false }
                try {
                    $r = Get-EXORecipient -Identity "$($m.ForwardingAddress)" -ErrorAction Stop
                    $smtp = if ($r.ExternalEmailAddress) { "$($r.ExternalEmailAddress)" } else { "$($r.PrimarySmtpAddress)" }
                    $info = Get-AddressInfo $smtp $internal
                }
                catch { $errors.Add([pscustomobject]@{ Section = 'Forwarding address'; Error = "$($m.PrimarySmtpAddress): $($_.Exception.Message)" }) }
                $targets += [pscustomobject]@{ Kind = 'Forwarding address'; Info = $info }
            }
            foreach ($t in $targets) {
                $keep = if ($m.DeliverToMailboxAndForward) { 'Yes' } else { 'No - mail is not kept in the mailbox' }
                $mailboxRows.Add([pscustomobject]@{
                    Mailbox = "$($m.PrimarySmtpAddress)"; DisplayName = $m.DisplayName; Type = "$($m.RecipientTypeDetails)"
                    ForwardingType = $t.Kind; ForwardsTo = $t.Info.Address; External = $t.Info.External; KeepsCopy = $keep
                })
                if ($t.Info.External) { Add-Finding 'High' "$($m.PrimarySmtpAddress)" 'Mailbox forwarding' 'Forwards to an external address' $t.Info.Address "$($t.Kind); keeps a copy: $keep" }
                else                  { Add-Finding 'Low'  "$($m.PrimarySmtpAddress)" 'Mailbox forwarding' 'Forwards internally' $t.Info.Address "$($t.Kind); keeps a copy: $keep" }
            }
        }

        # --- Inbox rules -------------------------------------------------

        $ruleRows = [System.Collections.Generic.List[object]]::new()
        $rulesChecked = 0
        if (-not $SkipInboxRules) {
            $ruleMailboxes = @($mailboxes | Where-Object { "$($_.RecipientTypeDetails)" -in 'UserMailbox', 'SharedMailbox' })
            $i = 0
            foreach ($m in $ruleMailboxes) {
                $i++
                Write-Progress -Activity 'Reading inbox rules' -Status "$i of $($ruleMailboxes.Count): $($m.PrimarySmtpAddress)" -PercentComplete ($i / [math]::Max(1, $ruleMailboxes.Count) * 100)
                try {
                    $ruleParams = @{ Mailbox = "$($m.PrimarySmtpAddress)"; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
                    if ($IncludeHiddenRules) { $ruleParams.IncludeHidden = $true }
                    $rules = @(Get-InboxRule @ruleParams)
                }
                catch { $errors.Add([pscustomobject]@{ Section = 'Inbox rules'; Error = "$($m.PrimarySmtpAddress): $($_.Exception.Message)" }); continue }
                $rulesChecked += $rules.Count

                foreach ($rule in $rules) {
                    $a = Get-InboxRuleAssessment $rule $internal
                    if (-not $a.Risk) { continue }
                    $ruleRows.Add([pscustomobject]@{
                        Risk = $a.Risk; Mailbox = "$($m.PrimarySmtpAddress)"; Rule = "$($rule.Name)"; Enabled = [bool]$rule.Enabled
                        Actions = $a.Actions; ForwardsTo = $a.Targets; Conditions = $a.Conditions
                        Issues = ($a.Issues -join '; '); RuleIdentity = "$($rule.Identity)"
                    })
                    Add-Finding $a.Risk "$($m.PrimarySmtpAddress)" 'Inbox rule' ($a.Short -join ' + ') $a.Targets "Rule '$($rule.Name)' ($(if ($rule.Enabled) { 'enabled' } else { 'disabled' })): $($a.Actions)"
                }
            }
            Write-Progress -Activity 'Reading inbox rules' -Completed
        }

        # --- Transport rules ---------------------------------------------

        $transportRows = [System.Collections.Generic.List[object]]::new()
        try {
            foreach ($t in @(Get-TransportRule -ResultSize Unlimited)) {
                $targets = @()
                foreach ($field in 'RedirectMessageTo', 'BlindCopyTo', 'CopyTo', 'AddToRecipients') {
                    foreach ($entry in @($t.$field | Where-Object { $_ })) { $targets += [pscustomobject]@{ Action = $field; Info = Get-AddressInfo "$entry" $internal } }
                }
                if (-not $targets.Count) { continue }
                $ext = @($targets | Where-Object { $_.Info.External })
                $transportRows.Add([pscustomobject]@{
                    Rule = $t.Name; State = "$($t.State)"; Actions = (@($targets.Action | Sort-Object -Unique) -join ', ')
                    Targets = (@($targets | ForEach-Object { "$($_.Info.Address)$(if ($_.Info.External) { ' (external)' })" }) -join ', ')
                    External = [bool]$ext.Count; Priority = $t.Priority; Modified = $t.WhenChanged
                })
                if ($ext.Count) { Add-Finding $(if ("$($t.State)" -eq 'Enabled') { 'High' } else { 'Low' }) '(organisation)' 'Transport rule' 'Copies or redirects mail to an external address' (($ext | ForEach-Object { $_.Info.Address }) -join ', ') "Rule '$($t.Name)' ($($t.State))" }
            }
        }
        catch { $errors.Add([pscustomobject]@{ Section = 'Transport rules'; Error = $_.Exception.Message }) }

        # --- Tenant settings ---------------------------------------------

        $settings = [System.Collections.Generic.List[object]]::new()
        $externalBlocked = $null
        try {
            foreach ($p in @(Get-HostedOutboundSpamFilterPolicy)) {
                $mode = "$($p.AutoForwardingMode)"
                $meaning = switch ($mode) { 'Off' { 'External automatic forwarding blocked' } 'On' { 'External automatic forwarding allowed' } default { 'Automatic (Microsoft blocks external automatic forwarding)' } }
                $settings.Add([pscustomobject]@{ Setting = 'Outbound spam policy'; Name = $p.Name; Value = $mode; Meaning = $meaning })
                if ($p.IsDefault -or $p.Name -eq 'Default') { $externalBlocked = $mode -ne 'On' }
                if ($mode -eq 'On') { Add-Finding $(if ($p.IsDefault -or $p.Name -eq 'Default') { 'Medium' } else { 'Low' }) '(organisation)' 'Outbound spam policy' 'External automatic forwarding is allowed' '' "Policy '$($p.Name)' has AutoForwardingMode = On" }
            }
        }
        catch { $errors.Add([pscustomobject]@{ Section = 'Outbound spam policy'; Error = $_.Exception.Message }) }
        try {
            foreach ($d in @(Get-RemoteDomain)) {
                $settings.Add([pscustomobject]@{ Setting = 'Remote domain'; Name = "$($d.Name) ($($d.DomainName))"; Value = "AutoForwardEnabled = $($d.AutoForwardEnabled)"; Meaning = if ($d.AutoForwardEnabled) { 'Forwarding to this domain allowed (if the outbound policy allows it)' } else { 'Forwarding to this domain blocked' } })
            }
        }
        catch { $errors.Add([pscustomobject]@{ Section = 'Remote domains'; Error = $_.Exception.Message }) }

        # --- Export ------------------------------------------------------

        $order = @{ Critical = 0; High = 1; Medium = 2; Low = 3 }
        $findingsSorted = @($findings | Sort-Object @{ Expression = { $order[$_.Risk] } }, Mailbox)
        $counts = [ordered]@{}
        foreach ($r in 'Critical', 'High', 'Medium', 'Low') { $counts[$r] = @($findingsSorted | Where-Object { $_.Risk -eq $r }).Count }
        $org = ''
        try { $org = @($internal | Where-Object { $_ -like '*.onmicrosoft.com' -and $_ -notlike '*.mail.onmicrosoft.com' })[0] } catch { $org = '' }
        if (-not $org) { $org = $internal | Select-Object -First 1 }

        $summary = [ordered]@{
            'Organisation'                 = "$org"
            'Run at'                       = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Mailboxes checked'            = $mailboxes.Count
            'Inbox rules checked'          = if ($SkipInboxRules) { 'Skipped (-SkipInboxRules)' } else { $rulesChecked }
            'Mailboxes forwarding'         = "$(@($mailboxRows | Select-Object -ExpandProperty Mailbox -Unique).Count) ($(@($mailboxRows | Where-Object External).Count) external)"
            'Risky inbox rules'            = $ruleRows.Count
            'Transport rules copying mail' = $transportRows.Count
            'Findings: Critical'           = $counts.Critical
            'Findings: High'               = $counts.High
            'Findings: Medium'             = $counts.Medium
            'Findings: Low'                = $counts.Low
            'External auto-forwarding'     = if ($null -eq $externalBlocked) { 'Unknown' } elseif ($externalBlocked) { 'Blocked by the default outbound spam policy - external forwards found here are failing, but should still be removed' } else { 'ALLOWED by the default outbound spam policy' }
            'Note'                         = 'Some rules created through other mail clients or APIs are hidden from Get-InboxRule. Use -IncludeHiddenRules to include them. For a suspected compromise, also review the sign-in logs and revoke sessions.'
        }

        $saved = Save-Workbook -FilePath $reportPath -Sheets @(
            @{ Name = 'Findings'; Rows = $findingsSorted; Headers = @('Risk', 'Mailbox', 'Source', 'Finding', 'Target', 'Detail'); Widths = @(10, 34, 20, 50, 40, 70) }
            @{ Name = 'Mailbox Forwarding'; Rows = $mailboxRows; Headers = @('Mailbox', 'DisplayName', 'Type', 'ForwardingType', 'ForwardsTo', 'External', 'KeepsCopy'); Widths = @(34, 26, 16, 18, 36, 10, 36) }
            @{ Name = 'Inbox Rules'; Rows = @($ruleRows | Sort-Object @{ Expression = { $order[$_.Risk] } }, Mailbox)
               Headers = @('Risk', 'Mailbox', 'Rule', 'Enabled', 'Actions', 'ForwardsTo', 'Conditions', 'Issues', 'RuleIdentity'); Widths = @(10, 34, 26, 9, 40, 40, 40, 70, 40) }
            @{ Name = 'Transport Rules'; Rows = $transportRows; Headers = @('Rule', 'State', 'Actions', 'Targets', 'External', 'Priority', 'Modified'); Widths = @(36, 10, 30, 50, 10, 9, 18) }
            @{ Name = 'Settings'; Rows = $settings; Headers = @('Setting', 'Name', 'Value', 'Meaning'); Widths = @(22, 36, 26, 60) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(22, 120) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(28, 110) }
        )
        if (-not $saved) { exit 1 }

        Write-Host ""
        Write-Host ("Findings: {0} critical, {1} high, {2} medium, {3} low" -f $counts.Critical, $counts.High, $counts.Medium, $counts.Low) -ForegroundColor $(if ($counts.Critical + $counts.High) { 'Red' } elseif ($counts.Medium) { 'Yellow' } else { 'Green' })
        $findingsSorted | Where-Object { $_.Risk -in 'Critical', 'High' } | Select-Object -First 15 | ForEach-Object { Write-Host ("  {0,-8} {1}: {2} {3}" -f $_.Risk, $_.Mailbox, $_.Finding, $(if ($_.Target) { "-> $($_.Target)" })) }
        Write-Host ("External auto-forwarding: {0}" -f $summary['External auto-forwarding'])
        if ($errors.Count) { Write-Host ("Some checks failed - see the 'Errors' sheet ({0})." -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report  : $reportPath"

        if ($SendEmail) {
            $important = @($findingsSorted | Where-Object { $_.Risk -ne 'Low' })
            if (-not $important.Count -and -not $AlwaysEmail) {
                Write-Host 'No critical, high or medium findings - no email sent (use -AlwaysEmail to send anyway).'
            }
            else {
                $html = New-HtmlReport -Title 'Mailbox forwarding audit' -IntroLines @("$($mailboxes.Count) mailbox(es) checked: $($counts.Critical) critical, $($counts.High) high, $($counts.Medium) medium.", "External auto-forwarding: $($summary['External auto-forwarding']).") -Sections @(
                    @{ Heading = "Findings ($($important.Count))"; Html = ConvertTo-HtmlTable -Rows $important -Columns @('Risk', 'Mailbox', 'Source', 'Finding', 'Target', 'Detail') -StatusColumn 'Risk' -EmptyText 'None.' }
                )
                $state = if ($counts.Critical) { 'CRITICAL' } elseif ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Medium' } else { 'OK' }
                if (Send-ToolkitMail -Config $config -Subject "[$state] Mailbox forwarding audit - $($important.Count) finding(s)" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
            }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    }
}
