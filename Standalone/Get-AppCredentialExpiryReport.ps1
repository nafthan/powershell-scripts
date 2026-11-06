<#
.SYNOPSIS
    Finds Entra app registration secrets and certificates that are expired or expiring soon.
.DESCRIPTION
    Covers app registrations and enterprise apps (including SAML signing certificates).
    Also flags long-lived secrets and apps with no owners. Read-only.
.EXAMPLE
    .\Get-AppCredentialExpiryReport.ps1
.EXAMPLE
    .\Get-AppCredentialExpiryReport.ps1 -WarningDays 60 -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: Microsoft.Graph.Authentication module, Global Reader (or Application Administrator)
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [ValidateRange(1, 3650)] [int]$WarningDays,
    [ValidateRange(0, 3650)] [int]$CriticalDays,
    [switch]$ExpiringOnly,

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

#endregion
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('WarningDays'))  { $WarningDays  = [int](Get-Setting $config 'AppCredentials' 'WarningDays' 30) }
if (-not $PSBoundParameters.ContainsKey('CriticalDays')) { $CriticalDays = [int](Get-Setting $config 'AppCredentials' 'CriticalDays' 7) }
$ignoreExpired = [int](Get-Setting $config 'AppCredentials' 'IgnoreExpiredOlderThanDays' 90)
$maxLifetime   = [int](Get-Setting $config 'AppCredentials' 'MaxSecretLifetimeDays' 730)

$scopes = @('Application.Read.All')

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'Global Reader (or Application Administrator)'
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'AppCredentials' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        $tenantName = Get-GraphTenantName
        Write-Host "Tenant : $tenantName"

        # --- Collect -----------------------------------------------------

        Write-Host 'Getting app registrations...'
        $apps = @(Get-GraphAll 'v1.0/applications?$select=id,appId,displayName,createdDateTime,signInAudience,passwordCredentials,keyCredentials&$top=999')
        Write-Host 'Getting enterprise apps...'
        $sps = @(Get-GraphAll 'v1.0/servicePrincipals?$select=id,appId,displayName,servicePrincipalType,preferredSingleSignOnMode,accountEnabled,passwordCredentials,keyCredentials&$top=999' |
                 Where-Object { @($_.passwordCredentials | Where-Object { $_ }).Count -or @($_.keyCredentials | Where-Object { $_ }).Count })

        $objects = @()
        $objects += @($apps | Where-Object { @($_.passwordCredentials | Where-Object { $_ }).Count -or @($_.keyCredentials | Where-Object { $_ }).Count } |
                      ForEach-Object { [pscustomobject]@{ Kind = 'App registration'; Path = 'applications'; Item = $_ } })
        $objects += @($sps | ForEach-Object { [pscustomobject]@{ Kind = 'Enterprise app'; Path = 'servicePrincipals'; Item = $_ } })
        Write-Host ("  {0} app registration(s), {1} with credentials; {2} enterprise app(s) with credentials" -f $apps.Count, @($objects | Where-Object { $_.Kind -eq 'App registration' }).Count, $sps.Count)

        # --- Build rows --------------------------------------------------

        $rows    = [System.Collections.Generic.List[object]]::new()
        $appRows = [System.Collections.Generic.List[object]]::new()
        $order   = @{ Expired = 0; Critical = 1; Warning = 2; OK = 3 }
        $i = 0
        foreach ($o in $objects) {
            $i++
            if ($i % 10 -eq 0 -or $i -eq 1) { Write-Progress -Activity 'Checking app credentials' -Status "$i of $($objects.Count)" -PercentComplete ($i / [math]::Max(1, $objects.Count) * 100) }
            $item = $o.Item

            $owners = ''
            try {
                $owners = (@(Get-GraphAll "v1.0/$($o.Path)/$($item.id)/owners?`$select=displayName,userPrincipalName" |
                             ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }) -join ', ')
            }
            catch { $errors.Add([pscustomobject]@{ Section = 'Owners'; Error = "$($item.displayName): $($_.Exception.Message)" }) }

            $isSaml = $o.Kind -eq 'Enterprise app' -and "$($item.preferredSingleSignOnMode)" -eq 'saml'
            $creds = @()
            foreach ($c in @($item.passwordCredentials | Where-Object { $_ })) {
                if ($isSaml) { continue }   # the SAML certificate's key password - covered by the certificate itself
                $creds += [pscustomobject]@{ Type = 'Client secret'; Cred = $c }
            }
            foreach ($c in @($item.keyCredentials | Where-Object { $_ })) {
                $creds += [pscustomobject]@{ Type = $(if ($isSaml) { 'SAML signing certificate' } else { 'Certificate' }); Cred = $c }
            }

            $seen = @{}
            $objRows = @()
            foreach ($entry in $creds) {
                $c      = $entry.Cred
                $end    = ConvertTo-LocalTime $c.endDateTime
                $start  = ConvertTo-LocalTime $c.startDateTime
                if (-not $end) { continue }
                # A certificate is stored twice (Sign + Verify) on SAML apps - list it once
                $key = "$($entry.Type)|$($end.ToString('s'))|$($c.customKeyIdentifier)"
                if ($seen.ContainsKey($key)) { continue }
                $seen[$key] = $true

                $daysLeft = [int][math]::Floor(($end - $now).TotalDays)
                if ($daysLeft -lt -$ignoreExpired) { continue }
                $status = if ($daysLeft -lt 0) { 'Expired' } elseif ($daysLeft -le $CriticalDays) { 'Critical' } elseif ($daysLeft -le $WarningDays) { 'Warning' } else { 'OK' }

                $issues = @()
                if ($entry.Type -eq 'Client secret' -and $start -and ($end - $start).TotalDays -gt $maxLifetime) {
                    $issues += "Long-lived secret ($([math]::Round(($end - $start).TotalDays / 365, 1)) years) - use a shorter lifetime or a certificate"
                }
                if ($status -eq 'Expired') { $issues += 'Expired - remove it if a newer credential has replaced it' }

                $objRows += [pscustomobject]@{
                    Status         = $status
                    DaysLeft       = $daysLeft
                    Expires        = $end
                    App            = $item.displayName
                    AppId          = $item.appId
                    ObjectType     = $o.Kind
                    CredentialType = $entry.Type
                    Description    = "$($c.displayName)"
                    KeyId          = $c.keyId
                    Created        = $start
                    Owners         = $owners
                    Issues         = ($issues -join '; ')
                }
            }
            if (-not $objRows.Count) { continue }

            $appIssues = @()
            if (-not $owners) { $appIssues += 'No owners - nobody is responsible for renewing it' }
            if (-not @($objRows | Where-Object { $_.Status -ne 'Expired' }).Count) { $appIssues += 'All credentials expired - the app may be broken or no longer used' }
            if ($o.Kind -eq 'Enterprise app' -and $item.accountEnabled -eq $false) { $appIssues += 'Disabled' }
            $appIssues += @($objRows | Where-Object { $_.Issues -like 'Long-lived*' } | ForEach-Object { 'Long-lived secret' } | Select-Object -First 1)

            $worst    = ($objRows | Sort-Object @{ Expression = { $order[$_.Status] } } | Select-Object -First 1).Status
            $upcoming = @($objRows | Where-Object { $_.DaysLeft -ge 0 } | Sort-Object Expires | Select-Object -First 1)
            $next     = if ($upcoming.Count) { $upcoming[0] } else { $objRows | Sort-Object Expires -Descending | Select-Object -First 1 }

            if (-not $ExpiringOnly -or $worst -ne 'OK' -or $appIssues.Count) {
                foreach ($r in $objRows) { if (-not $ExpiringOnly -or $r.Status -ne 'OK' -or $r.Issues) { $rows.Add($r) } }
                $appRows.Add([pscustomobject]@{
                    Status      = $worst
                    App         = $item.displayName
                    AppId       = $item.appId
                    ObjectType  = $o.Kind
                    Credentials = $objRows.Count
                    NextExpiry  = $next.Expires
                    DaysLeft    = $next.DaysLeft
                    Owners      = if ($owners) { $owners } else { '(none)' }
                    Issues      = (@($appIssues | Where-Object { $_ }) -join '; ')
                })
            }
        }
        Write-Progress -Activity 'Checking app credentials' -Completed

        # --- Export ------------------------------------------------------

        $rowsSorted = @($rows | Sort-Object @{ Expression = { $order[$_.Status] } }, DaysLeft, App)
        $appsSorted = @($appRows | Sort-Object @{ Expression = { $order[$_.Status] } }, DaysLeft, App)
        $counts = [ordered]@{}
        foreach ($s in 'Expired', 'Critical', 'Warning', 'OK') { $counts[$s] = @($rowsSorted | Where-Object { $_.Status -eq $s }).Count }

        $summary = [ordered]@{
            'Tenant'                    = $tenantName
            'Run at'                    = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Thresholds'                = "Warning <= $WarningDays days, Critical <= $CriticalDays days"
            'App registrations'         = $apps.Count
            'Apps with credentials'     = $appRows.Count
            'Expired'                   = $counts.Expired
            'Critical'                  = $counts.Critical
            'Warning'                   = $counts.Warning
            'OK'                        = if ($ExpiringOnly) { 'Not listed (-ExpiringOnly)' } else { $counts.OK }
            'Apps with no owners'       = @($appsSorted | Where-Object { $_.Owners -eq '(none)' }).Count
            'Long-lived secrets'        = @($rowsSorted | Where-Object { $_.Issues -like 'Long-lived*' }).Count
            'Note'                      = "Credentials that expired more than $ignoreExpired days ago are not listed. Secrets longer than $maxLifetime days are flagged."
            'Run by'                    = Get-GraphRunBy
        }

        $saved = Save-Workbook -FilePath $reportPath -Sheets @(
            @{ Name = 'Credentials'; Rows = $rowsSorted
               Headers = @('Status', 'DaysLeft', 'Expires', 'App', 'ObjectType', 'CredentialType', 'Description', 'Owners', 'Issues', 'Created', 'AppId', 'KeyId')
               Widths  = @(10, 10, 18, 36, 17, 24, 26, 40, 60, 18, 38, 38) }
            @{ Name = 'Apps'; Rows = $appsSorted
               Headers = @('Status', 'App', 'ObjectType', 'Credentials', 'NextExpiry', 'DaysLeft', 'Owners', 'Issues', 'AppId')
               Widths  = @(10, 36, 17, 12, 18, 10, 40, 60, 38) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(20, 120) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(24, 100) }
        )
        if (-not $saved) { exit 1 }

        Write-Host ""
        foreach ($s in 'Expired', 'Critical', 'Warning') {
            Write-Host ("{0,-9} {1}" -f $s, $counts[$s]) -ForegroundColor $(if ($counts[$s]) { if ($s -eq 'Warning') { 'Yellow' } else { 'Red' } } else { 'Green' })
        }
        Write-Host ("Apps with no owners: {0}" -f $summary['Apps with no owners'])
        Write-Host "Report : $reportPath"

        $attention = @($rowsSorted | Where-Object { $_.Status -ne 'OK' })
        if ($SendEmail) {
            if (-not $attention.Count -and -not $AlwaysEmail) {
                Write-Host 'Nothing expiring - no email sent (use -AlwaysEmail to send anyway).'
            }
            else {
                $intro = @("$tenantName - $($counts.Expired) expired, $($counts.Critical) critical (<= $CriticalDays days), $($counts.Warning) warning (<= $WarningDays days).")
                $html = New-HtmlReport -Title 'App secrets and certificates expiry' -IntroLines $intro -Sections @(
                    @{ Heading = "Needs attention ($($attention.Count))"
                       Html    = ConvertTo-HtmlTable -Rows $attention -Columns @('Status', 'DaysLeft', 'Expires', 'App', 'CredentialType', 'Owners') -StatusColumn 'Status' -EmptyText 'Nothing expiring.' }
                )
                $state = if ($counts.Expired -or $counts.Critical) { 'CRITICAL' } elseif ($counts.Warning) { 'Warning' } else { 'All OK' }
                if (Send-ToolkitMail -Config $config -Subject "[$state] App credentials: $($counts.Expired) expired, $($counts.Critical) critical, $($counts.Warning) warning" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
            }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
