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
    [string[]]$MailTo
)

#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)

$script:ToolkitRoot   = $PSScriptRoot

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)

$script:HtmlFont = 'font-family:Segoe UI,Arial,sans-serif;font-size:13px;'

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
