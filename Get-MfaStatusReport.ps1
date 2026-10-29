<#
.SYNOPSIS
    Reports MFA status for every Microsoft 365 user to Excel.
.DESCRIPTION
    Lists sign-in status, last sign-in, creation date, MFA registration and methods, admin roles,
    and flags users who need attention. Read-only.
.EXAMPLE
    .\Get-MfaStatusReport.ps1
.EXAMPLE
    .\Get-MfaStatusReport.ps1 -IncludeGuests -IncludePerUserMfaState -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: Microsoft.Graph.Authentication module, Global Reader (or Reports Reader + Security Reader)
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [switch]$IncludeGuests,
    [switch]$IncludePerUserMfaState,
    [switch]$UseMethodsApi,

    [ValidateRange(1, 3650)]
    [int]$StaleDays,

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

if (-not $PSBoundParameters.ContainsKey('TenantId'))              { $TenantId              = Get-Setting $config 'M365' 'TenantId' '' }
if (-not $PSBoundParameters.ContainsKey('ClientId'))              { $ClientId              = Get-Setting $config 'M365' 'ClientId' '' }
if (-not $PSBoundParameters.ContainsKey('CertificateThumbprint')) { $CertificateThumbprint = Get-Setting $config 'M365' 'CertificateThumbprint' '' }
if (-not $PSBoundParameters.ContainsKey('IncludeGuests'))         { $IncludeGuests         = [bool](Get-Setting $config 'M365' 'IncludeGuests' $false) }
if (-not $PSBoundParameters.ContainsKey('StaleDays'))             { $StaleDays             = [int](Get-Setting $config 'M365' 'StaleDays' 90) }
$newAccountDays = [int](Get-Setting $config 'M365' 'NewAccountGraceDays' 14)

$scopes = @('User.Read.All', 'AuditLog.Read.All', 'UserAuthenticationMethod.Read.All', 'Policy.Read.All', 'RoleManagement.Read.Directory')

# --- Graph helpers --------------------------------------------------------

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

# --- Authentication method names ------------------------------------------

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
$defaultMethodNames = @{
    push = 'Authenticator push'; oath = 'OATH code'; sms = 'SMS'; voiceMobile = 'Phone call (mobile)'
    voiceAlternateMobile = 'Phone call (alternate)'; voiceOffice = 'Phone call (office)'; none = ''
}
$phoneMethods     = @('mobilePhone', 'alternateMobilePhone', 'officePhone')
$appMethods       = @('microsoftAuthenticatorPush', 'microsoftAuthenticatorPasswordless', 'softwareOneTimePasscode', 'hardwareOneTimePasscode')
$phishResistant   = '^(fido2|windowsHello|passKey|macOsSecureEnclave|x509|certificate)'
$passwordlessRx   = '^(fido2|windowsHello|passKey|macOsSecureEnclave|microsoftAuthenticatorPasswordless)'

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

# --- Main -----------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'The Microsoft Graph PowerShell module is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # --- Connect ----------------------------------------------------------

    $connectedHere = $false
    $context = Get-MgContext
    if ($context -and (-not $TenantId -or $context.TenantId -eq $TenantId -or $TenantId -like '*.*')) {
        Write-Host "Using existing Graph connection ($($context.Account)$($context.AppName))."
        if ($context.AuthType -eq 'Delegated') {
            $missing = @($scopes | Where-Object { $context.Scopes -notcontains $_ })
            if ($missing.Count) { Write-Warning "Current connection is missing scope(s): $($missing -join ', '). Some sections may be empty. Run Disconnect-MgGraph and try again." }
        }
    }
    else {
        $connect = @{ NoWelcome = $true; ErrorAction = 'Stop' }
        if ($TenantId) { $connect.TenantId = $TenantId }
        if ($ClientId -and $CertificateThumbprint) {
            if (-not $TenantId) { throw 'App-only sign-in needs a tenant. Use -TenantId.' }
            $connect.ClientId = $ClientId
            $connect.CertificateThumbprint = $CertificateThumbprint
            Write-Host 'Connecting to Microsoft Graph (app-only, certificate)...'
        }
        else {
            $connect.Scopes = $scopes
            if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
            Write-Host 'Connecting to Microsoft Graph - sign in with an account that has Global Reader (or Reports Reader + Security Reader)...'
        }
        Connect-MgGraph @connect | Out-Null
        $connectedHere = $true
        $context = Get-MgContext
    }

    $now        = Get-Date
    $reportPath = New-ReportPath -Config $config -Prefix 'MfaStatus' -OutputFile $OutputFile
    $errors     = [System.Collections.Generic.List[object]]::new()
    function Add-ReportError([string]$Section, [string]$Message) {
        $errors.Add([pscustomobject]@{ Section = $Section; Error = $Message })
    }

    try {
        # --- Tenant -------------------------------------------------------

        $tenantName = $context.TenantId
        try {
            $org = @(Get-GraphAll 'v1.0/organization?$select=displayName,verifiedDomains') | Select-Object -First 1
            $defaultDomain = @($org.verifiedDomains | Where-Object { $_.isDefault }) | Select-Object -First 1
            $tenantName = "$($org.displayName) ($($defaultDomain.name))"
        }
        catch { Add-ReportError 'Tenant' $_.Exception.Message }
        Write-Host "Tenant : $tenantName"

        # --- Users --------------------------------------------------------

        Write-Host 'Getting users...'
        $select = 'id,displayName,userPrincipalName,mail,accountEnabled,userType,createdDateTime,assignedLicenses,onPremisesSyncEnabled,department,jobTitle'
        $signInAvailable = $true
        try {
            $users = @(Get-GraphAll ("v1.0/users?`$select=$select,signInActivity&`$top=500"))
        }
        catch {
            $signInAvailable = $false
            Add-ReportError 'Sign-in activity' "Last sign-in dates unavailable (needs Entra ID P1/P2 and AuditLog.Read.All): $($_.Exception.Message)"
            Write-Warning 'Last sign-in dates are unavailable (needs Entra ID P1/P2 and AuditLog.Read.All). Continuing without them.'
            $users = @(Get-GraphAll ("v1.0/users?`$select=$select&`$top=999"))
        }
        $allUserCount = $users.Count
        if (-not $IncludeGuests) { $users = @($users | Where-Object { $_.userType -ne 'Guest' }) }
        Write-Host ("  {0:N0} user(s){1}" -f $users.Count, $(if (-not $IncludeGuests) { ' (members only - use -IncludeGuests for guests)' } else { '' }))

        $userById = @{}
        foreach ($u in $users) { $userById[$u.id] = $u }

        # --- MFA registration --------------------------------------------

        $registration = @{}
        $source = 'Registration report'
        if (-not $UseMethodsApi) {
            Write-Host 'Getting MFA registration details...'
            try {
                foreach ($r in (Get-GraphAll 'v1.0/reports/authenticationMethods/userRegistrationDetails')) { $registration[$r.id] = $r }
            }
            catch {
                Add-ReportError 'Registration report' "Unavailable (needs Entra ID P1/P2 and AuditLog.Read.All), so each user's methods were read one by one: $($_.Exception.Message)"
                Write-Warning 'MFA registration report unavailable (needs Entra ID P1/P2). Reading each user''s methods instead - this is slower.'
            }
        }
        $needMethods = @($users | Where-Object { -not $registration.ContainsKey($_.id) })
        $methodsById = @{}
        if ($needMethods.Count) {
            if ($registration.Count) { $source = "Registration report, plus methods API for $($needMethods.Count) user(s) not in the report" }
            else                     { $source = 'Authentication methods API (per user)' }
            $i = 0
            foreach ($u in $needMethods) {
                $i++
                if ($i % 10 -eq 0 -or $i -eq 1) { Write-Progress -Activity 'Reading authentication methods' -Status "$i of $($needMethods.Count)" -PercentComplete ($i / $needMethods.Count * 100) }
                try {
                    $methodsById[$u.id] = @(Get-GraphAll "v1.0/users/$($u.id)/authentication/methods" | ForEach-Object { ConvertFrom-MethodObject $_ } | Where-Object { $_ } | Sort-Object -Unique)
                }
                catch {
                    Add-ReportError 'Authentication methods' "$($u.userPrincipalName): $($_.Exception.Message)"
                }
            }
            Write-Progress -Activity 'Reading authentication methods' -Completed
        }

        # --- Per-user MFA (legacy) ----------------------------------------

        $perUserMfa = @{}
        if ($IncludePerUserMfaState) {
            $i = 0
            foreach ($u in $users) {
                $i++
                if ($i % 10 -eq 0 -or $i -eq 1) { Write-Progress -Activity 'Reading per-user MFA state' -Status "$i of $($users.Count)" -PercentComplete ($i / $users.Count * 100) }
                try   { $perUserMfa[$u.id] = "$((Invoke-Graph "beta/users/$($u.id)/authentication/requirements").perUserMfaState)" }
                catch { $perUserMfa[$u.id] = 'Unknown'; Add-ReportError 'Per-user MFA' "$($u.userPrincipalName): $($_.Exception.Message)" }
            }
            Write-Progress -Activity 'Reading per-user MFA state' -Completed
        }

        # --- Admin roles --------------------------------------------------

        Write-Host 'Getting admin roles...'
        $rolesById = @{}   # user id -> list of role names
        try {
            $roleNames = @{}
            foreach ($d in (Get-GraphAll 'v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName')) { $roleNames[$d.id] = $d.displayName }

            $assignments = @()
            $assignments += @(Get-GraphAll 'v1.0/roleManagement/directory/roleAssignments?$select=principalId,roleDefinitionId' |
                              ForEach-Object { [pscustomobject]@{ PrincipalId = $_.principalId; Role = $roleNames[$_.roleDefinitionId]; Eligible = $false } })
            try {
                $assignments += @(Get-GraphAll 'v1.0/roleManagement/directory/roleEligibilitySchedules?$select=principalId,roleDefinitionId' |
                                  ForEach-Object { [pscustomobject]@{ PrincipalId = $_.principalId; Role = $roleNames[$_.roleDefinitionId]; Eligible = $true } })
            }
            catch { Add-ReportError 'Eligible roles (PIM)' "Not included (needs Entra ID P2): $($_.Exception.Message)" }

            $groupMembers = @{}
            foreach ($a in $assignments) {
                if (-not $a.Role) { continue }
                $label = if ($a.Eligible) { "$($a.Role) (eligible)" } else { $a.Role }
                $targets = @()
                if ($userById.ContainsKey($a.PrincipalId)) {
                    $targets = @($a.PrincipalId)
                }
                else {
                    # Role-assignable group (expand its members); service principals and excluded users are skipped
                    if (-not $groupMembers.ContainsKey($a.PrincipalId)) {
                        try   { $groupMembers[$a.PrincipalId] = @(Get-GraphAll "v1.0/groups/$($a.PrincipalId)/transitiveMembers/microsoft.graph.user?`$select=id" | ForEach-Object { $_.id }) }
                        catch { $groupMembers[$a.PrincipalId] = @() }
                    }
                    $targets = $groupMembers[$a.PrincipalId]
                    if ($targets.Count) { $label += ' (via group)' }
                }
                foreach ($t in $targets) {
                    if (-not $rolesById.ContainsKey($t)) { $rolesById[$t] = [System.Collections.Generic.List[string]]::new() }
                    if (-not $rolesById[$t].Contains($label)) { $rolesById[$t].Add($label) }
                }
            }
        }
        catch { Add-ReportError 'Admin roles' $_.Exception.Message }

        # --- Tenant MFA policies -----------------------------------------

        Write-Host 'Checking security defaults and Conditional Access...'
        $securityDefaults = 'Unknown'
        try {
            $sd = Invoke-Graph 'v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
            $securityDefaults = if ($sd.isEnabled) { 'Enabled' } else { 'Disabled' }
        }
        catch { Add-ReportError 'Security defaults' $_.Exception.Message }

        $caRows = @()
        try {
            $caRows = @(Get-GraphAll 'v1.0/identity/conditionalAccess/policies' | ForEach-Object {
                $gc = $_.grantControls
                $mfa = ($gc -and ((@($gc.builtInControls) -contains 'mfa') -or $gc.authenticationStrength))
                if (-not $mfa) { return }
                $cu   = $_.conditions.users
                $capps = $_.conditions.applications
                $describe = {
                    param($List, [string]$Noun)
                    $items = @($List | Where-Object { $_ })
                    if ($items -contains 'All') { "All $Noun" } elseif ($items -contains 'None') { 'None' } elseif ($items.Count) { "$($items.Count) $Noun" } else { '' }
                }
                $include = @((& $describe $cu.includeUsers 'users'), (& $describe $cu.includeGroups 'groups'), (& $describe $cu.includeRoles 'roles')) | Where-Object { $_ }
                $exclude = @($cu.excludeUsers | Where-Object { $_ }).Count + @($cu.excludeGroups | Where-Object { $_ }).Count + @($cu.excludeRoles | Where-Object { $_ }).Count
                [pscustomobject]@{
                    Policy        = $_.displayName
                    State         = switch ($_.state) { 'enabled' { 'Enabled' } 'enabledForReportingButNotEnforced' { 'Report-only' } default { 'Disabled' } }
                    Requires      = if ($gc.authenticationStrength) { "Authentication strength: $($gc.authenticationStrength.displayName)" } else { 'MFA' }
                    AppliesTo     = ($include -join ', ')
                    Exclusions    = $exclude
                    Applications  = (& $describe $capps.includeApplications 'apps')
                    Modified      = ConvertTo-LocalTime $(if ($_.modifiedDateTime) { $_.modifiedDateTime } else { $_.createdDateTime })
                }
            })
        }
        catch { Add-ReportError 'Conditional Access' "Could not read policies (needs Policy.Read.All and Entra ID P1): $($_.Exception.Message)" }
        $caEnabled = @($caRows | Where-Object { $_.State -eq 'Enabled' }).Count

        # --- Build rows --------------------------------------------------

        Write-Host 'Building report...'
        $riskRank = @{ OK = 0; Low = 1; Medium = 2; High = 3; Critical = 4 }
        function Add-Issue([string]$Level, [string]$Text) {
            # Uses $issues from the loop below and raises $script:currentRisk to the highest level seen
            $issues.Add($Text)
            if ($riskRank[$Level] -gt $riskRank[$script:currentRisk]) { $script:currentRisk = $Level }
        }
        $rows = foreach ($u in $users) {
            $reg = $registration[$u.id]
            if ($reg) {
                $methods      = @($reg.methodsRegistered | Where-Object { $_ })
                $mfaRegistered = [bool]$reg.isMfaRegistered
                $mfaCapable    = [bool]$reg.isMfaCapable
                $passwordless  = [bool]$reg.isPasswordlessCapable
                $ssprRegistered = if ($null -ne $reg.isSsprRegistered) { [bool]$reg.isSsprRegistered } else { $null }
                $preferred     = $defaultMethodNames["$($reg.userPreferredMethodForSecondaryAuthentication)"]
                if ($reg.isSystemPreferredAuthenticationMethodEnabled -and @($reg.systemPreferredAuthenticationMethods).Count) {
                    $sys = $defaultMethodNames["$(@($reg.systemPreferredAuthenticationMethods)[0])"]
                    $preferred = "System-preferred ($sys)"
                }
                if (-not $mfaRegistered) { $preferred = '' }
                $isAdmin = [bool]$reg.isAdmin
            }
            elseif ($methodsById.ContainsKey($u.id)) {
                $methods       = $methodsById[$u.id]
                $mfaRegistered = [bool](@($methods | Where-Object { Test-StrongMethod $_ }).Count)
                $mfaCapable    = $null
                $passwordless  = [bool](@($methods | Where-Object { $_ -match $passwordlessRx }).Count)
                $ssprRegistered = $null
                $preferred     = ''
                $isAdmin       = $false
            }
            else {
                $methods = @(); $mfaRegistered = $null; $mfaCapable = $null; $passwordless = $null; $ssprRegistered = $null; $preferred = ''; $isAdmin = $false
            }

            $roles   = if ($rolesById.ContainsKey($u.id)) { @($rolesById[$u.id] | Sort-Object) } else { @() }
            $isAdmin = $isAdmin -or ($roles.Count -gt 0)
            $strong  = @($methods | Where-Object { Test-StrongMethod $_ })
            $phishOk = [bool](@($methods | Where-Object { $_ -match $phishResistant }).Count)
            $phoneOnly = $strong.Count -gt 0 -and -not @($strong | Where-Object { $phoneMethods -notcontains $_ }).Count

            $sia            = $u.signInActivity
            $lastInteractive = ConvertTo-LocalTime $sia.lastSignInDateTime
            $lastNonInter    = ConvertTo-LocalTime $sia.lastNonInteractiveSignInDateTime
            $lastSuccessful  = ConvertTo-LocalTime $sia.lastSuccessfulSignInDateTime
            $lastSignIn      = if ($lastSuccessful) { $lastSuccessful } else { Get-Latest @($lastInteractive, $lastNonInter) }
            $created         = ConvertTo-LocalTime $u.createdDateTime
            $daysSince       = if ($lastSignIn) { [int][math]::Floor(($now - $lastSignIn).TotalDays) } else { $null }
            $isNew           = $created -and ($now - $created).TotalDays -le $newAccountDays

            # Risk rating (enabled accounts only)
            $issues = [System.Collections.Generic.List[string]]::new()
            $script:currentRisk = 'OK'
            if (-not $u.accountEnabled) {
                $script:currentRisk = 'Disabled'
            }
            else {
                if ($null -eq $mfaRegistered) { Add-Issue 'Medium' 'MFA status could not be read' }
                elseif (-not $mfaRegistered) {
                    if ($isAdmin) { Add-Issue 'Critical' 'Admin with no MFA registered' }
                    elseif ($u.userType -eq 'Guest') { Add-Issue 'Low' 'Guest with no MFA registered here - guests usually use MFA from their own tenant; check a Conditional Access policy requires MFA for guests' }
                    else          { Add-Issue 'High' $(if ($isNew) { "No MFA registered (new account - created $([int]($now - $created).TotalDays) day(s) ago)" } else { 'No MFA registered' }) }
                }
                else {
                    if ($phoneOnly) { Add-Issue $(if ($isAdmin) { 'High' } else { 'Medium' }) 'Only SMS/voice MFA registered' }
                    if ($mfaCapable -eq $false) { Add-Issue 'Medium' 'MFA method registered but not allowed by policy' }
                    if ($isAdmin -and -not $phishOk -and -not $phoneOnly) { Add-Issue 'Low' 'Admin without a phishing-resistant method' }
                }
                if ($signInAvailable) {
                    if (-not $lastSignIn -and -not $isNew) { Add-Issue 'Low' 'Never signed in - consider disabling' }
                    elseif ($null -ne $daysSince -and $daysSince -gt $StaleDays) { Add-Issue 'Low' "No sign-in for $daysSince days - consider disabling" }
                }
            }
            $risk = $script:currentRisk

            [pscustomobject]@{
                DisplayName              = $u.displayName
                UserPrincipalName        = $u.userPrincipalName
                Email                    = $u.mail
                SignInStatus             = if ($u.accountEnabled) { 'Allowed' } else { 'Blocked' }
                UserType                 = $u.userType
                Licensed                 = @($u.assignedLicenses).Count -gt 0
                AccountSource            = if ($u.onPremisesSyncEnabled) { 'Synced (AD)' } else { 'Cloud' }
                Created                  = $created
                LastSignIn               = if ($signInAvailable) { $lastSignIn } else { 'Unavailable' }
                DaysSinceSignIn          = $daysSince
                LastInteractiveSignIn    = $lastInteractive
                LastNonInteractiveSignIn = $lastNonInter
                MfaStatus                = if ($null -eq $mfaRegistered) { 'Unknown' } elseif ($mfaRegistered) { 'Registered' } else { 'Not registered' }
                MfaCapable               = if ($null -eq $mfaCapable) { '' } elseif ($mfaCapable) { 'Yes' } else { 'No' }
                PerUserMfa               = if ($IncludePerUserMfaState) { $perUserMfa[$u.id] } else { '' }
                DefaultMethod            = $preferred
                MethodsRegistered        = (@($methods | ForEach-Object { if ($methodNames.ContainsKey($_)) { $methodNames[$_] } else { $_ } }) -join ', ')
                PhishingResistant        = $phishOk
                PasswordlessCapable      = if ($null -eq $passwordless) { '' } elseif ($passwordless) { 'Yes' } else { 'No' }
                SsprRegistered           = if ($null -eq $ssprRegistered) { '' } elseif ($ssprRegistered) { 'Yes' } else { 'No' }
                IsAdmin                  = $isAdmin
                AdminRoles               = ($roles -join '; ')
                Risk                     = $risk
                Issues                   = ($issues -join '; ')
                Department               = $u.department
                JobTitle                 = $u.jobTitle
            }
        }
        $rows = @($rows)

        # --- Summaries ---------------------------------------------------

        $riskOrder = @{ Critical = 0; High = 1; Medium = 2; Low = 3; OK = 4; Disabled = 5 }
        $rows      = @($rows | Sort-Object @{ Expression = { $riskOrder[$_.Risk] } }, DisplayName)
        $attention = @($rows | Where-Object { $_.Risk -in 'Critical', 'High', 'Medium', 'Low' })
        $admins    = @($rows | Where-Object { $_.IsAdmin })
        $enabled   = @($rows | Where-Object { $_.SignInStatus -eq 'Allowed' })
        $enabledMfa = @($enabled | Where-Object { $_.MfaStatus -eq 'Registered' })
        $pct = if ($enabled.Count) { [math]::Round($enabledMfa.Count / $enabled.Count * 100, 1) } else { 0 }

        $methodCounts = @($enabled | ForEach-Object { $_.MethodsRegistered -split ', ' } | Where-Object { $_ } |
                          Group-Object | Sort-Object Count -Descending |
                          ForEach-Object { [pscustomobject]@{ Method = $_.Name; Users = $_.Count; PercentOfEnabled = if ($enabled.Count) { [math]::Round($_.Count / $enabled.Count * 100, 1) } else { 0 } } })

        $counts = [ordered]@{}
        foreach ($r in 'Critical', 'High', 'Medium', 'Low') { $counts[$r] = @($rows | Where-Object { $_.Risk -eq $r }).Count }

        $summary = [ordered]@{
            'Tenant'                           = $tenantName
            'Run at'                           = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Users in report'                  = "$($rows.Count) of $allUserCount" + $(if (-not $IncludeGuests) { ' (guests excluded)' } else { '' })
            'Sign-in allowed'                  = $enabled.Count
            'Sign-in blocked'                  = $rows.Count - $enabled.Count
            'Enabled users with MFA'           = "$($enabledMfa.Count) ($pct%)"
            'Enabled users without MFA'        = @($enabled | Where-Object { $_.MfaStatus -eq 'Not registered' }).Count
            'Admins'                           = $admins.Count
            'Admins without MFA'               = @($admins | Where-Object { $_.SignInStatus -eq 'Allowed' -and $_.MfaStatus -ne 'Registered' }).Count
            'Phishing-resistant users'         = @($enabled | Where-Object { $_.PhishingResistant }).Count
            'Risk: Critical'                   = $counts.Critical
            'Risk: High'                       = $counts.High
            'Risk: Medium'                     = $counts.Medium
            'Risk: Low'                        = $counts.Low
            'Security defaults'                = $securityDefaults
            'Conditional Access requiring MFA' = "$caEnabled enabled, $(@($caRows | Where-Object { $_.State -eq 'Report-only' }).Count) report-only"
            'MFA data source'                  = $source
            'Last sign-in data'                = if ($signInAvailable) { "Available (no sign-in for $StaleDays+ days is flagged)" } else { 'Unavailable - needs Entra ID P1/P2' }
            'Per-user MFA (legacy)'            = if ($IncludePerUserMfaState) { (@($rows | Group-Object PerUserMfa | ForEach-Object { "$($_.Name): $($_.Count)" }) -join ', ') } else { 'Not checked (use -IncludePerUserMfaState)' }
            'Note'                             = '"Registered" means the user has set up a strong MFA method. Whether MFA is actually required at sign-in depends on security defaults or Conditional Access.'
            'Run by'                           = if ($context.Account) { $context.Account } else { "App: $($context.AppName)" }
        }

        # --- Export ------------------------------------------------------

        $mainHeaders = @('DisplayName', 'UserPrincipalName', 'Email', 'SignInStatus', 'LastSignIn', 'DaysSinceSignIn', 'Created', 'MfaStatus', 'MethodsRegistered',
                         'DefaultMethod', 'PhishingResistant', 'MfaCapable', 'PasswordlessCapable', 'SsprRegistered', 'IsAdmin', 'AdminRoles', 'Risk', 'Issues',
                         'UserType', 'Licensed', 'AccountSource', 'LastInteractiveSignIn', 'LastNonInteractiveSignIn', 'Department', 'JobTitle')
        $mainWidths  = @(26, 34, 32, 13, 18, 15, 18, 15, 50, 26, 17, 12, 19, 15, 9, 40, 10, 50, 10, 9, 13, 21, 24, 20, 24)
        if ($IncludePerUserMfaState) {
            $idx = [array]::IndexOf($mainHeaders, 'MfaCapable')
            $mainHeaders = $mainHeaders[0..$idx] + 'PerUserMfa' + $mainHeaders[($idx + 1)..($mainHeaders.Count - 1)]
            $mainWidths  = $mainWidths[0..$idx] + 12 + $mainWidths[($idx + 1)..($mainWidths.Count - 1)]
        }

        $sheets = @(
            @{ Name = 'MFA Status'; Rows = $rows; Headers = $mainHeaders; Widths = $mainWidths }
            @{ Name = 'Needs Attention'; Rows = $attention
               Headers = @('Risk', 'Issues', 'DisplayName', 'UserPrincipalName', 'SignInStatus', 'LastSignIn', 'Created', 'MfaStatus', 'MethodsRegistered', 'AdminRoles')
               Widths  = @(10, 55, 26, 34, 13, 18, 18, 15, 45, 40) }
            @{ Name = 'Admins'; Rows = $admins
               Headers = @('DisplayName', 'UserPrincipalName', 'AdminRoles', 'SignInStatus', 'MfaStatus', 'MethodsRegistered', 'PhishingResistant', 'LastSignIn', 'Risk', 'Issues')
               Widths  = @(26, 34, 50, 13, 15, 45, 17, 18, 10, 45) }
            @{ Name = 'Methods'; Rows = $methodCounts; Headers = @('Method', 'Users', 'PercentOfEnabled'); Widths = @(32, 10, 17) }
            @{ Name = 'Conditional Access'; Rows = $caRows
               Headers = @('Policy', 'State', 'Requires', 'AppliesTo', 'Exclusions', 'Applications', 'Modified'); Widths = @(45, 12, 40, 35, 11, 15, 18) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(24, 120) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(34, 100) }
        )
        if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

        # --- Console -----------------------------------------------------

        Write-Host ""
        Write-Host ("Enabled users with MFA : {0} of {1} ({2}%)" -f $enabledMfa.Count, $enabled.Count, $pct) -ForegroundColor $(if ($pct -ge 100) { 'Green' } elseif ($pct -ge 90) { 'Yellow' } else { 'Red' })
        Write-Host ("Admins without MFA     : {0}" -f $summary['Admins without MFA']) -ForegroundColor $(if ($summary['Admins without MFA']) { 'Red' } else { 'Green' })
        Write-Host ("Needs attention        : {0} critical, {1} high, {2} medium, {3} low" -f $counts.Critical, $counts.High, $counts.Medium, $counts.Low)
        Write-Host ("Security defaults      : {0};  CA policies requiring MFA: {1} enabled" -f $securityDefaults, $caEnabled)
        if ($securityDefaults -ne 'Enabled' -and $caEnabled -eq 0) {
            Write-Host 'WARNING: neither security defaults nor an enabled Conditional Access policy requires MFA.' -ForegroundColor Red
        }
        if ($errors.Count) { Write-Host ("Some data was unavailable - see the 'Errors' sheet ({0})." -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report                 : $reportPath"

        # --- Email -------------------------------------------------------

        if ($SendEmail) {
            $intro = @(
                "$tenantName - $($enabledMfa.Count) of $($enabled.Count) enabled users have MFA registered ($pct%).",
                "Security defaults: $securityDefaults. Conditional Access policies requiring MFA: $caEnabled enabled."
            )
            if ($securityDefaults -ne 'Enabled' -and $caEnabled -eq 0) { $intro += 'WARNING: nothing in the tenant currently requires MFA at sign-in.' }
            $html = New-HtmlReport -Title 'Microsoft 365 MFA report' -IntroLines $intro -Sections @(
                @{ Heading = "Needs attention ($($attention.Count))"
                   Html    = ConvertTo-HtmlTable -Rows @($attention | Where-Object { $_.Risk -ne 'Low' } | Select-Object -First 100) -Columns @('Risk', 'DisplayName', 'UserPrincipalName', 'MfaStatus', 'LastSignIn', 'Issues') -StatusColumn 'Risk' -EmptyText 'No critical, high or medium issues.' }
                @{ Heading = "Admins ($($admins.Count))"
                   Html    = ConvertTo-HtmlTable -Rows $admins -Columns @('DisplayName', 'AdminRoles', 'MfaStatus', 'MethodsRegistered', 'Risk') -StatusColumn 'Risk' -EmptyText 'No admins found.' }
            )
            $state = if ($counts.Critical) { 'CRITICAL' } elseif ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Medium' } else { 'OK' }
            $subject = "[$state] M365 MFA report - $($enabled.Count - $enabledMfa.Count) enabled user(s) without MFA"
            if (Send-ToolkitMail -Config $config -Subject $subject -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
