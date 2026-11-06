<#
.SYNOPSIS
    Finds enterprise apps with risky Microsoft Graph, Exchange or SharePoint permissions.
.DESCRIPTION
    Rates each app's application and delegated permissions, and flags unverified publishers,
    apps consented by individual users and recently added apps (possible consent phishing). Read-only.
.EXAMPLE
    .\Get-RiskyAppPermissionReport.ps1
.EXAMPLE
    .\Get-RiskyAppPermissionReport.ps1 -MinimumRisk High -SendEmail
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

    [ValidateSet('Low', 'Medium', 'High', 'Critical')]
    [string]$MinimumRisk = 'Low',
    [switch]$IncludeMicrosoftApps,

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

$recentDays = [int](Get-Setting $config 'AppPermissions' 'RecentDays' 30)
$scopes = @('Application.Read.All', 'Directory.Read.All')

# Microsoft's own tenants - their first-party apps are skipped unless -IncludeMicrosoftApps
$microsoftTenants = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47', '33e01921-4d64-4f8c-a055-5bdaffd5e33d')

# --- Permission risk ----------------------------------------------------------

$criticalPermissions = '^(RoleManagement\.ReadWrite\.Directory|AppRoleAssignment\.ReadWrite\.All|Application\.ReadWrite\.All|Directory\.ReadWrite\.All|' +
                       'UserAuthenticationMethod\.ReadWrite\.All|Policy\.ReadWrite\.AuthenticationMethod|Domain\.ReadWrite\.All|' +
                       'Policy\.ReadWrite\.ConditionalAccess|PrivilegedAccess\.ReadWrite\.AzureAD|Exchange\.ManageAsApp|full_access_as_app|Sites\.FullControl\.All)$'
$highPermissions     = '^(Mail\.|MailboxSettings\.ReadWrite|Files\.(Read|ReadWrite)\.All|Sites\.(Read|ReadWrite|Manage)\.All|User\.ReadWrite\.All|' +
                       'Group\.ReadWrite\.All|GroupMember\.ReadWrite\.All|Chat\.|ChannelMessage\.|Calendars\.ReadWrite|Contacts\.ReadWrite|' +
                       'EWS\.AccessAsUser\.All|Directory\.AccessAsUser\.All|DeviceManagement.*ReadWrite|Policy\.ReadWrite|IdentityRiskyUser\.ReadWrite|' +
                       'Notes\.ReadWrite\.All|TeamMember\.ReadWrite\.All|User\.Invite\.All|AuditLog\.Read\.All|SecurityEvents\.ReadWrite\.All|' +
                       'Application\.ReadWrite\.OwnedBy|User-PasswordProfile\.ReadWrite\.All|UserAuthenticationMethod\.ReadWrite)'
$lowPermissions      = '^(User\.Read|openid|profile|email|offline_access|User\.ReadBasic\.All|Calendars\.Read|Contacts\.Read|People\.Read)$'

$riskRank = @{ None = 0; Low = 1; Medium = 2; High = 3; Critical = 4 }
$riskName = @('None', 'Low', 'Medium', 'High', 'Critical')

function Get-PermissionRisk([string]$Permission, [string]$Type) {
    $level = if ($Permission -match $criticalPermissions) { 4 }
             elseif ($Permission -match $highPermissions) { 3 }
             elseif ($Permission -match $lowPermissions) { 1 }
             elseif ($Permission -match 'ReadWrite|\.All$|Manage|Send') { 2 }
             else { 1 }
    # Delegated permissions only act as the signed-in user, so they rate one level lower
    if ($Type -eq 'Delegated' -and $level -gt 1) { $level-- }
    return $riskName[$level]
}

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'Global Reader'
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'AppPermissions' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        $tenantName = Get-GraphTenantName
        $myTenant   = "$((Get-MgContext).TenantId)"
        Write-Host "Tenant : $tenantName"

        # --- Collect -----------------------------------------------------

        Write-Host 'Getting enterprise apps...'
        $allSps = @(Get-GraphAll 'v1.0/servicePrincipals?$select=id,appId,displayName,appOwnerOrganizationId,servicePrincipalType,accountEnabled,createdDateTime,publisherName,verifiedPublisher,homepage&$top=999')
        $spById = @{}; foreach ($sp in $allSps) { $spById[$sp.id] = $sp }

        $candidates = @($allSps | Where-Object {
            $_.servicePrincipalType -eq 'Application' -and ($IncludeMicrosoftApps -or $microsoftTenants -notcontains "$($_.appOwnerOrganizationId)")
        })
        Write-Host ("  {0} enterprise app(s), checking {1}{2}" -f $allSps.Count, $candidates.Count, $(if (-not $IncludeMicrosoftApps) { ' (Microsoft first-party apps skipped)' } else { '' }))

        $resourceRoles = @{}   # resource SP id -> @{ appRoleId -> value }
        function Get-AppRoleName([string]$ResourceId, [string]$AppRoleId) {
            if (-not $resourceRoles.ContainsKey($ResourceId)) {
                $map = @{}
                try { foreach ($r in @((Invoke-Graph "v1.0/servicePrincipals/$ResourceId`?`$select=appRoles").appRoles)) { $map["$($r.id)"] = $r.value } }
                catch { $errors.Add([pscustomobject]@{ Section = 'Resource roles'; Error = "${ResourceId}: $($_.Exception.Message)" }) }
                $resourceRoles[$ResourceId] = $map
            }
            $name = $resourceRoles[$ResourceId][$AppRoleId]
            if ($name) { return $name }
            if ($AppRoleId -eq '00000000-0000-0000-0000-000000000000') { return '(default access)' }
            return $AppRoleId
        }

        $perms = [System.Collections.Generic.List[object]]::new()

        Write-Host 'Getting application permissions...'
        $i = 0
        foreach ($sp in $candidates) {
            $i++
            if ($i % 10 -eq 0 -or $i -eq 1) { Write-Progress -Activity 'Reading application permissions' -Status "$i of $($candidates.Count)" -PercentComplete ($i / [math]::Max(1, $candidates.Count) * 100) }
            try {
                foreach ($a in (Get-GraphAll "v1.0/servicePrincipals/$($sp.id)/appRoleAssignments")) {
                    $perm = Get-AppRoleName $a.resourceId $a.appRoleId
                    if ($perm -eq '(default access)') { continue }
                    $perms.Add([pscustomobject]@{
                        SpId = $sp.id; Type = 'Application'; Resource = $a.resourceDisplayName; Permission = $perm
                        Risk = Get-PermissionRisk $perm 'Application'; ConsentedBy = 'Admin'; Users = $null; Granted = ConvertTo-LocalTime $a.createdDateTime
                    })
                }
            }
            catch { $errors.Add([pscustomobject]@{ Section = 'Application permissions'; Error = "$($sp.displayName): $($_.Exception.Message)" }) }
        }
        Write-Progress -Activity 'Reading application permissions' -Completed

        Write-Host 'Getting delegated permissions...'
        $candidateIds = @{}; foreach ($sp in $candidates) { $candidateIds[$sp.id] = $true }
        try {
            $grants = @(Get-GraphAll 'v1.0/oauth2PermissionGrants' | Where-Object { $candidateIds.ContainsKey($_.clientId) })
            foreach ($group in ($grants | Group-Object clientId, resourceId)) {
                $first    = $group.Group[0]
                $resource = $spById[$first.resourceId]
                $scopeMap = @{}   # permission -> @{ Admin = bool; Users = set }
                foreach ($g in $group.Group) {
                    foreach ($scope in ("$($g.scope)" -split '\s+' | Where-Object { $_ })) {
                        if (-not $scopeMap.ContainsKey($scope)) { $scopeMap[$scope] = @{ Admin = $false; Users = @{} } }
                        if ($g.consentType -eq 'AllPrincipals') { $scopeMap[$scope].Admin = $true } else { $scopeMap[$scope].Users["$($g.principalId)"] = $true }
                    }
                }
                foreach ($scope in $scopeMap.Keys) {
                    $entry = $scopeMap[$scope]
                    $by = if ($entry.Admin) { 'Admin (all users)' } else { "$($entry.Users.Count) user(s)" }
                    $perms.Add([pscustomobject]@{
                        SpId = $first.clientId; Type = 'Delegated'; Resource = if ($resource) { $resource.displayName } else { $first.resourceId }; Permission = $scope
                        Risk = Get-PermissionRisk $scope 'Delegated'; ConsentedBy = $by; Users = if ($entry.Admin) { $null } else { $entry.Users.Count }; Granted = $null
                    })
                }
            }
        }
        catch { $errors.Add([pscustomobject]@{ Section = 'Delegated permissions'; Error = $_.Exception.Message }) }

        # --- Per app ------------------------------------------------------

        $apps = foreach ($group in ($perms | Group-Object SpId)) {
            $sp = $spById[$group.Name]
            $p  = $group.Group
            $level = ($p | ForEach-Object { $riskRank[$_.Risk] } | Measure-Object -Maximum).Maximum
            $owner = if ("$($sp.appOwnerOrganizationId)" -eq $myTenant) { 'This tenant' } elseif ($microsoftTenants -contains "$($sp.appOwnerOrganizationId)") { 'Microsoft' } else { 'External' }
            $verified = [bool]$sp.verifiedPublisher.displayName
            $created  = ConvertTo-LocalTime $sp.createdDateTime
            $userConsents = @($p | Where-Object { $_.Users } | ForEach-Object { $_.Users } | Measure-Object -Maximum).Maximum

            $issues = @()
            if (@($p | Where-Object { $_.Risk -eq 'Critical' }).Count) { $issues += 'Has permissions that could take over the tenant or read all mail/files' }
            if ($owner -eq 'External' -and -not $verified -and $level -ge 2) { $issues += 'Unverified publisher'; $level = [math]::Min(4, $level + 1) }
            if ($userConsents -and @($p | Where-Object { $_.Users -and $riskRank[$_.Risk] -ge 2 }).Count) { $issues += "Consented by $userConsents individual user(s) - check it is a known app" }
            if ($created -and ($now - $created).TotalDays -le $recentDays -and $level -ge 3) { $issues += "Added $([int]($now - $created).TotalDays) day(s) ago with high-risk permissions - possible consent phishing" }
            if ($sp.accountEnabled -eq $false) { $issues += 'Disabled (sign-in blocked)' }

            [pscustomobject]@{
                Risk                   = $riskName[$level]
                App                    = $sp.displayName
                Publisher              = if ($sp.verifiedPublisher.displayName) { $sp.verifiedPublisher.displayName } elseif ($sp.publisherName) { $sp.publisherName } else { '' }
                VerifiedPublisher      = $verified
                Owner                  = $owner
                Created                = $created
                Enabled                = [bool]$sp.accountEnabled
                ApplicationPermissions = (@($p | Where-Object Type -eq 'Application' | Sort-Object { -$riskRank[$_.Risk] }, Permission | ForEach-Object { "$($_.Permission) ($($_.Risk))" }) -join ', ')
                DelegatedPermissions   = (@($p | Where-Object Type -eq 'Delegated' | Sort-Object { -$riskRank[$_.Risk] }, Permission | ForEach-Object { "$($_.Permission) ($($_.Risk))" }) -join ', ')
                UserConsents           = if ($userConsents) { $userConsents } else { 0 }
                Issues                 = ($issues -join '; ')
                AppId                  = $sp.appId
                SpId                   = $sp.id
            }
        }
        $minLevel   = $riskRank[$MinimumRisk]
        $apps       = @($apps | Where-Object { $riskRank[$_.Risk] -ge $minLevel } | Sort-Object { -$riskRank[$_.Risk] }, App)
        $keep       = @{}; foreach ($a in $apps) { $keep[$a.SpId] = $a }
        $permRows   = @($perms | Where-Object { $keep.ContainsKey($_.SpId) } | ForEach-Object {
            $a = $keep[$_.SpId]
            [pscustomobject]@{
                Risk = $_.Risk; App = $a.App; Type = $_.Type; Resource = $_.Resource; Permission = $_.Permission
                ConsentedBy = $_.ConsentedBy; Granted = $_.Granted; Publisher = $a.Publisher; VerifiedPublisher = $a.VerifiedPublisher; Owner = $a.Owner; AppId = $a.AppId
            }
        } | Sort-Object { -$riskRank[$_.Risk] }, App, Type, Permission)

        # --- Export ------------------------------------------------------

        $counts = [ordered]@{}
        foreach ($r in 'Critical', 'High', 'Medium', 'Low') { $counts[$r] = @($apps | Where-Object { $_.Risk -eq $r }).Count }
        $summary = [ordered]@{
            'Tenant'                       = $tenantName
            'Run at'                       = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Enterprise apps checked'      = "$($candidates.Count) of $($allSps.Count)" + $(if (-not $IncludeMicrosoftApps) { ' (Microsoft first-party apps skipped)' } else { '' })
            'Apps with permissions listed' = "$($apps.Count) (minimum risk: $MinimumRisk)"
            'Apps: Critical'               = $counts.Critical
            'Apps: High'                   = $counts.High
            'Apps: Medium'                 = $counts.Medium
            'Apps: Low'                    = $counts.Low
            'Unverified external apps'     = @($apps | Where-Object { $_.Owner -eq 'External' -and -not $_.VerifiedPublisher }).Count
            'Apps consented by users'      = @($apps | Where-Object { $_.UserConsents -gt 0 }).Count
            'How risk is rated'            = 'Critical: can take over the tenant or read all mail/files. High: broad access to mail, files, chats or directory. Delegated permissions rate one level lower (they act as the signed-in user). Unverified external publishers are raised one level.'
            'Run by'                       = Get-GraphRunBy
        }

        $saved = Save-Workbook -FilePath $reportPath -Sheets @(
            @{ Name = 'Apps'; Rows = $apps
               Headers = @('Risk', 'App', 'Publisher', 'VerifiedPublisher', 'Owner', 'Created', 'Enabled', 'ApplicationPermissions', 'DelegatedPermissions', 'UserConsents', 'Issues', 'AppId')
               Widths  = @(10, 34, 26, 17, 12, 18, 9, 70, 70, 13, 70, 38) }
            @{ Name = 'Permissions'; Rows = $permRows
               Headers = @('Risk', 'App', 'Type', 'Resource', 'Permission', 'ConsentedBy', 'Granted', 'Publisher', 'VerifiedPublisher', 'Owner', 'AppId')
               Widths  = @(10, 34, 12, 30, 36, 18, 18, 26, 17, 12, 38) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(24, 120) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(28, 110) }
        )
        if (-not $saved) { exit 1 }

        Write-Host ""
        Write-Host ("Apps: {0} critical, {1} high, {2} medium, {3} low" -f $counts.Critical, $counts.High, $counts.Medium, $counts.Low) -ForegroundColor $(if ($counts.Critical + $counts.High) { 'Red' } elseif ($counts.Medium) { 'Yellow' } else { 'Green' })
        $apps | Where-Object { $_.Risk -in 'Critical', 'High' } | Select-Object -First 10 | ForEach-Object { Write-Host ("  {0,-9} {1}" -f $_.Risk, $_.App) }
        if ($errors.Count) { Write-Host ("Some data was unavailable - see the 'Errors' sheet ({0})." -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report : $reportPath"

        if ($SendEmail) {
            $top = @($apps | Where-Object { $_.Risk -in 'Critical', 'High' })
            $html = New-HtmlReport -Title 'Risky app permissions' -IntroLines @("$tenantName - $($counts.Critical) critical and $($counts.High) high-risk app(s). The attached workbook lists every permission.") -Sections @(
                @{ Heading = "Critical and high-risk apps ($($top.Count))"; Html = ConvertTo-HtmlTable -Rows $top -Columns @('Risk', 'App', 'Publisher', 'Owner', 'ApplicationPermissions', 'Issues') -StatusColumn 'Risk' -EmptyText 'None.' }
            )
            $state = if ($counts.Critical) { 'CRITICAL' } elseif ($counts.High) { 'HIGH' } else { 'OK' }
            if (Send-ToolkitMail -Config $config -Subject "[$state] App permissions - $($counts.Critical) critical, $($counts.High) high" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
