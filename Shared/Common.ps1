<#
    Shared helpers for the PowerShell scripts in this repo (config, Excel, HTML and email).
    Dot-sourced by each script; not run on its own.

    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
#>

$script:ToolkitRoot   = Split-Path -Parent $PSScriptRoot
$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)
$script:IsWindowsHost = ($PSVersionTable.PSEdition -eq 'Desktop') -or $IsWindows
$script:WarnedPlaceholders = @{}

# Values that mean "this is still the example config"
$script:PlaceholderPattern = 'example\.(com|org|net)|\bSERVER0\d\b|FILESERVER0\d|<[^>]+>|DC=example'

# --- Config ---------------------------------------------------------------

function Get-ToolkitConfig {
    <# Loads config.psd1 from the toolkit root. Returns an empty hashtable if it doesn't exist. #>
    $path = Join-Path $script:ToolkitRoot 'config.psd1'
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Warning ("config.psd1 not found in {0}. Copy config.example.psd1 to config.psd1 and fill in your values. Continuing with script parameters and built-in defaults only." -f $script:ToolkitRoot)
        return @{}
    }
    try {
        return Import-PowerShellDataFile -LiteralPath $path
    }
    catch {
        throw "config.psd1 could not be read - check it for syntax errors. Details: $($_.Exception.Message)"
    }
}

function Get-Setting {
    <#
        Reads Config[Section][Key]. Returns $Default when the setting is missing or empty.
        Warns (once per setting) if the value still looks like a placeholder.
    #>
    param(
        [hashtable]$Config,
        [string]$Section,
        [string]$Key,
        $Default = $null
    )
    $value = $null
    if ($Config -and $Config.ContainsKey($Section) -and $Config[$Section] -is [hashtable] -and $Config[$Section].ContainsKey($Key)) {
        $value = $Config[$Section][$Key]
    }
    if ($null -eq $value -or ($value -is [string] -and $value -eq '') -or ($value -is [array] -and $value.Count -eq 0)) {
        $value = $Default
    }

    $settingName = "$Section.$Key"
    if ($null -ne $value -and -not $script:WarnedPlaceholders.ContainsKey($settingName)) {
        $text = (@($value) | ForEach-Object { "$_" }) -join ' '
        if ($text -match $script:PlaceholderPattern) {
            Write-Warning "Setting $settingName still contains a placeholder value ('$text'). Update it in config.psd1."
            $script:WarnedPlaceholders[$settingName] = $true
        }
    }

    # Callers that need a list wrap the result in @(...), so a single item still works.
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

# --- Long path helpers ----------------------------------------------------
# The \\?\ prefix lets .NET read paths longer than 260 characters on Windows.

function Add-LongPathPrefix([string]$p) {
    if (-not $script:IsWindowsHost) { return $p }
    if ($p.StartsWith('\\?\')) { return $p }
    if ($p.StartsWith('\\'))   { return '\\?\UNC\' + $p.Substring(2) }
    return '\\?\' + $p
}

function Remove-LongPathPrefix([string]$p) {
    if ($p.StartsWith('\\?\UNC\')) { return '\\' + $p.Substring(8) }
    if ($p.StartsWith('\\?\'))     { return $p.Substring(4) }
    return $p
}

function Get-NormalizedRoot([string]$Path) {
    <# Full path of a folder without a trailing separator (except drive roots like C:\). #>
    $full = (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName
    if ($full.Length -gt 1) { $full = $full.TrimEnd('\', '/') }
    if ($full -match '^[A-Za-z]:$') { $full += '\' }
    if ($full -eq '') { $full = '/' }
    return $full
}

function Get-ErrorMessage($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    while ($ex.InnerException) { $ex = $ex.InnerException }
    return [pscustomobject]@{ Type = $ex.GetType().Name; Message = $ex.Message }
}

# --- Excel (.xlsx) writer -------------------------------------------------
# An .xlsx file is a zip of XML parts; these functions build one directly,
# so Excel and the ImportExcel module are not needed.

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

# --- HTML (email-friendly, inline styles) ---------------------------------

$script:HtmlFont = 'font-family:Segoe UI,Arial,sans-serif;font-size:13px;'

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

# --- Email ----------------------------------------------------------------

function Send-ToolkitMail {
    <#
        Sends an HTML email using the Email section of config.psd1.
        If Email.CredentialFile is set, the credential is loaded with Import-Clixml
        (create it with: Get-Credential | Export-Clixml .\smtp.cred.xml - it is encrypted
        for the current user on the current machine and ignored by git).
    #>
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
        Write-Warning 'Email not sent: set Email.SmtpServer, Email.From and Email.To in config.psd1.'
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

# --- Microsoft Graph ------------------------------------------------------
# Used by the M365 scripts. Needs the Microsoft.Graph.Authentication module.

function Connect-ToolkitGraph {
    <#
        Connects to Microsoft Graph, or reuses an existing connection.
        Sign-in settings come from the calling script's parameters (-Bound $PSBoundParameters) first,
        then the M365 section of the config. Returns $true if this call made the connection,
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
        if (-not $tenantId) { throw 'App-only sign-in needs a tenant. Set M365.TenantId in config.psd1 or use -TenantId.' }
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

# --- Exchange Online -------------------------------------------------------
# Used by the mailbox scripts. Needs the ExchangeOnlineManagement module.

# Folders attackers move mail into so the victim doesn't see replies
$hidingFolders   = 'RSS|Conversation History|Archive|Junk|Deleted Items|Notes|Sync Issues|Outbox'
$sensitiveWords  = 'invoice|payment|bank|wire|transfer|remit|password|credential|security|phish|hack|compromis|suspicious|helpdesk|it support|mfa|verify'

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
        calling script's parameters (-Bound $PSBoundParameters) first, then the M365 section of the config.
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
        if ($TenantId -notlike '*.*') { throw 'App-only sign-in to Exchange needs the tenant domain (e.g. contoso.onmicrosoft.com). Set M365.TenantId in config.psd1 or use -TenantId.' }
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

# --- Evidence packages (SOC scripts) ----------------------------------------

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
