<#
.SYNOPSIS
    Reports SharePoint Online sites, their sharing settings and storage, and the external users on each site.
.DESCRIPTION
    Lists every site (OneDrive optional) with owner, storage, last activity and sharing level, and flags
    sites shared with "Anyone" links, sites close to their storage limit and inactive sites. Read-only.
.EXAMPLE
    .\Get-SharePointSiteReport.ps1 -AdminUrl https://contoso-admin.sharepoint.com
.EXAMPLE
    .\Get-SharePointSiteReport.ps1 -IncludeExternalUsers
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: Microsoft.Online.SharePoint.PowerShell module, SharePoint Administrator (or Global Reader)
#>
[CmdletBinding()]
param(
    [string]$AdminUrl,
    [switch]$IncludeExternalUsers,
    [switch]$IncludeOneDrive,
    [ValidateRange(1, 3650)] [int]$InactiveDays,

    [string]$OutputFile
)

#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)

$script:ToolkitRoot   = $PSScriptRoot

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)

$script:IsWindowsHost = ($PSVersionTable.PSEdition -eq 'Desktop') -or $IsWindows

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

function Get-XlsxStatusStyle([string]$Value) {
    <# Cell style for a status value, matching the HTML badges: 2 red, 3 amber, 4 blue, 5 green, 0 plain. #>
    switch -Regex ($Value) {
        '^(Critical|Expired|Error|Unreachable|Failed|Not done|High)' { return 2 }
        '^(Warning|Expiring|Medium)'                                  { return 3 }
        '^Low'                                                        { return 4 }
        '^(OK|Done|Requested)'                                        { return 5 }
        default                                                       { return 0 }
    }
}

function Add-XlsxSheet {
    param($Zip, [int]$Index, [string[]]$Headers, $Rows, [int[]]$Widths)

    $inv     = [System.Globalization.CultureInfo]::InvariantCulture
    $letters = for ($i = 0; $i -lt $Headers.Count; $i++) { Get-ColumnLetter $i }
    # Columns whose values get status colours, and the label column of an Item/Value summary sheet
    $statusCols = @(for ($i = 0; $i -lt $Headers.Count; $i++) { if ($Headers[$i] -match '^(Status|Risk|Severity|Rating|Result|Overall)$') { $i } })
    $labelCol   = if ($Headers.Count -eq 2 -and $Headers[0] -eq 'Item' -and $Headers[1] -eq 'Value') { 0 } else { -1 }
    $entry   = $Zip.CreateEntry("xl/worksheets/sheet$Index.xml")
    $w       = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try {
        $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        $w.Write('<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        $w.Write('<sheetFormatPr defaultRowHeight="18"/>')
        $w.Write('<cols>')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $width = if ($Widths -and $i -lt $Widths.Count) { $Widths[$i] } else { 20 }
            $w.Write(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $width))
        }
        $w.Write('</cols><sheetData>')

        $w.Write('<row r="1" ht="24" customHeight="1">')
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
                $st  = if ($i -eq $labelCol) { 6 } elseif ($statusCols -contains $i -and $null -ne $val) { Get-XlsxStatusStyle "$val" } else { 0 }
                if ($null -eq $val) {
                    $w.Write("<c r=`"$ref`" s=`"0`"/>")   # empty but styled, so the row line is unbroken
                }
                elseif ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal] -or $val -is [single]) {
                    $w.Write("<c r=`"$ref`" s=`"$st`"><v>$(([double]$val).ToString('R', $inv))</v></c>")
                }
                elseif ($val -is [datetime]) {
                    $w.Write("<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t>$($val.ToString('yyyy-MM-dd HH:mm', $inv))</t></is></c>")
                }
                elseif ($val -is [bool]) {
                    $w.Write("<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t>$(if ($val) { 'Yes' } else { 'No' })</t></is></c>")
                }
                else {
                    $w.Write("<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText "$val")</t></is></c>")
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

        # Styles match the HTML reports: warm ink and greys, thin row lines, status colours like the badges.
        #   0 body  1 header  2 red  3 amber  4 blue  5 green  6 summary label
        $font = { param($Size, $Colour, $Name, [switch]$Bold) "<font>$(if ($Bold) { '<b/>' })<sz val=`"$Size`"/><color rgb=`"FF$Colour`"/><name val=`"$Name`"/></font>" }
        $fill = { param($Colour) "<fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF$Colour`"/><bgColor indexed=`"64`"/></patternFill></fill>" }
        $xf   = { param($Font, $Fill, $Border) "<xf numFmtId=`"0`" fontId=`"$Font`" fillId=`"$Fill`" borderId=`"$Border`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"left`" vertical=`"center`" indent=`"1`"/></xf>" }
        Add-ZipText $zip 'xl/styles.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<fonts count="7">' +
                (& $font 10 '1C1B19' 'Segoe UI') + (& $font 9 'E8E4DC' 'Consolas' -Bold) +
                (& $font 10 '9B2C2C' 'Segoe UI' -Bold) + (& $font 10 '875A0B' 'Segoe UI' -Bold) +
                (& $font 10 '2D5A86' 'Segoe UI' -Bold) + (& $font 10 '2F6A47' 'Segoe UI' -Bold) +
                (& $font 9 '78736A' 'Consolas') + '</fonts>' +
            '<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
                (& $fill '1C1B19') + (& $fill 'F9EBE9') + (& $fill 'FBF3E1') + (& $fill 'EAF1F8') + (& $fill 'E9F3EC') + '</fills>' +
            '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
                '<border><left/><right/><top/><bottom style="thin"><color rgb="FFE6E2DA"/></bottom><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="7">' +
                (& $xf 0 0 1) + (& $xf 1 2 0) + (& $xf 2 3 1) + (& $xf 3 4 1) + (& $xf 4 5 1) + (& $xf 5 6 1) + (& $xf 6 0 1) + '</cellXfs>' +
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

#endregion
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('AdminUrl'))     { $AdminUrl     = Get-Setting $config 'SharePoint' 'AdminUrl' '' }
if (-not $PSBoundParameters.ContainsKey('InactiveDays')) { $InactiveDays = [int](Get-Setting $config 'SharePoint' 'InactiveDays' 365) }
if (-not $AdminUrl) { throw 'No SharePoint admin URL. Use -AdminUrl https://<tenant>-admin.sharepoint.com.' }

$sharingNames = @{
    Disabled                         = 'Only people in your organisation'
    ExistingExternalUserSharingOnly  = 'Existing guests only'
    ExternalUserSharingOnly          = 'New and existing guests'
    ExternalUserAndGuestSharing      = 'Anyone (anonymous links)'
}

if ($MyInvocation.InvocationName -ne '.') {
    $module = 'Microsoft.Online.SharePoint.PowerShell'
    if (-not (Get-Module -ListAvailable -Name $module)) { throw "The SharePoint Online module is not installed. Run: Install-Module $module -Scope CurrentUser" }
    # The module is built for Windows PowerShell; in PowerShell 7 load it through the compatibility layer
    if ($PSVersionTable.PSEdition -eq 'Core' -and $script:IsWindowsHost) { Import-Module $module -UseWindowsPowerShell -WarningAction SilentlyContinue -ErrorAction Stop }
    else { Import-Module $module -DisableNameChecking -ErrorAction Stop }

    Write-Host "Connecting to $AdminUrl..."
    Connect-SPOService -Url $AdminUrl -ErrorAction Stop
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'SharePointSites' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()

        $tenantSharing = ''
        try { $tenantSharing = "$((Get-SPOTenant).SharingCapability)" } catch { $errors.Add([pscustomobject]@{ Site = '(tenant)'; Check = 'Tenant settings'; Error = $_.Exception.Message }) }

        Write-Host 'Getting sites...'
        $sites = @(Get-SPOSite -Limit All -IncludePersonalSite:([bool]$IncludeOneDrive) |
                   Where-Object { $IncludeOneDrive -or "$($_.Template)" -notlike 'SPSPERS*' })
        Write-Host "  $($sites.Count) site(s)"

        $siteRows = [System.Collections.Generic.List[object]]::new()
        $extRows  = [System.Collections.Generic.List[object]]::new()
        $i = 0
        foreach ($s in $sites) {
            $i++
            $external = @()
            $extFailed = $false
            if ($IncludeExternalUsers) {
                Write-Progress -Activity 'Checking sites for external users' -Status "$i of $($sites.Count): $($s.Url)" -PercentComplete ($i / [math]::Max(1, $sites.Count) * 100)
                try {
                    $external = @(Get-SPOUser -Site $s.Url -Limit All -ErrorAction Stop | Where-Object { $_.LoginName -like '*#ext#*' -or $_.LoginName -like '*urn:spo:guest*' })
                    foreach ($u in $external) {
                        # Guest login names look like i:0#.f|membership|jo_example.com#ext#@tenant.onmicrosoft.com
                        $email = if ($u.LoginName -match 'urn:spo:guest#(.+)$') { $Matches[1] }
                                 elseif ($u.LoginName -match '\|([^|]+)#ext#') { $Matches[1] -replace '_([^_]+)$', '@$1' }
                                 else { $u.LoginName }
                        $extRows.Add([pscustomobject]@{
                            Site = $s.Title; Url = $s.Url; DisplayName = $u.DisplayName
                            Email = $email
                            SiteAdmin = [bool]$u.IsSiteAdmin; LoginName = $u.LoginName
                        })
                    }
                }
                catch { $extFailed = $true; $errors.Add([pscustomobject]@{ Site = $s.Url; Check = 'External users'; Error = $_.Exception.Message }) }
            }

            $usedMb   = [double]$s.StorageUsageCurrent
            $quotaMb  = [double]$s.StorageQuota
            $pctUsed  = if ($quotaMb -gt 0) { [math]::Round($usedMb / $quotaMb * 100, 1) } else { $null }
            $lastMod  = $s.LastContentModifiedDate
            $idleDays = if ($lastMod -and $lastMod.Year -gt 2000) { [int]($now - $lastMod).TotalDays } else { $null }
            $sharing  = "$($s.SharingCapability)"

            $issues = @()
            $risk   = 'OK'
            if ($sharing -eq 'ExternalUserAndGuestSharing') { $issues += 'Anyone (anonymous) links allowed'; $risk = 'Medium' }
            if ($null -ne $pctUsed -and $pctUsed -ge 90)    { $issues += "Storage $pctUsed% of quota"; $risk = 'Medium' }
            if ($null -ne $idleDays -and $idleDays -gt $InactiveDays) { $issues += "No changes for $idleDays days - archive or delete?"; if ($risk -eq 'OK') { $risk = 'Low' } }
            if ("$($s.LockState)" -ne 'Unlock' -and "$($s.LockState)") { $issues += "Locked ($($s.LockState))" }
            if ($external.Count) { $issues += "$($external.Count) external user(s)" }

            $siteRows.Add([pscustomobject]@{
                Risk = $risk; Title = $s.Title; Url = $s.Url; Template = "$($s.Template)"; Owner = $s.Owner
                StorageGB = [math]::Round($usedMb / 1024, 2); QuotaGB = [math]::Round($quotaMb / 1024, 1); PercentUsed = $pctUsed
                LastModified = $lastMod; DaysSinceChange = $idleDays
                Sharing = if ($sharingNames.ContainsKey($sharing)) { $sharingNames[$sharing] } else { $sharing }
                ExternalUsers = if ($extFailed) { 'Error' } elseif ($IncludeExternalUsers) { $external.Count } else { '' }
                M365Group = if ("$($s.GroupId)" -and "$($s.GroupId)" -ne '00000000-0000-0000-0000-000000000000') { 'Yes' } else { 'No' }
                Issues = ($issues -join '; ')
            })
        }
        Write-Progress -Activity 'Checking sites for external users' -Completed

        $order = @{ Medium = 0; Low = 1; OK = 2 }
        $summary = [ordered]@{
            'Admin URL'               = $AdminUrl
            'Run at'                  = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Tenant sharing level'    = if ($sharingNames.ContainsKey($tenantSharing)) { $sharingNames[$tenantSharing] } else { $tenantSharing }
            'Sites'                   = $siteRows.Count
            'Total storage (GB)'      = [math]::Round((($siteRows | Measure-Object StorageGB -Sum).Sum), 2)
            'Sites allowing Anyone links' = @($siteRows | Where-Object { $_.Sharing -like 'Anyone*' }).Count
            "Inactive > $InactiveDays days" = @($siteRows | Where-Object { $_.DaysSinceChange -gt $InactiveDays }).Count
            'External users'          = if ($IncludeExternalUsers) { "$($extRows.Count) across $(@($extRows.Url | Select-Object -Unique).Count) site(s)" } else { 'Not checked (use -IncludeExternalUsers)' }
            'Note'                    = 'A site can never be shared more widely than the tenant sharing level allows.'
        }

        $sheets = @(
            @{ Name = 'Sites'; Rows = @($siteRows | Sort-Object @{ Expression = { $order[$_.Risk] } }, Title)
               Headers = @('Risk', 'Title', 'Url', 'Template', 'Owner', 'StorageGB', 'QuotaGB', 'PercentUsed', 'LastModified', 'DaysSinceChange', 'Sharing', 'ExternalUsers', 'M365Group', 'Issues')
               Widths  = @(9, 34, 60, 18, 34, 10, 9, 12, 18, 16, 28, 14, 11, 60) }
        )
        if ($IncludeExternalUsers) { $sheets += @{ Name = 'External Users'; Rows = @($extRows | Sort-Object Site, DisplayName); Headers = @('Site', 'Url', 'DisplayName', 'Email', 'SiteAdmin', 'LoginName'); Widths = @(34, 60, 28, 36, 10, 60) } }
        $sheets += @{ Name = 'Errors'; Rows = $errors; Headers = @('Site', 'Check', 'Error'); Widths = @(60, 16, 100) }
        $sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(28, 90) }
        if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

        Write-Host ""
        Write-Host ("Sites: {0}, {1} GB. Tenant sharing: {2}" -f $siteRows.Count, $summary['Total storage (GB)'], $summary['Tenant sharing level'])
        Write-Host ("Anyone links allowed on {0} site(s); {1} inactive site(s)" -f $summary['Sites allowing Anyone links'], $summary["Inactive > $InactiveDays days"])
        if ($IncludeExternalUsers) { Write-Host "External users: $($summary['External users'])" }
        if ($errors.Count) { Write-Host ("Errors: {0} - see the 'Errors' sheet" -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report: $reportPath"
    }
    finally {
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
}
