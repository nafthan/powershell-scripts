<#
.SYNOPSIS
    Searches the Microsoft 365 unified audit log for activity on a mailbox and saves it as evidence.
.DESCRIPTION
    Finds what was done by (and optionally to) a mailbox: mail read, sent and deleted, inbox rules and
    forwarding changed, and from which IP addresses. Replaces the retired Search-MailboxAuditLog. Read-only.
.EXAMPLE
    .\Search-MailboxActivity.ps1 -Mailbox jane.doe@contoso.com -Days 7 -CaseId IR-2026-014
.EXAMPLE
    .\Search-MailboxActivity.ps1 -Mailbox shared@contoso.com -StartDate '2026-10-20' -EndDate '2026-10-23' -IncludeAccessByOthers
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Requires: ExchangeOnlineManagement module; Audit Logs or View-Only Audit Logs role (e.g. Compliance Administrator)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Mailbox,

    [ValidateRange(1, 180)] [int]$Days = 7,
    [datetime]$StartDate,
    [datetime]$EndDate,

    [switch]$IncludeAccessByOthers,
    [switch]$AllWorkloads,
    [ValidateRange(100, 50000)] [int]$MaxRecords = 50000,

    [string]$CaseId,
    [string]$OutputPath,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode
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

#endregion
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('EndDate'))   { $EndDate = Get-Date }
if (-not $PSBoundParameters.ContainsKey('StartDate')) { $StartDate = $EndDate.AddDays(-$Days) }
if ($StartDate -ge $EndDate) { throw '-StartDate must be before -EndDate.' }

# Exchange record types that cover mailbox activity
$mailboxRecordTypes = 'ExchangeItem', 'ExchangeItemGroup', 'ExchangeItemAggregated', 'ExchangeAdmin'
$ruleOperations     = 'New-InboxRule', 'Set-InboxRule', 'Enable-InboxRule', 'UpdateInboxRules', 'Set-Mailbox', 'Set-CASMailbox', 'Add-MailboxPermission', 'Add-RecipientPermission'
$deleteOperations   = 'SoftDelete', 'HardDelete', 'MoveToDeletedItems'
$sendOperations     = 'Send', 'SendAs', 'SendOnBehalf'

function Invoke-AuditSearch([hashtable]$Filter) {
    <# Pages through Search-UnifiedAuditLog with a session, up to $MaxRecords results. #>
    $session = "MbxActivity_$([guid]::NewGuid())"
    $found = [System.Collections.Generic.List[object]]::new()
    do {
        $batch = @(Search-UnifiedAuditLog @Filter -StartDate $StartDate -EndDate $EndDate -SessionId $session -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
        foreach ($b in $batch) { $found.Add($b) }
        $total = if ($batch.Count) { [int]$batch[0].ResultCount } else { 0 }
    } while ($batch.Count -gt 0 -and $found.Count -lt $total -and $found.Count -lt $MaxRecords)
    return $found
}

function ConvertFrom-AuditRecord($Record) {
    $d = $null
    try { $d = $Record.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $d = [pscustomobject]@{} }
    $folders  = @($d.Folders | Where-Object { $_ })
    $items    = @($d.AffectedItems | Where-Object { $_ })
    $subjects = @()
    if ($d.Item.Subject) { $subjects += $d.Item.Subject }
    $subjects += @($items | ForEach-Object { $_.Subject } | Where-Object { $_ })
    $folder   = if ($d.Item.ParentFolder.Path) { $d.Item.ParentFolder.Path } elseif ($folders.Count) { (@($folders | ForEach-Object { $_.Path }) | Select-Object -Unique) -join '; ' } elseif ($items.Count) { (@($items | ForEach-Object { $_.ParentFolder.Path }) | Select-Object -Unique) -join '; ' } else { $d.DestFolder.Path }
    $count    = if ($folders.Count) { (@($folders | ForEach-Object { @($_.FolderItems).Count }) | Measure-Object -Sum).Sum } elseif ($items.Count) { $items.Count } elseif ($d.OperationCount) { $d.OperationCount } else { 1 }
    $params   = if ($d.Parameters) { (@($d.Parameters | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; ') } else { '' }
    $ip       = if ($d.ClientIPAddress) { $d.ClientIPAddress } else { $d.ClientIP }
    if ($ip -match '^\[?([0-9a-fA-F:.]+?)\]?(:\d+)?$' -and $ip -notmatch '^[0-9a-fA-F:]+$') { $ip = $Matches[1] }   # strip port from IPv4:port / [IPv6]:port

    [pscustomobject]@{
        Time          = ConvertTo-UtcText $(if ($d.CreationTime) { $d.CreationTime } else { $Record.CreationDate })
        Operation     = if ($d.Operation) { $d.Operation } else { $Record.Operations }
        User          = if ($d.UserId) { $d.UserId } else { $Record.UserIds }
        MailboxOwner  = $d.MailboxOwnerUPN
        LogonType     = switch ("$($d.LogonType)") { '0' { 'Owner' } '1' { 'Admin' } '2' { 'Delegate' } default { "$($d.LogonType)" } }
        ClientIP      = $ip
        Client        = (@($d.ClientInfoString, $d.ClientAppId, $d.AppId) | Where-Object { $_ } | Select-Object -First 1)
        Folder        = $folder
        Items         = $count
        Subjects      = (@($subjects | Select-Object -First 5) -join ' | ')
        Parameters    = $params
        Result        = $d.ResultStatus
        Workload      = $d.Workload
        RecordType    = "$($Record.RecordType)"
        RecordId      = $Record.Identity
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitExchange -Config $config -Bound $PSBoundParameters -RoleHint 'an audit log role (e.g. Compliance Administrator)'
    try {
        Write-Host ("Searching the audit log for {0}, {1} to {2} (UTC)..." -f $Mailbox, $StartDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm'), $EndDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm'))
        $log = [System.Collections.Generic.List[object]]::new()
        $raw = [System.Collections.Generic.List[object]]::new()

        # 1. Everything done by the account itself
        if ($AllWorkloads) { $searches = @(@{ Label = 'All workloads'; Filter = @{ UserIds = $Mailbox } }) }
        else { $searches = @($mailboxRecordTypes | ForEach-Object { @{ Label = $_; Filter = @{ UserIds = $Mailbox; RecordType = $_ } } }) }
        # 2. Optionally, mailbox activity by anyone, filtered to this mailbox afterwards (slower)
        if ($IncludeAccessByOthers) { $searches += @('ExchangeItem', 'ExchangeItemGroup', 'ExchangeItemAggregated' | ForEach-Object { @{ Label = "$_ (all users)"; Filter = @{ RecordType = $_ }; OthersOnly = $true } }) }

        foreach ($s in $searches) {
            Write-Host "  $($s.Label)..."
            try {
                $found = @(Invoke-AuditSearch $s.Filter)
                if ($s.OthersOnly) { $found = @($found | Where-Object { "$($_.AuditData)" -match ('"MailboxOwnerUPN"\s*:\s*"' + [regex]::Escape($Mailbox) + '"') -and "$($_.UserIds)" -ne $Mailbox }) }
                foreach ($f in $found) { $raw.Add($f) }
                $log.Add([pscustomobject]@{ Search = $s.Label; Status = 'OK'; Records = $found.Count; Detail = $(if ($found.Count -ge $MaxRecords) { "Stopped at $MaxRecords records - narrow the dates" } else { '' }) })
            }
            catch { $log.Add([pscustomobject]@{ Search = $s.Label; Status = 'Failed'; Records = 0; Detail = $_.Exception.Message }) }
        }

        $seen = @{}
        $rows = @($raw | Where-Object { if ($seen.ContainsKey($_.Identity)) { $false } else { $seen[$_.Identity] = $true; $true } } |
                  ForEach-Object { ConvertFrom-AuditRecord $_ } | Sort-Object Time)
        Write-Host "  $($rows.Count) event(s)"

        # --- Analysis --------------------------------------------------------
        $ips = @($rows | Where-Object ClientIP | Group-Object ClientIP | ForEach-Object {
            $g = @($_.Group)
            [pscustomobject]@{
                ClientIP = $_.Name; Events = $g.Count; FirstSeen = $g[0].Time; LastSeen = $g[-1].Time
                Users = (@($g.User | Select-Object -Unique) -join ', '); Operations = (@($g | Group-Object Operation | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ')
                Clients = (@($g.Client | Where-Object { $_ } | Select-Object -Unique | Select-Object -First 3) -join ' | ')
            }
        } | Sort-Object Events -Descending)
        $ops = @($rows | Group-Object Operation | ForEach-Object { [pscustomobject]@{ Operation = $_.Name; Events = $_.Count; Items = ($_.Group | Measure-Object Items -Sum).Sum } } | Sort-Object Events -Descending)

        $indicators = [System.Collections.Generic.List[object]]::new()
        foreach ($r in @($rows | Where-Object { $ruleOperations -contains $_.Operation })) {
            $sev = if ($r.Parameters -match '(?i)ForwardTo|RedirectTo|ForwardingSmtpAddress|DeleteMessage|MoveToFolder') { 'High' } else { 'Medium' }
            $indicators.Add([pscustomobject]@{ Severity = $sev; Indicator = "$($r.Operation) by $($r.User)"; Time = $r.Time; Detail = "$($r.Parameters) (IP $($r.ClientIP))" })
        }
        $deletes = @($rows | Where-Object { $deleteOperations -contains $_.Operation })
        $deleted = ($deletes | Measure-Object Items -Sum).Sum
        if ($deleted -ge 50) { $indicators.Add([pscustomobject]@{ Severity = 'Medium'; Indicator = "$deleted item(s) deleted"; Time = "$($deletes[0].Time) - $($deletes[-1].Time)"; Detail = (@($deletes.Operation | Select-Object -Unique) -join ', ') }) }
        $sends = @($rows | Where-Object { $sendOperations -contains $_.Operation })
        if ($sends.Count) { $indicators.Add([pscustomobject]@{ Severity = 'Low'; Indicator = "$($sends.Count) message(s) sent"; Time = "$($sends[0].Time) - $($sends[-1].Time)"; Detail = "From IP(s): $(@($sends.ClientIP | Where-Object { $_ } | Select-Object -Unique) -join ', ')" }) }
        $others = @($rows | Where-Object { $_.User -and $_.User -ne $Mailbox -and $_.MailboxOwner -eq $Mailbox })
        foreach ($g in ($others | Group-Object User)) { $indicators.Add([pscustomobject]@{ Severity = 'Low'; Indicator = "Mailbox accessed by $($g.Name)"; Time = "$($g.Group[0].Time) - $($g.Group[-1].Time)"; Detail = "$($g.Count) event(s): $(@($g.Group.Operation | Select-Object -Unique) -join ', ')" }) }
        if ($ips.Count -gt 1) { $indicators.Add([pscustomobject]@{ Severity = 'Low'; Indicator = "Activity from $($ips.Count) IP addresses"; Time = ''; Detail = (@($ips | Select-Object -First 10 | ForEach-Object { "$($_.ClientIP) ($($_.Events))" }) -join ', ') }) }
        $sevOrder = @{ High = 0; Medium = 1; Low = 2 }
        $indicators = @($indicators | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } }, Time)

        # --- Evidence package ------------------------------------------------
        $name   = (@($CaseId, 'MailboxActivity', ($Mailbox -split '@')[0], (Get-Date -Format 'yyyyMMdd_HHmmss')) | Where-Object { $_ }) -join '_'
        $folder = New-EvidenceFolder -Config $config -OutputPath $OutputPath -Name $name
        [void](Export-EvidenceCsv -Folder $folder -Name '01_Activity' -Rows $rows)
        [void](Export-EvidenceCsv -Folder $folder -Name '02_IPAddresses' -Rows $ips)
        [void](Export-EvidenceCsv -Folder $folder -Name '03_Operations' -Rows $ops)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_Indicators' -Rows $indicators)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_SearchLog' -Rows $log)

        $html = New-HtmlReport -Title "Mailbox activity - $Mailbox" -IntroLines @(
            "Period: $(ConvertTo-UtcText $StartDate) to $(ConvertTo-UtcText $EndDate). $($rows.Count) event(s). All times are UTC.",
            "Searched: activity by the account$(if ($IncludeAccessByOthers) { ', and by anyone else on this mailbox' }). Run by $env:USERDOMAIN\$env:USERNAME$(if ($CaseId) { " - case $CaseId" }).",
            'Some events (e.g. MailItemsAccessed) are only recorded with the right audit licensing and settings; absence of an event is not proof it did not happen.'
        ) -Sections @(
            @{ Heading = "Indicators ($($indicators.Count))"; Html = ConvertTo-HtmlTable -Rows $indicators -Columns @('Severity', 'Indicator', 'Time', 'Detail') -StatusColumn 'Severity' -EmptyText 'No automatic indicators.' }
            @{ Heading = 'IP addresses'; Html = ConvertTo-HtmlTable -Rows @($ips | Select-Object -First 25) -Columns @('ClientIP', 'Events', 'FirstSeen', 'LastSeen', 'Operations') -EmptyText 'No IP addresses recorded.' }
            @{ Heading = 'Operations'; Html = ConvertTo-HtmlTable -Rows $ops -Columns @('Operation', 'Events', 'Items') -EmptyText 'No events found.' }
            @{ Heading = 'Searches'; Html = ConvertTo-HtmlTable -Rows $log -Columns @('Search', 'Status', 'Records', 'Detail') -StatusColumn 'Status' }
        )
        [System.IO.File]::WriteAllText((Join-Path $folder '00_Summary.html'), $html, $script:Utf8NoBom)
        $package = Complete-EvidencePackage -Folder $folder

        Write-Host ""
        foreach ($i in $indicators | Select-Object -First 10) { Write-Host ("  {0,-6} {1} - {2}" -f $i.Severity, $i.Indicator, $i.Detail) -ForegroundColor $(switch ($i.Severity) { 'High' { 'Red' } 'Medium' { 'Yellow' } default { 'Gray' } }) }
        $failed = @($log | Where-Object Status -ne 'OK')
        if ($failed.Count) { Write-Host ("Some searches failed: {0}" -f (($failed | ForEach-Object { "$($_.Search): $($_.Detail)" }) -join '; ')) -ForegroundColor Yellow }
        Write-Host "Evidence : $($package.Zip)"
        Write-Host "SHA-256  : $($package.Sha256)  (record this in your case notes)"
    }
    finally {
        if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    }
}
