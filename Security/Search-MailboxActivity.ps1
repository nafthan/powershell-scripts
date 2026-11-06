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

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
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
