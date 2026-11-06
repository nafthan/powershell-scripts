<#
.SYNOPSIS
    Reports who has access to which Exchange Online mailboxes and calendars.
.DESCRIPTION
    Lists Full Access, Send As and Send on Behalf permissions, and optionally calendar permissions,
    for user, shared, room and equipment mailboxes. Flags permissions held by deleted users and
    calendars shared with the whole organisation. Read-only.
.EXAMPLE
    .\Get-MailboxPermissionReport.ps1
.EXAMPLE
    .\Get-MailboxPermissionReport.ps1 -Trustee jane.doe@contoso.com -IncludeCalendars
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: ExchangeOnlineManagement module, Exchange Administrator (or Global Reader)
#>
[CmdletBinding()]
param(
    [string[]]$Mailbox,
    [string]$Trustee,
    [switch]$IncludeCalendars,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config  = Get-ToolkitConfig
$Mailbox = Split-ListValue $Mailbox

# Calendar access levels that let others see more than free/busy
$broadCalendarRights = 'Reviewer|Author|Editor|Owner|PublishingEditor|PublishingAuthor|NonEditingAuthor|Contributor'

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitExchange -Config $config -Bound $PSBoundParameters
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'MailboxPermissions' -OutputFile $OutputFile
        $rows       = [System.Collections.Generic.List[object]]::new()
        $errors     = [System.Collections.Generic.List[object]]::new()

        Write-Host 'Getting mailboxes...'
        $types = 'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'
        if ($Mailbox.Count) { $mailboxes = @($Mailbox | ForEach-Object { Get-EXOMailbox -Identity $_ -Properties GrantSendOnBehalfTo -ErrorAction Stop }) }
        else { $mailboxes = @(Get-EXOMailbox -RecipientTypeDetails $types -ResultSize Unlimited -Properties GrantSendOnBehalfTo) }
        Write-Host "  $($mailboxes.Count) mailbox(es)"

        function Add-Row($Mbx, [string]$Permission, [string]$Who, [string]$Rights, [string]$Note) {
            if ($Trustee -and $Who -notlike "*$Trustee*") { return }
            $flag = ''
            if ($Who -match '^S-1-5-21-') { $flag = 'Held by a deleted user (orphaned SID) - remove it' }
            elseif ($Note) { $flag = $Note }
            $rows.Add([pscustomobject]@{
                Mailbox = "$($Mbx.PrimarySmtpAddress)"; DisplayName = $Mbx.DisplayName; MailboxType = "$($Mbx.RecipientTypeDetails)"
                Permission = $Permission; Trustee = $Who; AccessRights = $Rights; Flag = $flag
            })
        }

        $i = 0
        foreach ($m in $mailboxes) {
            $i++
            Write-Progress -Activity 'Reading mailbox permissions' -Status "$i of $($mailboxes.Count): $($m.PrimarySmtpAddress)" -PercentComplete ($i / [math]::Max(1, $mailboxes.Count) * 100)
            $id = "$($m.PrimarySmtpAddress)"

            try {
                foreach ($p in @(Get-EXOMailboxPermission -Identity $id -ErrorAction Stop)) {
                    if ($p.IsInherited -or "$($p.User)" -like 'NT AUTHORITY\*' -or "$($p.User)" -match '^S-1-5-(18|19|20|32)') { continue }
                    if (@($p.AccessRights) -notcontains 'FullAccess') { continue }
                    Add-Row $m 'Full Access' "$($p.User)" (@($p.AccessRights) -join ', ') $(if ($p.Deny) { 'Deny entry' })
                }
            }
            catch { $errors.Add([pscustomobject]@{ Mailbox = $id; Check = 'Full Access'; Error = $_.Exception.Message }) }

            try {
                foreach ($p in @(Get-EXORecipientPermission -Identity $id -ErrorAction Stop)) {
                    if ($p.IsInherited -or "$($p.Trustee)" -like 'NT AUTHORITY\*') { continue }
                    if (@($p.AccessRights) -notcontains 'SendAs') { continue }
                    Add-Row $m 'Send As' "$($p.Trustee)" 'SendAs' ''
                }
            }
            catch { $errors.Add([pscustomobject]@{ Mailbox = $id; Check = 'Send As'; Error = $_.Exception.Message }) }

            foreach ($d in @($m.GrantSendOnBehalfTo | Where-Object { $_ })) { Add-Row $m 'Send on Behalf' "$d" 'SendOnBehalf' '' }

            if ($IncludeCalendars) {
                try {
                    $calendar = "${id}:\Calendar"
                    try { $perms = @(Get-EXOMailboxFolderPermission -Identity $calendar -ErrorAction Stop) }
                    catch {
                        # Calendar folder has a different (localised) name
                        $folder = @(Get-EXOMailboxFolderStatistics -Identity $id -FolderScope Calendar -ErrorAction Stop | Where-Object { $_.FolderType -eq 'Calendar' })[0]
                        $calendar = "${id}:\$($folder.Name)"
                        $perms = @(Get-EXOMailboxFolderPermission -Identity $calendar -ErrorAction Stop)
                    }
                    foreach ($p in $perms) {
                        $who    = "$($p.User)"
                        $rights = (@($p.AccessRights) -join ', ')
                        if ($rights -eq 'None') { continue }
                        $note = ''
                        if ($who -eq 'Default' -and $rights -match $broadCalendarRights) { $note = 'Everyone in the organisation can see calendar details' }
                        elseif ($who -eq 'Default') { continue }   # Default AvailabilityOnly/LimitedDetails is normal
                        elseif ($who -eq 'Anonymous') { $note = 'Calendar published to anyone (anonymous access)' }
                        Add-Row $m 'Calendar' $who $rights $note
                    }
                }
                catch { $errors.Add([pscustomobject]@{ Mailbox = $id; Check = 'Calendar'; Error = $_.Exception.Message }) }
            }
        }
        Write-Progress -Activity 'Reading mailbox permissions' -Completed

        # --- Export ----------------------------------------------------------

        $sorted = @($rows | Sort-Object @{ Expression = { -not $_.Flag } }, Mailbox, Permission, Trustee)
        $byTrustee = @($rows | Where-Object { $_.Trustee -notin 'Default', 'Anonymous' } | Group-Object Trustee | ForEach-Object {
            $g = $_.Group
            [pscustomobject]@{
                Trustee      = $_.Name
                FullAccess   = @($g | Where-Object Permission -eq 'Full Access').Count
                SendAs       = @($g | Where-Object Permission -eq 'Send As').Count
                SendOnBehalf = @($g | Where-Object Permission -eq 'Send on Behalf').Count
                Calendars    = @($g | Where-Object Permission -eq 'Calendar').Count
                Mailboxes    = (@($g.Mailbox | Select-Object -Unique) -join ', ')
            }
        } | Sort-Object @{ Expression = { $_.FullAccess + $_.SendAs + $_.SendOnBehalf + $_.Calendars }; Descending = $true })

        $summary = [ordered]@{
            'Run at'                 = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Mailboxes checked'      = $mailboxes.Count
            'Filtered to trustee'    = if ($Trustee) { $Trustee } else { '(all)' }
            'Full Access'            = @($rows | Where-Object Permission -eq 'Full Access').Count
            'Send As'                = @($rows | Where-Object Permission -eq 'Send As').Count
            'Send on Behalf'         = @($rows | Where-Object Permission -eq 'Send on Behalf').Count
            'Calendar permissions'   = if ($IncludeCalendars) { @($rows | Where-Object Permission -eq 'Calendar').Count } else { 'Not checked (use -IncludeCalendars)' }
            'Flagged'                = @($rows | Where-Object Flag).Count
            'Note'                   = 'Inherited and system permissions are not listed. Full Access granted through groups shows the group name.'
        }

        $saved = Save-Workbook -FilePath $reportPath -Sheets @(
            @{ Name = 'Permissions'; Rows = $sorted; Headers = @('Mailbox', 'DisplayName', 'MailboxType', 'Permission', 'Trustee', 'AccessRights', 'Flag'); Widths = @(36, 26, 16, 15, 36, 24, 50) }
            @{ Name = 'By Trustee'; Rows = $byTrustee; Headers = @('Trustee', 'FullAccess', 'SendAs', 'SendOnBehalf', 'Calendars', 'Mailboxes'); Widths = @(36, 11, 9, 13, 10, 100) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Mailbox', 'Check', 'Error'); Widths = @(36, 14, 100) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(24, 90) }
        )
        if (-not $saved) { exit 1 }

        Write-Host ""
        Write-Host ("Permissions: {0} Full Access, {1} Send As, {2} Send on Behalf{3}" -f $summary['Full Access'], $summary['Send As'], $summary['Send on Behalf'], $(if ($IncludeCalendars) { ", $($summary['Calendar permissions']) calendar" }))
        if ($summary.Flagged) { Write-Host ("Flagged    : {0} (orphaned SIDs or calendars shared widely)" -f $summary.Flagged) -ForegroundColor Yellow }
        if ($errors.Count) { Write-Host ("Errors     : {0} - see the 'Errors' sheet" -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report     : $reportPath"
    }
    finally {
        if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    }
}
