<#
.SYNOPSIS
    Exports every Microsoft 365 group, Team, distribution list and security group with its members and owners.
.DESCRIPTION
    One row per member, plus a summary row per group that flags groups with no owner, guests and
    empty groups. Optionally includes private and shared Teams channel members. Read-only.
.EXAMPLE
    .\Get-GroupMembershipReport.ps1
.EXAMPLE
    .\Get-GroupMembershipReport.ps1 -GroupType Team -IncludeTeamsChannels
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: Microsoft.Graph.Authentication module, Global Reader
#>
[CmdletBinding()]
param(
    [string[]]$GroupType,
    [switch]$IncludeTeamsChannels,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config    = Get-ToolkitConfig
$GroupType = Split-ListValue $GroupType
$validTypes = 'Team', 'Microsoft 365 group', 'Distribution list', 'Mail-enabled security', 'Security'
foreach ($t in $GroupType) { if ($validTypes -notcontains $t) { throw "Unknown -GroupType '$t'. Valid: $($validTypes -join ', ')." } }

$scopes = @('Group.Read.All', 'User.Read.All')
if ($IncludeTeamsChannels) { $scopes += 'Channel.ReadBasic.All', 'ChannelMember.Read.All' }

function Get-GroupKind($G) {
    if (@($G.resourceProvisioningOptions) -contains 'Team') { return 'Team' }
    if (@($G.groupTypes) -contains 'Unified') { return 'Microsoft 365 group' }
    if ($G.mailEnabled -and -not $G.securityEnabled) { return 'Distribution list' }
    if ($G.mailEnabled -and $G.securityEnabled) { return 'Mail-enabled security' }
    return 'Security'
}

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'Global Reader'
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'GroupMembership' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        $tenantName = Get-GraphTenantName
        Write-Host "Tenant : $tenantName"

        Write-Host 'Getting groups...'
        $select = 'id,displayName,mail,description,createdDateTime,groupTypes,mailEnabled,securityEnabled,resourceProvisioningOptions,membershipRule,visibility,onPremisesSyncEnabled'
        $groups = @(Get-GraphAll "v1.0/groups?`$select=$select&`$top=999" | ForEach-Object {
            $_ | Add-Member -NotePropertyName Kind -NotePropertyValue (Get-GroupKind $_) -PassThru
        })
        if ($GroupType.Count) { $groups = @($groups | Where-Object { $GroupType -contains $_.Kind }) }
        Write-Host "  $($groups.Count) group(s)"

        $memberRows  = [System.Collections.Generic.List[object]]::new()
        $groupRows   = [System.Collections.Generic.List[object]]::new()
        $channelRows = [System.Collections.Generic.List[object]]::new()
        $i = 0
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity 'Reading members' -Status "$i of $($groups.Count): $($g.displayName)" -PercentComplete ($i / [math]::Max(1, $groups.Count) * 100)
            $members = @(); $owners = @(); $membersRead = $true
            try {
                # One request per page of members - no per-member lookups
                $members = @(Get-GraphAll "v1.0/groups/$($g.id)/members?`$select=id,displayName,userPrincipalName,mail,userType,accountEnabled&`$top=999")
            }
            catch { $membersRead = $false; $errors.Add([pscustomobject]@{ Group = $g.displayName; Check = 'Members'; Error = $_.Exception.Message }) }
            if ($g.Kind -in 'Team', 'Microsoft 365 group', 'Security', 'Mail-enabled security') {
                try { $owners = @(Get-GraphAll "v1.0/groups/$($g.id)/owners?`$select=displayName,userPrincipalName") }
                catch { $errors.Add([pscustomobject]@{ Group = $g.displayName; Check = 'Owners'; Error = $_.Exception.Message }) }
            }

            foreach ($m in $members) {
                $type = "$($m.'@odata.type')" -replace '^#microsoft\.graph\.', ''
                $memberRows.Add([pscustomobject]@{
                    GroupType   = $g.Kind
                    Group       = $g.displayName
                    GroupEmail  = $g.mail
                    Member      = $m.displayName
                    MemberEmail = if ($m.userPrincipalName) { $m.userPrincipalName } else { $m.mail }
                    MemberType  = switch ($type) { 'user' { if ($m.userType -eq 'Guest') { 'Guest' } else { 'User' } } 'group' { 'Nested group' } 'orgContact' { 'Contact' } 'device' { 'Device' } 'servicePrincipal' { 'App' } default { $type } }
                    Enabled     = if ($type -eq 'user') { [bool]$m.accountEnabled } else { '' }
                })
            }

            $guests  = @($members | Where-Object { $_.userType -eq 'Guest' }).Count
            $issues  = @()
            if ($g.Kind -in 'Team', 'Microsoft 365 group' -and -not $owners.Count) { $issues += 'No owner - nobody manages membership or renews it' }
            elseif ($g.Kind -in 'Team', 'Microsoft 365 group' -and $owners.Count -eq 1) { $issues += 'Only one owner' }
            if (-not $membersRead) { $issues += 'Members could not be read' }
            elseif (-not $members.Count) { $issues += 'No members' }
            if ($guests) { $issues += "$guests guest(s)" }
            $disabled = @($members | Where-Object { "$($_.'@odata.type')" -like '*user' -and $_.accountEnabled -eq $false }).Count
            if ($disabled) { $issues += "$disabled disabled account(s) still members" }

            $groupRows.Add([pscustomobject]@{
                GroupType  = $g.Kind
                Group      = $g.displayName
                Email      = $g.mail
                Members    = if ($membersRead) { $members.Count } else { 'Error' }
                Guests     = $guests
                Owners     = (@($owners | ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }) -join ', ')
                Dynamic    = if (@($g.groupTypes) -contains 'DynamicMembership') { 'Yes' } else { 'No' }
                Visibility = $g.visibility
                Source     = if ($g.onPremisesSyncEnabled) { 'Synced (AD)' } else { 'Cloud' }
                Created    = ConvertTo-LocalTime $g.createdDateTime
                Issues     = ($issues -join '; ')
                Description = $g.description
            })

            if ($IncludeTeamsChannels -and $g.Kind -eq 'Team') {
                try {
                    foreach ($c in @(Get-GraphAll "v1.0/teams/$($g.id)/channels?`$select=id,displayName,membershipType" | Where-Object { $_.membershipType -ne 'standard' })) {
                        foreach ($cm in @(Get-GraphAll "v1.0/teams/$($g.id)/channels/$($c.id)/members")) {
                            $channelRows.Add([pscustomobject]@{
                                Team = $g.displayName; Channel = $c.displayName; ChannelType = (Get-Culture).TextInfo.ToTitleCase("$($c.membershipType)")
                                Member = $cm.displayName; MemberEmail = $cm.email; Role = if (@($cm.roles).Count) { (@($cm.roles) -join ', ') } else { 'member' }
                            })
                        }
                    }
                }
                catch { $errors.Add([pscustomobject]@{ Group = $g.displayName; Check = 'Channels'; Error = $_.Exception.Message }) }
            }
        }
        Write-Progress -Activity 'Reading members' -Completed

        # --- Export ----------------------------------------------------------

        $summaryItems = [ordered]@{
            'Tenant'          = $tenantName
            'Run at'          = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Groups'          = $groupRows.Count
        }
        foreach ($k in $validTypes) { $summaryItems["  $k"] = @($groupRows | Where-Object GroupType -eq $k).Count }
        $summaryItems['Membership rows']          = $memberRows.Count
        $summaryItems['Teams/M365 groups with no owner'] = @($groupRows | Where-Object { $_.Issues -like 'No owner*' }).Count
        $summaryItems['Groups with guests']       = @($groupRows | Where-Object { $_.Guests -gt 0 }).Count
        $summaryItems['Empty groups']             = @($groupRows | Where-Object { $_.Members -is [int] -and $_.Members -eq 0 }).Count
        $summaryItems['Note']                     = 'Members are direct members (nested groups are listed as one row). Exchange dynamic distribution groups are not in Microsoft Graph and are not included.'

        $sheets = @(
            @{ Name = 'Groups'; Rows = @($groupRows | Sort-Object GroupType, Group)
               Headers = @('GroupType', 'Group', 'Email', 'Members', 'Guests', 'Owners', 'Dynamic', 'Visibility', 'Source', 'Created', 'Issues', 'Description')
               Widths  = @(20, 34, 34, 10, 8, 45, 9, 11, 12, 18, 50, 50) }
            @{ Name = 'Members'; Rows = @($memberRows | Sort-Object GroupType, Group, Member)
               Headers = @('GroupType', 'Group', 'GroupEmail', 'Member', 'MemberEmail', 'MemberType', 'Enabled'); Widths = @(20, 34, 34, 28, 36, 13, 9) }
        )
        if ($IncludeTeamsChannels) {
            $sheets += @{ Name = 'Private-Shared Channels'; Rows = @($channelRows | Sort-Object Team, Channel, Member)
                          Headers = @('Team', 'Channel', 'ChannelType', 'Member', 'MemberEmail', 'Role'); Widths = @(30, 30, 12, 28, 36, 12) }
        }
        $sheets += @{ Name = 'Errors'; Rows = $errors; Headers = @('Group', 'Check', 'Error'); Widths = @(34, 12, 100) }
        $sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summaryItems); Headers = @('Item', 'Value'); Widths = @(32, 100) }
        if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

        Write-Host ""
        Write-Host ("Groups: {0}  ({1})" -f $groupRows.Count, (($validTypes | ForEach-Object { "$(@($groupRows | Where-Object GroupType -eq $_).Count) $_" }) -join ', '))
        Write-Host ("No owner: {0}   With guests: {1}   Empty: {2}" -f $summaryItems['Teams/M365 groups with no owner'], $summaryItems['Groups with guests'], $summaryItems['Empty groups'])
        if ($errors.Count) { Write-Host ("Errors: {0} - see the 'Errors' sheet" -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report: $reportPath"
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
