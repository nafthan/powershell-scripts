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
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
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
