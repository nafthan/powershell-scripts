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
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
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
