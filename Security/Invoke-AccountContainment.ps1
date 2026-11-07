<#
.SYNOPSIS
    Contains a compromised user account in Entra ID / Microsoft 365 and Active Directory, after saving evidence.
.DESCRIPTION
    1. Collects evidence first: account details, groups, MFA methods, sign-ins, audit events, devices,
       app consents, mailbox forwarding and inbox rules.
    2. Then (asking before each step): disables the account, resets the password to a random value,
       revokes sessions, removes mail forwarding and disables suspicious inbox rules.
    Works for cloud-only and AD-synced users. Use -WhatIf to preview, or -EvidenceOnly to change nothing.
.EXAMPLE
    .\Invoke-AccountContainment.ps1 -User jane.doe@contoso.com -CaseId IR-2026-014 -WhatIf
.EXAMPLE
    .\Invoke-AccountContainment.ps1 -User jane.doe@contoso.com -CaseId IR-2026-014
.EXAMPLE
    .\Invoke-AccountContainment.ps1 -User jane.doe@contoso.com -EvidenceOnly
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: Microsoft.Graph.Authentication module (+ ExchangeOnlineManagement for mailbox steps, ActiveDirectory for synced users).
              Roles: User Administrator + Authentication Administrator (Privileged Authentication Administrator if the user is an admin),
              Exchange Administrator, and rights to reset passwords in AD.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$User,
    [string]$CaseId,

    [switch]$EvidenceOnly,
    [string[]]$Skip,
    [datetime]$RemoveMfaMethodsAddedAfter,

    [ValidateRange(1, 30)] [int]$SignInDays,
    [switch]$SkipExchange,
    [switch]$SkipAD,
    [switch]$SkipCloud,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [string]$OutputPath
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('SignInDays')) { $SignInDays = [int](Get-Setting $config 'SOC' 'SignInDays' 7) }
$Skip = @(Split-ListValue $Skip)
$validSteps = 'Disable', 'ResetPassword', 'RevokeSessions', 'Mailbox'
foreach ($s in $Skip) { if ($validSteps -notcontains $s) { throw "Unknown step '$s' in -Skip. Valid steps: $($validSteps -join ', ')." } }

$readScopes  = @('User.Read.All', 'UserAuthenticationMethod.Read.All', 'AuditLog.Read.All', 'Directory.Read.All')
$writeScopes = @('User.ReadWrite.All', 'User.RevokeSessions.All', 'UserAuthenticationMethod.ReadWrite.All', 'AuditLog.Read.All', 'Directory.Read.All')
$passwordMethodId = '28c10230-6103-485e-b985-444c60001490'   # fixed ID of the password authentication method

$methodSegments = @{
    microsoftAuthenticatorAuthenticationMethod  = 'microsoftAuthenticatorMethods'
    fido2AuthenticationMethod                   = 'fido2Methods'
    softwareOathAuthenticationMethod            = 'softwareOathMethods'
    windowsHelloForBusinessAuthenticationMethod = 'windowsHelloForBusinessMethods'
    phoneAuthenticationMethod                   = 'phoneMethods'
    emailAuthenticationMethod                   = 'emailMethods'
    temporaryAccessPassAuthenticationMethod     = 'temporaryAccessPassMethods'
    platformCredentialAuthenticationMethod      = 'platformCredentialMethods'
}

function Invoke-GraphChange([string]$Method, [string]$Uri, $Body) {
    $p = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $p.Body = ($Body | ConvertTo-Json -Depth 5 -Compress); $p.ContentType = 'application/json' }
    Invoke-MgGraphRequest @p
}

function New-RandomPassword([int]$Length = 24) {
    $sets = @('ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnopqrstuvwxyz', '23456789', '!#%^*-_=+?')
    $all  = -join $sets
    $rng  = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = New-Object byte[] $Length
    $rng.GetBytes($bytes)
    $chars = for ($i = 0; $i -lt $Length; $i++) { $set = if ($i -lt $sets.Count) { $sets[$i] } else { $all }; $set[$bytes[$i] % $set.Length] }
    return -join ($chars | Sort-Object { Get-Random })
}

if ($MyInvocation.InvocationName -ne '.') {
    $start   = Get-Date
    $since   = $start.AddDays(-$SignInDays)
    $sinceIso = $since.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $log     = [System.Collections.Generic.List[object]]::new()
    $actions = [System.Collections.Generic.List[object]]::new()
    $notes   = [System.Collections.Generic.List[string]]::new()
    function Add-Log([string]$Section, [string]$Status, [string]$Detail) { $log.Add([pscustomobject]@{ Section = $Section; Status = $Status; Detail = $Detail }) }
    function Add-Action([string]$Step, [string]$Result, [string]$Detail) {
        $actions.Add([pscustomobject]@{ Time = ConvertTo-UtcText (Get-Date); Step = $Step; Result = $Result; Detail = $Detail })
        $colour = switch -Regex ($Result) { '^(Done|Requested)' { 'Green' } '^(Failed|Not done)' { 'Red' } default { 'Yellow' } }
        Write-Host ("  {0,-16} {1}{2}" -f $Step, $Result, $(if ($Detail) { " - $Detail" })) -ForegroundColor $colour
    }
    $mode = if ($EvidenceOnly) { 'Evidence only (no changes)' } elseif ($WhatIfPreference) { 'WhatIf (no changes)' } else { 'Containment' }
    Write-Host "Mode: $mode"

    $graphConnected = $false; $exoConnected = $false
    $cloudUser = $null; $adUser = $null; $mailbox = $null
    try {
        # --- Connect and find the user -------------------------------------------

        $adAvailable = $false
        $needAd = -not $SkipAD -and ($SkipCloud -or $User -notmatch '@')
        if ($needAd -or -not $SkipAD) {
            if (Get-Module -ListAvailable -Name ActiveDirectory) {
                try { Import-Module ActiveDirectory -ErrorAction Stop; $adAvailable = $true } catch { Add-Log 'Active Directory' 'Unavailable' $_.Exception.Message }
            }
            elseif ($needAd) { Add-Log 'Active Directory' 'Unavailable' 'ActiveDirectory module not installed' }
        }

        $upn = $User
        if ($User -notmatch '@' -and $adAvailable) {
            try { $adUser = Get-ADUser -Identity $User -Properties * -ErrorAction Stop; $upn = $adUser.UserPrincipalName } catch { throw "User '$User' not found in Active Directory: $($_.Exception.Message)" }
        }

        if (-not $SkipCloud) {
            $scopes = if ($EvidenceOnly) { $readScopes } else { $writeScopes }
            $graphConnected = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'User Administrator + Authentication Administrator'
            $select = 'id,displayName,userPrincipalName,mail,accountEnabled,userType,onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesDomainName,createdDateTime,lastPasswordChangeDateTime,signInSessionsValidFromDateTime,assignedLicenses,jobTitle,department'
            try { $cloudUser = Invoke-Graph "v1.0/users/$([uri]::EscapeDataString($upn))?`$select=$select" }
            catch {
                if ($_.Exception.Message -match '404|NotFound|ResourceNotFound|does not exist') { Add-Log 'Entra ID' 'Not found' "No cloud account for $upn" }
                else { throw }
            }
        }
        if (-not $cloudUser -and -not $adUser) { throw "User '$User' was not found." }

        $isSynced = $cloudUser -and $cloudUser.onPremisesSyncEnabled
        if (($isSynced -or -not $cloudUser) -and -not $adUser -and -not $SkipAD) {
            if ($adAvailable) {
                $sam = if ($cloudUser.onPremisesSamAccountName) { $cloudUser.onPremisesSamAccountName } else { ($upn -split '@')[0] }
                try { $adUser = Get-ADUser -Identity $sam -Properties * -ErrorAction Stop } catch { Add-Log 'Active Directory' 'Not found' "${sam}: $($_.Exception.Message)" }
            }
            else {
                Add-Log 'Active Directory' 'Unavailable' 'The user is synced from AD but the ActiveDirectory module is not available - AD steps will be skipped'
            }
        }
        $display = if ($cloudUser) { "$($cloudUser.displayName) <$($cloudUser.userPrincipalName)>" } else { "$($adUser.Name) <$($adUser.UserPrincipalName)>" }
        if ($cloudUser) { $upn = $cloudUser.userPrincipalName }
        Write-Host "User: $display  ($(if ($isSynced) { 'synced from AD' } elseif ($cloudUser) { 'cloud-only' } else { 'AD only' }))"

        if (-not $SkipExchange -and $cloudUser) {
            try {
                $exoConnected = Connect-ToolkitExchange -Config $config -Bound $PSBoundParameters -RoleHint 'Exchange Administrator'
                $mailbox = Get-EXOMailbox -Identity $upn -Properties ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward -ErrorAction Stop
            }
            catch { Add-Log 'Exchange Online' 'Unavailable' $_.Exception.Message }
        }

        $folderName = (@($CaseId, 'Containment', (($upn -split '@')[0]), (Get-Date -Format 'yyyyMMdd_HHmmss')) | Where-Object { $_ }) -join '_'
        $folder = New-EvidenceFolder -Config $config -OutputPath $OutputPath -Name $folderName

        # --- 1. Evidence (before any change) --------------------------------------

        Write-Host 'Collecting evidence...'
        $evidence = [ordered]@{}
        function Save-Evidence([string]$Name, [scriptblock]$Block) {
            try {
                $rows = @(& $Block)
                $evidence[$Name] = $rows
                $count = Export-EvidenceCsv -Folder $folder -Name $Name -Rows $rows
                Add-Log $Name 'OK' "$count row(s)"
            }
            catch {
                $evidence[$Name] = @()
                Add-Log $Name 'Failed' $_.Exception.Message
            }
        }

        Save-Evidence '01_Account' {
            [pscustomobject]@{
                DisplayName = if ($cloudUser) { $cloudUser.displayName } else { $adUser.Name }
                UserPrincipalName = $upn; Mail = $cloudUser.mail; CloudObjectId = $cloudUser.id; UserType = $cloudUser.userType
                CloudEnabled = $cloudUser.accountEnabled; SyncedFromAD = [bool]$isSynced; Created = ConvertTo-UtcText $cloudUser.createdDateTime
                CloudPasswordLastChanged = ConvertTo-UtcText $cloudUser.lastPasswordChangeDateTime; SessionsValidFrom = ConvertTo-UtcText $cloudUser.signInSessionsValidFromDateTime
                Licensed = @($cloudUser.assignedLicenses).Count -gt 0; JobTitle = $cloudUser.jobTitle; Department = $cloudUser.department
                AdSamAccountName = $adUser.SamAccountName; AdDistinguishedName = $adUser.DistinguishedName; AdEnabled = $adUser.Enabled
                AdLockedOut = $adUser.LockedOut; AdPasswordLastSet = ConvertTo-UtcText $adUser.PasswordLastSet; AdLastLogon = ConvertTo-UtcText $adUser.LastLogonDate
                AdBadLogonCount = $adUser.badPwdCount; AdLastBadPassword = ConvertTo-UtcText $adUser.LastBadPasswordAttempt
            }
        }

        if ($cloudUser) {
            $id = $cloudUser.id
            Save-Evidence '02_Groups' {
                Get-GraphAll "v1.0/users/$id/memberOf?`$select=id,displayName" | ForEach-Object {
                    [pscustomobject]@{ Source = 'Entra ID'; Type = ("$($_.'@odata.type')" -replace '^#microsoft\.graph\.', ''); Name = $_.displayName; Id = $_.id }
                }
                foreach ($g in @($adUser.MemberOf | Where-Object { $_ })) { [pscustomobject]@{ Source = 'Active Directory'; Type = 'group'; Name = ($g -replace '^CN=([^,]+),.*$', '$1'); Id = $g } }
            }
            Save-Evidence '03_MfaMethods' {
                Get-GraphAll "v1.0/users/$id/authentication/methods" | ForEach-Object {
                    $type = "$($_.'@odata.type')" -replace '^#microsoft\.graph\.', ''
                    $name = ConvertFrom-MethodObject $_
                    if ($type -eq 'passwordAuthenticationMethod') { $name = 'Password' }
                    [pscustomobject]@{
                        Method  = if ($methodNames.ContainsKey("$name")) { $methodNames["$name"] } else { $name }
                        Detail  = (@($_.phoneNumber, $_.emailAddress, $_.displayName, $_.deviceTag, $_.model) | Where-Object { $_ }) -join ' / '
                        Created = ConvertTo-UtcText $_.createdDateTime
                        Type    = $type
                        Id      = $_.id
                    }
                }
            }
            Save-Evidence '04_SignIns' {
                $uri = "v1.0/auditLogs/signIns?`$filter=userId eq '$id' and createdDateTime ge $sinceIso&`$top=500"
                $n = 0
                while ($uri -and $n -lt 5000) {
                    $page = Invoke-Graph $uri
                    foreach ($s in @($page.value)) {
                        $n++
                        $ok = [int]"0$($s.status.errorCode)" -eq 0
                        [pscustomobject]@{
                            Time = ConvertTo-UtcText $s.createdDateTime; Result = if ($ok) { 'Success' } else { "Failed: $($s.status.failureReason) ($($s.status.errorCode))" }
                            App = $s.appDisplayName; Resource = $s.resourceDisplayName; IP = $s.ipAddress
                            City = $s.location.city; Country = $s.location.countryOrRegion; ClientApp = $s.clientAppUsed
                            Device = (@($s.deviceDetail.operatingSystem, $s.deviceDetail.browser) | Where-Object { $_ }) -join ' / '
                            MfaRequirement = $s.authenticationRequirement; ConditionalAccess = $s.conditionalAccessStatus
                            Risk = $s.riskLevelDuringSignIn; CorrelationId = $s.correlationId
                        }
                    }
                    $uri = $page.'@odata.nextLink'
                }
            }
            Save-Evidence '05_SignInLocations' {
                $evidence['04_SignIns'] | Group-Object IP | ForEach-Object {
                    $g = @($_.Group | Sort-Object Time)
                    [pscustomobject]@{
                        IP = $_.Name; Country = $g[0].Country; City = $g[0].City; SignIns = $g.Count
                        Successes = @($g | Where-Object Result -eq 'Success').Count; Failures = @($g | Where-Object Result -ne 'Success').Count
                        FirstSeen = $g[0].Time; LastSeen = $g[-1].Time; Apps = (@($g.App | Select-Object -Unique) -join ', ')
                    }
                } | Sort-Object SignIns -Descending
            }
            Save-Evidence '06_DirectoryAudits' {
                $rows = @()
                foreach ($q in @(
                    @{ Direction = 'Changes to this user'; Filter = "activityDateTime ge $sinceIso and targetResources/any(t:t/id eq '$id')" },
                    @{ Direction = 'Actions by this user'; Filter = "activityDateTime ge $sinceIso and initiatedBy/user/id eq '$id'" })) {
                    foreach ($a in (Get-GraphAll "v1.0/auditLogs/directoryAudits?`$filter=$($q.Filter)")) {
                        $rows += [pscustomobject]@{
                            Time = ConvertTo-UtcText $a.activityDateTime; Direction = $q.Direction; Activity = $a.activityDisplayName; Category = $a.category; Result = $a.result
                            InitiatedBy = (@($a.initiatedBy.user.userPrincipalName, $a.initiatedBy.app.displayName) | Where-Object { $_ }) -join ''
                            Targets = (@($a.targetResources | ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }) -join '; ')
                            Changes = (@($a.targetResources | ForEach-Object { $_.modifiedProperties } | Where-Object { $_ } | ForEach-Object { $_.displayName }) | Select-Object -Unique) -join ', '
                        }
                    }
                }
                $rows | Sort-Object Time
            }
            Save-Evidence '07_Devices' {
                $seen = @{}
                foreach ($kind in 'registeredDevices', 'ownedDevices') {
                    foreach ($d in (Get-GraphAll "v1.0/users/$id/$kind")) {
                        if ($seen.ContainsKey($d.id)) { continue }; $seen[$d.id] = $true
                        [pscustomobject]@{
                            Name = $d.displayName; OS = (@($d.operatingSystem, $d.operatingSystemVersion) | Where-Object { $_ }) -join ' '; TrustType = $d.trustType
                            Registered = ConvertTo-UtcText $d.registrationDateTime; LastSignIn = ConvertTo-UtcText $d.approximateLastSignInDateTime
                            Compliant = $d.isCompliant; Managed = $d.isManaged; Enabled = $d.accountEnabled; DeviceId = $d.deviceId
                        }
                    }
                }
            }
            Save-Evidence '08_AppConsents' {
                $names = @{}
                Get-GraphAll "v1.0/oauth2PermissionGrants?`$filter=principalId eq '$id'" | Where-Object { $_.principalId -eq $id } | ForEach-Object {
                    if (-not $names.ContainsKey($_.clientId)) {
                        try { $names[$_.clientId] = (Invoke-Graph "v1.0/servicePrincipals/$($_.clientId)?`$select=displayName,appOwnerOrganizationId").displayName } catch { $names[$_.clientId] = $_.clientId }
                    }
                    [pscustomobject]@{ App = $names[$_.clientId]; Permissions = $_.scope; ConsentType = $_.consentType; ClientId = $_.clientId }
                }
            }
        }

        $internal = @()
        if ($mailbox) {
            try { $internal = @(Get-AcceptedDomain | ForEach-Object { "$($_.DomainName)".ToLower() }) } catch { Add-Log 'Accepted domains' 'Failed' $_.Exception.Message }
            Save-Evidence '09_MailboxForwarding' {
                [pscustomobject]@{ Mailbox = "$($mailbox.PrimarySmtpAddress)"; ForwardingSmtpAddress = "$($mailbox.ForwardingSmtpAddress)"; ForwardingAddress = "$($mailbox.ForwardingAddress)"; DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward }
            }
            Save-Evidence '10_InboxRules' {
                Get-InboxRule -Mailbox $upn -IncludeHidden -ErrorAction Stop -WarningAction SilentlyContinue | ForEach-Object {
                    $a = Get-InboxRuleAssessment $_ $internal
                    [pscustomobject]@{
                        Risk = $a.Risk; Rule = "$($_.Name)"; Enabled = [bool]$_.Enabled; Actions = $a.Actions; ForwardsTo = $a.Targets; Conditions = $a.Conditions
                        Issues = ($a.Issues -join '; '); Description = "$($_.Description)" -replace '\s+', ' '; Identity = "$($_.Identity)"
                    }
                }
            }
        }

        # --- 2. Containment ----------------------------------------------------------

        $isAppOnly = (Get-Command Get-MgContext -ErrorAction SilentlyContinue) -and "$((Get-MgContext).AuthType)" -eq 'AppOnly'
        function Invoke-Step {
            # Uses the script's own -WhatIf/-Confirm handling ($PSCmdlet of the script)
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
            param([string]$Step, [string]$Target, [string]$What, [scriptblock]$Do)
            if ($PSCmdlet.ShouldProcess($Target, $What)) {
                try { $detail = & $Do; Add-Action $Step $(if ($detail -is [array]) { $detail[0] } else { 'Done' }) $(if ($detail -is [array]) { $detail[1] } else { "$detail" }) }
                catch { Add-Action $Step 'Failed' $_.Exception.Message }
            }
            elseif ($WhatIfPreference) { Add-Action $Step 'WhatIf' "Would: $What" }
            else { Add-Action $Step 'Skipped' 'Not confirmed' }
        }

        if ($EvidenceOnly) {
            Write-Host 'Evidence only - no changes made.'
        }
        else {
            Write-Host 'Containing account...'

            # Disable
            if ($Skip -contains 'Disable') { Add-Action 'Disable' 'Skipped' '-Skip' }
            elseif ($isSynced -or -not $cloudUser) {
                if ($adUser) {
                    Invoke-Step 'Disable' "$($adUser.SamAccountName) (Active Directory)" 'Disable the account in Active Directory' {
                        Disable-ADAccount -Identity $adUser.DistinguishedName -Confirm:$false -ErrorAction Stop
                        if ($isSynced) { 'Disabled in AD. Entra ID follows at the next sync - run Start-ADSyncSyncCycle -PolicyType Delta on the Entra Connect server to apply it now.' }
                        elseif ($SkipCloud) { 'Disabled in AD. If the account syncs to Entra ID, run a delta sync so the change applies there now.' }
                        else { 'Disabled in AD.' }
                    }
                }
                else { Add-Action 'Disable' 'Not done' 'The account is synced from AD but AD is not available here - disable it in AD now.' }
            }
            else {
                Invoke-Step 'Disable' $upn 'Block sign-in (disable the account) in Entra ID' {
                    Invoke-GraphChange PATCH "v1.0/users/$($cloudUser.id)" @{ accountEnabled = $false } | Out-Null
                    'Sign-in blocked in Entra ID.'
                }
            }

            # Reset password
            if ($Skip -contains 'ResetPassword') { Add-Action 'ResetPassword' 'Skipped' '-Skip' }
            elseif (($isSynced -or -not $cloudUser) -and $adUser) {
                Invoke-Step 'ResetPassword' "$($adUser.SamAccountName) (Active Directory)" 'Reset the password to a random value in Active Directory' {
                    $secure = New-Object System.Security.SecureString
                    foreach ($ch in (New-RandomPassword).ToCharArray()) { $secure.AppendChar($ch) }
                    $secure.MakeReadOnly()
                    Set-ADAccountPassword -Identity $adUser.DistinguishedName -Reset -NewPassword $secure -Confirm:$false -ErrorAction Stop
                    if ($isSynced -or $SkipCloud) { 'Random password set in AD (not shown or saved). It syncs to Entra ID if password hash sync is on.' } else { 'Random password set in AD (not shown or saved).' }
                }
            }
            elseif ($isSynced) { Add-Action 'ResetPassword' 'Not done' 'The account is synced from AD but AD is not available here - reset the password in AD now.' }
            elseif ($isAppOnly) { Add-Action 'ResetPassword' 'Not done' 'Password reset needs a signed-in admin (it is not available with app-only sign-in).' }
            else {
                Invoke-Step 'ResetPassword' $upn 'Reset the password to a random value in Entra ID' {
                    Invoke-GraphChange POST "v1.0/users/$($cloudUser.id)/authentication/methods/$passwordMethodId/resetPassword" @{ newPassword = (New-RandomPassword) } | Out-Null
                    , @('Requested', 'Random password set (not shown or saved). Microsoft applies it within a minute or so.')
                }
            }

            # Revoke sessions
            if ($Skip -contains 'RevokeSessions') { Add-Action 'RevokeSessions' 'Skipped' '-Skip' }
            elseif (-not $cloudUser) { Add-Action 'RevokeSessions' 'Skipped' 'No cloud account' }
            else {
                Invoke-Step 'RevokeSessions' $upn 'Revoke all sign-in sessions and refresh tokens' {
                    Invoke-GraphChange POST "v1.0/users/$($cloudUser.id)/revokeSignInSessions" $null | Out-Null
                    'All sessions and refresh tokens revoked (can take a few minutes to take effect).'
                }
            }

            # Mailbox
            if ($Skip -contains 'Mailbox') { Add-Action 'Mailbox' 'Skipped' '-Skip' }
            elseif (-not $mailbox) { Add-Action 'Mailbox' 'Skipped' $(if ($SkipExchange) { '-SkipExchange' } else { 'No mailbox found, or Exchange Online not connected (see 00_CollectionLog.csv)' }) }
            else {
                $fwd = (@(("$($mailbox.ForwardingSmtpAddress)" -replace '^(?i)smtp:', ''), "$($mailbox.ForwardingAddress)") | Where-Object { $_ }) -join ', '
                if ($fwd) {
                    Invoke-Step 'Mailbox' $upn "Remove mailbox forwarding to $fwd" {
                        Set-Mailbox -Identity $upn -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false -Confirm:$false -ErrorAction Stop
                        "Removed forwarding to $fwd"
                    }
                }
                $badRules = @($evidence['10_InboxRules'] | Where-Object { $_.Risk -in 'Critical', 'High', 'Medium' -and $_.Enabled })
                foreach ($r in $badRules) {
                    Invoke-Step 'Mailbox' "$upn rule '$($r.Rule)'" "Disable inbox rule '$($r.Rule)' ($($r.Issues))" {
                        Disable-InboxRule -Identity $r.Identity -Confirm:$false -ErrorAction Stop
                        "Disabled inbox rule '$($r.Rule)' (kept, not deleted, for evidence)"
                    }
                }
                if (-not $fwd -and -not $badRules.Count) { Add-Action 'Mailbox' 'Done' 'No forwarding or suspicious enabled inbox rules found' }
            }

            # MFA methods (only when asked)
            if ($PSBoundParameters.ContainsKey('RemoveMfaMethodsAddedAfter') -and $cloudUser) {
                $cutoffText = ConvertTo-UtcText $RemoveMfaMethodsAddedAfter
                $recent = @($evidence['03_MfaMethods'] | Where-Object { $_.Created -and $_.Created -gt $cutoffText -and $methodSegments.ContainsKey($_.Type) })
                if (-not $recent.Count) { Add-Action 'MfaMethods' 'Done' "No MFA methods added after $cutoffText" }
                foreach ($m in $recent) {
                    Invoke-Step 'MfaMethods' $upn "Remove MFA method $($m.Method) ($($m.Detail)) added $($m.Created)" {
                        Invoke-GraphChange DELETE "v1.0/users/$($cloudUser.id)/authentication/$($methodSegments[$m.Type])/$($m.Id)" $null | Out-Null
                        "Removed $($m.Method) ($($m.Detail)) added $($m.Created)"
                    }
                }
                $undated = @($evidence['03_MfaMethods'] | Where-Object { -not $_.Created -and $_.Method -ne 'Password' })
                if ($undated.Count) { $notes.Add("These MFA methods have no creation date, so they were not checked: $(($undated | ForEach-Object { "$($_.Method) $($_.Detail)" }) -join '; '). Confirm them with the user.") }
            }
        }

        # --- 3. Indicators and next steps -------------------------------------------

        $indicators = [System.Collections.Generic.List[object]]::new()
        $si = @($evidence['04_SignIns'])
        $countries = @($si | Where-Object { $_.Result -eq 'Success' -and $_.Country } | Select-Object -ExpandProperty Country -Unique)
        if ($countries.Count -gt 1) { $indicators.Add([pscustomobject]@{ Indicator = 'Successful sign-ins from more than one country'; Detail = ($countries -join ', ') }) }
        $failed = @($si | Where-Object { $_.Result -ne 'Success' })
        if ($failed.Count -ge 10) { $indicators.Add([pscustomobject]@{ Indicator = "$($failed.Count) failed sign-ins in $SignInDays days"; Detail = (@($failed | Select-Object -ExpandProperty IP -Unique | Select-Object -First 10) -join ', ') }) }
        $legacy = @($si | Where-Object { $_.ClientApp -match 'IMAP|POP|SMTP|Exchange ActiveSync|Other clients|Authenticated SMTP' -and $_.Result -eq 'Success' })
        if ($legacy.Count) { $indicators.Add([pscustomobject]@{ Indicator = 'Successful legacy-protocol sign-ins (no MFA possible)'; Detail = (@($legacy.ClientApp | Select-Object -Unique) -join ', ') }) }
        $recentMfa = @($evidence['03_MfaMethods'] | Where-Object { $_.Created -and $_.Created -ge (ConvertTo-UtcText $since) })
        if ($recentMfa.Count) { $indicators.Add([pscustomobject]@{ Indicator = "MFA method(s) added in the last $SignInDays days"; Detail = (($recentMfa | ForEach-Object { "$($_.Method) $($_.Detail) ($($_.Created))" }) -join '; ') }) }
        $badRuleRows = @($evidence['10_InboxRules'] | Where-Object { $_.Risk -in 'Critical', 'High', 'Medium' })
        if ($badRuleRows.Count) { $indicators.Add([pscustomobject]@{ Indicator = 'Suspicious inbox rules'; Detail = (($badRuleRows | ForEach-Object { "'$($_.Rule)': $($_.Issues -replace ' - common.*$', '')" }) -join '; ') }) }
        $fwdRow = @($evidence['09_MailboxForwarding'])[0]
        if ($fwdRow -and ($fwdRow.ForwardingSmtpAddress -or $fwdRow.ForwardingAddress)) { $indicators.Add([pscustomobject]@{ Indicator = 'Mailbox forwarding was set'; Detail = ("$($fwdRow.ForwardingSmtpAddress) $($fwdRow.ForwardingAddress)" -replace '(?i)smtp:', '').Trim() }) }
        $userConsents = @($evidence['08_AppConsents'] | Where-Object { $_.ConsentType -eq 'Principal' })
        if ($userConsents.Count) { $indicators.Add([pscustomobject]@{ Indicator = 'Apps the user consented to'; Detail = (@($userConsents.App | Select-Object -Unique) -join ', ') }) }
        $byUser = @($evidence['06_DirectoryAudits'] | Where-Object { $_.Direction -eq 'Actions by this user' -and $_.Activity -match '(?i)security info|consent|add member|add owner|add app|reset password|update application' })
        if ($byUser.Count) { $indicators.Add([pscustomobject]@{ Indicator = 'Sensitive changes made by this account'; Detail = (@($byUser | ForEach-Object { "$($_.Time) $($_.Activity)" }) | Select-Object -First 10) -join '; ' }) }

        $nextSteps = @(
            'Review 04_SignIns.csv and 05_SignInLocations.csv for attacker IPs. Block them (Conditional Access named location) and search other accounts for the same IPs.'
            'Check 03_MfaMethods.csv with the user - remove any method they did not add (or re-run with -RemoveMfaMethodsAddedAfter <date>).'
            'Review 08_AppConsents.csv and revoke any app the user does not recognise.'
            'Check whether the account sent phishing or read sensitive mail: message trace and the unified audit log (Purview) for the same period.'
            'Review 07_Devices.csv for devices the user does not recognise and disable them.'
        )
        if ($isSynced) { $nextSteps += 'The account is synced from AD: run a delta sync (Start-ADSyncSyncCycle -PolicyType Delta) so the disable reaches Entra ID now, and check the AD account for other signs of compromise.' }
        $nextSteps += 'To restore access: verify the user out-of-band, have them set a new password, require MFA re-registration, re-enable the account and re-enable any inbox rules that were legitimate.'
        $nextSteps += @($notes)

        [void](Export-EvidenceCsv -Folder $folder -Name '00_Actions' -Rows $actions)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_Indicators' -Rows $indicators)
        [void](Export-EvidenceCsv -Folder $folder -Name '00_CollectionLog' -Rows $log)

        $runBy = if ($graphConnected -or $cloudUser) { Get-GraphRunBy } else { "$env:USERDOMAIN\$env:USERNAME" }
        $html = New-HtmlReport -Title "Account containment - $upn" -IntroLines @(
            "User: $display - $(if ($isSynced) { 'synced from AD' } elseif ($cloudUser) { 'cloud-only' } else { 'AD only' })",
            "Mode: $mode. Run $(ConvertTo-UtcText $start) by $runBy$(if ($CaseId) { " - case $CaseId" }). All times are UTC.",
            "Evidence covers the last $SignInDays day(s)."
        ) -Sections @(
            @{ Heading = 'Actions'; Html = ConvertTo-HtmlTable -Rows $actions -Columns @('Time', 'Step', 'Result', 'Detail') -StatusColumn 'Result' -EmptyText 'No changes made.' }
            @{ Heading = "Indicators ($($indicators.Count))"; Html = ConvertTo-HtmlTable -Rows $indicators -Columns @('Indicator', 'Detail') -EmptyText 'No automatic indicators. Review the evidence files.' }
            @{ Heading = 'Sign-in locations'; Html = ConvertTo-HtmlTable -Rows @($evidence['05_SignInLocations'] | Select-Object -First 20) -Columns @('IP', 'Country', 'City', 'SignIns', 'Successes', 'Failures', 'LastSeen', 'Apps') -EmptyText 'No sign-in data (needs Entra ID P1, and AuditLog.Read.All).' }
            @{ Heading = 'Next steps'; Html = ConvertTo-HtmlTable -Rows @($nextSteps | ForEach-Object { [pscustomobject]@{ Step = $_ } }) -Columns @('Step') }
            @{ Heading = 'Evidence collected'; Html = ConvertTo-HtmlTable -Rows $log -Columns @('Section', 'Status', 'Detail') -StatusColumn 'Status' }
        )
        [System.IO.File]::WriteAllText((Join-Path $folder '00_Summary.html'), $html, $script:Utf8NoBom)
        $package = Complete-EvidencePackage -Folder $folder

        Write-Host ""
        if ($indicators.Count) {
            Write-Host 'Indicators:' -ForegroundColor Yellow
            foreach ($i in $indicators) { Write-Host "  - $($i.Indicator): $($i.Detail)" -ForegroundColor Yellow }
        }
        $problems = @($log | Where-Object { $_.Status -ne 'OK' })
        if ($problems.Count) { Write-Host ("Some evidence could not be collected: {0}" -f (($problems | ForEach-Object { "$($_.Section) ($($_.Status))" }) -join ', ')) -ForegroundColor Yellow }
        Write-Host "Evidence : $($package.Zip)"
        Write-Host "SHA-256  : $($package.Sha256)  (record this in your case notes)"
        Write-Host "Summary  : $(Join-Path $folder '00_Summary.html')"
    }
    finally {
        if ($exoConnected)   { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
        if ($graphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
