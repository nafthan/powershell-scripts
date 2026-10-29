<#
.SYNOPSIS
    Reports MFA status for every Microsoft 365 user to Excel.
.DESCRIPTION
    Lists sign-in status, last sign-in, creation date, MFA registration and methods, admin roles,
    and flags users who need attention. Read-only.
.EXAMPLE
    .\Get-MfaStatusReport.ps1
.EXAMPLE
    .\Get-MfaStatusReport.ps1 -IncludeGuests -IncludePerUserMfaState -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: Microsoft.Graph.Authentication module, Global Reader (or Reports Reader + Security Reader)
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [switch]$IncludeGuests,
    [switch]$IncludePerUserMfaState,
    [switch]$UseMethodsApi,

    [ValidateRange(1, 3650)]
    [int]$StaleDays,

    [switch]$SendEmail,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('TenantId'))              { $TenantId              = Get-Setting $config 'M365' 'TenantId' '' }
if (-not $PSBoundParameters.ContainsKey('ClientId'))              { $ClientId              = Get-Setting $config 'M365' 'ClientId' '' }
if (-not $PSBoundParameters.ContainsKey('CertificateThumbprint')) { $CertificateThumbprint = Get-Setting $config 'M365' 'CertificateThumbprint' '' }
if (-not $PSBoundParameters.ContainsKey('IncludeGuests'))         { $IncludeGuests         = [bool](Get-Setting $config 'M365' 'IncludeGuests' $false) }
if (-not $PSBoundParameters.ContainsKey('StaleDays'))             { $StaleDays             = [int](Get-Setting $config 'M365' 'StaleDays' 90) }
$newAccountDays = [int](Get-Setting $config 'M365' 'NewAccountGraceDays' 14)

$scopes = @('User.Read.All', 'AuditLog.Read.All', 'UserAuthenticationMethod.Read.All', 'Policy.Read.All', 'RoleManagement.Read.Directory')

# --- Graph helpers --------------------------------------------------------

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
$defaultMethodNames = @{
    push = 'Authenticator push'; oath = 'OATH code'; sms = 'SMS'; voiceMobile = 'Phone call (mobile)'
    voiceAlternateMobile = 'Phone call (alternate)'; voiceOffice = 'Phone call (office)'; none = ''
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

# --- Main -----------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'The Microsoft Graph PowerShell module is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # --- Connect ----------------------------------------------------------

    $connectedHere = $false
    $context = Get-MgContext
    if ($context -and (-not $TenantId -or $context.TenantId -eq $TenantId -or $TenantId -like '*.*')) {
        Write-Host "Using existing Graph connection ($($context.Account)$($context.AppName))."
        if ($context.AuthType -eq 'Delegated') {
            $missing = @($scopes | Where-Object { $context.Scopes -notcontains $_ })
            if ($missing.Count) { Write-Warning "Current connection is missing scope(s): $($missing -join ', '). Some sections may be empty. Run Disconnect-MgGraph and try again." }
        }
    }
    else {
        $connect = @{ NoWelcome = $true; ErrorAction = 'Stop' }
        if ($TenantId) { $connect.TenantId = $TenantId }
        if ($ClientId -and $CertificateThumbprint) {
            if (-not $TenantId) { throw 'App-only sign-in needs a tenant. Set M365.TenantId in config.psd1 or use -TenantId.' }
            $connect.ClientId = $ClientId
            $connect.CertificateThumbprint = $CertificateThumbprint
            Write-Host 'Connecting to Microsoft Graph (app-only, certificate)...'
        }
        else {
            $connect.Scopes = $scopes
            if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
            Write-Host 'Connecting to Microsoft Graph - sign in with an account that has Global Reader (or Reports Reader + Security Reader)...'
        }
        Connect-MgGraph @connect | Out-Null
        $connectedHere = $true
        $context = Get-MgContext
    }

    $now        = Get-Date
    $reportPath = New-ReportPath -Config $config -Prefix 'MfaStatus' -OutputFile $OutputFile
    $errors     = [System.Collections.Generic.List[object]]::new()
    function Add-ReportError([string]$Section, [string]$Message) {
        $errors.Add([pscustomobject]@{ Section = $Section; Error = $Message })
    }

    try {
        # --- Tenant -------------------------------------------------------

        $tenantName = $context.TenantId
        try {
            $org = @(Get-GraphAll 'v1.0/organization?$select=displayName,verifiedDomains') | Select-Object -First 1
            $defaultDomain = @($org.verifiedDomains | Where-Object { $_.isDefault }) | Select-Object -First 1
            $tenantName = "$($org.displayName) ($($defaultDomain.name))"
        }
        catch { Add-ReportError 'Tenant' $_.Exception.Message }
        Write-Host "Tenant : $tenantName"

        # --- Users --------------------------------------------------------

        Write-Host 'Getting users...'
        $select = 'id,displayName,userPrincipalName,mail,accountEnabled,userType,createdDateTime,assignedLicenses,onPremisesSyncEnabled,department,jobTitle'
        $signInAvailable = $true
        try {
            $users = @(Get-GraphAll ("v1.0/users?`$select=$select,signInActivity&`$top=500"))
        }
        catch {
            $signInAvailable = $false
            Add-ReportError 'Sign-in activity' "Last sign-in dates unavailable (needs Entra ID P1/P2 and AuditLog.Read.All): $($_.Exception.Message)"
            Write-Warning 'Last sign-in dates are unavailable (needs Entra ID P1/P2 and AuditLog.Read.All). Continuing without them.'
            $users = @(Get-GraphAll ("v1.0/users?`$select=$select&`$top=999"))
        }
        $allUserCount = $users.Count
        if (-not $IncludeGuests) { $users = @($users | Where-Object { $_.userType -ne 'Guest' }) }
        Write-Host ("  {0:N0} user(s){1}" -f $users.Count, $(if (-not $IncludeGuests) { ' (members only - use -IncludeGuests for guests)' } else { '' }))

        $userById = @{}
        foreach ($u in $users) { $userById[$u.id] = $u }

        # --- MFA registration --------------------------------------------

        $registration = @{}
        $source = 'Registration report'
        if (-not $UseMethodsApi) {
            Write-Host 'Getting MFA registration details...'
            try {
                foreach ($r in (Get-GraphAll 'v1.0/reports/authenticationMethods/userRegistrationDetails')) { $registration[$r.id] = $r }
            }
            catch {
                Add-ReportError 'Registration report' "Unavailable (needs Entra ID P1/P2 and AuditLog.Read.All), so each user's methods were read one by one: $($_.Exception.Message)"
                Write-Warning 'MFA registration report unavailable (needs Entra ID P1/P2). Reading each user''s methods instead - this is slower.'
            }
        }
        $needMethods = @($users | Where-Object { -not $registration.ContainsKey($_.id) })
        $methodsById = @{}
        if ($needMethods.Count) {
            if ($registration.Count) { $source = "Registration report, plus methods API for $($needMethods.Count) user(s) not in the report" }
            else                     { $source = 'Authentication methods API (per user)' }
            $i = 0
            foreach ($u in $needMethods) {
                $i++
                if ($i % 10 -eq 0 -or $i -eq 1) { Write-Progress -Activity 'Reading authentication methods' -Status "$i of $($needMethods.Count)" -PercentComplete ($i / $needMethods.Count * 100) }
                try {
                    $methodsById[$u.id] = @(Get-GraphAll "v1.0/users/$($u.id)/authentication/methods" | ForEach-Object { ConvertFrom-MethodObject $_ } | Where-Object { $_ } | Sort-Object -Unique)
                }
                catch {
                    Add-ReportError 'Authentication methods' "$($u.userPrincipalName): $($_.Exception.Message)"
                }
            }
            Write-Progress -Activity 'Reading authentication methods' -Completed
        }

        # --- Per-user MFA (legacy) ----------------------------------------

        $perUserMfa = @{}
        if ($IncludePerUserMfaState) {
            $i = 0
            foreach ($u in $users) {
                $i++
                if ($i % 10 -eq 0 -or $i -eq 1) { Write-Progress -Activity 'Reading per-user MFA state' -Status "$i of $($users.Count)" -PercentComplete ($i / $users.Count * 100) }
                try   { $perUserMfa[$u.id] = "$((Invoke-Graph "beta/users/$($u.id)/authentication/requirements").perUserMfaState)" }
                catch { $perUserMfa[$u.id] = 'Unknown'; Add-ReportError 'Per-user MFA' "$($u.userPrincipalName): $($_.Exception.Message)" }
            }
            Write-Progress -Activity 'Reading per-user MFA state' -Completed
        }

        # --- Admin roles --------------------------------------------------

        Write-Host 'Getting admin roles...'
        $rolesById = @{}   # user id -> list of role names
        try {
            $roleNames = @{}
            foreach ($d in (Get-GraphAll 'v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName')) { $roleNames[$d.id] = $d.displayName }

            $assignments = @()
            $assignments += @(Get-GraphAll 'v1.0/roleManagement/directory/roleAssignments?$select=principalId,roleDefinitionId' |
                              ForEach-Object { [pscustomobject]@{ PrincipalId = $_.principalId; Role = $roleNames[$_.roleDefinitionId]; Eligible = $false } })
            try {
                $assignments += @(Get-GraphAll 'v1.0/roleManagement/directory/roleEligibilitySchedules?$select=principalId,roleDefinitionId' |
                                  ForEach-Object { [pscustomobject]@{ PrincipalId = $_.principalId; Role = $roleNames[$_.roleDefinitionId]; Eligible = $true } })
            }
            catch { Add-ReportError 'Eligible roles (PIM)' "Not included (needs Entra ID P2): $($_.Exception.Message)" }

            $groupMembers = @{}
            foreach ($a in $assignments) {
                if (-not $a.Role) { continue }
                $label = if ($a.Eligible) { "$($a.Role) (eligible)" } else { $a.Role }
                $targets = @()
                if ($userById.ContainsKey($a.PrincipalId)) {
                    $targets = @($a.PrincipalId)
                }
                else {
                    # Role-assignable group (expand its members); service principals and excluded users are skipped
                    if (-not $groupMembers.ContainsKey($a.PrincipalId)) {
                        try   { $groupMembers[$a.PrincipalId] = @(Get-GraphAll "v1.0/groups/$($a.PrincipalId)/transitiveMembers/microsoft.graph.user?`$select=id" | ForEach-Object { $_.id }) }
                        catch { $groupMembers[$a.PrincipalId] = @() }
                    }
                    $targets = $groupMembers[$a.PrincipalId]
                    if ($targets.Count) { $label += ' (via group)' }
                }
                foreach ($t in $targets) {
                    if (-not $rolesById.ContainsKey($t)) { $rolesById[$t] = [System.Collections.Generic.List[string]]::new() }
                    if (-not $rolesById[$t].Contains($label)) { $rolesById[$t].Add($label) }
                }
            }
        }
        catch { Add-ReportError 'Admin roles' $_.Exception.Message }

        # --- Tenant MFA policies -----------------------------------------

        Write-Host 'Checking security defaults and Conditional Access...'
        $securityDefaults = 'Unknown'
        try {
            $sd = Invoke-Graph 'v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
            $securityDefaults = if ($sd.isEnabled) { 'Enabled' } else { 'Disabled' }
        }
        catch { Add-ReportError 'Security defaults' $_.Exception.Message }

        $caRows = @()
        try {
            $caRows = @(Get-GraphAll 'v1.0/identity/conditionalAccess/policies' | ForEach-Object {
                $gc = $_.grantControls
                $mfa = ($gc -and ((@($gc.builtInControls) -contains 'mfa') -or $gc.authenticationStrength))
                if (-not $mfa) { return }
                $cu   = $_.conditions.users
                $capps = $_.conditions.applications
                $describe = {
                    param($List, [string]$Noun)
                    $items = @($List | Where-Object { $_ })
                    if ($items -contains 'All') { "All $Noun" } elseif ($items -contains 'None') { 'None' } elseif ($items.Count) { "$($items.Count) $Noun" } else { '' }
                }
                $include = @((& $describe $cu.includeUsers 'users'), (& $describe $cu.includeGroups 'groups'), (& $describe $cu.includeRoles 'roles')) | Where-Object { $_ }
                $exclude = @($cu.excludeUsers | Where-Object { $_ }).Count + @($cu.excludeGroups | Where-Object { $_ }).Count + @($cu.excludeRoles | Where-Object { $_ }).Count
                [pscustomobject]@{
                    Policy        = $_.displayName
                    State         = switch ($_.state) { 'enabled' { 'Enabled' } 'enabledForReportingButNotEnforced' { 'Report-only' } default { 'Disabled' } }
                    Requires      = if ($gc.authenticationStrength) { "Authentication strength: $($gc.authenticationStrength.displayName)" } else { 'MFA' }
                    AppliesTo     = ($include -join ', ')
                    Exclusions    = $exclude
                    Applications  = (& $describe $capps.includeApplications 'apps')
                    Modified      = ConvertTo-LocalTime $(if ($_.modifiedDateTime) { $_.modifiedDateTime } else { $_.createdDateTime })
                }
            })
        }
        catch { Add-ReportError 'Conditional Access' "Could not read policies (needs Policy.Read.All and Entra ID P1): $($_.Exception.Message)" }
        $caEnabled = @($caRows | Where-Object { $_.State -eq 'Enabled' }).Count

        # --- Build rows --------------------------------------------------

        Write-Host 'Building report...'
        $riskRank = @{ OK = 0; Low = 1; Medium = 2; High = 3; Critical = 4 }
        function Add-Issue([string]$Level, [string]$Text) {
            # Uses $issues from the loop below and raises $script:currentRisk to the highest level seen
            $issues.Add($Text)
            if ($riskRank[$Level] -gt $riskRank[$script:currentRisk]) { $script:currentRisk = $Level }
        }
        $rows = foreach ($u in $users) {
            $reg = $registration[$u.id]
            if ($reg) {
                $methods      = @($reg.methodsRegistered | Where-Object { $_ })
                $mfaRegistered = [bool]$reg.isMfaRegistered
                $mfaCapable    = [bool]$reg.isMfaCapable
                $passwordless  = [bool]$reg.isPasswordlessCapable
                $ssprRegistered = if ($null -ne $reg.isSsprRegistered) { [bool]$reg.isSsprRegistered } else { $null }
                $preferred     = $defaultMethodNames["$($reg.userPreferredMethodForSecondaryAuthentication)"]
                if ($reg.isSystemPreferredAuthenticationMethodEnabled -and @($reg.systemPreferredAuthenticationMethods).Count) {
                    $sys = $defaultMethodNames["$(@($reg.systemPreferredAuthenticationMethods)[0])"]
                    $preferred = "System-preferred ($sys)"
                }
                if (-not $mfaRegistered) { $preferred = '' }
                $isAdmin = [bool]$reg.isAdmin
            }
            elseif ($methodsById.ContainsKey($u.id)) {
                $methods       = $methodsById[$u.id]
                $mfaRegistered = [bool](@($methods | Where-Object { Test-StrongMethod $_ }).Count)
                $mfaCapable    = $null
                $passwordless  = [bool](@($methods | Where-Object { $_ -match $passwordlessRx }).Count)
                $ssprRegistered = $null
                $preferred     = ''
                $isAdmin       = $false
            }
            else {
                $methods = @(); $mfaRegistered = $null; $mfaCapable = $null; $passwordless = $null; $ssprRegistered = $null; $preferred = ''; $isAdmin = $false
            }

            $roles   = if ($rolesById.ContainsKey($u.id)) { @($rolesById[$u.id] | Sort-Object) } else { @() }
            $isAdmin = $isAdmin -or ($roles.Count -gt 0)
            $strong  = @($methods | Where-Object { Test-StrongMethod $_ })
            $phishOk = [bool](@($methods | Where-Object { $_ -match $phishResistant }).Count)
            $phoneOnly = $strong.Count -gt 0 -and -not @($strong | Where-Object { $phoneMethods -notcontains $_ }).Count

            $sia            = $u.signInActivity
            $lastInteractive = ConvertTo-LocalTime $sia.lastSignInDateTime
            $lastNonInter    = ConvertTo-LocalTime $sia.lastNonInteractiveSignInDateTime
            $lastSuccessful  = ConvertTo-LocalTime $sia.lastSuccessfulSignInDateTime
            $lastSignIn      = if ($lastSuccessful) { $lastSuccessful } else { Get-Latest @($lastInteractive, $lastNonInter) }
            $created         = ConvertTo-LocalTime $u.createdDateTime
            $daysSince       = if ($lastSignIn) { [int][math]::Floor(($now - $lastSignIn).TotalDays) } else { $null }
            $isNew           = $created -and ($now - $created).TotalDays -le $newAccountDays

            # Risk rating (enabled accounts only)
            $issues = [System.Collections.Generic.List[string]]::new()
            $script:currentRisk = 'OK'
            if (-not $u.accountEnabled) {
                $script:currentRisk = 'Disabled'
            }
            else {
                if ($null -eq $mfaRegistered) { Add-Issue 'Medium' 'MFA status could not be read' }
                elseif (-not $mfaRegistered) {
                    if ($isAdmin) { Add-Issue 'Critical' 'Admin with no MFA registered' }
                    elseif ($u.userType -eq 'Guest') { Add-Issue 'Low' 'Guest with no MFA registered here - guests usually use MFA from their own tenant; check a Conditional Access policy requires MFA for guests' }
                    else          { Add-Issue 'High' $(if ($isNew) { "No MFA registered (new account - created $([int]($now - $created).TotalDays) day(s) ago)" } else { 'No MFA registered' }) }
                }
                else {
                    if ($phoneOnly) { Add-Issue $(if ($isAdmin) { 'High' } else { 'Medium' }) 'Only SMS/voice MFA registered' }
                    if ($mfaCapable -eq $false) { Add-Issue 'Medium' 'MFA method registered but not allowed by policy' }
                    if ($isAdmin -and -not $phishOk -and -not $phoneOnly) { Add-Issue 'Low' 'Admin without a phishing-resistant method' }
                }
                if ($signInAvailable) {
                    if (-not $lastSignIn -and -not $isNew) { Add-Issue 'Low' 'Never signed in - consider disabling' }
                    elseif ($null -ne $daysSince -and $daysSince -gt $StaleDays) { Add-Issue 'Low' "No sign-in for $daysSince days - consider disabling" }
                }
            }
            $risk = $script:currentRisk

            [pscustomobject]@{
                DisplayName              = $u.displayName
                UserPrincipalName        = $u.userPrincipalName
                Email                    = $u.mail
                SignInStatus             = if ($u.accountEnabled) { 'Allowed' } else { 'Blocked' }
                UserType                 = $u.userType
                Licensed                 = @($u.assignedLicenses).Count -gt 0
                AccountSource            = if ($u.onPremisesSyncEnabled) { 'Synced (AD)' } else { 'Cloud' }
                Created                  = $created
                LastSignIn               = if ($signInAvailable) { $lastSignIn } else { 'Unavailable' }
                DaysSinceSignIn          = $daysSince
                LastInteractiveSignIn    = $lastInteractive
                LastNonInteractiveSignIn = $lastNonInter
                MfaStatus                = if ($null -eq $mfaRegistered) { 'Unknown' } elseif ($mfaRegistered) { 'Registered' } else { 'Not registered' }
                MfaCapable               = if ($null -eq $mfaCapable) { '' } elseif ($mfaCapable) { 'Yes' } else { 'No' }
                PerUserMfa               = if ($IncludePerUserMfaState) { $perUserMfa[$u.id] } else { '' }
                DefaultMethod            = $preferred
                MethodsRegistered        = (@($methods | ForEach-Object { if ($methodNames.ContainsKey($_)) { $methodNames[$_] } else { $_ } }) -join ', ')
                PhishingResistant        = $phishOk
                PasswordlessCapable      = if ($null -eq $passwordless) { '' } elseif ($passwordless) { 'Yes' } else { 'No' }
                SsprRegistered           = if ($null -eq $ssprRegistered) { '' } elseif ($ssprRegistered) { 'Yes' } else { 'No' }
                IsAdmin                  = $isAdmin
                AdminRoles               = ($roles -join '; ')
                Risk                     = $risk
                Issues                   = ($issues -join '; ')
                Department               = $u.department
                JobTitle                 = $u.jobTitle
            }
        }
        $rows = @($rows)

        # --- Summaries ---------------------------------------------------

        $riskOrder = @{ Critical = 0; High = 1; Medium = 2; Low = 3; OK = 4; Disabled = 5 }
        $rows      = @($rows | Sort-Object @{ Expression = { $riskOrder[$_.Risk] } }, DisplayName)
        $attention = @($rows | Where-Object { $_.Risk -in 'Critical', 'High', 'Medium', 'Low' })
        $admins    = @($rows | Where-Object { $_.IsAdmin })
        $enabled   = @($rows | Where-Object { $_.SignInStatus -eq 'Allowed' })
        $enabledMfa = @($enabled | Where-Object { $_.MfaStatus -eq 'Registered' })
        $pct = if ($enabled.Count) { [math]::Round($enabledMfa.Count / $enabled.Count * 100, 1) } else { 0 }

        $methodCounts = @($enabled | ForEach-Object { $_.MethodsRegistered -split ', ' } | Where-Object { $_ } |
                          Group-Object | Sort-Object Count -Descending |
                          ForEach-Object { [pscustomobject]@{ Method = $_.Name; Users = $_.Count; PercentOfEnabled = if ($enabled.Count) { [math]::Round($_.Count / $enabled.Count * 100, 1) } else { 0 } } })

        $counts = [ordered]@{}
        foreach ($r in 'Critical', 'High', 'Medium', 'Low') { $counts[$r] = @($rows | Where-Object { $_.Risk -eq $r }).Count }

        $summary = [ordered]@{
            'Tenant'                           = $tenantName
            'Run at'                           = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Users in report'                  = "$($rows.Count) of $allUserCount" + $(if (-not $IncludeGuests) { ' (guests excluded)' } else { '' })
            'Sign-in allowed'                  = $enabled.Count
            'Sign-in blocked'                  = $rows.Count - $enabled.Count
            'Enabled users with MFA'           = "$($enabledMfa.Count) ($pct%)"
            'Enabled users without MFA'        = @($enabled | Where-Object { $_.MfaStatus -eq 'Not registered' }).Count
            'Admins'                           = $admins.Count
            'Admins without MFA'               = @($admins | Where-Object { $_.SignInStatus -eq 'Allowed' -and $_.MfaStatus -ne 'Registered' }).Count
            'Phishing-resistant users'         = @($enabled | Where-Object { $_.PhishingResistant }).Count
            'Risk: Critical'                   = $counts.Critical
            'Risk: High'                       = $counts.High
            'Risk: Medium'                     = $counts.Medium
            'Risk: Low'                        = $counts.Low
            'Security defaults'                = $securityDefaults
            'Conditional Access requiring MFA' = "$caEnabled enabled, $(@($caRows | Where-Object { $_.State -eq 'Report-only' }).Count) report-only"
            'MFA data source'                  = $source
            'Last sign-in data'                = if ($signInAvailable) { "Available (no sign-in for $StaleDays+ days is flagged)" } else { 'Unavailable - needs Entra ID P1/P2' }
            'Per-user MFA (legacy)'            = if ($IncludePerUserMfaState) { (@($rows | Group-Object PerUserMfa | ForEach-Object { "$($_.Name): $($_.Count)" }) -join ', ') } else { 'Not checked (use -IncludePerUserMfaState)' }
            'Note'                             = '"Registered" means the user has set up a strong MFA method. Whether MFA is actually required at sign-in depends on security defaults or Conditional Access.'
            'Run by'                           = if ($context.Account) { $context.Account } else { "App: $($context.AppName)" }
        }

        # --- Export ------------------------------------------------------

        $mainHeaders = @('DisplayName', 'UserPrincipalName', 'Email', 'SignInStatus', 'LastSignIn', 'DaysSinceSignIn', 'Created', 'MfaStatus', 'MethodsRegistered',
                         'DefaultMethod', 'PhishingResistant', 'MfaCapable', 'PasswordlessCapable', 'SsprRegistered', 'IsAdmin', 'AdminRoles', 'Risk', 'Issues',
                         'UserType', 'Licensed', 'AccountSource', 'LastInteractiveSignIn', 'LastNonInteractiveSignIn', 'Department', 'JobTitle')
        $mainWidths  = @(26, 34, 32, 13, 18, 15, 18, 15, 50, 26, 17, 12, 19, 15, 9, 40, 10, 50, 10, 9, 13, 21, 24, 20, 24)
        if ($IncludePerUserMfaState) {
            $idx = [array]::IndexOf($mainHeaders, 'MfaCapable')
            $mainHeaders = $mainHeaders[0..$idx] + 'PerUserMfa' + $mainHeaders[($idx + 1)..($mainHeaders.Count - 1)]
            $mainWidths  = $mainWidths[0..$idx] + 12 + $mainWidths[($idx + 1)..($mainWidths.Count - 1)]
        }

        $sheets = @(
            @{ Name = 'MFA Status'; Rows = $rows; Headers = $mainHeaders; Widths = $mainWidths }
            @{ Name = 'Needs Attention'; Rows = $attention
               Headers = @('Risk', 'Issues', 'DisplayName', 'UserPrincipalName', 'SignInStatus', 'LastSignIn', 'Created', 'MfaStatus', 'MethodsRegistered', 'AdminRoles')
               Widths  = @(10, 55, 26, 34, 13, 18, 18, 15, 45, 40) }
            @{ Name = 'Admins'; Rows = $admins
               Headers = @('DisplayName', 'UserPrincipalName', 'AdminRoles', 'SignInStatus', 'MfaStatus', 'MethodsRegistered', 'PhishingResistant', 'LastSignIn', 'Risk', 'Issues')
               Widths  = @(26, 34, 50, 13, 15, 45, 17, 18, 10, 45) }
            @{ Name = 'Methods'; Rows = $methodCounts; Headers = @('Method', 'Users', 'PercentOfEnabled'); Widths = @(32, 10, 17) }
            @{ Name = 'Conditional Access'; Rows = $caRows
               Headers = @('Policy', 'State', 'Requires', 'AppliesTo', 'Exclusions', 'Applications', 'Modified'); Widths = @(45, 12, 40, 35, 11, 15, 18) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(24, 120) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(34, 100) }
        )
        if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

        # --- Console -----------------------------------------------------

        Write-Host ""
        Write-Host ("Enabled users with MFA : {0} of {1} ({2}%)" -f $enabledMfa.Count, $enabled.Count, $pct) -ForegroundColor $(if ($pct -ge 100) { 'Green' } elseif ($pct -ge 90) { 'Yellow' } else { 'Red' })
        Write-Host ("Admins without MFA     : {0}" -f $summary['Admins without MFA']) -ForegroundColor $(if ($summary['Admins without MFA']) { 'Red' } else { 'Green' })
        Write-Host ("Needs attention        : {0} critical, {1} high, {2} medium, {3} low" -f $counts.Critical, $counts.High, $counts.Medium, $counts.Low)
        Write-Host ("Security defaults      : {0};  CA policies requiring MFA: {1} enabled" -f $securityDefaults, $caEnabled)
        if ($securityDefaults -ne 'Enabled' -and $caEnabled -eq 0) {
            Write-Host 'WARNING: neither security defaults nor an enabled Conditional Access policy requires MFA.' -ForegroundColor Red
        }
        if ($errors.Count) { Write-Host ("Some data was unavailable - see the 'Errors' sheet ({0})." -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report                 : $reportPath"

        # --- Email -------------------------------------------------------

        if ($SendEmail) {
            $intro = @(
                "$tenantName - $($enabledMfa.Count) of $($enabled.Count) enabled users have MFA registered ($pct%).",
                "Security defaults: $securityDefaults. Conditional Access policies requiring MFA: $caEnabled enabled."
            )
            if ($securityDefaults -ne 'Enabled' -and $caEnabled -eq 0) { $intro += 'WARNING: nothing in the tenant currently requires MFA at sign-in.' }
            $html = New-HtmlReport -Title 'Microsoft 365 MFA report' -IntroLines $intro -Sections @(
                @{ Heading = "Needs attention ($($attention.Count))"
                   Html    = ConvertTo-HtmlTable -Rows @($attention | Where-Object { $_.Risk -ne 'Low' } | Select-Object -First 100) -Columns @('Risk', 'DisplayName', 'UserPrincipalName', 'MfaStatus', 'LastSignIn', 'Issues') -StatusColumn 'Risk' -EmptyText 'No critical, high or medium issues.' }
                @{ Heading = "Admins ($($admins.Count))"
                   Html    = ConvertTo-HtmlTable -Rows $admins -Columns @('DisplayName', 'AdminRoles', 'MfaStatus', 'MethodsRegistered', 'Risk') -StatusColumn 'Risk' -EmptyText 'No admins found.' }
            )
            $state = if ($counts.Critical) { 'CRITICAL' } elseif ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Medium' } else { 'OK' }
            $subject = "[$state] M365 MFA report - $($enabled.Count - $enabledMfa.Count) enabled user(s) without MFA"
            if (Send-ToolkitMail -Config $config -Subject $subject -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
