<#
.SYNOPSIS
    Audits AD accounts for settings that attackers commonly abuse.
.DESCRIPTION
    Read-only. Rates each finding High, Medium or Low, with a recommended fix.
.EXAMPLE
    .\Get-RiskyAccountReport.ps1
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: ActiveDirectory module (RSAT)
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$Server,
    [switch]$IncludeDisabled,
    [switch]$SendEmail,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('SearchBase')) { $SearchBase = Get-Setting $config 'RiskyAccounts' 'SearchBase' '' }
if (-not $PSBoundParameters.ContainsKey('Server'))     { $Server     = Get-Setting $config 'ActiveDirectory' 'Server' '' }

$privMaxDays   = [int](Get-Setting $config 'RiskyAccounts' 'PrivilegedPasswordMaxDays' 365)
$krbtgtMaxDays = [int](Get-Setting $config 'RiskyAccounts' 'KrbtgtMaxDays' 180)
$accepted      = Get-Setting $config 'RiskyAccounts' 'AcceptedRisks' @{}
if (-not ($accepted -is [hashtable])) { $accepted = @{} }

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { throw 'The ActiveDirectory module is not installed. Install RSAT: Active Directory tools (or run on a domain controller).' }

$ad = @{}
if ($Server) { $ad.Server = $Server }
$search = @{} + $ad
if ($SearchBase) { $search.SearchBase = $SearchBase }

$now        = Get-Date
$reportPath = New-ReportPath -Config $config -Prefix 'RiskyAccounts' -OutputFile $OutputFile
$domain     = Get-ADDomain @ad
$domainSid  = $domain.DomainSID.Value

Write-Host "Domain : $($domain.DNSRoot)"
if ($SearchBase) { Write-Host "Search : $SearchBase" }
Write-Host ""

# --- Privileged accounts --------------------------------------------------

Write-Host 'Finding privileged accounts...'
$privGroups = [ordered]@{
    'Domain Admins'     = "$domainSid-512"
    'Enterprise Admins' = "$domainSid-519"
    'Schema Admins'     = "$domainSid-518"
    'Administrators'    = 'S-1-5-32-544'
    'Account Operators' = 'S-1-5-32-548'
    'Server Operators'  = 'S-1-5-32-549'
    'Backup Operators'  = 'S-1-5-32-551'
    'Print Operators'   = 'S-1-5-32-550'
}
$privileged = @{}   # DistinguishedName -> list of group names
foreach ($name in $privGroups.Keys) {
    try {
        $group = Get-ADGroup -Identity $privGroups[$name] @ad -ErrorAction Stop
    }
    catch { continue }   # Enterprise/Schema Admins only exist in the forest root domain
    # LDAP_MATCHING_RULE_IN_CHAIN: all users in the group, including through nested groups
    $members = @(Get-ADUser -LDAPFilter "(memberOf:1.2.840.113556.1.4.1941:=$($group.DistinguishedName))" @ad)
    foreach ($m in $members) {
        if (-not $privileged.ContainsKey($m.DistinguishedName)) { $privileged[$m.DistinguishedName] = [System.Collections.Generic.List[string]]::new() }
        $privileged[$m.DistinguishedName].Add($group.Name)
    }
}

# --- Users ----------------------------------------------------------------

Write-Host 'Checking user accounts...'
$UF = @{
    PasswordNotRequired   = 0x0020
    ReversibleEncryption  = 0x0080
    PasswordNeverExpires  = 0x10000
    TrustedForDelegation  = 0x80000
    NotDelegated          = 0x100000
    UseDesKeyOnly         = 0x200000
    DontRequirePreauth    = 0x400000
}

$props  = 'userAccountControl', 'servicePrincipalName', 'adminCount', 'PasswordLastSet', 'lastLogonTimestamp',
          'SIDHistory', 'Description', 'DisplayName', 'Enabled', 'objectSid', 'whenCreated'
$filter = if ($IncludeDisabled) { '*' } else { 'Enabled -eq $true' }
$users  = @(Get-ADUser -Filter $filter -Properties $props @search)

$findings = [System.Collections.Generic.List[object]]::new()

function Add-Finding($Account, [string]$Type, [string]$Severity, [string]$Finding, [string]$Detail, [string]$Recommendation) {
    $sam      = $Account.SamAccountName
    $isAccepted = $false
    if ($accepted.ContainsKey($sam)) { $isAccepted = @($accepted[$sam]) -contains $Finding -or @($accepted[$sam]) -contains '*' }
    $groups   = if ($privileged.ContainsKey($Account.DistinguishedName)) { ($privileged[$Account.DistinguishedName] | Sort-Object -Unique) -join ', ' } else { '' }
    $lastLogon = if ($Account.lastLogonTimestamp) { [datetime]::FromFileTime($Account.lastLogonTimestamp) } else { $null }
    $findings.Add([pscustomobject]@{
        Severity        = $Severity
        Status          = if ($isAccepted) { 'Accepted' } else { 'Open' }
        Finding         = $Finding
        ObjectType      = $Type
        SamAccountName  = $sam
        Name            = if ($Account.DisplayName) { $Account.DisplayName } else { $Account.Name }
        Enabled         = [bool]$Account.Enabled
        PrivilegedVia   = $groups
        Detail          = $Detail
        Recommendation  = $Recommendation
        PasswordLastSet = $Account.PasswordLastSet
        LastLogon       = $lastLogon
        OU              = ($Account.DistinguishedName -replace '^.+?(?<!\\),', '')
    })
}

foreach ($u in $users) {
    $uac      = [int]$u.userAccountControl
    $isPriv   = $privileged.ContainsKey($u.DistinguishedName)
    $spns     = @($u.servicePrincipalName | Where-Object { $_ })
    $pwdAge   = if ($u.PasswordLastSet) { [int]($now - $u.PasswordLastSet).TotalDays } else { $null }
    $rid      = ($u.objectSid.Value -split '-')[-1]

    if ($rid -eq '502') { continue }   # krbtgt is checked separately below
    if ($rid -eq '501' -and $u.Enabled) {
        Add-Finding $u 'User' 'High' 'Guest account enabled' 'The built-in Guest account is enabled.' 'Disable the Guest account.'
    }

    if ($uac -band $UF.PasswordNotRequired) {
        Add-Finding $u 'User' 'High' 'Password not required' 'PASSWD_NOTREQD is set, so the account can have an empty password.' `
            'Clear the flag: Set-ADUser <name> -PasswordNotRequired $false, then set a strong password.'
    }
    if ($uac -band $UF.ReversibleEncryption) {
        Add-Finding $u 'User' 'High' 'Reversible password encryption' 'The password is stored in a form that can be decrypted.' `
            'Turn off "Store password using reversible encryption" and reset the password.'
    }
    if ($uac -band $UF.DontRequirePreauth) {
        Add-Finding $u 'User' 'High' 'Kerberos pre-auth not required' 'Anyone can request data encrypted with this password and crack it offline (AS-REP roasting).' `
            'Turn off "Do not require Kerberos preauthentication" unless an application truly needs it.'
    }
    if ($uac -band $UF.TrustedForDelegation) {
        Add-Finding $u 'User' 'High' 'Unconstrained delegation' 'This account can impersonate any user that authenticates to it, anywhere.' `
            'Switch to constrained or resource-based constrained delegation.'
    }
    if ($spns.Count) {
        $spnText = ($spns | Select-Object -First 3) -join '; '
        if ($spns.Count -gt 3) { $spnText += " (+$($spns.Count - 3) more)" }
        if ($isPriv) {
            Add-Finding $u 'User' 'High' 'Kerberoastable privileged account' "Privileged account with SPN(s): $spnText. Its password hash can be requested by any user and cracked offline." `
                'Remove the SPN from the privileged account, or move the service to a gMSA. Otherwise use a 25+ character random password.'
        } else {
            Add-Finding $u 'User' 'Medium' 'Kerberoastable account' "Has SPN(s): $spnText. Any domain user can request a ticket and try to crack the password offline." `
                'Use a group Managed Service Account (gMSA), or a 25+ character random password, rotated regularly.'
        }
    }
    if ($uac -band $UF.UseDesKeyOnly) {
        Add-Finding $u 'User' 'Medium' 'DES-only Kerberos encryption' '"Use only Kerberos DES encryption types" is set; DES is broken.' 'Clear the flag and reset the password.'
    }
    if (@($u.SIDHistory).Count -and $u.SIDHistory) {
        Add-Finding $u 'User' 'Medium' 'SID history present' "$(@($u.SIDHistory).Count) SID(s) in sIDHistory - grants the rights of old accounts." `
            'Remove SID history once migrations are finished.'
    }

    if ($uac -band $UF.PasswordNeverExpires) {
        if ($isPriv) {
            Add-Finding $u 'User' 'Medium' 'Password never expires (privileged)' 'Privileged account whose password never has to change.' `
                'Require password changes, or use a vault/rotation process for this account.'
        } else {
            Add-Finding $u 'User' 'Low' 'Password never expires' 'The password never has to change.' `
                'Turn off "Password never expires", or use a gMSA for service accounts.'
        }
    }

    if ($isPriv) {
        if ($null -ne $pwdAge -and $pwdAge -gt $privMaxDays) {
            Add-Finding $u 'User' 'Medium' 'Stale privileged password' "Password last changed $pwdAge days ago." 'Change the password.'
        }
        if (-not ($uac -band $UF.NotDelegated)) {
            Add-Finding $u 'User' 'Low' 'Privileged account can be delegated' 'Not marked "Account is sensitive and cannot be delegated".' `
                'Tick "Account is sensitive and cannot be delegated", or add the account to Protected Users (test first).'
        }
    }
    elseif ($u.adminCount -eq 1) {
        Add-Finding $u 'User' 'Low' 'Orphaned adminCount' 'adminCount=1 but the account is no longer in a privileged group, so it keeps locked-down permissions.' `
            'Clear adminCount and re-enable permission inheritance on the account.'
    }
}

# --- krbtgt (always disabled, so checked on its own) ----------------------

try {
    $krbtgt = Get-ADUser -Identity "$domainSid-502" -Properties $props @ad -ErrorAction Stop
    $age = if ($krbtgt.PasswordLastSet) { [int]($now - $krbtgt.PasswordLastSet).TotalDays } else { $null }
    if ($null -ne $age -and $age -gt $krbtgtMaxDays) {
        Add-Finding $krbtgt 'User' 'Medium' 'krbtgt password old' "Last changed $age days ago (limit $krbtgtMaxDays)." `
            "Reset the krbtgt password twice, allowing replication in between, using Microsoft's documented procedure."
    }
} catch { Write-Warning "Could not read the krbtgt account: $($_.Exception.Message)" }

# --- Computers with unconstrained delegation ------------------------------

Write-Host 'Checking computers for unconstrained delegation...'
$computers = @(Get-ADComputer -Filter 'TrustedForDelegation -eq $true' -Properties 'TrustedForDelegation', 'PrimaryGroupID', 'OperatingSystem', 'lastLogonTimestamp', 'PasswordLastSet', 'Enabled' @search)
foreach ($c in $computers) {
    if ($c.PrimaryGroupID -in 516, 521) { continue }   # domain controllers / read-only DCs
    if (-not $IncludeDisabled -and -not $c.Enabled) { continue }
    Add-Finding $c 'Computer' 'High' 'Unconstrained delegation' "$($c.OperatingSystem) - can impersonate any user that connects to it." `
        'Switch to constrained or resource-based constrained delegation.'
}

# --- Export ---------------------------------------------------------------

$sevOrder = @{ High = 0; Medium = 1; Low = 2 }
$sorted   = @($findings | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } }, Finding, SamAccountName)
$open     = @($sorted | Where-Object { $_.Status -eq 'Open' })

# One row per flagged account
$accounts = @($sorted | Group-Object ObjectType, SamAccountName | ForEach-Object {
    $first = $_.Group[0]
    $worst = ($_.Group | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } } | Select-Object -First 1).Severity
    [pscustomobject]@{
        HighestSeverity = $worst
        SamAccountName  = $first.SamAccountName
        ObjectType      = $first.ObjectType
        Name            = $first.Name
        Enabled         = $first.Enabled
        PrivilegedVia   = $first.PrivilegedVia
        OpenFindings    = @($_.Group | Where-Object { $_.Status -eq 'Open' }).Count
        Findings        = ($_.Group | ForEach-Object { $_.Finding + $(if ($_.Status -eq 'Accepted') { ' (accepted)' } else { '' }) }) -join '; '
        PasswordLastSet = $first.PasswordLastSet
        LastLogon       = $first.LastLogon
        OU              = $first.OU
    }
} | Sort-Object @{ Expression = { $sevOrder[$_.HighestSeverity] } }, SamAccountName)

# All privileged users, for review even when nothing is flagged
$userByDn = @{}
foreach ($u in $users) { $userByDn[$u.DistinguishedName] = $u }
$privRows = @(foreach ($dn in $privileged.Keys) {
    $u = $userByDn[$dn]
    if (-not $u) { try { $u = Get-ADUser -Identity $dn -Properties $props @ad } catch { continue } }
    [pscustomobject]@{
        SamAccountName  = $u.SamAccountName
        Name            = if ($u.DisplayName) { $u.DisplayName } else { $u.Name }
        Enabled         = [bool]$u.Enabled
        Groups          = ($privileged[$dn] | Sort-Object -Unique) -join ', '
        PasswordLastSet = $u.PasswordLastSet
        PasswordAgeDays = if ($u.PasswordLastSet) { [int]($now - $u.PasswordLastSet).TotalDays } else { $null }
        LastLogon       = if ($u.lastLogonTimestamp) { [datetime]::FromFileTime($u.lastLogonTimestamp) } else { $null }
        OU              = ($dn -replace '^.+?(?<!\\),', '')
    }
}) | Sort-Object SamAccountName

$counts = [ordered]@{}
foreach ($s in 'High', 'Medium', 'Low') { $counts[$s] = @($open | Where-Object { $_.Severity -eq $s }).Count }

$summaryItems = [ordered]@{
    'Run at'              = $now.ToString('yyyy-MM-dd HH:mm:ss')
    'Domain'              = $domain.DNSRoot
    'Search base'         = if ($SearchBase) { $SearchBase } else { '(whole domain)' }
    'Accounts checked'    = "$($users.Count) user(s)" + $(if ($IncludeDisabled) { ' incl. disabled' } else { ' (enabled only)' })
    'Privileged users'    = $privileged.Count
    'Open findings'       = "$($open.Count) ($($counts.High) high, $($counts.Medium) medium, $($counts.Low) low)"
    'Accepted findings'   = @($sorted | Where-Object { $_.Status -eq 'Accepted' }).Count
}
foreach ($f in ($open | Group-Object Finding | Sort-Object Count -Descending)) { $summaryItems["Finding: $($f.Name)"] = $f.Count }
$summaryItems['Run by'] = "$env:USERDOMAIN\$env:USERNAME"

$saved = Save-Workbook -FilePath $reportPath -Sheets @(
    @{ Name = 'Findings'; Rows = $sorted
       Headers = @('Severity', 'Status', 'Finding', 'ObjectType', 'SamAccountName', 'Name', 'Enabled', 'PrivilegedVia', 'Detail', 'Recommendation', 'PasswordLastSet', 'LastLogon', 'OU')
       Widths  = @(10, 10, 34, 11, 22, 26, 9, 30, 60, 60, 18, 18, 50) }
    @{ Name = 'Accounts'; Rows = $accounts
       Headers = @('HighestSeverity', 'SamAccountName', 'ObjectType', 'Name', 'Enabled', 'PrivilegedVia', 'OpenFindings', 'Findings', 'PasswordLastSet', 'LastLogon', 'OU')
       Widths  = @(15, 22, 11, 26, 9, 30, 13, 80, 18, 18, 50) }
    @{ Name = 'Privileged Accounts'; Rows = $privRows
       Headers = @('SamAccountName', 'Name', 'Enabled', 'Groups', 'PasswordLastSet', 'PasswordAgeDays', 'LastLogon', 'OU')
       Widths  = @(22, 26, 9, 45, 18, 16, 18, 50) }
    @{ Name = 'Summary'; Rows = (New-SummaryRows $summaryItems); Headers = @('Item', 'Value'); Widths = @(44, 70) }
)
if (-not $saved) { exit 1 }

Write-Host ""
Write-Host ("Open findings: {0} high, {1} medium, {2} low" -f $counts.High, $counts.Medium, $counts.Low) -ForegroundColor $(if ($counts.High) { 'Red' } elseif ($counts.Medium) { 'Yellow' } else { 'Green' })
foreach ($f in ($open | Group-Object Finding | Sort-Object Count -Descending)) { Write-Host ("  {0,-38} {1}" -f $f.Name, $f.Count) }
Write-Host "Privileged users: $($privileged.Count)"
Write-Host "Report : $reportPath"

if ($SendEmail) {
    $intro = @("Domain $($domain.DNSRoot): $($counts.High) high, $($counts.Medium) medium and $($counts.Low) low open findings across $(@($accounts | Where-Object { $_.OpenFindings }).Count) account(s).",
               'Accepted risks are not listed here. The attached workbook has full details and recommendations.')
    $html = New-HtmlReport -Title 'Risky AD account settings' -IntroLines $intro -Sections @(
        @{ Heading = 'High and medium findings'; Html = ConvertTo-HtmlTable -Rows @($open | Where-Object { $_.Severity -ne 'Low' } | Select-Object -First 100) -Columns @('Severity', 'Finding', 'SamAccountName', 'PrivilegedVia', 'Recommendation') -StatusColumn 'Severity' -EmptyText 'None.' }
        @{ Heading = 'Low findings (count)'; Html = ConvertTo-HtmlTable -Rows @($open | Where-Object { $_.Severity -eq 'Low' } | Group-Object Finding | ForEach-Object { [pscustomobject]@{ Finding = $_.Name; Accounts = $_.Count } }) -Columns @('Finding', 'Accounts') -EmptyText 'None.' }
    )
    $state = if ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Medium' } else { 'Low' }
    if (Send-ToolkitMail -Config $config -Subject "[$state] Risky AD accounts - $($open.Count) open finding(s)" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
}
