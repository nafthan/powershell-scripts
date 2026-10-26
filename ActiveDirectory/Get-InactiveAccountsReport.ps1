<#
.SYNOPSIS
    Finds AD users and computers that haven't logged on for N days.
.DESCRIPTION
    Report only, unless -DisableAccounts is used. Then it disables the accounts and moves them, asking before each one.
.EXAMPLE
    .\Get-InactiveAccountsReport.ps1
.EXAMPLE
    .\Get-InactiveAccountsReport.ps1 -DisableAccounts -WhatIf
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
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [ValidateRange(1, 3650)]
    [int]$InactiveDays,

    [ValidateSet('Users', 'Computers', 'Both')]
    [string]$AccountType = 'Both',

    [string]$SearchBase,
    [string]$Server,
    [string[]]$Exclude,

    [switch]$DisableAccounts,
    [string]$MoveToOU,

    [switch]$SendEmail,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('InactiveDays')) { $InactiveDays = Get-Setting $config 'ActiveDirectory' 'InactiveDays' 90 }
if (-not $PSBoundParameters.ContainsKey('SearchBase'))   { $SearchBase   = Get-Setting $config 'ActiveDirectory' 'SearchBase' '' }
if (-not $PSBoundParameters.ContainsKey('Server'))       { $Server       = Get-Setting $config 'ActiveDirectory' 'Server' '' }
if (-not $PSBoundParameters.ContainsKey('Exclude'))      { $Exclude      = Get-Setting $config 'ActiveDirectory' 'Exclude' @() }
if (-not $PSBoundParameters.ContainsKey('MoveToOU') -and $DisableAccounts) { $MoveToOU = Get-Setting $config 'ActiveDirectory' 'DisabledOU' '' }

$Exclude = Split-ListValue $Exclude

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { throw 'The ActiveDirectory module is not installed. Install RSAT: Active Directory tools (or run on a domain controller).' }

$ad = @{}
if ($Server) { $ad.Server = $Server }
$search = @{} + $ad
if ($SearchBase) { $search.SearchBase = $SearchBase }

$now    = Get-Date
$cutoff = $now.AddDays(-$InactiveDays)
$reportPath = New-ReportPath -Config $config -Prefix 'InactiveAccounts' -OutputFile $OutputFile

Write-Host ("Looking for enabled accounts with no logon since {0:yyyy-MM-dd} ({1} days)" -f $cutoff, $InactiveDays)
if ($SearchBase) { Write-Host "Search base : $SearchBase" }
Write-Host ""

function Test-Excluded([string]$Name) {
    foreach ($pattern in $Exclude) { if ($Name -like $pattern) { return $true } }
    return $false
}

function Get-ParentOU([string]$DistinguishedName) {
    return ($DistinguishedName -replace '^.+?(?<!\\),', '')
}

function Get-InactiveInfo($Account) {
    $lastLogon = if ($Account.lastLogonTimestamp) { [datetime]::FromFileTime($Account.lastLogonTimestamp) } else { $null }
    $inactive  = if ($lastLogon) { $lastLogon -lt $cutoff } else { $Account.whenCreated -lt $cutoff }
    $days      = if ($lastLogon) { [int][math]::Floor(($now - $lastLogon).TotalDays) } else { $null }
    return @{ LastLogon = $lastLogon; Inactive = $inactive; Days = $days }
}

# --- Users ----------------------------------------------------------------

$users = @()
if ($AccountType -in 'Users', 'Both') {
    Write-Host 'Querying users...'
    $props = 'lastLogonTimestamp', 'whenCreated', 'PasswordLastSet', 'Description', 'UserPrincipalName', 'EmailAddress', 'DisplayName'
    $users = @(Get-ADUser -Filter 'Enabled -eq $true' -Properties $props @search | ForEach-Object {
        if (Test-Excluded $_.SamAccountName) { return }
        $info = Get-InactiveInfo $_
        if (-not $info.Inactive) { return }
        [pscustomobject]@{
            SamAccountName    = $_.SamAccountName
            DisplayName       = $_.DisplayName
            UserPrincipalName = $_.UserPrincipalName
            Email             = $_.EmailAddress
            LastLogon         = if ($info.LastLogon) { $info.LastLogon } else { 'Never' }
            DaysInactive      = $info.Days
            Created           = $_.whenCreated
            PasswordLastSet   = $_.PasswordLastSet
            Description       = $_.Description
            OU                = Get-ParentOU $_.DistinguishedName
            DistinguishedName = $_.DistinguishedName
            ObjectType        = 'User'
        }
    } | Sort-Object -Property @{ Expression = { if ($null -eq $_.DaysInactive) { [int]::MaxValue } else { $_.DaysInactive } }; Descending = $true })
}

# --- Computers ------------------------------------------------------------

$computers = @()
if ($AccountType -in 'Computers', 'Both') {
    Write-Host 'Querying computers...'
    $props = 'lastLogonTimestamp', 'whenCreated', 'OperatingSystem', 'OperatingSystemVersion', 'Description', 'DNSHostName'
    $computers = @(Get-ADComputer -Filter 'Enabled -eq $true' -Properties $props @search | ForEach-Object {
        if (Test-Excluded $_.SamAccountName) { return }
        # Never touch domain controllers
        if ($_.DistinguishedName -match 'OU=Domain Controllers,') { return }
        $info = Get-InactiveInfo $_
        if (-not $info.Inactive) { return }
        [pscustomobject]@{
            Name              = $_.Name
            DNSHostName       = $_.DNSHostName
            OperatingSystem   = $_.OperatingSystem
            OSVersion         = $_.OperatingSystemVersion
            LastLogon         = if ($info.LastLogon) { $info.LastLogon } else { 'Never' }
            DaysInactive      = $info.Days
            Created           = $_.whenCreated
            Description       = $_.Description
            OU                = Get-ParentOU $_.DistinguishedName
            DistinguishedName = $_.DistinguishedName
            SamAccountName    = $_.SamAccountName
            ObjectType        = 'Computer'
        }
    } | Sort-Object -Property @{ Expression = { if ($null -eq $_.DaysInactive) { [int]::MaxValue } else { $_.DaysInactive } }; Descending = $true })
}

# --- Optional: disable and move -------------------------------------------

$actions = [System.Collections.Generic.List[object]]::new()
if ($DisableAccounts) {
    $targets = @($users) + @($computers)
    $stamp   = $now.ToString('yyyy-MM-dd')
    $verb    = if ($MoveToOU) { "Disable and move to $MoveToOU" } else { 'Disable' }

    foreach ($acct in $targets) {
        $label  = "$($acct.ObjectType) $($acct.SamAccountName)"
        $result = ''
        if ($PSCmdlet.ShouldProcess($label, $verb)) {
            try {
                # -Confirm:$false: the confirmation was already given above
                Disable-ADAccount -Identity $acct.DistinguishedName @ad -Confirm:$false -ErrorAction Stop
                $result = 'Disabled'
                $note = "Disabled $stamp - no logon for $InactiveDays+ days."
                if ($acct.Description) { $note += " Was: $($acct.Description)" }
                if ($note.Length -gt 1024) { $note = $note.Substring(0, 1024) }
                Set-ADObject -Identity $acct.DistinguishedName -Replace @{ description = $note } @ad -Confirm:$false -ErrorAction Stop
                if ($MoveToOU) {
                    Move-ADObject -Identity $acct.DistinguishedName -TargetPath $MoveToOU @ad -Confirm:$false -ErrorAction Stop
                    $result = 'Disabled and moved'
                }
            }
            catch {
                $msg = $_.Exception.Message
                $result = if ($result) { "${result}, then failed: $msg" } else { "Failed: $msg" }
            }
        }
        elseif ($WhatIfPreference) { $result = 'WhatIf - no change made' }
        else                       { $result = 'Skipped (not confirmed)' }

        $actions.Add([pscustomobject]@{
            ObjectType        = $acct.ObjectType
            SamAccountName    = $acct.SamAccountName
            DaysInactive      = $acct.DaysInactive
            Action            = $verb
            Result            = $result
            OriginalOU        = $acct.OU
            DistinguishedName = $acct.DistinguishedName
        })
    }
}

# --- Export ---------------------------------------------------------------

$summary = [ordered]@{
    'Run at'                = $now.ToString('yyyy-MM-dd HH:mm:ss')
    'Inactive threshold'    = "$InactiveDays days (no logon since $($cutoff.ToString('yyyy-MM-dd')))"
    'Search base'           = if ($SearchBase) { $SearchBase } else { '(whole domain)' }
    'Excluded patterns'     = ($Exclude -join ', ')
    'Inactive users'        = $users.Count
    'Inactive computers'    = $computers.Count
    'Mode'                  = if (-not $DisableAccounts) { 'Report only' } elseif ($WhatIfPreference) { 'WhatIf (no changes)' } else { 'Disable' + $(if ($MoveToOU) { ' and move' } else { '' }) }
    'Accounts changed'      = @($actions | Where-Object { $_.Result -like 'Disabled*' }).Count
    'Note'                  = 'Last logon comes from lastLogonTimestamp, which can lag the real last logon by up to 14 days.'
    'Run by'                = "$env:USERDOMAIN\$env:USERNAME"
}

$sheets = @(
    @{ Name = 'Inactive Users'; Rows = $users
       Headers = @('SamAccountName', 'DisplayName', 'UserPrincipalName', 'Email', 'LastLogon', 'DaysInactive', 'Created', 'PasswordLastSet', 'Description', 'OU')
       Widths  = @(22, 28, 34, 32, 18, 13, 18, 18, 40, 60) }
    @{ Name = 'Inactive Computers'; Rows = $computers
       Headers = @('Name', 'DNSHostName', 'OperatingSystem', 'OSVersion', 'LastLogon', 'DaysInactive', 'Created', 'Description', 'OU')
       Widths  = @(20, 34, 32, 16, 18, 13, 18, 40, 60) }
)
if ($DisableAccounts) {
    $sheets += @{ Name = 'Actions'; Rows = $actions
                  Headers = @('ObjectType', 'SamAccountName', 'DaysInactive', 'Action', 'Result', 'OriginalOU', 'DistinguishedName')
                  Widths  = @(11, 22, 13, 40, 40, 60, 80) }
}
$sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(24, 90) }

if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

Write-Host ""
Write-Host ("Inactive users     : {0:N0}" -f $users.Count)     -ForegroundColor $(if ($users.Count) { 'Yellow' } else { 'Green' })
Write-Host ("Inactive computers : {0:N0}" -f $computers.Count) -ForegroundColor $(if ($computers.Count) { 'Yellow' } else { 'Green' })
if ($DisableAccounts) {
    $actions | Group-Object Result | ForEach-Object { Write-Host ("  {0,-40} {1:N0}" -f $_.Name, $_.Count) }
}
Write-Host "Report             : $reportPath"

# --- Email ----------------------------------------------------------------

if ($SendEmail) {
    $sections = @(
        @{ Heading = "Inactive users ($($users.Count))"
           Html    = ConvertTo-HtmlTable -Rows @($users | Select-Object -First 50) -Columns @('SamAccountName', 'DisplayName', 'LastLogon', 'DaysInactive', 'OU') -EmptyText 'No inactive users.' }
        @{ Heading = "Inactive computers ($($computers.Count))"
           Html    = ConvertTo-HtmlTable -Rows @($computers | Select-Object -First 50) -Columns @('Name', 'OperatingSystem', 'LastLogon', 'DaysInactive', 'OU') -EmptyText 'No inactive computers.' }
    )
    if ($DisableAccounts) {
        $sections += @{ Heading = 'Actions taken'; Html = ConvertTo-HtmlTable -Rows $actions -Columns @('ObjectType', 'SamAccountName', 'Result') }
    }
    $intro = @("Enabled accounts with no logon for $InactiveDays days or more.", 'Tables show up to 50 rows each - the attached workbook has the full list.')
    $html  = New-HtmlReport -Title 'Inactive AD accounts' -IntroLines $intro -Sections $sections
    $subject = "Inactive AD accounts: $($users.Count) users, $($computers.Count) computers"
    if (Send-ToolkitMail -Config $config -Subject $subject -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
}
