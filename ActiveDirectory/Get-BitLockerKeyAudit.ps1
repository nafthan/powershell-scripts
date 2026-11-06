<#
.SYNOPSIS
    Finds Active Directory computers with no BitLocker recovery key backed up to AD.
.DESCRIPTION
    Matches every computer account with its msFVE-RecoveryInformation objects (the keys themselves are
    never read) and reports coverage. Read-only.
.EXAMPLE
    .\Get-BitLockerKeyAudit.ps1
.EXAMPLE
    .\Get-BitLockerKeyAudit.ps1 -SearchBase 'OU=Workstations,DC=corp,DC=example,DC=com' -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: ActiveDirectory module (RSAT); an account that can see BitLocker recovery objects (e.g. Domain Admins)
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$Server,
    [switch]$IncludeDisabled,
    [switch]$IncludeServers,
    [ValidateRange(1, 3650)] [int]$StaleDays,

    [switch]$SendEmail,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('SearchBase')) { $SearchBase = Get-Setting $config 'BitLocker' 'SearchBase' '' }
if (-not $PSBoundParameters.ContainsKey('Server'))     { $Server     = Get-Setting $config 'ActiveDirectory' 'Server' '' }
if (-not $PSBoundParameters.ContainsKey('StaleDays'))  { $StaleDays  = [int](Get-Setting $config 'BitLocker' 'StaleDays' 90) }

if ($MyInvocation.InvocationName -ne '.') {
    try { Import-Module ActiveDirectory -ErrorAction Stop }
    catch { throw 'The ActiveDirectory module is not installed. Install RSAT: Active Directory tools.' }

    $ad = @{}
    if ($Server) { $ad.Server = $Server }
    if (-not $SearchBase) { $SearchBase = (Get-ADDomain @ad).DistinguishedName }
    $now        = Get-Date
    $reportPath = New-ReportPath -Config $config -Prefix 'BitLockerKeys' -OutputFile $OutputFile

    Write-Host "Search base: $SearchBase"
    Write-Host 'Getting computers...'
    $filter = if ($IncludeDisabled) { '*' } else { 'Enabled -eq $true' }
    $computers = @(Get-ADComputer -Filter $filter -SearchBase $SearchBase -Properties OperatingSystem, LastLogonDate, whenCreated, Enabled @ad |
                   Where-Object { "$($_.OperatingSystem)" -like 'Windows*' -and ($IncludeServers -or "$($_.OperatingSystem)" -notlike '*Server*') })

    # One query for every recovery object; they sit under the computer they belong to
    Write-Host 'Getting BitLocker recovery objects...'
    $keysByComputer = @{}
    foreach ($k in @(Get-ADObject -LDAPFilter '(objectClass=msFVE-RecoveryInformation)' -SearchBase $SearchBase -Properties whenCreated @ad)) {
        $parent = $k.DistinguishedName -replace '^.+?(?<!\\),', ''
        if (-not $keysByComputer.ContainsKey($parent)) { $keysByComputer[$parent] = [System.Collections.Generic.List[datetime]]::new() }
        $keysByComputer[$parent].Add($k.whenCreated)
    }

    $rows = foreach ($c in $computers) {
        $keys     = $keysByComputer[$c.DistinguishedName]
        $count    = if ($keys) { $keys.Count } else { 0 }
        $latest   = if ($count) { ($keys | Sort-Object -Descending | Select-Object -First 1) } else { $null }
        $lastSeen = if ($c.LastLogonDate) { [int]($now - $c.LastLogonDate).TotalDays } else { $null }
        $stale    = ($null -eq $lastSeen -and ($now - $c.whenCreated).TotalDays -gt $StaleDays) -or ($lastSeen -gt $StaleDays)

        $status = if ($count) { 'Key backed up' } elseif ($stale) { 'No key (stale computer)' } else { 'No key' }
        [pscustomobject]@{
            Status          = $status
            Computer        = $c.Name
            OperatingSystem = $c.OperatingSystem
            Enabled         = [bool]$c.Enabled
            LastLogon       = $c.LastLogonDate
            DaysSinceLogon  = $lastSeen
            Keys            = $count
            LatestKey       = $latest
            OU              = ($c.DistinguishedName -replace '^.+?(?<!\\),', '')
        }
    }
    $rows = @($rows)

    $order    = @{ 'No key' = 0; 'No key (stale computer)' = 1; 'Key backed up' = 2 }
    $sorted   = @($rows | Sort-Object @{ Expression = { $order[$_.Status] } }, Computer)
    $missing  = @($rows | Where-Object { $_.Status -eq 'No key' })
    $active   = @($rows | Where-Object { $_.Status -ne 'No key (stale computer)' })
    $coverage = if ($active.Count) { [math]::Round(@($active | Where-Object { $_.Keys -gt 0 }).Count / $active.Count * 100, 1) } else { 0 }

    $summary = [ordered]@{
        'Run at'                       = $now.ToString('yyyy-MM-dd HH:mm:ss')
        'Search base'                  = $SearchBase
        'Computers checked'            = "$($rows.Count) Windows $(if ($IncludeServers) { 'computers' } else { 'clients (servers excluded - use -IncludeServers)' })"
        'Key backed up'                = @($rows | Where-Object Keys -gt 0).Count
        'No key (active)'              = $missing.Count
        "No key (no logon for $StaleDays+ days)" = @($rows | Where-Object Status -eq 'No key (stale computer)').Count
        'Coverage of active computers' = "$coverage%"
        'Note'                         = 'Devices joined only to Entra ID (not AD) back up their keys to Entra ID and are not covered here. A computer may be encrypted without a backed-up key, or have a key from an old installation.'
    }

    $saved = Save-Workbook -FilePath $reportPath -Sheets @(
        @{ Name = 'Computers'; Rows = $sorted; Headers = @('Status', 'Computer', 'OperatingSystem', 'Enabled', 'LastLogon', 'DaysSinceLogon', 'Keys', 'LatestKey', 'OU'); Widths = @(24, 22, 34, 9, 18, 15, 7, 18, 60) }
        @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(36, 100) }
    )
    if (-not $saved) { exit 1 }

    Write-Host ""
    Write-Host ("Coverage: {0}% of active computers have a key backed up. {1} active computer(s) have none." -f $coverage, $missing.Count) -ForegroundColor $(if ($missing.Count) { 'Yellow' } else { 'Green' })
    $missing | Select-Object -First 15 | ForEach-Object { Write-Host ("  {0,-20} {1}" -f $_.Computer, $_.OperatingSystem) }
    Write-Host "Report  : $reportPath"

    if ($SendEmail) {
        $html = New-HtmlReport -Title 'BitLocker recovery key audit' -IntroLines @("$coverage% of active Windows computers have a recovery key backed up to AD. $($missing.Count) have none.") -Sections @(
            @{ Heading = "Active computers with no key ($($missing.Count))"; Html = ConvertTo-HtmlTable -Rows $missing -Columns @('Computer', 'OperatingSystem', 'LastLogon', 'OU') -EmptyText 'None - every active computer has a key backed up.' }
        )
        if (Send-ToolkitMail -Config $config -Subject "BitLocker keys: $coverage% coverage, $($missing.Count) computer(s) without a key" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
    }
}
