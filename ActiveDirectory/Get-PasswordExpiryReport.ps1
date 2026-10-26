<#
.SYNOPSIS
    Lists AD passwords that are expiring soon and can email users a reminder.
.DESCRIPTION
    Reminders are only sent on the days set in PasswordExpiry.NotifyOnDays.
.EXAMPLE
    .\Get-PasswordExpiryReport.ps1 -WarningDays 21
.EXAMPLE
    .\Get-PasswordExpiryReport.ps1 -NotifyUsers -SendReport
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
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateRange(0, 365)]
    [int]$WarningDays,

    [switch]$SkipExpired,

    [string]$SearchBase,
    [string]$Server,
    [string[]]$Exclude,

    [switch]$NotifyUsers,
    [switch]$NotifyAll,
    [string]$TestRecipient,

    [switch]$SendReport,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('WarningDays')) { $WarningDays = Get-Setting $config 'PasswordExpiry' 'WarningDays' 14 }
if (-not $PSBoundParameters.ContainsKey('SearchBase'))  { $SearchBase  = Get-Setting $config 'ActiveDirectory' 'SearchBase' '' }
if (-not $PSBoundParameters.ContainsKey('Server'))      { $Server      = Get-Setting $config 'ActiveDirectory' 'Server' '' }
if (-not $PSBoundParameters.ContainsKey('Exclude'))     { $Exclude     = Get-Setting $config 'ActiveDirectory' 'Exclude' @() }

$Exclude      = Split-ListValue $Exclude
$notifyOnDays = @(Get-Setting $config 'PasswordExpiry' 'NotifyOnDays' @(14, 7, 3, 2, 1, 0))

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { throw 'The ActiveDirectory module is not installed. Install RSAT: Active Directory tools (or run on a domain controller).' }

$search = @{}
if ($Server)     { $search.Server = $Server }
if ($SearchBase) { $search.SearchBase = $SearchBase }

$now        = Get-Date
$today      = $now.Date
$reportPath = New-ReportPath -Config $config -Prefix 'PasswordExpiry' -OutputFile $OutputFile

Write-Host "Checking passwords that expire within $WarningDays day(s)..."

# --- Query ----------------------------------------------------------------

$props = 'msDS-UserPasswordExpiryTimeComputed', 'PasswordLastSet', 'EmailAddress', 'DisplayName', 'GivenName', 'UserPrincipalName'
$filter = 'Enabled -eq $true -and PasswordNeverExpires -eq $false'

$users = @(Get-ADUser -Filter $filter -Properties $props @search | ForEach-Object {
    $u = $_
    foreach ($pattern in $Exclude) { if ($u.SamAccountName -like $pattern) { return } }
    if (-not $u.PasswordLastSet) { return }                      # must change at next logon

    $raw = $u.'msDS-UserPasswordExpiryTimeComputed'
    if (-not $raw -or $raw -eq 0 -or $raw -ge 9223372036854775807) { return }   # never expires

    $expires  = [datetime]::FromFileTime($raw)
    $daysLeft = [int][math]::Floor(($expires.Date - $today).TotalDays)
    if ($daysLeft -gt $WarningDays) { return }
    if ($expires -lt $now -and $SkipExpired) { return }

    [pscustomobject]@{
        SamAccountName  = $u.SamAccountName
        DisplayName     = $u.DisplayName
        GivenName       = $u.GivenName
        Email           = $u.EmailAddress
        PasswordLastSet = $u.PasswordLastSet
        ExpiresOn       = $expires
        DaysLeft        = $daysLeft
        Status          = if ($expires -lt $now) { 'Expired' } elseif ($daysLeft -le 3) { 'Expiring - urgent' } else { 'Expiring' }
        Notification    = ''
        OU              = ($u.DistinguishedName -replace '^.+?(?<!\\),', '')
    }
} | Sort-Object -Property ExpiresOn)

$expired  = @($users | Where-Object { $_.Status -eq 'Expired' })
$expiring = @($users | Where-Object { $_.Status -ne 'Expired' })

# --- Notify users ---------------------------------------------------------

$notifications = [System.Collections.Generic.List[object]]::new()

function New-UserNotificationHtml($User) {
    $helpdesk     = Get-Setting $config 'PasswordExpiry' 'HelpdeskContact' 'the IT Service Desk'
    $instructions = Get-Setting $config 'PasswordExpiry' 'ChangeInstructions' 'Press Ctrl+Alt+Delete on a domain-joined PC and choose "Change a password".'
    $enc  = { param($t) [System.Net.WebUtility]::HtmlEncode($t) }
    $name = if ($User.GivenName) { $User.GivenName } elseif ($User.DisplayName) { $User.DisplayName } else { $User.SamAccountName }
    $when = switch ($User.DaysLeft) {
        { $_ -le 0 } { 'today'; break }
        1            { 'tomorrow'; break }
        default      { "in $($User.DaysLeft) days" }
    }
    $font = 'font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1f2328;'
    return @"
<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="$font">
<p style="$font">Hello $(& $enc $name),</p>
<p style="$font">The password for your account <b>$(& $enc $User.SamAccountName)</b> expires <b>$when</b> ($($User.ExpiresOn.ToString('dddd d MMMM yyyy, HH:mm'))).</p>
<p style="$font">To change it: $(& $enc $instructions)</p>
<p style="$font">If you are working remotely, connect to the VPN before changing your password so your laptop picks up the new one.</p>
<p style="$font">If you need help, please contact $(& $enc $helpdesk).</p>
<p style="$font">This is an automated message - please do not reply.</p>
</body></html>
"@
}

if ($NotifyUsers) {
    foreach ($u in $expiring) {
        if (-not $NotifyAll -and $notifyOnDays -notcontains $u.DaysLeft) {
            $u.Notification = "Not due (notifies at $($notifyOnDays -join ', ') days left)"
            continue
        }
        if (-not $u.Email) { $u.Notification = 'No email address in AD'; continue }

        $to      = if ($TestRecipient) { $TestRecipient } else { $u.Email }
        $subject = if ($u.DaysLeft -le 0) { 'Your password expires today' } elseif ($u.DaysLeft -eq 1) { 'Your password expires tomorrow' } else { "Your password expires in $($u.DaysLeft) days" }

        if ($PSCmdlet.ShouldProcess($to, "Send password expiry reminder for $($u.SamAccountName)")) {
            $sent = Send-ToolkitMail -Config $config -To $to -Subject $subject -HtmlBody (New-UserNotificationHtml $u)
            $u.Notification = if ($sent) { "Sent to $to" } else { 'Send failed - see console warning' }
        }
        else {
            $u.Notification = "WhatIf - would send to $to"
        }
        $notifications.Add([pscustomobject]@{ SamAccountName = $u.SamAccountName; DaysLeft = $u.DaysLeft; SentTo = $to; Result = $u.Notification })
    }
    foreach ($u in $expired) { $u.Notification = 'Already expired - not emailed' }
}

# --- Export ---------------------------------------------------------------

$summary = [ordered]@{
    'Run at'                   = $now.ToString('yyyy-MM-dd HH:mm:ss')
    'Warning window'           = "$WarningDays days"
    'Search base'              = if ($SearchBase) { $SearchBase } else { '(whole domain)' }
    'Expiring within window'   = $expiring.Count
    'Already expired'          = if ($SkipExpired) { 'Not included (-SkipExpired)' } else { $expired.Count }
    'Expiring with no email'   = @($expiring | Where-Object { -not $_.Email }).Count
    'Notifications'            = if (-not $NotifyUsers) { 'Not sent (use -NotifyUsers)' } elseif ($TestRecipient) { "Test mode - all sent to $TestRecipient" } else { "Sent on days left: $($notifyOnDays -join ', ')" }
    'Notifications sent'       = @($notifications | Where-Object { $_.Result -like 'Sent*' }).Count
    'Run by'                   = "$env:USERDOMAIN\$env:USERNAME"
}

$sheets = @(
    @{ Name = 'Password Expiry'; Rows = $users
       Headers = @('SamAccountName', 'DisplayName', 'Email', 'PasswordLastSet', 'ExpiresOn', 'DaysLeft', 'Status', 'Notification', 'OU')
       Widths  = @(22, 28, 34, 18, 18, 10, 18, 50, 60) }
)
if ($NotifyUsers) {
    $sheets += @{ Name = 'Notifications'; Rows = $notifications; Headers = @('SamAccountName', 'DaysLeft', 'SentTo', 'Result'); Widths = @(22, 10, 34, 50) }
}
$sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(26, 80) }

if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

Write-Host ""
Write-Host ("Expiring within {0} days : {1:N0}" -f $WarningDays, $expiring.Count) -ForegroundColor $(if ($expiring.Count) { 'Yellow' } else { 'Green' })
if (-not $SkipExpired) { Write-Host ("Already expired          : {0:N0}" -f $expired.Count) -ForegroundColor $(if ($expired.Count) { 'Red' } else { 'Green' }) }
if ($NotifyUsers) { Write-Host ("Notifications sent       : {0:N0}" -f $summary['Notifications sent']) }
Write-Host "Report                   : $reportPath"

# --- Email the IT team ----------------------------------------------------

if ($SendReport) {
    $cols = @('SamAccountName', 'DisplayName', 'ExpiresOn', 'DaysLeft', 'Status', 'Notification')
    $sections = @(
        @{ Heading = "Expiring within $WarningDays days ($($expiring.Count))"; Html = ConvertTo-HtmlTable -Rows $expiring -Columns $cols -StatusColumn 'Status' -EmptyText 'None.' }
    )
    if (-not $SkipExpired) {
        $sections += @{ Heading = "Already expired ($($expired.Count))"; Html = ConvertTo-HtmlTable -Rows @($expired | Select-Object -First 100) -Columns $cols -StatusColumn 'Status' -EmptyText 'None.' }
    }
    $html = New-HtmlReport -Title 'Password expiry report' -IntroLines @("Enabled users whose passwords expire within $WarningDays days.") -Sections $sections
    if (Send-ToolkitMail -Config $config -Subject "Password expiry: $($expiring.Count) expiring, $($expired.Count) expired" -HtmlBody $html -Attachments $reportPath) {
        Write-Host 'Report emailed.'
    }
}
