<#
.SYNOPSIS
    Finds Entra app registration secrets and certificates that are expired or expiring soon.
.DESCRIPTION
    Covers app registrations and enterprise apps (including SAML signing certificates).
    Also flags long-lived secrets and apps with no owners. Read-only.
.EXAMPLE
    .\Get-AppCredentialExpiryReport.ps1
.EXAMPLE
    .\Get-AppCredentialExpiryReport.ps1 -WarningDays 60 -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: Microsoft.Graph.Authentication module, Global Reader (or Application Administrator)
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [ValidateRange(1, 3650)] [int]$WarningDays,
    [ValidateRange(0, 3650)] [int]$CriticalDays,
    [switch]$ExpiringOnly,

    [switch]$SendEmail,
    [switch]$AlwaysEmail,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('WarningDays'))  { $WarningDays  = [int](Get-Setting $config 'AppCredentials' 'WarningDays' 30) }
if (-not $PSBoundParameters.ContainsKey('CriticalDays')) { $CriticalDays = [int](Get-Setting $config 'AppCredentials' 'CriticalDays' 7) }
$ignoreExpired = [int](Get-Setting $config 'AppCredentials' 'IgnoreExpiredOlderThanDays' 90)
$maxLifetime   = [int](Get-Setting $config 'AppCredentials' 'MaxSecretLifetimeDays' 730)

$scopes = @('Application.Read.All')

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'Global Reader (or Application Administrator)'
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'AppCredentials' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        $tenantName = Get-GraphTenantName
        Write-Host "Tenant : $tenantName"

        # --- Collect -----------------------------------------------------

        Write-Host 'Getting app registrations...'
        $apps = @(Get-GraphAll 'v1.0/applications?$select=id,appId,displayName,createdDateTime,signInAudience,passwordCredentials,keyCredentials&$top=999')
        Write-Host 'Getting enterprise apps...'
        $sps = @(Get-GraphAll 'v1.0/servicePrincipals?$select=id,appId,displayName,servicePrincipalType,preferredSingleSignOnMode,accountEnabled,passwordCredentials,keyCredentials&$top=999' |
                 Where-Object { @($_.passwordCredentials | Where-Object { $_ }).Count -or @($_.keyCredentials | Where-Object { $_ }).Count })

        $objects = @()
        $objects += @($apps | Where-Object { @($_.passwordCredentials | Where-Object { $_ }).Count -or @($_.keyCredentials | Where-Object { $_ }).Count } |
                      ForEach-Object { [pscustomobject]@{ Kind = 'App registration'; Path = 'applications'; Item = $_ } })
        $objects += @($sps | ForEach-Object { [pscustomobject]@{ Kind = 'Enterprise app'; Path = 'servicePrincipals'; Item = $_ } })
        Write-Host ("  {0} app registration(s), {1} with credentials; {2} enterprise app(s) with credentials" -f $apps.Count, @($objects | Where-Object { $_.Kind -eq 'App registration' }).Count, $sps.Count)

        # --- Build rows --------------------------------------------------

        $rows    = [System.Collections.Generic.List[object]]::new()
        $appRows = [System.Collections.Generic.List[object]]::new()
        $order   = @{ Expired = 0; Critical = 1; Warning = 2; OK = 3 }
        $i = 0
        foreach ($o in $objects) {
            $i++
            if ($i % 10 -eq 0 -or $i -eq 1) { Write-Progress -Activity 'Checking app credentials' -Status "$i of $($objects.Count)" -PercentComplete ($i / [math]::Max(1, $objects.Count) * 100) }
            $item = $o.Item

            $owners = ''
            try {
                $owners = (@(Get-GraphAll "v1.0/$($o.Path)/$($item.id)/owners?`$select=displayName,userPrincipalName" |
                             ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }) -join ', ')
            }
            catch { $errors.Add([pscustomobject]@{ Section = 'Owners'; Error = "$($item.displayName): $($_.Exception.Message)" }) }

            $isSaml = $o.Kind -eq 'Enterprise app' -and "$($item.preferredSingleSignOnMode)" -eq 'saml'
            $creds = @()
            foreach ($c in @($item.passwordCredentials | Where-Object { $_ })) {
                if ($isSaml) { continue }   # the SAML certificate's key password - covered by the certificate itself
                $creds += [pscustomobject]@{ Type = 'Client secret'; Cred = $c }
            }
            foreach ($c in @($item.keyCredentials | Where-Object { $_ })) {
                $creds += [pscustomobject]@{ Type = $(if ($isSaml) { 'SAML signing certificate' } else { 'Certificate' }); Cred = $c }
            }

            $seen = @{}
            $objRows = @()
            foreach ($entry in $creds) {
                $c      = $entry.Cred
                $end    = ConvertTo-LocalTime $c.endDateTime
                $start  = ConvertTo-LocalTime $c.startDateTime
                if (-not $end) { continue }
                # A certificate is stored twice (Sign + Verify) on SAML apps - list it once
                $key = "$($entry.Type)|$($end.ToString('s'))|$($c.customKeyIdentifier)"
                if ($seen.ContainsKey($key)) { continue }
                $seen[$key] = $true

                $daysLeft = [int][math]::Floor(($end - $now).TotalDays)
                if ($daysLeft -lt -$ignoreExpired) { continue }
                $status = if ($daysLeft -lt 0) { 'Expired' } elseif ($daysLeft -le $CriticalDays) { 'Critical' } elseif ($daysLeft -le $WarningDays) { 'Warning' } else { 'OK' }

                $issues = @()
                if ($entry.Type -eq 'Client secret' -and $start -and ($end - $start).TotalDays -gt $maxLifetime) {
                    $issues += "Long-lived secret ($([math]::Round(($end - $start).TotalDays / 365, 1)) years) - use a shorter lifetime or a certificate"
                }
                if ($status -eq 'Expired') { $issues += 'Expired - remove it if a newer credential has replaced it' }

                $objRows += [pscustomobject]@{
                    Status         = $status
                    DaysLeft       = $daysLeft
                    Expires        = $end
                    App            = $item.displayName
                    AppId          = $item.appId
                    ObjectType     = $o.Kind
                    CredentialType = $entry.Type
                    Description    = "$($c.displayName)"
                    KeyId          = $c.keyId
                    Created        = $start
                    Owners         = $owners
                    Issues         = ($issues -join '; ')
                }
            }
            if (-not $objRows.Count) { continue }

            $appIssues = @()
            if (-not $owners) { $appIssues += 'No owners - nobody is responsible for renewing it' }
            if (-not @($objRows | Where-Object { $_.Status -ne 'Expired' }).Count) { $appIssues += 'All credentials expired - the app may be broken or no longer used' }
            if ($o.Kind -eq 'Enterprise app' -and $item.accountEnabled -eq $false) { $appIssues += 'Disabled' }
            $appIssues += @($objRows | Where-Object { $_.Issues -like 'Long-lived*' } | ForEach-Object { 'Long-lived secret' } | Select-Object -First 1)

            $worst    = ($objRows | Sort-Object @{ Expression = { $order[$_.Status] } } | Select-Object -First 1).Status
            $upcoming = @($objRows | Where-Object { $_.DaysLeft -ge 0 } | Sort-Object Expires | Select-Object -First 1)
            $next     = if ($upcoming.Count) { $upcoming[0] } else { $objRows | Sort-Object Expires -Descending | Select-Object -First 1 }

            if (-not $ExpiringOnly -or $worst -ne 'OK' -or $appIssues.Count) {
                foreach ($r in $objRows) { if (-not $ExpiringOnly -or $r.Status -ne 'OK' -or $r.Issues) { $rows.Add($r) } }
                $appRows.Add([pscustomobject]@{
                    Status      = $worst
                    App         = $item.displayName
                    AppId       = $item.appId
                    ObjectType  = $o.Kind
                    Credentials = $objRows.Count
                    NextExpiry  = $next.Expires
                    DaysLeft    = $next.DaysLeft
                    Owners      = if ($owners) { $owners } else { '(none)' }
                    Issues      = (@($appIssues | Where-Object { $_ }) -join '; ')
                })
            }
        }
        Write-Progress -Activity 'Checking app credentials' -Completed

        # --- Export ------------------------------------------------------

        $rowsSorted = @($rows | Sort-Object @{ Expression = { $order[$_.Status] } }, DaysLeft, App)
        $appsSorted = @($appRows | Sort-Object @{ Expression = { $order[$_.Status] } }, DaysLeft, App)
        $counts = [ordered]@{}
        foreach ($s in 'Expired', 'Critical', 'Warning', 'OK') { $counts[$s] = @($rowsSorted | Where-Object { $_.Status -eq $s }).Count }

        $summary = [ordered]@{
            'Tenant'                    = $tenantName
            'Run at'                    = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Thresholds'                = "Warning <= $WarningDays days, Critical <= $CriticalDays days"
            'App registrations'         = $apps.Count
            'Apps with credentials'     = $appRows.Count
            'Expired'                   = $counts.Expired
            'Critical'                  = $counts.Critical
            'Warning'                   = $counts.Warning
            'OK'                        = if ($ExpiringOnly) { 'Not listed (-ExpiringOnly)' } else { $counts.OK }
            'Apps with no owners'       = @($appsSorted | Where-Object { $_.Owners -eq '(none)' }).Count
            'Long-lived secrets'        = @($rowsSorted | Where-Object { $_.Issues -like 'Long-lived*' }).Count
            'Note'                      = "Credentials that expired more than $ignoreExpired days ago are not listed. Secrets longer than $maxLifetime days are flagged."
            'Run by'                    = Get-GraphRunBy
        }

        $saved = Save-Workbook -FilePath $reportPath -Sheets @(
            @{ Name = 'Credentials'; Rows = $rowsSorted
               Headers = @('Status', 'DaysLeft', 'Expires', 'App', 'ObjectType', 'CredentialType', 'Description', 'Owners', 'Issues', 'Created', 'AppId', 'KeyId')
               Widths  = @(10, 10, 18, 36, 17, 24, 26, 40, 60, 18, 38, 38) }
            @{ Name = 'Apps'; Rows = $appsSorted
               Headers = @('Status', 'App', 'ObjectType', 'Credentials', 'NextExpiry', 'DaysLeft', 'Owners', 'Issues', 'AppId')
               Widths  = @(10, 36, 17, 12, 18, 10, 40, 60, 38) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(20, 120) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(24, 100) }
        )
        if (-not $saved) { exit 1 }

        Write-Host ""
        foreach ($s in 'Expired', 'Critical', 'Warning') {
            Write-Host ("{0,-9} {1}" -f $s, $counts[$s]) -ForegroundColor $(if ($counts[$s]) { if ($s -eq 'Warning') { 'Yellow' } else { 'Red' } } else { 'Green' })
        }
        Write-Host ("Apps with no owners: {0}" -f $summary['Apps with no owners'])
        Write-Host "Report : $reportPath"

        $attention = @($rowsSorted | Where-Object { $_.Status -ne 'OK' })
        if ($SendEmail) {
            if (-not $attention.Count -and -not $AlwaysEmail) {
                Write-Host 'Nothing expiring - no email sent (use -AlwaysEmail to send anyway).'
            }
            else {
                $intro = @("$tenantName - $($counts.Expired) expired, $($counts.Critical) critical (<= $CriticalDays days), $($counts.Warning) warning (<= $WarningDays days).")
                $html = New-HtmlReport -Title 'App secrets and certificates expiry' -IntroLines $intro -Sections @(
                    @{ Heading = "Needs attention ($($attention.Count))"
                       Html    = ConvertTo-HtmlTable -Rows $attention -Columns @('Status', 'DaysLeft', 'Expires', 'App', 'CredentialType', 'Owners') -StatusColumn 'Status' -EmptyText 'Nothing expiring.' }
                )
                $state = if ($counts.Expired -or $counts.Critical) { 'CRITICAL' } elseif ($counts.Warning) { 'Warning' } else { 'All OK' }
                if (Send-ToolkitMail -Config $config -Subject "[$state] App credentials: $($counts.Expired) expired, $($counts.Critical) critical, $($counts.Warning) warning" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
            }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
