<#
.SYNOPSIS
    Reports SharePoint Online sites, their sharing settings and storage, and the external users on each site.
.DESCRIPTION
    Lists every site (OneDrive optional) with owner, storage, last activity and sharing level, and flags
    sites shared with "Anyone" links, sites close to their storage limit and inactive sites. Read-only.
.EXAMPLE
    .\Get-SharePointSiteReport.ps1 -AdminUrl https://contoso-admin.sharepoint.com
.EXAMPLE
    .\Get-SharePointSiteReport.ps1 -IncludeExternalUsers
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: Microsoft.Online.SharePoint.PowerShell module, SharePoint Administrator (or Global Reader)
#>
[CmdletBinding()]
param(
    [string]$AdminUrl,
    [switch]$IncludeExternalUsers,
    [switch]$IncludeOneDrive,
    [ValidateRange(1, 3650)] [int]$InactiveDays,

    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('AdminUrl'))     { $AdminUrl     = Get-Setting $config 'SharePoint' 'AdminUrl' '' }
if (-not $PSBoundParameters.ContainsKey('InactiveDays')) { $InactiveDays = [int](Get-Setting $config 'SharePoint' 'InactiveDays' 365) }
if (-not $AdminUrl) { throw 'No SharePoint admin URL. Use -AdminUrl https://<tenant>-admin.sharepoint.com.' }

$sharingNames = @{
    Disabled                         = 'Only people in your organisation'
    ExistingExternalUserSharingOnly  = 'Existing guests only'
    ExternalUserSharingOnly          = 'New and existing guests'
    ExternalUserAndGuestSharing      = 'Anyone (anonymous links)'
}

if ($MyInvocation.InvocationName -ne '.') {
    $module = 'Microsoft.Online.SharePoint.PowerShell'
    if (-not (Get-Module -ListAvailable -Name $module)) { throw "The SharePoint Online module is not installed. Run: Install-Module $module -Scope CurrentUser" }
    # The module is built for Windows PowerShell; in PowerShell 7 load it through the compatibility layer
    if ($PSVersionTable.PSEdition -eq 'Core' -and $script:IsWindowsHost) { Import-Module $module -UseWindowsPowerShell -WarningAction SilentlyContinue -ErrorAction Stop }
    else { Import-Module $module -DisableNameChecking -ErrorAction Stop }

    Write-Host "Connecting to $AdminUrl..."
    Connect-SPOService -Url $AdminUrl -ErrorAction Stop
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'SharePointSites' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()

        $tenantSharing = ''
        try { $tenantSharing = "$((Get-SPOTenant).SharingCapability)" } catch { $errors.Add([pscustomobject]@{ Site = '(tenant)'; Check = 'Tenant settings'; Error = $_.Exception.Message }) }

        Write-Host 'Getting sites...'
        $sites = @(Get-SPOSite -Limit All -IncludePersonalSite:([bool]$IncludeOneDrive) |
                   Where-Object { $IncludeOneDrive -or "$($_.Template)" -notlike 'SPSPERS*' })
        Write-Host "  $($sites.Count) site(s)"

        $siteRows = [System.Collections.Generic.List[object]]::new()
        $extRows  = [System.Collections.Generic.List[object]]::new()
        $i = 0
        foreach ($s in $sites) {
            $i++
            $external = @()
            $extFailed = $false
            if ($IncludeExternalUsers) {
                Write-Progress -Activity 'Checking sites for external users' -Status "$i of $($sites.Count): $($s.Url)" -PercentComplete ($i / [math]::Max(1, $sites.Count) * 100)
                try {
                    $external = @(Get-SPOUser -Site $s.Url -Limit All -ErrorAction Stop | Where-Object { $_.LoginName -like '*#ext#*' -or $_.LoginName -like '*urn:spo:guest*' })
                    foreach ($u in $external) {
                        # Guest login names look like i:0#.f|membership|jo_example.com#ext#@tenant.onmicrosoft.com
                        $email = if ($u.LoginName -match 'urn:spo:guest#(.+)$') { $Matches[1] }
                                 elseif ($u.LoginName -match '\|([^|]+)#ext#') { $Matches[1] -replace '_([^_]+)$', '@$1' }
                                 else { $u.LoginName }
                        $extRows.Add([pscustomobject]@{
                            Site = $s.Title; Url = $s.Url; DisplayName = $u.DisplayName
                            Email = $email
                            SiteAdmin = [bool]$u.IsSiteAdmin; LoginName = $u.LoginName
                        })
                    }
                }
                catch { $extFailed = $true; $errors.Add([pscustomobject]@{ Site = $s.Url; Check = 'External users'; Error = $_.Exception.Message }) }
            }

            $usedMb   = [double]$s.StorageUsageCurrent
            $quotaMb  = [double]$s.StorageQuota
            $pctUsed  = if ($quotaMb -gt 0) { [math]::Round($usedMb / $quotaMb * 100, 1) } else { $null }
            $lastMod  = $s.LastContentModifiedDate
            $idleDays = if ($lastMod -and $lastMod.Year -gt 2000) { [int]($now - $lastMod).TotalDays } else { $null }
            $sharing  = "$($s.SharingCapability)"

            $issues = @()
            $risk   = 'OK'
            if ($sharing -eq 'ExternalUserAndGuestSharing') { $issues += 'Anyone (anonymous) links allowed'; $risk = 'Medium' }
            if ($null -ne $pctUsed -and $pctUsed -ge 90)    { $issues += "Storage $pctUsed% of quota"; $risk = 'Medium' }
            if ($null -ne $idleDays -and $idleDays -gt $InactiveDays) { $issues += "No changes for $idleDays days - archive or delete?"; if ($risk -eq 'OK') { $risk = 'Low' } }
            if ("$($s.LockState)" -ne 'Unlock' -and "$($s.LockState)") { $issues += "Locked ($($s.LockState))" }
            if ($external.Count) { $issues += "$($external.Count) external user(s)" }

            $siteRows.Add([pscustomobject]@{
                Risk = $risk; Title = $s.Title; Url = $s.Url; Template = "$($s.Template)"; Owner = $s.Owner
                StorageGB = [math]::Round($usedMb / 1024, 2); QuotaGB = [math]::Round($quotaMb / 1024, 1); PercentUsed = $pctUsed
                LastModified = $lastMod; DaysSinceChange = $idleDays
                Sharing = if ($sharingNames.ContainsKey($sharing)) { $sharingNames[$sharing] } else { $sharing }
                ExternalUsers = if ($extFailed) { 'Error' } elseif ($IncludeExternalUsers) { $external.Count } else { '' }
                M365Group = if ("$($s.GroupId)" -and "$($s.GroupId)" -ne '00000000-0000-0000-0000-000000000000') { 'Yes' } else { 'No' }
                Issues = ($issues -join '; ')
            })
        }
        Write-Progress -Activity 'Checking sites for external users' -Completed

        $order = @{ Medium = 0; Low = 1; OK = 2 }
        $summary = [ordered]@{
            'Admin URL'               = $AdminUrl
            'Run at'                  = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Tenant sharing level'    = if ($sharingNames.ContainsKey($tenantSharing)) { $sharingNames[$tenantSharing] } else { $tenantSharing }
            'Sites'                   = $siteRows.Count
            'Total storage (GB)'      = [math]::Round((($siteRows | Measure-Object StorageGB -Sum).Sum), 2)
            'Sites allowing Anyone links' = @($siteRows | Where-Object { $_.Sharing -like 'Anyone*' }).Count
            "Inactive > $InactiveDays days" = @($siteRows | Where-Object { $_.DaysSinceChange -gt $InactiveDays }).Count
            'External users'          = if ($IncludeExternalUsers) { "$($extRows.Count) across $(@($extRows.Url | Select-Object -Unique).Count) site(s)" } else { 'Not checked (use -IncludeExternalUsers)' }
            'Note'                    = 'A site can never be shared more widely than the tenant sharing level allows.'
        }

        $sheets = @(
            @{ Name = 'Sites'; Rows = @($siteRows | Sort-Object @{ Expression = { $order[$_.Risk] } }, Title)
               Headers = @('Risk', 'Title', 'Url', 'Template', 'Owner', 'StorageGB', 'QuotaGB', 'PercentUsed', 'LastModified', 'DaysSinceChange', 'Sharing', 'ExternalUsers', 'M365Group', 'Issues')
               Widths  = @(9, 34, 60, 18, 34, 10, 9, 12, 18, 16, 28, 14, 11, 60) }
        )
        if ($IncludeExternalUsers) { $sheets += @{ Name = 'External Users'; Rows = @($extRows | Sort-Object Site, DisplayName); Headers = @('Site', 'Url', 'DisplayName', 'Email', 'SiteAdmin', 'LoginName'); Widths = @(34, 60, 28, 36, 10, 60) } }
        $sheets += @{ Name = 'Errors'; Rows = $errors; Headers = @('Site', 'Check', 'Error'); Widths = @(60, 16, 100) }
        $sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(28, 90) }
        if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

        Write-Host ""
        Write-Host ("Sites: {0}, {1} GB. Tenant sharing: {2}" -f $siteRows.Count, $summary['Total storage (GB)'], $summary['Tenant sharing level'])
        Write-Host ("Anyone links allowed on {0} site(s); {1} inactive site(s)" -f $summary['Sites allowing Anyone links'], $summary["Inactive > $InactiveDays days"])
        if ($IncludeExternalUsers) { Write-Host "External users: $($summary['External users'])" }
        if ($errors.Count) { Write-Host ("Errors: {0} - see the 'Errors' sheet" -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report: $reportPath"
    }
    finally {
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
}
