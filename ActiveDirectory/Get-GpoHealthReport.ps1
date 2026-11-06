<#
.SYNOPSIS
    Checks Group Policy for conflicting settings, unlinked, empty or broken GPOs, and risky link settings.
.DESCRIPTION
    Uses each OU's effective GPO list (including inherited, enforced and blocked links) to find settings
    configured differently by more than one GPO, and shows which GPO wins. Also flags unlinked, empty and
    disabled GPOs, disabled links, block inheritance and GPOs that no computer or user can apply. Read-only.
.EXAMPLE
    .\Get-GpoHealthReport.ps1
.EXAMPLE
    .\Get-GpoHealthReport.ps1 -SearchBase 'OU=Workstations,DC=corp,DC=example,DC=com' -SendEmail
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: GroupPolicy and ActiveDirectory modules (RSAT)
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$Server,

    [switch]$SendEmail,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig
if (-not $PSBoundParameters.ContainsKey('Server')) { $Server = Get-Setting $config 'ActiveDirectory' 'Server' '' }

# --- Reading GPO XML ----------------------------------------------------------------
# GPO reports use XML namespaces, so nodes are matched by local name.

function Select-Local($Node, [string]$Path) {
    $xpath = ($Path -split '/' | ForEach-Object { if ($_ -eq '') { '' } elseif ($_ -eq '..') { '..' } else { "*[local-name()='$_']" } }) -join '/'
    if ($Node) { return @($Node.SelectNodes($xpath)) }
    return @()
}
function Get-LocalText($Node, [string]$Name) {
    $n = @(Select-Local $Node $Name)
    if ($n.Count) { return $n[0].InnerText }
    return ''
}

function Get-GpoSettings($GpoNode) {
    <# Returns @{ Settings = list of Scope/Category/Name/Value; ScopesWithData = @('Computer','User') } #>
    $settings = [System.Collections.Generic.List[object]]::new()
    $withData = @()
    foreach ($scope in 'Computer', 'User') {
        $scopeNode = @(Select-Local $GpoNode $scope)[0]
        if (-not $scopeNode) { continue }
        $extensions = @(Select-Local $scopeNode 'ExtensionData/Extension')
        if ($extensions.Count) { $withData += $scope }
        if ((Get-LocalText $scopeNode 'Enabled') -eq 'false') { continue }   # this half of the GPO is switched off
        $add = { param($Category, $Name, $Value) if ($Name) { $settings.Add([pscustomobject]@{ Scope = $scope; Category = $Category; Name = "$Name"; Value = "$Value" }) } }

        foreach ($ext in $extensions) {
            foreach ($p in (Select-Local $ext 'Policy')) {
                $vals = @($p.ChildNodes | Where-Object { $_.LocalName -in 'EditText', 'Numeric', 'DropDownList', 'CheckBox', 'ListBox' } |
                          ForEach-Object { "$(Get-LocalText $_ 'Name')=$(Get-LocalText $_ 'Value')$(Get-LocalText $_ 'State')" })
                & $add 'Administrative Templates' "$(Get-LocalText $p 'Category')\$(Get-LocalText $p 'Name')" ((@(Get-LocalText $p 'State') + $vals) -join '; ')
            }
            foreach ($s in (Select-Local $ext 'SecurityOptions')) {
                $name = Get-LocalText $s 'KeyName'; if (-not $name) { $name = Get-LocalText $s 'SystemAccessPolicyName' }
                & $add 'Security Options' $name ((@('SettingNumber', 'SettingString', 'SettingBoolean') | ForEach-Object { Get-LocalText $s $_ } | Where-Object { $_ }) -join '; ')
            }
            foreach ($a in (Select-Local $ext 'Account')) {
                & $add "Account Policy ($(Get-LocalText $a 'Type'))" (Get-LocalText $a 'Name') ((@('SettingNumber', 'SettingBoolean', 'SettingString') | ForEach-Object { Get-LocalText $a $_ } | Where-Object { $_ }) -join '; ')
            }
            foreach ($a in (Select-Local $ext 'Audit')) {
                & $add 'Audit Policy' (Get-LocalText $a 'Name') "Success=$(Get-LocalText $a 'SuccessAttempts'); Failure=$(Get-LocalText $a 'FailureAttempts')"
            }
            foreach ($a in (Select-Local $ext 'AuditSetting')) { & $add 'Advanced Audit Policy' (Get-LocalText $a 'SubcategoryName') (Get-LocalText $a 'SettingValue') }
            foreach ($u in (Select-Local $ext 'UserRightsAssignment')) {
                & $add 'User Rights' (Get-LocalText $u 'Name') ((@(Select-Local $u 'Member') | ForEach-Object { Get-LocalText $_ 'Name' } | Sort-Object) -join ', ')
            }
            foreach ($r in (Select-Local $ext 'RegistrySettings/Registry')) {
                $props = @(Select-Local $r 'Properties')[0]
                if ($props) { & $add 'Registry Preference' "$($props.hive)\$($props.key)\$($props.name)" "$($props.action) $($props.type)=$($props.value)" }
            }
        }
    }
    return @{ Settings = $settings; ScopesWithData = $withData }
}

function Get-GpoPermissionIssue($GpoNode) {
    $perms = @(Select-Local $GpoNode 'SecurityDescriptor/Permissions/TrusteePermissions')
    if (-not $perms.Count) { return '' }
    $allow   = @($perms | Where-Object { (Get-LocalText $_ 'Type/PermissionType') -ne 'Deny' })
    $apply   = @($allow | Where-Object { (Get-LocalText $_ 'Standard/GPOGroupedAccessEnum') -eq 'Apply Group Policy' })
    $readers = @($allow | ForEach-Object { Get-LocalText $_ 'Trustee/SID' })
    if (-not $apply.Count) { return 'No one has Apply permission - the GPO applies to nobody' }
    $canRead = $readers | Where-Object { $_ -eq 'S-1-5-11' -or $_ -match '-515$' }
    if (-not $canRead) { return 'Neither Authenticated Users nor Domain Computers can read it - user settings may not apply (MS16-072)' }
    return ''
}

# --- Main ------------------------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    foreach ($m in 'GroupPolicy', 'ActiveDirectory') {
        try { Import-Module $m -ErrorAction Stop } catch { throw "The $m module is not installed. Install RSAT: Group Policy Management and Active Directory tools." }
    }
    $ad = @{}; $gp = @{}
    if ($Server) { $ad.Server = $Server; $gp.Server = $Server }
    $domain = Get-ADDomain @ad
    $gp.Domain = $domain.DNSRoot
    if (-not $SearchBase) { $SearchBase = $domain.DistinguishedName }
    $now        = Get-Date
    $reportPath = New-ReportPath -Config $config -Prefix 'GpoHealth' -OutputFile $OutputFile
    $findings   = [System.Collections.Generic.List[object]]::new()
    function Add-Finding([string]$Severity, [string]$Finding, [string]$Gpo, [string]$Where, [string]$Detail) {
        $findings.Add([pscustomobject]@{ Severity = $Severity; Finding = $Finding; GPO = $Gpo; Where = $Where; Detail = $Detail })
    }

    Write-Host "Domain      : $($domain.DNSRoot)"
    Write-Host "Search base : $SearchBase"

    # 1. GPOs and their settings - one XML report for all GPOs
    Write-Host 'Reading GPOs...'
    $gpos = @(Get-GPO -All @gp)
    $gpoById = @{}; foreach ($g in $gpos) { $gpoById[$g.Id.ToString().ToLower()] = $g }
    $info = @{}
    $xmlNodes = @()
    try {
        [xml]$all = Get-GPOReport -All -ReportType Xml @gp -ErrorAction Stop
        $xmlNodes = @(Select-Local $all '/report/GPO')
        if (-not $xmlNodes.Count) { $xmlNodes = @(Select-Local $all '//GPO') }
    }
    catch {
        Write-Warning "Reading all GPO reports at once failed ($($_.Exception.Message)) - reading them one by one."
        $n = 0
        foreach ($g in $gpos) {
            $n++; Write-Progress -Activity 'Reading GPO reports' -Status "$n of $($gpos.Count)" -PercentComplete ($n / [math]::Max(1, $gpos.Count) * 100)
            try { [xml]$one = Get-GPOReport -Guid $g.Id -ReportType Xml @gp -ErrorAction Stop; $xmlNodes += @(Select-Local $one '/GPO') }
            catch { Add-Finding 'Low' 'GPO report could not be read' $g.DisplayName '' $_.Exception.Message }
        }
        Write-Progress -Activity 'Reading GPO reports' -Completed
    }
    foreach ($node in $xmlNodes) {
        $id = (Get-LocalText $node 'Identifier/Identifier').Trim('{', '}').ToLower()
        if (-not $id) { continue }
        $s = Get-GpoSettings $node
        $info[$id] = [pscustomobject]@{ Settings = $s.Settings; ScopesWithData = $s.ScopesWithData; PermissionIssue = (Get-GpoPermissionIssue $node) }
    }

    # 2. Effective GPO links for the search base and every OU under it
    Write-Host 'Reading OUs and links...'
    $targets = @($SearchBase) + @(Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase @ad | Select-Object -ExpandProperty DistinguishedName | Sort-Object { ($_ -split '(?<!\\),').Count }, { $_ })
    $linkRows = [System.Collections.Generic.List[object]]::new()
    $effective = @{}
    $linkedIds = @{}
    $n = 0
    foreach ($t in $targets) {
        $n++; Write-Progress -Activity 'Reading GPO inheritance' -Status "$n of $($targets.Count)" -PercentComplete ($n / [math]::Max(1, $targets.Count) * 100)
        try { $inh = Get-GPInheritance -Target $t @gp -ErrorAction Stop }
        catch { Add-Finding 'Low' 'Could not read GPO inheritance' '' $t $_.Exception.Message; continue }
        if ($inh.GpoInheritanceBlocked) { Add-Finding 'Low' 'Block inheritance is set - GPOs from parent OUs do not apply here unless enforced' '' $t '' }
        foreach ($l in @($inh.GpoLinks)) {
            $id = "$($l.GpoId)".ToLower()
            $linkedIds[$id] = $true
            $linkRows.Add([pscustomobject]@{ OU = $t; GPO = $l.DisplayName; LinkOrder = $l.Order; Enabled = [bool]$l.Enabled; Enforced = [bool]$l.Enforced })
            if (-not $l.Enabled) { Add-Finding 'Low' 'Link is disabled - the GPO does not apply here' $l.DisplayName $t '' }
        }
        # InheritedGpoLinks is the effective list in precedence order (1 = wins)
        $effective[$t] = @($inh.InheritedGpoLinks | Where-Object { $_.Enabled } | Sort-Object Order)
    }
    Write-Progress -Activity 'Reading GPO inheritance' -Completed
    # GPOs linked above the search base (e.g. at the domain root) also count as linked
    foreach ($t in $effective.Keys) { foreach ($l in $effective[$t]) { $linkedIds["$($l.GpoId)".ToLower()] = $true } }

    # 3. Conflicts - the same setting with different values from more than one effective GPO
    Write-Host 'Looking for conflicts...'
    $conflicts = @{}
    foreach ($t in $effective.Keys) {
        $bySetting = @{}
        foreach ($l in $effective[$t]) {
            $i = $info["$($l.GpoId)".ToLower()]
            if (-not $i) { continue }
            foreach ($s in $i.Settings) {
                $key = "$($s.Scope)|$($s.Category)|$($s.Name)"
                if (-not $bySetting.ContainsKey($key)) { $bySetting[$key] = [System.Collections.Generic.List[object]]::new() }
                $bySetting[$key].Add([pscustomobject]@{ Gpo = $l.DisplayName; Order = $l.Order; Value = $s.Value; Setting = $s })
            }
        }
        foreach ($key in $bySetting.Keys) {
            $entries = @($bySetting[$key] | Sort-Object Order)
            if (@($entries.Value | Select-Object -Unique).Count -lt 2) { continue }
            $sig = "$key|" + (($entries | ForEach-Object { "$($_.Gpo)=$($_.Value)" }) -join '|')
            if (-not $conflicts.ContainsKey($sig)) {
                $conflicts[$sig] = [pscustomobject]@{
                    Scope = $entries[0].Setting.Scope; Category = $entries[0].Setting.Category; Setting = $entries[0].Setting.Name
                    WinningGPO = $entries[0].Gpo; WinningValue = $entries[0].Value
                    OverriddenBy = (($entries | Select-Object -Skip 1 | ForEach-Object { "$($_.Gpo) [$($_.Value)]" }) -join '; ')
                    OUs = [System.Collections.Generic.List[string]]::new()
                }
            }
            $conflicts[$sig].OUs.Add($t)
        }
    }
    $conflictRows = @($conflicts.Values | ForEach-Object {
        [pscustomobject]@{
            Scope = $_.Scope; Category = $_.Category; Setting = $_.Setting; WinningGPO = $_.WinningGPO; WinningValue = $_.WinningValue
            LosingGPOs = $_.OverriddenBy; AffectedOUs = $_.OUs.Count; ExampleOU = $_.OUs[0]
        }
    } | Sort-Object Category, Setting)
    foreach ($c in $conflictRows) { Add-Finding 'Medium' "Conflicting setting: $($c.Setting)" $c.WinningGPO $c.ExampleOU "Wins with [$($c.WinningValue)] over $($c.LosingGPOs) - $($c.AffectedOUs) OU(s)" }

    # 4. GPO-level checks
    $gpoRows = foreach ($g in $gpos) {
        $id  = $g.Id.ToString().ToLower()
        $i   = $info[$id]
        $linked = $linkedIds.ContainsKey($id)
        $status = "$($g.GpoStatus)"
        $count  = if ($i) { $i.Settings.Count } else { $null }
        $issues = @()
        if (-not $linked) { $issues += 'Not linked within the search base' }
        if ($i -and -not $i.ScopesWithData.Count) { $issues += 'Empty - no settings' }
        if ($status -eq 'AllSettingsDisabled') { $issues += 'All settings disabled' }
        elseif ($status -eq 'ComputerSettingsDisabled' -and $i -and $i.ScopesWithData -contains 'Computer') { $issues += 'Computer settings exist but are disabled' }
        elseif ($status -eq 'UserSettingsDisabled' -and $i -and $i.ScopesWithData -contains 'User') { $issues += 'User settings exist but are disabled' }
        if ($i -and $i.PermissionIssue) { $issues += $i.PermissionIssue }

        if ($linked -and $i -and $i.PermissionIssue) { Add-Finding 'High' $i.PermissionIssue $g.DisplayName '' 'Check the Delegation tab in Group Policy Management' }
        if ($linked -and $status -eq 'AllSettingsDisabled') { Add-Finding 'Medium' 'Linked GPO has all settings disabled - it does nothing' $g.DisplayName '' '' }
        if ($issues -match 'exist but are disabled') { Add-Finding 'Low' ($issues -match 'exist but are disabled')[0] $g.DisplayName '' 'Often left behind after troubleshooting' }
        if (-not $linked) { Add-Finding 'Low' 'GPO is not linked within the search base' $g.DisplayName '' "Last changed $($g.ModificationTime.ToString('yyyy-MM-dd'))" }
        elseif ($i -and -not $i.ScopesWithData.Count) { Add-Finding 'Low' 'Linked GPO is empty' $g.DisplayName '' '' }

        [pscustomobject]@{
            GPO = $g.DisplayName; Status = $status; Linked = $linked; Settings = $count; WmiFilter = if ($g.WmiFilter) { $g.WmiFilter.Name } else { '' }
            Created = $g.CreationTime; Modified = $g.ModificationTime; Issues = ($issues -join '; '); Id = $g.Id
        }
    }

    # --- Export --------------------------------------------------------------------
    $order = @{ High = 0; Medium = 1; Low = 2 }
    $sortedFindings = @($findings | Sort-Object @{ Expression = { $order[$_.Severity] } }, Finding, GPO)
    $counts = [ordered]@{}; foreach ($s in 'High', 'Medium', 'Low') { $counts[$s] = @($sortedFindings | Where-Object Severity -eq $s).Count }
    $summary = [ordered]@{
        'Domain'                 = $domain.DNSRoot
        'Search base'            = $SearchBase
        'Run at'                 = $now.ToString('yyyy-MM-dd HH:mm:ss')
        'GPOs'                   = $gpos.Count
        'OUs checked'            = $targets.Count
        'Conflicting settings'   = $conflictRows.Count
        'Unlinked GPOs'          = @($gpoRows | Where-Object { -not $_.Linked }).Count
        'Findings'               = "$($counts.High) high, $($counts.Medium) medium, $($counts.Low) low"
        'Settings compared'      = 'Administrative Templates, Security Options, Account/Audit policy, User Rights and Registry preferences. Other preference types and scripts are not compared.'
        'Note'                   = 'Site-linked GPOs and WMI/security filtering are not taken into account when working out which GPO wins.'
    }

    $saved = Save-Workbook -FilePath $reportPath -Sheets @(
        @{ Name = 'Findings'; Rows = $sortedFindings; Headers = @('Severity', 'Finding', 'GPO', 'Where', 'Detail'); Widths = @(10, 60, 36, 50, 80) }
        @{ Name = 'Conflicts'; Rows = $conflictRows; Headers = @('Scope', 'Category', 'Setting', 'WinningGPO', 'WinningValue', 'LosingGPOs', 'AffectedOUs', 'ExampleOU'); Widths = @(10, 26, 60, 30, 30, 60, 12, 50) }
        @{ Name = 'GPOs'; Rows = @($gpoRows | Sort-Object GPO); Headers = @('GPO', 'Status', 'Linked', 'Settings', 'WmiFilter', 'Created', 'Modified', 'Issues', 'Id'); Widths = @(36, 24, 8, 10, 20, 18, 18, 70, 38) }
        @{ Name = 'Links'; Rows = @($linkRows | Sort-Object OU, LinkOrder); Headers = @('OU', 'GPO', 'LinkOrder', 'Enabled', 'Enforced'); Widths = @(60, 36, 10, 9, 10) }
        @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(22, 110) }
    )
    if (-not $saved) { exit 1 }

    Write-Host ""
    Write-Host ("GPOs: {0}. Conflicts: {1}. Findings: {2}" -f $gpos.Count, $conflictRows.Count, $summary.Findings) -ForegroundColor $(if ($counts.High) { 'Red' } elseif ($counts.Medium) { 'Yellow' } else { 'Green' })
    $sortedFindings | Where-Object Severity -eq 'High' | Select-Object -First 10 | ForEach-Object { Write-Host "  HIGH  $($_.GPO): $($_.Finding)" -ForegroundColor Red }
    Write-Host "Report: $reportPath"

    if ($SendEmail) {
        $html = New-HtmlReport -Title 'Group Policy health report' -IntroLines @("$($domain.DNSRoot): $($gpos.Count) GPOs, $($conflictRows.Count) conflicting setting(s). Findings: $($summary.Findings).") -Sections @(
            @{ Heading = 'High and medium findings'; Html = ConvertTo-HtmlTable -Rows @($sortedFindings | Where-Object Severity -ne 'Low' | Select-Object -First 100) -Columns @('Severity', 'Finding', 'GPO', 'Where', 'Detail') -StatusColumn 'Severity' -EmptyText 'None.' }
        )
        if (Send-ToolkitMail -Config $config -Subject "Group Policy health: $($summary.Findings)" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
    }
}
