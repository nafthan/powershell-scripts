<#
.SYNOPSIS
    Finds Exchange Online mail forwarding to external addresses and suspicious inbox rules.
.DESCRIPTION
    Checks mailbox forwarding, inbox rules (forward, redirect, and rules that hide mail - a common
    sign of a compromised account), transport rules and the tenant's external forwarding settings. Read-only.
.EXAMPLE
    .\Get-MailboxForwardingAudit.ps1
.EXAMPLE
    .\Get-MailboxForwardingAudit.ps1 -Mailbox user@contoso.com, shared@contoso.com
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: ExchangeOnlineManagement module, Exchange Administrator (or Global Reader)
#>
[CmdletBinding()]
param(
    [string[]]$Mailbox,
    [switch]$SkipInboxRules,
    [switch]$IncludeHiddenRules,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$UseDeviceCode,

    [switch]$SendEmail,
    [switch]$AlwaysEmail,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

$Mailbox = Split-ListValue $Mailbox

# --- Main -----------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitExchange -Config $config -Bound $PSBoundParameters
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'MailForwarding' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        $findings   = [System.Collections.Generic.List[object]]::new()
        function Add-Finding([string]$Risk, [string]$Mailbox, [string]$Source, [string]$Finding, [string]$Target, [string]$Detail) {
            $findings.Add([pscustomobject]@{ Risk = $Risk; Mailbox = $Mailbox; Source = $Source; Finding = $Finding; Target = $Target; Detail = $Detail })
        }

        $internal = @(Get-AcceptedDomain | ForEach-Object { "$($_.DomainName)".ToLower() })
        Write-Host "Internal domains: $($internal -join ', ')"

        # --- Mailboxes and mailbox-level forwarding ----------------------

        Write-Host 'Getting mailboxes...'
        $props = 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward', 'WhenCreated'
        if ($Mailbox.Count) { $mailboxes = @($Mailbox | ForEach-Object { Get-EXOMailbox -Identity $_ -Properties $props -ErrorAction Stop }) }
        else { $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties $props) }
        Write-Host "  $($mailboxes.Count) mailbox(es)"

        $mailboxRows = [System.Collections.Generic.List[object]]::new()
        foreach ($m in $mailboxes) {
            $targets = @()
            if ($m.ForwardingSmtpAddress) { $targets += [pscustomobject]@{ Kind = 'SMTP forwarding'; Info = Get-AddressInfo "$($m.ForwardingSmtpAddress)" $internal } }
            if ($m.ForwardingAddress) {
                $info = [pscustomobject]@{ Address = "$($m.ForwardingAddress)"; External = $false }
                try {
                    $r = Get-EXORecipient -Identity "$($m.ForwardingAddress)" -ErrorAction Stop
                    $smtp = if ($r.ExternalEmailAddress) { "$($r.ExternalEmailAddress)" } else { "$($r.PrimarySmtpAddress)" }
                    $info = Get-AddressInfo $smtp $internal
                }
                catch { $errors.Add([pscustomobject]@{ Section = 'Forwarding address'; Error = "$($m.PrimarySmtpAddress): $($_.Exception.Message)" }) }
                $targets += [pscustomobject]@{ Kind = 'Forwarding address'; Info = $info }
            }
            foreach ($t in $targets) {
                $keep = if ($m.DeliverToMailboxAndForward) { 'Yes' } else { 'No - mail is not kept in the mailbox' }
                $mailboxRows.Add([pscustomobject]@{
                    Mailbox = "$($m.PrimarySmtpAddress)"; DisplayName = $m.DisplayName; Type = "$($m.RecipientTypeDetails)"
                    ForwardingType = $t.Kind; ForwardsTo = $t.Info.Address; External = $t.Info.External; KeepsCopy = $keep
                })
                if ($t.Info.External) { Add-Finding 'High' "$($m.PrimarySmtpAddress)" 'Mailbox forwarding' 'Forwards to an external address' $t.Info.Address "$($t.Kind); keeps a copy: $keep" }
                else                  { Add-Finding 'Low'  "$($m.PrimarySmtpAddress)" 'Mailbox forwarding' 'Forwards internally' $t.Info.Address "$($t.Kind); keeps a copy: $keep" }
            }
        }

        # --- Inbox rules -------------------------------------------------

        $ruleRows = [System.Collections.Generic.List[object]]::new()
        $rulesChecked = 0
        if (-not $SkipInboxRules) {
            $ruleMailboxes = @($mailboxes | Where-Object { "$($_.RecipientTypeDetails)" -in 'UserMailbox', 'SharedMailbox' })
            $i = 0
            foreach ($m in $ruleMailboxes) {
                $i++
                Write-Progress -Activity 'Reading inbox rules' -Status "$i of $($ruleMailboxes.Count): $($m.PrimarySmtpAddress)" -PercentComplete ($i / [math]::Max(1, $ruleMailboxes.Count) * 100)
                try {
                    $ruleParams = @{ Mailbox = "$($m.PrimarySmtpAddress)"; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
                    if ($IncludeHiddenRules) { $ruleParams.IncludeHidden = $true }
                    $rules = @(Get-InboxRule @ruleParams)
                }
                catch { $errors.Add([pscustomobject]@{ Section = 'Inbox rules'; Error = "$($m.PrimarySmtpAddress): $($_.Exception.Message)" }); continue }
                $rulesChecked += $rules.Count

                foreach ($rule in $rules) {
                    $a = Get-InboxRuleAssessment $rule $internal
                    if (-not $a.Risk) { continue }
                    $ruleRows.Add([pscustomobject]@{
                        Risk = $a.Risk; Mailbox = "$($m.PrimarySmtpAddress)"; Rule = "$($rule.Name)"; Enabled = [bool]$rule.Enabled
                        Actions = $a.Actions; ForwardsTo = $a.Targets; Conditions = $a.Conditions
                        Issues = ($a.Issues -join '; '); RuleIdentity = "$($rule.Identity)"
                    })
                    Add-Finding $a.Risk "$($m.PrimarySmtpAddress)" 'Inbox rule' ($a.Short -join ' + ') $a.Targets "Rule '$($rule.Name)' ($(if ($rule.Enabled) { 'enabled' } else { 'disabled' })): $($a.Actions)"
                }
            }
            Write-Progress -Activity 'Reading inbox rules' -Completed
        }

        # --- Transport rules ---------------------------------------------

        $transportRows = [System.Collections.Generic.List[object]]::new()
        try {
            foreach ($t in @(Get-TransportRule -ResultSize Unlimited)) {
                $targets = @()
                foreach ($field in 'RedirectMessageTo', 'BlindCopyTo', 'CopyTo', 'AddToRecipients') {
                    foreach ($entry in @($t.$field | Where-Object { $_ })) { $targets += [pscustomobject]@{ Action = $field; Info = Get-AddressInfo "$entry" $internal } }
                }
                if (-not $targets.Count) { continue }
                $ext = @($targets | Where-Object { $_.Info.External })
                $transportRows.Add([pscustomobject]@{
                    Rule = $t.Name; State = "$($t.State)"; Actions = (@($targets.Action | Sort-Object -Unique) -join ', ')
                    Targets = (@($targets | ForEach-Object { "$($_.Info.Address)$(if ($_.Info.External) { ' (external)' })" }) -join ', ')
                    External = [bool]$ext.Count; Priority = $t.Priority; Modified = $t.WhenChanged
                })
                if ($ext.Count) { Add-Finding $(if ("$($t.State)" -eq 'Enabled') { 'High' } else { 'Low' }) '(organisation)' 'Transport rule' 'Copies or redirects mail to an external address' (($ext | ForEach-Object { $_.Info.Address }) -join ', ') "Rule '$($t.Name)' ($($t.State))" }
            }
        }
        catch { $errors.Add([pscustomobject]@{ Section = 'Transport rules'; Error = $_.Exception.Message }) }

        # --- Tenant settings ---------------------------------------------

        $settings = [System.Collections.Generic.List[object]]::new()
        $externalBlocked = $null
        try {
            foreach ($p in @(Get-HostedOutboundSpamFilterPolicy)) {
                $mode = "$($p.AutoForwardingMode)"
                $meaning = switch ($mode) { 'Off' { 'External automatic forwarding blocked' } 'On' { 'External automatic forwarding allowed' } default { 'Automatic (Microsoft blocks external automatic forwarding)' } }
                $settings.Add([pscustomobject]@{ Setting = 'Outbound spam policy'; Name = $p.Name; Value = $mode; Meaning = $meaning })
                if ($p.IsDefault -or $p.Name -eq 'Default') { $externalBlocked = $mode -ne 'On' }
                if ($mode -eq 'On') { Add-Finding $(if ($p.IsDefault -or $p.Name -eq 'Default') { 'Medium' } else { 'Low' }) '(organisation)' 'Outbound spam policy' 'External automatic forwarding is allowed' '' "Policy '$($p.Name)' has AutoForwardingMode = On" }
            }
        }
        catch { $errors.Add([pscustomobject]@{ Section = 'Outbound spam policy'; Error = $_.Exception.Message }) }
        try {
            foreach ($d in @(Get-RemoteDomain)) {
                $settings.Add([pscustomobject]@{ Setting = 'Remote domain'; Name = "$($d.Name) ($($d.DomainName))"; Value = "AutoForwardEnabled = $($d.AutoForwardEnabled)"; Meaning = if ($d.AutoForwardEnabled) { 'Forwarding to this domain allowed (if the outbound policy allows it)' } else { 'Forwarding to this domain blocked' } })
            }
        }
        catch { $errors.Add([pscustomobject]@{ Section = 'Remote domains'; Error = $_.Exception.Message }) }

        # --- Export ------------------------------------------------------

        $order = @{ Critical = 0; High = 1; Medium = 2; Low = 3 }
        $findingsSorted = @($findings | Sort-Object @{ Expression = { $order[$_.Risk] } }, Mailbox)
        $counts = [ordered]@{}
        foreach ($r in 'Critical', 'High', 'Medium', 'Low') { $counts[$r] = @($findingsSorted | Where-Object { $_.Risk -eq $r }).Count }
        $org = ''
        try { $org = @($internal | Where-Object { $_ -like '*.onmicrosoft.com' -and $_ -notlike '*.mail.onmicrosoft.com' })[0] } catch { $org = '' }
        if (-not $org) { $org = $internal | Select-Object -First 1 }

        $summary = [ordered]@{
            'Organisation'                 = "$org"
            'Run at'                       = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Mailboxes checked'            = $mailboxes.Count
            'Inbox rules checked'          = if ($SkipInboxRules) { 'Skipped (-SkipInboxRules)' } else { $rulesChecked }
            'Mailboxes forwarding'         = "$(@($mailboxRows | Select-Object -ExpandProperty Mailbox -Unique).Count) ($(@($mailboxRows | Where-Object External).Count) external)"
            'Risky inbox rules'            = $ruleRows.Count
            'Transport rules copying mail' = $transportRows.Count
            'Findings: Critical'           = $counts.Critical
            'Findings: High'               = $counts.High
            'Findings: Medium'             = $counts.Medium
            'Findings: Low'                = $counts.Low
            'External auto-forwarding'     = if ($null -eq $externalBlocked) { 'Unknown' } elseif ($externalBlocked) { 'Blocked by the default outbound spam policy - external forwards found here are failing, but should still be removed' } else { 'ALLOWED by the default outbound spam policy' }
            'Note'                         = 'Some rules created through other mail clients or APIs are hidden from Get-InboxRule. Use -IncludeHiddenRules to include them. For a suspected compromise, also review the sign-in logs and revoke sessions.'
        }

        $saved = Save-Workbook -FilePath $reportPath -Sheets @(
            @{ Name = 'Findings'; Rows = $findingsSorted; Headers = @('Risk', 'Mailbox', 'Source', 'Finding', 'Target', 'Detail'); Widths = @(10, 34, 20, 50, 40, 70) }
            @{ Name = 'Mailbox Forwarding'; Rows = $mailboxRows; Headers = @('Mailbox', 'DisplayName', 'Type', 'ForwardingType', 'ForwardsTo', 'External', 'KeepsCopy'); Widths = @(34, 26, 16, 18, 36, 10, 36) }
            @{ Name = 'Inbox Rules'; Rows = @($ruleRows | Sort-Object @{ Expression = { $order[$_.Risk] } }, Mailbox)
               Headers = @('Risk', 'Mailbox', 'Rule', 'Enabled', 'Actions', 'ForwardsTo', 'Conditions', 'Issues', 'RuleIdentity'); Widths = @(10, 34, 26, 9, 40, 40, 40, 70, 40) }
            @{ Name = 'Transport Rules'; Rows = $transportRows; Headers = @('Rule', 'State', 'Actions', 'Targets', 'External', 'Priority', 'Modified'); Widths = @(36, 10, 30, 50, 10, 9, 18) }
            @{ Name = 'Settings'; Rows = $settings; Headers = @('Setting', 'Name', 'Value', 'Meaning'); Widths = @(22, 36, 26, 60) }
            @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(22, 120) }
            @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(28, 110) }
        )
        if (-not $saved) { exit 1 }

        Write-Host ""
        Write-Host ("Findings: {0} critical, {1} high, {2} medium, {3} low" -f $counts.Critical, $counts.High, $counts.Medium, $counts.Low) -ForegroundColor $(if ($counts.Critical + $counts.High) { 'Red' } elseif ($counts.Medium) { 'Yellow' } else { 'Green' })
        $findingsSorted | Where-Object { $_.Risk -in 'Critical', 'High' } | Select-Object -First 15 | ForEach-Object { Write-Host ("  {0,-8} {1}: {2} {3}" -f $_.Risk, $_.Mailbox, $_.Finding, $(if ($_.Target) { "-> $($_.Target)" })) }
        Write-Host ("External auto-forwarding: {0}" -f $summary['External auto-forwarding'])
        if ($errors.Count) { Write-Host ("Some checks failed - see the 'Errors' sheet ({0})." -f $errors.Count) -ForegroundColor Yellow }
        Write-Host "Report  : $reportPath"

        if ($SendEmail) {
            $important = @($findingsSorted | Where-Object { $_.Risk -ne 'Low' })
            if (-not $important.Count -and -not $AlwaysEmail) {
                Write-Host 'No critical, high or medium findings - no email sent (use -AlwaysEmail to send anyway).'
            }
            else {
                $html = New-HtmlReport -Title 'Mailbox forwarding audit' -IntroLines @("$($mailboxes.Count) mailbox(es) checked: $($counts.Critical) critical, $($counts.High) high, $($counts.Medium) medium.", "External auto-forwarding: $($summary['External auto-forwarding']).") -Sections @(
                    @{ Heading = "Findings ($($important.Count))"; Html = ConvertTo-HtmlTable -Rows $important -Columns @('Risk', 'Mailbox', 'Source', 'Finding', 'Target', 'Detail') -StatusColumn 'Risk' -EmptyText 'None.' }
                )
                $state = if ($counts.Critical) { 'CRITICAL' } elseif ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Medium' } else { 'OK' }
                if (Send-ToolkitMail -Config $config -Subject "[$state] Mailbox forwarding audit - $($important.Count) finding(s)" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
            }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    }
}
