<#
.SYNOPSIS
    Exports non-compliant devices from Intune and Microsoft Defender for Endpoint, with the reasons why.
.DESCRIPTION
    Lists Intune devices that are not compliant, with the policies and settings they fail, then adds what
    Defender knows about each one: onboarding, sensor health, exposure level, active alerts and antivirus
    state. Every device is also listed on an All Devices sheet. If Defender data is unavailable, those columns
    are left blank. Read-only.
.EXAMPLE
    .\Get-NonCompliantDeviceReport.ps1
.EXAMPLE
    .\Get-NonCompliantDeviceReport.ps1 -Platform Windows -SendEmail
.EXAMPLE
    .\Get-NonCompliantDeviceReport.ps1 -SkipDefender      # Intune only, e.g. without Defender for Endpoint
.EXAMPLE
    .\Get-NonCompliantDeviceReport.ps1 -AntivirusForAllDevices      # antivirus state for every Windows device, not just flagged ones
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: Microsoft.Graph.Authentication module, Global Reader (or Intune Administrator) plus Security Reader
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

    [string[]]$Platform,   # Windows, macOS, iOS, Android, Linux
    [switch]$IncludeUnknown,
    [switch]$SkipDefender,
    [switch]$SkipSettingDetail,
    [switch]$AntivirusForAllDevices,
    [ValidateRange(1, 3650)] [int]$StaleDays,

    [switch]$SendEmail,
    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('StaleDays')) { $StaleDays = [int](Get-Setting $config 'Devices' 'StaleSyncDays' 30) }
$Platform = @(Split-ListValue $Platform)
$bad = @($Platform | Where-Object { $_ -notin 'Windows', 'macOS', 'iOS', 'Android', 'Linux' })
if ($bad.Count) { throw "Unknown platform: $($bad -join ', '). Use Windows, macOS, iOS, Android or Linux." }

$scopes = @('DeviceManagementManagedDevices.Read.All', 'DeviceManagementConfiguration.Read.All')
if (-not $SkipDefender) { $scopes += 'ThreatHunting.Read.All', 'SecurityAlert.Read.All' }

# Intune compliance states that count as "not compliant"
$badStates = @('noncompliant', 'inGracePeriod', 'conflict', 'error')
$stateNames = @{
    compliant = 'Compliant'; noncompliant = 'Not compliant'; inGracePeriod = 'In grace period'; conflict = 'Conflict'
    error = 'Error'; unknown = 'Not evaluated'; configManager = 'Managed by ConfigMgr'
}

# Friendly names for common compliance settings (others are split from their internal name)
$settingNames = @{
    # Built-in checks from the default compliance policy
    RequireRemainContact                  = 'Device has not checked in within the compliance validity period'
    RequireDeviceCompliancePolicyAssigned = 'No compliance policy assigned'
    RequireUserExistence                  = 'Enrolled user no longer exists'
    # Windows
    BitLockerEnabled                      = 'BitLocker'
    SecureBootEnabled                     = 'Secure Boot'
    CodeIntegrityEnabled                  = 'Code integrity'
    StorageRequireEncryption              = 'Encryption of data storage'
    TpmRequired                           = 'TPM'
    ActiveFirewallRequired                = 'Firewall'
    AntivirusRequired                     = 'Antivirus'
    AntiSpywareRequired                   = 'Antispyware'
    DefenderEnabled                       = 'Microsoft Defender Antimalware'
    RtpEnabled                            = 'Real-time protection'
    SignatureOutOfDate                    = 'Defender security intelligence up to date'
    DeviceThreatProtectionEnabled         = 'Defender for Endpoint machine risk score'
    DeviceThreatProtectionRequiredSecurityLevel = 'Defender for Endpoint machine risk score'
    OsMinimumVersion                      = 'Minimum OS version'
    OsMaximumVersion                      = 'Maximum OS version'
    PasswordRequired                      = 'Password required'
    # Mobile
    PasscodeRequired                      = 'Passcode required'
    SecurityBlockJailbrokenDevices        = 'Jailbroken / rooted devices blocked'
}

# Defender sensor states that mean it isn't reporting properly
$healthySensor = @('Active', '')

# --- Helpers ---------------------------------------------------------------------

function Get-SettingLabel([string]$SettingName) {
    <# "Windows10CompliancePolicy.BitLockerEnabled" -> "BitLocker Enabled" #>
    $leaf = ($SettingName -split '\.')[-1]
    if ($settingNames.ContainsKey($leaf)) { return $settingNames[$leaf] }
    if (-not $leaf) { return $SettingName }
    return ([regex]::Replace($leaf, '(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])', ' ')).Trim()
}

function Get-PlatformName([string]$Os) {
    switch -Wildcard ($Os) {
        'Windows*' { 'Windows' } 'mac*' { 'macOS' } 'iOS*' { 'iOS' } 'iPadOS*' { 'iOS' } 'Android*' { 'Android' } 'Linux*' { 'Linux' } default { $Os }
    }
}

function Get-ShortName([string]$Name) {
    <# Defender device names are FQDNs; Intune names aren't. #>
    return (($Name -split '\.')[0]).ToLowerInvariant()
}

function Invoke-HuntingQuery([string]$Query) {
    <# Runs an advanced hunting query through Graph and returns the result rows. #>
    $body = @{ Query = $Query } | ConvertTo-Json -Compress
    for ($attempt = 1; ; $attempt++) {
        try {
            $res = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/security/runHuntingQuery' -Body $body -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
            return @($res.results)
        }
        catch {
            if ($attempt -lt 4 -and $_.Exception.Message -match '429|TooManyRequests|503|504') { Start-Sleep -Seconds (10 * $attempt); continue }
            # The useful part of a BadRequest (e.g. an unknown column) is in the response body
            $detail = "$($_.ErrorDetails.Message)"
            if ($detail -match '"message"\s*:\s*"([^"]+)"') { $detail = $Matches[1] }
            throw "$($_.Exception.Message)$(if ($detail) { " - $detail" })"
        }
    }
}

function Get-FirstValue($Object, [string[]]$Names) {
    <# Hunting results can come back with different key casing - read whichever exists. #>
    foreach ($n in $Names) { $p = $Object.PSObject.Properties[$n]; if ($p -and "$($p.Value)" -ne '') { return $p.Value } }
    return $null
}

# --- Main ------------------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    $connectedHere = Connect-ToolkitGraph -Config $config -Scopes $scopes -Bound $PSBoundParameters -RoleHint 'Global Reader (or Intune Administrator) plus Security Reader'
    try {
        $now        = Get-Date
        $reportPath = New-ReportPath -Config $config -Prefix 'NonCompliantDevices' -OutputFile $OutputFile
        $errors     = [System.Collections.Generic.List[object]]::new()
        function Add-ReportError([string]$Section, [string]$Message) { $errors.Add([pscustomobject]@{ Section = $Section; Error = $Message }) }
        $tenantName = Get-GraphTenantName
        Write-Host "Tenant : $tenantName"

        # --- 1. Intune devices -------------------------------------------------------

        Write-Host 'Getting Intune devices...'
        $select = 'id,deviceName,userPrincipalName,userDisplayName,operatingSystem,osVersion,complianceState,lastSyncDateTime,enrolledDateTime,' +
                  'model,manufacturer,serialNumber,managedDeviceOwnerType,isEncrypted,azureADDeviceId,complianceGracePeriodExpirationDateTime,managementAgent,jailBroken'
        $allDevices = @()
        try { $allDevices = @(Get-GraphAll "v1.0/deviceManagement/managedDevices?`$select=$select&`$top=999") }
        catch { throw "Could not read Intune devices (needs an Intune licence and DeviceManagementManagedDevices.Read.All): $($_.Exception.Message)" }
        $everyDevice = $allDevices   # unfiltered, to tell which Defender devices Intune manages
        if ($Platform.Count) { $allDevices = @($allDevices | Where-Object { $Platform -contains (Get-PlatformName $_.operatingSystem) }) }
        $inScope = if ($IncludeUnknown) { $badStates + 'unknown' } else { $badStates }
        $flagged = @($allDevices | Where-Object { $inScope -contains "$($_.complianceState)" })
        Write-Host ("  {0} device(s), {1} not compliant" -f $allDevices.Count, $flagged.Count)

        # --- 2. Why each device fails ------------------------------------------------

        $settingRows = [System.Collections.Generic.List[object]]::new()
        $reasons = @{}   # device id -> @{ Policies = list; Settings = list }
        if (-not $SkipSettingDetail) {
            $i = 0
            foreach ($d in $flagged) {
                $i++
                Write-Progress -Activity 'Reading failing compliance policies' -Status "$i of $($flagged.Count): $($d.deviceName)" -PercentComplete ($i / [math]::Max(1, $flagged.Count) * 100)
                $entry = @{ Policies = [System.Collections.Generic.List[string]]::new(); Settings = [System.Collections.Generic.List[string]]::new() }
                $reasons[$d.id] = $entry
                try {
                    $policies = @(Get-GraphAll "v1.0/deviceManagement/managedDevices/$($d.id)/deviceCompliancePolicyStates" |
                                  Where-Object { "$($_.state)" -in 'nonCompliant', 'error', 'conflict' })
                    foreach ($p in $policies) {
                        $policyName = if ($p.displayName) { $p.displayName } else { 'Default device compliance policy' }
                        if (-not $entry.Policies.Contains($policyName)) { $entry.Policies.Add($policyName) }
                        $settings = @(Get-GraphAll "v1.0/deviceManagement/managedDevices/$($d.id)/deviceCompliancePolicyStates/$($p.id)/settingStates" |
                                      Where-Object { "$($_.state)" -in 'nonCompliant', 'error', 'conflict' })
                        foreach ($s in $settings) {
                            $label = Get-SettingLabel "$($s.settingName)$(if (-not $s.settingName) { $s.setting })"
                            if (-not $entry.Settings.Contains($label)) { $entry.Settings.Add($label) }
                            $settingRows.Add([pscustomobject]@{
                                Device = $d.deviceName; User = $d.userPrincipalName; Platform = Get-PlatformName $d.operatingSystem
                                Policy = $policyName; Setting = $label; State = "$($s.state)"; SettingId = "$($s.settingName)"
                            })
                        }
                    }
                }
                catch { Add-ReportError 'Compliance detail' "$($d.deviceName): $($_.Exception.Message)" }
            }
            Write-Progress -Activity 'Reading failing compliance policies' -Completed
        }

        # --- 3. Defender -------------------------------------------------------------

        $defender = @()
        $defByAad = @{}; $defByName = @{}
        $alertsByDevice = @{}
        $alertRows = [System.Collections.Generic.List[object]]::new()
        $defenderAvailable = $false
        if (-not $SkipDefender) {
            Write-Host 'Getting Defender for Endpoint devices...'
            # The full query first; if the tenant's schema doesn't have every column, a simpler one
            $queries = @(
                @'
DeviceInfo
| where Timestamp > ago(30d) and (isempty(DeviceCategory) or DeviceCategory == "Endpoint")
| summarize arg_max(Timestamp, *) by DeviceId
| where isempty(MergedToDeviceId)
| project DeviceId, DeviceName, AadDeviceId, OSPlatform, OSVersion, OnboardingStatus, SensorHealthState, ExposureLevel, ClientVersion, MachineGroup, LastSeen = Timestamp
'@
                @'
DeviceInfo
| where Timestamp > ago(30d)
| summarize arg_max(Timestamp, DeviceName, AadDeviceId, OSPlatform, OSVersion, OnboardingStatus, SensorHealthState, ExposureLevel) by DeviceId
| project DeviceId, DeviceName, AadDeviceId, OSPlatform, OSVersion, OnboardingStatus, SensorHealthState, ExposureLevel, LastSeen = Timestamp
'@
                @'
DeviceInfo
| where Timestamp > ago(30d)
| summarize arg_max(Timestamp, DeviceName, OSPlatform, OnboardingStatus) by DeviceId
| project DeviceId, DeviceName, OSPlatform, OnboardingStatus, LastSeen = Timestamp
'@
            )
            try {
                $huntRows = $null
                $failures = @()
                foreach ($q in $queries) {
                    try { $huntRows = @(Invoke-HuntingQuery $q); break }
                    catch { $failures += $_.Exception.Message }
                }
                if ($null -eq $huntRows) { throw ($failures | Select-Object -Last 1) }
                if ($failures.Count) { Add-ReportError 'Defender devices' "Full Defender query failed, so a simpler one was used (some Defender columns may be blank): $($failures[0])" }
                $defender = @($huntRows | ForEach-Object {
                    [pscustomobject]@{
                        DeviceId      = Get-FirstValue $_ 'DeviceId', 'deviceId'
                        DeviceName    = Get-FirstValue $_ 'DeviceName', 'deviceName'
                        AadDeviceId   = "$(Get-FirstValue $_ 'AadDeviceId', 'aadDeviceId')".ToLowerInvariant()
                        OSPlatform    = Get-FirstValue $_ 'OSPlatform', 'osPlatform'
                        OSVersion     = Get-FirstValue $_ 'OSVersion', 'osVersion'
                        Onboarding    = "$(Get-FirstValue $_ 'OnboardingStatus', 'onboardingStatus')"
                        SensorHealth  = "$(Get-FirstValue $_ 'SensorHealthState', 'sensorHealthState')"
                        Exposure      = "$(Get-FirstValue $_ 'ExposureLevel', 'exposureLevel')"
                        ClientVersion = Get-FirstValue $_ 'ClientVersion', 'clientVersion'
                        MachineGroup  = Get-FirstValue $_ 'MachineGroup', 'machineGroup'
                        LastSeen      = ConvertTo-LocalTime (Get-FirstValue $_ 'LastSeen', 'lastSeen')
                    }
                })
                $defenderAvailable = $true
                foreach ($m in $defender) {
                    if ($m.AadDeviceId -and $m.AadDeviceId -ne '00000000-0000-0000-0000-000000000000') { $defByAad[$m.AadDeviceId] = $m }
                    if ($m.DeviceName) { $defByName[(Get-ShortName $m.DeviceName)] = $m }
                }
                Write-Host "  $($defender.Count) device(s) seen in the last 30 days"
            }
            catch {
                Add-ReportError 'Defender devices' "Advanced hunting unavailable, so Defender columns are blank (needs Defender for Endpoint, ThreatHunting.Read.All and Security Reader): $($_.Exception.Message)"
                Write-Warning 'Defender device data is unavailable - see the Errors sheet. Continuing without it.'
            }

            Write-Host 'Getting active Defender alerts...'
            try {
                $alerts = @(Get-GraphAll "v1.0/security/alerts_v2?`$filter=serviceSource eq 'microsoftDefenderForEndpoint' and status ne 'resolved'&`$top=2000")
                foreach ($a in $alerts) {
                    $devices = @($a.evidence | Where-Object { "$($_.'@odata.type')" -like '*deviceEvidence' })
                    foreach ($ev in $devices) {
                        $keys = @("$($ev.mdeDeviceId)", "aad:$("$($ev.azureAdDeviceId)".ToLowerInvariant())", "name:$(Get-ShortName "$($ev.deviceDnsName)")") | Where-Object { $_ -and $_ -notmatch ':$' }
                        foreach ($k in $keys) {
                            if (-not $alertsByDevice.ContainsKey($k)) { $alertsByDevice[$k] = [System.Collections.Generic.List[object]]::new() }
                            if (-not ($alertsByDevice[$k] | Where-Object { $_.id -eq $a.id })) { $alertsByDevice[$k].Add($a) }
                        }
                        $alertRows.Add([pscustomobject]@{
                            Severity = (Get-Culture).TextInfo.ToTitleCase("$($a.severity)"); Device = $ev.deviceDnsName; Alert = $a.title; Category = $a.category
                            Status = "$($a.status)"; Created = ConvertTo-LocalTime $a.createdDateTime; Link = $a.alertWebUrl
                        })
                    }
                }
                Write-Host "  $($alerts.Count) active alert(s)"
            }
            catch { Add-ReportError 'Defender alerts' "Alerts unavailable (needs SecurityAlert.Read.All and Security Reader): $($_.Exception.Message)" }
        }

        function Find-DefenderDevice($IntuneDevice) {
            $aad = "$($IntuneDevice.azureADDeviceId)".ToLowerInvariant()
            if ($aad -and $defByAad.ContainsKey($aad)) { return $defByAad[$aad] }
            $short = Get-ShortName $IntuneDevice.deviceName
            if ($defByName.ContainsKey($short)) { return $defByName[$short] }
            return $null
        }
        function Get-DeviceAlerts($Machine, $IntuneDevice) {
            $found = @()
            $keys = @()
            if ($Machine) { $keys += "$($Machine.DeviceId)"; $keys += "name:$(Get-ShortName $Machine.DeviceName)" }
            if ($IntuneDevice) { $keys += "aad:$("$($IntuneDevice.azureADDeviceId)".ToLowerInvariant())"; $keys += "name:$(Get-ShortName $IntuneDevice.deviceName)" }
            foreach ($k in $keys) { if ($alertsByDevice.ContainsKey($k)) { $found += @($alertsByDevice[$k]) } }
            return @($found | Sort-Object id -Unique)
        }
        function Get-DefenderIssues($Machine, [object[]]$Alerts) {
            $out = @()
            if ($Machine) {
                # "Unsupported" and "Insufficient info" are left out - they are usually phones, printers and the like
                if ($Machine.Onboarding -eq 'Can be onboarded') { $out += 'Defender: not onboarded (can be onboarded)' }
                elseif ($Machine.Onboarding -eq 'Onboarded' -and $healthySensor -notcontains $Machine.SensorHealth) { $out += "Defender sensor: $($Machine.SensorHealth)" }
                if ($Machine.Exposure -eq 'High') { $out += 'Defender exposure level High' }
            }
            $hi = @($Alerts | Where-Object { "$($_.severity)" -eq 'high' }).Count
            if ($Alerts.Count) { $out += "$($Alerts.Count) active Defender alert(s)$(if ($hi) { " ($hi high)" })" }
            return $out
        }

        # --- 4. Devices with Defender problems that Intune says are fine -----------

        $extraIntune = @()
        if ($defenderAvailable -or $alertsByDevice.Count) {
            $flaggedIds = @{}; foreach ($d in $flagged) { $flaggedIds[$d.id] = $true }
            $extraIntune = @($allDevices | Where-Object { -not $flaggedIds.ContainsKey($_.id) } | Where-Object {
                $m = Find-DefenderDevice $_
                @(Get-DefenderIssues $m (Get-DeviceAlerts $m $_)).Count -gt 0
            })
        }

        # --- 5. Antivirus state for Windows devices in the report -----------------

        $avById = @{}
        $avScope = if ($AntivirusForAllDevices) { $allDevices } else { @($flagged) + @($extraIntune) }
        $windowsInReport = @($avScope | Where-Object { (Get-PlatformName $_.operatingSystem) -eq 'Windows' })
        if ($windowsInReport.Count) {
            $i = 0
            foreach ($d in $windowsInReport) {
                $i++
                Write-Progress -Activity 'Reading antivirus state' -Status "$i of $($windowsInReport.Count): $($d.deviceName)" -PercentComplete ($i / $windowsInReport.Count * 100)
                try { $avById[$d.id] = Invoke-Graph "v1.0/deviceManagement/managedDevices/$($d.id)/windowsProtectionState" }
                catch { if ($_.Exception.Message -notmatch '404|NotFound') { Add-ReportError 'Antivirus state' "$($d.deviceName): $($_.Exception.Message)" } }
            }
            Write-Progress -Activity 'Reading antivirus state' -Completed
        }

        # --- 6. Build rows -----------------------------------------------------------

        Write-Host 'Building report...'
        $rank = @{ OK = 0; Low = 1; Medium = 2; High = 3 }
        $matchedDefender = @{}
        $attentionIds = @{}; foreach ($d in @(@($flagged) + @($extraIntune))) { $attentionIds[$d.id] = $true }
        $allRows = foreach ($d in $allDevices) {
            $m       = Find-DefenderDevice $d
            if ($m) { $matchedDefender[$m.DeviceId] = $true }
            $alerts  = @(Get-DeviceAlerts $m $d)
            $av      = $avById[$d.id]
            $why     = $reasons[$d.id]
            $state   = "$($d.complianceState)"
            $lastSync = ConvertTo-LocalTime $d.lastSyncDateTime
            $syncDays = if ($lastSync) { [int][math]::Floor(($now - $lastSync).TotalDays) } else { $null }
            $devPlatform = Get-PlatformName $d.operatingSystem

            $issues = [System.Collections.Generic.List[string]]::new()
            $current = @{ Risk = 'OK' }
            $raise = { param($Level, $Text) $issues.Add($Text); if ($rank[$Level] -gt $rank[$current.Risk]) { $current.Risk = $Level } }

            switch ($state) {
                'noncompliant'  { & $raise 'Medium' 'Not compliant in Intune' }
                'inGracePeriod' { & $raise 'Low' "In grace period until $((ConvertTo-LocalTime $d.complianceGracePeriodExpirationDateTime).ToString('yyyy-MM-dd HH:mm'))" }
                'conflict'      { & $raise 'Low' 'Compliance policy conflict' }
                'error'         { & $raise 'Low' 'Compliance evaluation error' }
                'unknown'       { & $raise 'Low' 'Compliance not evaluated' }
            }
            if ($state -ne 'compliant' -and $d.isEncrypted -eq $false -and $devPlatform -in 'Windows', 'macOS') { & $raise 'High' 'Not compliant and not encrypted' }
            if ("$($d.jailBroken)" -eq 'True') { & $raise 'High' 'Jailbroken / rooted' }
            if ($null -ne $syncDays -and $syncDays -gt $StaleDays) { & $raise 'Low' "No Intune check-in for $syncDays days" }

            if ($av) {
                if ($av.malwareProtectionEnabled -eq $false) { & $raise 'High' 'Antivirus off' }
                elseif ($av.realTimeProtectionEnabled -eq $false) { & $raise 'High' 'Real-time protection off' }
                if ($av.signatureUpdateOverdue -eq $true) { & $raise 'Medium' 'Antivirus signatures out of date' }
                if ($av.tamperProtectionEnabled -eq $false) { & $raise 'Medium' 'Tamper protection off' }
                if ($av.quickScanOverdue -eq $true -or $av.fullScanOverdue -eq $true) { & $raise 'Low' 'Antivirus scan overdue' }
            }

            if ($defenderAvailable -and $devPlatform -eq 'Windows' -and -not $m) { & $raise 'Medium' 'Not seen in Defender in the last 30 days' }
            foreach ($t in (Get-DefenderIssues $m $alerts)) {
                & $raise $(if ($t -match 'high\)|sensor|not onboarded') { 'High' } else { 'Medium' }) $t
            }

            [pscustomobject]@{
                Risk              = $current.Risk
                Device            = $d.deviceName
                User              = $d.userPrincipalName
                UserName          = $d.userDisplayName
                Platform          = $devPlatform
                OSVersion         = $d.osVersion
                IntuneCompliance  = if ($stateNames.ContainsKey($state)) { $stateNames[$state] } else { $state }
                FailingPolicies   = if ($why) { ($why.Policies -join '; ') } elseif ($SkipSettingDetail) { '(not read)' } else { '' }
                FailingSettings   = if ($why) { ($why.Settings -join '; ') } else { '' }
                Issues            = ($issues -join '; ')
                LastSync          = $lastSync
                DaysSinceSync     = $syncDays
                GraceExpires      = if ($state -eq 'inGracePeriod') { ConvertTo-LocalTime $d.complianceGracePeriodExpirationDateTime } else { $null }
                Encrypted         = $d.isEncrypted
                DefenderStatus    = if (-not $defenderAvailable) { 'Not checked' } elseif ($m) { $m.Onboarding } else { 'Not found' }
                SensorHealth      = if ($m) { $m.SensorHealth } else { '' }
                ExposureLevel     = if ($m) { $m.Exposure } else { '' }
                ActiveAlerts      = $alerts.Count
                RealTimeProtection = if ($av) { [bool]$av.realTimeProtectionEnabled } else { '' }
                SignaturesOverdue = if ($av) { [bool]$av.signatureUpdateOverdue } else { '' }
                TamperProtection  = if ($av) { [bool]$av.tamperProtectionEnabled } else { '' }
                SignatureVersion  = if ($av) { $av.signatureVersion } else { '' }
                Ownership         = "$($d.managedDeviceOwnerType)"
                Manufacturer      = $d.manufacturer
                Model             = $d.model
                SerialNumber      = $d.serialNumber
                Enrolled          = ConvertTo-LocalTime $d.enrolledDateTime
                IntuneDeviceId    = $d.id
                EntraDeviceId     = $d.azureADDeviceId
                DefenderDeviceId  = if ($m) { $m.DeviceId } else { '' }
            }
        }
        $allRows = @($allRows | Sort-Object @{ Expression = { $rank[$_.Risk] }; Descending = $true }, Device)
        $rows    = @($allRows | Where-Object { $attentionIds.ContainsKey($_.IntuneDeviceId) })

        # Defender devices with problems that aren't in Intune at all
        $defOnly = @()
        if ($defenderAvailable) {
            $intuneAad = @{}; $intuneNames = @{}
            foreach ($d in $everyDevice) { if ($d.azureADDeviceId) { $intuneAad["$($d.azureADDeviceId)".ToLowerInvariant()] = $true }; $intuneNames[(Get-ShortName $d.deviceName)] = $true }
            $defOnly = @($defender | Where-Object { -not $Platform.Count -or $Platform -contains (Get-PlatformName $_.OSPlatform) } | Where-Object { -not $matchedDefender.ContainsKey($_.DeviceId) -and -not $intuneAad.ContainsKey($_.AadDeviceId) -and -not $intuneNames.ContainsKey((Get-ShortName $_.DeviceName)) } |
                ForEach-Object {
                    $alerts = @(Get-DeviceAlerts $_ $null)
                    $iss = @(Get-DefenderIssues $_ $alerts)
                    if (-not $iss.Count) { return }
                    [pscustomobject]@{
                        Device = $_.DeviceName; Platform = $_.OSPlatform; OSVersion = $_.OSVersion; Onboarding = $_.Onboarding; SensorHealth = $_.SensorHealth
                        ExposureLevel = $_.Exposure; ActiveAlerts = $alerts.Count; Issues = ($iss -join '; '); LastSeen = $_.LastSeen
                        DeviceGroup = $_.MachineGroup; SensorVersion = $_.ClientVersion; DefenderDeviceId = $_.DeviceId
                    }
                } | Sort-Object Device)
        }

        $bySetting = @($settingRows | Group-Object Platform, Policy, Setting | ForEach-Object {
            $g = $_.Group
            [pscustomobject]@{ Setting = $g[0].Setting; Policy = $g[0].Policy; Platform = $g[0].Platform; Devices = @($g.Device | Select-Object -Unique).Count }
        } | Sort-Object Devices -Descending)

        # --- 7. Summary and export ---------------------------------------------------

        $counts = @{}; foreach ($k in 'High', 'Medium', 'Low', 'OK') { $counts[$k] = @($rows | Where-Object Risk -eq $k).Count }
        $stateCounts = $allDevices | Group-Object complianceState | ForEach-Object { "$(if ($stateNames.ContainsKey($_.Name)) { $stateNames[$_.Name] } else { $_.Name }): $($_.Count)" }
        $compliantPct = if ($allDevices.Count) { [math]::Round(@($allDevices | Where-Object complianceState -eq 'compliant').Count / $allDevices.Count * 100, 1) } else { 0 }
        $summary = [ordered]@{
            'Tenant'                         = $tenantName
            'Run at'                         = $now.ToString('yyyy-MM-dd HH:mm:ss')
            'Platforms'                      = if ($Platform.Count) { $Platform -join ', ' } else { 'All' }
            'Intune devices'                 = $allDevices.Count
            'Compliant'                      = "$compliantPct%"
            'By compliance state'            = ($stateCounts -join ', ')
            'Not compliant in Intune'        = $flagged.Count
            'Compliant, but Defender issues' = $extraIntune.Count
            'Defender-only devices with issues' = if ($defenderAvailable) { $defOnly.Count } else { 'Not checked' }
            'Risk'                           = "$($counts.High) high, $($counts.Medium) medium, $($counts.Low) low"
            'Most common failing setting'    = if ($bySetting.Count) { "$($bySetting[0].Setting) ($($bySetting[0].Devices) device(s))" } else { '' }
            'Note'                           = "Defender data covers devices seen in the last 30 days. Devices are matched by Entra device ID, then by name. Stale = no Intune check-in for $StaleDays+ days."
            'Run by'                         = Get-GraphRunBy
        }

        $deviceHeaders = @('Risk', 'Device', 'User', 'Platform', 'OSVersion', 'IntuneCompliance', 'FailingPolicies', 'FailingSettings', 'Issues', 'LastSync', 'DaysSinceSync', 'GraceExpires',
                           'Encrypted', 'DefenderStatus', 'SensorHealth', 'ExposureLevel', 'ActiveAlerts', 'RealTimeProtection', 'SignaturesOverdue', 'TamperProtection', 'SignatureVersion',
                           'Ownership', 'Manufacturer', 'Model', 'SerialNumber', 'Enrolled', 'UserName', 'IntuneDeviceId', 'EntraDeviceId', 'DefenderDeviceId')
        $deviceWidths  = @(9, 24, 32, 10, 16, 16, 40, 55, 70, 18, 14, 18, 10, 15, 16, 14, 12, 18, 17, 16, 16, 11, 16, 24, 18, 18, 24, 38, 38, 42)
        $sheets = @(
            @{ Name = 'Devices'; Rows = $rows; Headers = $deviceHeaders; Widths = $deviceWidths }
            @{ Name = 'All Devices'; Rows = $allRows; Headers = $deviceHeaders; Widths = $deviceWidths }
            @{ Name = 'Failing Settings'; Rows = @($settingRows | Sort-Object Device, Policy, Setting)
               Headers = @('Device', 'User', 'Platform', 'Policy', 'Setting', 'State', 'SettingId'); Widths = @(24, 32, 10, 40, 50, 13, 60) }
            @{ Name = 'By Setting'; Rows = $bySetting; Headers = @('Setting', 'Policy', 'Platform', 'Devices'); Widths = @(55, 40, 10, 9) }
        )
        if (-not $SkipDefender) {
            $sheets += @{ Name = 'Defender Only'; Rows = $defOnly
                          Headers = @('Device', 'Platform', 'OSVersion', 'Onboarding', 'SensorHealth', 'ExposureLevel', 'ActiveAlerts', 'Issues', 'LastSeen', 'DeviceGroup', 'SensorVersion', 'DefenderDeviceId')
                          Widths  = @(34, 12, 14, 16, 22, 14, 12, 60, 18, 20, 16, 42) }
            $sheets += @{ Name = 'Active Alerts'; Rows = @($alertRows | Sort-Object @{ Expression = { @{ High = 0; Medium = 1; Low = 2 }["$($_.Severity)"] } }, Created)
                          Headers = @('Severity', 'Device', 'Alert', 'Category', 'Status', 'Created', 'Link'); Widths = @(10, 30, 55, 20, 12, 18, 60) }
        }
        $sheets += @{ Name = 'Errors'; Rows = $errors; Headers = @('Section', 'Error'); Widths = @(22, 120) }
        $sheets += @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(34, 110) }
        if (-not (Save-Workbook -FilePath $reportPath -Sheets $sheets)) { exit 1 }

        # --- Console and email -------------------------------------------------------

        Write-Host ""
        Write-Host ("Compliant : {0}% of {1} Intune device(s)" -f $compliantPct, $allDevices.Count) -ForegroundColor $(if ($compliantPct -ge 95) { 'Green' } elseif ($compliantPct -ge 85) { 'Yellow' } else { 'Red' })
        Write-Host ("In report : {0} device(s) - {1} high, {2} medium, {3} low" -f $rows.Count, $counts.High, $counts.Medium, $counts.Low)
        if ($defenderAvailable) { Write-Host ("Defender-only devices with issues: {0}" -f $defOnly.Count) }
        foreach ($s in $bySetting | Select-Object -First 5) { Write-Host ("  {0,4}  {1} ({2})" -f $s.Devices, $s.Setting, $s.Platform) }
        if ($errors.Count) { Write-Host "Some data was unavailable - see the 'Errors' sheet ($($errors.Count))." -ForegroundColor Yellow }
        Write-Host "Report    : $reportPath"

        if ($SendEmail) {
            $top = @($rows | Where-Object { $_.Risk -in 'High', 'Medium' } | Select-Object -First 50)
            $html = New-HtmlReport -Title 'Non-compliant devices' -IntroLines @(
                "$tenantName - $compliantPct% of $($allDevices.Count) Intune devices are compliant. $($rows.Count) device(s) need attention: $($counts.High) high, $($counts.Medium) medium, $($counts.Low) low."
            ) -Sections @(
                @{ Heading = "High and medium risk ($(@($rows | Where-Object { $_.Risk -in 'High', 'Medium' }).Count))"
                   Html    = ConvertTo-HtmlTable -Rows $top -Columns @('Risk', 'Device', 'User', 'Platform', 'IntuneCompliance', 'Issues', 'LastSync') -StatusColumn 'Risk' -EmptyText 'None.' }
                @{ Heading = 'Most common failing settings'
                   Html    = ConvertTo-HtmlTable -Rows @($bySetting | Select-Object -First 10) -Columns @('Setting', 'Policy', 'Platform', 'Devices') -EmptyText 'None recorded.' }
            )
            $tag = if ($counts.High) { 'HIGH' } elseif ($counts.Medium) { 'Warning' } else { 'OK' }
            if (Send-ToolkitMail -Config $config -Subject "[$tag] Devices: $compliantPct% compliant, $($rows.Count) need attention" -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
        }
    }
    finally {
        if ($connectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
}
