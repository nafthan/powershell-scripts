<#
.SYNOPSIS
    Intune detection script: finds manufacturer (OEM) bloatware on Dell, HP and Lenovo devices.
.DESCRIPTION
    Picks the app list for this device's manufacturer, then checks installed programs (registry uninstall
    entries, not Win32_Product) and Store apps for anything on it. Exits 1 if anything is found, otherwise 0.
    Pair with Remediate-OemBloatware.ps1 and keep the same lists in both.
.EXAMPLE
    Intune > Devices > Scripts and remediations > Create: detection = this file, remediation = Remediate-OemBloatware.ps1
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Intune : Run as SYSTEM, 64-bit PowerShell. Test on a pilot group first.
#>

# Wildcard names to remove, by manufacturer. Keep driver/firmware tools you rely on (e.g. Dell Command | Update) off these lists.
$Bloatware = @{
    Dell   = @('Dell SupportAssist*', 'Dell Digital Delivery*', 'Dell Optimizer*', 'Dell Peripheral Manager', 'Dell Core Services', 'Dell Pair', 'Dell Display Manager*', 'DellInc.DellSupportAssistforPCs', 'DellInc.DellDigitalDelivery', 'DellInc.PartnerPromo')
    HP     = @('HP Wolf Security*', 'HP Sure Click*', 'HP Sure Run*', 'HP Sure Recover*', 'HP Security Update Service', 'HP Connection Optimizer', 'HP Documentation', 'HP Notifications', 'HP Support Assistant', 'AD2F1837.HPJumpStarts', 'AD2F1837.HPPrivacySettings', 'AD2F1837.HPSupportAssistant', 'AD2F1837.myHP', 'AD2F1837.HPQuickDrop')
    Lenovo = @('Lenovo Vantage*', 'Lenovo Welcome', 'Lenovo Now', 'E046963F.LenovoCompanion', 'E046963F.LenovoSettingsforEnterprise', 'LenovoCorporation.LenovoWelcome')
}

$LogFile = "$env:ProgramData\IntuneScripts\OemBloatware.log"

function Write-IntuneLog([string]$Message) {
    try {
        $dir = Split-Path -Parent $LogFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Add-Content -LiteralPath $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [Detect] $Message"
    } catch { Write-Verbose "Could not write log: $($_.Exception.Message)" }
}

function Get-OemPatterns {
    $make = "$((Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).Manufacturer)"
    switch -Wildcard ($make) {
        'Dell*'    { return $Bloatware.Dell }
        'HP*'      { return $Bloatware.HP }
        'Hewlett*' { return $Bloatware.HP }
        'Lenovo*'  { return $Bloatware.Lenovo }
    }
    return @()
}

function Test-NameMatch([string]$Name, [string[]]$Patterns) {
    foreach ($p in $Patterns) { if ($Name -like $p) { return $true } }
    return $false
}

function Get-InstalledPrograms {
    $keys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    Get-ItemProperty -Path $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and -not $_.SystemComponent }
}

function Find-Bloatware([string[]]$Patterns) {
    $found = [System.Collections.Generic.List[string]]::new()
    foreach ($a in @(Get-InstalledPrograms)) { if (Test-NameMatch $a.DisplayName $Patterns) { $found.Add("Program: $($a.DisplayName)") } }
    foreach ($a in @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue)) { if (Test-NameMatch $a.Name $Patterns) { $found.Add("Store app: $($a.Name)") } }
    foreach ($a in @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue)) { if (Test-NameMatch $a.DisplayName $Patterns) { $found.Add("Provisioned: $($a.DisplayName)") } }
    return $found
}

function Invoke-Detection {
    $patterns = @(Get-OemPatterns)
    if (-not $patterns.Count) { Write-Output 'Manufacturer not on the list - nothing to check'; return 0 }
    $found = @(Find-Bloatware $patterns | Select-Object -Unique)
    if ($found.Count) {
        Write-IntuneLog "Found: $($found -join ' | ')"
        Write-Output "Bloatware found ($($found.Count)): $(($found | Select-Object -First 3) -join '; ')"
        return 1
    }
    Write-Output 'No bloatware found'
    return 0
}

if ($MyInvocation.InvocationName -ne '.') { exit (Invoke-Detection) }
