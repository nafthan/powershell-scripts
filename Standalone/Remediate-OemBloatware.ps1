<#
.SYNOPSIS
    Intune remediation script: removes manufacturer (OEM) bloatware from Dell, HP and Lenovo devices.
.DESCRIPTION
    Picks the app list for this device's manufacturer, then silently uninstalls matching programs
    (QuietUninstallString, or msiexec /x for MSI installs), removes matching Store apps for all users and
    de-provisions them so new users don't get them. Exits 0 if nothing is left, otherwise 1.
    Pair with Detect-OemBloatware.ps1.
.EXAMPLE
    Intune > Devices > Scripts and remediations > Create: detection = Detect-OemBloatware.ps1, remediation = this file
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Intune : Run as SYSTEM, 64-bit PowerShell. Log: C:\ProgramData\IntuneScripts\OemBloatware.log
#>

# Keep these lists the same as in Detect-OemBloatware.ps1
$Bloatware = @{
    Dell   = @('Dell SupportAssist*', 'Dell Digital Delivery*', 'Dell Optimizer*', 'Dell Peripheral Manager', 'Dell Core Services', 'Dell Pair', 'Dell Display Manager*', 'DellInc.DellSupportAssistforPCs', 'DellInc.DellDigitalDelivery', 'DellInc.PartnerPromo')
    HP     = @('HP Wolf Security*', 'HP Sure Click*', 'HP Sure Run*', 'HP Sure Recover*', 'HP Security Update Service', 'HP Connection Optimizer', 'HP Documentation', 'HP Notifications', 'HP Support Assistant', 'AD2F1837.HPJumpStarts', 'AD2F1837.HPPrivacySettings', 'AD2F1837.HPSupportAssistant', 'AD2F1837.myHP', 'AD2F1837.HPQuickDrop')
    Lenovo = @('Lenovo Vantage*', 'Lenovo Welcome', 'Lenovo Now', 'E046963F.LenovoCompanion', 'E046963F.LenovoSettingsforEnterprise', 'LenovoCorporation.LenovoWelcome')
}

$TimeoutMinutes = 10
$LogFile = "$env:ProgramData\IntuneScripts\OemBloatware.log"

function Write-IntuneLog([string]$Message) {
    try {
        $dir = Split-Path -Parent $LogFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if ((Test-Path -LiteralPath $LogFile) -and (Get-Item -LiteralPath $LogFile).Length -gt 1MB) { Move-Item -LiteralPath $LogFile -Destination "$LogFile.old" -Force }
        Add-Content -LiteralPath $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [Remediate] $Message"
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

function Get-UninstallCommand($App) {
    <# Returns @(file, arguments) for a silent uninstall, or $null if there is no safe silent option. #>
    if ($App.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$' -and ($App.WindowsInstaller -eq 1 -or "$($App.UninstallString)" -match 'msiexec')) {
        return @('msiexec.exe', "/x $($App.PSChildName) /qn /norestart")
    }
    $cmd = "$($App.QuietUninstallString)"
    if (-not $cmd) { return $null }
    if ($cmd -match '^\s*"([^"]+)"\s*(.*)$') { return @($Matches[1], $Matches[2]) }
    if ($cmd -match '^\s*(\S+\.exe)\s*(.*)$') { return @($Matches[1], $Matches[2]) }
    return $null
}

function Invoke-Uninstall($App) {
    $cmd = Get-UninstallCommand $App
    if (-not $cmd) { Write-IntuneLog "SKIPPED $($App.DisplayName): no silent uninstall command"; return $false }
    Write-IntuneLog "Uninstalling $($App.DisplayName): $($cmd[0]) $($cmd[1])"
    try {
        $p = Start-Process -FilePath $cmd[0] -ArgumentList $cmd[1] -PassThru -WindowStyle Hidden -ErrorAction Stop
        if (-not $p.WaitForExit($TimeoutMinutes * 60000)) { Write-IntuneLog "TIMEOUT $($App.DisplayName) after $TimeoutMinutes min"; return $false }
        # 0 = OK, 1605 = not installed, 1641/3010 = OK, restart needed
        $ok = $p.ExitCode -in 0, 1605, 1641, 3010
        Write-IntuneLog "$($App.DisplayName) exit code $($p.ExitCode)"
        return $ok
    } catch { Write-IntuneLog "FAILED $($App.DisplayName): $($_.Exception.Message)"; return $false }
}

function Invoke-Remediation {
    $patterns = @(Get-OemPatterns)
    if (-not $patterns.Count) { Write-Output 'Manufacturer not on the list - nothing to do'; return 0 }

    # Programs - one pass, MSI and EXE installers
    foreach ($a in @(Get-InstalledPrograms | Where-Object { Test-NameMatch $_.DisplayName $patterns })) { [void](Invoke-Uninstall $a) }

    # Store apps for every user, then the provisioned copy so new profiles don't get them
    foreach ($a in @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { Test-NameMatch $_.Name $patterns })) {
        try { Remove-AppxPackage -Package $a.PackageFullName -AllUsers -ErrorAction Stop; Write-IntuneLog "Removed Store app $($a.Name)" }
        catch { Write-IntuneLog "FAILED Store app $($a.Name): $($_.Exception.Message)" }
    }
    foreach ($a in @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { Test-NameMatch $_.DisplayName $patterns })) {
        try { Remove-AppxProvisionedPackage -Online -PackageName $a.PackageName -ErrorAction Stop | Out-Null; Write-IntuneLog "De-provisioned $($a.DisplayName)" }
        catch { Write-IntuneLog "FAILED de-provision $($a.DisplayName): $($_.Exception.Message)" }
    }

    # Check what is left
    $left = @(Get-InstalledPrograms | Where-Object { Test-NameMatch $_.DisplayName $patterns } | ForEach-Object { $_.DisplayName })
    $left += @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { Test-NameMatch $_.Name $patterns } | ForEach-Object { $_.Name })
    if ($left.Count) { Write-IntuneLog "Still installed: $($left -join ' | ')"; Write-Output "Still installed: $(($left | Select-Object -Unique) -join ', ')"; return 1 }
    Write-Output 'Bloatware removed'
    return 0
}

if ($MyInvocation.InvocationName -ne '.') { exit (Invoke-Remediation) }
