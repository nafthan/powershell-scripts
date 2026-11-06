<#
.SYNOPSIS
    Intune remediation script: removes the OneStart.ai browser (a potentially unwanted program) for every user.
.DESCRIPTION
    Stops OneStart processes running from user profiles, then removes its scheduled tasks, startup entries,
    registry keys, uninstall entries, program folders and downloaded installers. Exits 0 if nothing is left,
    otherwise 1. Pair with Detect-OneStart.ps1.
.EXAMPLE
    Intune > Devices > Scripts and remediations > Create: detection = Detect-OneStart.ps1, remediation = this file
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Intune : Run as SYSTEM, 64-bit PowerShell. Log: C:\ProgramData\IntuneScripts\OneStart.log
#>

$Pattern = 'OneStart'
$LogFile = "$env:ProgramData\IntuneScripts\OneStart.log"

function Write-IntuneLog([string]$Message) {
    try {
        $dir = Split-Path -Parent $LogFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Add-Content -LiteralPath $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [Remediate] $Message"
    } catch { Write-Verbose "Could not write log: $($_.Exception.Message)" }
}

function Get-UserProfilePaths {
    Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object { -not $_.Special -and $_.LocalPath } | ForEach-Object { $_.LocalPath }
}

function Get-UserHives {
    Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^S-1-5-21-[\d-]+$' } | ForEach-Object { "Registry::HKEY_USERS\$($_.PSChildName)" }
}

function Remove-Safely([string]$What, [scriptblock]$Action) {
    try { & $Action; Write-IntuneLog "Removed $What" }
    catch { Write-IntuneLog "FAILED to remove ${What}: $($_.Exception.Message)" }
}

function Invoke-Remediation {
    $profiles = @(Get-UserProfilePaths)

    # 1. Stop processes - only ones running from a user profile, so nothing else with a similar name is touched
    foreach ($pr in @(Get-Process -Name "$Pattern*" -ErrorAction SilentlyContinue)) {
        $path = "$($pr.Path)"
        if ($path -and ($profiles | Where-Object { $path -like (Join-Path $_ '*') })) { Remove-Safely "process $($pr.Name) [$($pr.Id)]" { Stop-Process -Id $pr.Id -Force -ErrorAction Stop } }
        else { Write-IntuneLog "Left process $($pr.Name) [$($pr.Id)] running - not in a user profile ($path)" }
    }
    Start-Sleep -Seconds 2

    # 2. Scheduled tasks
    foreach ($t in @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "*$Pattern*" })) {
        Remove-Safely "scheduled task $($t.TaskPath)$($t.TaskName)" { Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction Stop }
    }

    # 3. Registry: startup entries, app key and uninstall entries in every loaded user hive
    foreach ($h in (Get-UserHives)) {
        $run = "$h\Software\Microsoft\Windows\CurrentVersion\Run"
        if (Test-Path -LiteralPath $run) {
            foreach ($v in @((Get-Item -LiteralPath $run).GetValueNames() | Where-Object { $_ -like "*$Pattern*" })) {
                Remove-Safely "startup entry $run\$v" { Remove-ItemProperty -LiteralPath $run -Name $v -Force -ErrorAction Stop }
            }
        }
        if (Test-Path -LiteralPath "$h\Software\OneStart.ai") { Remove-Safely "$h\Software\OneStart.ai" { Remove-Item -LiteralPath "$h\Software\OneStart.ai" -Recurse -Force -ErrorAction Stop } }
        foreach ($k in @(Get-ChildItem -LiteralPath "$h\Software\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue)) {
            if ("$($k.GetValue('DisplayName'))" -like "*$Pattern*") { Remove-Safely "uninstall entry $($k.Name)" { Remove-Item -LiteralPath $k.PSPath -Recurse -Force -ErrorAction Stop } }
        }
    }

    # 4. Program folders and downloaded installers in every profile
    foreach ($p in $profiles) {
        foreach ($sub in 'AppData\Local', 'AppData\Roaming') {
            foreach ($d in @(Get-ChildItem -LiteralPath (Join-Path $p $sub) -Directory -Filter "$Pattern*" -Force -ErrorAction SilentlyContinue)) {
                Remove-Safely "folder $($d.FullName)" { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction Stop }
            }
        }
        foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $p 'Downloads') -File -Filter "$Pattern*" -Force -ErrorAction SilentlyContinue)) {
            Remove-Safely "installer $($f.FullName)" { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop }
        }
    }

    # 5. Check
    $left = @()
    foreach ($p in $profiles) {
        foreach ($sub in 'AppData\Local', 'AppData\Roaming') { $left += @(Get-ChildItem -LiteralPath (Join-Path $p $sub) -Directory -Filter "$Pattern*" -Force -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }) }
    }
    $left += @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "*$Pattern*" } | ForEach-Object { $_.TaskName })
    if ($left.Count) { Write-IntuneLog "Still present: $($left -join ' | ')"; Write-Output "OneStart partly removed - still present: $(($left | Select-Object -First 3) -join '; ')"; return 1 }
    Write-Output 'OneStart removed'
    return 0
}

if ($MyInvocation.InvocationName -ne '.') { exit (Invoke-Remediation) }
