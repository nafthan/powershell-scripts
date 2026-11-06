<#
.SYNOPSIS
    Intune detection script: finds the OneStart.ai browser (a potentially unwanted program) for any user.
.DESCRIPTION
    Checks every user profile and loaded user registry hive for OneStart files, uninstall entries,
    startup entries, scheduled tasks, running processes and downloaded installers.
    Exits 1 if anything is found (run the remediation), otherwise 0. Pair with Remediate-OneStart.ps1.
.EXAMPLE
    Intune > Devices > Scripts and remediations > Create: detection = this file, remediation = Remediate-OneStart.ps1
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Intune : Run as SYSTEM, 64-bit PowerShell
#>

$Pattern = 'OneStart'
$LogFile = "$env:ProgramData\IntuneScripts\OneStart.log"

function Write-IntuneLog([string]$Message) {
    try {
        $dir = Split-Path -Parent $LogFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Add-Content -LiteralPath $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [Detect] $Message"
    } catch { Write-Verbose "Could not write log: $($_.Exception.Message)" }
}

function Get-UserProfilePaths {
    # Real user profiles only (skips Public, Default and system profiles)
    Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object { -not $_.Special -and $_.LocalPath } | ForEach-Object { $_.LocalPath }
}

function Get-UserHives {
    # Loaded user hives (signed-in or recently signed-in users)
    Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^S-1-5-21-[\d-]+$' } | ForEach-Object { "Registry::HKEY_USERS\$($_.PSChildName)" }
}

function Find-OneStart {
    $found = [System.Collections.Generic.List[string]]::new()
    foreach ($p in (Get-UserProfilePaths)) {
        foreach ($sub in 'AppData\Local', 'AppData\Roaming') {
            foreach ($d in @(Get-ChildItem -LiteralPath (Join-Path $p $sub) -Directory -Filter "$Pattern*" -Force -ErrorAction SilentlyContinue)) { $found.Add("Folder: $($d.FullName)") }
        }
        foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $p 'Downloads') -File -Filter "$Pattern*" -Force -ErrorAction SilentlyContinue)) { $found.Add("Installer: $($f.FullName)") }
    }
    foreach ($h in (Get-UserHives)) {
        if (Test-Path -LiteralPath "$h\Software\OneStart.ai") { $found.Add("Registry: $h\Software\OneStart.ai") }
        foreach ($k in @(Get-ChildItem -LiteralPath "$h\Software\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue)) {
            if ("$($k.GetValue('DisplayName'))" -like "*$Pattern*") { $found.Add("Uninstall entry: $($k.Name)") }
        }
        $run = "$h\Software\Microsoft\Windows\CurrentVersion\Run"
        if (Test-Path -LiteralPath $run) {
            foreach ($v in @((Get-Item -LiteralPath $run).GetValueNames() | Where-Object { $_ -like "*$Pattern*" })) { $found.Add("Startup entry: $run\$v") }
        }
    }
    foreach ($t in @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "*$Pattern*" })) { $found.Add("Scheduled task: $($t.TaskPath)$($t.TaskName)") }
    foreach ($pr in @(Get-Process -Name "$Pattern*" -ErrorAction SilentlyContinue)) { $found.Add("Process: $($pr.Name) [$($pr.Id)]") }
    return $found
}

function Invoke-Detection {
    $found = @(Find-OneStart)
    if ($found.Count) {
        Write-IntuneLog "Found: $($found -join ' | ')"
        Write-Output "OneStart found ($($found.Count) item(s)): $(($found | Select-Object -First 3) -join '; ')"
        return 1
    }
    Write-Output 'OneStart not found'
    return 0
}

if ($MyInvocation.InvocationName -ne '.') { exit (Invoke-Detection) }
