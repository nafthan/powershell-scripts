<#
.SYNOPSIS
    Intune remediation script: updates apps in a list with winget.
.DESCRIPTION
    Updates each app in $AppIds that is installed and has an update, silently. Exits 0 if all updates
    succeeded, 1 if any failed (see the log). Pair with Detect-WingetAppUpdates.ps1.
.EXAMPLE
    Intune > Devices > Scripts and remediations > Create: detection = Detect-WingetAppUpdates.ps1, remediation = this file
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Intune : Run as SYSTEM, 64-bit PowerShell. Log: C:\ProgramData\IntuneScripts\WingetAppUpdates.log
#>

# Keep this list the same as in Detect-WingetAppUpdates.ps1
$AppIds = @(
    'Google.Chrome.EXE'
    'Mozilla.Firefox'
    'Dell.CommandUpdate.Universal'
    'Dell.DisplayManager'
)

$LogFile = "$env:ProgramData\IntuneScripts\WingetAppUpdates.log"

function Write-IntuneLog([string]$Message) {
    try {
        $dir = Split-Path -Parent $LogFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if ((Test-Path -LiteralPath $LogFile) -and (Get-Item -LiteralPath $LogFile).Length -gt 1MB) { Move-Item -LiteralPath $LogFile -Destination "$LogFile.old" -Force }
        Add-Content -LiteralPath $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [Remediate] $Message"
    } catch { Write-Verbose "Could not write log: $($_.Exception.Message)" }
}

function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
    $dirs = @(Get-ChildItem -Path "$env:ProgramFiles\WindowsApps" -Directory -Filter "Microsoft.DesktopAppInstaller_*_${arch}__8wekyb3d8bbwe" -ErrorAction SilentlyContinue |
              Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'winget.exe') } |
              Sort-Object { try { [version](($_.Name -split '_')[1]) } catch { [version]'0.0' } } -Descending)
    if ($dirs.Count) { return (Join-Path $dirs[0].FullName 'winget.exe') }
    return $null
}

function Test-UpdateAvailable([string]$Winget, [string]$Id) {
    $out = & $Winget list --id $Id --exact --upgrade-available --accept-source-agreements --disable-interactivity 2>&1 | Out-String
    return ($out -match "(?im)(^|\s)$([regex]::Escape($Id))(\s|$)")
}

function Invoke-Remediation {
    $winget = Get-WingetPath
    if (-not $winget) { Write-IntuneLog 'winget not found'; Write-Output 'winget not found'; return 1 }

    $failed = @(); $updated = @()
    foreach ($id in $AppIds) {
        if (-not (Test-UpdateAvailable $winget $id)) { continue }
        Write-IntuneLog "Updating $id"
        $out = & $winget upgrade --id $id --exact --silent --accept-source-agreements --accept-package-agreements --disable-interactivity 2>&1 | Out-String
        $code = $LASTEXITCODE
        Write-IntuneLog "$id exit code $code ($('0x{0:X8}' -f $code))`n$($out.Trim())"
        if ($code -eq 0) { $updated += $id } else { $failed += "$id ($('0x{0:X8}' -f $code))" }
    }

    if ($failed.Count) { Write-Output "Updated: $($updated -join ', '). Failed: $($failed -join ', ')"; return 1 }
    Write-Output $(if ($updated.Count) { "Updated: $($updated -join ', ')" } else { 'Nothing to update' })
    return 0
}

if ($MyInvocation.InvocationName -ne '.') { exit (Invoke-Remediation) }
