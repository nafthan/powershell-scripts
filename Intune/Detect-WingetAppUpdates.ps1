<#
.SYNOPSIS
    Intune detection script: finds apps in a list that have a winget update available.
.DESCRIPTION
    Exits 1 (run the remediation) if any app in $AppIds is installed and has an update, otherwise 0.
    Pair with Remediate-WingetAppUpdates.ps1 and keep the same $AppIds in both.
.EXAMPLE
    Intune > Devices > Scripts and remediations > Create: detection = this file, remediation = Remediate-WingetAppUpdates.ps1
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Intune : Run as SYSTEM, 64-bit PowerShell. Only machine-wide installs can be updated from SYSTEM.
#>

# winget IDs to keep up to date - find them with: winget search <name>
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
        Add-Content -LiteralPath $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [Detect] $Message"
    } catch { Write-Verbose "Could not write log: $($_.Exception.Message)" }
}

function Get-WingetPath {
    <# winget.exe is not on PATH for SYSTEM - find the newest App Installer package for this CPU. #>
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
    <# $true if the app is installed and winget has a newer version. #>
    $out = & $Winget list --id $Id --exact --upgrade-available --accept-source-agreements --disable-interactivity 2>&1 | Out-String
    return ($out -match "(?im)(^|\s)$([regex]::Escape($Id))(\s|$)")
}

function Invoke-Detection {
    $winget = Get-WingetPath
    if (-not $winget) { Write-IntuneLog 'winget not found'; Write-Output 'winget not found - nothing to do'; return 0 }

    $pending = @($AppIds | Where-Object { Test-UpdateAvailable $winget $_ })
    if ($pending.Count) {
        Write-IntuneLog "Updates available: $($pending -join ', ')"
        Write-Output "Updates available: $($pending -join ', ')"
        return 1
    }
    Write-IntuneLog 'All listed apps are up to date (or not installed)'
    Write-Output 'All listed apps are up to date (or not installed)'
    return 0
}

if ($MyInvocation.InvocationName -ne '.') { exit (Invoke-Detection) }
