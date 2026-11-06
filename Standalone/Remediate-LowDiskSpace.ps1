<#
.SYNOPSIS
    Intune remediation script: frees space on the system drive by clearing safe-to-delete files.
.DESCRIPTION
    Clears Windows and user temp files older than $TempDays days, the Recycle Bin, the Delivery Optimization
    cache and queued Windows Error Reporting files, then reports how much space was freed. Never touches
    Downloads, Documents or other user data. Exits 0 if the drive is now above the threshold, otherwise 1.
    Pair with Detect-LowDiskSpace.ps1.
.EXAMPLE
    Intune > Devices > Scripts and remediations > Create: detection = Detect-LowDiskSpace.ps1, remediation = this file
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Intune : Run as SYSTEM, 64-bit PowerShell. Log: C:\ProgramData\IntuneScripts\LowDiskSpace.log
#>

# Keep the thresholds the same as in Detect-LowDiskSpace.ps1
$MinFreePercent = 10
$MinFreeGB      = 15
$TempDays       = 7

$LogFile = "$env:ProgramData\IntuneScripts\LowDiskSpace.log"

function Write-IntuneLog([string]$Message) {
    try {
        $dir = Split-Path -Parent $LogFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if ((Test-Path -LiteralPath $LogFile) -and (Get-Item -LiteralPath $LogFile).Length -gt 1MB) { Move-Item -LiteralPath $LogFile -Destination "$LogFile.old" -Force }
        Add-Content -LiteralPath $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [Remediate] $Message"
    } catch { Write-Verbose "Could not write log: $($_.Exception.Message)" }
}

function Get-SystemDriveSpace {
    $d = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'" -ErrorAction Stop
    [pscustomobject]@{
        Drive       = $d.DeviceID
        SizeGB      = [math]::Round($d.Size / 1GB, 1)
        FreeBytes   = [double]$d.FreeSpace
        FreeGB      = [math]::Round($d.FreeSpace / 1GB, 1)
        FreePercent = if ($d.Size) { [math]::Round(($d.FreeSpace / $d.Size) * 100, 1) } else { 0 }
    }
}

function Clear-OldFiles([string]$Path, [int]$Days) {
    <# Deletes files older than $Days under $Path, then empty folders. Files in use are skipped. #>
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $cutoff = (Get-Date).AddDays(-$Days)
    $n = 0
    foreach ($f in @(Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff })) {
        try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $n++ } catch { Write-Verbose "In use: $($f.FullName)" }
    }
    Get-ChildItem -LiteralPath $Path -Directory -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object { $_.FullName.Length } -Descending |
        Where-Object { -not (Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue) } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    Write-IntuneLog "Cleared $n file(s) from $Path"
}

function Invoke-Remediation {
    try { $before = Get-SystemDriveSpace } catch { Write-Output "Could not read disk: $($_.Exception.Message)"; return 1 }
    Write-IntuneLog "Before: $($before.FreeGB) GB free ($($before.FreePercent)%)"

    # Temp files: Windows and every user profile
    Clear-OldFiles "$env:SystemRoot\Temp" $TempDays
    foreach ($p in @(Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object { -not $_.Special -and $_.LocalPath })) {
        Clear-OldFiles (Join-Path $p.LocalPath 'AppData\Local\Temp') $TempDays
    }

    # Windows Error Reporting queue
    Clear-OldFiles "$env:ProgramData\Microsoft\Windows\WER\ReportQueue" 0
    Clear-OldFiles "$env:ProgramData\Microsoft\Windows\WER\ReportArchive" 0

    # Delivery Optimization cache (Windows re-downloads from Microsoft if needed)
    if (Get-Command Delete-DeliveryOptimizationCache -ErrorAction SilentlyContinue) {
        try { Delete-DeliveryOptimizationCache -Force -ErrorAction Stop; Write-IntuneLog 'Cleared Delivery Optimization cache' }
        catch { Write-IntuneLog "Delivery Optimization cache: $($_.Exception.Message)" }
    }

    # Recycle Bin for all users
    try { Clear-RecycleBin -DriveLetter $env:SystemDrive.TrimEnd(':') -Force -ErrorAction Stop; Write-IntuneLog 'Emptied Recycle Bin' }
    catch { Write-IntuneLog "Recycle Bin: $($_.Exception.Message)" }

    $after = Get-SystemDriveSpace
    $freed = [math]::Round(($after.FreeBytes - $before.FreeBytes) / 1GB, 2)
    $text = "Freed $freed GB. $($after.Drive) now $($after.FreeGB) GB free ($($after.FreePercent)%)"
    Write-IntuneLog $text
    if ($after.FreePercent -lt $MinFreePercent -or $after.FreeGB -lt $MinFreeGB) { Write-Output "Still low - $text"; return 1 }
    Write-Output $text
    return 0
}

if ($MyInvocation.InvocationName -ne '.') { exit (Invoke-Remediation) }
