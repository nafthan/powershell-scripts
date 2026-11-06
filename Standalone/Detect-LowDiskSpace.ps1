<#
.SYNOPSIS
    Intune detection script: flags devices where the system drive is low on free space.
.DESCRIPTION
    Exits 1 (run the remediation) if free space on the system drive is below $MinFreePercent
    or $MinFreeGB, otherwise 0. Pair with Remediate-LowDiskSpace.ps1.
.EXAMPLE
    Intune > Devices > Scripts and remediations > Create: detection = this file, remediation = Remediate-LowDiskSpace.ps1
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
    Intune : Run as SYSTEM, 64-bit PowerShell
#>

$MinFreePercent = 10
$MinFreeGB      = 15

function Get-SystemDriveSpace {
    $d = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'" -ErrorAction Stop
    [pscustomobject]@{
        Drive       = $d.DeviceID
        SizeGB      = [math]::Round($d.Size / 1GB, 1)
        FreeGB      = [math]::Round($d.FreeSpace / 1GB, 1)
        FreePercent = if ($d.Size) { [math]::Round(($d.FreeSpace / $d.Size) * 100, 1) } else { 0 }
    }
}

function Invoke-Detection {
    try { $s = Get-SystemDriveSpace } catch { Write-Output "Could not read disk: $($_.Exception.Message)"; return 0 }
    $text = "$($s.Drive) $($s.FreeGB) GB free of $($s.SizeGB) GB ($($s.FreePercent)%)"
    if ($s.FreePercent -lt $MinFreePercent -or $s.FreeGB -lt $MinFreeGB) { Write-Output "Low disk space: $text"; return 1 }
    Write-Output "OK: $text"
    return 0
}

if ($MyInvocation.InvocationName -ne '.') { exit (Invoke-Detection) }
