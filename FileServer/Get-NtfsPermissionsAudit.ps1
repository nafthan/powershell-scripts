<#
.SYNOPSIS
    Audits NTFS folder permissions and flags risky entries.
.DESCRIPTION
    Flags broken inheritance, orphaned SIDs, broad groups with write access, non-admin Full Control and Deny entries.
.EXAMPLE
    .\Get-NtfsPermissionsAudit.ps1 -Path '\\FILESERVER01\Shares\Finance' -MaxDepth 3
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Requires: Windows, rights to read permissions on the folders
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path,

    [ValidateRange(-1, 100)]
    [int]$MaxDepth,

    [switch]$AllEntries,

    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('Path'))     { $Path     = Get-Setting $config 'NtfsAudit' 'Path' }
if (-not $PSBoundParameters.ContainsKey('MaxDepth')) { $MaxDepth = Get-Setting $config 'NtfsAudit' 'MaxDepth' 3 }

if (-not $Path) { throw 'No folder given. Use -Path or set NtfsAudit.Path in config.psd1.' }
if (-not $script:IsWindowsHost) { throw 'This script reads NTFS permissions and must be run on Windows.' }
if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw "Folder not found or not accessible: $Path" }

# --- Permission helpers ---------------------------------------------------

$script:SidCache = @{}

function Resolve-Sid([System.Security.Principal.SecurityIdentifier]$Sid) {
    $key = $Sid.Value
    if (-not $script:SidCache.ContainsKey($key)) {
        try   { $script:SidCache[$key] = $Sid.Translate([System.Security.Principal.NTAccount]).Value }
        catch { $script:SidCache[$key] = $null }
    }
    return $script:SidCache[$key]
}

function Get-FolderAcl([System.IO.DirectoryInfo]$Dir) {
    $sections = [System.Security.AccessControl.AccessControlSections]'Access, Owner'
    if ($Dir.PSObject.Methods.Name -contains 'GetAccessControl') {
        return $Dir.GetAccessControl($sections)                                       # Windows PowerShell 5.1
    }
    return [System.IO.FileSystemAclExtensions]::GetAccessControl($Dir, $sections)     # PowerShell 7
}

function ConvertTo-SpecificRights([long]$Rights) {
    # Generic rights (used on CREATOR OWNER and inherit-only entries) -> file-system equivalents
    if ($Rights -lt 0) { $Rights += 4294967296 }
    $specific = $Rights -band 0x0FFFFFFF
    if ($Rights -band 0x10000000) { $specific = $specific -bor 0x1F01FF }   # GENERIC_ALL
    if ($Rights -band 0x20000000) { $specific = $specific -bor 0x1200A0 }   # GENERIC_EXECUTE
    if ($Rights -band 0x40000000) { $specific = $specific -bor 0x120116 }   # GENERIC_WRITE
    if ($Rights -band 0x80000000) { $specific = $specific -bor 0x120089 }   # GENERIC_READ
    return $specific
}

function Format-Rights([long]$Specific) {
    # Masks leave out Synchronize (0x100000), which Windows doesn't always include
    $full   = 0x0F01FF   # Full Control
    $modify = 0x0301BF   # Modify
    $rx     = 0x0200A9   # Read & Execute / List folder contents
    $read   = 0x020089   # Read
    $write  = 0x000116   # Write (WriteData, AppendData, WriteExtendedAttributes, WriteAttributes)

    if (($Specific -band $full)   -eq $full)   { return 'Full Control' }
    if (($Specific -band $modify) -eq $modify) { return 'Modify' }

    $parts = @()
    if     (($Specific -band $rx)   -eq $rx)   { $parts += 'Read & Execute' }
    elseif (($Specific -band $read) -eq $read) { $parts += 'Read' }
    if (($Specific -band $write) -eq $write)   { $parts += 'Write' }
    if ($parts.Count -gt 0) { return $parts -join ', ' }

    return 'Special: ' + [enum]::ToObject([System.Security.AccessControl.FileSystemRights], [int]$Specific).ToString()
}

function Format-AppliesTo($Rule) {
    $ci = ($Rule.InheritanceFlags -band [System.Security.AccessControl.InheritanceFlags]::ContainerInherit) -ne 0
    $oi = ($Rule.InheritanceFlags -band [System.Security.AccessControl.InheritanceFlags]::ObjectInherit) -ne 0
    $io = ($Rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0
    $np = ($Rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::NoPropagateInherit) -ne 0

    $text = switch ("$ci|$oi|$io") {
        'False|False|False' { 'This folder only' }
        'True|True|False'   { 'This folder, subfolders and files' }
        'True|False|False'  { 'This folder and subfolders' }
        'False|True|False'  { 'This folder and files' }
        'True|True|True'    { 'Subfolders and files only' }
        'True|False|True'   { 'Subfolders only' }
        'False|True|True'   { 'Files only' }
        default             { 'This folder only' }
    }
    if ($np -and ($ci -or $oi)) { $text += ' (one level only)' }
    return $text
}

# Well-known SIDs used for flags
$broadSids = @('S-1-1-0', 'S-1-5-11', 'S-1-5-32-545')       # Everyone, Authenticated Users, BUILTIN\Users
$adminSids = @('S-1-5-18', 'S-1-5-32-544', 'S-1-3-0')       # SYSTEM, Administrators, CREATOR OWNER
$writeMask = 0x2 -bor 0x4 -bor 0x10000 -bor 0x40000 -bor 0x80000   # write/append/delete/change perms/take ownership

function Get-RuleFlags($Rule, [long]$Specific, [string]$Sid, [string]$Name) {
    $flags = @()
    if (-not $Name) { $flags += 'Orphaned SID' }
    if ($Rule.AccessControlType -eq 'Deny') { $flags += 'Deny entry'; return $flags }

    $isBroad = ($broadSids -contains $Sid) -or ($Sid -match '^S-1-5-21-.+-513$')      # ...-513 = Domain Users
    $isAdmin = ($adminSids -contains $Sid) -or ($Sid -match '^S-1-5-21-.+-(512|519)$') # Domain / Enterprise Admins

    if ($isBroad -and ($Specific -band $writeMask)) { $flags += 'Broad group can write' }
    if (-not $isAdmin -and (($Specific -band 0x0F01FF) -eq 0x0F01FF)) { $flags += 'Full Control (non-admin)' }
    return $flags
}

# --- Scan -----------------------------------------------------------------

$root       = Get-NormalizedRoot $Path
$reportPath = New-ReportPath -Config $config -Prefix 'NtfsPermissions' -OutputFile $OutputFile

Write-Host "Auditing : $root"
Write-Host ("Depth    : {0}" -f $(if ($MaxDepth -lt 0) { 'all levels' } else { "$MaxDepth level(s)" }))
Write-Host ""

$rows    = [System.Collections.Generic.List[object]]::new()
$errors  = [System.Collections.Generic.List[object]]::new()
$stack   = [System.Collections.Generic.Stack[object]]::new()
$stack.Push(@{ Dir = (Add-LongPathPrefix $root); Level = 0 })

$foldersChecked   = 0
$brokenInheritance = 0
$startTime = Get-Date
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

while ($stack.Count -gt 0) {
    $node    = $stack.Pop()
    $display = Remove-LongPathPrefix $node.Dir
    $dir     = [System.IO.DirectoryInfo]::new($node.Dir)

    try {
        $acl = Get-FolderAcl $dir
    }
    catch {
        $err = Get-ErrorMessage $_
        $errors.Add([pscustomobject]@{ FullPath = $display; Stage = 'Read permissions'; ErrorType = $err.Type; Error = $err.Message })
        $acl = $null
    }

    if ($acl) {
        $foldersChecked++
        $inheritanceOn = -not $acl.AreAccessRulesProtected
        if (-not $inheritanceOn) { $brokenInheritance++ }

        $owner = ''
        try {
            $ownerSid = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
            if ($ownerSid) { $owner = Resolve-Sid $ownerSid; if (-not $owner) { $owner = $ownerSid.Value } }
        } catch { $owner = '(unreadable)' }

        $listAll = ($node.Level -eq 0) -or $AllEntries -or (-not $inheritanceOn)
        $rules   = $acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])

        foreach ($rule in $rules) {
            if (-not $listAll -and $rule.IsInherited) { continue }

            $sid      = $rule.IdentityReference.Value
            $name     = Resolve-Sid $rule.IdentityReference
            $specific = ConvertTo-SpecificRights ([long][int]$rule.FileSystemRights)
            $flags    = @(Get-RuleFlags $rule $specific $sid $name)

            $rows.Add([pscustomobject]@{
                Folder             = $display
                Level              = $node.Level
                Identity           = if ($name) { $name } else { $sid }
                Access             = "$($rule.AccessControlType)"
                Rights             = Format-Rights $specific
                AppliesTo          = Format-AppliesTo $rule
                Inherited          = [bool]$rule.IsInherited
                InheritanceEnabled = $inheritanceOn
                Owner              = $owner
                Flags              = $flags -join '; '
            })
        }

        if (-not $inheritanceOn -and $node.Level -gt 0) {
            # Make sure the folder appears in Flagged even if none of its entries are risky
            $rows.Add([pscustomobject]@{
                Folder = $display; Level = $node.Level; Identity = '(folder)'; Access = ''; Rights = ''
                AppliesTo = ''; Inherited = $false; InheritanceEnabled = $false; Owner = $owner
                Flags = 'Inheritance disabled'
            })
        }
    }

    # Subfolders
    if ($MaxDepth -lt 0 -or $node.Level -lt $MaxDepth) {
        try {
            foreach ($sub in $dir.GetDirectories()) {
                if (($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                $stack.Push(@{ Dir = $sub.FullName; Level = $node.Level + 1 })
            }
        }
        catch {
            $err = Get-ErrorMessage $_
            $errors.Add([pscustomobject]@{ FullPath = $display; Stage = 'List subfolders'; ErrorType = $err.Type; Error = $err.Message })
        }
    }

    if ($foldersChecked % 500 -eq 0) {
        Write-Progress -Activity 'Auditing NTFS permissions' -Status ("{0:N0} folders checked" -f $foldersChecked) -CurrentOperation $display
    }
}
Write-Progress -Activity 'Auditing NTFS permissions' -Completed
$stopwatch.Stop()

# --- Export ---------------------------------------------------------------

$sorted  = @($rows | Sort-Object -Property Folder, @{ Expression = { $_.Identity -eq '(folder)' }; Descending = $true }, Identity)
$flagged = @($sorted | Where-Object { $_.Flags })

$flagCounts = [ordered]@{}
foreach ($f in @('Inheritance disabled', 'Orphaned SID', 'Broad group can write', 'Full Control (non-admin)', 'Deny entry')) {
    $flagCounts["Flag: $f"] = @($flagged | Where-Object { $_.Flags -like "*$f*" }).Count
}

$summaryItems = [ordered]@{
    'Folder audited'            = $root
    'Scan started'              = $startTime.ToString('yyyy-MM-dd HH:mm:ss')
    'Duration (seconds)'        = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
    'Depth'                     = if ($MaxDepth -lt 0) { 'All levels' } else { $MaxDepth }
    'Folders checked'           = $foldersChecked
    'Folders with inheritance disabled' = $brokenInheritance
    'Permission entries listed' = @($sorted | Where-Object { $_.Identity -ne '(folder)' }).Count
    'Flagged entries'           = $flagged.Count
}
foreach ($k in $flagCounts.Keys) { $summaryItems[$k] = $flagCounts[$k] }
$summaryItems['Folders with errors'] = $errors.Count
$summaryItems['Listing mode'] = if ($AllEntries) { 'All entries on every folder' } else { 'Top folder in full; subfolders show explicit entries only (or all entries where inheritance is disabled)' }
$summaryItems['Scanned by'] = "$env:USERDOMAIN\$env:USERNAME"

$headers = @('Folder', 'Level', 'Identity', 'Access', 'Rights', 'AppliesTo', 'Inherited', 'InheritanceEnabled', 'Owner', 'Flags')
$widths  = @(70, 7, 35, 8, 22, 32, 10, 18, 30, 45)

$saved = Save-Workbook -FilePath $reportPath -Sheets @(
    @{ Name = 'Permissions';   Rows = $sorted;  Headers = $headers; Widths = $widths }
    @{ Name = 'Flagged';       Rows = $flagged; Headers = $headers; Widths = $widths }
    @{ Name = 'Access Errors'; Rows = $errors;  Headers = @('FullPath', 'Stage', 'ErrorType', 'Error'); Widths = @(70, 18, 30, 80) }
    @{ Name = 'Summary';       Rows = (New-SummaryRows $summaryItems); Headers = @('Item', 'Value'); Widths = @(36, 90) }
)
if (-not $saved) { exit 1 }

Write-Host ("Checked {0:N0} folders in {1:N1} seconds." -f $foldersChecked, $stopwatch.Elapsed.TotalSeconds)
Write-Host ("Flagged {0:N0} entries:" -f $flagged.Count) -ForegroundColor $(if ($flagged.Count) { 'Yellow' } else { 'Green' })
foreach ($k in $flagCounts.Keys) { Write-Host ("  {0,-36} {1:N0}" -f ($k -replace '^Flag: ', ''), $flagCounts[$k]) }
if ($errors.Count -gt 0) {
    Write-Host ("Could not read {0:N0} folder(s) - see the 'Access Errors' sheet." -f $errors.Count) -ForegroundColor Red
}
Write-Host ""
Write-Host "Report  : $reportPath"
