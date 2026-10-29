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
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
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

#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)

$script:ToolkitRoot   = $PSScriptRoot

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)

$script:IsWindowsHost = ($PSVersionTable.PSEdition -eq 'Desktop') -or $IsWindows

$script:HtmlFont = 'font-family:Segoe UI,Arial,sans-serif;font-size:13px;'

function Get-ToolkitConfig {
    # Standalone: no config file. Settings come from parameters and built-in defaults.
    $cfg = @{ General = @{ OutputFolder = (Get-Location).Path } }
    if ($SmtpServer) {
        $cfg.Email = @{
            SmtpServer = $SmtpServer; Port = $SmtpPort; UseSsl = [bool]$SmtpUseSsl; From = $MailFrom
            To = @($MailTo | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
    }
    
    return $cfg
}

function Get-Setting {
    param([hashtable]$Config, [string]$Section, [string]$Key, $Default = $null)
    $value = $null
    if ($Config -and $Config.ContainsKey($Section) -and $Config[$Section] -is [hashtable] -and $Config[$Section].ContainsKey($Key)) {
        $value = $Config[$Section][$Key]
    }
    if ($null -eq $value -or ($value -is [string] -and $value -eq '') -or ($value -is [array] -and $value.Count -eq 0)) { $value = $Default }
    return $value
}

function Resolve-ToolkitPath([string]$PathValue) {
    <# Turns a relative path from the config into a full path under the toolkit root. #>
    if (-not $PathValue) { return $PathValue }
    if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
    return [System.IO.Path]::GetFullPath((Join-Path $script:ToolkitRoot $PathValue))
}

function New-ReportPath {
    <#
        Returns the output file path. Uses -OutputFile if given (fixing the extension),
        otherwise <OutputFolder>\<Prefix>_<timestamp><Extension>. Creates the folder.
    #>
    param(
        [hashtable]$Config,
        [string]$Prefix,
        [string]$Extension = '.xlsx',
        [string]$OutputFile
    )
    if ($OutputFile) {
        if ([System.IO.Path]::GetExtension($OutputFile) -ne $Extension) {
            $OutputFile = [System.IO.Path]::ChangeExtension($OutputFile, $Extension)
        }
        $folder = Split-Path -Path $OutputFile -Parent
        $file   = $OutputFile
    }
    else {
        $folder = Resolve-ToolkitPath (Get-Setting $Config 'General' 'OutputFolder' 'Reports')
        $file   = Join-Path $folder ("{0}_{1}{2}" -f $Prefix, (Get-Date -Format 'yyyyMMdd_HHmmss'), $Extension)
    }
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force -WhatIf:$false -Confirm:$false | Out-Null
    }
    return $file
}

function Add-LongPathPrefix([string]$p) {
    if (-not $script:IsWindowsHost) { return $p }
    if ($p.StartsWith('\\?\')) { return $p }
    if ($p.StartsWith('\\'))   { return '\\?\UNC\' + $p.Substring(2) }
    return '\\?\' + $p
}

function Remove-LongPathPrefix([string]$p) {
    if ($p.StartsWith('\\?\UNC\')) { return '\\' + $p.Substring(8) }
    if ($p.StartsWith('\\?\'))     { return $p.Substring(4) }
    return $p
}

function Get-NormalizedRoot([string]$Path) {
    <# Full path of a folder without a trailing separator (except drive roots like C:\). #>
    $full = (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName
    if ($full.Length -gt 1) { $full = $full.TrimEnd('\', '/') }
    if ($full -match '^[A-Za-z]:$') { $full += '\' }
    if ($full -eq '') { $full = '/' }
    return $full
}

function Get-ErrorMessage($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    while ($ex.InnerException) { $ex = $ex.InnerException }
    return [pscustomobject]@{ Type = $ex.GetType().Name; Message = $ex.Message }
}

function ConvertTo-XmlText([string]$s) {
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[^\u0009\u000A\u000D\u0020-\uFFFD]', '')
    return [System.Security.SecurityElement]::Escape($s)
}

function Get-ColumnLetter([int]$Index) {
    # 0 -> A, 25 -> Z, 26 -> AA
    $name = ''
    $n = $Index + 1
    while ($n -gt 0) {
        $rem  = ($n - 1) % 26
        $name = [char](65 + $rem) + $name
        $n    = [math]::Floor(($n - 1) / 26)
    }
    return $name
}

function Add-ZipText($Zip, [string]$Name, [string]$Text) {
    $entry  = $Zip.CreateEntry($Name)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Add-XlsxSheet {
    param($Zip, [int]$Index, [string[]]$Headers, $Rows, [int[]]$Widths)

    $inv     = [System.Globalization.CultureInfo]::InvariantCulture
    $letters = for ($i = 0; $i -lt $Headers.Count; $i++) { Get-ColumnLetter $i }
    $entry   = $Zip.CreateEntry("xl/worksheets/sheet$Index.xml")
    $w       = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try {
        $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        $w.Write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        $w.Write('<cols>')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $width = if ($Widths -and $i -lt $Widths.Count) { $Widths[$i] } else { 20 }
            $w.Write(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $width))
        }
        $w.Write('</cols><sheetData>')

        $w.Write('<row r="1">')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $w.Write(('<c r="{0}1" s="1" t="inlineStr"><is><t>{1}</t></is></c>' -f $letters[$i], (ConvertTo-XmlText $Headers[$i])))
        }
        $w.Write('</row>')

        $r = 1
        foreach ($row in $Rows) {
            $r++
            $w.Write("<row r=`"$r`">")
            for ($i = 0; $i -lt $Headers.Count; $i++) {
                $ref = "$($letters[$i])$r"
                $val = $row.($Headers[$i])
                if ($null -eq $val) {
                    continue
                }
                elseif ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal] -or $val -is [single]) {
                    $w.Write("<c r=`"$ref`"><v>$(([double]$val).ToString('R', $inv))</v></c>")
                }
                elseif ($val -is [datetime]) {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t>$($val.ToString('yyyy-MM-dd HH:mm', $inv))</t></is></c>")
                }
                elseif ($val -is [bool]) {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t>$(if ($val) { 'Yes' } else { 'No' })</t></is></c>")
                }
                else {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText "$val")</t></is></c>")
                }
            }
            $w.Write('</row>')
        }
        $w.Write('</sheetData>')
        $w.Write("<autoFilter ref=`"A1:$($letters[-1])$r`"/>")
        $w.Write('</worksheet>')
    }
    finally { $w.Dispose() }
}

function Export-Xlsx {
    <#
        Writes a workbook. -Sheets is an array of hashtables:
            @{ Name = 'Sheet name'; Headers = @('Col1','Col2'); Rows = $objects; Widths = @(40, 12) }
        Row values are read by header name. Numbers stay numeric; dates are written as yyyy-MM-dd HH:mm.
    #>
    param([string]$FilePath, [object[]]$Sheets)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -LiteralPath $FilePath) { Remove-Item -LiteralPath $FilePath -Force -ErrorAction Stop -WhatIf:$false -Confirm:$false }
    $zip = [System.IO.Compression.ZipFile]::Open($FilePath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $overrides = ''; $entries = ''; $rels = ''
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $n    = $i + 1
            $name = $Sheets[$i].Name -replace '[\[\]\*\?/\\:]', '-'
            if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }
            $overrides += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
            $entries   += "<sheet name=`"$(ConvertTo-XmlText $name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
            $rels      += "<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
        }
        $stylesId = $Sheets.Count + 1

        Add-ZipText $zip '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
            $overrides + '</Types>')

        Add-ZipText $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>')

        Add-ZipText $zip 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheets>' + $entries + '</sheets></workbook>')

        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $rels +
            "<Relationship Id=`"rId$stylesId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>" +
            '</Relationships>')

        Add-ZipText $zip 'xl/styles.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
            '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor indexed="64"/></patternFill></fill></fills>' +
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>' +
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
            '</styleSheet>')

        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $s = $Sheets[$i]
            Add-XlsxSheet -Zip $zip -Index ($i + 1) -Headers $s.Headers -Rows $s.Rows -Widths $s.Widths
        }
    }
    finally { $zip.Dispose() }
}

function Save-Workbook {
    <# Export-Xlsx with a friendly error if the file is open in Excel. #>
    param([string]$FilePath, [object[]]$Sheets)
    try {
        Export-Xlsx -FilePath $FilePath -Sheets $Sheets
        return $true
    }
    catch {
        Write-Error ("Could not save {0}. If it is open in Excel, close it and run again. Details: {1}" -f $FilePath, $_.Exception.Message)
        return $false
    }
}

function New-SummaryRows([System.Collections.IDictionary]$Items) {
    <# Turns an ordered hashtable into Item/Value rows for a Summary sheet. #>
    foreach ($k in $Items.Keys) { [pscustomobject]@{ Item = $k; Value = $Items[$k] } }
}

#endregion
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('Path'))     { $Path     = Get-Setting $config 'NtfsAudit' 'Path' }
if (-not $PSBoundParameters.ContainsKey('MaxDepth')) { $MaxDepth = Get-Setting $config 'NtfsAudit' 'MaxDepth' 3 }

if (-not $Path) { throw 'No folder given. Use -Path.' }
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
