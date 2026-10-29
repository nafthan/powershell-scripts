<#
.SYNOPSIS
    Reports folder sizes under a path to Excel.
.DESCRIPTION
    Totals each folder down to -Depth levels, with file counts and last-changed dates.
.EXAMPLE
    .\Get-FolderSizeReport.ps1 -Path '\\FILESERVER01\Shares' -Depth 2
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.
.LINK
    https://github.com/nafthan/powershell-scripts
.LINK
    https://nafthan.co.uk
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path,

    [ValidateRange(1, 20)]
    [int]$Depth,

    [ValidateRange(0, 1000000)]
    [int]$Top,

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

if (-not $PSBoundParameters.ContainsKey('Path'))  { $Path  = Get-Setting $config 'FolderSize' 'Path' }
if (-not $PSBoundParameters.ContainsKey('Depth')) { $Depth = Get-Setting $config 'FolderSize' 'Depth' 2 }
if (-not $PSBoundParameters.ContainsKey('Top'))   { $Top   = Get-Setting $config 'FolderSize' 'Top' 0 }

if (-not $Path) { throw 'No folder given. Use -Path.' }
if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw "Folder not found or not accessible: $Path" }

$root       = Get-NormalizedRoot $Path
$reportPath = New-ReportPath -Config $config -Prefix 'FolderSizes' -OutputFile $OutputFile

Write-Host "Scanning : $root"
Write-Host "Depth    : $Depth level(s)"
Write-Host ""

# --- Scan -----------------------------------------------------------------

function New-FolderStat([string]$FolderPath, [int]$Level) {
    [pscustomobject]@{
        Folder     = $FolderPath
        Level      = $Level
        SizeBytes  = [long]0
        Files      = [long]0
        Subfolders = [long]0
        Newest     = [datetime]::MinValue
    }
}

$rootStat = New-FolderStat $root 0
$stats    = [System.Collections.Generic.List[object]]::new()
$errors   = [System.Collections.Generic.List[object]]::new()
$stack    = [System.Collections.Generic.Stack[object]]::new()

# Each stack entry carries the report folders it belongs to, so a file's size is
# added to its own report folder and every report folder above it.
$stack.Push(@{ Dir = (Add-LongPathPrefix $root); Level = 0; Owners = @($rootStat) })

$scanned   = 0
$startTime = Get-Date
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

while ($stack.Count -gt 0) {
    $node = $stack.Pop()
    try {
        $items = ([System.IO.DirectoryInfo]::new($node.Dir)).GetFileSystemInfos()
    }
    catch {
        $err = Get-ErrorMessage $_
        $failed = Remove-LongPathPrefix $node.Dir
        $errors.Add([pscustomobject]@{ FullPath = $failed; ErrorType = $err.Type; Error = $err.Message })
        continue
    }

    foreach ($item in $items) {
        $scanned++
        $isFolder = ($item.Attributes -band [System.IO.FileAttributes]::Directory) -ne 0

        if ($isFolder) {
            foreach ($o in $node.Owners) { $o.Subfolders++ }
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }

            $level  = $node.Level + 1
            $owners = $node.Owners
            if ($level -le $Depth) {
                $s = New-FolderStat (Remove-LongPathPrefix $item.FullName) $level
                $stats.Add($s)
                $owners = @($node.Owners) + $s
            }
            $stack.Push(@{ Dir = $item.FullName; Level = $level; Owners = $owners })
        }
        else {
            $len = $item.Length
            $lw  = $item.LastWriteTime
            foreach ($o in $node.Owners) {
                $o.SizeBytes += $len
                $o.Files++
                if ($lw -gt $o.Newest) { $o.Newest = $lw }
            }
        }

        if ($scanned % 5000 -eq 0) {
            Write-Progress -Activity 'Measuring folder sizes' `
                -Status ("{0:N0} items scanned, {1:N2} GB so far" -f $scanned, ($rootStat.SizeBytes / 1GB)) `
                -CurrentOperation (Remove-LongPathPrefix $node.Dir)
        }
    }
}
Write-Progress -Activity 'Measuring folder sizes' -Completed
$stopwatch.Stop()

# --- Build rows -----------------------------------------------------------

$total = [math]::Max($rootStat.SizeBytes, 1)
$now   = Get-Date

function ConvertTo-Row($s) {
    [pscustomobject]@{
        Folder         = $s.Folder
        Level          = $s.Level
        SizeGB         = [math]::Round($s.SizeBytes / 1GB, 2)
        SizeMB         = [math]::Round($s.SizeBytes / 1MB, 1)
        Files          = $s.Files
        Subfolders     = $s.Subfolders
        PercentOfTotal = [math]::Round(($s.SizeBytes / $total) * 100, 1)
        NewestFile     = if ($s.Newest -gt [datetime]::MinValue) { $s.Newest } else { $null }
        DaysSinceChange = if ($s.Newest -gt [datetime]::MinValue) { [int][math]::Floor(($now - $s.Newest).TotalDays) } else { $null }
    }
}

$rows = @($stats | Sort-Object -Property SizeBytes -Descending)
if ($Top -gt 0) { $rows = @($rows | Select-Object -First $Top) }
$rows = @(ConvertTo-Row $rootStat) + @($rows | ForEach-Object { ConvertTo-Row $_ })

$summary = New-SummaryRows ([ordered]@{
    'Folder scanned'              = $root
    'Scan started'                = $startTime.ToString('yyyy-MM-dd HH:mm:ss')
    'Duration (seconds)'          = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
    'Levels reported'             = $Depth
    'Total size (GB)'             = [math]::Round($rootStat.SizeBytes / 1GB, 2)
    'Total files'                 = $rootStat.Files
    'Total folders'               = $rootStat.Subfolders
    'Folders that failed to open' = $errors.Count
    'Note'                        = 'Row 1 of Folder Sizes is the scanned folder itself. Each size includes all subfolders.'
    'Scanned by'                  = "$env:USERDOMAIN\$env:USERNAME"
})

$saved = Save-Workbook -FilePath $reportPath -Sheets @(
    @{ Name = 'Folder Sizes'; Rows = $rows
       Headers = @('Folder', 'Level', 'SizeGB', 'SizeMB', 'Files', 'Subfolders', 'PercentOfTotal', 'NewestFile', 'DaysSinceChange')
       Widths  = @(90, 8, 10, 12, 12, 12, 15, 18, 16) }
    @{ Name = 'Access Errors'; Rows = $errors
       Headers = @('FullPath', 'ErrorType', 'Error')
       Widths  = @(90, 30, 80) }
    @{ Name = 'Summary'; Rows = $summary; Headers = @('Item', 'Value'); Widths = @(30, 90) }
)
if (-not $saved) { exit 1 }

# --- Console summary ------------------------------------------------------

Write-Host ("Scanned {0:N0} items in {1:N1} seconds. Total size {2:N2} GB." -f $scanned, $stopwatch.Elapsed.TotalSeconds, ($rootStat.SizeBytes / 1GB))
Write-Host ""
Write-Host "Largest folders:"
$rows | Select-Object -Skip 1 -First 10 | ForEach-Object {
    $size = if ($_.SizeGB -ge 1) { '{0:N2} GB' -f $_.SizeGB } else { '{0:N1} MB' -f $_.SizeMB }
    Write-Host ("  {0,12}  {1}" -f $size, $_.Folder)
}
if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host ("Could not open {0:N0} folder(s); their contents are not counted. See the 'Access Errors' sheet." -f $errors.Count) -ForegroundColor Red
}
Write-Host ""
Write-Host "Report  : $reportPath"
