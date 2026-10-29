<#
.SYNOPSIS
    Finds file and folder paths longer than the Windows 260-character limit.
.DESCRIPTION
    Saves an Excel workbook with the long paths, any folders that couldn't be read, and a summary.
.EXAMPLE
    .\Find-LongPaths.ps1 -Path "D:\Shares\Finance"
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
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [Alias('OutputCsv')]
    [string]$OutputFile = (Join-Path -Path (Get-Location).Path -ChildPath ("LongPaths_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [int]$MaxLength = 259
)

# --- Path helpers ---------------------------------------------------------

function Add-LongPathPrefix([string]$p) {
    if ($p.StartsWith('\\?\')) { return $p }
    if ($p.StartsWith('\\'))   { return '\\?\UNC\' + $p.Substring(2) }   # UNC path
    return '\\?\' + $p
}

function Remove-LongPathPrefix([string]$p) {
    if ($p.StartsWith('\\?\UNC\')) { return '\\' + $p.Substring(8) }
    if ($p.StartsWith('\\?\'))     { return $p.Substring(4) }
    return $p
}

# --- Excel (.xlsx) writer helpers -----------------------------------------
# An .xlsx file is a zip of XML files; these functions build one directly.

$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function ConvertTo-XmlText([string]$s) {
    if ($null -eq $s) { return '' }
    # Drop characters that are not allowed in XML, then escape & < > " '
    $s = [regex]::Replace($s, '[^\u0009\u000A\u000D\u0020-\uFFFD]', '')
    return [System.Security.SecurityElement]::Escape($s)
}

function Add-ZipText($Zip, [string]$Name, [string]$Text) {
    $entry  = $Zip.CreateEntry($Name)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Add-XlsxSheet {
    param(
        $Zip,
        [int]$Index,
        [string[]]$Headers,
        [object[]]$Rows,
        [int[]]$Widths
    )
    $letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $lastCol = $letters[$Headers.Count - 1]

    $entry = $Zip.CreateEntry("xl/worksheets/sheet$Index.xml")
    $w = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8NoBom)
    try {
        $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        # Freeze the header row
        $w.Write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')

        $w.Write('<cols>')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $w.Write(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $Widths[$i]))
        }
        $w.Write('</cols><sheetData>')

        # Header row (style 1 = bold white on blue)
        $w.Write('<row r="1">')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $w.Write(('<c r="{0}1" s="1" t="inlineStr"><is><t>{1}</t></is></c>' -f $letters[$i], (ConvertTo-XmlText $Headers[$i])))
        }
        $w.Write('</row>')

        # Data rows
        $r = 1
        foreach ($row in $Rows) {
            $r++
            $w.Write("<row r=`"$r`">")
            for ($i = 0; $i -lt $Headers.Count; $i++) {
                $ref = "$($letters[$i])$r"
                $val = $row.($Headers[$i])
                if ($val -is [int] -or $val -is [long]) {
                    $w.Write("<c r=`"$ref`"><v>$($val.ToString([System.Globalization.CultureInfo]::InvariantCulture))</v></c>")
                } else {
                    $w.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText "$val")</t></is></c>")
                }
            }
            $w.Write('</row>')
        }
        $w.Write('</sheetData>')
        # Filter buttons on the header row
        $w.Write("<autoFilter ref=`"A1:$lastCol$r`"/>")
        $w.Write('</worksheet>')
    }
    finally { $w.Dispose() }
}

function Export-Xlsx {
    param(
        [string]$FilePath,
        [object[]]$Sheets   # each: @{ Name; Headers; Rows; Widths }
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -LiteralPath $FilePath) { Remove-Item -LiteralPath $FilePath -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($FilePath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $sheetOverrides = ''
        $sheetEntries   = ''
        $sheetRels      = ''
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $n = $i + 1
            $sheetOverrides += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
            $sheetEntries   += "<sheet name=`"$(ConvertTo-XmlText $Sheets[$i].Name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
            $sheetRels      += "<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
        }
        $stylesId = $Sheets.Count + 1

        Add-ZipText $zip '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
            $sheetOverrides + '</Types>')

        Add-ZipText $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>')

        Add-ZipText $zip 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheets>' + $sheetEntries + '</sheets></workbook>')

        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $sheetRels +
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

# --- Validate input -------------------------------------------------------

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Error "Folder not found or not accessible: $Path"
    exit 1
}

if ([System.IO.Path]::GetExtension($OutputFile) -ne '.xlsx') {
    $OutputFile = [System.IO.Path]::ChangeExtension($OutputFile, '.xlsx')
    Write-Host "Note     : output will be saved as $OutputFile" -ForegroundColor Cyan
}

$root = (Get-Item -LiteralPath $Path).FullName.TrimEnd('\')
if ($root -match '^[A-Za-z]:$') { $root += '\' }   # keep "C:\" valid

Write-Host "Scanning : $root"
Write-Host "Limit    : paths longer than $MaxLength characters"
Write-Host ""

# --- Scan (iterative, so one unreadable folder doesn't stop the run) ------

$results = [System.Collections.Generic.List[object]]::new()
$errors  = [System.Collections.Generic.List[object]]::new()
$stack   = [System.Collections.Generic.Stack[string]]::new()
$stack.Push((Add-LongPathPrefix $root))

$scanned   = 0
$startTime = Get-Date
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

while ($stack.Count -gt 0) {
    $currentDir = $stack.Pop()

    try {
        $items = ([System.IO.DirectoryInfo]::new($currentDir)).GetFileSystemInfos()
    }
    catch {
        $ex = $_.Exception
        if ($ex.InnerException) { $ex = $ex.InnerException }
        $failedPath = Remove-LongPathPrefix $currentDir
        $errors.Add([pscustomobject]@{
            FullPath   = $failedPath
            PathLength = $failedPath.Length
            ErrorType  = $ex.GetType().Name
            Error      = $ex.Message
        })
        continue
    }

    foreach ($item in $items) {
        $scanned++
        $fullPath  = Remove-LongPathPrefix $item.FullName
        $isFolder  = ($item.Attributes -band [System.IO.FileAttributes]::Directory) -ne 0
        $isReparse = ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0

        if ($fullPath.Length -gt $MaxLength) {
            $results.Add([pscustomobject]@{
                FullPath     = $fullPath
                PathLength   = $fullPath.Length
                OverLimitBy  = $fullPath.Length - $MaxLength
                Type         = if ($isFolder) { 'Folder' } else { 'File' }
                Name         = $item.Name
                ParentFolder = $fullPath.Substring(0, $fullPath.Length - $item.Name.Length).TrimEnd('\')
            })
        }

        # Recurse into real folders only (skip junctions/symlinks)
        if ($isFolder -and -not $isReparse) {
            $stack.Push($item.FullName)
        }

        if ($scanned % 2000 -eq 0) {
            Write-Progress -Activity "Scanning for long paths" `
                -Status ("{0:N0} items scanned, {1:N0} long paths found" -f $scanned, $results.Count) `
                -CurrentOperation (Remove-LongPathPrefix $currentDir)
        }
    }
}

Write-Progress -Activity "Scanning for long paths" -Completed
$stopwatch.Stop()

# --- Export ---------------------------------------------------------------

$outDir = Split-Path -Path $OutputFile -Parent
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$summary = @(
    [pscustomobject]@{ Item = 'Folder scanned';              Value = $root }
    [pscustomobject]@{ Item = 'Scan started';                Value = $startTime.ToString('yyyy-MM-dd HH:mm:ss') }
    [pscustomobject]@{ Item = 'Duration (seconds)';          Value = [string][math]::Round($stopwatch.Elapsed.TotalSeconds, 1) }
    [pscustomobject]@{ Item = 'Path length limit';           Value = $MaxLength }
    [pscustomobject]@{ Item = 'Items scanned';               Value = $scanned }
    [pscustomobject]@{ Item = 'Paths over the limit';        Value = $results.Count }
    [pscustomobject]@{ Item = 'Folders that failed to open'; Value = $errors.Count }
    [pscustomobject]@{ Item = 'Scanned by';                  Value = "$env:USERDOMAIN\$env:USERNAME" }
)

$sheets = @(
    @{ Name    = 'Long Paths'
       Headers = @('FullPath', 'PathLength', 'OverLimitBy', 'Type', 'Name', 'ParentFolder')
       Rows    = @($results | Sort-Object -Property PathLength -Descending)
       Widths  = @(100, 12, 13, 9, 40, 80) }
    @{ Name    = 'Access Errors'
       Headers = @('FullPath', 'PathLength', 'ErrorType', 'Error')
       Rows    = @($errors)
       Widths  = @(100, 12, 30, 80) }
    @{ Name    = 'Summary'
       Headers = @('Item', 'Value')
       Rows    = $summary
       Widths  = @(30, 80) }
)

try {
    Export-Xlsx -FilePath $OutputFile -Sheets $sheets
}
catch {
    Write-Error ("Could not save {0}. If it is open in Excel, close it and run again. Details: {1}" -f $OutputFile, $_.Exception.Message)
    exit 1
}

# --- Console summary ------------------------------------------------------

Write-Host ("Scanned {0:N0} items in {1:N1} seconds." -f $scanned, $stopwatch.Elapsed.TotalSeconds)
Write-Host ("Found   {0:N0} paths longer than {1} characters." -f $results.Count, $MaxLength) -ForegroundColor $(if ($results.Count) { 'Yellow' } else { 'Green' })
Write-Host ""

if ($errors.Count -gt 0) {
    Write-Host ("Could not access {0:N0} folder(s). Their contents were NOT scanned:" -f $errors.Count) -ForegroundColor Red
    $errors | Select-Object -First 20 | ForEach-Object {
        Write-Host ("  {0}  [{1}]" -f $_.FullPath, $_.ErrorType) -ForegroundColor Red
    }
    if ($errors.Count -gt 20) { Write-Host ("  ...and {0:N0} more - see the 'Access Errors' sheet" -f ($errors.Count - 20)) -ForegroundColor Red }
} else {
    Write-Host "No access errors - every folder was scanned." -ForegroundColor Green
}

Write-Host ""
Write-Host "Results : $OutputFile  (sheets: Long Paths, Access Errors, Summary)"
