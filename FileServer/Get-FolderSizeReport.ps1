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

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('Path'))  { $Path  = Get-Setting $config 'FolderSize' 'Path' }
if (-not $PSBoundParameters.ContainsKey('Depth')) { $Depth = Get-Setting $config 'FolderSize' 'Depth' 2 }
if (-not $PSBoundParameters.ContainsKey('Top'))   { $Top   = Get-Setting $config 'FolderSize' 'Top' 0 }

if (-not $Path) { throw 'No folder given. Use -Path or set FolderSize.Path in config.psd1.' }
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
