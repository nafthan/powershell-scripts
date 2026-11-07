<#
.SYNOPSIS
    Creates a report theme from a Word (.docx) or Excel (.xlsx) template, so reports match your organisation's style.
.DESCRIPTION
    Reads the fonts and colours a template actually uses (header bars, accent lines, label cells, text) and its
    Office theme, then writes a theme file for the HTML reports, emails and Excel workbooks. Nothing else is
    copied from the template. Check the result and adjust any colour by hand - see default.psd1 for each setting.
.EXAMPLE
    .\New-ReportTheme.ps1 -Template '.\Company Template.docx' -Organisation 'Contoso Ltd' -OutFile .\contoso.psd1
.EXAMPLE
    .\New-ReportTheme.ps1 -Template .\Template.xlsx -Classification Confidential
.NOTES
    Author : nafthan
    Website: https://nafthan.co.uk
    Repo   : https://github.com/nafthan/powershell-scripts
    Keep themes made from work templates out of public repos: name them *.local.psd1 (ignored by git) or store them elsewhere.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Template,
    [string]$Organisation = '',
    [string]$Classification = 'Internal',
    [string]$OutFile
)

$ErrorActionPreference = 'Stop'

# --- Colour helpers -------------------------------------------------------------

function Get-Rgb([string]$Hex) {
    $h = $Hex.TrimStart('#')
    if ($h.Length -eq 8) { $h = $h.Substring(2) }
    return @([Convert]::ToInt32($h.Substring(0, 2), 16), [Convert]::ToInt32($h.Substring(2, 2), 16), [Convert]::ToInt32($h.Substring(4, 2), 16))
}
function Get-Hsl([string]$Hex) {
    $r, $g, $b = (Get-Rgb $Hex) | ForEach-Object { $_ / 255 }
    $max = [math]::Max($r, [math]::Max($g, $b)); $min = [math]::Min($r, [math]::Min($g, $b)); $l = ($max + $min) / 2
    $s = 0; $hue = 0
    if ($max -ne $min) {
        $d = $max - $min
        $s = if ($l -gt 0.5) { $d / (2 - $max - $min) } else { $d / ($max + $min) }
        $hue = if ($max -eq $r) { (($g - $b) / $d) % 6 } elseif ($max -eq $g) { ($b - $r) / $d + 2 } else { ($r - $g) / $d + 4 }
        $hue = ($hue * 60 + 360) % 360
    }
    return @{ H = $hue; S = $s; L = $l }
}
function ConvertTo-HexColour([string]$Hex) { $h = $Hex.TrimStart('#'); if ($h.Length -eq 8) { $h = $h.Substring(2) }; "#$($h.ToLowerInvariant())" }
function Test-Hex([string]$Hex) { $Hex -match '^(?:[0-9A-Fa-f]{2})?[0-9A-Fa-f]{6}$' }

function Add-Count([hashtable]$Table, [string]$Key, [int]$By = 1) {
    if (-not $Key) { return }
    if ($Table.ContainsKey($Key)) { $Table[$Key] += $By } else { $Table[$Key] = $By }
}
function Get-Top([hashtable]$Table, [scriptblock]$Where = { $true }) {
    $Table.GetEnumerator() | Where-Object { & $Where $_.Key } | Sort-Object Value -Descending | Select-Object -First 1 -ExpandProperty Key
}

# --- Read the template ----------------------------------------------------------

$path = (Resolve-Path -LiteralPath $Template).Path
$ext  = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
if ($ext -notin '.docx', '.dotx', '.xlsx', '.xltx') { throw 'Template must be a .docx, .dotx, .xlsx or .xltx file.' }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($path)
function Read-Part([string]$Pattern) {
    $out = @{}
    foreach ($e in $zip.Entries | Where-Object { $_.FullName -match $Pattern }) {
        $sr = New-Object System.IO.StreamReader($e.Open()); try { $out[$e.FullName] = $sr.ReadToEnd() } finally { $sr.Dispose() }
    }
    return $out
}

try {
    $isWord = $ext -in '.docx', '.dotx'
    $prefix = if ($isWord) { 'word' } else { 'xl' }
    $themeXml = (Read-Part "^$prefix/theme/theme\d*\.xml$").Values | Select-Object -First 1

    # Office theme colours and fonts, used to resolve theme references and as fallbacks
    $office = @{}
    if ($themeXml) {
        foreach ($m in [regex]::Matches($themeXml, '<a:(dk1|lt1|dk2|lt2|accent[1-6])>.*?(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"', 'Singleline')) { $office[$m.Groups[1].Value] = ConvertTo-HexColour $m.Groups[2].Value }
        if ($themeXml -match '<a:minorFont>\s*<a:latin typeface="([^"]+)"') { $office.MinorFont = $Matches[1] }
        if ($themeXml -match '<a:majorFont>\s*<a:latin typeface="([^"]+)"') { $office.MajorFont = $Matches[1] }
    }

    $fills = @{}; $text = @{}; $lines = @{}; $fonts = @{}; $headerFills = @{}

    if ($isWord) {
        $xml = (Read-Part '^word/(document|header\d*|footer\d*|styles)\.xml$').Values -join "`n"
        foreach ($m in [regex]::Matches($xml, 'w:fill="([0-9A-Fa-f]{6})"')) { Add-Count $fills (ConvertTo-HexColour $m.Groups[1].Value) }
        foreach ($m in [regex]::Matches($xml, '<w:color w:val="([0-9A-Fa-f]{6})"')) { Add-Count $text (ConvertTo-HexColour $m.Groups[1].Value) }
        foreach ($m in [regex]::Matches($xml, '<w:(?:top|bottom|left|right)\b[^>]*w:color="([0-9A-Fa-f]{6})"')) { Add-Count $lines (ConvertTo-HexColour $m.Groups[1].Value) }
        foreach ($m in [regex]::Matches($xml, 'w:ascii="([^"]+)"')) { Add-Count $fonts $m.Groups[1].Value }
        if (-not $fonts.Count -and $xml -match 'w:asciiTheme="minor') { Add-Count $fonts $office.MinorFont }
        # Header cells: shaded table cells whose text is white
        foreach ($m in [regex]::Matches($xml, '<w:tc>.*?</w:tc>', 'Singleline')) {
            if ($m.Value -match 'w:fill="([0-9A-Fa-f]{6})"' -and $m.Value -match '<w:color w:val="(FFFFFF|ffffff)"') { Add-Count $headerFills (ConvertTo-HexColour $Matches[1]) }
        }
    }
    else {
        $styles = (Read-Part '^xl/styles\.xml$').Values | Select-Object -First 1
        $fontList = @([regex]::Matches($styles, '<font>.*?</font>|<font/>', 'Singleline') | ForEach-Object { $_.Value })
        $fillList = @([regex]::Matches($styles, '<fill>.*?</fill>', 'Singleline') | ForEach-Object { $_.Value })
        $cellXfs  = if ($styles -match '<cellXfs[^>]*>(.*?)</cellXfs>') { @([regex]::Matches($Matches[1], '<xf [^>]*>') | ForEach-Object { $_.Value }) } else { @() }
        $fontColour = { param($f) if ($f -match 'color rgb="([0-9A-Fa-f]{8})"') { ConvertTo-HexColour $Matches[1] } elseif ($f -match 'color theme="(\d)"') { @($office.lt1, $office.dk1, $office.lt2, $office.dk2)[[int]$Matches[1]] } }
        $fillColour = { param($f) if ($f -match 'solid' -and $f -match 'fgColor rgb="([0-9A-Fa-f]{8})"') { ConvertTo-HexColour $Matches[1] } }
        foreach ($xf in $cellXfs) {
            $fi = if ($xf -match 'fontId="(\d+)"') { [int]$Matches[1] } else { 0 }
            $fl = if ($xf -match 'fillId="(\d+)"') { [int]$Matches[1] } else { 0 }
            $fc = if ($fi -lt $fontList.Count) { & $fontColour $fontList[$fi] }
            $bg = if ($fl -lt $fillList.Count) { & $fillColour $fillList[$fl] }
            if ($bg) { Add-Count $fills $bg; if ($fc -and (Get-Hsl $fc).L -gt 0.9) { Add-Count $headerFills $bg } }
            if ($fc) { Add-Count $text $fc }
            if ($fi -lt $fontList.Count -and $fontList[$fi] -match '<name val="([^"]+)"') { Add-Count $fonts $Matches[1] }
        }
        foreach ($m in [regex]::Matches($styles, '<(?:top|bottom|left|right)[^>]*>\s*<color rgb="([0-9A-Fa-f]{8})"')) { Add-Count $lines (ConvertTo-HexColour $m.Groups[1].Value) }
    }
}
finally { $zip.Dispose() }

# --- Pick the theme colours ------------------------------------------------------

$all = @{}
foreach ($t in $fills, $text, $lines) { foreach ($k in $t.Keys) { Add-Count $all $k $t[$k] } }
$isDark      = { param($c) (Get-Hsl $c).L -lt 0.35 }
$isLightGrey = { param($c) $x = Get-Hsl $c; $x.L -gt 0.82 -and $x.L -lt 0.985 -and $x.S -lt 0.15 }
$isTemplateYellow = { param($c) $x = Get-Hsl $c; $x.H -ge 40 -and $x.H -le 70 -and $x.S -gt 0.5 -and $x.L -gt 0.8 }
$isBright    = { param($c) $x = Get-Hsl $c; $x.S -gt 0.45 -and $x.L -gt 0.3 -and $x.L -lt 0.75 }

$primary = Get-Top $headerFills $isDark
if (-not $primary) { $primary = Get-Top $fills $isDark }
if (-not $primary) { $primary = if ($office.dk2) { $office.dk2 } else { '#1f3a5f' } }
$pHue = (Get-Hsl $primary).H

$accent = Get-Top $all { param($c) (& $isBright $c) -and [math]::Abs((Get-Hsl $c).H - $pHue) -gt 40 }
$subtitle = Get-Top $text { param($c) $x = Get-Hsl $c; $x.S -gt 0.3 -and $x.L -lt 0.6 -and $c -ne $primary -and $c -ne $accent }
$bodyText = Get-Top $text { param($c) (Get-Hsl $c).L -lt 0.2 -and (Get-Hsl $c).S -lt 0.25 }
$muted    = Get-Top $text { param($c) $x = Get-Hsl $c; $x.L -ge 0.25 -and $x.L -le 0.6 -and $x.S -lt 0.15 }
$labelFill = Get-Top $fills $isLightGrey
$stripe   = Get-Top $fills { param($c) $x = Get-Hsl $c; $x.L -gt 0.9 -and $x.L -lt 0.995 -and $c -ne $labelFill -and -not (& $isTemplateYellow $c) }
$line     = Get-Top $lines { param($c) (Get-Hsl $c).L -gt 0.7 -and (Get-Hsl $c).S -lt 0.2 }

$font = Get-Top $fonts { param($n) $n -and $n -notmatch '^(Calibri|Cambria)$' }
if (-not $font) { $font = Get-Top $fonts }
if (-not $font) { $font = if ($office.MinorFont) { $office.MinorFont } else { 'Segoe UI' } }
$font = $font -replace "[`"';<>{}]", ''

$fallback = if ($font -eq 'Arial') { "'Arial', Helvetica, sans-serif" } else { "'$font', 'Segoe UI', Arial, sans-serif" }
$theme = [ordered]@{
    Organisation    = $Organisation
    Classification  = $Classification
    Font            = $fallback
    LabelFont       = $fallback
    UppercaseLabels = $false
    ExcelFont       = $font
    ExcelHeaderFont = $font
    Primary         = $primary
    OnPrimary       = $(if (& $isDark $primary) { '#ffffff' } else { '#1a1a1a' })
    Accent          = $(if ($accent) { $accent } else { '' })
    Heading         = $primary
    Subtitle        = $(if ($subtitle) { $subtitle } else { $primary })
    Text            = $(if ($bodyText) { $bodyText } else { '#1a1a1a' })
    Muted           = $(if ($muted) { $muted } else { '#5f6368' })
    Page            = '#f3f4f6'
    Card            = '#ffffff'
    Line            = $(if ($line) { $line } else { '#d9dde2' })
    Rule            = $(if ($line) { $line } else { '#c4cad1' })
    RowStripe       = $(if ($stripe) { $stripe } else { '' })
    LabelFill       = $(if ($labelFill) { $labelFill } else { '' })
}

# --- Write the theme file --------------------------------------------------------

if (-not $OutFile) { $OutFile = Join-Path (Get-Location).Path 'my-theme.local.psd1' }
$q = { param($v)
    if ($v -is [bool]) { if ($v) { '$true' } else { '$false' } }
    elseif ("$v" -match "'") { '"' + ("$v" -replace '[`"$]', '') + '"' }   # font lists contain quotes
    else { "'$v'" }
}
$notes = @{
    Primary = 'header bars and table headings'; Accent = 'stripe under the top bar and line under headings'; Subtitle = 'intro text under the title'
    Text = 'table text'; Muted = 'labels and footer'; Line = 'row lines'; RowStripe = 'every other row'; LabelFill = 'details block and Excel summary labels'
}
$body = foreach ($k in $theme.Keys) {
    $line = "    {0,-15} = {1}" -f $k, (& $q $theme[$k])
    if ($notes[$k]) { $line = "{0,-62} # {1}" -f $line, $notes[$k] }
    $line
}
$header = @(
    '# ============================================================================='
    "#  Report theme created by New-ReportTheme.ps1 on $(Get-Date -Format 'yyyy-MM-dd') from a $ext template."
    '#  Check the colours and adjust as needed - see default.psd1 for every setting,'
    '#  including the status colours, which are left at their defaults here.'
    '#  Made from a work template? Keep this file out of public repos.'
    '# ============================================================================='
)
Set-Content -LiteralPath $OutFile -Value ($header + '@{' + $body + '}') -Encoding UTF8

Write-Host "Theme written to $OutFile" -ForegroundColor Green
foreach ($k in 'Font', 'Primary', 'Accent', 'Subtitle', 'Text', 'Muted', 'LabelFill', 'RowStripe', 'Line') {
    Write-Host ("  {0,-10} {1}" -f $k, $(if ($theme[$k]) { $theme[$k] } else { '(none)' }))
}
Write-Host 'Use it with -ThemeFile on a standalone script, or Theme = @{ File = ''...'' } in config.psd1.'
