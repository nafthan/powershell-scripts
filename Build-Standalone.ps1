<#
.SYNOPSIS
    Builds single-file versions of the toolkit scripts into the Standalone folder.
.DESCRIPTION
    Each standalone script gets the shared helpers it needs copied in, reads no config file,
    and takes email settings as parameters. Run this after changing any toolkit script.
.EXAMPLE
    .\build\Build-Standalone.ps1
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
    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'Standalone' }
$utf8 = New-Object System.Text.UTF8Encoding($false)

$sourceFolders = 'FileServer', 'ActiveDirectory', 'ServerHealth', 'Network', 'Security', 'M365'
$dotSource     = ". ([System.IO.Path]::Combine(`$PSScriptRoot, '..', 'Shared', 'Common.ps1'))"
$notesAnchor   = '    Repo   : https://github.com/nafthan/powershell-scripts'
$notesLine     = '    Version: standalone - single file, no config needed. The full toolkit (shared config) is in the repo.'

# --- Per-script additions ---------------------------------------------------

$extras = @{
    'Get-OpenPortReport.ps1' = @{
        Params  = @('[string[]]$AllowedPort')
        Config  = 'if ($AllowedPort) { $cfg.PortCheck = @{ Allowed = @{ Default = @($AllowedPort | ForEach-Object { "$_" -split "," } | Where-Object { $_.Trim() } | ForEach-Object { [int]$_.Trim() }) } } }'
        Replace = [ordered]@{
            'Compares open ports with the allowed list in config.psd1 and flags risky services.' = 'Flags risky services (Telnet, FTP, VNC...) and any open port not listed in -AllowedPort.'
            '.\Get-OpenPortReport.ps1 -ComputerName SERVER01, SERVER02'                           = '.\Get-OpenPortReport.ps1 -ComputerName SERVER01, SERVER02 -AllowedPort 135, 445, 3389, 5985'
            "'PortCheck.Allowed in config.psd1'"                                                   = "('-AllowedPort ' + (`$AllowedPort -join ', '))"
            ' or set PortCheck.ComputerName / Servers.ComputerName'                                = ''
        }
    }
    'Get-PasswordExpiryReport.ps1' = @{
        Params  = @('[string]$HelpdeskContact')
        Config  = 'if ($HelpdeskContact) { $cfg.PasswordExpiry = @{ HelpdeskContact = $HelpdeskContact } }'
        Replace = [ordered]@{
            'Reminders are only sent on the days set in PasswordExpiry.NotifyOnDays.' = 'Reminders are only sent at 14, 7, 3, 2, 1 and 0 days left.'
        }
    }
    'Get-MfaStatusReport.ps1' = @{
        Replace = [ordered]@{ 'Set M365.TenantId in config.psd1 or use -TenantId.' = 'Use -TenantId.' }
    }
}

$emailParams = @(
    '# Email settings (for -SendEmail / -NotifyUsers)'
    '[string]$SmtpServer'
    '[int]$SmtpPort = 25'
    '[switch]$SmtpUseSsl'
    '[pscredential]$SmtpCredential'
    '[string]$MailFrom'
    '[string[]]$MailTo'
)

$configFunction = @'
function Get-ToolkitConfig {
    # Standalone: no config file. Settings come from parameters and built-in defaults.
    $cfg = @{ General = @{ OutputFolder = (Get-Location).Path } }
    if ($SmtpServer) {
        $cfg.Email = @{
            SmtpServer = $SmtpServer; Port = $SmtpPort; UseSsl = [bool]$SmtpUseSsl; From = $MailFrom
            To = @($MailTo | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
    }
    #EXTRA#
    return $cfg
}
'@

# --- Shared helpers ---------------------------------------------------------

$commonAst = [System.Management.Automation.Language.Parser]::ParseFile([System.IO.Path]::Combine($root, 'Shared', 'Common.ps1'), [ref]$null, [ref]$null)
$functions = [ordered]@{}
$preamble  = @()
foreach ($statement in $commonAst.EndBlock.Statements) {
    if ($statement -is [System.Management.Automation.Language.FunctionDefinitionAst]) { $functions[$statement.Name] = $statement.Extent.Text }
    else { $preamble += $statement.Extent.Text }
}
$preamble = $preamble -replace [regex]::Escape('Split-Path -Parent $PSScriptRoot'), '$PSScriptRoot'

# Simpler Get-Setting (no config file, so no placeholder checks)
$functions['Get-Setting'] = @'
function Get-Setting {
    param([hashtable]$Config, [string]$Section, [string]$Key, $Default = $null)
    $value = $null
    if ($Config -and $Config.ContainsKey($Section) -and $Config[$Section] -is [hashtable] -and $Config[$Section].ContainsKey($Key)) {
        $value = $Config[$Section][$Key]
    }
    if ($null -eq $value -or ($value -is [string] -and $value -eq '') -or ($value -is [array] -and $value.Count -eq 0)) { $value = $Default }
    return $value
}
'@
$preamble = @($preamble | Where-Object { $_ -notmatch 'Placeholder' })

$functions['Send-ToolkitMail'] = [regex]::Replace($functions['Send-ToolkitMail'], '(?s)<#.*?#>', '<# Sends an HTML email using -SmtpServer, -SmtpPort, -SmtpUseSsl, -SmtpCredential, -MailFrom and -MailTo. #>', 1)
$functions['Send-ToolkitMail'] = $functions['Send-ToolkitMail'].
    Replace("'Email not sent: set Email.SmtpServer, Email.From and Email.To in config.psd1.'", "'Email not sent: use -SmtpServer, -MailFrom and -MailTo.'").
    Replace('    if ($Attachments) { $params.Attachments = $Attachments }', "    if (`$SmtpCredential) { `$params.Credential = `$SmtpCredential }`n    if (`$Attachments) { `$params.Attachments = `$Attachments }")

function Get-CommandNames([string]$Text) {
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$null, [ref]$null)
    @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ } | Sort-Object -Unique)
}

function Get-NeededFunctions([string]$Text) {
    $needed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $queue  = [System.Collections.Generic.Queue[string]]::new()
    foreach ($n in (Get-CommandNames $Text)) { if ($functions.Contains($n)) { $queue.Enqueue($n) } }
    while ($queue.Count) {
        $n = $queue.Dequeue()
        if (-not $needed.Add($n)) { continue }
        foreach ($d in (Get-CommandNames $functions[$n])) { if ($functions.Contains($d) -and -not $needed.Contains($d)) { $queue.Enqueue($d) } }
    }
    # Keep the original order from Common.ps1
    @($functions.Keys | Where-Object { $needed.Contains($_) })
}

# --- Build ------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

foreach ($folder in $sourceFolders) {
    foreach ($file in (Get-ChildItem -Path (Join-Path $root $folder) -Filter *.ps1 -ErrorAction SilentlyContinue)) {
        $text  = [System.IO.File]::ReadAllText($file.FullName)
        $extra = $extras[$file.Name]

        if ($text.Contains($dotSource)) {
            # 1. Extra parameters, added after the last existing parameter
            $ast    = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$null)
            $add    = @()
            if ($text -match 'Send-ToolkitMail') { $add += $emailParams }
            if ($extra -and $extra.Params) { $add += $extra.Params }
            if ($add.Count) {
                $insertAt = $ast.ParamBlock.Parameters[-1].Extent.EndOffset
                $lines    = foreach ($a in $add) { if ($a.StartsWith('#')) { "`n`n    $a" } else { "`n    $a," } }
                $block    = (',' + ($lines -join '')).TrimEnd(',')
                $text     = $text.Insert($insertAt, $block)
            }

            # 2. Shared helpers in place of the dot-source line
            $cfgFn = $configFunction.Replace('#EXTRA#', $(if ($extra -and $extra.Config) { $extra.Config } else { '' }))
            $names = Get-NeededFunctions ($text + "`n" + $cfgFn) | Where-Object { $_ -ne 'Get-ToolkitConfig' }
            $helpers = @('#region Shared helpers (copied in by build/Build-Standalone.ps1 - edit Shared/Common.ps1 in the repo instead)')
            $helpers += $preamble
            $helpers += $cfgFn.Trim()
            foreach ($n in $names) { $helpers += $functions[$n] }
            $helpers += '#endregion'
            $text = $text.Replace($dotSource, ($helpers -join "`n`n"))
        }

        # 3. Text clean-up
        $text = [regex]::Replace($text, ' or set [\w\.]+ in config\.psd1', '')
        if ($extra -and $extra.Replace) {
            foreach ($k in $extra.Replace.Keys) {
                if (-not $text.Contains($k)) { Write-Warning "$($file.Name): text not found for replacement: $k" }
                $text = $text.Replace($k, $extra.Replace[$k])
            }
        }
        if ($text.Contains($notesAnchor)) { $text = $text.Replace($notesAnchor, "$notesAnchor`n$notesLine") }

        # 4. Check and save
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$errors)
        if ($errors) { throw "$($file.Name): generated script has syntax errors: $($errors[0].Message)" }

        $target = Join-Path $OutputFolder $file.Name
        [System.IO.File]::WriteAllText($target, $text, $utf8)
        Write-Host ("{0,-34} {1}" -f $file.Name, $(if ($text -match 'config\.psd1') { 'built (still mentions config.psd1 - check)' } else { 'built' }))
    }
}
