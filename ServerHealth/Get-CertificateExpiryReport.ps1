<#
.SYNOPSIS
    Finds expired and expiring certificates.
.DESCRIPTION
    Checks server certificate stores and live HTTPS/TLS endpoints.
.EXAMPLE
    .\Get-CertificateExpiryReport.ps1 -ComputerName WEB01 -Endpoint 'https://intranet.example.com'
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
    [string[]]$ComputerName,
    [string[]]$Endpoint,

    [ValidateRange(1, 3650)] [int]$WarningDays,
    [ValidateRange(0, 3650)] [int]$CriticalDays,

    [switch]$ExpiringOnly,
    [switch]$SkipStores,

    [switch]$SendEmail,
    [switch]$AlwaysEmail,

    [string]$OutputFile
)

. ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Shared', 'Common.ps1'))
$config = Get-ToolkitConfig

if (-not $PSBoundParameters.ContainsKey('ComputerName')) { $ComputerName = Get-Setting $config 'Servers' 'ComputerName' @([System.Environment]::MachineName) }
if (-not $PSBoundParameters.ContainsKey('Endpoint'))     { $Endpoint     = Get-Setting $config 'Certificates' 'Endpoints' @() }
if (-not $PSBoundParameters.ContainsKey('WarningDays'))  { $WarningDays  = Get-Setting $config 'Certificates' 'WarningDays' 30 }
if (-not $PSBoundParameters.ContainsKey('CriticalDays')) { $CriticalDays = Get-Setting $config 'Certificates' 'CriticalDays' 7 }
$ComputerName  = Split-ListValue $ComputerName
$Endpoint      = Split-ListValue $Endpoint
$stores        = @(Get-Setting $config 'Certificates' 'Stores' @('Cert:\LocalMachine\My', 'Cert:\LocalMachine\WebHosting'))
$ignoreExpired = [int](Get-Setting $config 'Certificates' 'IgnoreExpiredOlderThanDays' 30)

# --- Helpers --------------------------------------------------------------

function Get-CertStatus([int]$DaysLeft) {
    if ($DaysLeft -lt 0)              { return 'Expired' }
    if ($DaysLeft -le $CriticalDays)  { return 'Critical' }
    if ($DaysLeft -le $WarningDays)   { return 'Warning' }
    return 'OK'
}

function Get-CommonName([string]$DistinguishedName) {
    if ($DistinguishedName -match '(?:^|,\s*)CN=([^,]+)') { return $Matches[1] }
    return $DistinguishedName
}

function Get-SanText([System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert) {
    $ext = $Cert.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' }
    if (-not $ext) { return '' }
    return (($ext.Format($false) -split ',\s*') | ForEach-Object { $_ -replace '^(DNS Name|DNS)[=:]', '' }) -join ', '
}

function ConvertTo-EndpointTarget([string]$Value) {
    $value = $Value.Trim()
    if ($value -match '^[a-z]+://') {
        $uri = [uri]$value
        $port = if ($uri.IsDefaultPort -and $uri.Scheme -ne 'https') { 443 } else { $uri.Port }
        return @{ Host = $uri.Host; Port = $port }
    }
    if ($value -match '^(.+):(\d+)$') { return @{ Host = $Matches[1]; Port = [int]$Matches[2] } }
    return @{ Host = $value; Port = 443 }
}

function Get-EndpointCertificate([string]$TargetHost, [int]$Port, [int]$TimeoutMs = 8000) {
    $tcp = New-Object System.Net.Sockets.TcpClient
    try {
        $connect = $tcp.BeginConnect($TargetHost, $Port, $null, $null)
        if (-not $connect.AsyncWaitHandle.WaitOne($TimeoutMs)) { throw "Timed out connecting to ${TargetHost}:$Port" }
        $tcp.EndConnect($connect)

        $state    = @{ Errors = 'None' }
        $callback = [System.Net.Security.RemoteCertificateValidationCallback] {
            param($senderObj, $cert, $chain, $policyErrors)
            $state.Errors = "$policyErrors"
            return $true   # accept so the certificate can be inspected even if it is invalid
        }.GetNewClosure()

        $ssl = New-Object System.Net.Security.SslStream($tcp.GetStream(), $false, $callback)
        try {
            $ssl.ReadTimeout  = $TimeoutMs
            $ssl.WriteTimeout = $TimeoutMs
            $protocols = [System.Security.Authentication.SslProtocols]'Tls, Tls11, Tls12'
            if ([enum]::GetNames([System.Security.Authentication.SslProtocols]) -contains 'Tls13') {
                $protocols = $protocols -bor [System.Security.Authentication.SslProtocols]::Tls13
            }
            $ssl.AuthenticateAsClient($TargetHost, $null, $protocols, $false)
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
            return [pscustomobject]@{ Certificate = $cert; PolicyErrors = $state.Errors; Protocol = "$($ssl.SslProtocol)" }
        }
        finally { $ssl.Dispose() }
    }
    finally { $tcp.Dispose() }
}

# Runs locally or on each remote server
$storeScript = {
    param([string[]]$StorePaths)
    foreach ($storePath in $StorePaths) {
        if (-not (Test-Path -Path $storePath)) { continue }
        Get-ChildItem -Path $storePath | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509Certificate2] } | ForEach-Object {
            $san = ''
            $ext = $_.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' }
            if ($ext) { $san = (($ext.Format($false) -split ',\s*') | ForEach-Object { $_ -replace '^(DNS Name|DNS)[=:]', '' }) -join ', ' }
            [pscustomobject]@{
                Store         = $storePath
                Subject       = $_.Subject
                FriendlyName  = $_.FriendlyName
                SAN           = $san
                Issuer        = $_.Issuer
                NotBefore     = $_.NotBefore
                NotAfter      = $_.NotAfter
                Thumbprint    = $_.Thumbprint
                HasPrivateKey = $_.HasPrivateKey
                Computer      = [System.Environment]::MachineName
            }
        }
    }
}

# --- Main -----------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {
    $now        = Get-Date
    $reportPath = New-ReportPath -Config $config -Prefix 'Certificates' -OutputFile $OutputFile
    $rows       = [System.Collections.Generic.List[object]]::new()
    $errors     = [System.Collections.Generic.List[object]]::new()

    function Add-CertRow([string]$Source, [string]$Location, [string]$Store, $Subject, $FriendlyName, $San, $Issuer, [datetime]$NotAfter, $Thumbprint, $HasKey, [string]$Notes) {
        $daysLeft = [int][math]::Floor(($NotAfter - $now).TotalDays)
        if ($daysLeft -lt -$ignoreExpired) { return }
        $status = Get-CertStatus $daysLeft
        if ($ExpiringOnly -and $status -eq 'OK') { return }
        $rows.Add([pscustomobject]@{
            Status        = $status
            DaysLeft      = $daysLeft
            Expires       = $NotAfter
            Source        = $Source
            Location      = $Location
            Store         = $Store
            CommonName    = Get-CommonName $Subject
            FriendlyName  = $FriendlyName
            SAN           = $San
            Issuer        = Get-CommonName $Issuer
            Thumbprint    = $Thumbprint
            HasPrivateKey = $HasKey
            Notes         = $Notes
        })
    }

    # Certificate stores
    if (-not $SkipStores -and $ComputerName) {
        if (-not $script:IsWindowsHost) {
            Write-Warning 'Certificate stores can only be read from Windows - skipping store checks.'
        }
        else {
            $local  = @($ComputerName | Where-Object { $n = $_.Split('.')[0]; $n -eq 'localhost' -or $n -eq '.' -or $n -eq [System.Environment]::MachineName })
            $remote = @($ComputerName | Where-Object { $local -notcontains $_ })

            $storeResults = @()
            if ($local.Count) {
                Write-Host "Reading certificate stores on this computer..."
                $storeResults += & $storeScript $stores
            }
            if ($remote.Count) {
                Write-Host "Reading certificate stores on $($remote.Count) remote server(s)..."
                $icmErrors = $null
                $storeResults += Invoke-Command -ComputerName $remote -ScriptBlock $storeScript -ArgumentList (, $stores) `
                                 -ErrorAction SilentlyContinue -ErrorVariable icmErrors
                foreach ($e in $icmErrors) {
                    $target = if ($e.OriginInfo -and $e.OriginInfo.PSComputerName) { $e.OriginInfo.PSComputerName } else { "$($e.TargetObject)" }
                    $errors.Add([pscustomobject]@{ Source = 'Certificate store'; Target = $target; Error = $e.Exception.Message })
                }
            }
            foreach ($c in $storeResults) {
                $loc = if ($c.PSComputerName) { $c.PSComputerName } else { $c.Computer }
                Add-CertRow 'Certificate store' $loc ($c.Store -replace '^Cert:\\', '') $c.Subject $c.FriendlyName $c.SAN $c.Issuer $c.NotAfter $c.Thumbprint $c.HasPrivateKey ''
            }
        }
    }

    # Endpoints
    foreach ($ep in $Endpoint) {
        if (-not $ep) { continue }
        $label = $ep
        try {
            $t = ConvertTo-EndpointTarget $ep
            $label = "$($t.Host):$($t.Port)"
            Write-Host "Checking endpoint $label..."
            $r = Get-EndpointCertificate $t.Host $t.Port
            $notes = @("Protocol $($r.Protocol)")
            if ($r.PolicyErrors -ne 'None') { $notes += "Validation problem: $($r.PolicyErrors)" }
            $c = $r.Certificate
            Add-CertRow 'Endpoint' $label '' $c.Subject '' (Get-SanText $c) $c.Issuer $c.NotAfter $c.Thumbprint '' ($notes -join '; ')
        }
        catch {
            $err = Get-ErrorMessage $_
            $errors.Add([pscustomobject]@{ Source = 'Endpoint'; Target = $label; Error = $err.Message })
        }
    }

    # --- Export -----------------------------------------------------------

    $order  = @{ Expired = 0; Critical = 1; Warning = 2; OK = 3 }
    $sorted = @($rows | Sort-Object @{ Expression = { $order[$_.Status] } }, DaysLeft, Location)
    $counts = [ordered]@{}
    foreach ($s in 'Expired', 'Critical', 'Warning', 'OK') { $counts[$s] = @($sorted | Where-Object { $_.Status -eq $s }).Count }

    $summary = [ordered]@{
        'Run at'          = $now.ToString('yyyy-MM-dd HH:mm:ss')
        'Thresholds'      = "Warning <= $WarningDays days, Critical <= $CriticalDays days"
        'Servers checked' = if ($SkipStores) { '(stores skipped)' } else { ($ComputerName -join ', ') }
        'Stores'          = ($stores -join ', ')
        'Endpoints'       = ($Endpoint -join ', ')
        'Expired'         = $counts.Expired
        'Critical'        = $counts.Critical
        'Warning'         = $counts.Warning
        'OK'              = if ($ExpiringOnly) { 'Not listed (-ExpiringOnly)' } else { $counts.OK }
        'Errors'          = $errors.Count
        'Note'            = "Certificates that expired more than $ignoreExpired days ago are not listed."
        'Run by'          = "$env:USERDOMAIN\$env:USERNAME"
    }

    $headers = @('Status', 'DaysLeft', 'Expires', 'Source', 'Location', 'Store', 'CommonName', 'FriendlyName', 'SAN', 'Issuer', 'Thumbprint', 'HasPrivateKey', 'Notes')
    $saved = Save-Workbook -FilePath $reportPath -Sheets @(
        @{ Name = 'Certificates'; Rows = $sorted; Headers = $headers; Widths = @(10, 10, 18, 18, 30, 22, 40, 25, 60, 35, 44, 14, 50) }
        @{ Name = 'Errors'; Rows = $errors; Headers = @('Source', 'Target', 'Error'); Widths = @(18, 40, 100) }
        @{ Name = 'Summary'; Rows = (New-SummaryRows $summary); Headers = @('Item', 'Value'); Widths = @(18, 90) }
    )
    if (-not $saved) { exit 1 }

    Write-Host ""
    foreach ($s in 'Expired', 'Critical', 'Warning', 'OK') {
        $colour = switch ($s) { 'Expired' { 'Red' } 'Critical' { 'Red' } 'Warning' { 'Yellow' } default { 'Green' } }
        Write-Host ("{0,-9} {1:N0}" -f $s, $counts[$s]) -ForegroundColor $colour
    }
    if ($errors.Count) { Write-Host ("Errors    {0:N0} - see the 'Errors' sheet" -f $errors.Count) -ForegroundColor Red }
    Write-Host "Report  : $reportPath"

    # --- Email --------------------------------------------------------------

    $attention = @($sorted | Where-Object { $_.Status -ne 'OK' })
    if ($SendEmail) {
        if (-not $attention.Count -and -not $errors.Count -and -not $AlwaysEmail) {
            Write-Host 'Nothing expiring - no email sent (use -AlwaysEmail to send anyway).'
        }
        else {
            $sections = @(
                @{ Heading = "Needs attention ($($attention.Count))"
                   Html    = ConvertTo-HtmlTable -Rows $attention -Columns @('Status', 'DaysLeft', 'Expires', 'Location', 'CommonName', 'Issuer', 'Thumbprint') -StatusColumn 'Status' -EmptyText 'Nothing expiring within the warning window.' }
            )
            if ($errors.Count) {
                $sections += @{ Heading = "Checks that failed ($($errors.Count))"; Html = ConvertTo-HtmlTable -Rows $errors -Columns @('Source', 'Target', 'Error') }
            }
            $intro = @("$($counts.Expired) expired, $($counts.Critical) critical (<= $CriticalDays days), $($counts.Warning) warning (<= $WarningDays days).")
            $html  = New-HtmlReport -Title 'Certificate expiry report' -IntroLines $intro -Sections $sections
            $state = if ($counts.Expired -or $counts.Critical) { 'CRITICAL' } elseif ($counts.Warning) { 'Warning' } else { 'All OK' }
            $subject = "[$state] Certificates: $($counts.Expired) expired, $($counts.Critical) critical, $($counts.Warning) warning"
            if (Send-ToolkitMail -Config $config -Subject $subject -HtmlBody $html -Attachments $reportPath) { Write-Host 'Email sent.' }
        }
    }
}
