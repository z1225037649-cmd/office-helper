# Standalone Windows PowerShell 5.1 tests. No certificate store writes.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\..\portable\certificates.ps1')

function Assert-True { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw ('ASSERT: ' + $Message) } }
function Assert-Throws { param([scriptblock]$Action,[string]$Message) $thrown=$false; try { & $Action | Out-Null } catch { $thrown=$true }; Assert-True $thrown $Message }

# Replace only the OS-store boundary and issuer; PEM and key validation remain real.
$script:trust = @{}
$script:events = New-Object 'Collections.Generic.List[string]'
$script:clock = [DateTime]::UtcNow
$script:failTrust=$false
$script:failManifest=$false
$script:failRemove=$false
function Get-HelperCertificateClock { return $script:clock }
function Get-HelperTrustedCertificate {
    param([string]$Thumbprint)
    if ($script:trust.ContainsKey($Thumbprint)) { return New-Object Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,$script:trust[$Thumbprint]) }
    return $null
}
function Add-HelperCertificateTrust {
    param([Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)
    $script:events.Add('add:'+$Certificate.Thumbprint)
    $script:trust[$Certificate.Thumbprint]=$Certificate.RawData
    if ($script:failTrust) { throw 'Injected trust failure after store write.' }
}
function Remove-HelperCertificateTrust {
    param([string]$Thumbprint)
    if ($script:failRemove) { throw 'Injected rollback failure.' }
    $script:events.Add('remove:'+$Thumbprint)
    $script:trust.Remove($Thumbprint)
}
function New-HelperCertificateMaterial {
    param([string]$VersionDirectory)
    $rsa=[Security.Cryptography.RSA]::Create(2048)
    $certificate=$null
    try {
        $request=[Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=Office Helper Portable localhost',$rsa,[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new($false,$false,0,$true))
        $oids=New-Object Security.Cryptography.OidCollection
        [void]$oids.Add((New-Object Security.Cryptography.Oid('1.3.6.1.5.5.7.3.1')))
        $request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($oids,$false))
        $san=[Security.Cryptography.X509Certificates.SubjectAlternativeNameBuilder]::new()
        $san.AddDnsName('localhost'); $san.AddIpAddress([Net.IPAddress]::Parse('127.0.0.1'))
        $request.CertificateExtensions.Add($san.Build())
        $certificate=$request.CreateSelfSigned([DateTimeOffset]$script:clock.AddMinutes(-5),[DateTimeOffset]$script:clock.AddDays(30))
        [OfficeHelperPortable.CertificateCrypto]::WriteMaterial($certificate,(Join-Path $VersionDirectory 'localhost.crt'),(Join-Path $VersionDirectory 'localhost.key'))
        return $certificate.Thumbprint
    } finally { if($certificate){$certificate.Dispose()}; $rsa.Dispose() }
}
$script:realWriteManifest = ${function:Write-HelperCertificateManifest}
function Write-HelperCertificateManifest {
    param([string]$StateDirectory,$Manifest)
    if ($script:failManifest) { throw 'Injected manifest write failure.' }
    & $script:realWriteManifest $StateDirectory $Manifest
}

$testRoot=Join-Path $env:TEMP ('OfficeHelperCertificateTests-'+[Guid]::NewGuid().ToString('N'))
$state=Join-Path $testRoot 'certificates'
[void][IO.Directory]::CreateDirectory($testRoot)
$script:passed=0
try {
    $missing=Get-HelperCertificateStatus $state
    Assert-True (-not $missing.valid -and -not $missing.trusted -and $missing.needsRenewal) 'Missing state must require setup.'
    Assert-Throws { Ensure-HelperCertificate $state } 'Missing consent must stop setup.'
    Assert-True (-not [IO.Directory]::Exists($state) -and $script:events.Count -eq 0) 'No-consent check cannot create files or trust.'
    $script:passed++

    $first=Ensure-HelperCertificate $state -AllowTrust
    Assert-True ($first.valid -and $first.trusted -and -not $first.needsRenewal) 'Initial material must be valid and trusted.'
    Assert-True ([IO.File]::Exists($first.certificatePath) -and [IO.File]::Exists($first.privateKeyPath)) 'HTTPS PEM paths must exist.'
    $manifestPath=Join-Path $state 'manifest.json'
    $firstManifest=[IO.File]::ReadAllText($manifestPath)
    $eventCount=$script:events.Count
    $again=Ensure-HelperCertificate $state
    Assert-True ($again.thumbprint -eq $first.thumbprint -and $script:events.Count -eq $eventCount) 'Healthy certificate must be reused without consent or trust writes.'
    $script:passed++

    $unrelated='AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    $script:trust[$unrelated]=[byte[]]@(1,2,3)
    $script:trust.Remove($first.thumbprint)
    $untrusted=Get-HelperCertificateStatus $state
    Assert-True ($untrusted.valid -and -not $untrusted.trusted) 'Unrelated Root certificate is not sufficient trust.'
    Assert-Throws { Ensure-HelperCertificate $state } 'Trust repair requires consent.'
    $repaired=Ensure-HelperCertificate $state -AllowTrust
    Assert-True ($repaired.thumbprint -eq $first.thumbprint -and $repaired.trusted) 'Trust repair should preserve healthy key material.'
    $script:passed++

    $expiry=[DateTime]::Parse($first.expiresAt).ToUniversalTime()
    $script:clock=$expiry.AddDays(-8)
    Assert-True (-not (Get-HelperCertificateStatus $state).needsRenewal) 'Eight days left does not renew.'
    $script:clock=$expiry.AddDays(-7)
    Assert-True ((Get-HelperCertificateStatus $state).needsRenewal) 'Seven days left requires renewal.'
    Assert-Throws { Ensure-HelperCertificate $state } 'Renewal requires consent.'
    Assert-True ([IO.File]::ReadAllText($manifestPath) -eq $firstManifest) 'Denied renewal must preserve current manifest.'
    $script:passed++

    $beforeRenewalEvents=$script:events.Count
    $renewed=Ensure-HelperCertificate $state -AllowTrust
    Assert-True ($renewed.valid -and $renewed.trusted -and $renewed.thumbprint -ne $first.thumbprint) 'Due certificate must be replaced.'
    Assert-True (-not [IO.File]::Exists($first.privateKeyPath) -and -not $script:trust.ContainsKey($first.thumbprint)) 'Recorded old version must be cleaned.'
    Assert-True ($script:events[$beforeRenewalEvents] -eq ('add:'+$renewed.thumbprint) -and $script:events[$beforeRenewalEvents+1] -eq ('remove:'+$first.thumbprint)) 'New trust must precede old cleanup.'
    Assert-True $script:trust.ContainsKey($unrelated) 'Unrelated trust must remain.'
    $script:passed++

    $stableManifest=[IO.File]::ReadAllText($manifestPath)
    $script:failTrust=$true
    Assert-Throws { Ensure-HelperCertificate $state -AllowTrust -ForceRenew } 'Injected trust failure must fail transaction.'
    $script:failTrust=$false
    Assert-True ([IO.File]::ReadAllText($manifestPath) -eq $stableManifest) 'Trust failure must preserve old manifest.'
    Assert-True ((Get-HelperCertificateStatus $state).thumbprint -eq $renewed.thumbprint -and $script:trust.Count -eq 2) 'Trust failure must roll back only new trust.'
    Assert-True (@(Get-ChildItem -LiteralPath $state -Directory).Count -eq 1) 'Trust failure must remove staged material.'
    $script:passed++

    $script:failManifest=$true
    Assert-Throws { Ensure-HelperCertificate $state -AllowTrust -ForceRenew } 'Injected manifest failure must fail transaction.'
    $script:failManifest=$false
    Assert-True ([IO.File]::ReadAllText($manifestPath) -eq $stableManifest -and $script:trust.Count -eq 2) 'Manifest failure must preserve old files and trust.'
    Assert-True (@(Get-ChildItem -LiteralPath $state -Directory).Count -eq 1) 'Manifest failure must remove staged material.'
    $script:passed++

    $script:failTrust=$true
    $script:failRemove=$true
    Assert-Throws { Ensure-HelperCertificate $state -AllowTrust -ForceRenew } 'Failed rollback must report retry requirement.'
    $script:failTrust=$false
    $script:failRemove=$false
    Assert-True ([IO.File]::Exists((Join-Path $state 'pending.json')) -and $script:trust.Count -eq 3) 'Failed rollback must retain exact ownership journal.'
    Assert-True ([IO.File]::ReadAllText($manifestPath) -eq $stableManifest) 'Failed rollback must not replace old active certificate.'
    $recovered=Ensure-HelperCertificate $state -AllowTrust
    Assert-True ($recovered.thumbprint -eq $renewed.thumbprint -and $script:trust.Count -eq 2 -and -not [IO.File]::Exists((Join-Path $state 'pending.json'))) 'Retry must clean pending trust and retain active certificate.'
    $script:passed++

    $otherState=Join-Path $testRoot 'other'
    $other=Ensure-HelperCertificate $otherState -AllowTrust
    $originalKey=[IO.File]::ReadAllText($renewed.privateKeyPath)
    [IO.File]::WriteAllText($renewed.privateKeyPath,[IO.File]::ReadAllText($other.privateKeyPath))
    $mismatch=Get-HelperCertificateStatus $state
    Assert-True (-not $mismatch.valid -and $mismatch.needsRenewal -and $mismatch.reason -eq 'key-mismatch') 'Wrong private key must invalidate status.'
    [IO.File]::WriteAllText($renewed.privateKeyPath,$originalKey)
    $script:passed++

    $script:clock=[DateTime]::Parse($renewed.expiresAt).ToUniversalTime().AddSeconds(1)
    $expired=Get-HelperCertificateStatus $state
    Assert-True (-not $expired.valid -and $expired.trusted -and $expired.needsRenewal) 'Trusted expired certificate is unusable.'
    $script:passed++

    $badManifest=($stableManifest | ConvertFrom-Json)
    $badManifest.activeId='..\outside'; $badManifest.certificates[0].id='..\outside'
    [IO.File]::WriteAllText($manifestPath,($badManifest | ConvertTo-Json -Depth 5))
    Assert-True (-not (Get-HelperCertificateStatus $state).valid) 'Traversal in manifest must be rejected.'
    Assert-Throws { Remove-HelperCertificate $state } 'Unsafe manifest cannot drive deletion.'
    [IO.File]::WriteAllText($manifestPath,$stableManifest)
    $script:passed++

    Remove-HelperCertificate $state | Out-Null
    Assert-True (-not [IO.File]::Exists($manifestPath) -and -not [IO.File]::Exists($renewed.privateKeyPath)) 'Removal cleans recorded files.'
    Assert-True (-not $script:trust.ContainsKey($renewed.thumbprint) -and $script:trust.ContainsKey($other.thumbprint) -and $script:trust.ContainsKey($unrelated)) 'Removal is limited to this manifest.'
    Remove-HelperCertificate $state | Out-Null
    Remove-HelperCertificate $otherState | Out-Null
    Assert-True ($script:trust.Count -eq 1 -and $script:trust.ContainsKey($unrelated)) 'Repeated removal is harmless.'
    $script:passed++
    Write-Output ('PASS: '+$script:passed+' certificate lifecycle tests (in-memory trust; no OS trust changes).')
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $tempPrefix=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\OfficeHelperCertificateTests-'
    if ($resolved.StartsWith($tempPrefix,[StringComparison]::OrdinalIgnoreCase) -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
