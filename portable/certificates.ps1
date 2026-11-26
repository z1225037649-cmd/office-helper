# Windows PowerShell 5.1. Dot-source this file; it performs no work on import.
# AllowTrust is explicit user consent to change CurrentUser\Root only.
if (-not ('OfficeHelperPortable.CertificateCrypto' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
namespace OfficeHelperPortable {
    public static class CertificateCrypto {
        public const string Subject = "CN=Office Helper Portable localhost";
        static byte[] Join(params byte[][] parts) { var b = new List<byte>(); foreach(var p in parts) b.AddRange(p); return b.ToArray(); }
        static byte[] Der(byte tag, byte[] body) {
            var length = new List<byte>(); int n = body.Length;
            if(n < 128) length.Add((byte)n); else { var l = new List<byte>(); while(n > 0) { l.Insert(0, (byte)(n & 255)); n >>= 8; } length.Add((byte)(128 | l.Count)); length.AddRange(l); }
            return Join(new byte[]{tag}, length.ToArray(), body);
        }
        static byte[] Integer(byte[] b) { int i = 0; while(i < b.Length - 1 && b[i] == 0) i++; var v = new byte[b.Length-i]; Buffer.BlockCopy(b,i,v,0,v.Length); if((v[0]&128)!=0) v=Join(new byte[]{0},v); return Der(2,v); }
        static byte[] Read(byte[] input, ref int offset, byte tag) {
            if(offset >= input.Length || input[offset++] != tag || offset >= input.Length) throw new InvalidDataException("Invalid DER tag.");
            int size = input[offset++];
            if((size & 128)!=0) { int count=size & 127; if(count==0 || count>4 || offset+count>input.Length) throw new InvalidDataException("Invalid DER length."); size=0; for(int i=0;i<count;i++) size=checked((size<<8)|input[offset++]); }
            if(size<0 || size>input.Length-offset) throw new InvalidDataException("Truncated DER.");
            var result=new byte[size]; Buffer.BlockCopy(input,offset,result,0,size); offset+=size; return result;
        }
        static byte[] ReadInteger(byte[] bytes, ref int offset) { var v=Read(bytes,ref offset,2); if(v.Length==0) throw new InvalidDataException("Empty integer."); if(v.Length>1 && v[0]==0) { var r=new byte[v.Length-1]; Buffer.BlockCopy(v,1,r,0,r.Length); return r; } return v; }
        static byte[] Pad(byte[] bytes,int size) { if(bytes.Length>size) throw new InvalidDataException("Oversized RSA integer."); var result=new byte[size]; Buffer.BlockCopy(bytes,0,result,size-bytes.Length,bytes.Length); return result; }
        static byte[] Pem(string path,string label) {
            var text=File.ReadAllText(path,Encoding.ASCII).Trim(); string begin="-----BEGIN "+label+"-----", end="-----END "+label+"-----";
            if(!text.StartsWith(begin,StringComparison.Ordinal) || !text.EndsWith(end,StringComparison.Ordinal)) throw new InvalidDataException("Unexpected PEM type.");
            return Convert.FromBase64String(text.Substring(begin.Length,text.Length-begin.Length-end.Length));
        }
        static string ToPem(string label, byte[] data) { return "-----BEGIN "+label+"-----\r\n"+Convert.ToBase64String(data,Base64FormattingOptions.InsertLineBreaks)+"\r\n-----END "+label+"-----\r\n"; }
        public static void WriteMaterial(X509Certificate2 certificate,string certPath,string keyPath) {
            using(var rsa=RSACertificateExtensions.GetRSAPrivateKey(certificate)) {
                if(rsa==null) throw new InvalidDataException("RSA private key unavailable.");
                var p=rsa.ExportParameters(true);
                var key=Der(48,Join(Integer(new byte[]{0}),Integer(p.Modulus),Integer(p.Exponent),Integer(p.D),Integer(p.P),Integer(p.Q),Integer(p.DP),Integer(p.DQ),Integer(p.InverseQ)));
                File.WriteAllText(keyPath,ToPem("RSA PRIVATE KEY",key),Encoding.ASCII);
                File.WriteAllText(certPath,ToPem("CERTIFICATE",certificate.RawData),Encoding.ASCII);
                Array.Clear(key,0,key.Length);
            }
        }
        public static X509Certificate2 ReadCertificate(string path) { return new X509Certificate2(Pem(path,"CERTIFICATE")); }
        public static bool IsProductCertificate(X509Certificate2 c) {
            if(c.Subject!=Subject || c.Issuer!=Subject) return false;
            bool constraints=false,usage=false,dns=false,ip=false;
            foreach(X509Extension e in c.Extensions) {
                if(e.Oid.Value=="2.5.29.19") { var bc=new X509BasicConstraintsExtension(e,e.Critical); constraints=!bc.CertificateAuthority; }
                if(e.Oid.Value=="2.5.29.37") { var eku=new X509EnhancedKeyUsageExtension(e,e.Critical); foreach(Oid oid in eku.EnhancedKeyUsages) if(oid.Value=="1.3.6.1.5.5.7.3.1") usage=true; }
                if(e.Oid.Value=="2.5.29.17") { int o=0; var names=Read(e.RawData,ref o,48); if(o!=e.RawData.Length) return false; o=0; while(o<names.Length) { byte tag=names[o]; var name=Read(names,ref o,tag); if(tag==130 && Encoding.ASCII.GetString(name)=="localhost") dns=true; if(tag==135 && name.Length==4 && name[0]==127 && name[1]==0 && name[2]==0 && name[3]==1) ip=true; } }
            }
            using(var rsa=RSACertificateExtensions.GetRSAPublicKey(c)) { return constraints && usage && dns && ip && rsa!=null && rsa.KeySize>=2048; }
        }
        public static bool MatchesPrivateKey(X509Certificate2 c,string keyPath) {
            var bytes=Pem(keyPath,"RSA PRIVATE KEY"); int o=0; var sequence=Read(bytes,ref o,48); if(o!=bytes.Length) return false; o=0;
            var version=ReadInteger(sequence,ref o); if(version.Length!=1 || version[0]!=0) return false;
            var p=new RSAParameters(); p.Modulus=ReadInteger(sequence,ref o); p.Exponent=ReadInteger(sequence,ref o);
            p.D=Pad(ReadInteger(sequence,ref o),p.Modulus.Length); int half=(p.Modulus.Length+1)/2;
            p.P=Pad(ReadInteger(sequence,ref o),half); p.Q=Pad(ReadInteger(sequence,ref o),half); p.DP=Pad(ReadInteger(sequence,ref o),half); p.DQ=Pad(ReadInteger(sequence,ref o),half); p.InverseQ=Pad(ReadInteger(sequence,ref o),half);
            if(o!=sequence.Length) return false;
            using(var privateKey=RSA.Create()) using(var publicKey=RSACertificateExtensions.GetRSAPublicKey(c)) {
                privateKey.ImportParameters(p); var proof=Encoding.ASCII.GetBytes("Office Helper Portable private key match");
                var signature=privateKey.SignData(proof,HashAlgorithmName.SHA256,RSASignaturePadding.Pkcs1);
                return publicKey!=null && publicKey.VerifyData(proof,signature,HashAlgorithmName.SHA256,RSASignaturePadding.Pkcs1);
            }
        }
    }
}
'@
}

function Resolve-HelperCertificateDirectory {
    param([string]$StateDirectory)
    if ([string]::IsNullOrWhiteSpace($StateDirectory)) {
        if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA is unavailable.' }
        $StateDirectory = Join-Path $env:LOCALAPPDATA 'OfficeHelperPortable\certificates'
    }
    return [IO.Path]::GetFullPath($StateDirectory)
}

function Get-HelperCertificateClock { return [DateTime]::UtcNow }

function Read-HelperCertificateManifest {
    param([string]$StateDirectory)
    $path = Join-Path $StateDirectory 'manifest.json'
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $manifest = [IO.File]::ReadAllText($path) | ConvertFrom-Json -ErrorAction Stop
    if ($manifest.schemaVersion -ne 1 -or $manifest.activeId -notmatch '^[a-f0-9]{32}$') { throw 'Invalid certificate manifest.' }
    $ids = @{}
    foreach ($record in @($manifest.certificates)) {
        if ($record.id -notmatch '^[a-f0-9]{32}$' -or $record.thumbprint -notmatch '^[A-F0-9]{40}$' -or $ids.ContainsKey($record.id)) { throw 'Invalid certificate record.' }
        $ids[$record.id] = $true
    }
    if (-not $ids.ContainsKey($manifest.activeId)) { throw 'Active certificate record missing.' }
    return $manifest
}

function Get-HelperCertificatePaths {
    param([string]$StateDirectory, [string]$Id)
    if ($Id -notmatch '^[a-f0-9]{32}$') { throw 'Invalid certificate version.' }
    $root = [IO.Path]::GetFullPath($StateDirectory).TrimEnd('\')
    $directory = [IO.Path]::GetFullPath((Join-Path $root $Id))
    if (-not $directory.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Certificate path escaped state directory.' }
    if ((Test-Path -LiteralPath $directory) -and ((Get-Item -LiteralPath $directory -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Certificate version cannot be a reparse point.' }
    return @{ directory=$directory; certificatePath=(Join-Path $directory 'localhost.crt'); privateKeyPath=(Join-Path $directory 'localhost.key') }
}

function Read-HelperPendingCertificate {
    param([string]$StateDirectory)
    $path=Join-Path $StateDirectory 'pending.json'
    if (-not [IO.File]::Exists($path)) { return $null }
    $record=[IO.File]::ReadAllText($path) | ConvertFrom-Json -ErrorAction Stop
    if ($record.id -notmatch '^[a-f0-9]{32}$' -or $record.thumbprint -notmatch '^[A-F0-9]{40}$') { throw 'Invalid pending certificate record.' }
    return $record
}

function Clear-HelperPendingCertificate {
    param([string]$StateDirectory,$Manifest)
    $pending=Read-HelperPendingCertificate $StateDirectory
    if (-not $pending) { return }
    if (-not $Manifest -or $Manifest.activeId -ne $pending.id) {
        Remove-HelperCertificateTrust $pending.thumbprint
        Remove-HelperCertificateFiles $StateDirectory $pending.id
    }
    [IO.File]::Delete((Join-Path $StateDirectory 'pending.json'))
}

function Get-HelperTrustedCertificate {
    param([string]$Thumbprint)
    if ($Thumbprint -notmatch '^[A-F0-9]{40}$') { throw 'Invalid certificate thumbprint.' }
    $store = New-Object Security.Cryptography.X509Certificates.X509Store('Root','CurrentUser')
    try {
        $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        foreach ($certificate in $store.Certificates.Find([Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,$Thumbprint,$false)) {
            if ($certificate.Thumbprint -eq $Thumbprint) { return $certificate }
        }
        return $null
    } finally { $store.Close() }
}

function Add-HelperCertificateTrust {
    param([Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)
    if (-not [OfficeHelperPortable.CertificateCrypto]::IsProductCertificate($Certificate)) { throw 'Refusing to trust an unrelated certificate.' }
    $publicOnly = New-Object Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,$Certificate.RawData)
    $store = New-Object Security.Cryptography.X509Certificates.X509Store('Root','CurrentUser')
    try { $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite); $store.Add($publicOnly) }
    finally { $store.Close(); $publicOnly.Dispose() }
}

function Remove-HelperCertificateTrust {
    param([string]$Thumbprint)
    $certificate = Get-HelperTrustedCertificate $Thumbprint
    if (-not $certificate) { return }
    if (-not [OfficeHelperPortable.CertificateCrypto]::IsProductCertificate($certificate)) { throw 'Refusing to remove an unrelated certificate.' }
    $store = New-Object Security.Cryptography.X509Certificates.X509Store('Root','CurrentUser')
    try { $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite); $store.Remove($certificate) }
    finally { $store.Close(); $certificate.Dispose() }
}

function Get-HelperCertificateStatus {
    [CmdletBinding()]
    param([string]$StateDirectory)
    $root = Resolve-HelperCertificateDirectory $StateDirectory
    $status = [ordered]@{ valid=$false; trusted=$false; expiresAt=$null; needsRenewal=$true; remainingDays=0; thumbprint=$null; certificatePath=$null; privateKeyPath=$null; reason='missing'; stateDirectory=$root }
    $certificate = $null
    $trusted = $null
    try {
        $manifest = Read-HelperCertificateManifest $root
        if (-not $manifest) { return [pscustomobject]$status }
        $record = @($manifest.certificates | Where-Object { $_.id -eq $manifest.activeId })[0]
        $paths = Get-HelperCertificatePaths $root $record.id
        $status.thumbprint = $record.thumbprint
        $status.certificatePath = $paths.certificatePath
        $status.privateKeyPath = $paths.privateKeyPath
        foreach ($path in @($paths.certificatePath,$paths.privateKeyPath)) {
            if (-not [IO.File]::Exists($path)) { $status.reason='missing-material'; return [pscustomobject]$status }
            if ((Get-Item -LiteralPath $path -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Certificate files cannot be reparse points.' }
        }
        $certificate = [OfficeHelperPortable.CertificateCrypto]::ReadCertificate($paths.certificatePath)
        if ($certificate.Thumbprint -ne $record.thumbprint -or -not [OfficeHelperPortable.CertificateCrypto]::IsProductCertificate($certificate)) { $status.reason='invalid-certificate'; return [pscustomobject]$status }
        $status.expiresAt = $certificate.NotAfter.ToUniversalTime().ToString('o')
        if (-not [OfficeHelperPortable.CertificateCrypto]::MatchesPrivateKey($certificate,$paths.privateKeyPath)) { $status.reason='key-mismatch'; return [pscustomobject]$status }
        $now = Get-HelperCertificateClock
        $status.remainingDays = [Math]::Floor(($certificate.NotAfter.ToUniversalTime()-$now).TotalDays)
        $status.valid = $certificate.NotBefore.ToUniversalTime() -le $now -and $certificate.NotAfter.ToUniversalTime() -gt $now
        $status.needsRenewal = (-not $status.valid) -or $certificate.NotAfter.ToUniversalTime() -le $now.AddDays(7)
        $trusted = Get-HelperTrustedCertificate $record.thumbprint
        $status.trusted = $null -ne $trusted -and [Convert]::ToBase64String($trusted.RawData) -eq [Convert]::ToBase64String($certificate.RawData)
        $status.reason = if (-not $status.valid) { 'expired-or-not-yet-valid' } elseif (-not $status.trusted) { 'untrusted' } elseif ($status.needsRenewal) { 'renewal-due' } else { 'ready' }
    } catch { $status.valid=$false; $status.needsRenewal=$true; $status.reason='invalid-state: ' + $_.Exception.Message }
    finally { if ($certificate) { $certificate.Dispose() }; if ($trusted) { $trusted.Dispose() } }
    return [pscustomobject]$status
}

function Initialize-HelperCertificateDirectory {
    param([string]$StateDirectory)
    [void][IO.Directory]::CreateDirectory($StateDirectory)
    if ((Get-Item -LiteralPath $StateDirectory -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Certificate state cannot be a reparse point.' }
    $acl = [IO.Directory]::GetAccessControl($StateDirectory,[Security.AccessControl.AccessControlSections]::Access)
    $acl.SetAccessRuleProtection($true,$false)
    foreach ($existing in @($acl.Access)) { $acl.RemoveAccessRuleSpecific($existing) }
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User
    $rule = New-Object Security.AccessControl.FileSystemAccessRule($sid,'FullControl','ContainerInherit, ObjectInherit','None','Allow')
    $acl.AddAccessRule($rule)
    [IO.Directory]::SetAccessControl($StateDirectory,$acl)
}

function New-HelperCertificateMaterial {
    param([string]$VersionDirectory)
    $certificate = $null
    try {
        $now = Get-HelperCertificateClock
        # The Windows PKI cmdlet interprets these parameters as local wall time.
        # Add the duration in UTC first so a DST transition does not shorten it.
        $notBefore = $now.AddMinutes(-5).ToLocalTime()
        $notAfter = $now.AddDays(30).ToLocalTime()
        $certificate = New-SelfSignedCertificate -Type Custom -Subject 'CN=Office Helper Portable localhost' -FriendlyName 'Office Helper Portable localhost' -CertStoreLocation 'Cert:\CurrentUser\My' -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -Provider 'Microsoft Software Key Storage Provider' -KeyUsage DigitalSignature,KeyEncipherment -NotBefore $notBefore -NotAfter $notAfter -TextExtension @('2.5.29.17={text}DNS=localhost&IPAddress=127.0.0.1','2.5.29.19={critical}{text}ca=false','2.5.29.37={text}1.3.6.1.5.5.7.3.1') -ErrorAction Stop
        [OfficeHelperPortable.CertificateCrypto]::WriteMaterial($certificate,(Join-Path $VersionDirectory 'localhost.crt'),(Join-Path $VersionDirectory 'localhost.key'))
        return $certificate.Thumbprint
    } finally {
        if ($certificate) {
            # Only the newly issued, exact personal-store certificate/key is removed.
            try {
                $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
                try {
                    if ($rsa -is [Security.Cryptography.RSACng]) { $rsa.Key.Delete() }
                    elseif ($rsa -is [Security.Cryptography.RSACryptoServiceProvider]) { $rsa.PersistKeyInCsp=$false; $rsa.Clear() }
                } finally { if ($rsa) { $rsa.Dispose() } }
            } finally {
                $store = New-Object Security.Cryptography.X509Certificates.X509Store('My','CurrentUser')
                try { $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite); $store.Remove($certificate) }
                finally { $store.Close(); $certificate.Dispose() }
            }
        }
    }
}

function Write-HelperCertificateManifest {
    param([string]$StateDirectory, $Manifest)
    $path = Join-Path $StateDirectory 'manifest.json'
    $temporary = Join-Path $StateDirectory ('manifest-' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $backup = $temporary + '.previous'
    try {
        [IO.File]::WriteAllText($temporary,($Manifest | ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
        if ([IO.File]::Exists($path)) { [IO.File]::Replace($temporary,$path,$backup) } else { [IO.File]::Move($temporary,$path) }
    } finally {
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
        if ([IO.File]::Exists($backup)) { [IO.File]::Delete($backup) }
    }
}

function Remove-HelperCertificateFiles {
    param([string]$StateDirectory,[string]$Id)
    $paths = Get-HelperCertificatePaths $StateDirectory $Id
    foreach ($path in @($paths.certificatePath,$paths.privateKeyPath)) { if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) } }
    # Nonrecursive deletion avoids removing unexpected files or following links.
    if ([IO.Directory]::Exists($paths.directory)) { [IO.Directory]::Delete($paths.directory,$false) }
}

function Ensure-HelperCertificate {
    [CmdletBinding()]
    param([string]$StateDirectory,[switch]$AllowTrust,[switch]$ForceRenew)
    $root = Resolve-HelperCertificateDirectory $StateDirectory
    $status = Get-HelperCertificateStatus $root
    if ($status.valid -and $status.trusted -and -not $status.needsRenewal -and -not $ForceRenew -and -not [IO.File]::Exists((Join-Path $root 'pending.json'))) { return $status }
    if (-not $AllowTrust) { throw 'TRUST_REQUIRED: Explicit consent is required to create, renew, or trust the portable localhost certificate.' }
    Initialize-HelperCertificateDirectory $root
    $lock = [IO.File]::Open((Join-Path $root 'operation.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    try {
        $status = Get-HelperCertificateStatus $root
        $manifest = Read-HelperCertificateManifest $root
        Clear-HelperPendingCertificate $root $manifest
        if ($status.valid -and -not $status.needsRenewal -and -not $ForceRenew) {
            if ($status.trusted) { return $status }
            $certificate = [OfficeHelperPortable.CertificateCrypto]::ReadCertificate($status.certificatePath)
            try {
                Add-HelperCertificateTrust $certificate
                $checked = Get-HelperCertificateStatus $root
                if (-not $checked.valid -or -not $checked.trusted) { throw 'Trust verification failed.' }
                return $checked
            } catch { Remove-HelperCertificateTrust $status.thumbprint; throw }
            finally { $certificate.Dispose() }
        }
        $id = [Guid]::NewGuid().ToString('N')
        $paths = Get-HelperCertificatePaths $root $id
        [void][IO.Directory]::CreateDirectory($paths.directory)
        $thumbprint=$null; $certificate=$null; $committed=$false; $trustAttempted=$false
        try {
            $thumbprint = New-HelperCertificateMaterial $paths.directory
            $certificate = [OfficeHelperPortable.CertificateCrypto]::ReadCertificate($paths.certificatePath)
            $now = Get-HelperCertificateClock
            if ($certificate.Thumbprint -ne $thumbprint -or -not [OfficeHelperPortable.CertificateCrypto]::IsProductCertificate($certificate) -or -not [OfficeHelperPortable.CertificateCrypto]::MatchesPrivateKey($certificate,$paths.privateKeyPath) -or $certificate.NotBefore.ToUniversalTime() -gt $now -or $certificate.NotAfter.ToUniversalTime() -lt $now.AddDays(7)) { throw 'New certificate verification failed.' }
            # Record exact ownership before store changes, including interrupted setup.
            $newRecord = [pscustomobject]@{id=$id;thumbprint=$thumbprint}
            [IO.File]::WriteAllText((Join-Path $root 'pending.json'),($newRecord | ConvertTo-Json),(New-Object Text.UTF8Encoding($false)))
            $trustAttempted=$true
            Add-HelperCertificateTrust $certificate
            $trusted=Get-HelperTrustedCertificate $thumbprint
            try { if (-not $trusted -or [Convert]::ToBase64String($trusted.RawData) -ne [Convert]::ToBase64String($certificate.RawData)) { throw 'New certificate trust verification failed.' } }
            finally { if ($trusted) { $trusted.Dispose() } }
            $oldRecords = if ($manifest) { @($manifest.certificates) } else { @() }
            $next = [pscustomobject]@{schemaVersion=1;activeId=$id;certificates=@($newRecord)+$oldRecords}
            Write-HelperCertificateManifest $root $next
            $committed=$true
        } catch {
            $failure=$_
            if (-not $committed) {
                if ($trustAttempted -and $thumbprint) {
                    try { Remove-HelperCertificateTrust $thumbprint }
                    catch { throw ('Certificate setup failed and trust rollback needs retry. pending.json retains the exact product thumbprint. ' + $_.Exception.Message) }
                }
                Remove-HelperCertificateFiles $root $id
                [IO.File]::Delete((Join-Path $root 'pending.json'))
            }
            throw $failure
        } finally { if ($certificate) { $certificate.Dispose() } }
        [IO.File]::Delete((Join-Path $root 'pending.json'))
        $remaining = @($newRecord)
        $cleanupWarnings=@()
        foreach ($record in $oldRecords) {
            try { Remove-HelperCertificateTrust $record.thumbprint; Remove-HelperCertificateFiles $root $record.id }
            catch { $remaining += $record; $cleanupWarnings += 'Old product certificate cleanup deferred: ' + $_.Exception.Message }
        }
        $next.certificates=$remaining
        try { Write-HelperCertificateManifest $root $next }
        catch { $cleanupWarnings += 'New certificate is active; cleanup record compaction deferred: ' + $_.Exception.Message }
        $result=Get-HelperCertificateStatus $root
        $result | Add-Member -NotePropertyName cleanupWarnings -NotePropertyValue @($cleanupWarnings)
        return $result
    } finally { $lock.Dispose() }
}

function Remove-HelperCertificate {
    [CmdletBinding()]
    param([string]$StateDirectory)
    $root = Resolve-HelperCertificateDirectory $StateDirectory
    if (-not [IO.Directory]::Exists($root)) { return [pscustomobject]@{removed=$true;stateDirectory=$root} }
    $lock = [IO.File]::Open((Join-Path $root 'operation.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    try {
        $manifest = Read-HelperCertificateManifest $root
        Clear-HelperPendingCertificate $root $manifest
        if ($manifest) {
            foreach ($record in @($manifest.certificates)) { Remove-HelperCertificateTrust $record.thumbprint; Remove-HelperCertificateFiles $root $record.id }
            [IO.File]::Delete((Join-Path $root 'manifest.json'))
        }
        return [pscustomobject]@{removed=$true;stateDirectory=$root}
    } finally { $lock.Dispose() }
}
