# Windows PowerShell 5.1. Run the actual controller in a child process with fake
# certificate APIs, registry cmdlets, app discovery and loopback-port probes.
# All filesystem writes go to a unique TEMP tree. No OS registry/trust changes.
$ErrorActionPreference='Stop'
$controllerSource=Join-Path $PSScriptRoot '..\..\portable\controller.ps1'
$controllerText=[IO.File]::ReadAllText($controllerSource)
$testRoot=Join-Path $env:TEMP ('OfficeHelperControllerTests-'+[Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
$script:passed=0
$script:failures=New-Object 'Collections.Generic.List[string]'
function Assert-True { param([bool]$Condition,[string]$Message) if(-not $Condition){throw $Message} }
function Write-TestFile { param([string]$Path,[string]$Content) [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path)); [IO.File]::WriteAllText($Path,$Content,(New-Object Text.UTF8Encoding($false))) }
function Test-Case { param([string]$Name,[scriptblock]$Body) try { & $Body; $script:passed++; Write-Output ('PASS: '+$Name) } catch { $script:failures.Add($Name+': '+$_.Exception.Message); Write-Output ('FAIL: '+$Name+': '+$_.Exception.Message) } }

$fakeCertificates=@'
function Get-HelperCertificateStatus {
  param([string]$StateDirectory)
  $p=Join-Path $StateDirectory 'fake.json'
  if(Test-Path -LiteralPath $p){return Get-Content -LiteralPath $p -Raw | ConvertFrom-Json}
  return [pscustomobject]@{valid=$false;trusted=$false;needsRenewal=$true;expiresAt=$null;certificatePath=$null;privateKeyPath=$null}
}
function Ensure-HelperCertificate {
  param([string]$StateDirectory,[switch]$AllowTrust,[switch]$ForceRenew)
  [IO.File]::AppendAllText((Join-Path $env:HELPER_TEST_ROOT 'trace.txt'),('ensure:'+ [bool]$AllowTrust+"`n"))
  if(-not $AllowTrust){throw 'TRUST_REQUIRED'}
  [void][IO.Directory]::CreateDirectory($StateDirectory)
  $value=[pscustomobject]@{valid=$true;trusted=$true;needsRenewal=$false;expiresAt='2030-01-01T00:00:00Z';certificatePath=(Join-Path $StateDirectory 'localhost.crt');privateKeyPath=(Join-Path $StateDirectory 'localhost.key')}
  [IO.File]::WriteAllText((Join-Path $StateDirectory 'fake.json'),($value|ConvertTo-Json))
  return $value
}
function Remove-HelperCertificate {
  param([string]$StateDirectory)
  [IO.File]::AppendAllText((Join-Path $env:HELPER_TEST_ROOT 'trace.txt'),"remove-certificate`n")
  $p=Join-Path $StateDirectory 'fake.json'
  if([IO.File]::Exists($p)){[IO.File]::Delete($p)}
  return [pscustomobject]@{removed=$true}
}
'@

$wrapper=@'
param([string]$TestRoot,[string]$PackageRoot,[string]$Action,[string]$Hosts,[switch]$AllowTrust,[switch]$AllowAppLaunch)
$ErrorActionPreference='Stop'
$env:HELPER_TEST_ROOT=$TestRoot
$env:HELPER_TEST_ALLOW_LAUNCH=[string][bool]$AllowAppLaunch
$env:APPDATA=Join-Path $TestRoot 'roaming'
$env:LOCALAPPDATA=Join-Path $TestRoot 'local'
$env:ProgramFiles=Join-Path $TestRoot 'programs'
${env:ProgramFiles(x86)}=Join-Path $TestRoot 'programs-x86'
$global:apps=Get-Content -LiteralPath (Join-Path $TestRoot 'apps.json') -Raw | ConvertFrom-Json
function Read-TestRegistry { $p=Join-Path $env:HELPER_TEST_ROOT 'registry.json'; if(Test-Path -LiteralPath $p){return Get-Content -LiteralPath $p -Raw|ConvertFrom-Json}; return [pscustomobject]@{} }
function Save-TestRegistry($Value){[IO.File]::WriteAllText((Join-Path $env:HELPER_TEST_ROOT 'registry.json'),($Value|ConvertTo-Json))}
function Get-Item {
  [CmdletBinding()]param([string]$LiteralPath,[string]$Path,[switch]$Force)
  $p=if($LiteralPath){$LiteralPath}else{$Path}
  if($p -match '^HK(CU|LM):'){
    $name=Split-Path -Leaf $p; $property=$global:apps.PSObject.Properties[$name]
    if($property -and $property.Value){$item=[pscustomobject]@{AppPath=[string]$property.Value}; $item|Add-Member ScriptMethod GetValue {param($name) return $this.AppPath}; return $item}; return $null
  }
  return Microsoft.PowerShell.Management\Get-Item @PSBoundParameters
}
function New-Item {
  [CmdletBinding()]param([string]$Path,[string]$ItemType,[switch]$Force)
  if($Path -match '^HK(CU|LM):'){if($Path -notlike 'HKCU:*'){throw 'Unexpected machine registry write'};return}
  return Microsoft.PowerShell.Management\New-Item @PSBoundParameters
}
function New-ItemProperty {
  [CmdletBinding()]param([string]$Path,[string]$Name,$Value,[string]$PropertyType,[switch]$Force)
  if($Path -ne 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'){throw 'Unexpected registry target'}
  $r=Read-TestRegistry; $r|Add-Member NoteProperty $Name $Value -Force; Save-TestRegistry $r
  [IO.File]::AppendAllText((Join-Path $env:HELPER_TEST_ROOT 'trace.txt'),"office-register`n")
}
function Remove-ItemProperty {
  [CmdletBinding()]param([string]$Path,[string]$Name)
  if($Path -ne 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'){throw 'Unexpected registry target'}
  $r=Read-TestRegistry; $r.PSObject.Properties.Remove($Name); Save-TestRegistry $r
  [IO.File]::AppendAllText((Join-Path $env:HELPER_TEST_ROOT 'trace.txt'),"office-remove`n")
}
function Get-ItemProperty {
  [CmdletBinding()]param([string]$Path,[string]$Name)
  if($Path -ne 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'){throw 'Unexpected registry read'}
  $r=Read-TestRegistry
  if($r.PSObject.Properties[$Name]){return $r};return $null
}
function New-Object {
  [CmdletBinding()]param([Parameter(Position=0)][string]$TypeName,[Parameter(Position=1)][object[]]$ArgumentList)
  if($TypeName -eq 'Net.Sockets.TcpListener'){
    $v=[pscustomobject]@{}
    $v|Add-Member ScriptMethod Start {if([IO.File]::Exists((Join-Path $env:HELPER_TEST_ROOT 'port-busy'))){throw 'Injected port conflict'}}
    $v|Add-Member ScriptMethod Stop {}; return $v
  }
  return Microsoft.PowerShell.Utility\New-Object @PSBoundParameters
}
function Start-Process {
  [CmdletBinding()]param([string]$FilePath,[string[]]$ArgumentList,[string]$WindowStyle)
  if($env:HELPER_TEST_ALLOW_LAUNCH -ne 'True'){throw 'Tests must not start real applications.'}
  $capture=@{filePath=$FilePath;arguments=$ArgumentList;windowStyle=$WindowStyle;windowStyleSpecified=$PSBoundParameters.ContainsKey('WindowStyle')}
  [IO.File]::WriteAllText((Join-Path $env:HELPER_TEST_ROOT 'launch.json'),($capture|ConvertTo-Json -Compress))
}
& (Join-Path $TestRoot 'controller\controller.ps1') -Action $Action -Hosts $Hosts -AllowTrust:$AllowTrust -PackageRoot $PackageRoot -StateDirectory (Join-Path $env:LOCALAPPDATA 'OfficeHelperPortable')
exit $LASTEXITCODE
'@

function New-Scenario {
  $root=Join-Path $testRoot ([Guid]::NewGuid().ToString('N'))
  $package=Join-Path $root 'package one'
  foreach($path in @('site\index.html','runtime\node.exe','app\server.mjs','site\wps\writer\ribbon.xml','site\wps\sheet\ribbon.xml','samples\OfficeHelper-Sample.docx','samples\OfficeHelper-Sample.xlsx')){Write-TestFile (Join-Path $package $path) ('fixture:'+ $path)}
  Write-TestFile (Join-Path $package 'manifests\office.xml') '<OfficeApp><Id>f5202c69-1f59-4e09-b510-829a3748a31b</Id></OfficeApp>'
  Write-TestFile (Join-Path $root 'controller\controller.ps1') $controllerText
  Write-TestFile (Join-Path $root 'controller\certificates.ps1') $fakeCertificates
  Write-TestFile (Join-Path $root 'wrapper.ps1') $wrapper
  $apps=@{}
  foreach($name in @('WINWORD.EXE','EXCEL.EXE','wps.exe')){$file=Join-Path $root ('fake-apps\'+$name);Write-TestFile $file ''; $apps[$name]=$file}
  Write-TestFile (Join-Path $root 'apps.json') ($apps|ConvertTo-Json)
  Write-TestFile (Join-Path $root 'registry.json') (@{'existing-addon'='C:\existing.xml';'26b5f5ec-853a-4c06-9c1c-8ac9ea33421a'='C:\dev.xml'}|ConvertTo-Json)
  $publish=Join-Path $root 'roaming\kingsoft\wps\jsaddons\publish.xml'
  Write-TestFile $publish '<jsplugins><!--keep-comment--><jspluginonline name="ExistingPlugin" url="https://example.test/"/><jspluginonline name="OfficeHelperWriter" url="http://127.0.0.1:5173/wps/writer/"/><jspluginonline name="OfficeHelperSheet" url="http://127.0.0.1:5173/wps/sheet/"/></jsplugins>'
  return [pscustomobject]@{root=$root;package=$package;publish=$publish;state=(Join-Path $root 'local\OfficeHelperPortable')}
}
function Invoke-ControllerTest {
  param($Scenario,[string]$Action='Prepare',[string]$Hosts='Both',[switch]$AllowTrust,[switch]$AllowAppLaunch,[string]$PackageRoot)
  if(-not $PackageRoot){$PackageRoot=$Scenario.package}
  $arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $Scenario.root 'wrapper.ps1'),'-TestRoot',$Scenario.root,'-PackageRoot',$PackageRoot,'-Action',$Action,'-Hosts',$Hosts)
  if($AllowTrust){$arguments+='-AllowTrust'}
  if($AllowAppLaunch){$arguments+='-AllowAppLaunch'}
  $output=& powershell.exe @arguments
  $code=$LASTEXITCODE
  try{$result=($output -join "`n")|ConvertFrom-Json}catch{throw ('Controller did not return JSON: '+($output -join ' '))}
  return [pscustomobject]@{exitCode=$code;data=$result}
}
function Read-Trace($Scenario){$p=Join-Path $Scenario.root 'trace.txt';if([IO.File]::Exists($p)){return [IO.File]::ReadAllText($p)};return ''}
function Read-Registry($Scenario){return [IO.File]::ReadAllText((Join-Path $Scenario.root 'registry.json'))|ConvertFrom-Json}
$portableId='f5202c69-1f59-4e09-b510-829a3748a31b'
try {
  Test-Case 'Inspect is read-only and reports discovered applications' {
    $s=New-Scenario; $result=Invoke-ControllerTest $s -Action Inspect
    Assert-True ($result.data.ok -and $result.data.applications.word -and $result.data.applications.wps) 'Installed apps should be reported.'
    Assert-True (-not [IO.Directory]::Exists($s.state) -and (Read-Trace $s) -eq '') 'Inspect must not mutate state or trust.'
  }
  Test-Case 'WPS registration preserves unrelated and development entries' {
    $s=New-Scenario; $r=Invoke-ControllerTest $s -Hosts Wps
    Assert-True $r.data.ok 'WPS prepare failed.'
    $xml=[xml][IO.File]::ReadAllText($s.publish)
    Assert-True ($xml.jsplugins.jspluginonline.Count -eq 5) 'Only two portable nodes should be added.'
    Assert-True ($xml.SelectSingleNode('//jspluginonline[@name="OfficeHelperWriter"]').GetAttribute('url') -like '*:5173/*') 'Development registration must remain.'
    Assert-True ([IO.File]::ReadAllText($s.publish).Contains('keep-comment')) 'Existing XML comments must remain.'
    Assert-True ((Read-Trace $s) -eq '') 'WPS-only prepare must not touch certificate or Office registration.'
    $again=Invoke-ControllerTest $s -Hosts Wps
    Assert-True ($again.data.ok -and ([xml][IO.File]::ReadAllText($s.publish)).jsplugins.jspluginonline.Count -eq 5) 'Repeated setup must not duplicate nodes.'
    Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $s.state 'backups') -File).Count -ge 1) 'Existing WPS config needs a backup.'
  }
  Test-Case 'Portable package can move while user samples and registrations stay stable' {
    $s=New-Scenario; $r=Invoke-ControllerTest $s -Hosts Wps
    $sample=Join-Path $s.state 'samples\OfficeHelper-Sample.docx';Write-TestFile $sample 'user-edited'
    $moved=Join-Path $s.root 'moved package';Copy-Item -LiteralPath $s.package -Destination $moved -Recurse
    $r=Invoke-ControllerTest $s -Hosts Wps -PackageRoot $moved
    $runtime=[IO.File]::ReadAllText((Join-Path $s.state 'runtime.json'))|ConvertFrom-Json
    Assert-True ($r.data.ok -and $runtime.siteDirectory -eq (Join-Path $moved 'site')) 'Runtime must use new package location.'
    Assert-True ([IO.File]::ReadAllText($sample) -eq 'user-edited') 'Existing sample must not be overwritten.'
  }
  Test-Case 'Missing selected host fails before configuration changes' {
    $s=New-Scenario;Write-TestFile (Join-Path $s.root 'apps.json') '{}'
    foreach($hosts in @('Office','Wps')){$r=Invoke-ControllerTest $s -Hosts $hosts -AllowTrust;Assert-True (-not $r.data.ok -and $r.data.error -like '*NOT_FOUND*') 'Missing host needs an explicit error.'}
    Assert-True (-not [IO.Directory]::Exists($s.state) -and (Read-Trace $s) -eq '') 'Missing host cannot trigger certificate setup.'
  }
  Test-Case 'A busy portable port stops preparation without touching configuration' {
    $s=New-Scenario;Write-TestFile (Join-Path $s.root 'port-busy') ''
    $r=Invoke-ControllerTest $s -AllowTrust
    Assert-True (-not $r.data.ok -and $r.data.error -like 'PORT_IN_USE:*') 'Port conflict should be explicit.'
    Assert-True (-not [IO.Directory]::Exists($s.state) -and (Read-Trace $s) -eq '') 'Port conflict cannot trigger trust or registration changes.'
  }
  Test-Case 'Office consent is required and persisted for later maintenance' {
    $s=New-Scenario;$r=Invoke-ControllerTest $s -Hosts Office
    Assert-True (-not $r.data.ok -and $r.data.error -eq 'TRUST_REQUIRED') ('Office must require consent; response: '+($r.data|ConvertTo-Json -Compress))
    $r=Invoke-ControllerTest $s -Hosts Office -AllowTrust
    Assert-True $r.data.ok 'Consented setup should succeed.'
    $settings=[IO.File]::ReadAllText((Join-Path $s.state 'settings.json'))|ConvertFrom-Json
    Assert-True $settings.allowCertificateMaintenance 'Consent must be persisted.'
    $r=Invoke-ControllerTest $s -Hosts Office
    Assert-True $r.data.ok 'Persisted consent should permit maintenance.'
    $registry=Read-Registry $s
    Assert-True ($registry.$portableId -eq (Join-Path $s.state 'office-manifest.xml') -and $registry.'existing-addon' -eq 'C:\existing.xml') 'Only portable HKCU value should be set.'
  }
  Test-Case 'Word and Excel launch their sample documents in visible windows' {
    foreach($app in @(
      @{action='OpenWord';name='WINWORD.EXE';sample='OfficeHelper-Sample.docx'},
      @{action='OpenExcel';name='EXCEL.EXE';sample='OfficeHelper-Sample.xlsx'}
    )){
      $s=New-Scenario
      $prepared=Invoke-ControllerTest $s -Hosts Office -AllowTrust
      Assert-True $prepared.data.ok 'Office setup for the launch test failed.'
      $opened=Invoke-ControllerTest $s -Action $app.action -Hosts Office -AllowAppLaunch
      Assert-True $opened.data.ok ($app.action+' should launch successfully.')
      $launch=Get-Content -LiteralPath (Join-Path $s.root 'launch.json') -Raw|ConvertFrom-Json
      $samplePath=Join-Path $s.state ('samples\'+$app.sample)
      Assert-True ($launch.filePath -eq (Get-Content -LiteralPath (Join-Path $s.root 'apps.json') -Raw|ConvertFrom-Json).$($app.name)) ($app.action+' should use the discovered office application.')
      Assert-True (($launch.arguments -join ' ') -eq ('"'+$samplePath+'"')) ($app.action+' should pass the quoted portable sample path.')
      Assert-True ($launch.windowStyleSpecified -and $launch.windowStyle -eq 'Normal') ($app.action+' must show an ordinary application window.')
    }
  }
  Test-Case 'Remove clears portable ownership only and preserves user files' {
    $s=New-Scenario;$null=Invoke-ControllerTest $s -AllowTrust
    $sample=Join-Path $s.state 'samples\OfficeHelper-Sample.docx';Write-TestFile $sample 'user-edited'
    $r=Invoke-ControllerTest $s -Action Remove
    $registry=Read-Registry $s;$xml=[xml][IO.File]::ReadAllText($s.publish)
    Assert-True ($r.data.ok -and -not $registry.PSObject.Properties[$portableId] -and $registry.'existing-addon') 'Uninstall should remove only portable Office registration.'
    Assert-True ($xml.jsplugins.jspluginonline.Count -eq 3 -and $xml.SelectSingleNode('//jspluginonline[@name="OfficeHelperWriter"]')) 'Uninstall must retain development WPS registration.'
    Assert-True ([IO.File]::ReadAllText($sample) -eq 'user-edited' -and (Read-Trace $s).Contains('remove-certificate')) 'Samples stay; dedicated certificate is removed.'
  }
  Test-Case 'Incomplete package is rejected before certificate or registration changes' {
    $s=New-Scenario;[IO.File]::Delete((Join-Path $s.package 'samples\OfficeHelper-Sample.xlsx'))
    $r=Invoke-ControllerTest $s -AllowTrust
    Assert-True (-not $r.data.ok) 'Missing sample must fail preparation.'
    Assert-True ((Read-Trace $s) -eq '') 'Incomplete package must be detected before trust or registrations are changed.'
  }
  Test-Case 'Malformed WPS config is rejected before Office changes in Both mode' {
    $s=New-Scenario;Write-TestFile $s.publish '<unexpected/>'
    $r=Invoke-ControllerTest $s -AllowTrust
    Assert-True (-not $r.data.ok) 'Unexpected WPS config should fail.'
    Assert-True ((Read-Trace $s) -eq '') 'WPS preflight must precede Office certificate/registry mutation.'
  }
  Test-Case 'WPS uninstall failure does not prevent Office and certificate cleanup' {
    $s=New-Scenario;$null=Invoke-ControllerTest $s -AllowTrust
    Write-TestFile $s.publish '<unexpected/>'
    $r=Invoke-ControllerTest $s -Action Remove
    $registry=Read-Registry $s
    Assert-True (-not $r.data.ok) 'Partial uninstall should report an error.'
    Assert-True (-not $registry.PSObject.Properties[$portableId] -and (Read-Trace $s).Contains('remove-certificate')) 'Independent Office and certificate cleanup must still run.'
  }
  Test-Case 'Damaged preferences do not block uninstalling owned configuration' {
    $s=New-Scenario;$prepared=Invoke-ControllerTest $s -AllowTrust
    Assert-True $prepared.data.ok 'Fixture setup must succeed.'
    Write-TestFile (Join-Path $s.state 'settings.json') '{broken-json'
    $r=Invoke-ControllerTest $s -Action Remove
    $registry=Read-Registry $s
    Assert-True ($r.data.ok -and -not $registry.PSObject.Properties[$portableId] -and (Read-Trace $s).Contains('remove-certificate')) 'Uninstall must clean owned registrations and certificates even when preferences are unreadable.'
  }
  Write-Output ('RESULT: '+$script:passed+' passed; '+$script:failures.Count+' failed. No real registry, certificates, hosts, or ports were modified.')
  if($script:failures.Count){throw ($script:failures -join "`n")}
} finally {
  $resolved=[IO.Path]::GetFullPath($testRoot);$prefix=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\OfficeHelperControllerTests-'
  if($resolved.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -and [IO.Directory]::Exists($resolved)){Remove-Item -LiteralPath $resolved -Recurse -Force}
}
