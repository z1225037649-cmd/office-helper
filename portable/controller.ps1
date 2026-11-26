param(
  [ValidateSet('Inspect','Prepare','Remove','OpenWord','OpenExcel','OpenWps')][string]$Action = 'Inspect',
  [ValidateSet('Wps','Office','Both')][string]$Hosts = 'Both',
  [switch]$AllowTrust,
  [string]$PackageRoot = (Split-Path -Parent $PSScriptRoot),
  [string]$StateDirectory = (Join-Path $env:LOCALAPPDATA 'OfficeHelperPortable')
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
. (Join-Path $PSScriptRoot 'certificates.ps1')
$officeId = 'f5202c69-1f59-4e09-b510-829a3748a31b'
$wpsNames = @('OfficeHelperPortableWriter','OfficeHelperPortableSheet')
$registryPath = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$settingsPath = Join-Path $StateDirectory 'settings.json'
$certificateDirectory = Join-Path $StateDirectory 'certificates'

function Read-Settings {
  if (Test-Path -LiteralPath $settingsPath) { return Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json }
  return [PSCustomObject]@{ allowCertificateMaintenance=$false; hosts='Both' }
}
function Write-Json($Path,$Value) {
  $temp = $Path + '.tmp'
  [IO.File]::WriteAllText($temp,($Value | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
  Move-Item -LiteralPath $temp -Destination $Path -Force
}
function Find-Application([string]$Executable) {
  foreach ($base in @('HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths','HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths')) {
    $key = Get-Item -LiteralPath (Join-Path $base $Executable) -ErrorAction SilentlyContinue
    if ($key) { $value = $key.GetValue(''); if ($value -and (Test-Path -LiteralPath $value)) { return $value } }
  }
  if ($Executable -in @('WINWORD.EXE','EXCEL.EXE')) {
    foreach ($base in @($env:ProgramFiles,${env:ProgramFiles(x86)})) {
      if (-not $base) { continue }
      $value = Join-Path $base "Microsoft Office\root\Office16\$Executable"
      if (Test-Path -LiteralPath $value) { return $value }
    }
  }
  return $null
}
function Merge-Wps([bool]$Remove) {
  $directory = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
  $path = Join-Path $directory 'publish.xml'
  if ($Remove -and -not (Test-Path -LiteralPath $path)) { return }
  $xml = New-Object Xml.XmlDocument
  $xml.PreserveWhitespace = $true
  if (Test-Path -LiteralPath $path) { $xml.Load($path) } else { $xml.LoadXml('<jsplugins/>') }
  if ($xml.DocumentElement.Name -ne 'jsplugins') { throw 'WPS_CONFIG_UNEXPECTED' }
  $backup = Join-Path $StateDirectory 'backups'
  New-Item -ItemType Directory -Path $backup -Force | Out-Null
  if (Test-Path -LiteralPath $path) { Copy-Item -LiteralPath $path -Destination (Join-Path $backup ('publish-' + [DateTime]::UtcNow.Ticks + '.xml')) }
  foreach ($name in $wpsNames) {
    foreach ($entry in @($xml.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.GetAttribute('name') -eq $name })) { [void]$xml.DocumentElement.RemoveChild($entry) }
  }
  if (-not $Remove) {
    for ($i=0; $i -lt 2; $i++) {
      $entry = $xml.CreateElement('jspluginonline')
      $entry.SetAttribute('name',$wpsNames[$i]); $entry.SetAttribute('type',@('wps','et')[$i])
      $entry.SetAttribute('url',('http://127.0.0.1:18765/wps/' + @('writer','sheet')[$i] + '/'))
      $entry.SetAttribute('enable','enable_dev'); $entry.SetAttribute('install','null'); $entry.SetAttribute('debug','')
      [void]$xml.DocumentElement.AppendChild($entry)
    }
  }
  New-Item -ItemType Directory -Path $directory -Force | Out-Null
  $temp = $path + '.office-helper-portable.tmp'
  $xml.Save($temp)
  Move-Item -LiteralPath $temp -Destination $path -Force
}
function Assert-PortFree([int]$Port,[string]$Address) {
  $listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Parse($Address),$Port)
  try { $listener.Start() } catch { throw "PORT_IN_USE:$Port" } finally { $listener.Stop() }
}
function Assert-Package([bool]$UseWps,[bool]$UseOffice) {
  $required = @('site\index.html','runtime\node.exe','app\server.mjs','samples\OfficeHelper-Sample.docx','samples\OfficeHelper-Sample.xlsx')
  if ($UseOffice) { $required += 'manifests\office.xml' }
  if ($UseWps) { $required += @('site\wps\writer\ribbon.xml','site\wps\sheet\ribbon.xml') }
  foreach ($file in $required) { if (-not (Test-Path -LiteralPath (Join-Path $PackageRoot $file) -PathType Leaf)) { throw 'PACKAGE_INCOMPLETE' } }
  if ($UseOffice) {
    $manifest = New-Object Xml.XmlDocument
    $manifest.Load((Join-Path $PackageRoot 'manifests\office.xml'))
    if ($manifest.DocumentElement.LocalName -ne 'OfficeApp' -or $manifest.DocumentElement.Id -ne $officeId) { throw 'PACKAGE_INCOMPLETE' }
  }
  if ($UseWps) {
    $publish = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons\publish.xml'
    if (Test-Path -LiteralPath $publish) {
      $xml = New-Object Xml.XmlDocument
      try { $xml.Load($publish) } catch { throw 'WPS_CONFIG_UNEXPECTED' }
      if ($xml.DocumentElement.Name -ne 'jsplugins') { throw 'WPS_CONFIG_UNEXPECTED' }
    }
  }
}
try {
  $settings = if ($Action -eq 'Remove') { $null } else { Read-Settings }
  $applications = @{ word=(Find-Application 'WINWORD.EXE'); excel=(Find-Application 'EXCEL.EXE'); wps=(Find-Application 'wps.exe') }
  if ($Action -eq 'Inspect') {
    $result = @{ ok=$true; applications=$applications; hosts=$settings.hosts; allowCertificateMaintenance=[bool]$settings.allowCertificateMaintenance; certificate=(Get-HelperCertificateStatus -StateDirectory $certificateDirectory) }
  } elseif ($Action -eq 'Prepare') {
    $useOffice = $Hosts -in @('Office','Both')
    $useWps = $Hosts -in @('Wps','Both')
    if ($useOffice -and -not ($applications.word -or $applications.excel)) { throw 'OFFICE_NOT_FOUND' }
    if ($useWps -and -not $applications.wps) { throw 'WPS_NOT_FOUND' }
    Assert-Package $useWps $useOffice
    Assert-PortFree 18765 '127.0.0.1'
    if ($useOffice) { Assert-PortFree 18766 '127.0.0.1'; Assert-PortFree 18766 '::1' }
    New-Item -ItemType Directory -Path $StateDirectory -Force | Out-Null
    $cert = $null
    if ($useOffice) {
      $permitted = $AllowTrust -or $settings.allowCertificateMaintenance
      $cert = Ensure-HelperCertificate -StateDirectory $certificateDirectory -AllowTrust:$permitted
      $manifest = Join-Path $StateDirectory 'office-manifest.xml'
      Copy-Item -LiteralPath (Join-Path $PackageRoot 'manifests\office.xml') -Destination $manifest -Force
      New-Item -Path $registryPath -Force | Out-Null
      New-ItemProperty -Path $registryPath -Name $officeId -Value $manifest -PropertyType String -Force | Out-Null
    }
    if ($useWps) { Merge-Wps $false }
    $sampleDirectory = Join-Path $StateDirectory 'samples'
    New-Item -ItemType Directory -Path $sampleDirectory -Force | Out-Null
    foreach ($file in @('OfficeHelper-Sample.docx','OfficeHelper-Sample.xlsx')) {
      if (-not (Test-Path -LiteralPath (Join-Path $sampleDirectory $file))) { Copy-Item -LiteralPath (Join-Path $PackageRoot "samples\$file") -Destination $sampleDirectory }
    }
    Write-Json $settingsPath @{ allowCertificateMaintenance=([bool]$settings.allowCertificateMaintenance -or [bool]$AllowTrust); hosts=$Hosts }
    $config = @{ siteDirectory=(Join-Path $PackageRoot 'site'); httpPort=18765; httpsPort=18766; office=$useOffice; certificatePath=$null; privateKeyPath=$null }
    if ($useOffice) { $config.certificatePath=$cert.certificatePath; $config.privateKeyPath=$cert.privateKeyPath }
    $configPath = Join-Path $StateDirectory 'runtime.json'
    Write-Json $configPath $config
    $result = @{ ok=$true; configPath=$configPath; certificate=$cert; hosts=$Hosts }
  } elseif ($Action -eq 'Remove') {
    New-Item -ItemType Directory -Path $StateDirectory -Force | Out-Null
    $removeErrors = @()
    try { Merge-Wps $true } catch { $removeErrors += 'WPS: ' + $_.Exception.Message }
    try {
      if (Get-ItemProperty -Path $registryPath -Name $officeId -ErrorAction SilentlyContinue) { Remove-ItemProperty -Path $registryPath -Name $officeId }
    } catch { $removeErrors += 'Office: ' + $_.Exception.Message }
    try { Remove-HelperCertificate -StateDirectory $certificateDirectory | Out-Null } catch { $removeErrors += 'Certificate: ' + $_.Exception.Message }
    try { Write-Json $settingsPath @{ allowCertificateMaintenance=$false; hosts='Both' } } catch { $removeErrors += 'Settings: ' + $_.Exception.Message }
    if ($removeErrors.Count -gt 0) { throw ('REMOVE_PARTIAL: ' + ($removeErrors -join '; ')) }
    $result = @{ ok=$true; removed=$true }
  } else {
    $name = @{ OpenWord='word'; OpenExcel='excel'; OpenWps='wps' }[$Action]
    $executable = $applications[$name]
    if (-not $executable) { throw 'APPLICATION_NOT_FOUND' }
    if ($name -eq 'wps') { Start-Process -FilePath $executable -WindowStyle Hidden }
    else {
      $extension = if ($name -eq 'word') { 'docx' } else { 'xlsx' }
      $sample = Join-Path $StateDirectory "samples\OfficeHelper-Sample.$extension"
      if (-not (Test-Path -LiteralPath $sample)) { throw 'START_HELPER_FIRST' }
      Start-Process -FilePath $executable -ArgumentList ('"' + $sample + '"') -WindowStyle Hidden
    }
    $result = @{ ok=$true }
  }
  $result | ConvertTo-Json -Depth 8 -Compress
} catch {
  @{ ok=$false; error=$_.Exception.Message } | ConvertTo-Json -Compress
  exit 1
}
