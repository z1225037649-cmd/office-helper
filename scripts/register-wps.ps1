param([ValidateSet('Install','Remove','Test')][string]$Action = 'Install')
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$addonDirectory = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$publishPath = Join-Path $addonDirectory 'publish.xml'
$names = @('OfficeHelperWriter','OfficeHelperSheet')
$xml = New-Object System.Xml.XmlDocument
$xml.PreserveWhitespace = $true
if (Test-Path -LiteralPath $publishPath) { $xml.Load($publishPath) } else { $xml.LoadXml('<?xml version="1.0" encoding="UTF-8"?><jsplugins></jsplugins>') }
if ($xml.DocumentElement.Name -ne 'jsplugins') { throw 'Unexpected WPS configuration. No changes made.' }
$backupDirectory = Join-Path $projectRoot '.local\wps-registration-backups'
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
if (Test-Path -LiteralPath $publishPath) { Copy-Item -LiteralPath $publishPath -Destination (Join-Path $backupDirectory ('publish-' + [DateTime]::UtcNow.Ticks + '.xml')) }
foreach ($name in $names) {
  $matches = @($xml.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.GetAttribute('name') -eq $name })
  foreach ($match in $matches) { [void]$xml.DocumentElement.RemoveChild($match) }
}
if ($Action -ne 'Remove') {
  for ($i=0;$i -lt 2;$i++) {
    $component = @('writer','sheet')[$i]
    $url = if ($Action -eq 'Test') { "http://127.0.0.1:5173/wps-native-test/$component/" } else { "http://127.0.0.1:5173/wps/$component/" }
    $entry = $xml.CreateElement('jspluginonline')
    $entry.SetAttribute('name',$names[$i]); $entry.SetAttribute('type',@('wps','et')[$i])
    $entry.SetAttribute('url',$url); $entry.SetAttribute('enable','enable_dev')
    $entry.SetAttribute('debug',''); $entry.SetAttribute('install','null')
    [void]$xml.DocumentElement.AppendChild($entry)
  }
}
New-Item -ItemType Directory -Path $addonDirectory -Force | Out-Null
$temporaryPath = $publishPath + '.office-helper.tmp'
$xml.Save($temporaryPath)
Move-Item -LiteralPath $temporaryPath -Destination $publishPath -Force
Write-Output "WPS registration: $Action. Only OfficeHelperWriter/OfficeHelperSheet entries changed."
