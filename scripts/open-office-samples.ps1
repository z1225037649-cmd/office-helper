param([ValidateSet('Word','Excel')][string]$App = 'Word')
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$extension = if ($App -eq 'Word') { 'docx' } else { 'xlsx' }
$executable = if ($App -eq 'Word') { 'WINWORD.EXE' } else { 'EXCEL.EXE' }
$samplePath = Join-Path $projectRoot "output\samples\OfficeHelper-Sample.$extension"
$officeExecutable = Join-Path $env:ProgramFiles "Microsoft Office\root\Office16\$executable"
if (-not (Test-Path -LiteralPath $samplePath)) { throw 'Run npm run setup:office first.' }
if (-not (Test-Path -LiteralPath $officeExecutable)) { throw 'Open the sample from Microsoft Office manually; this installation path was not found.' }
# File associations and Word.Application may point to WPS; use the Microsoft executable.
# /x limits Word's DDE startup and creates a separate Excel instance. Leave existing Office processes alone.
Start-Process -FilePath $officeExecutable -ArgumentList @('/x', ('"' + $samplePath + '"')) -WindowStyle Normal
Write-Output "Opened the sample with Microsoft $App."
