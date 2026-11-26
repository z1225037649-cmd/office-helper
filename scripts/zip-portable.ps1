param([Parameter(Mandatory=$true)][string]$Source,[Parameter(Mandatory=$true)][string]$Destination)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\output\portable')) + [IO.Path]::DirectorySeparatorChar
$sourcePath = [IO.Path]::GetFullPath($Source)
$destinationPath = [IO.Path]::GetFullPath($Destination)
if (-not $sourcePath.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or -not $destinationPath.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)) { throw 'Package paths must stay inside output/portable.' }
if (Test-Path -LiteralPath $destinationPath) { Remove-Item -LiteralPath $destinationPath }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::CreateFromDirectory($sourcePath,$destinationPath,[IO.Compression.CompressionLevel]::Optimal,$true)
