param([ValidateSet('Word','Excel')][string]$App = 'Word')
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$extension = if ($App -eq 'Word') { 'docx' } else { 'xlsx' }
$directory = Join-Path $projectRoot 'output\native\wps'
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$testPath = Join-Path $directory "OfficeHelper-Native-Test.$extension"
if (-not (Test-Path -LiteralPath $testPath)) { Copy-Item -LiteralPath (Join-Path $projectRoot "output\native\OfficeHelper-Native-Test.$extension") -Destination $testPath }
$progId = if ($App -eq 'Word') { 'kwps.Application' } else { 'ket.Application' }
function Get-ComProperty($Object,[string]$Name) { return $Object.GetType().InvokeMember($Name,[Reflection.BindingFlags]::GetProperty,$null,$Object,$null) }
function Set-ComProperty($Object,[string]$Name,$Value) { [void]$Object.GetType().InvokeMember($Name,[Reflection.BindingFlags]::SetProperty,$null,$Object,@($Value)) }
function Invoke-ComMethod($Object,[string]$Name,[object[]]$Arguments) { return $Object.GetType().InvokeMember($Name,[Reflection.BindingFlags]::InvokeMethod,$null,$Object,$Arguments) }
try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject($progId) } catch { $application = New-Object -ComObject $progId }
$application.Visible = $true
if ($App -eq 'Word') {
  $doc = $application.Documents.Open($testPath)
  $application.Selection.SetRange(0,25)
  $application.Selection.Font.Size = 18
} else {
  $book = $application.Workbooks.Open($testPath)
  $sheet = $book.Worksheets.Item(1)
  $sheet.Cells.Item(1,1).Value2 = 'OFFICE_HELPER_NATIVE_TEST'
  $values = @(120,85,45.5,200)
  for($i=0;$i -lt 4;$i++) { $sheet.Cells.Item($i+2,2).Value2 = [double]$values[$i] }
  $sheet.Range('B2:B5').Select()
}
Write-Output "WPS $App test document prepared."
