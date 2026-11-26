param([ValidateSet('Word','Excel')][string]$App = 'Word')
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$extension = if ($App -eq 'Word') { 'docx' } else { 'xlsx' }
$testPath = Join-Path $projectRoot "output\native\OfficeHelper-Native-Test.$extension"
if (-not (Test-Path -LiteralPath $testPath)) { throw 'Run prepare-native-check.mjs first.' }
try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject("$App.Application") } catch { $application = New-Object -ComObject "$App.Application" }
if ($application.Path -notlike '*Microsoft Office*') { throw 'The COM association points to another application. Open the test file with the Microsoft Office executable first.' }
$application.Visible = $true
if ($App -eq 'Word') {
  $doc = $null
  foreach ($candidate in $application.Documents) { if ($candidate.FullName -eq $testPath) { $doc = $candidate } }
  if (-not $doc) { $doc = $application.Documents.Open($testPath) }
  $doc.Content.Text = "OFFICE_HELPER_NATIVE_TEST`rOnly this generated test document is used."
  $doc.Range(0,25).Font.Size = 18
  $doc.Save()
  $doc.Range(0,25).Select()
  Write-Output 'Word test document prepared; native add-in verification will run inside its task pane.'
} else {
  $book = $null
  foreach ($candidate in $application.Workbooks) { if ($candidate.FullName -eq $testPath) { $book = $candidate } }
  if (-not $book) { $book = $application.Workbooks.Open($testPath) }
  $sheet = $book.Worksheets.Item(1)
  $sheet.Cells.Item(1,1).Value2 = 'OFFICE_HELPER_NATIVE_TEST'
  $values = @(120,85,45.5,200)
  for($i=0;$i -lt 4;$i++) { $sheet.Cells.Item($i+2,2).Value2 = [double]$values[$i] }
  $book.Save()
  $sheet.Range('B2:B5').Select()
  Write-Output 'Excel test document prepared; native add-in verification will run inside its task pane.'
}
# 保留专用文档供人工查看；从不读取或关闭其他已打开的文件。
