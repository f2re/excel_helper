[CmdletBinding()]
param([switch]$RequireVbaProjectAccess)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Legacy-сборка доступна только в Windows.' }

$result = [ordered]@{
  Windows = $true
  PowerShell = $PSVersionTable.PSVersion.ToString()
  ExcelAvailable = $false
  ExcelVersion = ''
  ExcelProcessBitness = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
  VbaProjectAccess = $false
  InnoSetup = ''
  SignTool = ''
}
$excel = $null
$workbook = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $result.ExcelAvailable = $true
  $result.ExcelVersion = [string]$excel.Version
  $workbook = $excel.Workbooks.Add()
  try {
    $null = $workbook.VBProject.VBComponents.Count
    $result.VbaProjectAccess = $true
  } catch {
    $result.VbaProjectAccess = $false
  }
} finally {
  if ($null -ne $workbook) {
    try { $workbook.Close($false) } catch { }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch { }
  }
  if ($null -ne $excel) {
    try { $excel.Quit() } catch { }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$innoCandidates = @(
  (Get-Command ISCC.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue),
  (Join-Path "${env:ProgramFiles(x86)}" 'Inno Setup 6\ISCC.exe'),
  (Join-Path "$env:ProgramFiles" 'Inno Setup 6\ISCC.exe')
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
$result.InnoSetup = [string]($innoCandidates | Select-Object -First 1)
$result.SignTool = [string](Get-Command signtool.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue)
$result | Format-List
if (-not $result.ExcelAvailable) { throw 'Microsoft Excel не обнаружен.' }
if ($RequireVbaProjectAccess -and -not $result.VbaProjectAccess) {
  throw 'Отключён доверенный доступ к объектной модели VBA. Включите его в Центре управления безопасностью Excel только на время сборки.'
}
$result
