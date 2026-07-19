[CmdletBinding()]
param([switch]$RequireVbaProjectAccess)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Legacy-сборка доступна только в Windows.' }
$processBitness = 'x86'
if ([Environment]::Is64BitProcess) { $processBitness = 'x64' }

$result = [ordered]@{
  Windows = $true
  PowerShell = $PSVersionTable.PSVersion.ToString()
  ExcelAvailable = $false
  ExcelVersion = ''
  ExcelProcessBitness = $processBitness
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
  $workbook = $excel.Workbooks.Add(-4167)
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

$innoCandidates = @()
$innoCommand = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if ($null -ne $innoCommand) { $innoCandidates += $innoCommand.Source }
$innoCandidates += (Join-Path "${env:ProgramFiles(x86)}" 'Inno Setup 6\ISCC.exe')
$innoCandidates += (Join-Path "$env:ProgramFiles" 'Inno Setup 6\ISCC.exe')
$result.InnoSetup = [string]($innoCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1)
$signCommand = Get-Command signtool.exe -ErrorAction SilentlyContinue
if ($null -ne $signCommand) { $result.SignTool = [string]$signCommand.Source }
$result | Format-List
if (-not $result.ExcelAvailable) { throw 'Microsoft Excel не обнаружен.' }
if ($RequireVbaProjectAccess -and -not $result.VbaProjectAccess) {
  throw 'Отключён доверенный доступ к объектной модели VBA. Включите его в Центре управления безопасностью Excel только на время сборки.'
}
$result
