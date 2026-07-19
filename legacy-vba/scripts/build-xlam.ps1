[CmdletBinding()]
param(
  [string]$OutputPath = (Join-Path $PSScriptRoot '..\dist\ProfiExcelHelper-Legacy.xlam'),
  [string]$Version = '1.2.0'
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Profi.Build.Common.psm1') -Force
. (Join-Path $PSScriptRoot 'Inject-Ribbon.ps1')
$excel = $null
$workbook = $null
try {
  $fullOutput = [IO.Path]::GetFullPath($OutputPath)
  New-Item -ItemType Directory -Force -Path (Split-Path $fullOutput -Parent) | Out-Null
  if (Test-Path -LiteralPath $fullOutput) { Remove-Item -LiteralPath $fullOutput -Force }

  $excel = New-ProfiExcelApplication
  $workbook = $excel.Workbooks.Add(-4167)
  [void](Import-ProfiVbaModules -Workbook $workbook -Exclude @('modProfiTemplate.bas'))
  Set-ProfiDocumentProperties -Workbook $workbook -Title 'ПрофиПомощник Legacy XLAM' -Version $Version
  $workbook.IsAddin = $true
  $workbook.SaveAs($fullOutput, 55)
  $workbook.Close($false)
  Close-ProfiComObject $workbook
  $workbook = $null
  $excel.Quit()
  Close-ProfiComObject $excel
  $excel = $null

  Add-ProfiRibbon -WorkbookPath $fullOutput -RibbonXmlPath (Join-Path $PSScriptRoot '..\ribbon\customUI.xml')
  if (-not (Test-Path -LiteralPath $fullOutput)) { throw "Excel не создал файл: $fullOutput" }
  if ((Get-Item -LiteralPath $fullOutput).Length -lt 4096) { throw 'Созданный XLAM имеет подозрительно малый размер.' }
  Write-Host "Built XLAM: $fullOutput"
} finally {
  if ($null -ne $workbook) { try { $workbook.Close($false) } catch { }; Close-ProfiComObject $workbook }
  if ($null -ne $excel) { try { $excel.Quit() } catch { }; Close-ProfiComObject $excel }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
