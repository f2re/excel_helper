[CmdletBinding()]
param(
  [string]$XlamPath = (Join-Path $PSScriptRoot '..\dist\ProfiExcelHelper-Legacy.xlam'),
  [string]$XltmPath = (Join-Path $PSScriptRoot '..\dist\ProfiExcelHelper-Template.xltm')
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Profi.Build.Common.psm1') -Force
$excel = $null
$addin = $null
$templateWorkbook = $null
try {
  $xlam = [IO.Path]::GetFullPath($XlamPath)
  $xltm = [IO.Path]::GetFullPath($XltmPath)
  if (-not (Test-Path -LiteralPath $xlam)) { throw "XLAM не найден: $xlam" }
  if (-not (Test-Path -LiteralPath $xltm)) { throw "XLTM не найден: $xltm" }

  $excel = New-ProfiExcelApplication
  $addin = $excel.Workbooks.Open($xlam, 0, $true)
  if (-not $addin.IsAddin) { throw 'Файл XLAM открылся не как Excel Add-in.' }
  Assert-ProfiVbaProjectAccess -Workbook $addin
  if ($addin.VBProject.VBComponents.Item('modProfiSchedule') -eq $null) { throw 'В XLAM отсутствует modProfiSchedule.' }

  $templateWorkbook = $excel.Workbooks.Add($xltm)
  if ($templateWorkbook.IsAddin) { throw 'Рабочая книга из XLTM ошибочно помечена как Add-in.' }
  Assert-ProfiVbaProjectAccess -Workbook $templateWorkbook
  if ($templateWorkbook.Worksheets.Item('Старт') -eq $null) { throw 'В XLTM отсутствует стартовый лист.' }
  if ($templateWorkbook.VBProject.VBComponents.Item('modProfiTemplate') -eq $null) { throw 'В XLTM отсутствует modProfiTemplate.' }
  $version = $templateWorkbook.Names.Item('PROFI_TEMPLATE_VERSION').RefersToRange.Value2
  if ([string]::IsNullOrWhiteSpace([string]$version)) { throw 'В XLTM отсутствует версия шаблона.' }
  Write-Host "Verified XLAM and XLTM, template version: $version"
} finally {
  if ($null -ne $templateWorkbook) { try { $templateWorkbook.Close($false) } catch { }; Close-ProfiComObject $templateWorkbook }
  if ($null -ne $addin) { try { $addin.Close($false) } catch { }; Close-ProfiComObject $addin }
  if ($null -ne $excel) { try { $excel.Quit() } catch { }; Close-ProfiComObject $excel }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
