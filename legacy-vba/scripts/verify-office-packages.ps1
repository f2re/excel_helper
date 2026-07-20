[CmdletBinding()]
param(
  [string]$XlamPath = (Join-Path $PSScriptRoot '..\dist\ProfiExcelHelper-Legacy.xlam'),
  [string]$XltmPath = (Join-Path $PSScriptRoot '..\dist\ProfiExcelHelper-Template.xltm')
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Profi.Build.Common.psm1') -Force
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Assert-ProfiRibbonPackage {
  param([Parameter(Mandatory = $true)][string]$Path)
  $archive = [IO.Compression.ZipFile]::OpenRead($Path)
  try {
    $ribbon = $archive.GetEntry('customUI/customUI.xml')
    if ($null -eq $ribbon) { throw "В пакете отсутствует customUI/customUI.xml: $Path" }
    $relationships = $archive.GetEntry('_rels/.rels')
    if ($null -eq $relationships) { throw "В пакете отсутствует _rels/.rels: $Path" }
    $stream = $ribbon.Open()
    $reader = New-Object IO.StreamReader($stream)
    try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose(); $stream.Dispose() }
    if ($xml -notmatch 'ProfiRibbonOnLoad' -or $xml -notmatch 'ProfiRibbonComposeSchedule') {
      throw "RibbonX не содержит обязательные callback-процедуры: $Path"
    }
  } finally {
    $archive.Dispose()
  }
}

$excel = $null
$addin = $null
$templateWorkbook = $null
try {
  $xlam = [IO.Path]::GetFullPath($XlamPath)
  $xltm = [IO.Path]::GetFullPath($XltmPath)
  if (-not (Test-Path -LiteralPath $xlam)) { throw "XLAM не найден: $xlam" }
  if (-not (Test-Path -LiteralPath $xltm)) { throw "XLTM не найден: $xltm" }
  Assert-ProfiRibbonPackage -Path $xlam
  Assert-ProfiRibbonPackage -Path $xltm

  $excel = New-ProfiExcelApplication
  $addin = $excel.Workbooks.Open($xlam, 0, $true)
  if (-not $addin.IsAddin) { throw 'Файл XLAM открылся не как Excel Add-in.' }
  Assert-ProfiVbaProjectAccess -Workbook $addin
  foreach ($moduleName in @('modProfiSchedule', 'modProfiRibbon', 'modProfiFunctions')) {
    if ($addin.VBProject.VBComponents.Item($moduleName) -eq $null) { throw "В XLAM отсутствует $moduleName." }
  }

  $templateWorkbook = $excel.Workbooks.Add($xltm)
  if ($templateWorkbook.IsAddin) { throw 'Рабочая книга из XLTM ошибочно помечена как Add-in.' }
  Assert-ProfiVbaProjectAccess -Workbook $templateWorkbook
  if ($templateWorkbook.Worksheets.Item('Старт') -eq $null) { throw 'В XLTM отсутствует стартовый лист.' }
  foreach ($moduleName in @('modProfiTemplate', 'modProfiSchedule', 'modProfiRibbon')) {
    if ($templateWorkbook.VBProject.VBComponents.Item($moduleName) -eq $null) { throw "В XLTM отсутствует $moduleName." }
  }
  $version = $templateWorkbook.Names.Item('PROFI_TEMPLATE_VERSION').RefersToRange.Value2
  if ([string]::IsNullOrWhiteSpace([string]$version)) { throw 'В XLTM отсутствует версия шаблона.' }
  if ([string]$version -ne '1.2.0') { throw "Версия XLTM не совпадает: $version" }
  Write-Host "Verified XLAM, XLTM, RibbonX and VBA modules; template version: $version"
} finally {
  if ($null -ne $templateWorkbook) { try { $templateWorkbook.Close($false) } catch { }; Close-ProfiComObject $templateWorkbook }
  if ($null -ne $addin) { try { $addin.Close($false) } catch { }; Close-ProfiComObject $addin }
  if ($null -ne $excel) { try { $excel.Quit() } catch { }; Close-ProfiComObject $excel }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
