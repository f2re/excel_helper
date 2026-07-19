[CmdletBinding()]
param(
  [string]$PayloadRoot = '',
  [ValidateSet('', 'Full', 'LegacyFull', 'AddinOnly', 'TemplateOnly', 'ModernOnly', 'AddinModern', 'TemplateModern')][string]$Mode = '',
  [switch]$Silent
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ProfiInstaller.Common.psm1') -Force
$paths = Get-ProfiInstallPaths
if ([string]::IsNullOrWhiteSpace($Mode) -and (Test-Path -LiteralPath $paths.Manifest)) {
  try { $Mode = [string](Get-Content -LiteralPath $paths.Manifest -Raw -Encoding UTF8 | ConvertFrom-Json).mode } catch { }
}
if ([string]::IsNullOrWhiteSpace($Mode)) {
  $root = $PayloadRoot
  if ([string]::IsNullOrWhiteSpace($root)) {
    $root = @($PSScriptRoot, (Join-Path $PSScriptRoot '..\payload')) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
  }
  $hasAddin = Test-Path -LiteralPath (Join-Path $root 'legacy\ProfiExcelHelper-Legacy.xlam')
  $hasTemplate = Test-Path -LiteralPath (Join-Path $root 'template\ProfiExcelHelper-Template.xltm')
  $hasModern = Test-Path -LiteralPath (Join-Path $root 'officejs\manifest.xml')
  if ($hasAddin -and $hasTemplate -and $hasModern) { $Mode = 'Full' }
  elseif ($hasAddin -and $hasTemplate) { $Mode = 'LegacyFull' }
  elseif ($hasAddin -and $hasModern) { $Mode = 'AddinModern' }
  elseif ($hasTemplate -and $hasModern) { $Mode = 'TemplateModern' }
  elseif ($hasAddin) { $Mode = 'AddinOnly' }
  elseif ($hasTemplate) { $Mode = 'TemplateOnly' }
  elseif ($hasModern) { $Mode = 'ModernOnly' }
  else { throw 'Не удалось определить компоненты для восстановления.' }
}

$arguments = @{ Mode = $Mode; Silent = $Silent }
if (-not [string]::IsNullOrWhiteSpace($PayloadRoot)) { $arguments.PayloadRoot = $PayloadRoot }
& (Join-Path $PSScriptRoot 'install.ps1') @arguments
