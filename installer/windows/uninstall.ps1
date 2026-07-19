[CmdletBinding()]
param([switch]$Silent, [switch]$Force)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ProfiInstaller.Common.psm1') -Force
$paths = Get-ProfiInstallPaths
New-Item -ItemType Directory -Force -Path $paths.Logs | Out-Null
$log = Join-Path $paths.Logs ("uninstall-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Start-Transcript -Path $log -Force | Out-Null
try {
  $addin = Join-Path $paths.Addins 'ProfiExcelHelper-Legacy.xlam'
  if (Test-Path -LiteralPath $addin) {
    Assert-ProfiExcelClosed -Force:$Force -Silent:$Silent
    Unregister-ProfiExcelAddin -Path $addin
  }
  Remove-Item -LiteralPath $addin -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath (Join-Path $paths.Templates 'ProfiExcelHelper-Template.xltm') -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $paths.StartMenu -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $paths.Manifest -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath (Join-Path $paths.Root 'officejs') -Recurse -Force -ErrorAction SilentlyContinue
  Get-ChildItem -LiteralPath $paths.Root -Directory -Filter 'backup-*' -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
  if ((Test-Path -LiteralPath $paths.Templates) -and -not (Get-ChildItem -LiteralPath $paths.Templates -Force -ErrorAction SilentlyContinue)) {
    Remove-Item -LiteralPath $paths.Templates -Force -ErrorAction SilentlyContinue
  }
  Write-Host 'ПрофиПомощник удалён из профиля текущего пользователя.'
  Write-Host "Журнал удаления сохранён: $log"
} finally {
  Stop-Transcript | Out-Null
}
