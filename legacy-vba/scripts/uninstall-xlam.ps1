[CmdletBinding()]
param([switch]$Force)

$ErrorActionPreference = 'Stop'
$installerModule = Join-Path $PSScriptRoot '..\..\installer\windows\ProfiInstaller.Common.psm1'
Import-Module $installerModule -Force
Assert-ProfiExcelClosed -Force:$Force
$paths = Get-ProfiInstallPaths
$target = Join-Path $paths.Addins 'ProfiExcelHelper-Legacy.xlam'
if (Test-Path -LiteralPath $target) { Unregister-ProfiExcelAddin -Path $target }
Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
Write-Host "Removed XLAM: $target"
