[CmdletBinding()]
param(
  [string]$XlamPath = (Join-Path $PSScriptRoot '..\dist\ProfiExcelHelper-Legacy.xlam'),
  [switch]$Force
)

$ErrorActionPreference = 'Stop'
$installerModule = Join-Path $PSScriptRoot '..\..\installer\windows\ProfiInstaller.Common.psm1'
Import-Module $installerModule -Force
if (-not (Test-Path -LiteralPath $XlamPath)) {
  & (Join-Path $PSScriptRoot 'build-xlam.ps1') -OutputPath $XlamPath
}
Assert-ProfiExcelClosed -Force:$Force
$paths = Get-ProfiInstallPaths
$target = Join-Path $paths.Addins 'ProfiExcelHelper-Legacy.xlam'
Copy-ProfiFile -Source ([IO.Path]::GetFullPath($XlamPath)) -Destination $target
Register-ProfiExcelAddin -Path $target
Write-Host "Installed XLAM: $target"
