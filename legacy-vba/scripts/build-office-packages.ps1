[CmdletBinding()]
param(
  [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\dist'),
  [string]$Version = '1.2.0',
  [switch]$SkipVerification
)

$ErrorActionPreference = 'Stop'
$dist = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $dist | Out-Null
$xlam = Join-Path $dist 'ProfiExcelHelper-Legacy.xlam'
$xltm = Join-Path $dist 'ProfiExcelHelper-Template.xltm'

& (Join-Path $PSScriptRoot 'build-xlam.ps1') -OutputPath $xlam -Version $Version
& (Join-Path $PSScriptRoot 'build-xltm.ps1') -OutputPath $xltm -Version $Version
if (-not $SkipVerification) {
  & (Join-Path $PSScriptRoot 'verify-office-packages.ps1') -XlamPath $xlam -XltmPath $xltm
}
Write-Host "Legacy Office packages are ready in $dist"
