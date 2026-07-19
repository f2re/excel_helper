[CmdletBinding()]
param(
  [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path,
  [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\..\release'),
  [string]$Version = '1.2.0'
)

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath($RepositoryRoot)
$release = [IO.Path]::GetFullPath($OutputDirectory)
$payload = Join-Path $release 'payload'
Remove-Item -LiteralPath $payload -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path (Join-Path $payload 'legacy'), (Join-Path $payload 'template'), (Join-Path $payload 'officejs'), (Join-Path $payload 'docs') | Out-Null

Copy-Item -LiteralPath (Join-Path $root 'legacy-vba\dist\ProfiExcelHelper-Legacy.xlam') -Destination (Join-Path $payload 'legacy') -Force
Copy-Item -LiteralPath (Join-Path $root 'legacy-vba\dist\ProfiExcelHelper-Template.xltm') -Destination (Join-Path $payload 'template') -Force
Copy-Item -LiteralPath (Join-Path $root 'dist\manifest.xml') -Destination (Join-Path $payload 'officejs') -Force
Copy-Item -LiteralPath (Join-Path $root 'dist\manifest-office2016.xml') -Destination (Join-Path $payload 'officejs') -Force
Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination (Join-Path $payload 'docs\README.md') -Force
Copy-Item -LiteralPath (Join-Path $root 'docs\INSTALLATION.md') -Destination (Join-Path $payload 'docs') -Force
Copy-Item -LiteralPath (Join-Path $root 'docs\LEGACY_OFFICE.md') -Destination (Join-Path $payload 'docs') -Force
Copy-Item -LiteralPath (Join-Path $root 'docs\DISTRIBUTION.md') -Destination (Join-Path $payload 'docs') -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'install.ps1') -Destination $payload -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'uninstall.ps1') -Destination $payload -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'ProfiInstaller.Common.psm1') -Destination $payload -Force

$zip = Join-Path $release "ProfiExcelHelper-Portable-$Version.zip"
Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $payload '*') -DestinationPath $zip -CompressionLevel Optimal
Write-Host "Portable bundle built: $zip"
