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
$required = @(
  'legacy-vba\dist\ProfiExcelHelper-Legacy.xlam',
  'legacy-vba\dist\ProfiExcelHelper-Template.xltm',
  'dist\manifest.xml',
  'dist\manifest-office2016.xml'
)
foreach ($file in $required) {
  if (-not (Test-Path -LiteralPath (Join-Path $root $file))) { throw "Перед упаковкой отсутствует $file" }
}

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
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'repair.ps1') -Destination $payload -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'ProfiInstaller.Common.psm1') -Destination $payload -Force

@'
@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -PayloadRoot "%~dp0" -Mode Full
if errorlevel 1 pause
'@ | Set-Content -LiteralPath (Join-Path $payload 'Установить.cmd') -Encoding ASCII
@'
@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
if errorlevel 1 pause
'@ | Set-Content -LiteralPath (Join-Path $payload 'Удалить.cmd') -Encoding ASCII
@'
@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0repair.ps1" -PayloadRoot "%~dp0"
if errorlevel 1 pause
'@ | Set-Content -LiteralPath (Join-Path $payload 'Восстановить.cmd') -Encoding ASCII

$zip = Join-Path $release "ProfiExcelHelper-Portable-$Version.zip"
Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $payload '*') -DestinationPath $zip -CompressionLevel Optimal
if ((Get-Item -LiteralPath $zip).Length -lt 4096) { throw 'Переносимый ZIP имеет подозрительно малый размер.' }
Write-Host "Portable bundle built: $zip"
