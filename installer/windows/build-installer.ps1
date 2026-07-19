[CmdletBinding()]
param(
  [string]$PayloadRoot = (Join-Path $PSScriptRoot '..\..\release\payload'),
  [string]$OutputDirectory = (Join-Path $PSScriptRoot 'dist')
)

$ErrorActionPreference = 'Stop'
$payload = [IO.Path]::GetFullPath($PayloadRoot)
$output = [IO.Path]::GetFullPath($OutputDirectory)
foreach ($required in @(
  'legacy\ProfiExcelHelper-Legacy.xlam',
  'template\ProfiExcelHelper-Template.xltm',
  'officejs\manifest.xml',
  'officejs\manifest-office2016.xml'
)) {
  if (-not (Test-Path -LiteralPath (Join-Path $payload $required))) { throw "В payload отсутствует $required" }
}

$candidates = @(
  "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
  "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
if ($candidates.Count -eq 0) { throw 'Не найден Inno Setup 6 (ISCC.exe).' }
$iscc = $candidates[0]
New-Item -ItemType Directory -Force -Path $output | Out-Null
Push-Location $PSScriptRoot
try {
  & $iscc "/DPayloadDir=$payload" (Join-Path $PSScriptRoot 'ProfiExcelHelper.iss')
  if ($LASTEXITCODE -ne 0) { throw "Inno Setup завершился с кодом $LASTEXITCODE" }
  $compiled = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'dist\ProfiExcelHelper-Setup-1.2.0.exe'))
  $target = [IO.Path]::GetFullPath((Join-Path $output 'ProfiExcelHelper-Setup-1.2.0.exe'))
  if (-not (Test-Path -LiteralPath $compiled)) { throw "Inno Setup не создал $compiled" }
  if (-not [string]::Equals($compiled, $target, [StringComparison]::OrdinalIgnoreCase)) {
    Copy-Item -LiteralPath $compiled -Destination $target -Force
  }
} finally {
  Pop-Location
}
Write-Host "Installer built in $output"
