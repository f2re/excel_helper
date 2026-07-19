[CmdletBinding()]
param(
  [ValidateSet('Full', 'AddinOnly', 'TemplateOnly')][string]$Mode = 'Full',
  [string]$PayloadRoot = '',
  [switch]$Silent,
  [switch]$Force,
  [switch]$SkipRegistration
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ProfiInstaller.Common.psm1') -Force
$paths = Get-ProfiInstallPaths
$office = Get-ProfiOfficeInfo
if ([string]::IsNullOrWhiteSpace($PayloadRoot)) {
  $candidates = @(
    $PSScriptRoot,
    (Join-Path $PSScriptRoot '..\payload'),
    (Join-Path $PSScriptRoot '..\..\release\payload')
  )
  $PayloadRoot = $candidates | Where-Object { Test-Path -LiteralPath (Join-Path $_ 'legacy') } | Select-Object -First 1
  if (-not $PayloadRoot) { throw 'Не найден payload. Передайте -PayloadRoot явно.' }
}
$payload = [IO.Path]::GetFullPath($PayloadRoot)
$addinSource = Join-Path $payload 'legacy\ProfiExcelHelper-Legacy.xlam'
$templateSource = Join-Path $payload 'template\ProfiExcelHelper-Template.xltm'
$addinTarget = Join-Path $paths.Addins 'ProfiExcelHelper-Legacy.xlam'
$templateTarget = Join-Path $paths.Templates 'ProfiExcelHelper-Template.xltm'

if ($Mode -in @('Full', 'AddinOnly')) {
  if (-not (Test-Path -LiteralPath $addinSource)) { throw "Не найден XLAM: $addinSource" }
  if (-not $SkipRegistration -and -not $office.Installed) { throw 'Microsoft Excel не найден. Для копирования без регистрации используйте -SkipRegistration.' }
}
if ($Mode -in @('Full', 'TemplateOnly')) {
  if (-not (Test-Path -LiteralPath $templateSource)) { throw "Не найден XLTM: $templateSource" }
}

New-Item -ItemType Directory -Force -Path $paths.Root, $paths.Logs | Out-Null
$log = Join-Path $paths.Logs ("install-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$backup = Join-Path $paths.Root ("backup-{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$addinBackedUp = $false
$templateBackedUp = $false
$addinRegistered = $false
Start-Transcript -Path $log -Force | Out-Null
try {
  Assert-ProfiExcelClosed -Force:$Force -Silent:$Silent
  New-Item -ItemType Directory -Force -Path $backup | Out-Null

  if ($Mode -in @('Full', 'AddinOnly')) {
    if (Test-Path -LiteralPath $addinTarget) {
      Copy-Item -LiteralPath $addinTarget -Destination (Join-Path $backup 'ProfiExcelHelper-Legacy.xlam') -Force
      $addinBackedUp = $true
    }
    Copy-ProfiFile -Source $addinSource -Destination $addinTarget
    if (-not $SkipRegistration) {
      Register-ProfiExcelAddin -Path $addinTarget
      $addinRegistered = $true
    }
  }

  if ($Mode -in @('Full', 'TemplateOnly')) {
    if (Test-Path -LiteralPath $templateTarget) {
      Copy-Item -LiteralPath $templateTarget -Destination (Join-Path $backup 'ProfiExcelHelper-Template.xltm') -Force
      $templateBackedUp = $true
    }
    Copy-ProfiFile -Source $templateSource -Destination $templateTarget
    New-ProfiShortcut -ShortcutPath (Join-Path $paths.StartMenu 'ПрофиПомощник — новый проект.lnk') -TargetPath $templateTarget -Description 'Создать новую книгу ПрофиПомощник из XLTM-шаблона'
  }

  $manifestSource = Join-Path $payload 'officejs'
  if (Test-Path -LiteralPath $manifestSource) {
    $manifestTarget = Join-Path $paths.Root 'officejs'
    New-Item -ItemType Directory -Force -Path $manifestTarget | Out-Null
    Copy-Item -Path (Join-Path $manifestSource '*') -Destination $manifestTarget -Recurse -Force
  }

  $record = [ordered]@{
    version = '1.2.0'
    installedAt = (Get-Date).ToString('o')
    mode = $Mode
    officeInstalled = $office.Installed
    officeVersion = $office.Version
    officeBitness = $office.Bitness
    officeProduct = $office.Product
    addinPath = if ($Mode -in @('Full', 'AddinOnly')) { $addinTarget } else { $null }
    templatePath = if ($Mode -in @('Full', 'TemplateOnly')) { $templateTarget } else { $null }
    backup = $backup
    log = $log
  }
  $record | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $paths.Manifest -Encoding UTF8
  Write-Host "ПрофиПомощник установлен. Режим: $Mode"
  Write-Host "Excel: $($office.Version) $($office.Bitness)"
} catch {
  Write-Error $_
  if ($addinRegistered -and (Test-Path -LiteralPath $addinTarget)) { Unregister-ProfiExcelAddin -Path $addinTarget }
  if ($Mode -in @('Full', 'AddinOnly')) {
    if ($addinBackedUp) { Copy-Item -LiteralPath (Join-Path $backup 'ProfiExcelHelper-Legacy.xlam') -Destination $addinTarget -Force }
    else { Remove-Item -LiteralPath $addinTarget -Force -ErrorAction SilentlyContinue }
  }
  if ($Mode -in @('Full', 'TemplateOnly')) {
    if ($templateBackedUp) { Copy-Item -LiteralPath (Join-Path $backup 'ProfiExcelHelper-Template.xltm') -Destination $templateTarget -Force }
    else { Remove-Item -LiteralPath $templateTarget -Force -ErrorAction SilentlyContinue }
  }
  throw
} finally {
  Stop-Transcript | Out-Null
}
