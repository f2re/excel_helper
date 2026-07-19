[CmdletBinding()]
param(
  [ValidateSet('Full', 'LegacyFull', 'AddinOnly', 'TemplateOnly', 'ModernOnly', 'AddinModern', 'TemplateModern')][string]$Mode = 'Full',
  [string]$PayloadRoot = '',
  [switch]$Silent,
  [switch]$Force,
  [switch]$SkipRegistration
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ProfiInstaller.Common.psm1') -Force
$paths = Get-ProfiInstallPaths
$office = Get-ProfiOfficeInfo
$installAddin = $Mode -in @('Full', 'LegacyFull', 'AddinOnly', 'AddinModern')
$installTemplate = $Mode -in @('Full', 'LegacyFull', 'TemplateOnly', 'TemplateModern')
$installModern = $Mode -in @('Full', 'ModernOnly', 'AddinModern', 'TemplateModern')

if ([string]::IsNullOrWhiteSpace($PayloadRoot)) {
  $candidates = @(
    $PSScriptRoot,
    (Join-Path $PSScriptRoot '..\payload'),
    (Join-Path $PSScriptRoot '..\..\release\payload')
  )
  $PayloadRoot = $candidates | Where-Object {
    (Test-Path -LiteralPath $_) -and (
      (Test-Path -LiteralPath (Join-Path $_ 'legacy')) -or
      (Test-Path -LiteralPath (Join-Path $_ 'template')) -or
      (Test-Path -LiteralPath (Join-Path $_ 'officejs'))
    )
  } | Select-Object -First 1
  if (-not $PayloadRoot) { throw 'Не найден payload. Передайте -PayloadRoot явно.' }
}

$payload = [IO.Path]::GetFullPath($PayloadRoot)
$addinSource = Join-Path $payload 'legacy\ProfiExcelHelper-Legacy.xlam'
$templateSource = Join-Path $payload 'template\ProfiExcelHelper-Template.xltm'
$modernSource = Join-Path $payload 'officejs'
$addinTarget = Join-Path $paths.Addins 'ProfiExcelHelper-Legacy.xlam'
$templateTarget = Join-Path $paths.Templates 'ProfiExcelHelper-Template.xltm'
$modernTarget = Join-Path $paths.Root 'officejs'

if ($installAddin) {
  if (-not (Test-Path -LiteralPath $addinSource)) { throw "Не найден XLAM: $addinSource" }
  if (-not $SkipRegistration -and -not $office.Installed) { throw 'Microsoft Excel не найден. Для копирования XLAM без регистрации используйте -SkipRegistration.' }
}
if ($installTemplate -and -not (Test-Path -LiteralPath $templateSource)) { throw "Не найден XLTM: $templateSource" }
if ($installModern) {
  foreach ($manifestName in @('manifest.xml', 'manifest-office2016.xml')) {
    if (-not (Test-Path -LiteralPath (Join-Path $modernSource $manifestName))) { throw "В Office.js payload отсутствует $manifestName" }
  }
}

New-Item -ItemType Directory -Force -Path $paths.Root, $paths.Logs | Out-Null
$log = Join-Path $paths.Logs ("install-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$backup = Join-Path $paths.Root ("backup-{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$addinBackedUp = $false
$templateBackedUp = $false
$modernBackedUp = $false
$addinRegistered = $false
Start-Transcript -Path $log -Force | Out-Null
try {
  if ($installAddin) { Assert-ProfiExcelClosed -Force:$Force -Silent:$Silent }
  New-Item -ItemType Directory -Force -Path $backup | Out-Null

  if ($installAddin) {
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

  if ($installTemplate) {
    if (Test-Path -LiteralPath $templateTarget) {
      Copy-Item -LiteralPath $templateTarget -Destination (Join-Path $backup 'ProfiExcelHelper-Template.xltm') -Force
      $templateBackedUp = $true
    }
    Copy-ProfiFile -Source $templateSource -Destination $templateTarget
    New-ProfiShortcut -ShortcutPath (Join-Path $paths.StartMenu 'ПрофиПомощник — новый проект.lnk') -TargetPath $templateTarget -Description 'Создать новую книгу ПрофиПомощник из XLTM-шаблона'
  }

  if ($installModern) {
    if (Test-Path -LiteralPath $modernTarget) {
      Copy-Item -LiteralPath $modernTarget -Destination (Join-Path $backup 'officejs') -Recurse -Force
      $modernBackedUp = $true
      Remove-Item -LiteralPath $modernTarget -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $modernTarget | Out-Null
    Copy-Item -Path (Join-Path $modernSource '*') -Destination $modernTarget -Recurse -Force
  }

  $record = [ordered]@{
    version = '1.2.0'
    installedAt = (Get-Date).ToString('o')
    mode = $Mode
    officeInstalled = $office.Installed
    officeVersion = $office.Version
    officeBitness = $office.Bitness
    officeProduct = $office.Product
    addinPath = $(if ($installAddin) { $addinTarget } else { $null })
    templatePath = $(if ($installTemplate) { $templateTarget } else { $null })
    officeJsPath = $(if ($installModern) { $modernTarget } else { $null })
    backup = $backup
    log = $log
  }
  $record | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $paths.Manifest -Encoding UTF8
  Write-Host "ПрофиПомощник установлен. Режим: $Mode"
  if ($office.Installed) { Write-Host "Excel: $($office.Version) $($office.Bitness)" }
} catch {
  $failure = $_
  Write-Warning "Установка не завершена: $($failure.Exception.Message). Выполняется откат."
  if ($addinRegistered -and (Test-Path -LiteralPath $addinTarget)) { Unregister-ProfiExcelAddin -Path $addinTarget }
  if ($installAddin) {
    if ($addinBackedUp) {
      Copy-Item -LiteralPath (Join-Path $backup 'ProfiExcelHelper-Legacy.xlam') -Destination $addinTarget -Force
      if (-not $SkipRegistration -and $office.Installed) { Register-ProfiExcelAddin -Path $addinTarget }
    } else {
      Remove-Item -LiteralPath $addinTarget -Force -ErrorAction SilentlyContinue
    }
  }
  if ($installTemplate) {
    if ($templateBackedUp) { Copy-Item -LiteralPath (Join-Path $backup 'ProfiExcelHelper-Template.xltm') -Destination $templateTarget -Force }
    else { Remove-Item -LiteralPath $templateTarget -Force -ErrorAction SilentlyContinue }
  }
  if ($installModern) {
    Remove-Item -LiteralPath $modernTarget -Recurse -Force -ErrorAction SilentlyContinue
    if ($modernBackedUp) { Copy-Item -LiteralPath (Join-Path $backup 'officejs') -Destination $modernTarget -Recurse -Force }
  }
  throw $failure
} finally {
  Stop-Transcript | Out-Null
}
