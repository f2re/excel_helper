[CmdletBinding()]
param(
  [ValidateSet('Full', 'AddinOnly', 'TemplateOnly')][string]$Mode = 'Full',
  [string]$PayloadRoot = (Join-Path $PSScriptRoot '..\..\release\payload'),
  [switch]$Silent,
  [switch]$Force,
  [switch]$SkipRegistration
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ProfiInstaller.Common.psm1') -Force
$paths = Get-ProfiInstallPaths
$office = Get-ProfiOfficeInfo
New-Item -ItemType Directory -Force -Path $paths.Root, $paths.Logs | Out-Null
$log = Join-Path $paths.Logs ("install-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Start-Transcript -Path $log -Force | Out-Null
try {
  Assert-ProfiExcelClosed -Force:$Force -Silent:$Silent
  $payload = [IO.Path]::GetFullPath($PayloadRoot)
  $addinSource = Join-Path $payload 'legacy\ProfiExcelHelper-Legacy.xlam'
  $templateSource = Join-Path $payload 'template\ProfiExcelHelper-Template.xltm'
  $addinTarget = Join-Path $paths.Addins 'ProfiExcelHelper-Legacy.xlam'
  $templateTarget = Join-Path $paths.Templates 'ProfiExcelHelper-Template.xltm'

  if ($Mode -in @('Full', 'AddinOnly')) {
    Copy-ProfiFile -Source $addinSource -Destination $addinTarget
    if (-not $SkipRegistration) {
      if (-not $office.Installed) { throw 'Microsoft Excel не найден; XLAM скопирован, но не может быть зарегистрирован.' }
      Register-ProfiExcelAddin -Path $addinTarget
    }
  }

  if ($Mode -in @('Full', 'TemplateOnly')) {
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
    addinPath = if ($Mode -in @('Full', 'AddinOnly')) { $addinTarget } else { $null }
    templatePath = if ($Mode -in @('Full', 'TemplateOnly')) { $templateTarget } else { $null }
    log = $log
  }
  $record | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $paths.Manifest -Encoding UTF8
  Write-Host "ПрофиПомощник установлен. Режим: $Mode"
  Write-Host "Excel: $($office.Version) $($office.Bitness)"
} finally {
  Stop-Transcript | Out-Null
}
