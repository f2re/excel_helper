Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-ProfiInstallPaths {
  [CmdletBinding()]
  param()
  $appData = [Environment]::GetFolderPath('ApplicationData')
  $localAppData = [Environment]::GetFolderPath('LocalApplicationData')
  $startMenu = [Environment]::GetFolderPath('Programs')
  [pscustomobject]@{
    Root = Join-Path $localAppData 'ProfiExcelHelper'
    Addins = Join-Path $appData 'Microsoft\AddIns'
    Templates = Join-Path $appData 'Microsoft\Templates\ProfiExcelHelper'
    StartMenu = Join-Path $startMenu 'ПрофиПомощник'
    Logs = Join-Path $localAppData 'ProfiExcelHelper\logs'
    Manifest = Join-Path $localAppData 'ProfiExcelHelper\install.json'
  }
}

function Get-ProfiOfficeInfo {
  [CmdletBinding()]
  param()
  $version = $null
  $bitness = $null
  $product = $null
  $excelPath = $null
  $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
  if ($null -ne $c2r) {
    $version = $c2r.VersionToReport
    $bitness = $c2r.Platform
    $product = $c2r.ProductReleaseIds
    $excelPath = Join-Path $c2r.InstallationPath 'root\Office16\EXCEL.EXE'
  }
  foreach ($officeVersion in @('16.0', '15.0', '14.0')) {
    if ($null -eq $excelPath -or -not (Test-Path -LiteralPath $excelPath)) {
      foreach ($base in @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office')) {
        $root = Get-ItemProperty "$base\$officeVersion\Excel\InstallRoot" -ErrorAction SilentlyContinue
        if ($null -ne $root -and $root.Path) {
          $candidate = Join-Path $root.Path 'EXCEL.EXE'
          if (Test-Path -LiteralPath $candidate) {
            $excelPath = $candidate
            if (-not $version) { $version = $officeVersion }
            if (-not $bitness) { $bitness = if ($base -like '*WOW6432Node*') { 'x86' } else { 'x64' } }
            break
          }
        }
      }
    }
  }
  if ($excelPath -and (Test-Path -LiteralPath $excelPath)) {
    try {
      $fileVersion = (Get-Item -LiteralPath $excelPath).VersionInfo.ProductVersion
      if ($fileVersion) { $version = $fileVersion }
    } catch { }
  }
  [pscustomobject]@{
    Installed = [bool]($excelPath -and (Test-Path -LiteralPath $excelPath))
    Version = $version
    Bitness = $bitness
    Product = $product
    ExcelPath = $excelPath
  }
}

function Test-ProfiExcelRunning {
  [CmdletBinding()]
  param()
  return [bool](Get-Process EXCEL -ErrorAction SilentlyContinue)
}

function Assert-ProfiExcelClosed {
  [CmdletBinding()]
  param([switch]$Force, [switch]$Silent)
  if (-not (Test-ProfiExcelRunning)) { return }
  if ($Force) { return }
  if ($Silent) { throw 'Excel запущен. Закройте все окна Excel и повторите установку.' }
  $answer = Read-Host 'Excel запущен. Закройте все окна Excel, затем введите Y для продолжения'
  if ($answer -notmatch '^[YyДд]$') { throw 'Установка отменена: Excel остаётся запущенным.' }
  if (Test-ProfiExcelRunning) { throw 'Excel всё ещё запущен.' }
}

function Register-ProfiExcelAddin {
  [CmdletBinding()]
  param([Parameter(Mandatory = $true)][string]$Path)
  $excel = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $existing = $null
    foreach ($addin in $excel.AddIns) {
      if ([string]::Equals([string]$addin.FullName, $Path, [StringComparison]::OrdinalIgnoreCase) -or [string]::Equals([string]$addin.Name, [IO.Path]::GetFileName($Path), [StringComparison]::OrdinalIgnoreCase)) {
        $existing = $addin
        break
      }
    }
    if ($null -eq $existing) { $existing = $excel.AddIns.Add($Path, $false) }
    $existing.Installed = $true
  } finally {
    if ($null -ne $excel) {
      try { $excel.Quit() } catch { }
      try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Unregister-ProfiExcelAddin {
  [CmdletBinding()]
  param([Parameter(Mandatory = $true)][string]$Path)
  $excel = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    foreach ($addin in $excel.AddIns) {
      if ([string]::Equals([string]$addin.FullName, $Path, [StringComparison]::OrdinalIgnoreCase) -or [string]::Equals([string]$addin.Name, [IO.Path]::GetFileName($Path), [StringComparison]::OrdinalIgnoreCase)) {
        $addin.Installed = $false
      }
    }
  } catch {
    Write-Warning "Не удалось снять регистрацию XLAM через Excel: $($_.Exception.Message)"
  } finally {
    if ($null -ne $excel) {
      try { $excel.Quit() } catch { }
      try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function New-ProfiShortcut {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)][string]$ShortcutPath,
    [Parameter(Mandatory = $true)][string]$TargetPath,
    [string]$Description = 'ПрофиПомощник для Excel'
  )
  New-Item -ItemType Directory -Force -Path (Split-Path $ShortcutPath -Parent) | Out-Null
  $shell = New-Object -ComObject WScript.Shell
  try {
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = Split-Path $TargetPath -Parent
    $shortcut.Description = $Description
    $shortcut.Save()
  } finally {
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell) } catch { }
  }
}

function Copy-ProfiFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination
  )
  if (-not (Test-Path -LiteralPath $Source)) { throw "Не найден файл поставки: $Source" }
  New-Item -ItemType Directory -Force -Path (Split-Path $Destination -Parent) | Out-Null
  Copy-Item -LiteralPath $Source -Destination $Destination -Force
  Unblock-File -LiteralPath $Destination -ErrorAction SilentlyContinue
}

Export-ModuleMember -Function Get-ProfiInstallPaths, Get-ProfiOfficeInfo, Test-ProfiExcelRunning, Assert-ProfiExcelClosed, Register-ProfiExcelAddin, Unregister-ProfiExcelAddin, New-ProfiShortcut, Copy-ProfiFile
