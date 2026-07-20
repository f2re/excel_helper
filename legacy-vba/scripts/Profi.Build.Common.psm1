Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-ProfiLegacyRoot {
  [CmdletBinding()]
  param()
  return (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}

function Get-ProfiSourceModules {
  [CmdletBinding()]
  param([string[]]$Exclude = @())
  $root = Get-ProfiLegacyRoot
  $source = Join-Path $root 'src'
  $excluded = @{}
  foreach ($name in $Exclude) { $excluded[$name.ToLowerInvariant()] = $true }
  return @(Get-ChildItem -LiteralPath $source -Filter '*.bas' -File | Sort-Object Name | Where-Object {
    -not $excluded.ContainsKey($_.Name.ToLowerInvariant())
  })
}

function New-ProfiExcelApplication {
  [CmdletBinding()]
  param()
  if ($env:OS -ne 'Windows_NT') {
    throw 'Сборка XLAM/XLTM выполняется только в Windows с установленным Microsoft Excel.'
  }
  try {
    $excel = New-Object -ComObject Excel.Application
  } catch {
    throw "Microsoft Excel COM не найден. Установите настольный Excel и повторите сборку. $($_.Exception.Message)"
  }
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.EnableEvents = $false
  return $excel
}

function Assert-ProfiVbaProjectAccess {
  [CmdletBinding()]
  param([Parameter(Mandatory = $true)]$Workbook)
  try {
    $count = $Workbook.VBProject.VBComponents.Count
    if ($count -lt 1) { throw 'VBProject пуст.' }
  } catch {
    throw @'
Excel запретил программный доступ к проекту VBA.
Включите: Файл → Параметры → Центр управления безопасностью → Параметры центра → Параметры макросов → «Доверять доступ к объектной модели проектов VBA».
После сборки этот параметр можно выключить.
'@
  }
}

function Import-ProfiVbaModules {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]$Workbook,
    [string[]]$Exclude = @()
  )
  Assert-ProfiVbaProjectAccess -Workbook $Workbook
  $files = Get-ProfiSourceModules -Exclude $Exclude
  if ($files.Count -eq 0) { throw 'Не найдены VBA-модули legacy-vba/src/*.bas.' }
  foreach ($file in $files) {
    $text = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    $match = [regex]::Match($text, 'Attribute\s+VB_Name\s*=\s*"([^"]+)"', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { throw "В $($file.Name) отсутствует Attribute VB_Name." }
    $moduleName = $match.Groups[1].Value
    $code = ($text -split "`r?`n" | Where-Object { $_ -notmatch '^Attribute\s+VB_' }) -join "`r`n"
    $existing = $null
    try { $existing = $Workbook.VBProject.VBComponents.Item($moduleName) } catch { }
    if ($null -ne $existing) { $Workbook.VBProject.VBComponents.Remove($existing) }
    $component = $Workbook.VBProject.VBComponents.Add(1)
    $component.Name = $moduleName
    $component.CodeModule.AddFromString($code)
  }
  return $files
}

function Set-ProfiThisWorkbookCode {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]$Workbook,
    [Parameter(Mandatory = $true)][string]$Code
  )
  Assert-ProfiVbaProjectAccess -Workbook $Workbook
  $component = $null
  try { $component = $Workbook.VBProject.VBComponents.Item([string]$Workbook.CodeName) } catch { }
  if ($null -eq $component) {
    foreach ($candidate in $Workbook.VBProject.VBComponents) {
      if ($candidate.Type -eq 100 -and $candidate.CodeModule.CountOfDeclarationLines -ge 0) {
        $component = $candidate
        if ([string]$candidate.Name -match 'ThisWorkbook|ЭтаКнига') { break }
      }
    }
  }
  if ($null -eq $component) { throw 'Не найден модуль рабочей книги для событий XLTM.' }
  $module = $component.CodeModule
  if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
  $module.AddFromString($Code)
}

function Set-ProfiDocumentProperties {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]$Workbook,
    [Parameter(Mandatory = $true)][string]$Title,
    [string]$Subject = 'ПрофиПомощник для Excel',
    [string]$Version = '1.2.0'
  )
  foreach ($pair in @(
    @('Title', $Title),
    @('Subject', $Subject),
    @('Comments', "Версия $Version. https://github.com/f2re/excel_helper")
  )) {
    try { $Workbook.BuiltinDocumentProperties.Item($pair[0]).Value = $pair[1] } catch { }
  }
}

function Add-ProfiStartButton {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]$Worksheet,
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][string]$Caption,
    [Parameter(Mandatory = $true)][string]$Macro,
    [double]$Left,
    [double]$Top,
    [double]$Width = 255,
    [double]$Height = 34
  )
  try { $Worksheet.Shapes.Item($Name).Delete() } catch { }
  $shape = $Worksheet.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
  $shape.Name = $Name
  $shape.OnAction = $Macro
  try {
    $shape.TextFrame.Characters().Text = $Caption
    $shape.TextFrame.HorizontalAlignment = -4108
    $shape.TextFrame.VerticalAlignment = -4108
  } catch { }
  try {
    $shape.Fill.ForeColor.RGB = 0x4A8630
    $shape.Line.ForeColor.RGB = 0x316420
    $shape.TextFrame.Characters().Font.Color = 0xFFFFFF
    $shape.TextFrame.Characters().Font.Bold = $true
    $shape.TextFrame.Characters().Font.Size = 11
  } catch { }
  return $shape
}

function Initialize-ProfiTemplateSheet {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]$Workbook,
    [string]$Version = '1.2.0'
  )
  $sheet = $Workbook.Worksheets.Item(1)
  $sheet.Name = 'Старт'
  $sheet.Cells.Clear()
  $sheet.Range('A1').Value2 = 'ПрофиПомощник для Excel'
  $sheet.Range('A2').Value2 = "Переносимый шаблон XLTM · версия $Version"
  $sheet.Range('A4').Value2 = 'Начните с кнопки «Создать проект». Служебные листы и таблицы будут созданы автоматически.'
  $sheet.Range('A5').Value2 = 'Для расписания откройте или скопируйте групповой лист, затем выберите группу вручную, по имени листа или из ячейки.'
  $sheet.Range('A6').Value2 = 'Если парсер ошибся, измените строки, столбцы и смещения в мастере либо в tblParserProfiles.'
  $sheet.Range('A1').Font.Bold = $true
  $sheet.Range('A1').Font.Size = 22
  $sheet.Range('A2').Font.Size = 12
  $sheet.Range('A4:A6').WrapText = $true
  $sheet.Columns.Item('A').ColumnWidth = 92
  $sheet.Rows.Item(4).RowHeight = 34
  $sheet.Rows.Item(5).RowHeight = 44
  $sheet.Rows.Item(6).RowHeight = 44
  [void](Add-ProfiStartButton -Worksheet $sheet -Name 'btnProfiCreate' -Caption '1. Создать или восстановить проект' -Macro 'ProfiTemplateCreateProject' -Left 20 -Top 170)
  [void](Add-ProfiStartButton -Worksheet $sheet -Name 'btnProfiSource' -Caption '2. Подключить текущий лист группы' -Macro 'ProfiTemplateAddSource' -Left 20 -Top 215)
  [void](Add-ProfiStartButton -Worksheet $sheet -Name 'btnProfiSchedule' -Caption '3. Составить сводное расписание' -Macro 'ProfiTemplateComposeSchedule' -Left 20 -Top 260)
  [void](Add-ProfiStartButton -Worksheet $sheet -Name 'btnProfiTeachers' -Caption 'Справочник преподавателей' -Macro 'ProfiTemplateOpenTeachers' -Left 300 -Top 170 -Width 220)
  [void](Add-ProfiStartButton -Worksheet $sheet -Name 'btnProfiHelp' -Caption 'Справка и диагностика' -Macro 'ProfiTemplateShowHelp' -Left 300 -Top 215 -Width 220)
  [void](Add-ProfiStartButton -Worksheet $sheet -Name 'btnProfiHidden' -Caption 'Показать / скрыть служебные листы' -Macro 'ProfiTemplateToggleServiceSheets' -Left 300 -Top 260 -Width 220)
  $sheet.Range('Z1').Value2 = $Version
  $sheet.Range('Z1').Name = 'PROFI_TEMPLATE_VERSION'
  $sheet.Columns.Item('Z').Hidden = $true
  $sheet.Activate()
  try { $Workbook.Application.ActiveWindow.DisplayGridlines = $false } catch { }
  return $sheet
}

function Close-ProfiComObject {
  [CmdletBinding()]
  param($Object)
  if ($null -ne $Object) {
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch { }
  }
}

Export-ModuleMember -Function Get-ProfiLegacyRoot, Get-ProfiSourceModules, New-ProfiExcelApplication, Assert-ProfiVbaProjectAccess, Import-ProfiVbaModules, Set-ProfiThisWorkbookCode, Set-ProfiDocumentProperties, Initialize-ProfiTemplateSheet, Close-ProfiComObject
