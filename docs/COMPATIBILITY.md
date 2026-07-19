# 🖥 Совместимость

## Матрица

| Версия | Современный манифест | Совместимая панель | Legacy XLAM |
|---|---:|---:|---:|
| Microsoft 365 Windows | ✅ | не требуется | опционально |
| Microsoft 365 macOS/Web | ✅ | не требуется | нет гарантии |
| Excel 2016 Windows | зависит от канала и обновлений | ✅ при ExcelApi 1.1 | ✅ рекомендуется |
| Excel 2013 Windows | ❌ | ❌ | ✅ |
| Excel 2010 Windows | ❌ | ❌ | ✅ |

Microsoft не выпускала Office 2012. В разговорной формулировке обычно подразумевается Office 2010 или Office 2013.

## Почему нужны два контура

Office.js Custom Functions, Shared Runtime и новые ExcelApi requirement sets появились позже и не могут быть одинаково доступны в старых perpetual-версиях. Поэтому:

- современный контур использует `manifest.xml`, 114 функций `PROFI`, shared runtime и полную панель;
- Excel 2016 получает task-pane-only `manifest-office2016.xml` на ExcelApi 1.1;
- Excel 2010/2013/2016 Windows получают VBA/XLAM-контур.

## Ограничения legacy

- VBA-UDF используют английские имена с префиксом `PROFI_`, например `PROFI_FIO_SHORT`;
- реализован основной востребованный набор UDF, а массовые операции и расписание доступны через меню;
- установка и автоматическая сборка ориентированы на Windows COM Automation;
- старые выпуски Office уже не получают обновления безопасности от Microsoft;
- XLAM с макросами должен быть подписан и развёрнут из доверенного расположения.

Актуальные ссылки Microsoft:

- [Office Add-ins requirement sets](https://learn.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Excel JavaScript API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/excel/excel-api-requirement-sets)
- [Custom Functions requirements](https://learn.microsoft.com/office/dev/add-ins/excel/custom-functions-requirements)
