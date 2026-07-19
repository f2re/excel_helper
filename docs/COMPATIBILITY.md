# 🖥 Совместимость

## Целевые клиенты

| Возможность | Windows | macOS | Excel для браузера | iPad/iOS |
|---|---:|---:|---:|---:|
| Task pane и апплеты | ✅ Microsoft 365 | ✅ Microsoft 365 | ✅ | зависит от API, не целевая платформа |
| Shared runtime | ✅ при поддерживаемой сборке | ✅ при поддерживаемой версии | ✅ | не целевая платформа |
| Пользовательские функции | ✅ Microsoft 365 | ✅ | ✅ | ❌ по документации Microsoft |
| Импорт листов Base64 | ✅ | ✅ | ✅ с ограничениями | ❌ |
| Fallback: скопировать лист вручную | ✅ | ✅ | ✅ | зависит от клиента |

Актуальная матрица Microsoft может меняться с каналом обновления, поэтому надстройка проверяет критичные requirement sets во время выполнения.

## Пользовательские функции

Microsoft указывает, что custom functions Office.js не поддерживаются в Office на iPad и volume-licensed perpetual Office 2021 или более ранних версиях Windows:

- [Create custom functions in Excel — Supported platforms](https://learn.microsoft.com/office/dev/add-ins/excel/custom-functions-overview#supported-platforms)

## Shared runtime

Надстройка требует `SharedRuntime 1.1` и использует один runtime для task pane, функций и команд:

- [Configure your Office Add-in to use a shared runtime](https://learn.microsoft.com/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Shared runtime requirement sets](https://learn.microsoft.com/office/dev/add-ins/reference/requirement-sets/shared-runtime-requirement-sets)

## Импорт внешней книги

`Workbook.insertWorksheetsFromBase64` поддерживается в Excel на Windows, Mac и web, но не на iOS. В Excel для браузера исходные листы с PivotTable, Chart, Comment или Slicer могут быть отклонены:

- [Manage Excel workbooks with the Excel JavaScript API](https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-workbooks)

Fallback встроен в пользовательский flow: скопировать лист в текущую книгу и выбрать его в мастере.

## Локали

Интерфейс и display names функций русские. Формулы вставляются через `formulasLocal`, поэтому разделитель аргументов определяется локалью Excel. В справочнике примеры приведены для русской локали с `;`.

## Защищённые книги

Создание листов и таблиц невозможно, если структура книги защищена. Перед запуском мастера снимите защиту структуры или предоставьте надстройке рабочую незапрещённую книгу.
