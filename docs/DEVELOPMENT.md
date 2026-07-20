# 🧑‍💻 Разработка

## Первый запуск

```bash
npm ci
npm start
```

`npm start` автоматически генерирует метаданные функций, собирает localhost-версию, создаёт HTTPS-сертификат и запускает browser dashboard с watch-режимом. Подробности: [LOCAL_DEVELOPMENT.md](LOCAL_DEVELOPMENT.md).

Полная проверка:

```bash
npm run check
```

Проект не использует runtime-зависимости npm. Скрипты сборки и тестирования работают на Node.js 20.

## Новая функция

1. добавьте определение в `src/core/function-catalog.js`;
2. добавьте реализацию в `src/core/functions-core.js`;
3. добавьте unit-тест;
4. для legacy-паритета добавьте VBA-UDF в `legacy-vba/src/modProfiFunctions.bas`;
5. сохраните файл при работающем `npm start` — генерация и browser preview обновятся автоматически;
6. проверьте функцию в browser-стенде и внутри Excel;
7. выполните `npm run check`.

`npm run generate` остаётся доступным для ручной регенерации, но `build`, `build:dev`, `start` и `check` больше от него не зависят как от отдельного предварительного шага.

## Новый апплет

Добавьте элемент в `src/core/applet-catalog.js` и обработчик в `src/actions/applets.js`. Апплет не должен молча изменять исходные листы. В browser preview он должен отображаться в каталоге, а при попытке изменить книгу вне Excel — выдавать понятное ограничение.

## Панель

- `src/ui/taskpane.html` — структура панели;
- `src/ui/taskpane.css` — стили;
- `src/ui/taskpane.js` — поиск, каталог, browser preview и Excel-команды;
- `src/dev/` — dashboard локального сервера;
- `scripts/dev-server.mjs` — HTTPS, `/health`, watch и live reload.

После изменения панели browser preview перезагружается автоматически. Excel task pane проверяйте отдельно, поскольку обычный браузер не предоставляет `Excel.run`.

## Новый шаблон расписания

Расширяйте `detectProfile` и сохраняйте ручные параметры в `tblParserProfiles`. Автоматическое правило должно иметь диагностику и не отменять подтверждённое ручное значение. Добавляйте кейс в [USE_CASES.md](USE_CASES.md).

## macOS

Современный контур разрабатывается без PowerShell:

```bash
npm run start:mac
npm run package:modern
```

XLAM/XLTM из `.bas` автоматически компилируются только Windows Excel. Подробности и варианты библиотек: [MACOS.md](MACOS.md).

## Общие VBA-модули

Каталог `legacy-vba/src` является единственным источником VBA-кода:

- общие модули импортируются в XLAM и XLTM;
- `modProfiMenu.bas` исключается из XLTM;
- `modProfiTemplate.bas` исключается из XLAM;
- события `Workbook_Open` XLTM добавляются сборщиком;
- пользовательские данные всегда записываются в результат `ProfiHostWorkbook`.

Не создавайте отдельные копии одного VBA-модуля для XLAM и XLTM. Различия оформляйте небольшими package-specific модулями.

## Правила совместимости VBA

- не используйте WinAPI без условных VBA7/x86/x64-объявлений;
- предпочтительно полностью избегайте `Declare`, `PtrSafe`, `LongPtr` и `LongLong`;
- используйте позднее связывание: `CreateObject("Scripting.Dictionary")`;
- не используйте `ThisWorkbook` как безусловный источник пользовательских данных;
- не добавляйте ссылки, отсутствующие в Office 2010;
- проверяйте проект в 32- и 64-разрядном Excel;
- не храните персональные данные в тестах.

## Сборка legacy

```powershell
npm run legacy:prereq
npm run legacy:build
npm run legacy:build-template
npm run legacy:build-all
npm run legacy:verify
```

Сборка требует Windows, Excel и доверенного доступа к VBProject. После сборки этот параметр следует выключить.

## Установщик

PowerShell-ядро находится в `installer/windows`. Оно должно:

- работать без прав администратора;
- поддерживать `Full`, `LegacyFull`, `AddinOnly`, `TemplateOnly`, `ModernOnly`, `AddinModern`, `TemplateModern`;
- проверять закрытие Excel только для операций XLAM;
- создавать резервную копию заменяемых файлов;
- откатываться при ошибке;
- иметь симметричный uninstall и repair;
- не изменять корпоративные политики макросов.

Inno Setup является оболочкой над теми же PowerShell-скриптами. Бизнес-логику не следует дублировать в `.iss`.

## CI

- Linux CI: Node-тесты, документация, манифесты и web-сборка;
- Windows-hosted CI: те же проверки плюс PowerShell parser;
- installer compile smoke: компиляция Inno Setup на GitHub-hosted Windows;
- Windows self-hosted + Excel: настоящие XLAM/XLTM и COM-проверка.
