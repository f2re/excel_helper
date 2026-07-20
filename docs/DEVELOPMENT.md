# 🧑‍💻 Разработка

## Проверка

```bash
npm ci
npm run check
```

Проект не использует runtime-зависимости npm. Скрипты сборки и тестирования работают на Node.js 20.

## Новая функция

1. добавьте имя в `src/core/function-catalog.js`;
2. добавьте реализацию в `src/core/functions-core.js`;
3. добавьте тест;
4. для legacy-паритета добавьте VBA-UDF в `legacy-vba/src/modProfiFunctions.bas`;
5. выполните `npm run generate` и `npm run check`.

## Новый апплет

Добавьте элемент в `src/core/applet-catalog.js` и обработчик в `src/actions/applets.js`. Апплет не должен молча изменять исходные листы.

## Новый шаблон расписания

Расширяйте `detectProfile` и сохраняйте ручные параметры в `tblParserProfiles`. Автоматическое правило должно иметь диагностику и не отменять ручное значение.

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
npm run legacy:build
npm run legacy:build-template
npm run legacy:build-all
npm run legacy:verify
```

Сборка требует Windows, Excel и доверенного доступа к VBProject. После сборки этот параметр следует выключить.

## Установщик

PowerShell-ядро находится в `installer/windows`. Оно должно:

- работать без прав администратора;
- поддерживать `Full`, `AddinOnly`, `TemplateOnly`;
- проверять закрытие Excel;
- создавать резервную копию заменяемых файлов;
- откатываться при ошибке;
- иметь симметричный uninstall;
- не изменять корпоративные политики макросов.

Inno Setup является оболочкой над теми же PowerShell-скриптами. Бизнес-логику не следует дублировать в `.iss`.

## CI

- Linux CI: Node-тесты, документация, манифесты и web-сборка;
- Windows-hosted CI: PowerShell parser и статическая проверка поставки;
- Windows self-hosted + Excel: настоящие XLAM/XLTM и COM-проверка;
- Inno Setup на self-hosted runner: EXE-установщик.
