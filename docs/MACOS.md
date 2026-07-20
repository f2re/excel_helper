# 🍎 Разработка и сборка на macOS

Современная Office.js-надстройка полностью разрабатывается, тестируется и собирается на macOS без PowerShell. Windows требуется только для автоматической компиляции исходников VBA в бинарные XLAM/XLTM и для сборки EXE-установщика.

## Что можно сделать полностью на Mac

| Задача | macOS |
|---|---:|
| генерация `functions.json` | ✅ |
| browser preview панели | ✅ |
| sideload `manifest.xml` в Excel Mac | ✅ |
| Custom Functions в поддерживаемом Excel Mac | ✅ |
| unit-тесты и `npm run check` | ✅ |
| production `dist/` | ✅ |
| ZIP современной Office.js-сборки | ✅ |
| анализ VBA-исходников | ✅ |
| компиляция `.bas` в `.xlam/.xltm` текущими скриптами | ❌, нужен Windows Excel COM |
| EXE/Inno Setup | ❌, Windows |

## Подготовка

```bash
xcode-select --install
node --version
npm --version
openssl version
```

Рекомендуется Node.js 20 или новее. Если Node установлен через Homebrew:

```bash
brew install node
```

## Самый короткий путь

```bash
git clone https://github.com/f2re/excel_helper.git
cd excel_helper
npm ci
npm run start:mac
```

`start:mac` выполняет:

1. `build:dev` с адресом `https://localhost:3000`;
2. создание локального сертификата;
3. добавление сертификата в login keychain;
4. копирование `dist/manifest.xml` в папку sideload Excel;
5. запуск HTTPS-сервера и browser dashboard.

После первого запуска полностью закройте и снова откройте Excel.

## Ручной вариант

### 1. Создать и доверить сертификат

```bash
npm run certs:trust:mac
```

Сертификаты находятся в `.certs/`. Команда `security add-trusted-cert` может запросить пароль пользователя macOS.

### 2. Собрать dev-версию

```bash
npm run build:dev
```

### 3. Установить манифест

```bash
npm run sideload:mac
```

Скрипт копирует манифест в:

```text
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

Путь соответствует официальной процедуре Microsoft: [Sideload Office Add-ins on Mac for testing](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).

### 4. Запустить сервер

```bash
npm run start:no-open
```

### 5. Открыть панель в Excel

Перезапустите Excel, откройте книгу и выберите надстройку на вкладке **Главная → Надстройки**.

Удаление dev-манифеста:

```bash
npm run unsideload:mac
```

## Production-сборка на Mac

```bash
npm ci
npm run check
npm run build
npm run package:modern
```

Результаты:

```text
dist/
release/profi-excel-helper-modern-1.2.0.zip
```

`package:modern` использует системную команду `zip`. В стандартной macOS она обычно уже установлена.

## Опциональные инструменты Microsoft

Проект не добавляет эти пакеты в обязательные зависимости, но их можно запускать через `npx`.

### office-addin-debugging

Регистрирует и загружает манифест в Excel:

```bash
npm start
npx --yes office-addin-debugging@6.1.1 start dist/manifest.xml desktop --debug-method direct --no-live-reload
```

Остановка и снятие регистрации:

```bash
npx --yes office-addin-debugging@6.1.1 stop dist/manifest.xml
```

Microsoft рекомендует завершать отладочную сессию командой `stop`, поскольку простое закрытие терминала не удаляет dev-регистрацию.

### office-addin-dev-certs

Альтернатива встроенному генератору сертификатов:

```bash
npx --yes office-addin-dev-certs@2.0.10 install
```

### office-addin-manifest

Дополнительная официальная проверка манифеста:

```bash
npx --yes office-addin-manifest@2.1.6 validate dist/manifest.xml
```

## Почему XLAM/XLTM не собираются автоматически на Mac

VBA-исходники в репозитории представлены файлами `.bas`. Чтобы получить реальный `.xlam` или `.xltm`, Excel должен создать бинарный `vbaProject.bin` и связать его с OOXML-пакетом.

Текущий воспроизводимый процесс использует:

```text
PowerShell → Excel.Application COM → VBProject → SaveAs XLAM/XLTM
```

COM Automation доступна только в Windows. Простое создание ZIP с расширением `.xlam` не создаёт исполняемый VBA-проект.

## Реальные варианты для legacy-пакетов при разработке на Mac

### Вариант A — рекомендованный

Разрабатывать код и документацию на Mac, а бинарные XLAM/XLTM собирать:

- на выделенной Windows VM;
- на self-hosted GitHub Actions runner с установленным Excel;
- на отдельной машине сборки организации.

В репозитории для этого есть workflow `Build Legacy Office packages`.

### Вариант B — ручная сборка в Excel

Можно вручную открыть редактор VBA в настольном Excel, импортировать `.bas` и сохранить файл в нужном формате. Такой процесс не является полностью воспроизводимым и требует отдельной проверки RibbonX.

### Вариант C — заранее скомпилированный seed

Библиотеки могут переносить уже готовый `vbaProject.bin`:

- [XlsxWriter — Working with VBA Macros](https://xlsxwriter.readthedocs.io/working_with_macros.html);
- [openpyxl `keep_vba`](https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file).

Они **не компилируют `.bas` в VBA-проект**. Такой подход возможен только при наличии предварительно собранного seed-файла, который затем нужно обновлять при изменениях VBA.

### Вариант D — xlwings

[xlwings](https://docs.xlwings.org/) умеет автоматизировать Excel на macOS и Windows для работы с книгами. Он полезен для интеграционных тестов и подготовки данных, но не является гарантированной заменой Windows VBProject/COM-сборке XLAM/XLTM.

### Что не рекомендуется

- переименовывать `.xlsx` в `.xlam` или `.xltm`;
- считать LibreOffice эквивалентом Excel для проверки VBA и RibbonX;
- хранить вручную собранный бинарный файл без исходников и проверки версии;
- распространять неподписанные макросы в корпоративной среде.

## Типовые проблемы macOS

### Браузер предупреждает о сертификате

```bash
npm run certs:trust:mac
```

Затем полностью перезапустите браузер и Excel.

### Надстройка не появилась

1. Проверьте наличие XML в папке `wef`.
2. Закройте все процессы Excel.
3. Выполните `npm run sideload:mac` повторно.
4. Убедитесь, что `https://localhost:3000/health` открывается без предупреждения.
5. При необходимости очистите кэш Office.

### Панель открывается, но кнопки книги недоступны

Вы открыли browser preview либо Excel не инициализировал Office.js. В верхней части панели должен отображаться статус **Excel подключён**.

### `PROFI.*` возвращает `#NAME?`

Проверьте поддержку Custom Functions конкретной редакцией Excel и загрузку `functions.json`. Для старых perpetual-редакций Windows используйте XLAM-функции `PROFI_*`.
