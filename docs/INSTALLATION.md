# ⬇️ Установка и развёртывание

## Выбор варианта

| Сценарий | Файл или команда |
|---|---|
| Разработка современной панели | `npm start` |
| Microsoft 365, web или macOS | `dist/manifest.xml` |
| Excel 2016 с минимальной web-панелью | `dist/manifest-office2016.xml` |
| Команды во всех книгах Excel 2010–2019 Windows | `ProfiExcelHelper-Legacy.xlam` |
| Один переносимый файл-проект | `ProfiExcelHelper-Template.xltm` |
| Обычная установка «Далее → Готово» | `ProfiExcelHelper-Setup-1.2.0.exe` |
| Установка без EXE | `ProfiExcelHelper-Portable-1.2.0.zip` |

## Современная Office.js-надстройка

### Локальная разработка

```bash
git clone https://github.com/f2re/excel_helper.git
cd excel_helper
npm ci
npm start
```

`npm start` не устанавливает надстройку в Excel автоматически. Команда:

1. генерирует `src/functions.json` и каталоги;
2. собирает `dist/` для `https://localhost:3000`;
3. создаёт локальный HTTPS-сертификат;
4. запускает сервер, watch и browser dashboard.

Откройте:

```text
https://localhost:3000/
```

Browser preview работает без Excel, но не может изменять книгу. Для реальной работы загрузите `dist/manifest.xml` в Excel.

Подробности: [LOCAL_DEVELOPMENT.md](LOCAL_DEVELOPMENT.md).

### Production

```bash
npm ci
npm run check
npm run build
```

Результат находится в `dist/`. ZIP современной сборки на macOS/Linux:

```bash
npm run package:modern
```

## macOS

Самый короткий вариант без PowerShell:

```bash
npm ci
npm run start:mac
```

Команда доверяет сертификату, копирует dev-манифест в папку `wef` Excel и запускает сервер.

Ручные шаги:

```bash
npm run certs:trust:mac
npm run build:dev
npm run sideload:mac
npm run start:no-open
```

Удаление dev-манифеста:

```bash
npm run unsideload:mac
```

Полная инструкция и ограничения XLAM/XLTM: [MACOS.md](MACOS.md).

## Загрузка манифеста Microsoft 365

### Excel Desktop

Используйте штатную процедуру sideload вашей платформы или официальный инструмент:

```bash
npm start
npx --yes office-addin-debugging@6.1.1 start dist/manifest.xml desktop --debug-method direct --no-live-reload
```

После тестирования:

```bash
npx --yes office-addin-debugging@6.1.1 stop dist/manifest.xml
```

### Excel Web

Разместите содержимое `dist/` на HTTPS-хостинге и загрузите production-манифест. Для локальной разработки можно использовать официальный `office-addin-debugging` с платформой `web` и документом в OneDrive/SharePoint.

## Excel 2016 task pane

Используйте `dist/manifest-office2016.xml` только на клиентах, где доступен `ExcelApi 1.1`.

Этот манифест:

- не объявляет Custom Functions;
- не использует Shared Runtime;
- открывает облегчённую ES5-панель.

Для полного набора legacy-команд на Excel 2016 Windows установите XLAM.

## Ручная установка XLAM

1. Закройте Excel.
2. Откройте Excel заново.
3. Перейдите: **Файл → Параметры → Надстройки**.
4. Внизу выберите **Надстройки Excel → Перейти**.
5. Нажмите **Обзор** и выберите `ProfiExcelHelper-Legacy.xlam`.
6. Убедитесь, что надстройка отмечена флажком.
7. Проверьте вкладку или меню **ПрофиПомощник**.

## Использование XLTM

1. Откройте `ProfiExcelHelper-Template.xltm`.
2. Excel создаст новую книгу на основе шаблона.
3. Нажмите «Создать или восстановить проект».
4. Сохраните рабочую книгу как `.xlsm`, чтобы VBA-код сохранился.

Не сохраняйте рабочую книгу как `.xlsx`: этот формат удалит VBA-проект.

## Сборка XLAM и XLTM

Требуется Windows с установленным настольным Excel:

```powershell
npm run legacy:prereq
npm run legacy:build-all
npm run legacy:verify
```

Результаты:

```text
legacy-vba/dist/ProfiExcelHelper-Legacy.xlam
legacy-vba/dist/ProfiExcelHelper-Template.xltm
```

Программный доступ к VBProject следует включать только на время сборки.

## PowerShell-установка Windows

После подготовки `release/payload`:

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\install.ps1 `
  -PayloadRoot release\payload `
  -Mode Full
```

Поддерживаемые режимы:

| Режим | Компоненты |
|---|---|
| `Full` | XLAM + XLTM + Office.js payload |
| `LegacyFull` | XLAM + XLTM |
| `AddinOnly` | только XLAM |
| `TemplateOnly` | только XLTM |
| `ModernOnly` | только Office.js payload |
| `AddinModern` | XLAM + Office.js |
| `TemplateModern` | XLTM + Office.js |

Установка выполняется в профиль текущего пользователя, определяет Excel и его разрядность, создаёт резервную копию и откатывает изменения при ошибке.

## EXE-установщик

```powershell
npm run build
npm run legacy:build-all
npm run installer:bundle
npm run installer:build
```

Для EXE нужен Inno Setup 6. Администраторские права не требуются.

Подпись:

```powershell
npm run installer:sign -- `
  -InstallerPath release\ProfiExcelHelper-Setup-1.2.0.exe `
  -CertificateThumbprint YOUR_THUMBPRINT
```

## Удаление и восстановление

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\uninstall.ps1
powershell -ExecutionPolicy Bypass -File installer\windows\repair.ps1 -PayloadRoot release\payload
```

## Безопасность макросов

Перед массовым развёртыванием:

- подпишите XLAM, XLTM и EXE сертификатом организации;
- используйте доверенное расположение или доверенного издателя;
- не отключайте системные политики безопасности Office;
- не распространяйте книги с реальными персональными данными;
- храните журналы из `%LOCALAPPDATA%\ProfiExcelHelper\logs`;
- проверяйте контрольные суммы релиза.

## Проверка установки

### Office.js

1. Откройте `https://localhost:3000/health`.
2. Убедитесь, что указаны 114 функций и 44 апплета.
3. Откройте browser preview — вкладки не должны быть пустыми.
4. Загрузите манифест в Excel.
5. Проверьте статус «Excel подключён».
6. Выполните `=PROFI.ФИОКРАТКО(A1)`.

### XLAM/XLTM

1. Проверьте наличие вкладки Ribbon.
2. Создайте чистую книгу.
3. Нажмите «Проверить/создать проект».
4. Убедитесь, что `_PROFI_*` листы созданы.
5. Подключите тестовый групповой лист и сформируйте расписание.
