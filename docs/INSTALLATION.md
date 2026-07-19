# ⬇️ Установка, локальный запуск и развёртывание

## 1. Требования

| Компонент | Требование |
|---|---|
| Excel | Microsoft 365 Excel с поддержкой Office Add-ins и SharedRuntime 1.1 |
| Node.js | 20 или новее |
| npm | Поставляется с Node.js |
| OpenSSL | Нужен только для создания локального development-сертификата |
| Браузер | Современный Edge, Chrome или Safari для Excel в браузере |

Проверка окружения:

```bash
node --version
npm --version
openssl version
```

## 2. Получение и проверка исходного кода

```bash
git clone https://github.com/f2re/excel_helper.git
cd excel_helper
npm ci
npm run check
```

Ожидаемый результат:

- TypeScript проходит без ошибок;
- тесты показывают успешные проверки;
- создаётся папка `dist/`;
- внутренний валидатор подтверждает манифест и 114 функций.

## 3. Локальный HTTPS

```bash
npm start
```

Команда выполняет development-сборку, создаёт сертификат и запускает сервер:

```text
https://localhost:3000/taskpane.html
```

Сертификаты находятся в:

```text
.certs/localhost-cert.pem
.certs/localhost-key.pem
```

### Доверие сертификату

Самоподписанный сертификат должен быть добавлен в доверенные корневые сертификаты **только на машине разработки**.

#### Windows PowerShell от администратора

```powershell
certutil -addstore -f Root .certs\localhost-cert.pem
```

#### macOS

```bash
sudo security add-trusted-cert -d -r trustRoot \
  -k /Library/Keychains/System.keychain \
  .certs/localhost-cert.pem
```

Также можно открыть сертификат в Keychain Access и установить доверие вручную.

> Закрытый ключ `.certs/localhost-key.pem` не следует копировать или публиковать.

## 4. Установка в Excel для браузера

1. Оставьте `npm start` запущенным.
2. Откройте книгу в Excel for the web.
3. Выберите **Главная → Надстройки → Дополнительные параметры**.
4. Выберите **Загрузить мою надстройку**.
5. Укажите `dist/manifest.xml`.
6. На ленте появится вкладка **ПрофиПомощник**.

Официальная инструкция Microsoft: [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

## 5. Локальный запуск в Excel Desktop

### Автоматизированный вариант

```bash
npm run start:desktop
```

Команда использует официальный пакет `office-addin-debugging` через `npx`. Для остановки и удаления development-регистрации:

```bash
npm run stop:desktop
```

### Ручной вариант

Можно использовать доверенный сетевой каталог или административное развёртывание XML-манифеста. Статические файлы надстройки при этом всё равно должны быть доступны по HTTPS по адресам, указанным в манифесте.

Официальная навигация по вариантам: [Deploy and publish Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/publish/publish).

## 6. Production-сборка

Укажите публичный HTTPS-адрес, по которому будет размещено содержимое `dist/`:

```bash
ADDIN_BASE_URL=https://excel-helper.example.org/ npm run build
```

Сборщик:

- создаст `dist/taskpane.html`, `taskpane.js`, CSS, метаданные функций и иконки;
- заменит `https://localhost:3000/` в копии манифеста на `ADDIN_BASE_URL`;
- сохранит production-манифест в `dist/manifest.xml`.

Публикация:

1. загрузите **всё содержимое `dist/`** на HTTPS-хостинг;
2. проверьте доступность `taskpane.html`, `functions.json` и всех иконок;
3. передайте `dist/manifest.xml` администратору Microsoft 365 или разместите его выбранным способом;
4. после изменения XML-манифеста распространите его заново; обычные изменения JavaScript/CSS достаточно развернуть на веб-хостинге.

Microsoft описывает Office Add-in как веб-приложение плюс манифест; для production веб-файлы размещаются на сервере, а адрес в манифесте указывает на этот сервер: [Develop Office Add-ins with Visual Studio Code](https://learn.microsoft.com/office/dev/add-ins/develop/develop-add-ins-vscode).

## 7. Проверка манифеста

Локальная обязательная проверка:

```bash
npm run validate
```

Официальная проверка схемы Microsoft, требующая доступа в интернет:

```bash
npm run validate:office
```

## 8. Обновление установленной версии

```bash
git pull
npm ci
npm run check
```

Для локальной версии перезапустите сервер. Для production загрузите новый `dist/` на тот же URL. Если изменились идентификаторы, команды, разрешения или адреса в манифесте, распространите новый `dist/manifest.xml`.

## 9. Совместимость

### Пользовательские функции

Microsoft указывает, что пользовательские функции Office.js поддерживаются большинством актуальных клиентов Excel, но не Excel на iPad и не бессрочными корпоративными выпусками Office 2021 и старше: [Create custom functions in Excel](https://learn.microsoft.com/office/dev/add-ins/excel/custom-functions-overview#supported-platforms).

### Shared runtime

Надстройка использует SharedRuntime 1.1, чтобы команды ленты, task pane и функции работали в одном browser runtime: [Configure your Office Add-in to use a shared runtime](https://learn.microsoft.com/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime).

### Импорт внешней книги

Прямой импорт листов использует `Workbook.insertWorksheetsFromBase64`:

- поддерживается в Excel для Windows, macOS и браузера;
- не поддерживается в iOS;
- в Excel для браузера исходные листы с PivotTable, Chart, Comment или Slicer могут вернуть `UnsupportedFeature`.

Источник: [Manage Excel workbooks with the Excel JavaScript API](https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-workbooks).

В любом несовместимом случае предусмотрен рабочий fallback: скопировать нужный лист в текущую книгу, затем выбрать его в мастере.
