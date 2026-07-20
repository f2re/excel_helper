# 🛠 Локальная разработка

Этот документ описывает современный Office.js-контур. Для XLAM/XLTM см. [LEGACY_OFFICE.md](LEGACY_OFFICE.md), для macOS — [MACOS.md](MACOS.md).

## Требования

- Node.js 20 или новее;
- npm;
- OpenSSL для локального HTTPS;
- Microsoft Excel Microsoft 365 для проверки реальной книги;
- доверенный локальный сертификат.

## Первый запуск

```bash
npm ci
npm start
```

Команда заканчивается работающим HTTPS-сервером, а не готовой установкой в Excel. В терминале отображаются адреса:

| URL | Назначение |
|---|---|
| `https://localhost:3000/` | панель запуска и диагностика |
| `https://localhost:3000/src/ui/taskpane.html` | browser preview боковой панели |
| `https://localhost:3000/manifest.xml` | манифест Microsoft 365 |
| `https://localhost:3000/manifest-office2016.xml` | облегчённый манифест Excel 2016 |
| `https://localhost:3000/health` | JSON-статус сборки |

Сервер работает в foreground. Для остановки нажмите `Ctrl+C`.

## Что происходит по шагам

### 1. Автоматическая генерация

`npm run build`, `npm run build:dev` и `npm start` всегда запускают генератор перед копированием файлов. Генерируются:

```text
src/functions.json
docs/FUNCTIONS.md
docs/APPLETS.md
```

Отдельная команда `npm run generate` нужна только для ручной проверки результата.

### 2. Сборка

`npm run build:dev` заменяет `__BASE_URL__` на `https://localhost:3000` и формирует:

```text
dist/
├── index.html
├── build-info.json
├── manifest.xml
├── manifest-office2016.xml
├── functions.json
├── assets/
└── src/
```

`npm run build` использует production URL GitHub Pages.

### 3. Сертификат

`npm run certs` создаёт:

```text
.certs/key.pem
.certs/cert.pem
```

Скрипт повторно использует действующий сертификат и пересоздаёт его, если срок действия подходит к концу.

На macOS доверие добавляется командой:

```bash
npm run certs:trust:mac
```

### 4. Сервер и watch

`npm start` запускает `scripts/dev-server.mjs` с параметрами `--watch --open`.

Сервер:

- обслуживает только `dist/`;
- запрещает выход за корневой каталог;
- отдаёт правильные MIME-типы;
- отключает кэш;
- предоставляет `/health`;
- следит за `src/`, `assets/` и манифестами;
- при изменении повторно собирает проект;
- перезагружает browser preview через Server-Sent Events.

Excel task pane автоматически не перезагружается, чтобы не прерывать операции с книгой. Закройте и снова откройте панель либо перезагрузите её через инструменты разработчика.

## Browser preview и Excel runtime

### Browser preview

Открывается по адресу `https://localhost:3000/src/ui/taskpane.html`.

Работают:

- вкладки;
- поиск;
- каталог функций;
- каталог апплетов;
- копирование примеров формул;
- стенд вычисления функций;
- демонстрация flow расписания;
- ссылки на документацию.

Не работают операции, которым требуется `Excel.run`:

- создание служебных листов;
- изменение выбранного диапазона;
- запуск массового апплета;
- импорт группового расписания;
- формирование итоговых листов.

Панель явно показывает режим «Браузерный preview», поэтому отсутствие Excel API не выглядит как поломка.

### Внутри Excel

После загрузки `dist/manifest.xml` панель определяет Excel runtime и включает команды книги. В шапке появляется статус «Excel подключён».

## Загрузка в Excel

### macOS

```bash
npm run start:mac
```

Команда:

1. делает dev-сборку;
2. создаёт и доверяет сертификату;
3. копирует манифест в папку `wef` Excel;
4. запускает сервер и browser dashboard.

Перезапустите Excel после первого sideload.

### Ручной macOS sideload

```bash
npm run sideload:mac
npm run start:no-open
```

Удаление:

```bash
npm run unsideload:mac
```

### Windows и Excel Web

Можно использовать стандартный `office-addin-debugging`:

```bash
npm start
npx --yes office-addin-debugging@6.1.1 start dist/manifest.xml desktop --debug-method direct --no-live-reload
```

После завершения отладки снимите регистрацию:

```bash
npx --yes office-addin-debugging@6.1.1 stop dist/manifest.xml
```

`npx` здесь является опциональным инструментом, поэтому пакет не включён в runtime-зависимости проекта.

## Полная проверка

```bash
npm run check
```

Порядок:

1. синтаксис JavaScript;
2. legacy audit;
3. distribution audit;
4. unit-тесты;
5. production-сборка с автоматической генерацией;
6. валидация структуры и манифестов;
7. ссылки документации;
8. smoke-тест `dist/`.

## Команды

| Команда | Результат |
|---|---|
| `npm start` | dev-сервер с watch и браузером |
| `npm run start:no-open` | dev-сервер без открытия браузера |
| `npm run serve` | сервер для готового `dist/` |
| `npm run build` | production `dist/` |
| `npm run build:dev` | localhost `dist/` |
| `npm run generate` | метаданные и каталоги |
| `npm run check` | полный контроль |
| `npm run certs` | создание сертификата |
| `npm run certs:trust:mac` | доверие сертификату в macOS |
| `npm run sideload:mac` | установка dev-манифеста Excel Mac |
| `npm run unsideload:mac` | удаление dev-манифеста Excel Mac |

## Типовой цикл изменения функции

1. Измените `src/core/function-catalog.js` и `src/core/functions-core.js`.
2. Добавьте тест.
3. При работающем `npm start` сохраните файл.
4. Watch автоматически запустит генератор и пересборку.
5. Browser preview перезагрузится.
6. Проверьте функцию в стенде.
7. Проверьте функцию внутри Excel.
8. Выполните `npm run check`.

## Типовой цикл изменения панели

1. Измените `src/ui/taskpane.html`, `taskpane.css` или `taskpane.js`.
2. Сохраните файл.
3. Открытый browser preview перезагрузится автоматически.
4. После визуальной проверки откройте панель в Excel и проверьте `Excel.run`-операции.
