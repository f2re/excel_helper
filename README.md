# 🧰 ПрофиПомощник для Excel

Надстройка для кадровых служб, рекрутинга, преподавателей, работников и учебных подразделений. Проект объединяет пользовательские функции, массовые апплеты и составитель сводного преподавательского расписания.

## ✨ Возможности

- **114 функций `PROFI`** для ФИО, кадров, рекрутинга, занятий, задач и очистки данных;
- **44 апплета** для массовых операций и готовых рабочих листов;
- составитель сводного расписания с автоматически создаваемыми `_PROFI_*` листами;
- группа вручную, из имени листа или из выбранной ячейки;
- автоматическое распознавание сетки и легенды с ручной правкой каждого параметра;
- алиасы преподавателей, потоковые занятия, детерминированный Round-Robin и контроль конфликтов;
- современная Office.js-панель, Excel 2016 task pane, XLAM, XLTM, переносимый ZIP и EXE-установщик;
- поддержка разработки на **macOS без PowerShell** для современного Office.js-контура.

> Отдельного выпуска Microsoft Office 2012 не существовало. Для компьютеров этого периода рассматриваются Office 2010 или Office 2013.

## 📦 Варианты поставки

| Клиент | Пакет | Возможности |
|---|---|---|
| Microsoft 365 Excel Windows/macOS/Web | `manifest.xml` | полная панель, Custom Functions, апплеты и расписание |
| Excel 2016 Windows с ExcelApi 1.1 | `manifest-office2016.xml` | облегчённая ES5-панель |
| Excel 2010/2013/2016/2019 Windows | `ProfiExcelHelper-Legacy.xlam` | глобальные VBA-UDF, Ribbon и составитель |
| Excel 2010/2013/2016/2019 Windows | `ProfiExcelHelper-Template.xltm` | автономный проект из одного шаблона |
| Windows | `ProfiExcelHelper-Setup-1.2.0.exe` | per-user установка выбранных компонентов |
| Windows без EXE | `ProfiExcelHelper-Portable-1.2.0.zip` | переносимый пакет и сценарии установки |

Подробнее: [варианты поставки](docs/DISTRIBUTION.md), [совместимость](docs/COMPATIBILITY.md), [legacy-режим](docs/LEGACY_OFFICE.md).

## 🚀 Быстрый старт современной надстройки

```bash
git clone https://github.com/f2re/excel_helper.git
cd excel_helper
npm ci
npm start
```

После запуска автоматически откроется:

```text
https://localhost:3000/
```

Это **панель разработки**, а не сама надстройка Excel. Она показывает состояние сборки, ссылки на манифесты и открывает функциональный browser preview боковой панели.

### Что именно делает `npm start`

1. запускает генератор `src/functions.json` и справочников функций/апплетов;
2. собирает `dist/` с локальными HTTPS-URL;
3. создаёт локальный сертификат в `.certs/`, если его ещё нет;
4. запускает HTTPS-сервер на порту `3000`;
5. включает слежение за `src/`, манифестами и ресурсами;
6. открывает локальную панель запуска в браузере.

Browser preview позволяет:

- пользоваться поиском;
- просматривать 114 функций и 44 апплета;
- вычислять функции в стенде без Excel;
- проверять форму и flow составителя расписания;
- открывать манифесты и документацию.

Операции, которые читают или изменяют книгу, работают только после загрузки `dist/manifest.xml` в Excel.

## 🧭 Команды разработки

| Команда | Что делает |
|---|---|
| `npm start` | dev-сборка, сертификат, сервер, watch и открытие браузера |
| `npm run start:no-open` | то же, но без автоматического открытия браузера |
| `npm run serve` | обслуживает уже существующий `dist/` без сборки |
| `npm run build` | production-сборка; метаданные функций генерируются автоматически |
| `npm run package:modern` | production-сборка и ZIP современной надстройки |
| `npm run build:dev` | сборка с URL `https://localhost:3000` |
| `npm run generate` | ручная регенерация `functions.json`, `FUNCTIONS.md`, `APPLETS.md` |
| `npm run check` | полный воспроизводимый контроль проекта |
| `npm run certs` | создаёт локальный сертификат |
| `npm run certs:trust:mac` | создаёт и добавляет сертификат в login keychain macOS |
| `npm run start:mac` | сборка, доверие сертификату, sideload в Excel Mac и запуск сервера |
| `npm run sideload:mac` | копирует dev-манифест в папку `wef` Excel Mac |
| `npm run unsideload:mac` | удаляет dev-манифест из Excel Mac |

Полное описание: [LOCAL_DEVELOPMENT.md](docs/LOCAL_DEVELOPMENT.md).

## 🍎 macOS без PowerShell

Современный Office.js-контур полностью собирается на macOS:

```bash
npm ci
npm run certs:trust:mac
npm run start:mac
```

Для обычной production-сборки:

```bash
npm run check
npm run build
```

XLAM/XLTM — отдельный Windows/VBA-контур. Репозиторий содержит исходники и проверки на любой ОС, но автоматическая компиляция `.bas` в бинарные `.xlam/.xltm` выполняется Excel COM на Windows. Варианты для macOS, ограничения и подходящие библиотеки описаны в [MACOS.md](docs/MACOS.md).

## 🧩 XLAM и XLTM

На Windows с установленным Excel:

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

## 🪟 Переносимый пакет и EXE

```powershell
npm run build
npm run legacy:build-all
npm run installer:bundle
npm run installer:build
```

Последняя команда требует Inno Setup 6. Установка выполняется без прав администратора.

## 📅 Сводное расписание

```mermaid
flowchart LR
    A[Чистая книга или XLTM] --> B[Составить сводное расписание]
    B --> C[Автосоздание служебных листов]
    C --> D[Выбор или импорт листа группы]
    D --> E[Группа вручную / имя листа / выбранная ячейка]
    E --> F[Автоопределение сетки и легенды]
    F --> G{Распознано верно?}
    G -- нет --> H[Ручная правка строк, столбцов и смещений]
    H --> F
    G -- да --> I[Занятия, преподаватели и алиасы]
    I --> J[Потоки и Round-Robin]
    J --> K[Конфликты и итоговые листы]
```

Все сервисные страницы создаются при первом запуске либо восстанавливаются, если пользователь их удалил. Полный набор сценариев, включая ошибочный парсинг, одинаковые фамилии, несколько источников, обновление и конфликты: [USE_CASES.md](docs/USE_CASES.md).

## 🧪 Контроль качества

```bash
npm run check
```

Проверяются современный код, 114 функций, 44 апплета, browser preview, автоматическая генерация метаданных, legacy VBA, XLAM/XLTM-сборщики, установщик, оба манифеста, документация, production-сборка и smoke-тест.

Настоящие бинарные XLAM/XLTM создаются на self-hosted Windows runner с установленным Excel. EXE дополнительно требует Inno Setup 6.

## 📚 Документация

- [Индекс](docs/INDEX.md)
- [Локальная разработка](docs/LOCAL_DEVELOPMENT.md)
- [Разработка на macOS](docs/MACOS.md)
- [Сценарии и кейсы](docs/USE_CASES.md)
- [Установка](docs/INSTALLATION.md)
- [Варианты поставки](docs/DISTRIBUTION.md)
- [Руководство пользователя](docs/USER_GUIDE.md)
- [Сводное расписание](docs/SCHEDULE_WORKFLOW.md)
- [Совместимость](docs/COMPATIBILITY.md)
- [Office 2010–2019](docs/LEGACY_OFFICE.md)
- [Архитектура](docs/ARCHITECTURE.md)
- [Тестирование](docs/TESTING.md)
- [Диагностика](docs/TROUBLESHOOTING.md)
- [Функции](docs/FUNCTIONS.md)
- [Апплеты](docs/APPLETS.md)
