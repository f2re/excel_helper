# 🧰 ПрофиПомощник для Excel

Надстройка для кадровых служб, рекрутинга, преподавателей, работников и учебных подразделений. Проект объединяет функции в ячейках, массовые апплеты и составитель сводного преподавательского расписания.

## ✨ Возможности

- **114 функций `PROFI`** в современном Office.js-режиме;
- **44 апплета** для ФИО, кадров, рекрутинга, обучения и планирования;
- импорт или выбор групповых листов;
- название группы вручную, из имени листа или из выбранной ячейки;
- автоматическое распознавание сетки и легенды с ручным исправлением каждого параметра;
- преподавательские алиасы, потоковые занятия, детерминированный Round-Robin и контроль конфликтов;
- автоматическое создание и восстановление всех `_PROFI_*` листов;
- сводное расписание, выбранная неделя, нагрузка преподавателей и занятость аудиторий;
- Legacy Office-пакеты для Excel 2010, 2013, 2016 и 2019 Windows.

> Отдельного выпуска Microsoft Office 2012 не существовало. Для компьютеров этого периода поддерживаются варианты Office 2010 и Office 2013.

## 📦 Варианты поставки

| Клиент | Пакет | Возможности |
|---|---|---|
| Microsoft 365 Excel Windows/macOS/Web | `manifest.xml` | полная панель, Custom Functions, апплеты и расписание |
| Excel 2016 Windows с ExcelApi 1.1 | `manifest-office2016.xml` | облегчённая ES5-панель |
| Excel 2010/2013/2016/2019 Windows | `ProfiExcelHelper-Legacy.xlam` | глобальные VBA-UDF, меню и составитель |
| Excel 2010/2013/2016/2019 Windows | `ProfiExcelHelper-Template.xltm` | автономный проект из одного шаблона |
| Windows | `ProfiExcelHelper-Setup-1.2.0.exe` | per-user установка XLAM + XLTM |
| Windows без EXE | `ProfiExcelHelper-Portable-1.2.0.zip` | переносимый пакет и PowerShell-установка |

Подробнее: [варианты поставки](docs/DISTRIBUTION.md), [совместимость](docs/COMPATIBILITY.md), [legacy-режим](docs/LEGACY_OFFICE.md).

## 🚀 Современная надстройка

```bash
git clone https://github.com/f2re/excel_helper.git
cd excel_helper
npm ci
npm run check
npm start
```

После запуска используйте `dist/manifest.xml`. Для облегчённой панели Excel 2016 — `dist/manifest-office2016.xml`.

## 🧩 XLAM и XLTM

На Windows с установленным Excel:

```powershell
npm run legacy:build
npm run legacy:build-template
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

## 📅 Составление сводного расписания

```mermaid
flowchart LR
    A[Чистая книга или XLTM] --> B[Составить сводное расписание]
    B --> C[Автосоздание служебных листов]
    C --> D[Выбор или импорт листа группы]
    D --> E[Группа вручную / имя листа / выбранная ячейка]
    E --> F[Автоопределение сетки и легенды]
    F --> G{Распознано верно?}
    G -- нет --> H[Ручное изменение строк, столбцов и смещений]
    H --> F
    G -- да --> I[Занятия, преподаватели и алиасы]
    I --> J[Потоки и Round-Robin]
    J --> K[Конфликты и итоговые листы]
```

Все сервисные страницы создаются при первом запуске либо восстанавливаются, если пользователь их удалил.

## 🧪 Контроль качества

```bash
npm run check
```

Проверяются современный код, 114 функций, 44 апплета, legacy VBA, XLAM/XLTM-сборщики, per-user установщик, PowerShell-синтаксис, оба манифеста, документация, production-сборка и smoke-тест.

Настоящие бинарные XLAM/XLTM создаются на self-hosted Windows runner с установленным Excel. EXE дополнительно требует Inno Setup 6.

## 📚 Документация

- [Индекс](docs/INDEX.md)
- [Установка](docs/INSTALLATION.md)
- [Варианты поставки](docs/DISTRIBUTION.md)
- [Руководство пользователя](docs/USER_GUIDE.md)
- [Сводное расписание](docs/SCHEDULE_WORKFLOW.md)
- [Совместимость](docs/COMPATIBILITY.md)
- [Office 2010–2019](docs/LEGACY_OFFICE.md)
- [Архитектура](docs/ARCHITECTURE.md)
- [Разработка](docs/DEVELOPMENT.md)
- [Тестирование](docs/TESTING.md)
- [Диагностика](docs/TROUBLESHOOTING.md)
- [Функции](docs/FUNCTIONS.md)
- [Апплеты](docs/APPLETS.md)
