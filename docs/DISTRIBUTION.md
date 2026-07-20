# 📦 Варианты поставки

ПрофиПомощник выпускается в четырёх формах. Они используют одну предметную модель, но рассчитаны на разные поколения Excel.

| Пакет | Файл | Назначение |
|---|---|---|
| Современная надстройка | `manifest.xml` | Microsoft 365, Excel Windows/macOS/Web |
| Совместимая панель | `manifest-office2016.xml` | Excel 2016 Windows с `ExcelApi 1.1` |
| Глобальная legacy-надстройка | `ProfiExcelHelper-Legacy.xlam` | Excel 2010/2013/2016/2019 Windows, команды во всех книгах |
| Переносимый шаблон | `ProfiExcelHelper-Template.xltm` | Новая автономная книга без постоянной установки |
| Установщик | `ProfiExcelHelper-Setup-1.2.0.exe` | Per-user установка XLAM, XLTM, ярлыка и файлов Office.js |
| Переносимый архив | `ProfiExcelHelper-Portable-1.2.0.zip` | Установка PowerShell-скриптом либо ручное использование файлов |

## 🧩 XLAM

`XLAM` устанавливается один раз и загружается Excel при запуске. Функции `PROFI_*`, меню, кадровые мастера и составитель расписания доступны во всех пользовательских книгах.

Служебные листы создаются не внутри XLAM, а в активной книге пользователя. Исходные групповые листы не изменяются.

## 📄 XLTM

`XLTM` открывается как шаблон и создаёт новую книгу с макросами. В ней уже есть стартовый лист с кнопками:

1. создать или восстановить проект;
2. подключить текущий групповой лист;
3. составить сводное расписание;
4. открыть преподавателей;
5. показать справку;
6. показать или скрыть служебные страницы.

Код, размещённый внутри созданной книги, корректно работает с `ThisWorkbook`. При первом запуске служебные листы создаются автоматически.

## 🪟 EXE-установщик

Установщик построен на Inno Setup 6 и работает без прав администратора. Он размещает файлы в профиле текущего пользователя:

```text
%APPDATA%\Microsoft\AddIns\ProfiExcelHelper-Legacy.xlam
%APPDATA%\Microsoft\Templates\ProfiExcelHelper\ProfiExcelHelper-Template.xltm
%LOCALAPPDATA%\ProfiExcelHelper\
```

Затем он регистрирует XLAM через Excel COM, создаёт ярлык шаблона и сохраняет `install.json` с обнаруженной версией и разрядностью Office. Удаление снимает регистрацию XLAM и удаляет установленные файлы.

Один установщик подходит для 32- и 64-разрядного Office, потому что legacy-код не использует WinAPI-декларации, `LongPtr` или `PtrSafe`.

## 🔨 Команды сборки

На Windows с установленным Excel:

```powershell
npm run legacy:build
npm run legacy:build-template
npm run legacy:build-all
npm run legacy:verify
```

После `npm run build` и `npm run legacy:build-all`:

```powershell
npm run installer:bundle
npm run installer:build
```

`installer:build` требует Inno Setup 6. Сборка XLAM и XLTM требует включённого параметра «Доверять доступ к объектной модели проектов VBA».

## ✅ Проверка

Обычный GitHub-hosted CI проверяет исходники, тесты, манифесты, ES5-панель, PowerShell-синтаксис, форматы SaveAs и структуру установщика.

Создание настоящих бинарных `.xlam` и `.xltm` выполняется workflow `Build Legacy Office packages` на self-hosted Windows runner с установленным Excel. Этот же runner может собрать EXE, если установлен Inno Setup.

## 🔐 Подпись

Перед корпоративной поставкой следует подписать:

- VBA-проекты XLAM и XLTM;
- EXE-установщик;
- при наличии — MSI/VSTO-компоненты будущих версий.

Установщик снимает Mark of the Web с локально размещённых файлов, но политика макросов организации остаётся приоритетной.
