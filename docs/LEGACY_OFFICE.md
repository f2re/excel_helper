# 🧩 Excel 2010–2019: XLAM, XLTM и установщик

Отдельного выпуска Microsoft Office 2012 не было. Для компьютеров этого периода используются Office 2010 или Office 2013.

## Поддерживаемые варианты

| Excel | XLAM | XLTM | EXE | Office.js |
|---|---:|---:|---:|---:|
| 2010 Windows | ✅ | ✅ | ✅ | — |
| 2013 Windows | ✅ | ✅ | ✅ | ограниченно |
| 2016 Windows | ✅ | ✅ | ✅ | панель `ExcelApi 1.1` по фактической сборке |
| 2019 Windows | ✅ | ✅ | ✅ | зависит от редакции и канала обновлений |
| Microsoft 365 | ✅ | ✅ | ✅ | полный современный режим |

## XLAM

Глобальная надстройка содержит VBA-UDF, меню и составитель расписания. Сборка:

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\build-xlam.ps1
```

## XLTM

Переносимый шаблон создаёт новую макросодержащую книгу. Сборка:

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\build-xltm.ps1
```

Шаблон не зависит от установленной XLAM. Он не импортирует модуль глобального меню и использует стартовый лист с кнопками, поэтому не конфликтует с одновременно установленной надстройкой.

## Полная сборка и проверка

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\build-office-packages.ps1
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\verify-office-packages.ps1
```

Проверка открывает XLAM и книгу, созданную из XLTM, через настоящий Excel COM и подтверждает наличие модулей, стартового листа и версии шаблона.

## Установка

Ручная XLAM:

```text
Файл → Параметры → Надстройки → Управление: Надстройки Excel → Перейти → Обзор
```

Шаблон можно открыть напрямую или поместить в каталог пользовательских шаблонов.

Автоматическая установка:

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\install.ps1 -PayloadRoot release\payload -Mode Full
```

Удаление:

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\uninstall.ps1
```

## Ограничения

- legacy-контур предназначен для Windows;
- VBA UserForm и лист управления заменяют настоящую dockable task pane;
- сборка бинарных файлов требует установленного Excel;
- для импорта VBA-модулей при сборке нужен доверенный доступ к VBProject;
- макросы должны быть разрешены политикой организации;
- цифровая подпись не включается в открытый исходный код без сертификата владельца.
