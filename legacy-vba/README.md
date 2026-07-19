# 🧩 ПрофиПомощник Legacy Office

Legacy-контур предназначен для Excel 2010, 2013, 2016 и 2019 под Windows. Отдельного Office 2012 не существовало; для компьютеров этого периода используйте вариант Office 2010/2013.

## Пакеты

| Файл | Назначение |
|---|---|
| `ProfiExcelHelper-Legacy.xlam` | глобальная надстройка для всех книг |
| `ProfiExcelHelper-Template.xltm` | переносимый шаблон новой рабочей книги |

## Сборка

В Excel разрешите доступ к объектной модели проекта VBA, затем выполните:

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\build-office-packages.ps1
```

Отдельные команды:

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\build-xlam.ps1
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\build-xltm.ps1
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\verify-office-packages.ps1
```

## Установка

Старый совместимый способ XLAM:

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\install-xlam.ps1
```

Полная поставка XLAM + XLTM:

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\install.ps1 -PayloadRoot release\payload -Mode Full
```

## Возможности

- кадровые UDF с префиксом `PROFI_`;
- мастер ФИО, дубли и недельный план;
- автоматическая служебная схема;
- выбор группы вручную, из имени листа или из выбранной ячейки;
- редактируемый профиль парсера;
- алиасы преподавателей, потоки, Round-Robin и конфликты;
- сводное расписание и лист контроля;
- стартовая страница XLTM, не требующая постоянной установки.

Перед корпоративной поставкой подпишите VBA-проекты доверенным сертификатом.
