# ⬇️ Установка и развёртывание

## Выбор варианта

| Сценарий | Файл |
|---|---|
| Microsoft 365, web или macOS | `dist/manifest.xml` |
| Excel 2016 с минимальной web-панелью | `dist/manifest-office2016.xml` |
| Команды во всех книгах Excel 2010–2019 Windows | `ProfiExcelHelper-Legacy.xlam` |
| Один переносимый файл-проект | `ProfiExcelHelper-Template.xltm` |
| Обычная установка «Далее → Готово» | `ProfiExcelHelper-Setup-1.2.0.exe` |
| Установка без EXE | `ProfiExcelHelper-Portable-1.2.0.zip` |

## Современная Office.js-надстройка

```bash
git clone https://github.com/f2re/excel_helper.git
cd excel_helper
npm ci
npm run check
npm start
```

Для production используйте `npm run build` и файлы из `dist/`.

## Ручная установка XLAM

1. Закройте все книги с важными несохранёнными изменениями.
2. Откройте Excel.
3. Перейдите: Файл → Параметры → Надстройки.
4. Внизу выберите «Надстройки Excel» → «Перейти».
5. Нажмите «Обзор» и выберите `ProfiExcelHelper-Legacy.xlam`.
6. Убедитесь, что надстройка отмечена флажком.

## Использование XLTM

Откройте `ProfiExcelHelper-Template.xltm`. Excel создаст новую книгу на основе шаблона. Сохраните рабочий проект как `.xlsm`, чтобы VBA-код сохранился.

## PowerShell-установка

После сборки payload:

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\install.ps1 `
  -PayloadRoot release\payload `
  -Mode Full
```

Доступны режимы `Full`, `AddinOnly`, `TemplateOnly`.

Установка выполняется в профиль текущего пользователя, определяет Excel и его разрядность, регистрирует XLAM, устанавливает XLTM и создаёт ярлык.

## EXE-установщик

Соберите XLAM/XLTM, современный `dist`, затем:

```powershell
npm run installer:bundle
npm run installer:build
```

Для EXE нужен Inno Setup 6. Администраторские права не требуются.

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
- храните журнал установки из `%LOCALAPPDATA%\ProfiExcelHelper\logs`.

## Диагностика

Если XLAM не появился:

1. проверьте, закрыт ли Excel во время установки;
2. откройте список надстроек и включите файл вручную;
3. проверьте `%APPDATA%\Microsoft\AddIns`;
4. запустите `repair.ps1`;
5. изучите `install.json` и последний лог установки.
