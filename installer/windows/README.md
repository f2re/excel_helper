# 📦 Windows-установщик

Каталог содержит per-user установщик для Excel 2010, 2013, 2016 и 2019 под Windows.

## Состав

- `install.ps1` — копирует и регистрирует XLAM, устанавливает XLTM и ярлык;
- `uninstall.ps1` — снимает регистрацию и удаляет файлы текущего пользователя;
- `repair.ps1` — повторная установка поверх текущей версии;
- `ProfiExcelHelper.iss` — проект Inno Setup 6;
- `build-portable-bundle.ps1` — переносимый ZIP без EXE;
- `build-installer.ps1` — сборка `ProfiExcelHelper-Setup-1.2.0.exe`.

Установка не требует прав администратора. Файлы размещаются в `%APPDATA%\Microsoft\AddIns`, `%APPDATA%\Microsoft\Templates\ProfiExcelHelper` и `%LOCALAPPDATA%\ProfiExcelHelper`.

Перед установкой Excel должен быть закрыт. Скрипт определяет установленную редакцию и разрядность Office, но один и тот же VBA-пакет используется для x86 и x64, поскольку в проекте нет WinAPI-деклараций.

## Сборка

1. Соберите `XLAM` и `XLTM` на Windows с Excel.
2. Выполните современную web-сборку `npm run build`.
3. Создайте payload:

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\build-portable-bundle.ps1
```

4. При наличии Inno Setup 6:

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\build-installer.ps1
```

Для производственной поставки подпишите EXE и VBA-проекты сертификатом организации.
