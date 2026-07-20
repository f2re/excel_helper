# 📦 Windows-установщик

Каталог содержит per-user установщик для Excel 2010, 2013, 2016 и 2019 под Windows.

## Состав

- `install.ps1` — устанавливает выбранные компоненты и регистрирует XLAM;
- `uninstall.ps1` — снимает регистрацию и удаляет файлы текущего пользователя;
- `repair.ps1` — повторная полная установка поверх текущей версии;
- `ProfiExcelHelper.iss` — проект Inno Setup 6;
- `build-portable-bundle.ps1` — самоустанавливающийся ZIP без EXE;
- `build-installer.ps1` — сборка `ProfiExcelHelper-Setup-1.2.0.exe`;
- `sign-installer.ps1` — SHA-256 подпись EXE и проверка подписи.

## Режимы

| Mode | Компоненты |
|---|---|
| `Full` | XLAM + XLTM + Office.js-манифесты |
| `LegacyFull` | XLAM + XLTM |
| `AddinOnly` | только глобальная XLAM |
| `TemplateOnly` | только переносимый XLTM |
| `ModernOnly` | только Office.js-манифесты |
| `AddinModern` | XLAM + Office.js |
| `TemplateModern` | XLTM + Office.js |

Inno Setup отображает эти варианты как типы и компоненты установки.

Установка не требует прав администратора. Файлы размещаются в `%APPDATA%\Microsoft\AddIns`, `%APPDATA%\Microsoft\Templates\ProfiExcelHelper` и `%LOCALAPPDATA%\ProfiExcelHelper`.

Перед регистрацией XLAM Excel должен быть закрыт. Скрипт определяет установленную версию и разрядность Office. Один VBA-пакет используется для x86 и x64, поскольку проект не содержит WinAPI-деклараций.

## Сборка

1. Проверьте среду:

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\Test-LegacyPrerequisites.ps1 -RequireVbaProjectAccess
```

2. Соберите XLAM и XLTM на Windows с Excel.
3. Выполните `npm run build` для Office.js.
4. Создайте переносимый пакет:

```powershell
npm run installer:bundle
```

5. При наличии Inno Setup 6:

```powershell
npm run installer:build
```

6. При наличии сертификата:

```powershell
powershell -ExecutionPolicy Bypass -File installer\windows\sign-installer.ps1 `
  -InstallerPath release\ProfiExcelHelper-Setup-1.2.0.exe `
  -CertificateThumbprint YOUR_SHA1_THUMBPRINT
```

PowerShell-установка создаёт резервную копию заменяемых файлов и выполняет откат при ошибке. Для производственной поставки подпишите также VBA-проекты XLAM и XLTM сертификатом организации.
