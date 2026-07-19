# ⬇️ Установка

## 1. Microsoft 365 и современные Excel

```bash
npm ci
npm run check
npm start
```

Доверьте `.certs/cert.pem`, затем загрузите `dist/manifest.xml` в Excel. Production-файлы должны размещаться на HTTPS-хостинге; GitHub Pages workflow публикует содержимое `dist/`.

## 2. Excel 2016: облегчённая панель

Соберите проект и загрузите `dist/manifest-office2016.xml`. Этот манифест требует только `ExcelApi 1.1`, не использует Shared Runtime и Custom Functions. Панель написана в ES5 и предоставляет базовые операции.

Для формул и полноценного составителя расписания установите Legacy XLAM.

## 3. Excel 2010, 2013 и 2016 Windows: XLAM

1. Установите настольный Excel.
2. В Центре управления безопасностью временно разрешите доступ к объектной модели VBA-проекта для автоматической сборки.
3. Выполните:

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\build-xlam.ps1
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\install-xlam.ps1
```

4. Перезапустите Excel.
5. Откройте вкладку **Надстройки → ПрофиПомощник**.

После сборки рекомендуется снова отключить программный доступ к VBA-проекту. В организации подпишите XLAM доверенным сертификатом и разверните его через доверенное расположение.

## Ручная сборка XLAM

Если автоматический скрипт запрещён политиками:

1. создайте пустую книгу;
2. откройте редактор VBA (`Alt+F11`);
3. импортируйте все `.bas` из `legacy-vba/src`;
4. сохраните книгу как Excel Add-In (`*.xlam`);
5. подключите её через Файл → Параметры → Надстройки.

## Удаление legacy-надстройки

```powershell
powershell -ExecutionPolicy Bypass -File legacy-vba\scripts\uninstall-xlam.ps1
```

См. также [LEGACY_OFFICE.md](LEGACY_OFFICE.md) и [TROUBLESHOOTING.md](TROUBLESHOOTING.md).
