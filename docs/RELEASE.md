# 🚢 Выпуск версии

## Общая проверка

```bash
npm ci
npm run check
```

## Современные артефакты

```bash
npm run build
```

Проверьте `dist/manifest.xml`, `dist/manifest-office2016.xml`, `dist/functions.json` и web-файлы панели.

## Legacy Office

На self-hosted Windows с Excel:

```powershell
npm run legacy:build-all
npm run legacy:verify
```

Должны появиться:

```text
legacy-vba/dist/ProfiExcelHelper-Legacy.xlam
legacy-vba/dist/ProfiExcelHelper-Template.xltm
```

## Поставки Windows

```powershell
npm run installer:bundle
npm run installer:build
```

Проверьте переносимый ZIP и EXE на чистом профиле пользователя. Выполните установку, восстановление и удаление.

## Матрица приёмки

- Excel 2010 x86;
- Excel 2013 x86;
- Excel 2016 x86 и x64;
- Excel 2019 x86 и x64;
- Microsoft 365;
- современный Office.js в браузере и на macOS при наличии целевых сред.

## Подпись

До публичной или корпоративной поставки подпишите VBA-проекты и установщик. Не коммитьте закрытые ключи и сертификаты в репозиторий.

## GitHub Release

Создайте тег `v1.2.0`. Linux workflow публикует modern dist, полный source и portable-sources. Self-hosted workflow публикует XLAM, XLTM, переносимый ZIP и, при наличии Inno Setup, EXE.
