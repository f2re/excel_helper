# 🚢 Выпуск версии

1. обновить `package.json`, `CHANGELOG.md` и версии манифестов;
2. выполнить `npm ci` и `npm run check`;
3. проверить современную панель;
4. проверить Office 2016 task-pane manifest;
5. собрать и подписать Legacy XLAM на Windows;
6. выполнить ручную приёмку по [TESTING.md](TESTING.md);
7. создать тег `vX.Y.Z`;
8. дождаться workflow release и проверить SHA-256 архивов.

В релиз рекомендуется включать:

- source archive;
- production `dist`;
- подписанный `ProfiExcelHelper-Legacy.xlam`;
- `SHA256SUMS.txt`;
- перечень проверенных клиентов Excel.
