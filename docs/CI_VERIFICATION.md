# 🔍 CI-проверка версии 1.1.0

Этот файл запускает воспроизводимую pull-request проверку полного дерева проекта после добавления совместимых контуров:

- современная Office.js-надстройка;
- облегчённая панель ExcelApi 1.1 для Excel 2016;
- Legacy XLAM/VBA для Excel 2010, 2013 и 2016 Windows.

Проверяемая команда:

```bash
npm ci
npm run check
```

Сборка XLAM через COM выполняется отдельно на self-hosted Windows runner с установленным Microsoft Excel.
