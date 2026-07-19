# ✅ Воспроизводимая проверка

## Linux/macOS CI

```bash
rm -rf node_modules dist
npm ci
npm run check
npm audit --audit-level=high
```

## Windows с Excel

```powershell
npm ci
npm run check
npm run legacy:build
```

После сборки установите XLAM и выполните сценарии из [TESTING.md](TESTING.md).

Автоматическая проверка подтверждает структуру исходников, алгоритмы, манифесты и web-сборку. Она не подтверждает интерактивную работу COM Excel без Windows runner с установленным Office.
