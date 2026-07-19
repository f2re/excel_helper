# 🧑‍💻 Разработка

## 1. Команды

| Команда | Назначение |
|---|---|
| `npm ci` | установка точных зависимостей |
| `npm run typecheck` | проверка TypeScript |
| `npm test` | unit-тесты чистой логики |
| `npm run generate:functions` | обновление metadata функций |
| `npm run generate:docs` | обновление каталогов функций и апплетов |
| `npm run generate:icons` | генерация PNG-иконок |
| `npm run build` | production-сборка |
| `npm run build:dev` | localhost-сборка |
| `npm run validate` | локальная проверка манифеста |
| `npm run validate:office` | официальная сетевая проверка Microsoft |
| `npm run check` | полный обязательный pipeline |
| `npm start` | локальный HTTPS-сервер |
| `npm run start:desktop` | sideload в Excel Desktop |

## 2. Добавление пользовательской функции

1. Добавьте запись в `scripts/function-catalog.mjs`:

```js
{
  id: "MY_FUNCTION",
  name: "МОЯФУНКЦИЯ",
  category: "Данные",
  description: "Описание результата.",
  parameters: [
    { name: "value", displayName: "Значение", description: "Исходная ячейка" }
  ],
  result: { type: "string", description: "Результат" },
  example: "=PROFI.МОЯФУНКЦИЯ(A2)",
  insertArguments: "A2"
}
```

2. Реализуйте экспорт в `src/functions/implementations.ts`.
3. Добавьте ID и функцию в `functionImplementations`.
4. При необходимости вынесите чистый алгоритм в `src/functions/core.ts`.
5. Добавьте тест.
6. Выполните:

```bash
npm run generate:functions
npm run generate:docs
npm run check
```

Тест требует точного совпадения ID каталога и реализаций.

## 3. Добавление апплета

1. Добавьте запись `Axx` в `src/actions/catalog.ts`.
2. Реализуйте действие:
   - массовая операция — `helperActions.ts` или `advancedActions.ts`;
   - шаблон — `templates.ts`;
   - сложный мастер — отдельный модуль.
3. Добавьте action в `REGISTERED_APPLET_ACTIONS`.
4. Добавьте ветку в `executeApplet()`.
5. Обновите тесты и выполните `npm run check`.

## 4. Добавление служебной таблицы

1. Добавьте имя в `TABLES` (`src/workbook/schema.ts`).
2. Добавьте `TableSpec` с уникальным anchor.
3. Добавьте маппер row ↔ domain type.
4. Убедитесь, что операция вызывает `ensureProjectScaffold()`.
5. Если нужны seed-строки, добавьте идемпотентную логику в `seedDefaults()`.
6. Документируйте допустимое ручное редактирование.

## 5. Новый профиль группового расписания

Профиль не требует отдельного кода, если структура укладывается в модель:

- строка недель;
- первый столбец;
- строка месяцев;
- начало сетки;
- фиксированное число строк на день/пару;
- offsets кода, дисциплины и аудитории;
- легенда и три её столбца.

Можно создать профиль через UI и сохранить его вместе с источником. Для встроенного шаблона добавьте seed-профиль в `scaffold.ts`.

## 6. Тестирование

Чистая бизнес-логика не должна зависеть от `Excel.run`. Тестируемые функции размещаются в:

```text
src/core
src/functions/core.ts
src/schedule/parser.ts
src/schedule/teacherResolver.ts
src/schedule/assignmentEngine.ts
```

Office-dependent интеграции проверяются smoke-тестом в Excel:

1. чистая книга;
2. автосоздание служебных листов;
3. импорт файла;
4. ручная корректировка координат;
5. сохранение источника;
6. добавление преподавателей и алиасов;
7. расчёт;
8. восстановление версии;
9. удаление одной служебной таблицы и самовосстановление.

## 7. Правила совместимости

- не используйте API без `Office.context.requirements.isSetSupported`, если есть fallback;
- не оставляйте книгу без видимого листа;
- не удаляйте пользовательские данные при повторном создании шаблона;
- не используйте нестабильные адреса как единственный ID;
- пользовательские индексы показывайте в 1-based виде, внутренние храните 0-based;
- все ошибки должны содержать действие для восстановления.

## 8. Стиль кода

- TypeScript strict;
- чистые функции для расчётов;
- Office API только в orchestration-слое;
- русские сообщения интерфейса, латинские стабильные ID;
- отсутствие скрытых сетевых запросов;
- идемпотентные операции книги.
