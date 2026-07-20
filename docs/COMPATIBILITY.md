# 💻 Совместимость

| Платформа | Рекомендуемый пакет | Статус |
|---|---|---|
| Microsoft 365 Excel Windows | `manifest.xml` | полный режим |
| Microsoft 365 Excel macOS/Web | `manifest.xml` | полный Office.js-режим без VBA |
| Excel 2019 Windows | XLAM/XLTM; Office.js по requirement sets | поддерживается legacy-контуром |
| Excel 2016 Windows | `manifest-office2016.xml` + XLAM/XLTM | поддерживается |
| Excel 2013 Windows | XLAM/XLTM | поддерживается |
| Excel 2010 Windows | XLAM/XLTM | поддерживается |
| Office «2012» | использовать матрицу 2010/2013 | отдельной версии не существовало |
| iPad | ограниченно | без legacy VBA |

## Почему несколько пакетов

Современные Custom Functions и Shared Runtime зависят от requirement sets Office.js. Старые и бессрочные редакции Excel не дают одинакового набора API, поэтому гарантированный контур 2010–2019 реализован через VBA в `.xlam` и `.xltm`.

## Разрядность

Legacy VBA не содержит `Declare`, `PtrSafe`, `LongPtr` и других WinAPI-зависимостей. Один бинарный XLAM/XLTM подходит для x86 и x64 Office. Установщик всё равно определяет разрядность и записывает её в журнал диагностики.

## Windows и macOS

XLAM/XLTM ориентированы на Windows, поскольку автоматическая сборка, регистрация и установщик используют Excel COM. На macOS следует применять современный Office.js-манифест либо вручную проверять переносимость конкретных VBA-модулей.
