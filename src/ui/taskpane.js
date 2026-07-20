import { FUNCTION_CATALOG } from "../core/function-catalog.js";
import { APPLET_CATALOG } from "../core/applet-catalog.js";
import { evaluateFunction } from "../core/functions-core.js";
import { runApplet } from "../actions/applets.js";
import {
  addCurrentSheetAsSource,
  composeSchedule,
  ensureWorkbookStructure,
  setServiceSheetsVisible
} from "../schedule/excel-service.js";
import { detectOfficeCapabilities, chooseCompatibilityMode, compatibilityMessage } from "../core/compatibility.js";

const $ = (id) => document.getElementById(id);
const CATEGORY_LABELS = {
  data: "🧹 Данные",
  hr: "👥 Кадры",
  recruiting: "🎯 Рекрутинг",
  education: "🎓 Обучение",
  planning: "✅ Планирование"
};
const DEMO_ARGUMENTS = {
  ФИОКРАТКО: ["Иванов Иван Иванович"],
  ФИОНОРМА: ["  иВАНОВ   иВАН  иВАНОВИЧ "],
  ФИОСРАВНИТЬ: ["Иванов И. И.", "Иванов Иван Иванович"],
  ВОЗРАСТ: ["1990-04-15", "2026-07-20"],
  ТЕЛЕФОННОРМА: ["8 999 123-45-67"],
  EMAILПРОВЕРКА: ["person@example.org"],
  НАВЫКИСОВПАЛИ: ["Excel; Power BI; SQL", "SQL; Python; Excel"],
  АКАДЧАСЫ: [180, 45],
  СТАТУСЗАДАЧИ: ["В работе", "2026-07-30"],
  ПОХОЖЕСТЬ: ["ООО Ромашка", "Ромашка ООО"]
};

let initialized = false;
let runtimeContext = {};
let currentTab = "home";

const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&#39;"
})[character]);

function toast(message, kind = "info") {
  const element = $("toast");
  element.textContent = message;
  element.dataset.kind = kind;
  element.style.display = "block";
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => (element.style.display = "none"), 3800);
}

function hasExcelRuntime() {
  const officeHost = runtimeContext.officeInfo?.host || globalThis.Office?.context?.host;
  const excelHost = globalThis.Office?.HostType?.Excel;
  return typeof globalThis.Excel !== "undefined" && (!officeHost || officeHost === excelHost || officeHost === "Excel");
}

function environmentInfo() {
  const capabilities = detectOfficeCapabilities();
  if (hasExcelRuntime()) {
    const mode = chooseCompatibilityMode(capabilities);
    return { kind: "excel", title: "Excel подключён", description: compatibilityMessage(mode), mode };
  }
  if (runtimeContext.officeError) {
    return { kind: "error", title: "Office.js не инициализирован", description: runtimeContext.officeError.message || String(runtimeContext.officeError) };
  }
  return {
    kind: "preview",
    title: "Браузерный preview",
    description: "Поиск, каталог и стенд функций работают. Изменение книги доступно после загрузки манифеста в Excel.",
    mode: "browser-preview"
  };
}

function activate(name) {
  currentTab = name;
  document.querySelectorAll(".tab").forEach((element) => element.classList.remove("active"));
  document.querySelectorAll("nav button").forEach((element) => element.classList.toggle("active", element.dataset.tab === name));
  const target = $(`tab-${name}`) || $(name);
  if (target) target.classList.add("active");
}

function actionButton(label, action, className = "") {
  return `<button type="button" class="${escapeHtml(className)}" data-run="${escapeHtml(action)}">${escapeHtml(label)}</button>`;
}

function card({ title, body = "", tags = [], actions = [], extra = "" }) {
  const tagHtml = tags.map((tag) => `<span class="tag">${escapeHtml(tag)}</span>`).join("");
  const actionsHtml = actions.length ? `<div class="card-actions">${actions.map((item) => actionButton(item.label, item.action, item.className)).join("")}</div>` : "";
  return `<article class="card"><h3>${escapeHtml(title)}</h3>${tagHtml}<p>${escapeHtml(body)}</p>${extra}${actionsHtml}</article>`;
}

function functionActions(definition) {
  const actions = [{ label: "Скопировать формулу", action: `copy-formula:${definition.name}`, className: "ghost" }];
  if (hasExcelRuntime()) actions.unshift({ label: "Вставить в ячейку", action: `insert-formula:${definition.name}` });
  return actions;
}

function renderEnvironment() {
  const info = environmentInfo();
  const banner = $("environment-banner");
  banner.className = `environment-banner ${info.kind}`;
  $("environment-title").textContent = info.title;
  $("environment-description").textContent = info.description;
  $("ensure-project").disabled = !hasExcelRuntime();
  $("ensure-project").title = hasExcelRuntime() ? "Создать или восстановить служебную структуру" : "Откройте панель внутри Excel";
}

function renderHome() {
  const info = environmentInfo();
  const defaultFunction = FUNCTION_CATALOG.find((item) => item.name === "ФИОКРАТКО") || FUNCTION_CATALOG[0];
  $("tab-home").innerHTML = `
    <article class="card">
      <h2>Быстрый старт</h2>
      <div class="grid">
        ${actionButton("Проверить книгу", "ensure", hasExcelRuntime() ? "" : "ghost")}
        ${actionButton("Сводное расписание", "open-schedule")}
        ${actionButton("Мастер ФИО", "A01", hasExcelRuntime() ? "secondary" : "ghost")}
        ${actionButton("Найти дубли", "A03", hasExcelRuntime() ? "secondary" : "ghost")}
      </div>
      <p class="notice ${info.kind === "preview" ? "preview" : ""}">${escapeHtml(info.description)}</p>
    </article>
    <article class="card">
      <h2>Стенд функций</h2>
      <p>Работает и в обычном браузере. Аргументы задаются JSON-массивом.</p>
      <label>Функция<select id="demo-function">${FUNCTION_CATALOG.map((item) => `<option value="${escapeHtml(item.name)}" ${item.name === defaultFunction.name ? "selected" : ""}>PROFI.${escapeHtml(item.name)}</option>`).join("")}</select></label>
      <label>Аргументы<textarea id="demo-arguments">${escapeHtml(JSON.stringify(DEMO_ARGUMENTS[defaultFunction.name] || ["Пример"], null, 2))}</textarea></label>
      <div class="card-actions">${actionButton("Вычислить", "demo-evaluate")}${actionButton("Скопировать пример", `copy-formula:${defaultFunction.name}`, "ghost")}</div>
      <pre id="demo-result" class="demo-result">Результат появится здесь.</pre>
    </article>
    <article class="card">
      <h2>Каталог</h2>
      <p><strong>${FUNCTION_CATALOG.length}</strong> функций и <strong>${APPLET_CATALOG.length}</strong> апплета. Используйте поиск сверху или вкладки каталога.</p>
      <div class="card-actions">${actionButton("Функции", "open-functions", "ghost")}${actionButton("Апплеты", "open-applets", "ghost")}${actionButton("Справка", "open-help", "ghost")}</div>
    </article>`;
}

function renderApplets() {
  const grouped = new Map();
  for (const item of APPLET_CATALOG) {
    const list = grouped.get(item.category) || [];
    list.push(item);
    grouped.set(item.category, list);
  }
  $("tab-applets").innerHTML = [...grouped.entries()].map(([category, items]) => `
    <h2 class="category-title">${CATEGORY_LABELS[category] || escapeHtml(category)}</h2>
    ${items.map((item) => card({
      title: `${item.id} · ${item.title}`,
      body: item.description,
      tags: [CATEGORY_LABELS[item.category] || item.category],
      actions: [{ label: hasExcelRuntime() ? "Запустить" : "Требуется Excel", action: item.id, className: hasExcelRuntime() ? "" : "ghost" }]
    })).join("")}`).join("");
}

function renderFunctions() {
  const grouped = new Map();
  for (const item of FUNCTION_CATALOG) {
    const list = grouped.get(item.category) || [];
    list.push(item);
    grouped.set(item.category, list);
  }
  $("tab-functions").innerHTML = [...grouped.entries()].map(([category, items]) => `
    <h2 class="category-title">${CATEGORY_LABELS[category] || escapeHtml(category)}</h2>
    ${items.map((item) => card({
      title: `PROFI.${item.name}`,
      body: item.description,
      tags: [item.matrix ? "массив" : "ячейка", `${item.parameters.length} арг.`],
      extra: `<code class="formula">${escapeHtml(item.example)}</code>`,
      actions: functionActions(item)
    })).join("")}`).join("");
}

function renderSchedule() {
  const disabled = hasExcelRuntime() ? "" : "disabled";
  $("tab-schedule").innerHTML = `
    <article class="card">
      <h2>📅 Составитель сводного расписания</h2>
      <p>В чистой книге служебные листы создаются автоматически. После подключения источника все параметры парсера можно исправить вручную.</p>
      <label>Группа<input id="group" placeholder="например, 522"></label>
      <label>Источник названия группы<select id="group-mode"><option value="manual">Ввести вручную</option><option value="sheet">Имя листа</option><option value="cell">Выбранная ячейка</option></select></label>
      <label>Адрес ячейки группы<input id="group-cell" placeholder="A1"></label>
      <div class="grid">
        <button type="button" data-run="source" ${disabled}>Подключить активный лист</button>
        <button type="button" data-run="compose" ${disabled}>Разобрать и составить</button>
        <button type="button" class="secondary" data-run="service" ${disabled}>Показать служебные листы</button>
        <button type="button" class="secondary" data-run="hide" ${disabled}>Скрыть служебные листы</button>
      </div>
      ${hasExcelRuntime() ? "" : `<div class="card-actions">${actionButton("Показать демонстрационный flow", "schedule-demo", "ghost")}</div>`}
      <p class="notice">Если парсер ошибся, измените значения в <code>tblParserProfiles</code>: строки недель, первый столбец, начало сетки, легенду и смещения.</p>
      <div id="schedule-demo-result"></div>
    </article>
    <article class="card">
      <h3>Порядок обработки</h3>
      <ol class="steps"><li>Создание или восстановление `_PROFI_*` листов.</li><li>Подключение листа группы и выбор способа определения группы.</li><li>Автоопределение сетки и легенды с ручным подтверждением.</li><li>Нормализация занятий и сопоставление преподавателей.</li><li>Потоки, Round-Robin, конфликты и итоговые листы.</li></ol>
    </article>`;
}

function renderHelp() {
  $("tab-help").innerHTML = `
    ${card({ title: "Локальная разработка", body: "Что делает npm start, адреса сервера, browser preview и загрузка в Excel.", actions: [{ label: "Открыть", action: "docs:LOCAL_DEVELOPMENT.md" }] })}
    ${card({ title: "macOS", body: "Сборка современной надстройки, доверие сертификату и sideload без PowerShell.", actions: [{ label: "Открыть", action: "docs:MACOS.md" }] })}
    ${card({ title: "Сценарии использования", body: "Кадры, преподаватели, планировщик и все ветки flow расписания.", actions: [{ label: "Открыть", action: "docs:USE_CASES.md" }] })}
    ${card({ title: "Диагностика", body: "Сертификаты, кэш Office, пустая панель, манифест и старые версии Excel.", actions: [{ label: "Открыть", action: "docs:TROUBLESHOOTING.md" }] })}`;
}

function bindStaticControls() {
  document.querySelectorAll("[data-run]").forEach((button) => {
    button.onclick = () => executeAction(button.dataset.run).catch((error) => toast(error.message || String(error), "error"));
  });
  const demoSelect = $("demo-function");
  if (demoSelect) {
    demoSelect.onchange = () => {
      const args = DEMO_ARGUMENTS[demoSelect.value] || ["Пример"];
      $("demo-arguments").value = JSON.stringify(args, null, 2);
      const copyButton = document.querySelector('[data-run^="copy-formula:"]');
      if (copyButton) copyButton.dataset.run = `copy-formula:${demoSelect.value}`;
    };
  }
}

function render() {
  renderEnvironment();
  renderHome();
  renderApplets();
  renderFunctions();
  renderSchedule();
  renderHelp();
  $("catalog-summary").textContent = `${FUNCTION_CATALOG.length} функций · ${APPLET_CATALOG.length} апплета`;
  bindStaticControls();
  activate(currentTab);
}

async function copyText(value) {
  if (navigator.clipboard?.writeText) return navigator.clipboard.writeText(value);
  const area = document.createElement("textarea");
  area.value = value;
  area.style.position = "fixed";
  area.style.opacity = "0";
  document.body.append(area);
  area.select();
  document.execCommand("copy");
  area.remove();
}

function functionDefinition(name) {
  const definition = FUNCTION_CATALOG.find((item) => item.name === name);
  if (!definition) throw new Error(`Функция ${name} не найдена.`);
  return definition;
}

async function insertFormula(name) {
  if (!hasExcelRuntime()) throw new Error("Вставка формулы доступна только внутри Excel.");
  const definition = functionDefinition(name);
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.formulas = [[definition.example]];
    await context.sync();
  });
}

async function executeAction(action) {
  if (!action) return;
  if (action === "open-functions") return activate("functions");
  if (action === "open-applets") return activate("applets");
  if (action === "open-help") return activate("help");
  if (action === "open-schedule") return activate("schedule");
  if (action === "demo-evaluate") {
    const name = $("demo-function").value;
    let args;
    try {
      const parsed = JSON.parse($("demo-arguments").value);
      args = Array.isArray(parsed) ? parsed : [parsed];
    } catch {
      throw new Error("Аргументы должны быть корректным JSON-массивом, например [\"Иванов Иван Иванович\"].");
    }
    const result = evaluateFunction(name, args);
    $("demo-result").textContent = typeof result === "string" ? result : JSON.stringify(result, null, 2);
    return;
  }
  if (action.startsWith("copy-formula:")) {
    const definition = functionDefinition(action.slice("copy-formula:".length));
    await copyText(definition.example);
    toast(`Скопировано: ${definition.example}`);
    return;
  }
  if (action.startsWith("insert-formula:")) {
    await insertFormula(action.slice("insert-formula:".length));
    toast("Формула вставлена в выбранную ячейку.");
    return;
  }
  if (action.startsWith("docs:")) {
    const file = action.slice("docs:".length);
    window.open(`https://github.com/f2re/excel_helper/blob/main/docs/${encodeURIComponent(file)}`, "_blank", "noopener");
    return;
  }
  if (action === "schedule-demo") {
    $("schedule-demo-result").innerHTML = `<p class="notice preview"><strong>Демо:</strong> 3 источника → 126 занятий → 18 преподавателей → 4 предполагаемых потока → 2 конфликта для ручной проверки. Это только визуальный пример; реальные данные читаются из книги Excel.</p>`;
    return;
  }
  if (!hasExcelRuntime()) throw new Error("Эта операция изменяет книгу и доступна только внутри Excel. В браузере используйте поиск и стенд функций.");
  if (action === "ensure") await ensureWorkbookStructure();
  else if (action === "source") await addCurrentSheetAsSource({ group: $("group").value, groupMode: $("group-mode").value, groupCell: $("group-cell").value });
  else if (action === "compose") await composeSchedule({ reparse: true });
  else if (action === "service") await setServiceSheetsVisible(true);
  else if (action === "hide") await setServiceSheetsVisible(false);
  else await runApplet(action);
  toast("Операция выполнена.");
}

function renderSearch(query) {
  const normalized = query.trim().toLocaleLowerCase("ru-RU");
  if (!normalized) {
    activate("home");
    return;
  }
  const entries = [
    ...APPLET_CATALOG.map((item) => ({ title: `${item.id} · ${item.title}`, text: `${item.description} ${item.category}`, action: item.id, kind: "апплет" })),
    ...FUNCTION_CATALOG.map((item) => ({ title: `PROFI.${item.name}`, text: `${item.description} ${item.example} ${item.category}`, action: hasExcelRuntime() ? `insert-formula:${item.name}` : `copy-formula:${item.name}`, kind: "функция" })),
    { title: "Локальная разработка", text: "npm start localhost preview manifest", action: "docs:LOCAL_DEVELOPMENT.md", kind: "справка" },
    { title: "macOS", text: "mac build sideload certificate no powershell", action: "docs:MACOS.md", kind: "справка" },
    { title: "Сводное расписание", text: "группа парсер потоки конфликты", action: "open-schedule", kind: "раздел" }
  ];
  const found = entries.filter((entry) => `${entry.title} ${entry.text}`.toLocaleLowerCase("ru-RU").includes(normalized));
  $("search-results").innerHTML = found.length
    ? found.map((entry) => card({ title: entry.title, body: entry.text, tags: [entry.kind], actions: [{ label: "Открыть", action: entry.action }] })).join("")
    : '<p class="empty">Ничего не найдено. Попробуйте «ФИО», «дубли», «расписание» или «macOS».</p>';
  activate("search-results");
  bindStaticControls();
}

export function initializeTaskpane(context = {}) {
  if (typeof document === "undefined") return;
  runtimeContext = { ...runtimeContext, ...context };
  if (!initialized) {
    initialized = true;
    document.querySelectorAll("nav button").forEach((button) => (button.onclick = () => activate(button.dataset.tab)));
    $("search").oninput = (event) => renderSearch(event.target.value);
    $("clear-search").onclick = () => {
      $("search").value = "";
      activate("home");
    };
    $("ensure-project").onclick = () => executeAction("ensure").catch((error) => toast(error.message || String(error), "error"));
  }
  render();
}
