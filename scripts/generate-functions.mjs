import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FUNCTION_CATALOG } from "../src/core/function-catalog.js";
import { APPLET_CATALOG } from "../src/core/applet-catalog.js";

const scriptPath = fileURLToPath(import.meta.url);
const root = path.resolve(path.dirname(scriptPath), "..");

export async function generateArtifacts({ quiet = false } = {}) {
  const metadata = {
    $schema: "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    allowCustomDataForDataTypeAny: true,
    functions: FUNCTION_CATALOG.map((item) => ({
      id: item.id,
      name: item.name,
      description: item.description,
      helpUrl: "https://github.com/f2re/excel_helper/blob/main/docs/FUNCTIONS.md",
      parameters: item.parameters.map((parameter) => ({
        name: parameter.name,
        description: parameter.description,
        type: "any",
        optional: Boolean(parameter.optional)
      })),
      result: {
        type: "any",
        dimensionality: item.matrix ? "matrix" : "scalar"
      }
    }))
  };

  await mkdir(path.join(root, "src"), { recursive: true });
  await mkdir(path.join(root, "docs"), { recursive: true });
  await writeFile(path.join(root, "src", "functions.json"), `${JSON.stringify(metadata, null, 2)}\n`);

  const categories = new Map();
  for (const item of FUNCTION_CATALOG) {
    const list = categories.get(item.category) || [];
    list.push(item);
    categories.set(item.category, list);
  }

  const categoryNames = {
    hr: "👥 Кадры и ФИО",
    recruiting: "🎯 Рекрутинг",
    education: "🎓 Преподаватели и занятия",
    planning: "✅ Планирование",
    data: "🧹 Данные"
  };

  let functionsDoc = `# ƒ Пользовательские функции\n\nВсего: **${FUNCTION_CATALOG.length}**. Пространство имён Excel: \`PROFI\`.\n\n`;
  for (const [category, items] of categories) {
    functionsDoc += `## ${categoryNames[category] || category}\n\n| Функция | Назначение | Пример |\n|---|---|---|\n`;
    for (const item of items) {
      functionsDoc += `| \`PROFI.${item.name}\` | ${item.description.replace(/\|/g, "\\|")} | \`${item.example.replace(/\|/g, "\\|")}\` |\n`;
    }
    functionsDoc += "\n";
  }
  await writeFile(path.join(root, "docs", "FUNCTIONS.md"), functionsDoc);

  let appletsDoc = `# 🧩 Апплеты\n\nВсего: **${APPLET_CATALOG.length}**. Все апплеты доступны через боковую панель и поиск.\n\n| Код | Апплет | Категория | Действие |\n|---|---|---|---|\n`;
  for (const item of APPLET_CATALOG) {
    appletsDoc += `| ${item.id} | ${item.title} | ${item.category} | ${item.description.replace(/\|/g, "\\|")} |\n`;
  }
  await writeFile(path.join(root, "docs", "APPLETS.md"), appletsDoc);

  const result = { functions: FUNCTION_CATALOG.length, applets: APPLET_CATALOG.length };
  if (!quiet) console.log(`Generated ${result.functions} function definitions and ${result.applets} applet descriptions.`);
  return result;
}

const invokedDirectly = process.argv[1] && path.resolve(process.argv[1]) === scriptPath;
if (invokedDirectly) await generateArtifacts();
