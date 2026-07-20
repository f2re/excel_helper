import { access, readFile, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(root, "dist");
const required = [
  "index.html",
  "build-info.json",
  "manifest.xml",
  "manifest-office2016.xml",
  "functions.json",
  "src/runtime.js",
  "src/ui/taskpane.html",
  "src/ui/taskpane.css",
  "src/ui/taskpane.js",
  "src/dev/dashboard.js",
  "src/dev/dashboard.css",
  "src/schedule/excel-service.js",
  "assets/icon-16.png",
  "assets/icon-32.png",
  "assets/icon-80.png",
  ".nojekyll"
];
for (const file of required) await access(path.join(dist, file));

const manifest = await readFile(path.join(dist, "manifest.xml"), "utf8");
if (manifest.includes("__BASE_URL__") || !manifest.includes("/src/ui/taskpane.html")) {
  throw new Error("Modern manifest is not built correctly");
}
const legacyManifest = await readFile(path.join(dist, "manifest-office2016.xml"), "utf8");
if (legacyManifest.includes("__BASE_URL__") || !legacyManifest.includes("taskpane-office2016.html")) {
  throw new Error("Office 2016 manifest is not built correctly");
}
const metadata = JSON.parse(await readFile(path.join(dist, "functions.json"), "utf8"));
if (metadata.functions.length !== 114) throw new Error("В dist отсутствуют метаданные 114 функций");
const buildInfo = JSON.parse(await readFile(path.join(dist, "build-info.json"), "utf8"));
if (buildInfo.functions !== 114 || buildInfo.applets !== 44) throw new Error("build-info.json содержит неверный каталог");
const index = await readFile(path.join(dist, "index.html"), "utf8");
if (!index.includes("Интерактивный preview") || !index.includes("taskpane.html")) throw new Error("Development dashboard is incomplete");

async function walk(directory) {
  const result = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const full = path.join(directory, entry.name);
    if (entry.isDirectory()) result.push(...(await walk(full)));
    else if (entry.name.endsWith(".js")) result.push(full);
  }
  return result;
}
for (const file of await walk(path.join(dist, "src"))) {
  const text = await readFile(file, "utf8");
  for (const match of text.matchAll(/from\s+["'](\.[^"']+)["']/g)) {
    await access(path.resolve(path.dirname(file), match[1]));
  }
}
await import(`${pathToFileURL(path.join(dist, "src/core/functions-core.js")).href}?smoke=${Date.now()}`);
console.log(`dist smoke test passed: ${required.length} required files, ${metadata.functions.length} custom functions, browser dashboard ready.`);
