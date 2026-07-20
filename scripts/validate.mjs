import { access, readFile } from "node:fs/promises";
import { FUNCTION_CATALOG } from "../src/core/function-catalog.js";
import { APPLET_CATALOG } from "../src/core/applet-catalog.js";
import { TABLE_SPECS } from "../src/schedule/schema.js";

const required = [
  "manifest.xml",
  "manifest-office2016.xml",
  "package.json",
  "package-lock.json",
  "src/runtime.js",
  "src/ui/taskpane.html",
  "src/ui/taskpane-office2016.html",
  "src/dev/index.html",
  "src/dev/dashboard.js",
  "scripts/dev-server.mjs",
  "scripts/dev-cert.mjs",
  "scripts/sideload-macos.mjs",
  "scripts/package-modern.mjs",
  "src/schedule/excel-service.js",
  "legacy-vba/src/modProfiSchedule.bas",
  "docs/INSTALLATION.md",
  "docs/LOCAL_DEVELOPMENT.md",
  "docs/MACOS.md",
  "docs/USE_CASES.md",
  "docs/COMPATIBILITY.md"
];
for (const file of required) await access(new URL(`../${file}`, import.meta.url));
if (FUNCTION_CATALOG.length !== 114) throw new Error("Ожидалось 114 функций");
if (APPLET_CATALOG.length !== 44) throw new Error("Ожидалось 44 апплета");
if (Object.keys(TABLE_SPECS).length < 15) throw new Error("Неполная служебная схема");

const modern = await readFile(new URL("../manifest.xml", import.meta.url), "utf8");
const legacy = await readFile(new URL("../manifest-office2016.xml", import.meta.url), "utf8");
if (!modern.includes("CustomFunctions") || !modern.includes("SharedRuntime")) throw new Error("Modern manifest incomplete");
if (legacy.includes("CustomFunctions") || legacy.includes("SharedRuntime")) throw new Error("Legacy manifest must not use modern extensions");

const packageJson = JSON.parse(await readFile(new URL("../package.json", import.meta.url), "utf8"));
for (const script of ["start", "start:no-open", "start:mac", "sideload:mac", "unsideload:mac", "certs:trust:mac", "package:modern"]) {
  if (!packageJson.scripts[script]) throw new Error(`Missing development script: ${script}`);
}
console.log(`Validation passed: ${FUNCTION_CATALOG.length} functions, ${APPLET_CATALOG.length} applets, ${Object.keys(TABLE_SPECS).length} tables, cross-platform dev workflow.`);
