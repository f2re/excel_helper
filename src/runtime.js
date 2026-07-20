import { FUNCTION_CATALOG } from "./core/function-catalog.js";
import { evaluateFunction } from "./core/functions-core.js";
import { initializeTaskpane } from "./ui/taskpane.js";
import { registerRibbonCommands } from "./commands.js";

export function registerCustomFunctions() {
  if (!globalThis.CustomFunctions?.associate) return 0;
  for (const definition of FUNCTION_CATALOG) {
    CustomFunctions.associate(definition.id, (...args) => evaluateFunction(definition.id, args));
  }
  return FUNCTION_CATALOG.length;
}

function bootTaskpane(context = {}) {
  if (typeof document === "undefined") return;
  initializeTaskpane(context);
}

registerCustomFunctions();
registerRibbonCommands();

if (typeof document !== "undefined") {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => bootTaskpane({ source: "dom" }), { once: true });
  } else {
    queueMicrotask(() => bootTaskpane({ source: "dom" }));
  }
}

if (globalThis.Office?.onReady) {
  Office.onReady((officeInfo) => bootTaskpane({ source: "office", officeInfo: officeInfo || {} })).catch((officeError) => {
    console.warn("Office.js initialization failed; browser preview remains available.", officeError);
    bootTaskpane({ source: "office", officeError });
  });
}
