import { FUNCTION_CATALOG } from "./core/function-catalog.js";
import { evaluateFunction } from "./core/functions-core.js";
import { initializeTaskpane } from "./ui/taskpane.js";
import { registerRibbonCommands } from "./commands.js";

export function registerCustomFunctions() {
  if (!globalThis.CustomFunctions?.associate) return 0;
  for (const definition of FUNCTION_CATALOG) CustomFunctions.associate(definition.id, (...args) => evaluateFunction(definition.id, args));
  return FUNCTION_CATALOG.length;
}
registerCustomFunctions(); registerRibbonCommands();
if (globalThis.Office?.onReady) Office.onReady((info) => { if (info.host === Office.HostType.Excel) initializeTaskpane(); });
else if (typeof document !== "undefined") document.addEventListener("DOMContentLoaded", () => initializeTaskpane());
