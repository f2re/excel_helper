import { composeSchedule, ensureWorkbookStructure, setServiceSheetsVisible } from "./schedule/excel-service.js";
import { runApplet } from "./actions/applets.js";
async function withEvent(event, action) { try { await action(); } catch (error) { console.error(error); if (typeof window !== "undefined") window.alert(error instanceof Error ? error.message : String(error)); } finally { event?.completed?.(); } }
export function registerRibbonCommands() {
  if (!globalThis.Office?.actions?.associate) return;
  Office.actions.associate("ensureProject", (event) => withEvent(event, () => ensureWorkbookStructure()));
  Office.actions.associate("composeSchedule", (event) => withEvent(event, () => composeSchedule({ reparse: true })));
  Office.actions.associate("auditSchedule", (event) => withEvent(event, () => runApplet("A24")));
  Office.actions.associate("openServiceSheets", (event) => withEvent(event, () => setServiceSheetsVisible(true)));
  Office.actions.associate("hideServiceSheets", (event) => withEvent(event, () => setServiceSheetsVisible(false)));
  Office.actions.associate("fioWizard", (event) => withEvent(event, () => runApplet("A01")));
  Office.actions.associate("duplicateManager", (event) => withEvent(event, () => runApplet("A03")));
  Office.actions.associate("weeklyPlanner", (event) => withEvent(event, () => runApplet("A31")));
}
