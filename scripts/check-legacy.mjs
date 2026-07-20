import { access, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const required = [
  "legacy-vba/src/modProfiCommon.bas",
  "legacy-vba/src/modProfiFunctions.bas",
  "legacy-vba/src/modProfiProject.bas",
  "legacy-vba/src/modProfiMenu.bas",
  "legacy-vba/src/modProfiSchedule.bas",
  "legacy-vba/src/modProfiTemplate.bas",
  "legacy-vba/src/modProfiRibbon.bas",
  "legacy-vba/ribbon/customUI.xml",
  "legacy-vba/scripts/Profi.Build.Common.psm1",
  "legacy-vba/scripts/Inject-Ribbon.ps1",
  "legacy-vba/scripts/Test-LegacyPrerequisites.ps1",
  "legacy-vba/scripts/build-xlam.ps1",
  "legacy-vba/scripts/build-xltm.ps1",
  "legacy-vba/scripts/build-office-packages.ps1",
  "legacy-vba/scripts/verify-office-packages.ps1",
  "installer/windows/ProfiInstaller.Common.psm1",
  "installer/windows/install.ps1",
  "installer/windows/uninstall.ps1",
  "installer/windows/repair.ps1",
  "installer/windows/build-portable-bundle.ps1",
  "installer/windows/build-installer.ps1",
  "installer/windows/sign-installer.ps1",
  "installer/windows/ProfiExcelHelper.iss",
  "manifest-office2016.xml",
  "src/ui/taskpane-office2016.html",
  "src/ui/taskpane-office2016.js",
  "docs/DISTRIBUTION.md"
];
for (const file of required) await access(path.join(root, file));

const basFiles = required.filter((file) => file.endsWith(".bas"));
const modules = await Promise.all(basFiles.map((file) => readFile(path.join(root, file), "utf8")));
const source = modules.join("\n");
for (const name of [
  "ProfiEnsureProject",
  "ProfiComposeSchedule",
  "ProfiConfigureParser",
  "ProfiInstallMenu",
  "ProfiTemplateCreateProject",
  "ProfiTemplateComposeSchedule",
  "ProfiRibbonOnLoad",
  "PROFI_FIO_SHORT",
  "PROFI_POSITION_COUNT",
  "PROFI_TASK_STATUS"
]) {
  if (!new RegExp(`\\b(?:Sub|Function)\\s+${name}\\b`, "i").test(source)) {
    throw new Error(`Legacy procedure missing: ${name}`);
  }
}
if (/\b(?:Declare|PtrSafe|LongPtr|LongLong)\b/i.test(source)) {
  throw new Error("Legacy VBA contains architecture-specific WinAPI declarations");
}
if (!/PROFI_VERSION\s+As\s+String\s*=\s*"1\.2\.0"/i.test(source)) {
  throw new Error("Legacy VBA version must be 1.2.0");
}
const common = await readFile(path.join(root, "legacy-vba/src/modProfiCommon.bas"), "utf8");
if (!/ActiveWorkbook\.IsAddin[\s\S]*Set ProfiHostWorkbook = Application\.ActiveWorkbook/i.test(common)) {
  throw new Error("ProfiHostWorkbook must support code hosted inside XLTM/XLSM workbooks");
}

const ribbonXml = await readFile(path.join(root, "legacy-vba/ribbon/customUI.xml"), "utf8");
const ribbonVba = await readFile(path.join(root, "legacy-vba/src/modProfiRibbon.bas"), "utf8");
const callbacks = [...ribbonXml.matchAll(/(?:onAction|onLoad)="([^"]+)"/g)].map((match) => match[1]);
if (callbacks.length < 8) throw new Error("Ribbon must expose the main project commands");
for (const callback of callbacks) {
  if (!new RegExp(`\\bSub\\s+${callback}\\b`, "i").test(ribbonVba)) throw new Error(`Ribbon callback missing: ${callback}`);
}

const xlam = await readFile(path.join(root, "legacy-vba/scripts/build-xlam.ps1"), "utf8");
const xltm = await readFile(path.join(root, "legacy-vba/scripts/build-xltm.ps1"), "utf8");
if (!/SaveAs\(\$fullOutput,\s*55\)/.test(xlam)) throw new Error("XLAM builder must use file format 55");
if (!/SaveAs\(\$fullOutput,\s*53\)/.test(xltm)) throw new Error("XLTM builder must use file format 53");
if (!/Exclude\s+@\('modProfiTemplate\.bas'\)/.test(xlam)) throw new Error("XLAM builder must exclude template-only module");
if (!/Exclude\s+@\('modProfiMenu\.bas'\)/.test(xltm)) throw new Error("XLTM builder must avoid global legacy menu conflicts");
for (const builder of [xlam, xltm]) {
  if (!/Add-ProfiRibbon/.test(builder)) throw new Error("Both legacy packages must contain RibbonX");
}

const install = await readFile(path.join(root, "installer/windows/install.ps1"), "utf8");
const installerCommon = await readFile(path.join(root, "installer/windows/ProfiInstaller.Common.psm1"), "utf8");
const installerSource = `${installerCommon}\n${install}`;
const uninstall = await readFile(path.join(root, "installer/windows/uninstall.ps1"), "utf8");
const iss = await readFile(path.join(root, "installer/windows/ProfiExcelHelper.iss"), "utf8");
const signing = await readFile(path.join(root, "installer/windows/sign-installer.ps1"), "utf8");
for (const token of ["Microsoft\\AddIns", "Microsoft\\Templates", "Register-ProfiExcelAddin", "install.json"]) {
  if (!installerSource.includes(token)) throw new Error(`Installer missing token: ${token}`);
}
if (!uninstall.includes("Unregister-ProfiExcelAddin")) throw new Error("Uninstaller must unregister XLAM");
if (!/PrivilegesRequired=lowest/.test(iss)) throw new Error("Installer must support per-user installation");
if (!/AppVersion=\{#MyAppVersion\}/.test(iss) || !/1\.2\.0/.test(iss)) throw new Error("Inno Setup version mismatch");
for (const token of ["signtool.exe", "SHA256", "TimestampUrl"]) {
  if (!signing.includes(token)) throw new Error(`Signing hook missing token: ${token}`);
}

const legacyJs = await readFile(path.join(root, "src/ui/taskpane-office2016.js"), "utf8");
for (const token of ["=>", "?.", "??", "import ", "export "]) {
  if (legacyJs.includes(token)) throw new Error(`Office 2016 ES5 bundle contains unsupported token: ${token}`);
}
console.log(`Legacy distribution verified: ${required.length} files, ${modules.length} VBA modules, ${callbacks.length} Ribbon callbacks`);
