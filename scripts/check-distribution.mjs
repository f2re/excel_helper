import { access, readFile, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const files = [
  "docs/DISTRIBUTION.md",
  "installer/windows/README.md",
  "installer/windows/ProfiExcelHelper.iss",
  "installer/windows/build-portable-bundle.ps1",
  "installer/windows/sign-installer.ps1",
  "legacy-vba/ribbon/customUI.xml",
  "legacy-vba/scripts/Test-LegacyPrerequisites.ps1",
  ".github/workflows/legacy-office.yml",
  ".github/workflows/windows-static.yml",
  ".github/workflows/installer-compile.yml"
];
for (const file of files) await access(path.join(root, file));

const pkg = JSON.parse(await readFile(path.join(root, "package.json"), "utf8"));
const lock = JSON.parse(await readFile(path.join(root, "package-lock.json"), "utf8"));
if (pkg.version !== "1.2.0" || lock.version !== "1.2.0" || lock.packages?.[""]?.version !== "1.2.0") {
  throw new Error("Package versions must all be 1.2.0");
}
for (const script of ["legacy:build", "legacy:build-template", "legacy:build-all", "installer:bundle", "installer:build", "installer:sign"]) {
  if (!pkg.scripts[script]) throw new Error(`Package script missing: ${script}`);
}

const modernManifest = await readFile(path.join(root, "manifest.xml"), "utf8");
const office2016Manifest = await readFile(path.join(root, "manifest-office2016.xml"), "utf8");
if (!modernManifest.includes("<Version>1.2.0.0</Version>")) throw new Error("Modern manifest version mismatch");
if (!office2016Manifest.includes("<Version>1.2.0.0</Version>")) throw new Error("Office 2016 manifest version mismatch");

const release = await readFile(path.join(root, ".github/workflows/release.yml"), "utf8");
for (const token of ["legacy-vba", "installer", "portable-sources", "SHA256SUMS.txt"]) {
  if (!release.includes(token)) throw new Error(`Release workflow missing token: ${token}`);
}

const legacyWorkflow = await readFile(path.join(root, ".github/workflows/legacy-office.yml"), "utf8");
for (const token of ["self-hosted", "legacy:build-all", "ProfiExcelHelper-Template.xltm", "ProfiExcelHelper-Legacy.xlam", "installer:bundle"]) {
  if (!legacyWorkflow.includes(token)) throw new Error(`Legacy workflow missing token: ${token}`);
}
const windowsWorkflow = await readFile(path.join(root, ".github/workflows/windows-static.yml"), "utf8");
for (const token of ["windows-latest", "Parse project PowerShell sources", "Distribution audit"]) {
  if (!windowsWorkflow.includes(token)) throw new Error(`Windows workflow missing token: ${token}`);
}
const installerWorkflow = await readFile(path.join(root, ".github/workflows/installer-compile.yml"), "utf8");
for (const token of ["choco install innosetup", "build-installer.ps1", "compile-only payload fixture"]) {
  if (!installerWorkflow.includes(token)) throw new Error(`Installer compile workflow missing token: ${token}`);
}

const installer = await readFile(path.join(root, "installer/windows/ProfiExcelHelper.iss"), "utf8");
for (const token of ["[Types]", "[Components]", "LegacyFull", "ModernOnly", "PrivilegesRequired=lowest"]) {
  if (!installer.includes(token)) throw new Error(`Installer project missing token: ${token}`);
}

let localizedPowerShell = 0;
for (const directory of ["legacy-vba/scripts", "installer/windows"]) {
  const entries = await readdir(path.join(root, directory), { recursive: true, withFileTypes: true });
  for (const entry of entries) {
    if (!entry.isFile() || !/\.ps(?:1|m1)$/i.test(entry.name)) continue;
    const relativeParent = entry.parentPath ? path.relative(path.join(root, directory), entry.parentPath) : "";
    const fullPath = path.join(root, directory, relativeParent, entry.name);
    const bytes = await readFile(fullPath);
    const text = bytes.toString("utf8");
    if (!/[^\x00-\x7F]/.test(text)) continue;
    localizedPowerShell += 1;
    if (!(bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)) {
      throw new Error(`Localized Windows PowerShell source must use UTF-8 BOM: ${path.relative(root, fullPath)}`);
    }
  }
}
console.log(`Distribution metadata verified for ${pkg.version}: ${files.length} required files, ${localizedPowerShell} localized PowerShell files`);
