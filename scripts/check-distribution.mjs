import { access, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const files = [
  "docs/DISTRIBUTION.md",
  "installer/windows/README.md",
  "installer/windows/ProfiExcelHelper.iss",
  "installer/windows/build-portable-bundle.ps1",
  ".github/workflows/legacy-office.yml",
  ".github/workflows/windows-static.yml"
];
for (const file of files) await access(path.join(root, file));
const pkg = JSON.parse(await readFile(path.join(root, "package.json"), "utf8"));
const lock = JSON.parse(await readFile(path.join(root, "package-lock.json"), "utf8"));
if (pkg.version !== "1.2.0" || lock.version !== "1.2.0" || lock.packages?.[""]?.version !== "1.2.0") {
  throw new Error("Package versions must all be 1.2.0");
}
const release = await readFile(path.join(root, ".github/workflows/release.yml"), "utf8");
for (const token of ["legacy-vba", "installer", "portable-sources"]) {
  if (!release.includes(token)) throw new Error(`Release workflow missing token: ${token}`);
}
const legacyWorkflow = await readFile(path.join(root, ".github/workflows/legacy-office.yml"), "utf8");
for (const token of ["self-hosted", "legacy:build-all", "ProfiExcelHelper-Template.xltm", "ProfiExcelHelper-Legacy.xlam"]) {
  if (!legacyWorkflow.includes(token)) throw new Error(`Legacy workflow missing token: ${token}`);
}
console.log(`Distribution metadata verified for ${pkg.version}`);
