import { mkdir, readFile, rm, stat } from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packageJson = JSON.parse(await readFile(path.join(root, "package.json"), "utf8"));
const release = path.join(root, "release");
const dist = path.join(root, "dist");
const archive = path.join(release, `profi-excel-helper-modern-${packageJson.version}.zip`);

await import(`./build.mjs?package=${Date.now()}`);
await mkdir(release, { recursive: true });
await rm(archive, { force: true });
const result = spawnSync("zip", ["-qr", archive, "."], { cwd: dist, stdio: "inherit" });
if (result.error) {
  throw new Error("Команда zip не найдена. На macOS установите Command Line Tools или `brew install zip`.");
}
if (result.status !== 0) throw new Error(`zip завершился с кодом ${result.status}`);
const size = (await stat(archive)).size;
if (size < 4096) throw new Error(`Архив подозрительно мал: ${size} байт`);
console.log(`Modern Office.js package: ${archive} (${size} bytes)`);
