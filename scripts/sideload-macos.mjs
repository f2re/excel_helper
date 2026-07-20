import { access, copyFile, mkdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

if (process.platform !== "darwin") {
  throw new Error("Команда sideload:mac предназначена только для macOS.");
}

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const remove = process.argv.includes("--remove");
const skipBuild = process.argv.includes("--skip-build");
const reveal = process.argv.includes("--reveal");
const source = path.join(root, "dist", "manifest.xml");
const wefDirectory = path.join(os.homedir(), "Library", "Containers", "com.microsoft.Excel", "Data", "Documents", "wef");
const target = path.join(wefDirectory, "ProfiExcelHelper-development.xml");

function run(command, args) {
  const result = spawnSync(command, args, { cwd: root, stdio: "inherit", env: process.env });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${command} завершился с кодом ${result.status}`);
}

if (remove) {
  await rm(target, { force: true });
  console.log(`Manifest удалён: ${target}`);
  console.log("Перезапустите Excel. При необходимости очистите кэш Office.");
  process.exit(0);
}

if (!skipBuild) {
  run(process.execPath, [path.join(root, "scripts", "build.mjs"), "--base-url=https://localhost:3000"]);
}
try {
  await access(source);
} catch {
  throw new Error("dist/manifest.xml не найден. Выполните npm run build:dev.");
}
await mkdir(wefDirectory, { recursive: true });
await copyFile(source, target);
console.log(`Manifest установлен для Excel на macOS: ${target}`);
console.log("Перезапустите Excel и откройте Главная → Надстройки → ПрофиПомощник.");
console.log("Локальный сервер должен быть доступен по адресу https://localhost:3000.");
if (reveal) spawnSync("open", ["-R", target], { stdio: "ignore" });
