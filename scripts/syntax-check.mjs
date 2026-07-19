import { readdir } from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), ".."); const roots = [path.join(root, "src"), path.join(root, "scripts")]; const files = [];
async function walk(directory) { for (const entry of await readdir(directory, { withFileTypes: true })) { const fullPath = path.join(directory, entry.name); if (entry.isDirectory()) await walk(fullPath); else if (/\.(?:m?js)$/.test(entry.name)) files.push(fullPath); } }
for (const directory of roots) await walk(directory); files.sort();
for (const file of files) { const result = spawnSync(process.execPath, ["--check", file], { encoding: "utf8" }); if (result.status !== 0) { process.stderr.write(result.stderr || result.stdout || `Syntax check failed: ${file}\n`); process.exit(result.status || 1); } }
console.log(`JavaScript syntax: ${files.length} files checked.`);
