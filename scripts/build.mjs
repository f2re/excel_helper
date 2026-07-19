import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), ".."); const dist = path.join(root, "dist"); const baseUrl = (process.env.BASE_URL || "https://f2re.github.io/excel_helper").replace(/\/$/, "");
await rm(dist, { recursive: true, force: true }); await mkdir(path.join(dist, "assets"), { recursive: true }); await cp(path.join(root, "src"), path.join(dist, "src"), { recursive: true }); await cp(path.join(root, "src", "functions.json"), path.join(dist, "functions.json"));
for (const manifestName of ["manifest.xml", "manifest-office2016.xml"]) { const template = await readFile(path.join(root, manifestName), "utf8"); await writeFile(path.join(dist, manifestName), template.replaceAll("__BASE_URL__", baseUrl)); }
const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAANUlEQVR42mNgGAWjYBSMglEwCkbBKBgFQwYGBgaG/0D8////UQwMDAwMDAz/oRgFo2AUjIJRMAoGAQCAuQkQnVyf8QAAAABJRU5ErkJggg==", "base64"); for (const size of [16,32,64,80,128]) await writeFile(path.join(dist, "assets", `icon-${size}.png`), png);
await cp(path.join(root, "assets", "icon.svg"), path.join(dist, "assets", "icon.svg")); await writeFile(path.join(dist, ".nojekyll"), ""); console.log(`Built dist for ${baseUrl}`);
