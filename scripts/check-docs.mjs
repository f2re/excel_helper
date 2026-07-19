import { access, readFile, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
async function walk(directory, predicate) { const result = []; for (const entry of await readdir(directory, { withFileTypes: true })) { const full = path.join(directory, entry.name); if (entry.isDirectory()) result.push(...await walk(full, predicate)); else if (predicate(full)) result.push(full); } return result; }
const markdown = await walk(root, (file) => file.endsWith(".md") && !file.includes(`${path.sep}node_modules${path.sep}`) && !file.includes(`${path.sep}dist${path.sep}`)); let links = 0;
for (const file of markdown) { const content = await readFile(file, "utf8"); for (const match of content.matchAll(/\[[^\]]*\]\(([^)]+)\)/g)) { const target = match[1].trim().replace(/^<|>$/g, ""); if (!target || /^(?:https?:|mailto:|#)/i.test(target)) continue; const [relative] = target.split("#"); if (!relative) continue; await access(path.resolve(path.dirname(file), decodeURIComponent(relative))); links += 1; } }
console.log(`Documentation: ${markdown.length} Markdown files, ${links} internal links checked.`);
