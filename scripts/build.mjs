import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FUNCTION_CATALOG } from "../src/core/function-catalog.js";
import { APPLET_CATALOG } from "../src/core/applet-catalog.js";
import { generateArtifacts } from "./generate-functions.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(root, "dist");
const baseUrlArgument = process.argv.find((value) => value.startsWith("--base-url="));
const baseUrl = (baseUrlArgument?.slice("--base-url=".length) || process.env.BASE_URL || "https://f2re.github.io/excel_helper").replace(/\/$/, "");
const packageJson = JSON.parse(await readFile(path.join(root, "package.json"), "utf8"));

await generateArtifacts({ quiet: true });
await rm(dist, { recursive: true, force: true });
await mkdir(path.join(dist, "assets"), { recursive: true });
await cp(path.join(root, "src"), path.join(dist, "src"), { recursive: true });
await cp(path.join(root, "src", "functions.json"), path.join(dist, "functions.json"));
await cp(path.join(root, "src", "dev", "index.html"), path.join(dist, "index.html"));

for (const manifestName of ["manifest.xml", "manifest-office2016.xml"]) {
  const template = await readFile(path.join(root, manifestName), "utf8");
  await writeFile(path.join(dist, manifestName), template.replaceAll("__BASE_URL__", baseUrl));
}

const png = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAANUlEQVR42mNgGAWjYBSMglEwCkbBKBgFQwYGBgaG/0D8////UQwMDAwMDAz/oRgFo2AUjIJRMAoGAQCAuQkQnVyf8QAAAABJRU5ErkJggg==",
  "base64"
);
for (const size of [16, 32, 64, 80, 128]) {
  await writeFile(path.join(dist, "assets", `icon-${size}.png`), png);
}
await cp(path.join(root, "assets", "icon.svg"), path.join(dist, "assets", "icon.svg"));
await writeFile(path.join(dist, ".nojekyll"), "");
await writeFile(
  path.join(dist, "build-info.json"),
  `${JSON.stringify(
    {
      name: packageJson.name,
      version: packageJson.version,
      baseUrl,
      mode: baseUrl.includes("localhost") ? "development" : "production",
      builtAt: new Date().toISOString(),
      functions: FUNCTION_CATALOG.length,
      applets: APPLET_CATALOG.length
    },
    null,
    2
  )}\n`
);

console.log(
  `Built ${packageJson.name}@${packageJson.version} for ${baseUrl}: ${FUNCTION_CATALOG.length} functions, ${APPLET_CATALOG.length} applets.`
);
