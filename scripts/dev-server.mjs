import https from "node:https";
import { watch } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const distRoot = path.join(projectRoot, "dist");
const args = process.argv.slice(2);
const hasFlag = (name) => args.includes(name);
const option = (name, fallback) => {
  const prefix = `${name}=`;
  const entry = args.find((value) => value.startsWith(prefix));
  return entry ? entry.slice(prefix.length) : fallback;
};
const port = Number(option("--port", process.env.PORT || "3000"));
const host = option("--host", process.env.HOST || "127.0.0.1");
const shouldOpen = hasFlag("--open") && !hasFlag("--no-open") && process.env.CI !== "true";
const shouldWatch = hasFlag("--watch");

const mime = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8"
};

const packageJson = JSON.parse(await readFile(path.join(projectRoot, "package.json"), "utf8"));
const tls = {
  key: await readFile(path.join(projectRoot, ".certs", "key.pem")),
  cert: await readFile(path.join(projectRoot, ".certs", "cert.pem"))
};
const reloadClients = new Set();

async function buildStatus() {
  const metadata = JSON.parse(await readFile(path.join(distRoot, "functions.json"), "utf8"));
  const buildInfo = JSON.parse(await readFile(path.join(distRoot, "build-info.json"), "utf8"));
  return {
    ok: true,
    name: packageJson.name,
    version: packageJson.version,
    mode: "development",
    functions: metadata.functions.length,
    applets: buildInfo.applets,
    builtAt: buildInfo.builtAt,
    urls: {
      dashboard: `https://localhost:${port}/`,
      taskpane: `https://localhost:${port}/src/ui/taskpane.html`,
      manifest: `https://localhost:${port}/manifest.xml`,
      office2016Manifest: `https://localhost:${port}/manifest-office2016.xml`
    }
  };
}

function commonHeaders(contentType) {
  return {
    "Content-Type": contentType,
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Cache-Control": "no-store, max-age=0",
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "no-referrer"
  };
}

async function resolveStaticFile(urlPath) {
  let relative = decodeURIComponent(urlPath).replace(/^\/+/, "");
  if (!relative) relative = "index.html";
  let file = path.resolve(distRoot, relative);
  if (!file.startsWith(`${distRoot}${path.sep}`) && file !== path.join(distRoot, "index.html")) {
    throw new Error("Forbidden");
  }
  if ((await stat(file)).isDirectory()) file = path.join(file, "index.html");
  return file;
}

const server = https.createServer(tls, async (request, response) => {
  try {
    const url = new URL(request.url || "/", `https://${host}:${port}`);
    if (request.method === "OPTIONS") {
      response.writeHead(204, commonHeaders("text/plain; charset=utf-8"));
      response.end();
      return;
    }
    if (url.pathname === "/health" || url.pathname === "/health.json") {
      const body = JSON.stringify(await buildStatus(), null, 2);
      response.writeHead(200, commonHeaders("application/json; charset=utf-8"));
      response.end(body);
      return;
    }
    if (url.pathname === "/events") {
      response.writeHead(200, {
        ...commonHeaders("text/event-stream; charset=utf-8"),
        Connection: "keep-alive"
      });
      response.write("event: ready\ndata: connected\n\n");
      reloadClients.add(response);
      request.on("close", () => reloadClients.delete(response));
      return;
    }
    const file = await resolveStaticFile(url.pathname);
    const body = await readFile(file);
    response.writeHead(200, commonHeaders(mime[path.extname(file).toLowerCase()] || "application/octet-stream"));
    response.end(body);
  } catch (error) {
    const isMissing = error?.code === "ENOENT" || error?.code === "ENOTDIR";
    response.writeHead(isMissing ? 404 : 500, commonHeaders("text/plain; charset=utf-8"));
    response.end(isMissing ? "Not found" : `Development server error: ${error.message}`);
  }
});

function openBrowser(url) {
  const command = process.platform === "darwin" ? "open" : process.platform === "win32" ? "cmd" : "xdg-open";
  const commandArgs = process.platform === "win32" ? ["/c", "start", "", url] : [url];
  const child = spawn(command, commandArgs, { detached: true, stdio: "ignore" });
  child.on("error", () => {});
  child.unref();
}

let rebuilding = false;
let rebuildQueued = false;
let rebuildTimer;
async function rebuild() {
  if (rebuilding) {
    rebuildQueued = true;
    return;
  }
  rebuilding = true;
  const child = spawn(process.execPath, [path.join(projectRoot, "scripts", "build.mjs")], {
    cwd: projectRoot,
    env: { ...process.env, BASE_URL: `https://localhost:${port}` },
    stdio: "inherit"
  });
  const exitCode = await new Promise((resolve) => child.on("exit", resolve));
  rebuilding = false;
  if (exitCode === 0) {
    for (const client of reloadClients) client.write(`event: reload\ndata: ${Date.now()}\n\n`);
    console.log("Изменения собраны. Обновите Excel task pane; браузерная панель перезагрузится автоматически.");
  } else {
    console.error(`Пересборка завершилась с кодом ${exitCode}.`);
  }
  if (rebuildQueued) {
    rebuildQueued = false;
    await rebuild();
  }
}

function scheduleRebuild(filename = "") {
  if (filename.endsWith("functions.json")) return;
  clearTimeout(rebuildTimer);
  rebuildTimer = setTimeout(() => rebuild().catch((error) => console.error(error)), 180);
}

const watchers = [];
if (shouldWatch) {
  for (const directory of [path.join(projectRoot, "src"), path.join(projectRoot, "assets")]) {
    try {
      watchers.push(watch(directory, { recursive: true }, (_event, filename) => scheduleRebuild(String(filename || ""))));
    } catch (error) {
      console.warn(`Наблюдение за ${directory} отключено: ${error.message}`);
    }
  }
  for (const file of [path.join(projectRoot, "manifest.xml"), path.join(projectRoot, "manifest-office2016.xml")]) {
    try {
      watchers.push(watch(file, () => scheduleRebuild(file)));
    } catch (error) {
      console.warn(`Наблюдение за ${file} отключено: ${error.message}`);
    }
  }
}

server.listen(port, host, async () => {
  const status = await buildStatus();
  console.log(`\nПрофиПомощник ${status.version} — локальная разработка`);
  console.log(`Панель запуска:         ${status.urls.dashboard}`);
  console.log(`Task pane preview:      ${status.urls.taskpane}`);
  console.log(`Manifest Microsoft 365: ${status.urls.manifest}`);
  console.log(`Manifest Excel 2016:    ${status.urls.office2016Manifest}`);
  console.log(`Health:                 https://localhost:${port}/health`);
  console.log(`Каталог: ${status.functions} функций, ${status.applets} апплета.`);
  console.log("В браузере доступен интерактивный preview; операции с книгой выполняются только внутри Excel.");
  console.log("Остановка сервера: Ctrl+C.\n");
  if (shouldOpen) openBrowser(status.urls.dashboard);
});

function shutdown() {
  for (const watcher of watchers) watcher.close();
  for (const client of reloadClients) client.end();
  server.close(() => process.exit(0));
  setTimeout(() => process.exit(0), 1500).unref();
}
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
