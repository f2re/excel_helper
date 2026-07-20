import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = (file) => readFile(new URL(`../${file}`, import.meta.url), "utf8");

test("build always regenerates custom-function metadata", async () => {
  const build = await read("scripts/build.mjs");
  assert.match(build, /generateArtifacts/);
  assert.match(build, /functions\.json/);
  assert.match(build, /build-info\.json/);
});

test("npm start has an explicit browser dashboard and watch mode", async () => {
  const pkg = JSON.parse(await read("package.json"));
  assert.match(pkg.scripts.start, /build:dev/);
  assert.match(pkg.scripts.start, /dev-server\.mjs --watch --open/);
  assert.ok(pkg.scripts["start:no-open"]);
  assert.ok(pkg.scripts.serve);
  assert.ok(pkg.scripts["package:modern"]);
});

test("task pane initializes outside Excel and exposes a function playground", async () => {
  const runtime = await read("src/runtime.js");
  const taskpane = await read("src/ui/taskpane.js");
  assert.match(runtime, /DOMContentLoaded/);
  assert.match(runtime, /browser preview remains available/);
  assert.match(taskpane, /Стенд функций/);
  assert.match(taskpane, /evaluateFunction/);
  assert.match(taskpane, /Вставка формулы доступна только внутри Excel/);
});

test("macOS workflow does not require PowerShell", async () => {
  const pkg = JSON.parse(await read("package.json"));
  const sideload = await read("scripts/sideload-macos.mjs");
  const certs = await read("scripts/dev-cert.mjs");
  assert.doesNotMatch(pkg.scripts["start:mac"], /powershell/i);
  assert.match(sideload, /com\.microsoft\.Excel/);
  assert.match(sideload, /Documents.*wef/s);
  assert.match(certs, /security/);
  assert.match(certs, /openssl/);
});

test("development dashboard and documentation are part of the build", async () => {
  const index = await read("src/dev/index.html");
  const validate = await read("scripts/validate.mjs");
  assert.match(index, /Что делает <code>npm start<\/code>/);
  assert.match(index, /Интерактивный preview/);
  for (const doc of ["LOCAL_DEVELOPMENT.md", "MACOS.md", "USE_CASES.md"]) {
    assert.match(validate, new RegExp(doc.replace(".", "\\.")));
  }
});
