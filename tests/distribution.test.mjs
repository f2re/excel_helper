import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = (file) => readFile(new URL(`../${file}`, import.meta.url), "utf8");

test("package exposes modern, XLAM, XLTM and installer commands", async () => {
  const pkg = JSON.parse(await read("package.json"));
  assert.equal(pkg.version, "1.2.0");
  for (const name of ["legacy:check", "legacy:build", "legacy:build-template", "legacy:build-all", "distribution:check"]) {
    assert.ok(pkg.scripts[name], `missing npm script ${name}`);
  }
});

test("portable template builder uses macro-enabled template format", async () => {
  const script = await read("legacy-vba/scripts/build-xltm.ps1");
  assert.match(script, /SaveAs\(\$fullOutput,\s*53\)/);
  assert.match(script, /Initialize-ProfiTemplateSheet/);
  assert.match(script, /Workbook_Open/);
});

test("legacy add-in builder uses Open XML add-in format", async () => {
  const script = await read("legacy-vba/scripts/build-xlam.ps1");
  assert.match(script, /SaveAs\(\$fullOutput,\s*55\)/);
  assert.match(script, /IsAddin\s*=\s*\$true/);
});

test("installer is per-user and reversible", async () => {
  const install = await read("installer/windows/install.ps1");
  const common = await read("installer/windows/ProfiInstaller.Common.psm1");
  const uninstall = await read("installer/windows/uninstall.ps1");
  const iss = await read("installer/windows/ProfiExcelHelper.iss");
  assert.match(common, /LocalApplicationData/);
  assert.match(`${common}\n${install}`, /Register-ProfiExcelAddin/);
  assert.match(uninstall, /Unregister-ProfiExcelAddin/);
  assert.match(iss, /PrivilegesRequired=lowest/);
  assert.doesNotMatch(iss, /PrivilegesRequired=admin/);
});

test("template-hosted VBA can target its own workbook", async () => {
  const common = await read("legacy-vba/src/modProfiCommon.bas");
  assert.match(common, /If Not Application\.ActiveWorkbook\.IsAddin Then/);
  assert.match(common, /Set ProfiHostWorkbook = Application\.ActiveWorkbook/);
});
