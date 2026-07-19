import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = (file) => readFile(new URL(`../${file}`, import.meta.url), "utf8");

test("package exposes modern, XLAM, XLTM, installer and signing commands", async () => {
  const pkg = JSON.parse(await read("package.json"));
  assert.equal(pkg.version, "1.2.0");
  for (const name of [
    "legacy:check",
    "legacy:prereq",
    "legacy:build",
    "legacy:build-template",
    "legacy:build-all",
    "legacy:verify",
    "installer:bundle",
    "installer:build",
    "installer:sign",
    "distribution:check"
  ]) {
    assert.ok(pkg.scripts[name], `missing npm script ${name}`);
  }
  assert.equal(pkg.scripts.test, "node --test --test-reporter spec");
});

test("portable template builder uses macro-enabled template format and events", async () => {
  const script = await read("legacy-vba/scripts/build-xltm.ps1");
  assert.match(script, /SaveAs\(\$fullOutput,\s*53\)/);
  assert.match(script, /Initialize-ProfiTemplateSheet/);
  assert.match(script, /Workbook_Open/);
  assert.match(script, /Add-ProfiRibbon/);
  assert.match(script, /Exclude\s+@\('modProfiMenu\.bas'\)/);
});

test("legacy add-in builder uses Open XML add-in format and RibbonX", async () => {
  const script = await read("legacy-vba/scripts/build-xlam.ps1");
  assert.match(script, /SaveAs\(\$fullOutput,\s*55\)/);
  assert.match(script, /IsAddin\s*=\s*\$true/);
  assert.match(script, /Add-ProfiRibbon/);
  assert.match(script, /Exclude\s+@\('modProfiTemplate\.bas'\)/);
});

test("Ribbon callbacks map to VBA procedures", async () => {
  const xml = await read("legacy-vba/ribbon/customUI.xml");
  const vba = await read("legacy-vba/src/modProfiRibbon.bas");
  const callbacks = [...xml.matchAll(/(?:onAction|onLoad)="([^"]+)"/g)].map((match) => match[1]);
  assert.ok(callbacks.length >= 10);
  for (const callback of callbacks) {
    assert.match(vba, new RegExp(`\\bSub\\s+${callback}\\b`, "i"), `missing ${callback}`);
  }
});

test("installer offers selectable components and is per-user", async () => {
  const install = await read("installer/windows/install.ps1");
  const common = await read("installer/windows/ProfiInstaller.Common.psm1");
  const uninstall = await read("installer/windows/uninstall.ps1");
  const iss = await read("installer/windows/ProfiExcelHelper.iss");
  assert.match(common, /LocalApplicationData/);
  assert.match(`${common}\n${install}`, /Register-ProfiExcelAddin/);
  assert.match(install, /LegacyFull/);
  assert.match(install, /ModernOnly/);
  assert.match(install, /backup-/);
  assert.match(uninstall, /Unregister-ProfiExcelAddin/);
  assert.match(iss, /PrivilegesRequired=lowest/);
  assert.match(iss, /\[Types\]/);
  assert.match(iss, /\[Components\]/);
  assert.match(iss, /GetInstallMode/);
  assert.doesNotMatch(iss, /PrivilegesRequired=admin/);
});

test("installer signing uses SHA-256 and timestamping", async () => {
  const signing = await read("installer/windows/sign-installer.ps1");
  assert.match(signing, /signtool\.exe/);
  assert.match(signing, /SHA256/);
  assert.match(signing, /TimestampUrl/);
  assert.match(signing, /verify \/pa/);
});

test("template-hosted VBA can target its own workbook", async () => {
  const common = await read("legacy-vba/src/modProfiCommon.bas");
  assert.match(common, /If Not Application\.ActiveWorkbook\.IsAddin Then/);
  assert.match(common, /Set ProfiHostWorkbook = Application\.ActiveWorkbook/);
  assert.match(common, /If Not ThisWorkbook\.IsAddin Then/);
});

test("portable bundle contains direct install, repair and uninstall launchers", async () => {
  const bundle = await read("installer/windows/build-portable-bundle.ps1");
  for (const token of ["Установить.cmd", "Восстановить.cmd", "Удалить.cmd", "Compress-Archive"]) {
    assert.ok(bundle.includes(token), `portable bundle missing ${token}`);
  }
});
