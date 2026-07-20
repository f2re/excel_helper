import { access, mkdir, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const certDirectory = path.join(root, ".certs");
const keyPath = path.join(certDirectory, "key.pem");
const certPath = path.join(certDirectory, "cert.pem");
const configPath = path.join(certDirectory, "openssl.cnf");
const trustRequested = process.argv.includes("--trust");
const force = process.argv.includes("--force");

function run(command, args, options = {}) {
  const result = spawnSync(command, args, { stdio: options.capture ? "pipe" : "inherit", encoding: "utf8" });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    const details = [result.stdout, result.stderr].filter(Boolean).join("\n").trim();
    throw new Error(`${command} завершился с кодом ${result.status}${details ? `: ${details}` : ""}`);
  }
  return result;
}

function opensslAvailable() {
  const result = spawnSync("openssl", ["version"], { stdio: "ignore" });
  return !result.error && result.status === 0;
}

async function exists(file) {
  try {
    await access(file);
    return true;
  } catch {
    return false;
  }
}

await mkdir(certDirectory, { recursive: true });
if (!opensslAvailable()) {
  throw new Error(
    "OpenSSL не найден. На macOS установите Command Line Tools или `brew install openssl`; альтернативно используйте `npx office-addin-dev-certs install`."
  );
}

let mustGenerate = force || !(await exists(keyPath)) || !(await exists(certPath));
if (!mustGenerate) {
  const check = spawnSync("openssl", ["x509", "-checkend", "604800", "-noout", "-in", certPath], { stdio: "ignore" });
  mustGenerate = check.status !== 0;
}

if (mustGenerate) {
  const config = `[req]\ndistinguished_name=req_distinguished_name\nx509_extensions=v3_req\nprompt=no\n[req_distinguished_name]\nCN=localhost\n[v3_req]\nsubjectAltName=@alt_names\nkeyUsage=digitalSignature,keyEncipherment\nextendedKeyUsage=serverAuth\n[alt_names]\nDNS.1=localhost\nIP.1=127.0.0.1\n`;
  await writeFile(configPath, config);
  run("openssl", [
    "req",
    "-x509",
    "-newkey",
    "rsa:2048",
    "-nodes",
    "-sha256",
    "-days",
    "825",
    "-keyout",
    keyPath,
    "-out",
    certPath,
    "-config",
    configPath
  ]);
  console.log(`Создан сертификат разработки: ${certPath}`);
} else {
  console.log(`Используется действующий сертификат: ${certPath}`);
}

if (trustRequested) {
  if (process.platform !== "darwin") {
    throw new Error("Автоматическое доверие через этот скрипт реализовано только для macOS. Импортируйте cert.pem вручную.");
  }
  const keychain = path.join(os.homedir(), "Library", "Keychains", "login.keychain-db");
  run("security", ["add-trusted-cert", "-r", "trustRoot", "-k", keychain, certPath]);
  console.log("Сертификат добавлен в login keychain текущего пользователя как доверенный корневой сертификат.");
} else if (process.platform === "darwin") {
  console.log("Для доверия на macOS выполните: npm run certs:trust:mac");
} else {
  console.log("Доверьте .certs/cert.pem в системном хранилище перед загрузкой надстройки в Excel.");
}
