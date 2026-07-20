const $ = (id) => document.getElementById(id);
const toast = (message) => {
  const element = $("toast");
  element.textContent = message;
  element.style.display = "block";
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => (element.style.display = "none"), 2800);
};

async function loadStatus() {
  const candidates = [new URL("health", location.href), new URL("build-info.json", location.href)];
  let lastError;
  for (const url of candidates) {
    try {
      const response = await fetch(url, { cache: "no-store" });
      if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
      return await response.json();
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError;
}

function renderStatus(status) {
  $("version").textContent = status.version || "—";
  $("function-count").textContent = status.functions ?? "—";
  $("applet-count").textContent = status.applets ?? "—";
  $("build-time").textContent = status.builtAt ? new Date(status.builtAt).toLocaleTimeString("ru-RU") : "—";
  const badge = $("server-status");
  badge.textContent = status.ok === false ? "ошибка" : status.mode === "production" ? "production" : "сервер работает";
  badge.className = "status ok";
}

for (const button of document.querySelectorAll("[data-copy]")) {
  button.addEventListener("click", async () => {
    const url = new URL(button.dataset.copy, location.href).href;
    await navigator.clipboard.writeText(url);
    toast(`Скопировано: ${url}`);
  });
}

async function main() {
  try {
    renderStatus(await loadStatus());
  } catch (error) {
    const badge = $("server-status");
    badge.textContent = "статус недоступен";
    badge.className = "status error";
    toast(`Не удалось прочитать статус сборки: ${error.message}`);
  }

  if (location.hostname === "localhost" || location.hostname === "127.0.0.1") {
    const events = new EventSource(new URL("events", location.href));
    events.addEventListener("reload", () => location.reload());
    events.onerror = () => {};
  }
}

main();
