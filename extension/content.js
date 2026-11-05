(function () {
  const HOST_ID = "avito-excel-host";
  let lastPayload = null;
  let clearArmed = false;
  let busy = false;
  let ui = null;

  window.addEventListener("message", (event) => {
    if (event.source !== window) return;
    const data = event.data;
    if (!data || data.source !== "avito-excel") return;
    if (data.type === "items" || data.type === "latest") {
      if (data.payload) lastPayload = data.payload;
      renderHint();
    }
  });

  function requestLatest() {
    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        window.removeEventListener("message", onMessage);
        resolve(lastPayload);
      }, 700);

      function onMessage(event) {
        if (event.source !== window) return;
        const data = event.data;
        if (!data || data.source !== "avito-excel" || data.type !== "latest") return;
        clearTimeout(timer);
        window.removeEventListener("message", onMessage);
        if (data.payload) lastPayload = data.payload;
        resolve(lastPayload);
      }

      window.addEventListener("message", onMessage);
      window.postMessage({ source: "avito-excel-ui", type: "request-latest" }, "*");
    });
  }

  function send(message) {
    return chrome.runtime.sendMessage(message);
  }

  function currentItems() {
    return typeof globalThis.extractItems === "function" ? globalThis.extractItems(lastPayload) : [];
  }

  function renderHint() {
    if (!ui) return;
    const count = currentItems().length;
    const totalText = lastPayload && (lastPayload.mainCount || lastPayload.totalCount);
    if (count) {
      ui.hint.textContent = totalText
        ? `В последней выдаче ${count} объявлений из ${totalText}.`
        : `В последней выдаче ${count} объявлений.`;
    } else {
      ui.hint.textContent = "Список ещё не пойман. Обновите выдачу или откройте другую страницу.";
    }
  }

  function setStatus(text, isError) {
    if (!ui) return;
    ui.status.textContent = text || "";
    ui.status.style.color = isError ? "#8c3a3a" : "#3d5248";
  }

  function setBusy(value) {
    busy = value;
    if (!ui) return;
    ui.save.disabled = value;
    ui.download.disabled = value;
    ui.clear.disabled = value;
  }

  async function refreshStats() {
    try {
      const response = await send({ type: "stats" });
      if (!ui) return;
      const total = response && response.ok ? response.total : 0;
      ui.stats.textContent = `В базе: ${total}`;
    } catch (error) {
      if (ui) ui.stats.textContent = "База недоступна. Обновите расширение.";
    }
  }

  async function onSave() {
    if (busy) return;
    setBusy(true);
    setStatus("Читаю объявления…");
    try {
      await requestLatest();
      const items = currentItems();
      renderHint();
      if (!items.length) {
        setStatus("На этой странице нет выдачи. Откройте поиск и дождитесь списка.", true);
        return;
      }
      const records = items.map((item) => globalThis.toRecord(item, location.href));
      const response = await send({ type: "save", records });
      if (!response || !response.ok) {
        setStatus((response && response.error) || "Не удалось сохранить.", true);
        return;
      }
      ui.stats.textContent = `В базе: ${response.total}`;
      setStatus(`Добавлено ${response.added}, обновлено ${response.updated}.`);
    } catch (error) {
      setStatus(error && error.message ? error.message : "Не удалось сохранить.", true);
    } finally {
      setBusy(false);
    }
  }

  async function onDownload() {
    if (busy) return;
    setBusy(true);
    setStatus("Собираю Excel…");
    try {
      const response = await send({ type: "download" });
      if (!response || !response.ok) {
        setStatus((response && response.error) || "Не удалось создать файл.", true);
        return;
      }
      setStatus(`Файл в «Загрузках»: ${response.filename}`);
    } catch (error) {
      setStatus(error && error.message ? error.message : "Не удалось создать файл.", true);
    } finally {
      setBusy(false);
    }
  }

  async function onClear() {
    if (busy) return;
    if (!clearArmed) {
      clearArmed = true;
      setStatus("Нажмите «Очистить» ещё раз, чтобы удалить собранное.");
      setTimeout(() => {
        clearArmed = false;
      }, 4000);
      return;
    }
    clearArmed = false;
    setBusy(true);
    try {
      const response = await send({ type: "clear" });
      if (!response || !response.ok) {
        setStatus((response && response.error) || "Не удалось очистить.", true);
        return;
      }
      ui.stats.textContent = "В базе: 0";
      setStatus("Собранные объявления удалены.");
    } catch (error) {
      setStatus(error && error.message ? error.message : "Не удалось очистить.", true);
    } finally {
      setBusy(false);
    }
  }

  function mount() {
    if (document.getElementById(HOST_ID)) return;
    const parent = document.documentElement;
    if (!parent) return;

    const host = document.createElement("div");
    host.id = HOST_ID;
    host.style.cssText = [
      "position:fixed",
      "top:72px",
      "right:16px",
      "z-index:2147483647",
      "width:280px",
      "font-family:Segoe UI,Arial,sans-serif",
    ].join(";");
    parent.appendChild(host);

    const shadow = host.attachShadow({ mode: "open" });
    shadow.innerHTML = `
      <style>
        .card { background:#fff; color:#243028; border:1px solid #d7e3dc; border-radius:14px; box-shadow:0 10px 30px rgba(28,48,38,.16); padding:14px; }
        .head { display:flex; align-items:center; justify-content:space-between; gap:8px; }
        h1 { font-size:15px; line-height:1.2; margin:0; font-weight:700; }
        button.toggle { border:0; background:#eef5f1; color:#1f7a4d; border-radius:8px; padding:4px 8px; cursor:pointer; font:inherit; }
        .hint, .stats, .status { font-size:12px; line-height:1.35; margin:8px 0 0; }
        .hint { color:#3d5248; }
        .stats { color:#1f7a4d; font-weight:650; }
        .status { min-height:2.6em; color:#3d5248; }
        button.action { width:100%; margin-top:8px; border:0; border-radius:10px; padding:9px 10px; font:inherit; font-weight:650; cursor:pointer; }
        button.action:disabled { opacity:.55; cursor:default; }
        button.primary { background:#1f7a4d; color:#fff; }
        button.secondary { background:#e7f3ec; color:#1f7a4d; }
        button.ghost { background:transparent; color:#8c3a3a; }
        .body.hidden { display:none; }
      </style>
      <div class="card">
        <div class="head">
          <h1>Авито → Excel</h1>
          <button class="toggle" type="button">Свернуть</button>
        </div>
        <div class="body">
          <p class="hint"></p>
          <p class="stats">В базе: …</p>
          <button class="action primary save" type="button">Выгрузить страницу</button>
          <button class="action secondary download" type="button">Скачать Excel</button>
          <button class="action ghost clear" type="button">Очистить</button>
          <p class="status"></p>
        </div>
      </div>
    `;

    const body = shadow.querySelector(".body");
    const toggle = shadow.querySelector(".toggle");
    toggle.addEventListener("click", () => {
      const hidden = body.classList.toggle("hidden");
      toggle.textContent = hidden ? "Открыть" : "Свернуть";
    });

    ui = {
      hint: shadow.querySelector(".hint"),
      stats: shadow.querySelector(".stats"),
      status: shadow.querySelector(".status"),
      save: shadow.querySelector(".save"),
      download: shadow.querySelector(".download"),
      clear: shadow.querySelector(".clear"),
    };
    ui.save.addEventListener("click", onSave);
    ui.download.addEventListener("click", onDownload);
    ui.clear.addEventListener("click", onClear);
    renderHint();
    refreshStats();
    requestLatest();
  }

  function watch() {
    mount();
    const observer = new MutationObserver(() => {
      if (!document.getElementById(HOST_ID)) mount();
    });
    observer.observe(document.documentElement, { childList: true });
  }

  if (document.documentElement) watch();
  else document.addEventListener("DOMContentLoaded", watch, { once: true });
})();
