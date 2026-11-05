(function () {
  if (window.__avitoExcelHooked) return;
  window.__avitoExcelHooked = true;

  let lastPayload = null;

  function isItemsUrl(url) {
    if (!url) return false;
    const value = String(url);
    return value.includes("/web/1/js/items") || value.includes("/web/1/main/items");
  }

  function remember(payload) {
    if (!payload || typeof payload !== "object") return;
    const items = Array.isArray(payload.items)
      ? payload.items
      : payload.catalog && Array.isArray(payload.catalog.items)
        ? payload.catalog.items
        : null;
    if (!items) return;
    lastPayload = payload;
    window.postMessage({ source: "avito-excel", type: "items", payload }, "*");
  }

  function looksLikePayload(node) {
    if (!node || typeof node !== "object") return false;
    const items = Array.isArray(node.items)
      ? node.items
      : node.catalog && Array.isArray(node.catalog.items)
        ? node.catalog.items
        : null;
    if (!items) return false;
    return items.some((item) => item && typeof item === "object" && item.id != null && (item.urlPath || item.title));
  }

  function findPayload(node, depth) {
    if (!node || depth > 7 || typeof node !== "object") return null;
    if (looksLikePayload(node)) return node;
    const values = Array.isArray(node) ? node.slice(0, 20) : Object.values(node).slice(0, 40);
    for (const value of values) {
      if (!value || typeof value !== "object") continue;
      const found = findPayload(value, depth + 1);
      if (found) return found;
    }
    return null;
  }

  function scanPage() {
    const seeds = [window.__staticRouterHydrationData, window.__initialData__, window.__data];
    for (const seed of seeds) {
      let data = seed;
      if (typeof data === "string") {
        try {
          data = JSON.parse(data);
        } catch (error) {
          continue;
        }
      }
      const found = findPayload(data, 0);
      if (found) return found;
    }
    return null;
  }

  const originalFetch = window.fetch;
  if (typeof originalFetch === "function") {
    window.fetch = function () {
      const pending = originalFetch.apply(this, arguments);
      try {
        const input = arguments[0];
        const url = typeof input === "string" ? input : input && input.url;
        if (isItemsUrl(url) && pending && typeof pending.then === "function") {
          pending
            .then((response) => {
              if (!response || typeof response.clone !== "function") return;
              response.clone().json().then(remember).catch(() => {});
            })
            .catch(() => {});
        }
      } catch (error) {}
      return pending;
    };
  }

  const originalOpen = XMLHttpRequest.prototype.open;
  const originalSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (method, url) {
    this.__avitoExcelUrl = url;
    return originalOpen.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function () {
    this.addEventListener("load", function () {
      try {
        if (!isItemsUrl(this.__avitoExcelUrl) || !this.responseText) return;
        remember(JSON.parse(this.responseText));
      } catch (error) {}
    });
    return originalSend.apply(this, arguments);
  };

  window.addEventListener("message", (event) => {
    if (event.source !== window) return;
    const data = event.data;
    if (!data || data.source !== "avito-excel-ui" || data.type !== "request-latest") return;
    if (!lastPayload) {
      const found = scanPage();
      if (found) lastPayload = found;
    }
    window.postMessage({ source: "avito-excel", type: "latest", payload: lastPayload }, "*");
  });
})();
