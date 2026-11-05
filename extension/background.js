importScripts("phrases.js", "parse.js", "storage.js", "vendor/exceljs.min.js", "excel.js");

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

async function downloadBuffer(buffer, filename) {
  const blob = new Blob([buffer], { type: XLSX_TYPE });
  const url = URL.createObjectURL(blob);
  try {
    await chrome.downloads.download({ url, filename, saveAs: false });
  } catch (error) {
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    let binary = "";
    const step = 0x4000;
    for (let index = 0; index < bytes.length; index += step) {
      binary += String.fromCharCode.apply(null, bytes.subarray(index, index + step));
    }
    await chrome.downloads.download({
      url: `data:${XLSX_TYPE};base64,${btoa(binary)}`,
      filename,
      saveAs: false,
    });
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
}

function fileStamp(date) {
  const value = date || new Date();
  const pad = (part) => String(part).padStart(2, "0");
  return [
    value.getFullYear(),
    pad(value.getMonth() + 1),
    pad(value.getDate()),
  ].join("-") + `-${pad(value.getHours())}${pad(value.getMinutes())}`;
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || !message.type) return;

  const task = (async () => {
    if (message.type === "save") {
      const records = (message.records || []).filter((record) => record && record.id != null);
      if (!records.length) return { ok: false, error: "На этой странице нет объявлений." };
      const stats = await saveRecords(records);
      return { ok: true, ...stats };
    }
    if (message.type === "stats") {
      return { ok: true, total: await countItems() };
    }
    if (message.type === "clear") {
      await clearItems();
      return { ok: true, total: 0 };
    }
    if (message.type === "download") {
      const records = await getAllItems();
      if (!records.length) {
        return { ok: false, error: "База пуста. Сначала выгрузите страницу поиска." };
      }
      const buffer = await buildWorkbookBuffer(records);
      const filename = `avito-${fileStamp()}.xlsx`;
      await downloadBuffer(buffer, filename);
      return { ok: true, filename };
    }
    return { ok: false, error: "Неизвестная команда" };
  })();

  task
    .then(sendResponse)
    .catch((error) => {
      sendResponse({ ok: false, error: error && error.message ? error.message : String(error) });
    });
  return true;
});
