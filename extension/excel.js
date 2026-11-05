(function (root, factory) {
  const api = factory(root);
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.buildWorkbookBuffer = api.buildWorkbookBuffer;
})(typeof globalThis !== "undefined" ? globalThis : this, function (root) {
  const COLUMNS = [
    { header: "Ссылка", key: "url", width: 46 },
    { header: "ID", key: "id", width: 16 },
    { header: "Заголовок", key: "title", width: 38 },
    { header: "Цена", key: "price", width: 22 },
    { header: "Цена за м²", key: "pricePerM2", width: 18 },
    { header: "Город", key: "city", width: 22 },
    { header: "Адрес", key: "address", width: 34 },
    { header: "Метро", key: "metro", width: 28 },
    { header: "ЖК", key: "development", width: 28 },
    { header: "Тип сделки", key: "dealType", width: 24 },
    { header: "Продавец", key: "seller", width: 22 },
    { header: "Фото", key: "photos", width: 10 },
    { header: "Описание", key: "description", width: 64 },
    { header: "Найденные фразы", key: "phrases", width: 36 },
    { header: "Дата выгрузки", key: "savedAt", width: 22 },
  ];

  const THEMES = {
    open: {
      header: "FF1F7A4D",
      band: "FFF3FBF6",
      tab: "FF1F7A4D",
    },
    closed: {
      header: "FF8C3A3A",
      band: "FFFDF4F4",
      tab: "FF8C3A3A",
    },
  };

  function formatSavedAt(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    return date.toLocaleString("ru-RU", { hour12: false });
  }

  function addSheet(workbook, title, rows, themeName) {
    const theme = THEMES[themeName];
    const sheet = workbook.addWorksheet(title, {
      views: [{ state: "frozen", ySplit: 1 }],
      properties: { tabColor: { argb: theme.tab } },
    });
    sheet.columns = COLUMNS.map((column) => ({
      header: column.header,
      key: column.key,
      width: column.width,
    }));

    const header = sheet.getRow(1);
    header.height = 24;
    header.font = { name: "Calibri", bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    header.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    header.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: theme.header } };
      cell.border = thinBorder();
    });

    rows.forEach((row, index) => {
      const excelRow = sheet.addRow({
        url: row.url,
        id: row.id,
        title: row.title,
        price: row.price,
        pricePerM2: row.pricePerM2,
        city: row.city,
        address: row.address,
        metro: row.metro,
        development: row.development,
        dealType: row.dealType,
        seller: row.seller,
        photos: row.photos === "" ? null : row.photos,
        description: row.description,
        phrases: row.phrases,
        savedAt: formatSavedAt(row.savedAt),
      });
      excelRow.height = 48;
      excelRow.alignment = { vertical: "top", wrapText: true };
      excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.border = thinBorder();
        cell.font = { name: "Calibri", size: 11, color: { argb: "FF243028" } };
        if (index % 2 === 1) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: theme.band } };
        }
        if (colNumber === 12 && typeof row.photos === "number") {
          cell.alignment = { vertical: "top", horizontal: "center" };
        }
      });

      const linkCell = excelRow.getCell(1);
      if (row.url) {
        linkCell.value = { text: row.url, hyperlink: row.url };
        linkCell.font = { name: "Calibri", size: 11, color: { argb: "FF0563C1" }, underline: true };
      }
      excelRow.getCell(2).numFmt = "@";
    });

    const lastRow = Math.max(1, sheet.rowCount);
    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: lastRow, column: COLUMNS.length },
    };
    sheet.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    return sheet;
  }

  function thinBorder() {
    const edge = { style: "thin", color: { argb: "FFD5DDD6" } };
    return { top: edge, left: edge, bottom: edge, right: edge };
  }

  async function buildWorkbookBuffer(records) {
    const ExcelJS = root.ExcelJS;
    if (!ExcelJS) throw new Error("Библиотека Excel не загрузилась");
    const phrases = root.PHRASES || [];
    const sorted = [...records].sort(
      (a, b) => (a.savedAt || 0) - (b.savedAt || 0) || String(a.id).localeCompare(String(b.id), "ru")
    );
    const open = [];
    const closed = [];
    for (const record of sorted) {
      const found = root.findPhrases(record.title, record.description, phrases);
      const row = { ...record, phrases: found.join(", ") };
      if (found.length) closed.push(row);
      else open.push(row);
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Авито → Excel";
    workbook.created = new Date();
    addSheet(workbook, "Можно писать", open, "open");
    addSheet(workbook, "Не беспокоить", closed, "closed");
    return workbook.xlsx.writeBuffer();
  }

  return { buildWorkbookBuffer };
});
