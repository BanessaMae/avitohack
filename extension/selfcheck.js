const assert = require("assert");
const ExcelJS = require("exceljs");

global.ExcelJS = ExcelJS;
require("./phrases.js");
require("./parse.js");
const { buildWorkbookBuffer } = require("./excel.js");

const payload = {
  mainCount: 3,
  items: [
    { code: "context_4", type: "banner" },
    {
      id: 101,
      type: "item",
      title: "Квартира у метро",
      description: "Светлая квартира.\nРиёлторов не беспокоить.",
      priceDetailed: { fullString: "9 500 000 ₽" },
      normalizedPrice: "250 000 ₽ за м²",
      addressDetailed: { locationName: "Санкт-Петербург" },
      coords: { address_user: "Невский проспект, 1" },
      geo: { geoReferences: [{ content: "Невский проспект", after: " 300 м" }] },
      developmentName: "ЖК «Тест»",
      realtyInfo: { categoryType: "ss", transactionType: "sell" },
      userLogo: { link: "/brand/setl_group?source=search_icon" },
      imagesCount: 4,
      urlPath: "/sankt-peterburg/kvartiry/test_101",
    },
    {
      id: 102,
      type: "item",
      title: "Можно с агентом",
      description: "Собственник рассматривает покупателей с агентом.",
      priceDetailed: { fullString: "12 000 000 ₽" },
      normalizedPrice: "200 000 ₽ за м²",
      addressDetailed: { locationName: "Санкт-Петербург" },
      coords: { address_user: "Лиговский проспект, 10" },
      realtyInfo: { categoryType: "nd", transactionType: "sell" },
      userLogo: { name: "Частное лицо" },
      imagesCount: 8,
      urlPath: "/sankt-peterburg/kvartiry/test_102",
    },
    {
      id: 103,
      type: "item",
      title: "Только свои",
      description: "Продаю сам. Только собственник.",
      priceDetailed: { string: "8 100 000" },
      urlPath: "https://www.avito.ru/sankt-peterburg/kvartiry/test_103",
    },
    {
      id: 101,
      type: "item",
      title: "дубль",
      description: "дубль",
      urlPath: "/dup",
    },
  ],
  catalog: {
    items: [
      {
        id: 104,
        type: "item",
        title: "Уже есть в items",
        description: "без посредников, звоните сами",
        urlPath: "/extra",
      },
    ],
  },
};

const items = extractItems(payload);
assert.strictEqual(items.length, 4, "banner and duplicate id are skipped, catalog item is kept");

const records = items.map((item, index) => ({
  ...toRecord(item, "https://www.avito.ru/sankt-peterburg/kvartiry/prodam"),
  savedAt: 1_700_000_000_000 + index * 1000,
}));

assert.strictEqual(records[0].seller, "setl group");
assert.strictEqual(records[0].metro, "Невский проспект 300 м");
assert.strictEqual(records[0].dealType, "Вторичка, Продажа");
assert.strictEqual(records[0].price, "9 500 000 ₽");
assert.strictEqual(records[1].dealType, "Новостройка, Продажа");
assert.strictEqual(records[1].seller, "Частное лицо");
assert.strictEqual(records[2].url, "https://www.avito.ru/sankt-peterburg/kvartiry/test_103");
assert.deepStrictEqual(findPhrases(records[0].title, records[0].description, PHRASES), [
  "риелторов не беспокоить",
]);
assert.deepStrictEqual(findPhrases("Только свои", "Только собственник.", PHRASES), ["только собственник"]);
assert.deepStrictEqual(findPhrases("Можно с агентом", "Собственник рассматривает покупателей с агентом.", PHRASES), []);

const merged = mergeRecord({ id: "102", savedAt: 10, title: "старое" }, records[1], 99);
assert.strictEqual(merged.savedAt, 10);
assert.strictEqual(merged.title, "Можно с агентом");
assert.strictEqual(merged.updatedAt, 99);

(async () => {
  const buffer = await buildWorkbookBuffer(records);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  assert.deepStrictEqual(
    workbook.worksheets.map((sheet) => sheet.name),
    ["Можно писать", "Не беспокоить"]
  );

  const open = workbook.getWorksheet("Можно писать");
  const closed = workbook.getWorksheet("Не беспокоить");
  assert.strictEqual(open.getRow(1).getCell(1).value, "Ссылка");
  assert.strictEqual(open.rowCount, 2);
  assert.strictEqual(closed.rowCount, 4);
  assert.strictEqual(open.getRow(2).getCell(2).value, "102");
  assert.strictEqual(open.getRow(2).getCell(3).value, "Можно с агентом");
  const link = open.getRow(2).getCell(1);
  assert.strictEqual(link.value.text || link.value, "https://www.avito.ru/sankt-peterburg/kvartiry/test_102");
  assert.ok(link.hyperlink || (link.value && link.value.hyperlink));
  assert.strictEqual(open.views[0].state, "frozen");
  assert.ok(open.autoFilter);
  assert.strictEqual(closed.getRow(2).getCell(14).value, "риелторов не беспокоить");
  assert.strictEqual(closed.getRow(4).getCell(14).value, "без посредников");

  const browserBuffer = await buildWithBrowserBundle(records);
  const browserBook = new ExcelJS.Workbook();
  await browserBook.xlsx.load(browserBuffer);
  assert.strictEqual(browserBook.getWorksheet("Можно писать").rowCount, 2);
  console.log("selfcheck ok");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});

async function buildWithBrowserBundle(records) {
  const fs = require("fs");
  const vm = require("vm");
  const sandbox = {
    console,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    Buffer,
    TextEncoder,
    TextDecoder,
    URL,
    Blob,
    atob,
    btoa,
    navigator: { userAgent: "chrome" },
  };
  sandbox.self = sandbox;
  sandbox.globalThis = sandbox;
  vm.runInNewContext(fs.readFileSync(`${__dirname}/vendor/exceljs.min.js`, "utf8"), sandbox, { timeout: 20000 });
  vm.runInNewContext(fs.readFileSync(`${__dirname}/phrases.js`, "utf8"), sandbox);
  vm.runInNewContext(fs.readFileSync(`${__dirname}/parse.js`, "utf8"), sandbox);
  vm.runInNewContext(fs.readFileSync(`${__dirname}/excel.js`, "utf8"), sandbox);
  return sandbox.buildWorkbookBuffer(records);
}
