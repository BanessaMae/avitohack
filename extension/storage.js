(function (root, factory) {
  const api = factory(root);
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  Object.assign(root, api);
})(typeof globalThis !== "undefined" ? globalThis : this, function (root) {
  const DB_NAME = "avito-excel";
  const STORE = "items";

  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: "id" });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  function requestToPromise(request) {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  function transactionDone(tx) {
    return new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error || new Error("Операция с базой прервана"));
    });
  }

  async function getAllItems() {
    const db = await openDb();
    try {
      const tx = db.transaction(STORE, "readonly");
      const rows = await requestToPromise(tx.objectStore(STORE).getAll());
      await transactionDone(tx);
      return rows || [];
    } finally {
      db.close();
    }
  }

  async function countItems() {
    const db = await openDb();
    try {
      const tx = db.transaction(STORE, "readonly");
      const total = await requestToPromise(tx.objectStore(STORE).count());
      await transactionDone(tx);
      return total || 0;
    } finally {
      db.close();
    }
  }

  async function saveRecords(records) {
    const db = await openDb();
    try {
      const now = Date.now();
      const readTx = db.transaction(STORE, "readonly");
      const existing = await requestToPromise(readTx.objectStore(STORE).getAll());
      await transactionDone(readTx);
      const previousById = new Map((existing || []).map((row) => [String(row.id), row]));

      let added = 0;
      let updated = 0;
      const writeTx = db.transaction(STORE, "readwrite");
      const store = writeTx.objectStore(STORE);
      for (const record of records) {
        const previous = previousById.get(String(record.id));
        if (previous) updated += 1;
        else added += 1;
        store.put(root.mergeRecord(previous, record, now));
      }
      await transactionDone(writeTx);
      const total = await countFrom(db);
      return { added, updated, total, pageCount: records.length };
    } finally {
      db.close();
    }
  }

  function countFrom(db) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readonly");
      const request = tx.objectStore(STORE).count();
      request.onsuccess = () => resolve(request.result || 0);
      request.onerror = () => reject(request.error);
    });
  }

  async function clearItems() {
    const db = await openDb();
    try {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).clear();
      await transactionDone(tx);
    } finally {
      db.close();
    }
  }

  return { getAllItems, countItems, saveRecords, clearItems };
});
