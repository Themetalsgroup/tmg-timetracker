// db.js — tiny IndexedDB wrapper: a key/value cache + an append-only outbox.
// The outbox holds entries saved offline until they're flushed to the workbook.

const DB_NAME = "tmg-tt";
const DB_VERSION = 1;
let _dbp;

function db() {
  if (!_dbp) {
    _dbp = new Promise((resolve, reject) => {
      const r = indexedDB.open(DB_NAME, DB_VERSION);
      r.onupgradeneeded = () => {
        const d = r.result;
        if (!d.objectStoreNames.contains("cache")) d.createObjectStore("cache");
        if (!d.objectStoreNames.contains("outbox")) d.createObjectStore("outbox", { keyPath: "id", autoIncrement: true });
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }
  return _dbp;
}

async function run(store, mode, op) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const req = op(d.transaction(store, mode).objectStore(store));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export const getCache = (key) => run("cache", "readonly", (s) => s.get(key));
export const setCache = (key, val) => run("cache", "readwrite", (s) => s.put(val, key));

export const addOutbox = (rec) => run("outbox", "readwrite", (s) => s.add({ ...rec, createdAt: Date.now() }));
export const allOutbox = () => run("outbox", "readonly", (s) => s.getAll());
export const deleteOutbox = (id) => run("outbox", "readwrite", (s) => s.delete(id));
export const countOutbox = () => run("outbox", "readonly", (s) => s.count());
