import { openDB, type IDBPDatabase } from "idb";

let dbPromise: Promise<IDBPDatabase> | null = null;

export function database(): Promise<IDBPDatabase> {
  if (!dbPromise) {
    dbPromise = openDB("paut-field", 1, {
      upgrade(db) {
        db.createObjectStore("kv");
      },
    });
  }
  return dbPromise;
}

export async function kvGet<T>(key: string): Promise<T | undefined> {
  const db = await database();
  return db.get("kv", key) as Promise<T | undefined>;
}

export async function kvSet(key: string, value: unknown): Promise<void> {
  const db = await database();
  await db.put("kv", value, key);
}

export async function kvDel(key: string): Promise<void> {
  const db = await database();
  await db.delete("kv", key);
}
