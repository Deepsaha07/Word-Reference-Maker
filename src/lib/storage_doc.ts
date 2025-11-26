/// <reference types="office-js" />

import type * as Bib from "./bibtex"; // adjust path if needed
type BibEntry = Bib.BibEntry;

const LIB_KEY   = "wordrefLibraryV1";
const ORDER_KEY = "wordrefCitedOrderV1";

type Library = Record<string, BibEntry>;

/* ----------------- helpers ----------------- */

function getDocSetting<T>(key: string, fallback: T): T {
  // Reads are synchronous
  const v = Office.context.document.settings.get(key);
  return (v ?? fallback) as T;
}

function setDocSetting<T>(key: string, val: T): void {
  Office.context.document.settings.set(key, val);
}

/** Office settings saveAsync promisified */
function saveSettings(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(res.error?.message || "settings.saveAsync failed"));
    });
  });
}

/* ----------------- API: library ----------------- */

export async function getLibrary(): Promise<Library> {
  return getDocSetting<Library>(LIB_KEY, {});
}

export async function upsertEntry(e: BibEntry): Promise<void> {
  const lib = getDocSetting<Library>(LIB_KEY, {});
  lib[e.id] = e;
  setDocSetting(LIB_KEY, lib);
  await saveSettings();
}

export async function clearLibrary(): Promise<void> {
  setDocSetting(LIB_KEY, {});
  await saveSettings();
}

/* ----------------- API: cited order ----------------- */

export async function getCitedOrder(): Promise<string[]> {
  return getDocSetting<string[]>(ORDER_KEY, []);
}

export async function setCitedOrder(ids: string[]): Promise<void> {
  // store unique in order (safety)
  const seen = new Set<string>();
  const uniq = ids.filter(id => !seen.has(id) && (seen.add(id), true));
  setDocSetting(ORDER_KEY, uniq);
  await saveSettings();
}

export async function clearCitedOrder(): Promise<void> {
  setDocSetting(ORDER_KEY, []);
  await saveSettings();
}

export async function markCited(id: string): Promise<void> {
  const order = getDocSetting<string[]>(ORDER_KEY, []);
  if (!order.includes(id)) {
    order.push(id);
    setDocSetting(ORDER_KEY, order);
    await saveSettings();
  }
}

/* ----------------- optional: bulk upsert ----------------- */
export async function upsertMany(entries: BibEntry[]): Promise<void> {
  const lib = getDocSetting<Library>(LIB_KEY, {});
  for (const e of entries) lib[e.id] = e;
  setDocSetting(LIB_KEY, lib);
  await saveSettings();
}

/* ----------------- optional: migration hook ----------------- */
/** Call once on startup if you want to migrate from your old/global store. */
export async function migrateFromGlobal(
  readOld: () => Promise<{ lib?: Library; order?: string[] }>,
  clearOld?: () => Promise<void>
): Promise<void> {
  const curLib   = getDocSetting<Library>(LIB_KEY, {});
  const curOrder = getDocSetting<string[]>(ORDER_KEY, []);

  if (Object.keys(curLib).length || curOrder.length) return; // already has per-doc data

  const { lib, order } = await readOld();
  if (lib) setDocSetting(LIB_KEY, lib);
  if (order) setDocSetting(ORDER_KEY, Array.from(new Set(order)));
  if (lib || order) await saveSettings();

  if (clearOld) await clearOld();
}