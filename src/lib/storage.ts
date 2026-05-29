/// <reference types="office-js" />
import type { BibEntry } from "./bibtex";
const LIB_KEY = "wordreff.library";
const ORDER_KEY = "wordreff.citedOrder";
const CITED_ORDER_KEY = "wordref.citedOrder";

const DOC_NS = "https://wordreff.local/schema/1.0";
const DOC_ROOT = "wordreff-data";

type WordRefDocData = {
  version: string;
  library: Record<string, any>;
  citedOrder: string[];
  updatedAt: string;
};

/* =========================================================
   Fallback storage: OfficeRuntime.storage → localStorage
   ========================================================= */

function getStore() {
  const or = (globalThis as any).OfficeRuntime;
  const storage = or?.storage;

  if (storage) return storage;

  return {
    async getItem(k: string) {
      return localStorage.getItem(k);
    },
    async setItem(k: string, v: string) {
      localStorage.setItem(k, v);
    },
    async removeItem(k: string) {
      localStorage.removeItem(k);
    },
  };
}

function emptyData(): WordRefDocData {
  return {
    version: "1.0",
    library: {},
    citedOrder: [],
    updatedAt: new Date().toISOString(),
  };
}

async function readFallbackData(): Promise<WordRefDocData> {
  const rawLib = await getStore().getItem(LIB_KEY);
  const rawOrder = await getStore().getItem(ORDER_KEY);

  return {
    version: "1.0",
    library: rawLib ? JSON.parse(rawLib) : {},
    citedOrder: rawOrder ? JSON.parse(rawOrder) : [],
    updatedAt: new Date().toISOString(),
  };
}

async function writeFallbackData(data: WordRefDocData): Promise<void> {
  await getStore().setItem(LIB_KEY, JSON.stringify(data.library || {}));
  await getStore().setItem(ORDER_KEY, JSON.stringify(data.citedOrder || []));
}

async function clearFallbackData(): Promise<void> {
  await getStore().removeItem(LIB_KEY);
  await getStore().removeItem(ORDER_KEY);
}

/* =========================================================
   Encoding helpers for Custom XML Part
   ========================================================= */

function encodeBase64Unicode(str: string): string {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  bytes.forEach((b) => (binary += String.fromCharCode(b)));
  return btoa(binary);
}

function decodeBase64Unicode(base64: string): string {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new TextDecoder().decode(bytes);
}

function buildXml(data: WordRefDocData): string {
  const json = JSON.stringify(data);
  const encoded = encodeBase64Unicode(json);

  return `<?xml version="1.0" encoding="UTF-8"?>
<${DOC_ROOT} xmlns="${DOC_NS}">
  <payload>${encoded}</payload>
</${DOC_ROOT}>`;
}

function parseXmlPayload(xml: string): WordRefDocData | null {
  try {
    const dom = new DOMParser().parseFromString(xml, "application/xml");
    const payload = dom.querySelector("payload")?.textContent?.trim();

    if (!payload) return null;

    const json = decodeBase64Unicode(payload);
    const parsed = JSON.parse(json);

    return {
      version: parsed.version || "1.0",
      library: parsed.library || {},
      citedOrder: parsed.citedOrder || [],
      updatedAt: parsed.updatedAt || new Date().toISOString(),
    };
  } catch {
    return null;
  }
}

function canUseWordCustomXml(): boolean {
  return typeof Word !== "undefined" && typeof Word.run === "function";
}

/* =========================================================
   Document-backed storage using Word Custom XML Parts
   ========================================================= */

async function readDocumentData(): Promise<WordRefDocData | null> {
  if (!canUseWordCustomXml()) return null;

  try {
    return await Word.run(async (ctx) => {
      const parts = ctx.document.customXmlParts.getByNamespace(DOC_NS);
      parts.load("items");
      await ctx.sync();

      if (!parts.items.length) return null;

      const xmlResult = parts.items[0].getXml();
      await ctx.sync();

      return parseXmlPayload(xmlResult.value);
    });
  } catch (e) {
    console.warn("[WordReff] Could not read document storage:", e);
    return null;
  }
}

async function writeDocumentData(data: WordRefDocData): Promise<void> {
  if (!canUseWordCustomXml()) return;

  try {
    const payload: WordRefDocData = {
      ...data,
      version: data.version || "1.0",
      library: data.library || {},
      citedOrder: data.citedOrder || [],
      updatedAt: new Date().toISOString(),
    };

    const xml = buildXml(payload);

    await Word.run(async (ctx) => {
      const parts = ctx.document.customXmlParts.getByNamespace(DOC_NS);
      parts.load("items");
      await ctx.sync();

      // Remove old WordReff data parts to avoid duplicates.
      for (const part of parts.items) {
        part.delete();
      }

      ctx.document.customXmlParts.add(xml);
      await ctx.sync();
    });
  } catch (e) {
    console.warn("[WordReff] Could not write document storage:", e);
  }
}

/* =========================================================
   Unified data access
   Priority:
   1. Document Custom XML Part
   2. OfficeRuntime.storage / localStorage fallback
   ========================================================= */

async function readData(): Promise<WordRefDocData> {
  const docData = await readDocumentData();

  if (docData) {
    await writeFallbackData(docData);
    return docData;
  }

  const fallbackData = await readFallbackData();

  // If document has no embedded data yet, seed it from fallback.
  if (
    Object.keys(fallbackData.library || {}).length > 0 ||
    (fallbackData.citedOrder || []).length > 0
  ) {
    await writeDocumentData(fallbackData);
  }

  return fallbackData;
}

async function writeData(data: WordRefDocData): Promise<void> {
  const clean: WordRefDocData = {
    version: "1.0",
    library: data.library || {},
    citedOrder: data.citedOrder || [],
    updatedAt: new Date().toISOString(),
  };

  await writeFallbackData(clean);
  await writeDocumentData(clean);
}

/* =========================================================
   Public API used by taskpane.ts
   Keep these names unchanged.
   ========================================================= */

   function getSetting<T>(key: string, fallback: T): T {

    try {
  
      const value = Office.context.document.settings.get(key);
  
      if (!value) return fallback;
  
      if (typeof value === "string") {
  
        return JSON.parse(value) as T;
  
      }
  
      return value as T;
  
    } catch {
  
      return fallback;
  
    }
  
  }
  
  function setSetting(key: string, value: unknown): Promise<void> {
  
    Office.context.document.settings.set(key, JSON.stringify(value));
  
    return new Promise((resolve, reject) => {
  
      Office.context.document.settings.saveAsync((result) => {
  
        if (result.status === Office.AsyncResultStatus.Succeeded) {
  
          resolve();
  
        } else {
  
          reject(result.error);
  
        }
  
      });
  
    });
  
  }
  
  export async function getLibrary(): Promise<Record<string, BibEntry>> {
  
    return getSetting<Record<string, BibEntry>>(LIB_KEY, {});
  
  }
  
  export async function saveLibrary(lib: Record<string, BibEntry>): Promise<void> {
  
    await setSetting(LIB_KEY, lib);
  
  }
  
  export async function upsertEntry(entry: BibEntry): Promise<void> {
  
    const lib = await getLibrary();
  
    lib[entry.id] = entry;
  
    await saveLibrary(lib);
  
  }
  
  export async function clearLibrary(): Promise<void> {
  
    await saveLibrary({});
  
  }
  
  export async function getCitedOrder(): Promise<string[]> {
  
    return getSetting<string[]>(CITED_ORDER_KEY, []);
  
  }
  
  export async function setCitedOrder(order: string[]): Promise<void> {
  
    await setSetting(CITED_ORDER_KEY, order);
  
  }
  
  export async function markCited(id: string): Promise<void> {
  
    const order = await getCitedOrder();
  
    if (!order.includes(id)) {
  
      order.push(id);
  
      await setCitedOrder(order);
  
    }
  
  }
  
  export async function clearCitedOrder(): Promise<void> {
  
    await setCitedOrder([]);
  
  }

/* Optional: use only if you want to fully wipe WordReff from local fallback too */
export async function clearAllStorage(): Promise<void> {
  await clearFallbackData();

  if (!canUseWordCustomXml()) return;

  try {
    await Word.run(async (ctx) => {
      const parts = ctx.document.customXmlParts.getByNamespace(DOC_NS);
      parts.load("items");
      await ctx.sync();

      for (const part of parts.items) {
        part.delete();
      }

      await ctx.sync();
    });
  } catch (e) {
    console.warn("[WordReff] Could not clear document storage:", e);
  }
}