/// <reference types="office-js" />

function getStore() {
  const or = (globalThis as any).OfficeRuntime;
  const storage = or?.storage;
  if (storage) return storage;
  // Fallback for normal browsers
  return {
    async getItem(k: string) { return localStorage.getItem(k); },
    async setItem(k: string, v: string) { localStorage.setItem(k, v); },
    async removeItem(k: string) { localStorage.removeItem(k); },
  };
}

const LIB_KEY = "wordref.library";
const ORDER_KEY = "wordref.citedOrder";

export async function getLibrary(): Promise<Record<string, any>> {
  const raw = await getStore().getItem(LIB_KEY);
  return raw ? JSON.parse(raw) : {};
}
export async function saveLibrary(lib: Record<string, any>) {
  await getStore().setItem(LIB_KEY, JSON.stringify(lib));
}
export async function upsertEntry(entry: any) {
  const lib = await getLibrary();
  lib[entry.id] = { ...(lib[entry.id] || {}), ...entry };
  await saveLibrary(lib);
}
export async function getCitedOrder(): Promise<string[]> {
  const raw = await getStore().getItem(ORDER_KEY);
  return raw ? JSON.parse(raw) : [];
}
export async function setCitedOrder(ids: string[]) {
  await getStore().setItem(ORDER_KEY, JSON.stringify(ids));
}
export async function markCited(id: string) {
  const order = await getCitedOrder();
  if (!order.includes(id)) {
    order.push(id);
    await setCitedOrder(order);
  }
}
export async function clearCitedOrder() {
  await getStore().removeItem(ORDER_KEY);
}
export async function clearLibrary() {
  await getStore().removeItem(LIB_KEY);
}