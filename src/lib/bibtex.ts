// src/lib/bibtex.ts
// Use require to avoid ESM/CJS interop issues
// eslint-disable-next-line @typescript-eslint/no-var-requires
const bib = require("bibtex-parse-js");

export type BibEntry = {
  id: string;
  type: string;
  fields: Record<string, string>;
  createdAt: string;
  notes?: { where?: string; text?: string }[];
};

export function parseBibtex(raw: string): BibEntry[] {
  const text = (raw || "").trim();
  if (!text) return [];
  const parsed = bib.toJSON(text) || [];
  return parsed.map((it: any) => ({
    id: it.citationKey || it.entryTags?.citationKey || cryptoId(),
    type: it.entryType,
    fields: it.entryTags || {},
    createdAt: new Date().toISOString(),
  }));
}

function cryptoId() {
  return "id_" + Math.random().toString(36).slice(2, 9);
}