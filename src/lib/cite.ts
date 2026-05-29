/// <reference types="office-js" />
import type { BibEntry } from "./bibtex";

/* ======================= types ======================= */


/* ======================= helpers ======================= */

function normalizeStyle(style?: string): string {
  return (style || "apa").toLowerCase();
}
function firstAuthorSurname(e: BibEntry): string {
  const a = e.fields.author || e.fields.editor || "Anon";
  return a.split(" and ")[0].split(",")[0].trim();
}
function isNumericStyle(sty: string) {
  sty = normalizeStyle(sty);
  return sty === "ieee" || sty === "numeric" || sty === "vancouver" || sty === "acs";
}

async function ensureSelectionAtEnd(): Promise<void> {
  await Word.run(async (ctx) => {
    const end = ctx.document.body.getRange("End");
    end.select();
    await ctx.sync();
  });
}

function decodeLatex(value: string): string {
  return String(value || "")
    .replace(/\{\\'([A-Za-z])\}/g, (_, c) => {
      const map: Record<string, string> = {
        a: "á", e: "é", i: "í", o: "ó", u: "ú", y: "ý",
        A: "Á", E: "É", I: "Í", O: "Ó", U: "Ú", Y: "Ý",
        c: "ć", C: "Ć",
      };
      return map[c] || c;
    })
    .replace(/\{\\"([A-Za-z])\}/g, (_, c) => {
      const map: Record<string, string> = {
        a: "ä", e: "ë", i: "ï", o: "ö", u: "ü", y: "ÿ",
        A: "Ä", E: "Ë", I: "Ï", O: "Ö", U: "Ü",
      };
      return map[c] || c;
    })
    .replace(/\{\\`([A-Za-z])\}/g, (_, c) => {
      const map: Record<string, string> = {
        a: "à", e: "è", i: "ì", o: "ò", u: "ù",
        A: "À", E: "È", I: "Ì", O: "Ò", U: "Ù",
      };
      return map[c] || c;
    })
    .replace(/\{\\\^([A-Za-z])\}/g, (_, c) => {
      const map: Record<string, string> = {
        a: "â", e: "ê", i: "î", o: "ô", u: "û",
        A: "Â", E: "Ê", I: "Î", O: "Ô", U: "Û",
      };
      return map[c] || c;
    })
    .replace(/\{\\~([A-Za-z])\}/g, (_, c) => {
      const map: Record<string, string> = {
        a: "ã", n: "ñ", o: "õ",
        A: "Ã", N: "Ñ", O: "Õ",
      };
      return map[c] || c;
    })
    .replace(/--/g, "–")
    .replace(/[{}]/g, "");
}

/* ================= in-text formatting ================= */

export function formatInText(e: BibEntry, style: string, index: number): string {
  const sty = normalizeStyle(style);
  const author = firstAuthorSurname(e);
  const year = e.fields.year || "n.d.";

  if (isNumericStyle(sty)) return `[${index}]`;
  // author–year styles
  return `(${author}, ${year})`;
}

/* =============== bibliography formatting =============== */

export function formatBibliographyEntry(e: BibEntry, style: string, index: number): string {
  const sty = normalizeStyle(style);

  const a = decodeLatex(e.fields.author || e.fields.editor || "Anon");
  const y = decodeLatex(e.fields.year || "n.d.");
  const t = decodeLatex(e.fields.title || "");
  const j = decodeLatex(e.fields.journal || e.fields.booktitle || "");
  const v = decodeLatex(e.fields.volume ? `${e.fields.volume}` : "");
  const nRaw = decodeLatex(e.fields.number ? `${e.fields.number}` : "");
  const n = nRaw ? `(${nRaw})` : "";
  const pagesRaw = decodeLatex(e.fields.pages || "");
  const p = pagesRaw ? `, ${pagesRaw}` : "";
  const doi = e.fields.doi ? ` https://doi.org/${decodeLatex(e.fields.doi)}` : "";

  switch (sty) {
    case "ieee":
    case "numeric":
      return `[${index}] ${a}. “${t},” ${j} ${v}${n}${p}, ${y}.${doi}`;

    case "vancouver":
      return `${index}. ${a}. ${t}. ${j} ${y}${v ? `;${v}` : ""}${nRaw ? `(${nRaw})` : ""}${pagesRaw ? `:${pagesRaw}` : ""}.`;

    case "harvard":
      return `${a}, ${y}. ${t}. ${j}${v ? `, ${v}` : ""}${n ? ` ${n}` : ""}${p}.`;

    case "acs":
      return `${a}. ${t}. ${j} ${y}${v ? `, ${v}` : ""}${n ? ` ${n}` : ""}${p}.`;

    case "apa":
    default:
      return `${a} (${y}). ${t}. ${j}${v ? `, ${v}` : ""}${n ? ` ${n}` : ""}${p}.${doi}`;
  }
}

/* ============== heading + bibliography block ============== */

// replaces the old ensureReferencesHeading()
function ensureReferencesHeadingInContext(ctx: Word.RequestContext): Promise<Word.Paragraph> {
  return new Promise(async (resolve) => {
    const body = ctx.document.body;
    const paras = body.paragraphs;
    paras.load("items,text");
    await ctx.sync();

    let heading = paras.items.find(p => p.text.trim().toLowerCase() === "references");
    if (!heading) {
      heading = body.insertParagraph("References", Word.InsertLocation.end);
      heading.styleBuiltIn = Word.BuiltInStyleName.heading1;
      body.insertParagraph("", Word.InsertLocation.end);
      await ctx.sync();
    }
    resolve(heading!);
  });
}

/**
 * Build/replace the bibliography *per entry paragraph*, each wrapped in its own CC:
 *  - Deletes any old WordRef bibliography CCs (legacy block + per-entry) to avoid duplicates.
 *  - Inserts new entries directly after the "References" heading.
 *  - Applies best-effort formatting (font, size, alignment, line spacing).
 */
// cite.ts
export type BibFormatOptions = {
  fontName?: string;
  fontSize?: number;      // pt
  lineSpacing?: number;   // pt (best-effort)
  alignment?: Word.Alignment; // "Left" | "Right" | "Centered" | "Justified"
  color?: string;         // CSS color name or hex like "#333333"
};

// Helper: create/find the single bibliography CC after the “References” heading
async function getOrCreateBibCC(ctx: Word.RequestContext, heading: Word.Paragraph): Promise<Word.ContentControl> {
  const existing = ctx.document.contentControls.getByTag("wordreff-bibliography");
  existing.load("items");
  await ctx.sync();

  if (existing.items.length > 0) return existing.items[0];

  const afterHeading = heading.getRange("End");
  const p = afterHeading.insertParagraph("", Word.InsertLocation.after);
  const cc = p.insertContentControl();
  cc.tag = "wordreff-bibliography";
  cc.title = "WordReff Bibliography";
  return cc;
}

export async function updateBibliography(
  entries: BibEntry[],
  style: string = "apa",
  opts?: BibFormatOptions
): Promise<void> {
  // de-dup + sort for author-year styles
  const seen = new Set<string>();
  let list = entries.filter(e => !seen.has(e.id) && (seen.add(e.id), true));

  if (["apa", "harvard", "mla"].includes((style || "apa").toLowerCase())) {
    list = [...list].sort((a, b) =>
      (a.fields.author || "").localeCompare(b.fields.author || "")
    );
  }

  const isWordWeb =
    Office.context.platform === Office.PlatformType.OfficeOnline;

  await Word.run(async (ctx) => {
    // ensure heading, then one CC
    const heading = await (async () => {
      const body = ctx.document.body;
      const paras = body.paragraphs;

      paras.load("items,text");
      await ctx.sync();

      let h = paras.items.find(
        p => p.text.trim().toLowerCase() === "references"
      );

      if (!h) {
        h = body.insertParagraph("References", Word.InsertLocation.end);
        h.styleBuiltIn = Word.BuiltInStyleName.heading1;
        body.insertParagraph("", Word.InsertLocation.end);
        await ctx.sync();
      }

      return h!;
    })();

    const bibCC = await getOrCreateBibCC(ctx, heading);

    // Build full text once
    const lines = list.map((e, i) =>
      formatBibliographyEntry(e, style, i + 1)
    );

    const bibText = lines.join("\n");

    // Replace the whole CC with fresh text
    bibCC.insertText(bibText, Word.InsertLocation.replace);
    await ctx.sync();

    // Word Web fallback:
    // Avoid paragraph formatting inside content controls because it can throw
    // Sys.ArgumentOutOfRangeException: Parameter name: index.
    if (isWordWeb) {
      const body = ctx.document.body;
    
      const paras = body.paragraphs;
      paras.load("items/text");
      await ctx.sync();
    
      let heading = paras.items.find(
        p => (p.text || "").trim().toLowerCase() === "references"
      );
    
      if (!heading) {
        heading = body.insertParagraph("References", Word.InsertLocation.end);
        body.insertParagraph("", Word.InsertLocation.end);
        await ctx.sync();
      }
    
      const bibText = lines.join("\n");
    
      body.insertParagraph(bibText, Word.InsertLocation.end);
      await ctx.sync();
    
      console.warn("[WordReff] Word Web: bibliography inserted as plain text for compatibility.");
      return;
    }
    // Desktop Word formatting path - unchanged behavior
    try {
      const rng = bibCC.getRange();
      const paras = rng.paragraphs;

      paras.load("items");
      await ctx.sync();

      for (const p of paras.items) {
        if (opts?.fontName) p.font.name = opts.fontName;
        if (opts?.fontSize !== undefined) p.font.size = opts.fontSize;
        if (opts?.color) p.font.color = opts.color;
        if (opts?.alignment !== undefined) p.alignment = opts.alignment;
        if (opts?.lineSpacing !== undefined) p.lineSpacing = opts.lineSpacing;
      }

      await ctx.sync();
    } catch (err) {
      console.error("[WordReff] Bibliography formatting failed:", err);
    }

    await ctx.sync();
  });
}

/* =================== scan & re-render =================== */

/** Return cited IDs in reading order based on content controls tagged "wordref-cite:<id>" */
export async function scanCitationsInDoc(): Promise<string[]> {
  return await Word.run(async (ctx) => {
    const controls = ctx.document.contentControls;
    controls.load("items/tag");
    await ctx.sync();

    const ids: string[] = [];
    for (const cc of controls.items) {
      const tag = (cc.tag || "");
      if (tag.startsWith("wordreff-cite:")) {
        const id = tag.split(":")[1];
        if (id) ids.push(id);
      }
    }
    return ids;
  });
}

export async function rerenderAllCitations(style: string, order: string[], lib: Record<string, BibEntry>) {
  await Word.run(async (ctx) => {
    const controls = ctx.document.contentControls;
    controls.load("items/tag");
    await ctx.sync();

    const items = controls.items.filter(cc => (cc.tag || "").startsWith("wordreff-cite:"));
    for (const cc of items) {
      const id = (cc.tag || "").split(":")[1];
      const e = lib[id];
      if (!e) continue;

      const idx = Math.max(0, order.indexOf(id)) + 1;
      const newText = formatInText(e, style, idx);
      cc.insertText(newText, Word.InsertLocation.replace);
    }
    await ctx.sync();
  });
}

/* ================== grouping for numeric ================== */

function parseBracketIndices(text: string): number[] {
  // Accept "[2]", "[2, 3-5]" / "[2][3]" → we’ll normalize later
  const inside = text.match(/\[(.*?)\]/g);
  const bag: number[] = [];
  if (!inside) return bag;
  for (const block of inside) {
    const s = block.replace(/^\[|\]$/g, "");
    for (const part of s.split(",").map(t => t.trim()).filter(Boolean)) {
      if (/^\d+$/.test(part)) {
        bag.push(parseInt(part, 10));
      } else if (/^\d+\s*-\s*\d+$/.test(part)) {
        const [a, b] = part.split("-").map(x => parseInt(x.trim(), 10));
        for (let i = Math.min(a, b); i <= Math.max(a, b); i++) bag.push(i);
      }
    }
  }
  return bag;
}

function formatCitationGroup(indices: number[]): string {
  const arr = Array.from(new Set(indices)).sort((a, b) => a - b);
  if (!arr.length) return "[]";
  const out: string[] = [];
  let start = arr[0], prev = arr[0];

  for (let i = 1; i <= arr.length; i++) {
    const cur = arr[i];
    if (cur === prev + 1) { prev = cur; continue; }
    if (start === prev) out.push(`${start}`);
    else if (prev === start + 1) out.push(`${start}`, `${prev}`);
    else out.push(`${start}-${prev}`);
    start = cur ?? NaN;
    prev = cur ?? NaN;
  }
  return `[${out.join(", ")}]`;
}

export async function mergeAdjacentCitationsInParagraph() {
  await Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    const para = sel.paragraphs.getFirst();
    para.load("contentControls/items/tag");
    await ctx.sync();

    // Collect contiguous runs of our cite CCs
    const items = para.contentControls.items;
    let i = 0;
    while (i < items.length) {
      // start a group if this is a wordref-cite
      const group: Word.ContentControl[] = [];
      while (i < items.length && (items[i].tag || "").startsWith("wordreff-cite:")) {
        group.push(items[i]);
        i++;
      }

      if (group.length > 1) {
        // Load text of all CCs in the group
        group.forEach(cc => cc.getRange().load("text"));
        await ctx.sync();

        // Parse/merge numbers from each CC
        const allNums: number[] = [];
        for (const cc of group) {
          const txt = cc.getRange().text || "";
          allNums.push(...parseBracketIndices(txt));
        }
        const merged = formatCitationGroup(allNums);

        // Replace text in the FIRST CC, delete the rest
        group[0].insertText(merged, Word.InsertLocation.replace);
        for (let k = 1; k < group.length; k++) {
          // delete control + its contents so duplicates don’t remain
          group[k].delete(false);
        }
        await ctx.sync();
      }

      // Skip non-citation CC or advance
      while (i < items.length && !(items[i].tag || "").startsWith("wordreff-cite:")) {
        i++;
      }
    }
  });
}
/* ============== insert citation (with merge) ============== */

/**
 * Insert citation at selection. For numeric styles, if the user inserts
 * *inside* an existing WordReff citation CC, we merge indices into one bracket
 * e.g. `[2]` + `[3,4]` → `[2, 3–4]`. For author–year styles we just replace text.
 */
// cite.ts
// cite.ts (replace the whole function)
export async function insertCitationControl(
  e: BibEntry,
  text: string,
  style?: string,
  index?: number
) {
  const sty = (style || "apa").toLowerCase();
  const isNumeric = ["ieee", "numeric", "vancouver", "acs"].includes(sty);

  const isWordWeb =
    Office.context.platform === Office.PlatformType.OfficeOnline;

  const dbg = (window as any).__WORDREFF_DEBUG__ || {};
  const FORCE_APPEND_END = !!dbg.FORCE_APPEND_END;
  const NO_CC_FALLBACK = dbg.NO_CONTENT_CONTROL_FALLBACK !== false;
  const SAFE_MODE_NO_MERGE = !!dbg.SAFE_MODE_NO_MERGE;

  const setSelectedTextAsync = (t: string) =>
    new Promise<boolean>((resolve) => {
      Office.context.document.setSelectedDataAsync(
        t,
        { coercionType: Office.CoercionType.Text },
        (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded)
      );
    });

  return Word.run(async (ctx) => {
    const body = ctx.document.body;

    const insertPlainText = async (range: Word.Range) => {
      range.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
      return true;
    };

    if (FORCE_APPEND_END) {
      const p = body.insertParagraph("", Word.InsertLocation.end);
      const r = p.getRange("Start");
      await insertPlainText(r);
      return;
    }

    const sel = ctx.document.getSelection();
    sel.load("text,parentContentControl,paragraphs");
    await ctx.sync();

    // Word Web: avoid content controls entirely.
    if (isWordWeb) {
      try {
        await insertPlainText(sel);
        return;
      } catch {
        if (await setSelectedTextAsync(text)) return;

        const p = body.insertParagraph("", Word.InsertLocation.end);
        const r = p.getRange("Start");
        await insertPlainText(r);
        return;
      }
    }

    // Desktop only: merge if cursor is inside a WordReff citation.
    if (!SAFE_MODE_NO_MERGE) {
      const parent = sel.parentContentControl;

      if (
        parent &&
        !parent.isNullObject &&
        (parent.tag || "").startsWith("wordreff-cite:") &&
        isNumeric &&
        typeof index === "number"
      ) {
        const r = parent.getRange();
        r.load("text");
        await ctx.sync();

        const nums = parseBracketIndices(r.text);
        nums.push(index);

        parent.insertText(formatCitationGroup(nums), Word.InsertLocation.replace);
        await ctx.sync();
        return;
      }
    }

    // Desktop only: if just after a WordReff citation, merge with previous citation.
    if (!SAFE_MODE_NO_MERGE && isNumeric && typeof index === "number") {
      try {
        const para = sel.paragraphs.getFirst();
        para.load("text");
        await ctx.sync();

        const all = ctx.document.contentControls;
        all.load("items/tag");
        await ctx.sync();

        const items = all.items.filter((cc) =>
          (cc.tag || "").startsWith("wordreff-cite:")
        );

        for (const cc of items) {
          const rr = cc.getRange();
          rr.load("text");
          await ctx.sync();

          const end = rr.getRange("End");
          const cmp = end.compareLocationWith(sel);
          await ctx.sync();

          if (cmp.value === Word.LocationRelation.equal) {
            const nums = parseBracketIndices(rr.text);
            nums.push(index);

            cc.insertText(formatCitationGroup(nums), Word.InsertLocation.replace);
            await ctx.sync();
            return;
          }
        }
      } catch {
        // ignore and continue to normal insertion
      }
    }

    const tryCreateCC = async (range: Word.Range) => {
      try {
        const cc = range.insertContentControl();
        cc.tag = `wordreff-cite:${e.id}`;
        cc.title = "WordReff Citation";
        cc.appearance = "BoundingBox";
        cc.insertText(text, Word.InsertLocation.replace);
        await ctx.sync();
        return true;
      } catch {
        if (NO_CC_FALLBACK) {
          const ok = await setSelectedTextAsync(text);
          if (ok) return true;
        }
        return false;
      }
    };

    if (await tryCreateCC(sel)) return;

    try {
      const parent = sel.parentContentControl;

      if (
        parent &&
        !parent.isNullObject &&
        !(parent.tag || "").startsWith("wordreff-cite:")
      ) {
        const para = sel.paragraphs.getFirst();
        const afterPara = para
          .getRange("End")
          .insertParagraph("", Word.InsertLocation.after);

        const r = afterPara.getRange("Start");
        if (await tryCreateCC(r)) return;
      }
    } catch {
      // ignore
    }

    if (await setSelectedTextAsync(text)) return;

    const p = body.insertParagraph("", Word.InsertLocation.end);
    const r = p.getRange("Start");

    if (!(await tryCreateCC(r))) {
      await insertPlainText(r);
    }
  });
}