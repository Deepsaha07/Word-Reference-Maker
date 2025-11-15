


/// <reference types="office-js" />

import * as Bib from "../lib/bibtex";
import {
  getLibrary,
  upsertEntry,
  getCitedOrder,
  setCitedOrder,
  markCited,
  clearCitedOrder,
  clearLibrary,
} from "../lib/storage";
import type { BibFormatOptions } from "../lib/cite";
import * as Cite from "../lib/cite";

const DEBUG_SIMPLE_INSERT = true;
const {
  formatInText,
  insertCitationControl,
  updateBibliography,
  scanCitationsInDoc,
  rerenderAllCitations,
} = Cite;

const GROUP_PREFIX = "wordref-group:";
const GROUP_TITLE  = "WordRef Citation (Group)";
/* ============= Types & helpers ============= */

type BibEntry = Bib.BibEntry;
type CsvRow = Record<string, string>;

const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
  document.getElementById(id) as T | null;

const getStyle = () =>
  (($("styleSelect") as HTMLSelectElement | null)?.value || "apa").toLowerCase();

function show(id: string) { document.getElementById(id)?.classList.remove("hidden"); }
function hide(id: string) { document.getElementById(id)?.classList.add("hidden"); }

function normalizeAlignment(v: string): Word.Alignment {
  switch ((v || "").toLowerCase()) {
    case "center":
    case "centered":
      return Word.Alignment.centered;
    case "right":
      return Word.Alignment.right;
    case "justify":
    case "justified":
      return Word.Alignment.justified;
    case "left":
    default:
      return Word.Alignment.left;
  }
}

function getBibFormatOpts(): BibFormatOptions {
  const fontName = (document.getElementById("bibFont") as HTMLSelectElement)?.value || "Times New Roman";
  const fontSize = parseInt((document.getElementById("bibSize") as HTMLInputElement)?.value || "12", 10);
  const lineSpacing = parseInt((document.getElementById("bibSpacing") as HTMLInputElement)?.value || "14", 10);
  const alignRaw = (document.getElementById("bibAlign") as HTMLSelectElement)?.value || "Left";
  const color = (document.getElementById("bibColor") as HTMLInputElement)?.value || "#333333";
  const alignment = normalizeAlignment(alignRaw);
  return { fontName, fontSize, lineSpacing, alignment, color };
}

const nowIso = () => new Date().toISOString();
function normalizeEntry(e: Partial<BibEntry>): BibEntry {
  if (!e || !e.id) throw new Error("normalizeEntry: missing id");
  return {
    id: String(e.id),
    type: (e.type as any) || "misc",
    fields: e.fields || {},
    notes: e.notes || [],
    createdAt: e.createdAt ?? nowIso(),
  };
}

/** Return array with first-appearance ordering (stable unique). */
function uniqueInOrder(ids: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const id of ids) {
    if (!seen.has(id)) { seen.add(id); out.push(id); }
  }
  return out;
}

/* ============= Renumbering & merge helpers ============= */

/** Re-scan doc order → save order → rerender all in-text → rebuild bibliography. */
/** Re-scan (singles + groups) → save order → rerender singles → rebuild bibliography. */
/** Re-scan doc order → save order → rerender all in-text → rebuild bibliography. */
async function refreshNumbersAndBibliography(): Promise<void> {
  const style = getStyle();
  const lib   = await getLibrary();

  // 1) Get IDs in true document order (singles + groups expanded)
  const idsInDoc   = await getAllCitationIdsInDoc(); // e.g., ["A","B","B","C","A",...]
  const uniqueOrder = uniqueInOrder(idsInDoc);       // ["A","B","C",...]

  // 2) Store first-appearance order
  await setCitedOrder(uniqueOrder);

  // 3) Re-render every in-text citation using the unique order for indices
  await rerenderAllCitations(style, uniqueOrder, lib);

  // 4) Bibliography = unique cited entries in that order
  const cited = uniqueOrder
    .map(id => lib[id])
    .filter(Boolean);

  await updateBibliography(cited, style, getBibFormatOpts());
  await refreshCitedBibtexPanel();
}

let opInFlight = false;
async function safeRefreshAll(): Promise<void> {
  if (opInFlight) return;
  opInFlight = true;
  try {
    await refreshNumbersAndBibliography();
    await rerenderGroupCitations(getStyle());
  } finally {
    opInFlight = false;
  }
}

/** Compress consecutive numbers: [1,2,3,5,6,9] -> "1–3,5–6,9" */
function compressNumericList(nums: number[]): string {
  if (!nums.length) return "";
  const out: string[] = [];
  let start = nums[0], prev = nums[0];
  for (let i = 1; i < nums.length; i++) {
    const n = nums[i];
    if (n === prev + 1) { prev = n; continue; }
    out.push(start === prev ? String(start) : `${start}–${prev}`);
    start = prev = n;
  }
  out.push(start === prev ? String(start) : `${start}–${prev}`);
  return out.join(",");
}

/** Scan both single and group WordRef CCs in true document order (may include duplicates). */
async function scanCitationsEverywhere(): Promise<string[]> {
  let result: string[] = [];

  await Word.run(async (ctx) => {
    type Item = { ids: string[]; start: Word.Range };
    const items: Item[] = [];

    // Singles
    const singles = ctx.document.contentControls.getByTitle("WordRef Citation");
    singles.load("items/tag,items/id");
    // Groups
    const groups = ctx.document.contentControls.getByTitle(GROUP_TITLE);
    groups.load("items/tag,items/id");

    await ctx.sync();

    for (const cc of singles.items) {
      const tag = cc.tag || "";
      if (!tag.startsWith("wordref-cite:")) continue;
      const id = tag.slice("wordref-cite:".length).trim();
      if (!id) continue;
      items.push({ ids: [id], start: cc.getRange("Start") });
    }

    for (const cc of groups.items) {
      const tag = cc.tag || "";
      if (!tag.startsWith(GROUP_PREFIX)) continue;
      const ids = tag.slice(GROUP_PREFIX.length)
        .split(",")
        .map(s => s.trim())
        .filter(Boolean);
      if (!ids.length) continue;
      items.push({ ids, start: cc.getRange("Start") });
    }

    // Stable sort by position in the document
    const sorted: Item[] = [];
    for (const it of items) {
      let placed = false;
      for (let i = 0; i < sorted.length; i++) {
        const rel = it.start.compareLocationWith(sorted[i].start);
        await ctx.sync();
        if (rel.value === Word.LocationRelation.before) {
          sorted.splice(i, 0, it);
          placed = true;
          break;
        }
      }
      if (!placed) sorted.push(it);
    }

    // Flatten in order; duplicates allowed (we’ll uniq later)
    result = sorted.flatMap(it => it.ids);
  });

  return result;
}
/**
 * Merge adjacent WordRef citation content controls per paragraph.
 * Numeric styles → “[2–4,6]”; Text styles → “(Smith, 2019; Lee & Kim, 2020)”.
 */
async function mergeAdjacentCitationsInParagraph(): Promise<void> {
  const style = getStyle();
  const numericLike = ["ieee", "numeric", "vancouver", "acs"].includes(style);

  await Word.run(async (ctx) => {
    const doc = ctx.document;

    const allCcs = doc.contentControls.getByTitle("WordRef Citation");
    allCcs.load("items/tag,items/id,items/paragraphs");
    await ctx.sync();

    type CcInfo = { cc: Word.ContentControl; id: number; tag: string; paraKey: string; start: Word.Range };
    const perParagraph: Record<string, CcInfo[]> = {};

    for (const cc of allCcs.items) {
      const tag = (cc.tag || "");
      if (!tag.startsWith("wordref-cite:")) continue;

      const start = cc.getRange("Start");
      const para = cc.paragraphs.getFirst();
      const oox = para.getRange().getOoxml();
      await ctx.sync(); // resolves ClientResult<string>

      const paraKey = oox.value || `para-${Math.random().toString(36).slice(2)}`;
      (perParagraph[paraKey] ||= []).push({ cc, id: cc.id, tag, paraKey, start });
    }

    // For each paragraph, sort by location and merge
    for (const group of Object.values(perParagraph)) {
      const sorted: CcInfo[] = [];

      for (const it of group) {
        let inserted = false;
        for (let i = 0; i < sorted.length; i++) {
          const cmp = it.start.compareLocationWith(sorted[i].start);
          await ctx.sync(); // resolves ClientResult<LocationRelation>
          if (cmp.value === Word.LocationRelation.before) {
            sorted.splice(i, 0, it);
            inserted = true;
            break;
          }
        }
        if (!inserted) sorted.push(it);
      }

      if (sorted.length < 2) continue;

      const run = sorted;
      const ids = run.map(({ tag }) => tag.replace("wordref-cite:", ""));
      const order = await getCitedOrder();

      if (numericLike) {
        const numbers = ids
          .map((id) => order.indexOf(id) + 1)
          .filter((n) => n > 0)
          .sort((a, b) => a - b);
        const combined = `[${compressNumericList(numbers)}]`;

        const first = run[0].cc;
        first.insertText(combined, Word.InsertLocation.replace);
        for (let i = 1; i < run.length; i++) run[i].cc.delete(false);
      } else {
        const lib = await getLibrary();
        const pieces: string[] = [];
        for (const id of ids) {
          const e = (lib as any)[id];
          if (!e) continue;
          const idx = order.indexOf(id) + 1;
          pieces.push(formatInText(e, style, idx));
        }
        const stripped = pieces.map((p) => p.replace(/^[(\[]\s*|\s*[)\]]$/g, ""));
        const combined = `(${stripped.join("; ")})`;

        const first = run[0].cc;
        first.insertText(combined, Word.InsertLocation.replace);
        for (let i = 1; i < run.length; i++) run[i].cc.delete(false);
      }
    }

    await ctx.sync();
  });
}

// ====== Global error diagnostics and safe wrappers ======
try {
  OfficeExtension.config.extendedErrorLogging = true;
} catch {}

// Log all unhandled rejections / runtime errors for debugging
window.addEventListener("unhandledrejection", (e) => {
  const err = e.reason as any;
  const dbg = err?.debugInfo ? ` [${err.debugInfo.code || ""} @ ${err.debugInfo.errorLocation || ""}]` : "";
  console.error("Unhandled rejection:", err, err?.debugInfo);
  try { showToast("Unexpected error: " + (err?.message || String(err)) + dbg); } catch {}
});

window.addEventListener("error", (e) => {
  const err = e.error as any;
  const dbg = err?.debugInfo ? ` [${err.debugInfo.code || ""} @ ${err.debugInfo.errorLocation || ""}]` : "";
  console.error("Window error:", err, err?.debugInfo);
  try { showToast("Unexpected error: " + (err?.message || String(err)) + dbg); } catch {}
});

/** Wrapper to safely run async Word operations */
async function guard<T>(label: string, fn: () => Promise<T>): Promise<T | undefined> {
  try {
    return await fn();
  } catch (err: any) {
    const dbg = err?.debugInfo ? ` [${err.debugInfo.code || ""} @ ${err.debugInfo.errorLocation || ""}]` : "";
    console.error(`[WordRef] ${label} failed:`, err, err?.debugInfo);
    showToast(`${label} failed: ${(err?.message || String(err))}${dbg}`);
    return undefined;
  }
}
/* ============= UI boot ============= */

Office.onReady(async () => {
  $("btnAddOnly")?.addEventListener("click", onAddOnly);
  $("btnAddCite")?.addEventListener("click", onAddAndCite);
  $("btnInsertCite")?.addEventListener("click", onInsertOnly);
  $("btnUpdateBib")?.addEventListener("click", onUpdateBib);
  $("btnResetNumbering")?.addEventListener("click", onResetNumbering);
  $("btnClearLibrary")?.addEventListener("click", onClearLibraryClick);
  $("btnReload")?.addEventListener("click", () => {
    void guard("Reload UI", async () => {
      // Clear main input + note fields
      const bibbox = document.getElementById("bibtexInput") as HTMLTextAreaElement | null;
      if (bibbox) bibbox.value = "";
  
      const np = document.getElementById("notePage") as HTMLInputElement | null;
      if (np) np.value = "";
  
      const nt = document.getElementById("noteText") as HTMLInputElement | null;
      if (nt) nt.value = "";
  
      // Reset search box
      const search = document.getElementById("search") as HTMLInputElement | null;
      if (search) search.value = "";
  
      // Repaint list + cited BibTeX + style badge
      await refreshResults("");
      await refreshCitedBibtexPanel();
      updateStyleBadge();
  
      showToast("WordRef panel reloaded.");
    });
  });
  // Open panels
$("btnOpenLibPanel")?.addEventListener("click", async () => {
  show("citedLibPanel");
  await refreshCitedBibtexPanel(); // always current
});

$("btnOpenImportPanel")?.addEventListener("click", () => {
  (document.getElementById("importBibtexBox") as HTMLTextAreaElement).value = "";
  show("importLibPanel");
});

// Cited panel actions
$("btnCitedClose")?.addEventListener("click", () => hide("citedLibPanel"));
$("btnCitedRefresh")?.addEventListener("click", () => { void refreshCitedBibtexPanel(); });
$("btnCitedCopy")?.addEventListener("click", copyCitedBibtexToClipboard);

// Import panel actions
$("btnImportClose")?.addEventListener("click", () => hide("importLibPanel"));
$("btnImportParse")?.addEventListener("click", () => { void onImportBibtexPasted(); });

// ensure panel shows correct content on first open
await refreshCitedBibtexPanel();

  // Export / Import
  document.getElementById("btnExport")?.addEventListener("click", async () => {
    try {
      await onExportLibrary();
      setIoStatus(platformDownloadHint());
    } catch (e) {
      setIoStatus("Export failed. See console.", 5000);
      console.error(e);
    }
  });
  document.getElementById("btnImport")?.addEventListener("click", () => {
    (document.getElementById("importFile") as HTMLInputElement)?.click();
  });
  (document.getElementById("importFile") as HTMLInputElement)
    ?.addEventListener("change", (ev) => { void onImportFilePicked(ev); });

  // Help modal
  document.getElementById("helpIcon")?.addEventListener("click", () => {
    document.getElementById("instructionsModal")?.classList.add("show");
  });
  document.getElementById("helpText")?.addEventListener("click", () => {
    document.getElementById("instructionsModal")?.classList.add("show");
  });
  document.getElementById("closeInstructions")?.addEventListener("click", () => {
    document.getElementById("instructionsModal")?.classList.remove("show");
  });

  document.getElementById("btnMergeSelected")
  ?.addEventListener("click", () => { void onMergeSelectedCitations(); });

  document.getElementById("btnUnmergeSelected")
    ?.addEventListener("click", () => { void onUnmergeSelectedCitation(); });

  // Live color swatch
  const fontColorInput = document.getElementById("bibColor") as HTMLInputElement;
  fontColorInput?.addEventListener("input", () => {
    fontColorInput.style.backgroundColor = fontColorInput.value;
  });
  const openImportModal = () => {
    document.getElementById("importModal")?.classList.remove("hidden");
    document.body.classList.add("body-no-scroll");
  };
  const closeImportModal = () => {
    document.getElementById("importModal")?.classList.add("hidden");
    document.body.classList.remove("body-no-scroll");
  };
  
  document.getElementById("btnOpenImportPanel")?.addEventListener("click", () => {
    // don’t overlap with the sheet
    document.getElementById("citedLibPanel")?.classList.add("hidden");
    const box = document.getElementById("importBibtexBox");
    if (box) (box as HTMLTextAreaElement).value = "";
    openImportModal();
  });
  document.getElementById("btnImportClose")?.addEventListener("click", closeImportModal);
  document.getElementById("importBackdrop")?.addEventListener("click", closeImportModal);
  
  // When user confirms:
  document.getElementById("btnImportParse")?.addEventListener("click", async () => {
    await onImportBibtexPasted(); // your existing parser that calls upsertEntry(...)
    closeImportModal();
  });
  // Search
  const search = $("search") as HTMLInputElement | null;
  search?.addEventListener("input", async () => {
    const q = search.value;
    if (isLikelyBibtex(q)) {
      const match = await findExistingFromBibtex(q);
      await refreshResults(""); // show all for highlight
      if (match?.id) {
        const container = document.getElementById("results")!;
        flashResultRow(container, match.id);
        const reasonMsg = match.reason === "citationKey" ? "citation key" : match.reason.toUpperCase();
        showToast(`Already in your library (matched by ${reasonMsg}).`);
        return;
      }
    }
    await refreshResults(q);
    await refreshCitedBibtexPanel();
  });

  // Style badge + on change
  updateStyleBadge();
  (document.getElementById("styleSelect") as HTMLSelectElement | null)
    ?.addEventListener("change", onStyleChanged);

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => { void onSelectionChanged(); }
  );
  // Initial results
  await refreshResults("");
});

/* ============= Insert paths ============= */

async function simpleInsert(e: BibEntry, style: string) {
  // compute numeric index from first-cited order
  let order = await getCitedOrder();
  let pos = order.indexOf(e.id);
  if (pos === -1) {
    await markCited(e.id);
    order = await getCitedOrder();
    pos = order.indexOf(e.id);
  }
  const idx = pos >= 0 ? pos + 1 : 1;
  const text = formatInText(e, style, idx);

  await Word.run(async (ctx) => {
    const body = ctx.document.body;

    // Ensure selection; else end-of-doc
    const ensureSelection = async (): Promise<Word.Range> => {
      try {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        return sel;
      } catch {
        const end = body.getRange("End");
        end.select();
        await ctx.sync();
        return ctx.document.getSelection();
      }
    };

    try {
      const sel = await ensureSelection();
      const cc = sel.insertContentControl();
      cc.tag = `wordref-cite:${e.id}`;
      cc.title = "WordRef Citation";
      cc.appearance = "BoundingBox";
      cc.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
      return;
    } catch (err) {
      console.warn("[WordRef] selection CC insert failed; will append at end", err);
    }

    // Fallback: append at end
    const p = body.insertParagraph("", Word.InsertLocation.end);
    const r = p.getRange("Start");
    const cc = r.insertContentControl();
    cc.tag = `wordref-cite:${e.id}`;
    cc.title = "WordRef Citation";
    cc.appearance = "BoundingBox";
    cc.insertText(text, Word.InsertLocation.replace);
    await ctx.sync();
  });
}

function sizeResultsPanel() {
  const container = document.getElementById("results");
  if (!container) return;

  const first = container.querySelector<HTMLElement>(".item");
  if (!first) return;

  const card = first.getBoundingClientRect().height || 120; // fallback
  const gap  = 8;   // vertical gap between cards
  const pad  = 16;  // top+bottom padding of #results

  // 2 cards + 1 gap + padding
  const max = Math.round(card * 2 + gap * 1 + pad);
  container.style.maxHeight = `${max}px`;
}
window.addEventListener("resize", () => sizeResultsPanel());

/* ============= Style badge & numbering helpers ============= */

function updateStyleBadge() {
  const badge = document.getElementById("styleBadge");
  if (!badge) return;
  const v = getStyle();
  const label =
    ({
      ieee: "IEEE (numeric)",
      numeric: "Numeric",
      apa: "APA",
      mla: "MLA",
      harvard: "Harvard",
      acs: "ACS",
      vancouver: "Vancouver",
    } as Record<string, string>)[v] || v.toUpperCase();
  badge.textContent = label;
}

async function ensureCitedIndex(id: string): Promise<number> {
  let order = await getCitedOrder();
  let pos = order.indexOf(id);
  if (pos === -1) {
    await markCited(id);
    order = await getCitedOrder();
    pos = order.indexOf(id);
  }
  return pos + 1;
}

/* ============= Button handlers ============= */

async function onAddOnly() {
  const bibbox = $("bibtexInput") as HTMLTextAreaElement | null;
  if (!bibbox) return;
  const raw = bibbox.value.trim();
  if (!raw) return;

  let entries: BibEntry[] = [];
  try {
    const parsed = Bib.parseBibtex(raw);
    entries = parsed.map((e) => normalizeEntry(e));
  } catch {
    showToast("Could not parse the BibTeX you pasted.");
    return;
  }

  const lib = await getLibrary();
  for (const e of entries) {
    if (lib[e.id]) {
      await refreshResults("");
      const container = document.getElementById("results")!;
      flashResultRow(container, e.id);
      showToast(`Already in your library (key: ${e.id}).`);
      continue;
    }
    const where = ($("notePage") as HTMLInputElement | null)?.value.trim() || "";
    const note = ($("noteText") as HTMLInputElement | null)?.value.trim() || "";
    const withNotes = where || note ? normalizeEntry({ ...e, notes: [{ where, text: note }] }) : e;
    await upsertEntry(withNotes);
  }
  bibbox.value = "";
  await refreshResults("");
}

/** When the user clicks inside a WordRef citation, jump to its bibliography entry. */
async function onSelectionChanged(): Promise<void> {
  await guard("Selection change", async () => {
    await Word.run(async (ctx) => {
      const doc = ctx.document;
      const sel = doc.getSelection();

      // Load the actual selection text so we can ignore pure caret moves
      sel.load("text");
      const parent = sel.parentContentControlOrNullObject;
      parent.load("tag,title,isNullObject");
      await ctx.sync();

      // If there is no text selected (just a caret), do nothing
      if (!sel.text || !sel.text.trim()) {
        return;
      }

      if (parent.isNullObject) {
        // Not inside any CC → ignore
        return;
      }

      const tag   = parent.tag || "";
      const title = parent.title || "";

      // Decide if this is a WordRef single or group citation
      let ids: string[] = [];
      if (title === "WordRef Citation" && tag.startsWith("wordref-cite:")) {
        ids = [tag.slice("wordref-cite:".length).trim()];
      } else if (title === GROUP_TITLE && tag.startsWith(GROUP_PREFIX)) {
        ids = tag
          .slice(GROUP_PREFIX.length)
          .split(",")
          .map(s => s.trim())
          .filter(Boolean);
      } else {
        // Some other CC → ignore
        return;
      }

      if (!ids.length) return;

      const firstId = ids[0];

      // Lookup bib entry
      const lib   = await getLibrary();
      const entry = lib[firstId];
      if (!entry) return;

      const titleText = (entry.fields.title || "").trim();
      if (!titleText) return;

      // Search for the title (usually unique in bibliography)
      const body = doc.body;
      const hits = body.search(titleText, {
        matchCase: false,
        matchWholeWord: false,
        matchPrefix: false,
        matchSuffix: false,
        matchWildcards: false,
        ignoreSpace: true,
        ignorePunct: true,
      });

      hits.load("items");
      await ctx.sync();

      if (!hits.items.length) {
        // No match → silently do nothing
        return;
      }

      // Take the first match, go to its paragraph, and select it
      const para = hits.items[0].paragraphs.getFirst();
      para.load("text");
      await ctx.sync();

      para.select();
      await ctx.sync();
    });
  });
}

/** Jump to the bibliography paragraph for a given entry id and highlight it. */
async function jumpToBibliographyEntry(entryId: string, order: string[]): Promise<void> {
  if (!entryId) return;
  const idxInOrder = order.indexOf(entryId);
  if (idxInOrder < 0) return;

  await Word.run(async (ctx) => {
    const body  = ctx.document.body;
    const paras = body.paragraphs;
    paras.load("items/text,items/font");
    await ctx.sync();

    // Find the bibliography heading paragraph (e.g., "References" / "Bibliography")
    let bibStart = -1;
    for (let i = 0; i < paras.items.length; i++) {
      const t = paras.items[i].text || "";
      if (/references|bibliography/i.test(t)) {
        bibStart = i;
        break;
      }
    }
    if (bibStart < 0) {
      // No heading found; silently do nothing
      return;
    }

    // We assume updateBibliography wrote entries as one paragraph per cited item
    const targetIdx = bibStart + 1 + idxInOrder;
    if (targetIdx >= paras.items.length) return;

    const targetPara = paras.items[targetIdx];

    // Clear highlight from all bibliography entries (optional)
    for (let j = bibStart + 1; j < paras.items.length; j++) {
      paras.items[j].font.highlightColor = "NoColor" as any;
    }

    // Highlight the target entry and select it
    targetPara.font.highlightColor = "Yellow" as any;
    const r = targetPara.getRange();
    r.select();

    await ctx.sync();
  });
}

async function onAddAndCite() {
  const bibbox = $("bibtexInput") as HTMLTextAreaElement | null;
  if (!bibbox) return;
  const raw = bibbox.value.trim();
  if (!raw) return;

  const where = ($("notePage") as HTMLInputElement | null)?.value.trim() || "";
  const note = ($("noteText") as HTMLInputElement | null)?.value.trim() || "";
  const style = getStyle();

  let entries: BibEntry[] = [];
  try {
    const parsed = Bib.parseBibtex(raw);
    entries = parsed.map((e) => normalizeEntry(e));
  } catch {
    showToast("Could not parse the BibTeX you pasted.");
    return;
  }

  let lib = await getLibrary();
  for (const base of entries) {
    if (lib[base.id]) {
      await refreshResults("");
      const container = document.getElementById("results")!;
      flashResultRow(container, base.id);
      showToast(`Already cited/added (${base.id}).`);
      continue;
    }
    const e = where || note ? normalizeEntry({ ...base, notes: [{ where, text: note }] }) : base;

    await upsertEntry(e);
    lib[e.id] = e;

    const idx = await ensureCitedIndex(e.id);
    const citeText = formatInText(e, style, idx);

    if (DEBUG_SIMPLE_INSERT) {
      await simpleInsert(e, style);
      await refreshNumbersAndBibliography();
      await rerenderGroupCitations(getStyle());
    } else {
      await insertCitationControl(e, citeText, style, idx);
      await refreshNumbersAndBibliography();
      await rerenderGroupCitations(getStyle());
    }
  }

  bibbox.value = "";
  const np = $("notePage") as HTMLInputElement | null; if (np) np.value = "";
  const nt = $("noteText") as HTMLInputElement | null; if (nt) nt.value = "";
  await refreshResults("");
  await refreshCitedBibtexPanel();
}

async function onInsertOnly() {
  const lib = await getLibrary();
  const list = Object.values(lib) as BibEntry[];
  if (!list.length) return;
  const e = list[0];
  const style = getStyle();

  await guard("Insert citation", async () => {
    const idx = await ensureCitedIndex(e.id);
    const citeText = formatInText(e, style, idx);
    if (DEBUG_SIMPLE_INSERT) {
      await simpleInsert(e, style);
    } else {
      await insertCitationControl(e, citeText, style, idx);
    }
    await refreshNumbersAndBibliography();
    await rerenderGroupCitations(getStyle());
  });
}

async function onResetNumbering() {
  if (!(await askConfirm("Reset numbering? This re-numbers all citations from [1] based on current document order."))) return;
  await clearCitedOrder(); // drop stored order
  await refreshNumbersAndBibliography();
  await rerenderGroupCitations(getStyle());
  showToast("Numbering reset and citations reflowed.");
}

async function onClearLibraryClick() {
  const ok = await askConfirm(
    "Clear ENTIRE library and numbering? This removes all in-document WordRef citations and resets your library. This cannot be undone."
  );
  if (!ok) return;

  try {
    // 0) Move cursor out of any CC to reduce edit conflicts
    await Word.run(async (ctx) => {
      ctx.document.body.getRange("End").select();
      await ctx.sync();
    });
  } catch {}

  // 1) Delete all WordRef content controls everywhere (body + headers/footers)
  try {
    await deleteAllWordRefContentControlsEverywhere();
  } catch (e) {
    console.warn("[WordRef] deleteAllWordRefContentControlsEverywhere failed:", e);
  }

  // 2) Clear storage last, so we still know tags while deleting
  try { await clearLibrary(); } catch (e) { console.warn("[WordRef] clearLibrary failed:", e); }
  try { await clearCitedOrder(); } catch (e) { console.warn("[WordRef] clearCitedOrder failed:", e); }

  // 3) Rebuild an empty bibliography (guard against selection/CC issues)
  try {
    await Word.run(async (ctx) => {
      ctx.document.body.getRange("End").select(); // avoid being inside a CC
      await ctx.sync();
    });
    await updateBibliography([], getStyle());
  } catch (e) {
    console.warn("[WordRef] updateBibliography([]) failed (non-fatal):", e);
  }

  // 4) UI refresh
  try { await refreshResults(""); } catch {}
  try { await refreshCitedBibtexPanel(); } catch {}

  showToast("Library and in-document citations cleared.");
}

/** Delete all WordRef citation CCs in body + headers/footers, even if cannotDelete is true. */
/** Delete all WordRef citation CCs in body + headers/footers, even if cannotDelete is true. */
async function deleteAllWordRefContentControlsEverywhere(): Promise<void> {
  await Word.run(async (ctx) => {
    // Collect victims from the document BODY first
    const bodyCcs = ctx.document.contentControls;
    bodyCcs.load("items/tag,items/id,items/cannotDelete");
    const sections = ctx.document.sections;
    sections.load("items");
    await ctx.sync();

    const victims: Word.ContentControl[] = [];

    // Helper: push WordRef CCs from a content-control collection into victims
    const pushWordRef = (ccs: Word.ContentControlCollection) => {
      ccs.load("items/tag,items/id,items/cannotDelete");
      return ccs;
    };

    // BODY victims
    pushWordRef(bodyCcs);
    await ctx.sync();
    victims.push(
      ...bodyCcs.items.filter((cc) => {
        const t = cc.tag || "";
        return t.startsWith("wordref-cite:") || t.startsWith("wordref-group:");
      })
    );

    // HEADER/FOOTER victims via Section.getHeader / getFooter
    const headerTypes: Word.HeaderFooterType[] = [
      Word.HeaderFooterType.primary,
      Word.HeaderFooterType.firstPage,
      Word.HeaderFooterType.evenPages,
    ];

    for (const sec of sections.items) {
      for (const ht of headerTypes) {
        try {
          const h = sec.getHeader(ht);
          const hRange = h.getRange();
          const hCcs = hRange.contentControls;
          pushWordRef(hCcs);
          await ctx.sync();
          victims.push(
            ...hCcs.items.filter((cc) => {
              const t = cc.tag || "";
              return t.startsWith("wordref-cite:") || t.startsWith("wordref-group:");
            })
          );
        } catch { /* header might not exist; ignore */ }
      }
      for (const ht of headerTypes) {
        try {
          const f = sec.getFooter(ht);
          const fRange = f.getRange();
          const fCcs = fRange.contentControls;
          pushWordRef(fCcs);
          await ctx.sync();
          victims.push(
            ...fCcs.items.filter((cc) => {
              const t = cc.tag || "";
              return t.startsWith("wordref-cite:") || t.startsWith("wordref-group:");
            })
          );
        } catch { /* footer might not exist; ignore */ }
      }
    }

    // Move caret to a safe spot (avoid being inside a CC while deleting)
    try { ctx.document.body.getRange("End").select(); await ctx.sync(); } catch {}

    // Unlock -> blank content -> delete wrapper
    for (const cc of victims) {
      try {
        if (cc.cannotDelete) cc.cannotDelete = false;
        await ctx.sync();

        try { cc.getRange().insertText("", Word.InsertLocation.replace); } catch {}
        await ctx.sync();

        try { cc.delete(true); } catch { cc.delete(false); }
        await ctx.sync();
      } catch (e) {
        console.warn("[WordRef] CC delete failed for id", cc.id, e);
      }
    }
  });
}

async function onUpdateBib() {
  try {
    await refreshNumbersAndBibliography();
    await rerenderGroupCitations(getStyle());
  } catch (err: any) {
    console.error("updateBibliography failed", err);
    const info = (err && err.debugInfo) ? JSON.stringify(err.debugInfo) : String(err);
    showToast("Update bibliography failed. " + info);
  }
}

let styleChangeTimer: number | null = null;
let styleChangeInFlight = false;

async function onStyleChanged() {
  if (styleChangeTimer) clearTimeout(styleChangeTimer);
  styleChangeTimer = window.setTimeout(async () => {
    if (styleChangeInFlight) return;
    styleChangeInFlight = true;

    const sel = document.getElementById("styleSelect") as HTMLSelectElement | null;
    if (sel) sel.disabled = true;

    await guard("Change citation style", async () => {
      updateStyleBadge();

      // Move cursor to end to avoid GeneralException when updating CCs
      await Word.run(async (ctx) => {
        ctx.document.body.getRange("End").select();
        await ctx.sync();
      });

      await refreshNumbersAndBibliography();
      await rerenderGroupCitations(getStyle());
      await refreshCitedBibtexPanel();
    });

    if (sel) sel.disabled = false;
    styleChangeInFlight = false;
  }, 150) as any as number;
}

/* ============= Export / Import ============= */

function toCsvValue(v: unknown): string {
  const s = (v ?? "").toString();
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
function rowsToCsv(rows: CsvRow[]): string {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  const lines = [
    headers.map(h => toCsvValue(h)).join(","),
    ...rows.map(r => headers.map(h => toCsvValue(r[h] ?? "")).join(",")),
  ];
  return lines.join("\n");
}
function libraryToCsv(lib: Record<string, BibEntry>): string {
  const rows: CsvRow[] = Object.values(lib).map(e => ({
    id: e.id,
    title: e.fields.title || "",
    author: e.fields.author || e.fields.editor || "",
    year: e.fields.year || "",
    journal: e.fields.journal || "",
    booktitle: e.fields.booktitle || "",
    volume: e.fields.volume || "",
    number: e.fields.number || "",
    pages: e.fields.pages || "",
    doi: e.fields.doi || "",
    notes: (e.notes || []).map(n => `${n.where || ""} ${n.text || ""}`.trim()).join(" | "),
  }));
  return rowsToCsv(rows);
}
function libraryToXml(lib: Record<string, BibEntry>): string {
  const esc = (s: string) =>
    s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
     .replace(/"/g,"&quot;").replace(/'/g,"&apos;");
  const entries = Object.values(lib).map(e => {
    const fields = Object.entries(e.fields || {})
      .map(([k,v]) => `<field name="${esc(k)}">${esc(String(v ?? ""))}</field>`)
      .join("");
    const notes = (e.notes || [])
      .map(n => `<note where="${esc(n.where || "")}">${esc(n.text || "")}</note>`)
      .join("");
    return `<entry id="${esc(e.id)}">${fields}${notes}</entry>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8"?><wordref>${entries}</wordref>`;
}

function sanitizeFileName(s: string): string {
  return (s || "").replace(/[\\/:*?"<>|]/g, "").trim() || "wordref-library";
}
function platformDownloadHint(): string {
  const p = Office?.context?.platform;
  switch (p) {
    case Office.PlatformType.Mac:          return "Saved to your Mac Downloads folder (~/Downloads).";
    case Office.PlatformType.OfficeOnline: return "Saved to your browser's Downloads folder.";
    case Office.PlatformType.PC:           return "Saved to your Windows Downloads folder.";
    default:                               return "Saved to your Downloads folder.";
  }
}

let ioStatusTimer: number | null = null;
function setIoStatus(msg: string, ttlMs = 3500) {
  const el = document.getElementById("ioStatus");
  if (!el) return;
  el.textContent = msg;
  el.classList.remove("hide");
  if (ioStatusTimer) window.clearTimeout(ioStatusTimer);
  ioStatusTimer = window.setTimeout(() => {
    el.textContent = "";
    el.classList.add("hide");
    ioStatusTimer = null;
  }, ttlMs);
}

/** Use the same page as the dialog by switching to #download-dialog (handled in HTML) */
function openDownloadDialogUrl() {
  const base = new URL(window.location.href);
  base.hash = "download-dialog";
  return base.toString();
}

function separatorForStyle(style: string): string {
  const s = (style || "").toLowerCase();
  // Numeric families usually show “[1], [2]”
  if (["ieee", "numeric", "vancouver", "acs"].includes(s)) return ", ";
  // Author–year styles usually “(Smith, 2019; Lee & Kim, 2020)”
  return "; ";
}

function isMac() {
  return Office?.context?.platform === Office.PlatformType.Mac;
}

function exportViaDialog(filename: string, data: string, mime: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const url = openDownloadDialogUrl(); // same origin, same HTML
    Office.context.ui.displayDialogAsync(
      url,
      { height: 35, width: 30, displayInIframe: !isMac() ? true : false },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded || !res.value) {
          reject(new Error("Dialog failed to open"));
          return;
        }
        const dialog = res.value;

        const onMsg = (args: any) => {
          try {
            const payload = JSON.parse(String(args?.message ?? "{}"));
            if (payload?.ok) resolve();
            else reject(new Error(payload?.error || "Download failed"));
          } catch (e) {
            reject(e instanceof Error ? e : new Error(String(e)));
          } finally {
            try { dialog.close(); } catch {}
          }
        };
        const onErr = (args: any) => {
          try { reject(new Error(`Dialog error code: ${args?.error}`)); }
          finally { try { dialog.close(); } catch {} }
        };

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, onMsg);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, onErr);

        dialog.messageChild(JSON.stringify({ filename, data, mime }));
      }
    );
  });
}

async function onExportLibrary() {
  const lib = await getLibrary();
  const fmt = ((document.getElementById("exportFormat") as HTMLSelectElement)?.value || "json").toLowerCase();

  let data = "";
  let mime = "application/json";
  let ext  = "json";

  if (fmt === "csv") { data = libraryToCsv(lib); mime = "text/csv"; ext = "csv"; }
  else if (fmt === "xml") { data = libraryToXml(lib); mime = "application/xml"; ext = "xml"; }
  else { data = JSON.stringify(lib, null, 2); }

  const stamp = new Date().toISOString().slice(0,19).replace(/[:T]/g,"-");
  const filename = sanitizeFileName(`wordref-library-${stamp}.${ext}`);

  const isWin = Office?.context?.platform === Office.PlatformType.PC;
  const isWeb = Office?.context?.platform === Office.PlatformType.OfficeOnline;
  const isMacPlat = isMac();

  if (isWin || isWeb) {
    try {
      const blob = new Blob([data], { type: mime });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = filename; a.rel = "noopener"; a.style.display = "none";
      document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
      showToast(`${filename} saved to your Downloads folder.`);
      return;
    } catch {}
  }

  if (isMacPlat) {
    try {
      await exportViaDialog(filename, data, mime);
      showToast(`Exported as ${filename} (check ~/Downloads).`);
      return;
    } catch {}
  }

  // Fallback visible link
  try {
    const blob = new Blob([data], { type: mime });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.textContent = `Download ${filename}`;
    link.href = url; link.download = filename; link.target = "_blank"; link.rel = "noopener";
    link.style.display = "inline-block"; link.style.marginTop = "8px"; link.style.wordBreak = "break-all";
    (document.getElementById("toast") || document.body).appendChild(link);
    showToast("Click to download. If blocked, try right-click → Save link as…");
    setTimeout(() => { try { URL.revokeObjectURL(url); } catch {} }, 60_000);
  } catch {}

  if (mime === "application/json" || mime.startsWith("text/")) {
    try { await navigator.clipboard.writeText(data); showToast(`Also copied to clipboard (${ext.toUpperCase()}).`); } catch {}
  }
}

/* ============= CSV/XML parsing, results, utilities ============= */

function splitCsvLine(line: string): string[] {
  const out: string[] = []; let cur = "", inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQ) {
      if (ch === '"' && line[i+1] === '"') { cur += '"'; i++; }
      else if (ch === '"') inQ = false;
      else cur += ch;
    } else {
      if (ch === '"') inQ = true;
      else if (ch === ",") { out.push(cur); cur = ""; }
      else cur += ch;
    }
  }
  out.push(cur);
  return out;
}

function parseCsvToEntries(csv: string): BibEntry[] {
  const lines = csv.replace(/\r/g, "").split("\n").filter(Boolean);
  if (!lines.length) return [];
  const headers = splitCsvLine(lines[0]).map(h => h.trim().toLowerCase());
  const idx = (k: string) => headers.indexOf(k);

  const out: BibEntry[] = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitCsvLine(lines[i]);
    const get = (k: string) => { const j = idx(k); return j >= 0 ? (cols[j] || "").trim() : ""; };
    const id = get("id") || slugId(get("title"), get("author"), get("year"));
    const fields: any = {
      title: get("title"),
      author: get("author"),
      year: get("year"),
      journal: get("journal"),
      booktitle: get("booktitle"),
      volume: get("volume"),
      number: get("number"),
      pages: get("pages"),
      doi: get("doi"),
    };
    Object.keys(fields).forEach(k => { if (!fields[k]) delete fields[k]; });
    const notesRaw = get("notes");
    const notes = notesRaw ? notesRaw.split(" | ").map(s => ({ where: "", text: s })) : [];
    out.push(normalizeEntry({ id, type: "misc", fields, notes }));
  }
  return out;
}

function parseXmlToEntries(xml: string): BibEntry[] {
  const dom = new DOMParser().parseFromString(xml, "application/xml");
  if (dom.querySelector("parsererror")) return [];
  const entries: BibEntry[] = [];
  dom.querySelectorAll("entry").forEach(node => {
    const id = node.getAttribute("id") || "";
    const fields: Record<string, string> = {};
    node.querySelectorAll("field").forEach(f => {
      const k = f.getAttribute("name") || "";
      fields[k] = f.textContent || "";
    });
    const notes = Array.from(node.querySelectorAll("note")).map(n => ({
      where: n.getAttribute("where") || "",
      text: n.textContent || ""
    }));
    if (id) entries.push(normalizeEntry({ id, type: "misc", fields, notes }));
  });
  return entries;
}

function slugId(title = "", author = "", year = ""): string {
  const t = (title || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").slice(0, 32).replace(/^-|-$/g, "");
  const a = (author || "").split(/[,\s]/)[0]?.toLowerCase() || "anon";
  const y = year || "nd";
  return [a, t || "item", y].filter(Boolean).join("-");
}

async function refreshBibliographyNow(style: string) {
  const lib = await getLibrary();
  const ids = await scanCitationsEverywhere();
  const seen = new Set<string>();
  const cited = ids.filter(id => (lib as any)[id] && !seen.has(id) && (seen.add(id), true))
                   .map(id => (lib as any)[id]);
  await updateBibliography(cited, style, getBibFormatOpts());
}


async function refreshResults(query: string) {
  const lib = await getLibrary();
  const all = Object.values(lib) as BibEntry[];
  const q = norm(query || "");
  const filtered = q
    ? all.filter((e) => {
        const blob = [
          e.id, e.fields.title, e.fields.author, e.fields.year,
          e.fields.journal, e.fields.booktitle, e.fields.doi,
          ...(e.notes?.map((n) => `${n.where || ""} ${n.text || ""}`) || []),
        ].join(" ").toLowerCase();
        return blob.includes(q);
      })
    : all;

  const container = $("results");
  if (!container) return;

  const markify = (s: string) => (q ? highlightText(s, query) : s);
  container.innerHTML = filtered.map((e) => {
    const title = markify(e.fields.title || e.id);
    const meta = markify(`${e.fields.author || "Anon"} — ${e.fields.year || "n.d."}`);
    const notes = (e.notes?.map((n) => `<div class="note">${markify(`🔖 ${n.where || ""} ${n.text || ""}`)}</div>`) || []).join("");
    return `
      <div class="row item" data-id="${e.id}">
        <div class="grow">
          <div class="title">${title}</div>
          <div class="meta">${meta}</div>
          ${notes}
        </div>
        <button class="cite-btn" data-id="${e.id}">Cite</button>
      </div>
    `;
  }).join("");

  container.querySelectorAll<HTMLButtonElement>("button.cite-btn").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const id = btn.getAttribute("data-id"); if (!id) return;
      const lib = await getLibrary();
      const e = (lib as any)[id] as BibEntry | undefined; if (!e) return;
      const style = getStyle();
      const idx = await ensureCitedIndex(e.id);
      const citeText = formatInText(e, style, idx);
      try {
        if (DEBUG_SIMPLE_INSERT) {
          await simpleInsert(e, style);
          await refreshNumbersAndBibliography();
          await rerenderGroupCitations(getStyle());
        } else {
          await insertCitationControl(e, citeText, style, idx);
          await refreshNumbersAndBibliography();
          await rerenderGroupCitations(getStyle());
        }
      } catch (err: any) {
        console.error("Insert citation error:", err, err?.debugInfo);
        const info = err?.debugInfo ? JSON.stringify(err.debugInfo) : String(err);
        showToast("Insert citation failed: " + info);
      }
    });
  });

  const count = document.getElementById("libCount");
  if (count) count.textContent = `${filtered.length} item${filtered.length === 1 ? "" : "s"}`;
  sizeResultsPanel();
}

/* ============= Utilities ============= */

function norm(s: string) { return (s || "").toLowerCase().replace(/\s+/g, " ").trim(); }
function normTitle(s: string) { return norm(s).replace(/[^a-z0-9 ]/g, ""); }
function isLikelyBibtex(s: string) { return /@\w+\s*\{[\s\S]*\}/.test(s); }

type BibMatch = { id?: string; reason: "citationKey" | "doi" | "title"; };

async function findExistingFromBibtex(raw: string): Promise<BibMatch | null> {
  const lib = await getLibrary();
  const all = Object.values(lib) as BibEntry[];
  let parsed: BibEntry[];
  try { parsed = Bib.parseBibtex(raw); } catch { return null; }
  if (!parsed.length) return null;
  const needle = parsed[0];

  const nid = (needle.id || "").trim();
  if (nid) { const hit = all.find((e) => e.id === nid); if (hit) return { id: hit.id, reason: "citationKey" }; }

  const nd = norm(needle.fields.doi || "");
  if (nd) { const hit = all.find((e) => norm(e.fields.doi || "") === nd); if (hit) return { id: hit.id, reason: "doi" }; }

  const nt = normTitle(needle.fields.title || "");
  if (nt) { const hit = all.find((e) => normTitle(e.fields.title || "") === nt); if (hit) return { id: hit.id, reason: "title" }; }

  return null;
}

/** Scan document for all WordRef citations (single + group) in document order. */
async function getAllCitationIdsInDoc(): Promise<string[]> {
  return Word.run(async (ctx) => {
    const doc = ctx.document;
    const allCcs = doc.contentControls;
    allCcs.load("items/tag,items/title");
    await ctx.sync();

    type ItemInfo = { ids: string[]; start: Word.Range };
    const items: ItemInfo[] = [];

    for (const cc of allCcs.items) {
      const tag = cc.tag || "";
      const title = cc.title || "";

      if (!tag) continue;

      // Single citation
      if (title === "WordRef Citation" && tag.startsWith("wordref-cite:")) {
        const id = tag.slice("wordref-cite:".length).trim();
        if (!id) continue;
        const start = cc.getRange("Start");
        items.push({ ids: [id], start });
      }

      // Group citation
      if (title === GROUP_TITLE && tag.startsWith(GROUP_PREFIX)) {
        const raw = tag.slice(GROUP_PREFIX.length);
        const ids = raw.split(",").map(s => s.trim()).filter(Boolean);
        if (!ids.length) continue;
        const start = cc.getRange("Start");
        items.push({ ids, start });
      }
    }

    // Sort items by location in doc
    const sorted: ItemInfo[] = [];
    for (const it of items) {
      let placed = false;
      for (let i = 0; i < sorted.length; i++) {
        const cmp = it.start.compareLocationWith(sorted[i].start);
        await ctx.sync();
        if (cmp.value === Word.LocationRelation.before) {
          sorted.splice(i, 0, it);
          placed = true;
          break;
        }
      }
      if (!placed) sorted.push(it);
    }

    // Flatten to id sequence, preserving per-group order
    const out: string[] = [];
    for (const it of sorted) out.push(...it.ids);
    return out;
  });
}

function showToast(msg: string) {
  const t = document.getElementById("toast"); if (!t) return;
  t.textContent = msg; t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 1800);
}

function highlightText(haystackHtml: string, query: string) {
  const q = query.trim(); if (!q) return haystackHtml;
  const tokens = q.split(/\s+/).filter(Boolean).map(tok => tok.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  if (!tokens.length) return haystackHtml;
  const rx = new RegExp(`(${tokens.join("|")})`, "gi");
  return haystackHtml.replace(rx, "<mark>$1</mark>");
}

function flashResultRow(container: HTMLElement, id: string) {
  const row = container.querySelector<HTMLElement>(`.item[data-id="${CSS.escape(id)}"]`);
  if (!row) return;
  row.classList.add("hit-highlight");
  row.scrollIntoView({ behavior: "smooth", block: "center" });
  setTimeout(() => row.classList.remove("hit-highlight"), 1200);
}

/** Inline confirm bar */
function askConfirm(message: string): Promise<boolean> {
  return new Promise((resolve) => {
    const bar = document.getElementById("confirmBar")!;
    const msg = document.getElementById("confirmMsg")!;
    const yes = document.getElementById("confirmYes")!;
    const no = document.getElementById("confirmNo")!;
    msg.textContent = message;
    bar.style.display = "flex";
    const cleanup = () => { bar.style.display = "none"; yes.removeEventListener("click", onYes); no.removeEventListener("click", onNo); };
    const onYes = () => { cleanup(); resolve(true); };
    const onNo  = () => { cleanup(); resolve(false); };
    yes.addEventListener("click", onYes);
    no.addEventListener("click", onNo);
  });
}

/* ============= Import handler (single definition) ============= */

async function onImportFilePicked(ev: Event): Promise<void> {
  const input = ev.currentTarget as HTMLInputElement;
  const file = input.files && input.files[0];
  if (!file) return;

  try {
    const text = await file.text();
    const name = file.name.toLowerCase();
    let entries: Bib.BibEntry[] = [];

    if (name.endsWith(".json")) {
      const obj = JSON.parse(text);
      const list: any[] = Array.isArray(obj) ? obj : Object.values(obj || {});
      entries = list
        .filter(x => x && x.id)
        .map(x => normalizeEntry(x as Partial<Bib.BibEntry>));
    } else if (name.endsWith(".csv")) {
      entries = parseCsvToEntries(text);
    } else if (name.endsWith(".xml")) {
      entries = parseXmlToEntries(text);
    } else {
      showToast("Unsupported file. Use .json, .csv, or .xml.");
      input.value = "";
      return;
    }

    if (!entries.length) {
      showToast("No entries found in file.");
      input.value = "";
      return;
    }

    for (const e of entries) {
      await upsertEntry(e);
    }

    await refreshResults("");
    showToast(`Imported ${entries.length} entr${entries.length === 1 ? "y" : "ies"}.`);
  } catch (err) {
    console.error("Import failed:", err);
    showToast("Import failed. Check console for details.");
  } finally {
    input.value = ""; // allow re-select of same file
  }
}

const BIB_BOOKMARK = "_WORDREF_BIB";

async function onUnmergeSelectedCitation(): Promise<void> {
  const style = getStyle();
  const lib   = await getLibrary();
  const order = await getCitedOrder();

  await guard("Unmerge citations", async () => {
    await Word.run(async (ctx) => {
      const sel      = ctx.document.getSelection();
      const selRange = sel.getRange();

      const groups = ctx.document.contentControls.getByTitle(GROUP_TITLE);
      groups.load("items/tag");
      await ctx.sync();

      let target: Word.ContentControl | null = null;

      // Find the group CC containing the selection
      for (const cc of groups.items) {
        const r   = cc.getRange();
        const rel = r.compareLocationWith(selRange);
        await ctx.sync();

        const val = rel.value;
        const touches =
          val === Word.LocationRelation.inside ||
          val === Word.LocationRelation.equal;

        if (touches) {
          target = cc;
          break;
        }
      }

      if (!target) {
        showToast("Place the cursor inside a merged citation to unmerge.");
        return;
      }

      const tag = target.tag || "";
      if (!tag.startsWith(GROUP_PREFIX)) {
        showToast("This citation is not a merged group.");
        return;
      }

      const ids = tag
        .slice(GROUP_PREFIX.length)
        .split(",")
        .map(s => s.trim())
        .filter(Boolean);

      if (!ids.length) {
        showToast("Group is empty.");
        return;
      }

      // We just need valid IDs; numbering & text will be rebuilt later
      const validIds = ids.filter(id => !!lib[id]);
      if (!validIds.length) {
        showToast("No valid entries in group.");
        return;
      }

      // Insert new single citation CCs BEFORE the group CC
      const groupRange = target.getRange();
      let cursor       = groupRange.insertText("", Word.InsertLocation.before);

      for (let i = 0; i < validIds.length; i++) {
        const id = validIds[i];

        if (i > 0) {
          // space between them, e.g. [1] [2] -> will be combined later if you merge again
          cursor = cursor.insertText(" ", Word.InsertLocation.after);
        }

        // Create an empty CC; text will be set by rerenderAllCitations
        const ccRange = cursor.insertText("", Word.InsertLocation.after);
        const cc      = ccRange.insertContentControl();
        cc.title      = "WordRef Citation";
        cc.tag        = `wordref-cite:${id}`;
        cc.appearance = "BoundingBox";

        cursor = cc.getRange("End");
      }

      // Remove the old merged group completely
      try {
        target.delete(false); // delete wrapper + content
      } catch {
        try {
          target.getRange().insertText("", Word.InsertLocation.replace);
          await ctx.sync();
          target.delete(true);
        } catch {}
      }

      await ctx.sync();
    });
  });

  // Re-number and rebuild bibliography so everything is consistent
  await guard("Refresh after unmerge", async () => {
    await refreshNumbersAndBibliography();
    await rerenderGroupCitations(getStyle());
  });

  showToast("Citations unmerged.");
}


/** Ensure a bookmark exists at the bibliography start. Call it after UpdateBibliography. */
async function ensureBibliographyBookmark(): Promise<void> {
  await Word.run(async (ctx) => {
    // Look for a paragraph whose text contains “References” or your bibliography heading
    const body = ctx.document.body;
    const paras = body.paragraphs;
    paras.load("items/text");
    await ctx.sync();

    let target = paras.items.find(p => /references|bibliography/i.test(p.text || ""));
    if (!target) return; // silent if user hasn’t inserted one yet

    const r = target.getRange("Start");
    r.insertBookmark(BIB_BOOKMARK);
    await ctx.sync();
  });
}

/** Replace citation display text with a hyperlink pointing to the bibliography bookmark. */
function escapeHtml(s: string) {
  return (s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Replace the content of a citation CC with clickable text.
 * If `href` is omitted, it just inserts plain text.
 */
async function setCitationTextAsHyperlink(
  cc: Word.ContentControl,
  displayText: string,
  href?: string
): Promise<void> {
  return Word.run(async (ctx) => {
    const r = cc.getRange();
    if (href) {
      const html = `<a href="${escapeHtml(href)}" target="_self">${escapeHtml(displayText)}</a>`;
      r.insertHtml(html, Word.InsertLocation.replace);
    } else {
      r.insertText(displayText, Word.InsertLocation.replace);
    }
    await ctx.sync();
  });
}

/** Minimal BibTeX generator from your BibEntry shape */
function toBibtex(e: BibEntry): string {
  const t = (e.type || "misc");
  const k = e.id;
  const f = e.fields || {};
  const lines: string[] = [`@${t}{${k},`];
  for (const [key, val] of Object.entries(f)) {
    if (!val) continue;
    lines.push(`  ${key} = {${String(val)}},`);
  }
  lines[lines.length-1] = lines[lines.length-1].replace(/,+$/, ""); // trim last comma
  lines.push("}");
  return lines.join("\n");
}

async function buildCitedBibtex(): Promise<string> {
  const order = await getCitedOrder();
  const lib = await getLibrary();
  const cited = order.map(id => lib[id]).filter(Boolean);
  return cited.map(toBibtex).join("\n\n");
}

async function refreshCitedBibtexPanel(): Promise<void> {
  const box = document.getElementById("citedBibtexBox") as HTMLTextAreaElement | null;
  if (!box) return;
  box.value = await buildCitedBibtex();
}

/** Delete all WordRef citation content controls ([single] and [group]) from the document. */
async function deleteAllCitationControls(): Promise<void> {
  await Word.run(async (ctx) => {
    const ccs = ctx.document.contentControls;
    ccs.load("items/tag");
    await ctx.sync();

    const victims = ccs.items.filter(cc => {
      const t = cc.tag || "";
      return t.startsWith("wordref-cite:") || t.startsWith("wordref-group:");
    });

    for (const v of victims) {
      try { v.delete(false); } catch {}
    }
    await ctx.sync();
  });
}
async function onImportBibtexPasted(): Promise<void> {
  const box = document.getElementById("importBibtexBox") as HTMLTextAreaElement | null;
  if (!box) return;
  const raw = (box.value || "").trim();
  if (!raw) { showToast("Paste some BibTeX first."); return; }

  let entries: BibEntry[] = [];
  try {
    entries = Bib.parseBibtex(raw).map(e => normalizeEntry(e));
  } catch (e) {
    console.error("BibTeX parse failed:", e);
    showToast("Could not parse BibTeX. Check the format.");
    return;
  }
  if (!entries.length) { showToast("No BibTeX entries found."); return; }

  const lib = await getLibrary();
  let added = 0, skipped = 0;

  for (const e of entries) {
    if (!e.id) { skipped++; continue; }
    if (lib[e.id]) { skipped++; continue; }
    await upsertEntry(e);
    (lib as any)[e.id] = e;
    added++;
  }

  await refreshResults("");
  hide("importLibPanel");
  showToast(`Imported ${added} entr${added===1?"y":"ies"}${skipped?`, skipped ${skipped} duplicate${skipped===1?"":"s"}`:""}.`);
}
function isNumericStyle(style: string): boolean {
  const s = (style || "").toLowerCase();
  return ["ieee", "numeric", "vancouver", "acs"].includes(s);
}

async function formatGroupFromIds(ids: string[], style: string): Promise<string> {
  const lib   = await getLibrary();
  const order = await getCitedOrder();

  // unique preserve order
  const uniq = ids.filter((x, i, a) => a.indexOf(x) === i);

  const numericIdx: number[] = [];
  const alphaBits: string[] = [];

  for (const id of uniq) {
    const e = lib[id];
    if (!e) continue;
    const idx = order.indexOf(id) + 1;
    if (idx > 0) numericIdx.push(idx);

    // author–year with formatInText; strip outer () or []
    let t = formatInText(e, style, Math.max(idx, 1));
    t = t.replace(/^[\s\[(]+|[\s\])]+$/g, "");
    if (!isNumericStyle(style)) alphaBits.push(t);
  }

  numericIdx.sort((a,b)=>a-b);

  if (isNumericStyle(style)) {
    const numText = numericIdx.length ? `[${compressNumericList(numericIdx)}]` : "";
    return numText || "[]";
  } else {
    // Author–year style: just comma/semi-colon join. Most styles prefer “; ”
    const alphaText = alphaBits.length ? `(${alphaBits.join("; ")})` : "()";
    return alphaText;
  }
}

/** Update all group controls to current style */
async function rerenderGroupCitations(style: string): Promise<void> {
  await Word.run(async (ctx) => {
    const groups = ctx.document.contentControls.getByTitle(GROUP_TITLE);
    groups.load("items/tag,items/id");
    await ctx.sync();

    for (const cc of groups.items) {
      const tag = cc.tag || "";
      if (!tag.startsWith(GROUP_PREFIX)) continue;
      const ids = tag.slice(GROUP_PREFIX.length).split(",").map(s => s.trim()).filter(Boolean);
      const text = await formatGroupFromIds(ids, style);
      cc.insertText(text, Word.InsertLocation.replace);
    }
    await ctx.sync();
  });
}
async function formatMixedFromIds(ids: string[], style: string): Promise<string> {
  const lib   = await getLibrary();
  const order = await getCitedOrder();

  // de-dup while preserving first appearance within the selection
  const uniq = ids.filter((x, i, a) => a.indexOf(x) === i);

  const numericIdx: number[] = [];
  const alphaBits: string[] = [];

  for (const id of uniq) {
    const e = lib[id];
    if (!e) continue;
    const idx = order.indexOf(id) + 1;

    if (idx > 0) {
      numericIdx.push(idx);
    } else {
      // fallback to author-year style text if not numbered yet
      let s = formatInText(e, style, 1);
      s = s.replace(/^[\s\[(]+|[\s\])]+$/g, "");
      alphaBits.push(s);
    }
  }

  numericIdx.sort((a,b)=>a-b);
  const numText   = numericIdx.length ? `[${compressNumericList(numericIdx)}]` : "";
  const alphaText = alphaBits.length  ? `(${alphaBits.join("; ")})`           : "";

  if (numText && alphaText) return `${numText}, ${alphaText}`;
  return numText || alphaText || "";
}

function guid(): string {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {
    const r = (Math.random() * 16) | 0, v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

async function onMergeSelectedCitations(): Promise<void> {
  const style = getStyle();
  const isNumericStyleLocal = (s: string) =>
    ["ieee", "numeric", "vancouver", "acs"].includes((s || "").toLowerCase());

  const renderMergedText = async (ids: string[], styleNow: string): Promise<string> => {
    const lib = await getLibrary();
    const order = await getCitedOrder();
    const uniq = ids.filter((x, i, a) => a.indexOf(x) === i);

    if (isNumericStyleLocal(styleNow)) {
      const nums = uniq.map(id => order.indexOf(id) + 1).filter(n => n > 0).sort((a, b) => a - b);
      const body = nums.length ? compressNumericList(nums) : "";
      return body ? `[${body}]` : "[]";
    } else {
      const bits: string[] = [];
      for (const id of uniq) {
        const e = (lib as any)[id];
        if (!e) continue;
        const idx = Math.max(order.indexOf(id) + 1, 1);
        let t = formatInText(e, styleNow, idx);
        t = t.replace(/^[\s\[(]+|[\s\])]+$/g, "");
        bits.push(t);
      }
      return bits.length ? `(${bits.join("; ")})` : "()";
    }
  };

  await guard("Merge selected citations", async () => {
    await Word.run(async (ctx) => {
      const selRange = ctx.document.getSelection(); // correct
      const singles = ctx.document.contentControls.getByTitle("WordRef Citation");
      singles.load("items/tag,items/id");
      await ctx.sync();

      type Pick = { cc: Word.ContentControl; tag: string; start: Word.Range };
      const picked: Pick[] = [];

      for (const cc of singles.items) {
        const tag = cc.tag || "";
        if (!tag.startsWith("wordref-cite:")) continue;

        const rStart = cc.getRange("Start");
        const rEnd = cc.getRange("End");
        const cmpA = rStart.compareLocationWith(selRange);
        const cmpB = rEnd.compareLocationWith(selRange);
        await ctx.sync();

        const inside =
          (cmpA.value === Word.LocationRelation.inside || cmpA.value === Word.LocationRelation.equal) &&
          (cmpB.value === Word.LocationRelation.inside || cmpB.value === Word.LocationRelation.equal);
        if (inside) picked.push({ cc, tag, start: rStart });
      }

      if (picked.length < 2) {
        showToast("Select two or more WordRef citations to merge.");
        return;
      }

      // Safe manual sort
      const sorted: Pick[] = [];
      for (const it of picked) {
        let placed = false;
        for (let i = 0; i < sorted.length; i++) {
          const cmp = it.start.compareLocationWith(sorted[i].start);
          await ctx.sync();
          if (cmp.value === Word.LocationRelation.before) {
            sorted.splice(i, 0, it);
            placed = true;
            break;
          }
        }
        if (!placed) sorted.push(it);
      }

      const ids = sorted.map(p => p.tag.replace("wordref-cite:", ""));
      const mergedText = await renderMergedText(ids, style);
      const groupTag = "wordref-group:" + ids.join(",");

      // Move caret away to avoid edit conflicts
      ctx.document.body.getRange("End").select();
      await ctx.sync();

      // First CC becomes group
      const firstCC = sorted[0].cc;
      firstCC.title = "WordRef Citation (Group)";
      firstCC.tag = groupTag;
      firstCC.appearance = "BoundingBox";
      firstCC.insertText(mergedText, Word.InsertLocation.replace);
      await ctx.sync();

      // Delete the rest
      for (let i = 1; i < sorted.length; i++) {
        try {
          sorted[i].cc.delete(false);
        } catch {
          try {
            sorted[i].cc.getRange().insertText("", Word.InsertLocation.replace);
            await ctx.sync();
            sorted[i].cc.delete(true);
          } catch {}
        }
      }

      await ctx.sync();
    });
  });

  // Recalculate numbering and repaint
  await guard("Post-merge refresh", async () => {
    await refreshNumbersAndBibliography();
    await rerenderGroupCitations(getStyle());
  });

  showToast("Citations merged.");
}
async function safeDeleteContentControlsById(ids: number[], keepContent = true): Promise<void> {
  await Word.run(async (ctx) => {
    const set = ctx.document.contentControls;
    for (const id of ids) {
      try {
        const cc = set.getById(id);
        cc.load("id,cannotDelete,isNullObject");
        await ctx.sync();

        if (cc.isNullObject) continue;
        if (cc.cannotDelete) cc.cannotDelete = false;
        await ctx.sync();

        cc.delete(keepContent);
      } catch (e) {
        console.warn("safeDelete failed for CC id:", id, e);
      }
    }
    await ctx.sync();
  });
}

function copyCitedBibtexToClipboard(): void {
  const box = document.getElementById("citedBibtexBox") as HTMLTextAreaElement | null;
  if (!box) return;
  box.select();
  document.execCommand("copy");
  showToast("Cited BibTeX copied to clipboard.");
}


