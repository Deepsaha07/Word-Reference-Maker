/// <reference types="office-js" />

import * as Bib from "../lib/bibtex";

(window as any).__WORDREF_DEBUG__ = {
  NO_CONTENT_CONTROL_FALLBACK: true,   // allow fallback to plain text if CC insert fails
  FORCE_APPEND_END: false,             // if true, always append at end of doc
  SAFE_MODE_NO_MERGE: false            // if true, disables numeric merge logic
};

// Storage helpers
import {
  getLibrary,
  upsertEntry,
  getCitedOrder,
  setCitedOrder,
  markCited,
  clearCitedOrder,
  clearLibrary,
} from "../lib/storage";

// Cite helpers
import type { BibFormatOptions } from "../lib/cite";
import * as Cite from "../lib/cite";

const DEBUG_SIMPLE_INSERT = true;
const {
  formatInText,
  insertCitationControl,  // must exist in cite.ts
  updateBibliography,
  scanCitationsInDoc,
  rerenderAllCitations,
} = Cite;


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
  const fontName    = (document.getElementById("bibFont") as HTMLSelectElement)?.value || "Times New Roman";
  const fontSize    = parseInt((document.getElementById("bibSize") as HTMLInputElement)?.value || "12", 10);
  const lineSpacing = parseInt((document.getElementById("bibSpacing") as HTMLInputElement)?.value || "14", 10);
  const alignRaw    = (document.getElementById("bibAlign") as HTMLSelectElement)?.value || "Left";
  const color       = (document.getElementById("bibColor") as HTMLInputElement)?.value || "#333333";
  const alignment   = normalizeAlignment(alignRaw);
  return { fontName, fontSize, lineSpacing, alignment, color };
}



type BibEntry = Bib.BibEntry;

const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
  document.getElementById(id) as T | null;

const getStyle = () =>
  (($("styleSelect") as HTMLSelectElement | null)?.value || "apa").toLowerCase();

/* ───────────────────────── UI init ───────────────────────── */

Office.onReady(async () => {

  // best: typed
  
  ///try {
    ///OfficeExtension.config.extendedErrorLogging = true;
  ///} catch (err) {
    ///console.warn("Could not enable extended error logging", err);
  ///}

  // Buttons
  $("btnAddOnly")?.addEventListener("click", onAddOnly);
  $("btnAddCite")?.addEventListener("click", onAddAndCite);
  $("btnInsertCite")?.addEventListener("click", onInsertOnly);
  $("btnUpdateBib")?.addEventListener("click", onUpdateBib);
  $("btnResetNumbering")?.addEventListener("click", onResetNumbering);
  $("btnClearLibrary")?.addEventListener("click", onClearLibraryClick);
  $("btnReload")?.addEventListener("click", () => location.reload());

  document.getElementById("helpIcon")?.addEventListener("click", () => {
    document.getElementById("instructionsModal")?.classList.add("show");
  });
  document.getElementById("closeInstructions")?.addEventListener("click", () => {
    document.getElementById("instructionsModal")?.classList.remove("show");
  });

  // 🎨 Font color preview update
  const fontColorInput = document.getElementById("bibColor") as HTMLInputElement;
  fontColorInput?.addEventListener("input", () => {
    fontColorInput.style.backgroundColor = fontColorInput.value;
  });

  // Initial render
  await refreshResults("");

  // Search
  const search = $("search") as HTMLInputElement | null;
  search?.addEventListener("input", async () => {
    const q = search.value;

    // If the search looks like raw BibTeX, try to find a direct match and pop/highlight it
    if (isLikelyBibtex(q)) {
      const match = await findExistingFromBibtex(q);
      await refreshResults(""); // show full list so we can highlight
      if (match?.id) {
        const container = document.getElementById("results")!;
        flashResultRow(container, match.id);
        const reasonMsg = match.reason === "citationKey" ? "citation key" : match.reason.toUpperCase();
        showToast(`Already in your library (matched by ${reasonMsg}).`);
        return; // don't filter this time (we already highlighted the match)
      }
      // otherwise fall through to normal filtering
    }

    await refreshResults(q);
  });

  // Style badge + style change handling
  updateStyleBadge();
  (document.getElementById("styleSelect") as HTMLSelectElement | null)
    ?.addEventListener("change", onStyleChanged);

  // Initial render
  await refreshResults("");
});
//const DEBUG_SIMPLE_INSERT = false; // set to true if you want to force the simple path

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

    // Helper: make sure we have a writable selection. If not, select end-of-doc.
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
        const sel2 = ctx.document.getSelection();
        return sel2;
      }
    };

    let sel: Word.Range;
    try {
      sel = await ensureSelection();
      // Try inserting a content control at the (recovered) selection
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

    // Final fallback: append to end
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
async function insertOrAppendCitationShim(e: BibEntry, style: string, lib: Record<string, BibEntry>) {
  try {
    const order = await getCitedOrder();
    let idx = order.indexOf(e.id);
    if (idx === -1) {
      // ensure it's marked cited
      await markCited(e.id);
      const refreshed = await getCitedOrder();
      idx = refreshed.indexOf(e.id);
    }
    const citeIndex = idx >= 0 ? idx + 1 : 1;

    const text = formatInText(e, style, citeIndex);
    console.log("Inserting citation:", { id: e.id, style, citeIndex, text });

    await insertCitationControl(e, text, style, citeIndex);
  } catch (err) {
    console.error("insertOrAppendCitationShim failed", err);
    throw err; // let caller showToast
  }
}
/* ───────────────────── Style badge ───────────────────── */

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

/* ───────────── first-cited index helpers (1-based) ───────────── */

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


/* ───────────────────── Button handlers ───────────────────── */

async function onAddOnly() {
  const bibbox = $("bibtexInput") as HTMLTextAreaElement | null;
  if (!bibbox) return;

  const raw = bibbox.value.trim();
  if (!raw) return;

  let entries: BibEntry[] = [];
  try {
    entries = Bib.parseBibtex(raw);
  } catch {
    showToast("Could not parse the BibTeX you pasted.");
    return;
  }

  const lib = await getLibrary();

  for (const e of entries) {
    if (lib[e.id]) {
      // Duplicate detected → highlight + toast
      await refreshResults("");
      const container = document.getElementById("results")!;
      flashResultRow(container, e.id);
      showToast(`Already in your library (key: ${e.id}).`);
      continue; // skip inserting duplicate
    }

    // Save notes when using Add (was missing)
    const where = ($("notePage") as HTMLInputElement | null)?.value.trim() || "";
    const note  = ($("noteText") as HTMLInputElement | null)?.value.trim() || "";
    if (where || note) (e as BibEntry).notes = [{ where, text: note }];

    await upsertEntry(e);
  }

  bibbox.value = "";
  await refreshResults("");
}

async function onAddAndCite() {
  const bibbox = $("bibtexInput") as HTMLTextAreaElement | null;
  if (!bibbox) return;

  const raw = bibbox.value.trim();
  if (!raw) return;

  const where = ($("notePage") as HTMLInputElement | null)?.value.trim() || "";
  const note  = ($("noteText") as HTMLInputElement | null)?.value.trim() || "";
  const style = getStyle();

  let entries: BibEntry[] = [];
  try {
    entries = Bib.parseBibtex(raw);
  } catch {
    showToast("Could not parse the BibTeX you pasted.");
    return;
  }

  let lib = await getLibrary();

  for (const e of entries) {
    if (lib[e.id]) {
      await refreshResults("");
      const container = document.getElementById("results")!;
      flashResultRow(container, e.id);
      showToast(`Already cited/added (${e.id}).`);
      continue;
    }

    if (where || note) (e as BibEntry).notes = [{ where, text: note }];
    await upsertEntry(e);
    lib[e.id] = e; // update cache

    // Ensure this ID is in first-cited order
    // Ensure this ID is in first-cited order and compute text
    const idx = await ensureCitedIndex(e.id);
    const citeText = formatInText(e, style, idx);

    // Insert (robust)
    if (DEBUG_SIMPLE_INSERT) {
      await simpleInsert(e, style);
      await refreshBibliographyNow(style);

    } else {
      await insertCitationControl(e, citeText, style, idx);
      await refreshBibliographyNow(style);

      if (["ieee","numeric","vancouver","acs"].includes(getStyle())) {
        try { await Cite.mergeAdjacentCitationsInParagraph(); } catch {}
      }
    }
  }

  // Clear inputs
  bibbox.value = "";
  const np = $("notePage") as HTMLInputElement | null; if (np) np.value = "";
  const nt = $("noteText") as HTMLInputElement | null; if (nt) nt.value = "";

  await refreshResults("");
}

async function onInsertOnly() {
  const lib = await getLibrary();
  const list = Object.values(lib) as BibEntry[];
  if (list.length === 0) return;

  const e = list[0];
  const style = getStyle();

  try {
    const idx = await ensureCitedIndex(e.id);
    const citeText = formatInText(e, style, idx);

    if (DEBUG_SIMPLE_INSERT) {
      await simpleInsert(e, style);
      await refreshBibliographyNow(style);

    } else {
      await insertCitationControl(e, citeText, style, idx);
      await refreshBibliographyNow(style);

      if (["ieee","numeric","vancouver","acs"].includes(getStyle())) {
        try { await Cite.mergeAdjacentCitationsInParagraph(); } catch {}
      }
    }
  } catch (err: any) {
    console.error("Insert citation error:", err);
    const info = err?.debugInfo ? JSON.stringify(err.debugInfo) : String(err);
    showToast("Insert citation failed: " + info);
  }
}




async function onResetNumbering() {
  if (!(await askConfirm("Reset numbering? This keeps your library but re-numbers citations from [1]."))) return;
  await clearCitedOrder();
  const style = getStyle();
  const lib = await getLibrary();
  await Cite.rerenderAllCitations(style, [], lib);
  await Cite.updateBibliography([], style);
  showToast("Numbering reset.");
}

async function onClearLibraryClick() {
  if (!(await askConfirm("Clear ENTIRE library and numbering? This cannot be undone."))) return;
  await clearLibrary();
  await clearCitedOrder();
  await refreshResults("");
  try {
    await updateBibliography([], getStyle());
  } catch (err) {
    // ignore
  }
  showToast("Library cleared.");
}
// taskpane.ts
async function refreshBibliographyNow(style: string) {
  const lib = await getLibrary();
  const ids = await scanCitationsInDoc();
  const seen = new Set<string>();
  const cited = ids
    .filter(id => (lib as any)[id] && !seen.has(id) && (seen.add(id), true))
    .map(id => (lib as any)[id]);

  await updateBibliography(cited, style, getBibFormatOpts());
}
/* ───────────────────── Results list + search ───────────────────── */

async function refreshResults(query: string) {
  const lib = await getLibrary();
  const all = Object.values(lib) as BibEntry[];

  const q = norm(query || "");
  const filtered = q
    ? all.filter((e) => {
        const blob = [
          e.id,
          e.fields.title,
          e.fields.author,
          e.fields.year,
          e.fields.journal,
          e.fields.booktitle,
          e.fields.doi,
          ...(e.notes?.map((n) => `${n.where || ""} ${n.text || ""}`) || []),
        ]
          .join(" ")
          .toLowerCase();
        return blob.includes(q);
      })
    : all;

  const container = $("results");
  if (!container) return;

  const markify = (s: string) => (q ? highlightText(s, query) : s);

  container.innerHTML = filtered
    .map((e) => {
      const title = markify(e.fields.title || e.id);
      const meta = markify(`${e.fields.author || "Anon"} — ${e.fields.year || "n.d."}`);
      const notes =
        e.notes
          ?.map(
            (n) => `<div class="note">${markify(`🔖 ${n.where || ""} ${n.text || ""}`)}</div>`
          )
          .join("") || "";

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
    })
    .join("");

  // Per-row Cite buttons
  // Per-row Cite buttons
  // Per-row Cite buttons
container.querySelectorAll<HTMLButtonElement>("button.cite-btn").forEach((btn) => {
  btn.addEventListener("click", async () => {
    const id = btn.getAttribute("data-id");
    if (!id) return;

    const lib = await getLibrary();
    const e = (lib as any)[id] as BibEntry | undefined;
    if (!e) return;

    const style = getStyle();
    const idx = await ensureCitedIndex(e.id);
    const citeText = formatInText(e, style, idx);

    try {
      if (DEBUG_SIMPLE_INSERT) {
        await simpleInsert(e, style);
      } else {
        await insertCitationControl(e, citeText, style, idx);
      }
    } catch (err: any) {
      console.error("Insert citation error:", err, err?.debugInfo);
      const info = err?.debugInfo ? JSON.stringify(err.debugInfo) : String(err);
      showToast("Insert citation failed: " + info);
    }
  });
});
}

/* ─────────── search helpers + feedback (toast, flash) ─────────── */

function norm(s: string) {
  return (s || "").toLowerCase().replace(/\s+/g, " ").trim();
}
function normTitle(s: string) {
  return norm(s).replace(/[^a-z0-9 ]/g, "");
}
function isLikelyBibtex(s: string) {
  return /@\w+\s*\{[\s\S]*\}/.test(s);
}

type BibMatch = {
  id?: string;
  reason: "citationKey" | "doi" | "title";
};

async function findExistingFromBibtex(raw: string): Promise<BibMatch | null> {
  const lib = await getLibrary();
  const all = Object.values(lib) as BibEntry[];
  let parsed: BibEntry[];
  try {
    parsed = Bib.parseBibtex(raw);
  } catch {
    return null;
  }
  if (!parsed.length) return null;
  const needle = parsed[0];

  // 1) citationKey/id
  const nid = (needle.id || "").trim();
  if (nid) {
    const hit = all.find((e) => e.id === nid);
    if (hit) return { id: hit.id, reason: "citationKey" };
  }

  // 2) DOI
  const nd = norm(needle.fields.doi || "");
  if (nd) {
    const hit = all.find((e) => norm(e.fields.doi || "") === nd);
    if (hit) return { id: hit.id, reason: "doi" };
  }

  // 3) normalized title
  const nt = normTitle(needle.fields.title || "");
  if (nt) {
    const hit = all.find((e) => normTitle(e.fields.title || "") === nt);
    if (hit) return { id: hit.id, reason: "title" };
  }

  return null;
}

function showToast(msg: string) {
  const t = document.getElementById("toast");
  if (!t) return;
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 1800);
}

function highlightText(haystackHtml: string, query: string) {
  const q = query.trim();
  if (!q) return haystackHtml;
  const tokens = q
    .split(/\s+/)
    .filter(Boolean)
    .map((tok) => tok.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
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

/* ─────────────── askConfirm (taskpane UI, no alerts) ─────────────── */

function askConfirm(message: string): Promise<boolean> {
  return new Promise((resolve) => {
    const bar = document.getElementById("confirmBar")!;
    const msg = document.getElementById("confirmMsg")!;
    const yes = document.getElementById("confirmYes")!;
    const no = document.getElementById("confirmNo")!;
    msg.textContent = message;
    bar.style.display = "flex";
    const cleanup = () => {
      bar.style.display = "none";
      yes.removeEventListener("click", onYes);
      no.removeEventListener("click", onNo);
    };
    const onYes = () => { cleanup(); resolve(true); };
    const onNo  = () => { cleanup(); resolve(false); };
    yes.addEventListener("click", onYes);
    no.addEventListener("click", onNo);
  });
}



async function onUpdateBib() {
  const style = getStyle();
  const lib = await getLibrary();

  const currentIds = await scanCitationsInDoc();
  await setCitedOrder(currentIds);

  const seen = new Set<string>();
  const cited = currentIds.filter(id => (lib as any)[id] && !seen.has(id) && (seen.add(id), true))
                          .map(id => (lib as any)[id]);

  try {
    await updateBibliography(cited, style, getBibFormatOpts());
  } catch (err: any) {
    console.error("updateBibliography failed", err);
    const info = (err && err.debugInfo) ? JSON.stringify(err.debugInfo) : String(err);
    showToast("Update bibliography failed. " + info);
  }
}

async function onStyleChanged() {
  updateStyleBadge();

  const style = getStyle();
  const lib = await getLibrary();
  const currentIds = await scanCitationsInDoc();
  await setCitedOrder(currentIds);

  await rerenderAllCitations(style, currentIds, lib);

  const seen = new Set<string>();
  const cited = currentIds.filter(id => (lib as any)[id] && !seen.has(id) && (seen.add(id), true))
                          .map(id => (lib as any)[id]);

  try {
    await updateBibliography(cited, style, getBibFormatOpts());
  } catch (err: any) {
    console.error("updateBibliography failed", err);
    const info = (err && err.debugInfo) ? JSON.stringify(err.debugInfo) : String(err);
    showToast("Update bibliography failed. " + info);
  }
}