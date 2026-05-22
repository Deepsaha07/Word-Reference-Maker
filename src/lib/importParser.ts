import * as Bib from "./bibtex";
import type { BibEntry } from "./bibtex";

function slugId(title = "", author = "", year = ""): string {
  const firstAuthor =
    author
      .split(" and ")[0]
      ?.replace(/[,.\s]+/g, "-")
      .toLowerCase()
      .replace(/^-|-$/g, "") || "anon";

  const shortTitle =
    title
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 36) || "item";

  return `${firstAuthor}-${shortTitle}-${year || "nd"}`;
}

function extractYear(text: string): string {
  const matches = text.match(/\b(19|20)\d{2}\b/g);
  return matches ? matches[matches.length - 1] : "";
}

function extractDoi(text: string): string {
  const m = text.match(/\b10\.\d{4,9}\/[-._;()/:A-Z0-9]+/i);
  return m ? m[0].replace(/[.,;]+$/, "") : "";
}

function parseIeeePlainText(raw: string): BibEntry[] {
  const text = raw.replace(/\s+/g, " ").trim();

  const titleMatch = text.match(/"([^"]+)"/);
  const title = titleMatch?.[1]?.trim() || "";

  const doi = extractDoi(text);
  const year = extractYear(text);

  const beforeTitle = titleMatch ? text.slice(0, titleMatch.index).trim() : "";
  const authorText = beforeTitle.replace(/,\s*$/, "");

  const journalMatch = text.match(/\bin\s+([^,]+),\s*vol\./i);
  const journal = journalMatch?.[1]?.trim() || "";

  const volumeMatch = text.match(/\bvol\.\s*([^,\s]+)/i);
  const numberMatch = text.match(/\bno\.\s*([^,\s]+)/i);
  const pagesMatch = text.match(/\bpp\.\s*([^,]+)/i);

  const fields: Record<string, string> = {
    author: authorText,
    title,
    journal,
    year,
    doi,
  };

  if (volumeMatch?.[1]) fields.volume = volumeMatch[1];
  if (numberMatch?.[1]) fields.number = numberMatch[1];
  if (pagesMatch?.[1]) fields.pages = pagesMatch[1].trim();

  Object.keys(fields).forEach((k) => {
    if (!fields[k]) delete fields[k];
  });

  const id = slugId(title, authorText, year);

  return [
    {
      id,
      type: "article",
      fields,
      notes: [],
      createdAt: new Date().toISOString(),
    } as BibEntry,
  ];
}

function isBibtex(raw: string): boolean {
  return /@\w+\s*\{[\s\S]*\}/.test(raw);
}

function isRis(raw: string): boolean {
  return /\bTY\s*-\s*/.test(raw) && /\bER\s*-\s*/.test(raw);
}

function convertPubmedAuthors(authorText: string): string {
    // Input: Robinson K, Bontekoe K, Muellenbach J
    // Output: Robinson, K. and Bontekoe, K. and Muellenbach, J.
  
    return authorText
      .split(/\s*,\s*/)
      .map((a) => a.trim())
      .filter(Boolean)
      .map((a) => {
        const m = a.match(/^(.+?)\s+([A-Z][A-Z.\s-]*)$/);
        if (!m) return a;
  
        const surname = m[1].trim();
        const initials = m[2]
          .replace(/\s+/g, "")
          .split("")
          .filter(Boolean)
          .map((ch) => (/[A-Z]/.test(ch) ? `${ch}.` : ch))
          .join(" ");
  
        return `${surname}, ${initials}`;
      })
      .join(" and ");
  }

function parsePubmedVancouverPlainText(raw: string): BibEntry[] {
    const text = raw.replace(/\s+/g, " ").trim();
  
    const doi = extractDoi(text);
    const year = extractYear(text);
  
    // Remove DOI/PMID/PMCID tail before structural parsing
    const cleaned = text
      .replace(/\s*doi:\s*10\.\d{4,9}\/[-._;()/:A-Z0-9]+\.?/i, "")
      .replace(/\s*PMID:\s*\d+\.?/i, "")
      .replace(/\s*PMCID:\s*[A-Z0-9]+\.?/i, "")
      .trim();
  
    // Expected PubMed/Vancouver:
    // Authors. Title. Journal. 2025 Apr 18;113(2):184-188.
    const parts = cleaned
      .split(". ")
      .map((p) => p.trim())
      .filter(Boolean);
  
    const authorText = parts[0] || "";
    const title = parts[1] || "";
  
    let journal = "";
    let volume = "";
    let number = "";
    let pages = "";
  
    // Usually remaining = Journal. Date;Volume(Issue):Pages
    const rest = parts.slice(2).join(". ").trim();
  
    const journalMatch = rest.match(/^(.+?)\.\s*(?:\d{4}|[A-Za-z]{3,9})/);
    if (journalMatch?.[1]) {
      journal = journalMatch[1].trim();
    } else {
      journal = parts[2] || "";
    }
  
    const volIssuePages = rest.match(/;(\d+)(?:\(([^)]+)\))?:(\S+)/);
    if (volIssuePages) {
      volume = volIssuePages[1] || "";
      number = volIssuePages[2] || "";
      pages = (volIssuePages[3] || "").replace(/[.;]+$/, "");
    }
  
    const fields: Record<string, string> = {
      author: convertPubmedAuthors(authorText),
      title,
      journal,
      year,
      doi,
    };
  
    if (volume) fields.volume = volume;
    if (number) fields.number = number;
    if (pages) fields.pages = pages;
  
    Object.keys(fields).forEach((k) => {
      if (!fields[k]) delete fields[k];
    });
  
    const id = slugId(title, fields.author, year);
  
    return [
      {
        id,
        type: "article",
        fields,
        notes: [],
        createdAt: new Date().toISOString(),
      } as BibEntry,
    ];
  }

function isPubmedVancouverPlainText(raw: string): boolean {
    const text = raw.trim();

    return (
        /\bdoi:\s*10\./i.test(text) &&
        /\b(19|20)\d{2}\b/.test(text) &&
        /;\d+(?:\([^)]+\))?:\S+/.test(text)
    );
    }

function parseRis(raw: string): BibEntry[] {
  const entries = raw.split(/\nER\s*-\s*/i).map(s => s.trim()).filter(Boolean);

  return entries.map((block) => {
    const lines = block.split(/\r?\n/);
    const fields: Record<string, string> = {};
    const authors: string[] = [];

    for (const line of lines) {
      const m = line.match(/^([A-Z0-9]{2})\s*-\s*(.*)$/);
      if (!m) continue;

      const tag = m[1];
      const val = m[2].trim();

      if (tag === "AU") authors.push(val);
      if (tag === "TI" || tag === "T1") fields.title = val;
      if (tag === "JO" || tag === "JF" || tag === "JA") fields.journal = val;
      if (tag === "PY" || tag === "Y1") fields.year = val.match(/\d{4}/)?.[0] || val;
      if (tag === "VL") fields.volume = val;
      if (tag === "IS") fields.number = val;
      if (tag === "SP") fields.pages = val;
      if (tag === "DO") fields.doi = val;
    }

    if (authors.length) fields.author = authors.join(" and ");

    const id = slugId(fields.title, fields.author, fields.year);

    return {
      id,
      type: "article",
      fields,
      notes: [],
      createdAt: new Date().toISOString(),
    } as BibEntry;
  });
}

export function parseCitationInput(raw: string): BibEntry[] {
    const text = raw.trim();
    if (!text) return [];
  
    if (isBibtex(text)) {
      return Bib.parseBibtex(text) as BibEntry[];
    }
  
    if (isRis(text)) {
      return parseRis(text);
    }
  
    if (isPubmedVancouverPlainText(text)) {
      return parsePubmedVancouverPlainText(text);
    }
  
    if (extractDoi(text)) {
      return parseIeeePlainText(text);
    }
  
    return parseIeeePlainText(text);
  }
