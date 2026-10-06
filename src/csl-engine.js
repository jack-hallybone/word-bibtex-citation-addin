/* Shared verbatim with the vendored copy in word-bibtex-citation-addin (its scripts/sync-bib-core.mjs re-copies this file) — edit here, never the copy. ES2019 syntax, since the add-in bundles it to that floor. How both apps read BibTeX for citeproc-js, which each then drives its own way: bibliojson's parser and CSL exporter, with every fix the two need. bibliojson and citeproc resolve from each app's own node_modules and outweigh the rest of the reference manager, so index.mjs leaves this file out: an app loads them only by importing it. */

import CSL from "citeproc";
import { BibLatexParser, CSLExporter } from "bibliojson";

/** citeproc HTML-escapes every other "<" it emits (verified against citeproc-js's own text_escape), so anything outside this fixed vocabulary means unvetted output, to be refused rather than shown or inserted. */
export const CITEPROC_HTML_TAG =
  /<\/?(?:i|em|b|sup|sub)>|<span style="(?:font-style:normal|font-variant:small-caps|font-variant:normal|font-weight:normal|text-decoration:none|text-decoration:underline);">|<span style="baseline">|<div class="csl-(?:entry|left-margin|right-inline|block|indent)">|<\/(?:span|div)>/g;

/** The identifier fields the reference manager writes TeX-escaped like prose but means as plain text: bib-render's TEX_ESCAPED_IDENT_FIELDS, listed again since the add-in vendors this file without bib-render. */
const TEX_ESCAPED_IDENT_FIELDS = [
  "volume",
  "volumes",
  "number",
  "part",
  "issue",
  "pages",
  "pagetotal",
  "chapter",
  "edition",
  "version",
];

/** bib-render's decodeTexEscapes, in one pass: undoes exactly the escapes the reference manager writes (`\# \$ \% \& \_ \textasciicircum{} \textasciitilde{}`) and nothing else. It reads text after bibliojson's file-wide TeX pass, which leaves those escapes as written except `\textasciitilde{}`, rewritten as `\~{}`; a `\~{}` in the file itself becomes a combining tilde, so the two never meet. */
function decodeTexEscapes(value) {
  return value.replace(/\\textasciicircum\{\}|\\~\{\}|\\[#$%&_]/g, (escaped) =>
    escaped === "\\textasciicircum{}" ? "^" : escaped.charAt(1),
  );
}

/** A parsed literal holding `text` as it stands. */
function textNodes(text) {
  return text ? [{ type: "text", text }] : [];
}

/** bibliojson 4.0.4's own `pages` split (ranges at ",", each at "--" or else "-"), so the style still sets the range's dash: citeproc prints a `12--15` it is handed as two hyphens. */
function rangeNodes(text) {
  return text.split(",").map((range) => {
    const dash = range.includes("--") ? "--" : "-";
    const [first = "", ...rest] = range.split(dash);
    const bounds = rest.length > 0 ? [first, rest.join(dash)] : [first];
    return bounds.map((bound) => textNodes(bound.trim()));
  });
}

/** bibliojson's parser with two of its internals overridden, so recheck both on any bibliojson upgrade. */
class LibraryParser extends BibLatexParser {
  /** bibliojson 4.0.4 skips a `%` comment by looking for its line break with no end-of-input check, so a comment without one loops forever (appending one is no fix: TeX substitution eats the break after `\ldots` or `\ss`), and it recurses once per comment line. The same scan, bounded and iterative. */
  skipWhitespace() {
    for (;;) {
      while (this.isWhitespace(this.input.charAt(this.pos))) this.pos++;
      if (this.input.charAt(this.pos) !== "%") return;
      const lineEnd = this.input.indexOf("\n", this.pos);
      this.pos = lineEnd === -1 ? this.input.length : lineEnd;
    }
  }

  /** Reads the identifier fields as plain text rather than as LaTeX, which drops `\textasciicircum{}` and makes a raw `_` (written before these fields were escaped, or by another tool) subscript the next character. */
  processFields() {
    super.processFields();
    const fields = this.currentEntry && this.currentEntry.fields;
    const raw = this.currentRawFields;
    if (!fields || !raw) return;
    for (const name of TEX_ESCAPED_IDENT_FIELDS) {
      const value = raw[name];
      if (!(name in fields) || typeof value !== "string") continue;
      // The placeholders bibliojson leaves for an undefined @string name or a `\textbackslash` only its own reading resolves.
      if (/[\u0870\u0871]/.test(value)) continue;
      // So does TeX another tool wrote (`2\textsuperscript{nd}`): the reference manager writes no command but its own escapes.
      if (/\\(?!textasciicircum\{\})[A-Za-z]/.test(value)) continue;
      // A line break inside a value reads as a space, as it does in BibTeX. The braces and backslashes left once the escapes are undone go, as the reference manager shows these fields.
      const text = decodeTexEscapes(value.replace(/\s*\n\s*/g, " ")).replace(
        /[{}\\]/g,
        "",
      );
      fields[name] = name === "pages" ? rangeNodes(text) : textNodes(text);
    }
  }
}

/** True when `text` opens more brace groups than it closes; a backslash escapes the next character, as in the parser's own value scan. */
function leavesBraceOpen(text) {
  let depth = 0;
  for (let index = 0; index < text.length; index++) {
    const character = text[index];
    if (character === "\\") index++;
    else if (character === "{") depth++;
    else if (character === "}") depth--;
  }
  return depth > 0;
}

/** `bibtex` as bibliojson reads it, through LibraryParser. Neither failure throws, so each app refuses in its own words, and must refuse rather than format what was half read: `errors` lists what bibliojson could not read, and `cutShort` is the key of the first entry it stopped reading partway through without saying so. @returns {import("bibliojson").BiblatexParseResult & {cutShort: string | undefined}} */
export function parseBibtex(bibtex) {
  // Locations rather than `includeRawText`, so the check below holds no second copy of a 20 MB library.
  const parser = new LibraryParser(bibtex, {
    processUnexpected: true,
    processUnknown: true,
    // BibTeX takes any text as a url, and the parser's own check refuses intranet and local addresses such as http://wiki/page; kept as written, since neither app makes it a link: the add-in gives it to Word as text, and the reference manager copies it as text.
    processInvalidURIs: true,
    includeLocation: true,
  });
  const parsed = parser.parse();
  // A quoted value ends at its first `"`, even one inside braces (BibTeX's own `"The {"}Big{"} Idea"`), and the parser then skips the rest of the entry without a diagnostic, leaving a `{` open where it stopped. Locations index `parser.input`, the text after TeX substitution, which keeps braces balanced.
  const cutShort = Object.values(parsed.entries).find(
    ({ location }) =>
      location === undefined ||
      leavesBraceOpen(parser.input.slice(location.start, location.end + 1)),
  );
  return { ...parsed, cutShort: cutShort && cutShort.entry_key };
}

/** bibliojson resolves `@article` to CSL type "article" but its field table keys on "article-journal", so without this IEEE drops the journal fields and APA renders it as a book chapter. */
function normaliseJournalArticles(entries) {
  for (const id of Object.keys(entries)) {
    const entry = entries[Number(id)];
    if (entry && entry.bib_type === "article")
      entry.bib_type = "article-journal";
  }
  return entries;
}

/** The fields naming a journal, a series or an event, and the CSL variable each becomes. bibliojson keeps only the fields an entry's type may have, so no entry has both a journaltitle and a booktitle, which also becomes container-title. */
const NAME_FIELDS = [
  ["journaltitle", "container-title"],
  ["series", "collection-title"],
  ["eventtitle", "event"],
];

/** bibliojson's exporter with two behaviours overridden, so recheck both on any bibliojson upgrade. It sentence-cases every title field in English, as CSL expects a work's own title to be stored, but a journal, series or event keeps its name as written, as BibTeX prints it: IEEE, which prints the name as stored, listed "Annalen der physik". */
class LibraryExporter extends CSLExporter {
  constructor(...args) {
    super(...args);
    /** _reformTitle sentence-cases only English text, or text in no language named. */
    this.asWritten = new CSLExporter({}, false, { language: "as written" });
  }

  getCSLEntry(id) {
    // bibliojson keeps the language of the last entry that named one (its `langid`) for every entry after it, so one German entry left the English titles after it unconverted.
    delete this.config.language;
    const entry = super.getCSLEntry(id);
    const record = this.bibDB[Number(id)];
    const fields = record && record.fields;
    for (const [field, variable] of NAME_FIELDS) {
      if (
        fields &&
        fields[field] !== undefined &&
        entry[variable] !== undefined
      )
        entry[variable] = this.asWritten._reformTitle(fields[field]);
    }
    return entry;
  }
}

/** Each of parseBibtex's `entries` as a CSL-JSON item with its key as its id, in the entries' order, or undefined for an entry the exporter dropped, which each app refuses in its own words. @returns {Array<[string, import("bibliojson").CSLEntry | undefined]>} */
export function cslItems(entries) {
  const cslByNumber = new LibraryExporter(
    normaliseJournalArticles(entries),
  ).parse();
  return Object.entries(entries).map(([numericId, entry]) => {
    const item = cslByNumber[numericId];
    if (item) item.id = entry.entry_key;
    return [entry.entry_key, item];
  });
}

/** bibliojson gives a BibTeX type with no CSL counterpart, such as @misc, CSL 1.0.1's catch-all type "article", which today's styles read as a preprint ("Preprint at (2021)" in Nature); "document" is CSL 1.0.2's catch-all. Every @article is already "article-journal" (normaliseJournalArticles), so no "article" here is a preprint. Applied as citeproc reads each item, not in cslItems, so it also reaches an item an app holds from elsewhere, as the add-in holds a document's embedded copy. */
function renderedItem(item) {
  return item && item.type === "article" ? { ...item, type: "document" } : item;
}

/**
 * A citeproc-js processor for one stylesheet, citing `items` by key.
 * @param {string} style the stylesheet's XML
 * @param {string} locale the locale it renders in
 * @param {(lang: string) => string} retrieveLocale each locale's XML, by the name citeproc asks for
 * @param {ReadonlyMap<string, unknown>} items the items it may cite, by key, as cslItems gives them
 */
export function createProcessor(style, locale, retrieveLocale, items) {
  return new CSL.Engine(
    { retrieveLocale, retrieveItem: (id) => renderedItem(items.get(id)) },
    style,
    locale,
  );
}
