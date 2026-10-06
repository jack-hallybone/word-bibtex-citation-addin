// The only module that touches citeproc, through bib-core's citation engine (src/csl-engine.js, vendored from reference-manager/bib-core), which reads BibTeX for it with every fix this add-in needs, as it does for the reference manager.

import type {
  Engine as CiteprocEngine,
  Citation as CiteprocCitation,
  CitationMode,
} from "citeproc";
import type { BiblatexParseResult } from "bibliojson";

import {
  CITEPROC_HTML_TAG,
  createProcessor,
  cslItems,
  parseBibtex,
} from "../csl-engine.js";
import { AssetCache, AssetLoader, cacheFor, fetchText } from "./assets";
import { BibliographyLayout, bibliographyOoxml, entryIndent } from "./ooxml";
import {
  describeTruncatedNameLists,
  isBenignFieldWarning,
  stripPercentComments,
} from "./sourceFidelity";
import { validateCslItem } from "./snapshot";
import { CitationFormat, builtinStyle, isStyleId, StyleId } from "./styles";
import {
  CitationDetails,
  CslItem,
  Library,
  MAX_BIB_BYTES,
  hasUnsafeNonLayoutChars,
  safeDisplayText,
  validateCitation,
  validateCitationKeys,
} from "./types";

const MAX_LIBRARY_ENTRIES = 50_000;
/** Upper bound on fidelity notes kept, so a broken library cannot flood the UI. */
const MAX_FIDELITY_NOTES = 20;
const MAX_PARSER_WARNINGS = 50;

export interface RenderResult {
  text: string;
  missing: string[];
}

/** Unlike a citation's plain `text`, the reference list carries the CSL style's own formatting. */
export interface BibliographyResult {
  /** citeproc's own rendering, restricted to `assertSafeBibliographyHtml`'s fixed tag vocabulary. */
  html: string;
  /** The same entries as the Office Open XML Word inserts: a paragraph each, laid out as the style asks. Empty exactly when `html` is. */
  ooxml: string;
  /** Each entry's hanging indent in points, 0 for none, which wordDoc.ts also gives the last entry directly. */
  indent: number;
  missing: string[];
  /** False means there is nothing to list: no cited key resolved, or each one that did is cited in the text only. Callers must then leave existing bibliography content alone rather than write an empty string over it. */
  renderable: boolean;
  /** Distinct cited keys the library defines. With `renderable` false, zero means nothing resolved; more means each one that did is cited in the text only. */
  keysResolved: number;
}

export interface DocumentRenderResult {
  citations: RenderResult[];
  bibliography: BibliographyResult;
}

export interface LoadReport {
  entries: number;
  warnings: string[];
  /** Where the render will not match the .bib source, which is content wrong in the document rather than input the parser disliked. */
  fidelity: string[];
}

/** JSON, without letting a hostile diagnostic throw out of the error path. */
function jsonText(value: unknown): string {
  try {
    // JSON.stringify yields undefined for undefined/symbols/functions, and throws on a circular structure or BigInt.
    return JSON.stringify(value) ?? `<${typeof value}>`;
  } catch {
    return `<unserialisable ${typeof value}>`;
  }
}

/** Render one untyped field of a parser diagnostic; bare `String()` on an object yields "[object Object]". */
function diagnosticText(value: unknown): string {
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean")
    return String(value);
  if (value === null) return "null";
  return jsonText(value);
}

function describeDiagnostic(value: unknown): string {
  if (!value || typeof value !== "object") return diagnosticText(value);
  const diagnostic = value as Record<string, unknown>;
  const location =
    typeof diagnostic.line === "number" ? `line ${diagnostic.line}: ` : "";
  const type =
    typeof diagnostic.type === "string"
      ? diagnostic.type.replace(/_/g, " ")
      : "parse issue";
  const entry =
    typeof diagnostic.entry === "string"
      ? ` in entry "${diagnostic.entry}"`
      : "";
  const field =
    typeof diagnostic.field_name === "string"
      ? ` field "${diagnostic.field_name}"`
      : "";
  const expected =
    diagnostic.expected === undefined
      ? ""
      : `; expected ${diagnosticText(diagnostic.expected)}`;
  const found =
    diagnostic.found === undefined
      ? ""
      : `, found ${jsonText(diagnostic.found)}`;
  const offending =
    diagnostic.value === undefined
      ? ""
      : ` (${JSON.stringify(diagnosticText(diagnostic.value).slice(0, 60))})`;
  return safeDisplayText(
    `${location}${type}${entry}${field}${offending}${expected}${found}`,
  );
}

function cloneItem(item: CslItem): CslItem {
  return JSON.parse(JSON.stringify(item)) as CslItem;
}

function missingFlag(keys: readonly string[]): string {
  // Percent-encoding keeps a key containing `][?` from forging a second flag.
  return keys.map((key) => `[?${encodeURIComponent(key)}]`).join(" ");
}

/** Cap a note list, replacing the remainder with one summary line. */
function capped(notes: string[], limit: number, noun: string): string[] {
  if (notes.length <= limit) return notes;
  return [
    ...notes.slice(0, limit),
    `plus ${notes.length - limit} more ${noun}.`,
  ];
}

/** Refuse a render, safe since nothing has been written at any call site yet. */
function abort(detail: string): never {
  throw new Error(
    `The citation processor ${detail}, so nothing was rewritten. ` +
      "Try refreshing again, and report this if it persists.",
  );
}

function assertSafeRenderedText(value: string, kind: string): void {
  if (hasUnsafeNonLayoutChars(value)) {
    abort(
      `returned ${kind} containing unsafe invisible or directional characters`,
    );
  }
}

/** No `<a>` is allowed: turning on citeproc's `wrap_url_and_doi` would need each href from the .bib sanitised, not just allowed here. */
function assertSafeBibliographyHtml(value: string): void {
  if (value.replace(CITEPROC_HTML_TAG, "").includes("<")) {
    abort(
      "returned bibliography markup outside its known formatting vocabulary",
    );
  }
}

interface BibliographyMetadata {
  entry_ids?: unknown;
  bibliography_errors?: unknown;
  hangingindent?: unknown;
  "second-field-align"?: unknown;
  maxoffset?: unknown;
}

/** The layout citeproc's metadata asks for. `maxoffset` is the widest first field, in characters, of a style that sets that field apart (IEEE's "[1]"); one that is not a small whole number lays the entries out without the column rather than at a width it cannot trust. */
function bibliographyLayout(
  metadata: BibliographyMetadata,
): BibliographyLayout {
  const { maxoffset } = metadata;
  return {
    hangingIndent: metadata.hangingindent === true,
    numberColumn:
      metadata["second-field-align"] &&
      typeof maxoffset === "number" &&
      Number.isInteger(maxoffset) &&
      maxoffset > 0 &&
      maxoffset <= 20
        ? maxoffset
        : 0,
  };
}

/** citeproc's CSL.ERROR_NO_RENDERED_FORM: the style rendered nothing for a work in its reference list. */
const NO_RENDERED_FORM = 1;

/** Verify citeproc's bibliography identity metadata as well as its entry count. `inTextOnly` names the cited works the style leaves out of its list on purpose; citeproc reports each one as having no rendered form, and only those may be missing. */
export function verifyBibliographyResult(
  value: unknown,
  expectedIds: ReadonlySet<string>,
  inTextOnly: ReadonlySet<string> = new Set(),
): string[] {
  if (!Array.isArray(value) || value.length !== 2 || !Array.isArray(value[1])) {
    abort("produced no bibliography");
  }
  const metadata: unknown = value[0];
  if (
    metadata === null ||
    typeof metadata !== "object" ||
    Array.isArray(metadata)
  ) {
    abort("returned bibliography entries without identity metadata");
  }
  const { entry_ids: entryIds, bibliography_errors: errors } =
    metadata as BibliographyMetadata;
  if (!Array.isArray(errors)) abort("reported bibliography processing errors");
  const unlisted = new Set<string>();
  for (const error of errors) {
    const { itemID, error_code: code } = (error ?? {}) as Record<
      string,
      unknown
    >;
    if (
      code !== NO_RENDERED_FORM ||
      typeof itemID !== "string" ||
      !inTextOnly.has(itemID) ||
      !expectedIds.has(itemID) ||
      unlisted.has(itemID)
    ) {
      abort("reported bibliography processing errors");
    }
    unlisted.add(itemID);
  }
  if (
    !Array.isArray(entryIds) ||
    entryIds.length !== value[1].length + unlisted.size
  ) {
    abort("returned incomplete bibliography identity metadata");
  }

  const ids: string[] = [];
  for (const group of entryIds) {
    if (
      !Array.isArray(group) ||
      group.length !== 1 ||
      typeof group[0] !== "string"
    ) {
      abort("returned ambiguous bibliography identity metadata");
    }
    if (!unlisted.has(group[0])) ids.push(group[0]);
  }
  if (
    ids.length !== expectedIds.size - unlisted.size ||
    new Set(ids).size !== ids.length ||
    ids.some((id) => !expectedIds.has(id))
  ) {
    abort("returned bibliography entries for the wrong cited works");
  }

  const entries = value[1].map((entry) => {
    if (typeof entry !== "string" || !entry.trim())
      abort("returned an empty bibliography entry");
    const trimmed = entry.trim();
    assertSafeRenderedText(trimmed, "a bibliography entry");
    assertSafeBibliographyHtml(trimmed);
    return trimmed;
  });
  return entries;
}

/** Not the exported `BibDB`: an interface gets no implicit string index signature, so `Object.entries` over one degrades to `any`. */
type ParsedEntries = BiblatexParseResult["entries"];
type ParsedEntry = NonNullable<ParsedEntries[number]>;

/** Refuse a .bib the pane cannot hold, checked on the file exactly as it is on disk. */
function assertLoadableSource(bibtex: string): void {
  if (bibtex.length > MAX_BIB_BYTES) {
    throw new Error(
      `This .bib file exceeds the ${MAX_BIB_BYTES / (1024 * 1024)} MB safety limit.`,
    );
  }
  const bytes = new TextEncoder().encode(bibtex).length;
  if (bytes > MAX_BIB_BYTES) {
    throw new Error(
      `This .bib file is ${(bytes / (1024 * 1024)).toFixed(1)} MB, over the ${
        MAX_BIB_BYTES / (1024 * 1024)
      } MB safety limit. Split it into smaller libraries.`,
    );
  }
  if (bibtex.includes("\0")) {
    throw new Error("This .bib file contains NUL bytes and is not valid text.");
  }
}

/** Refuses the whole file on any parse failure: a partially understood library would drop entries silently, surfacing as a missing key long after load reported success. */
function parseLibrarySource(source: string): BiblatexParseResult {
  const parsed = parseBibtex(source);

  if (parsed.errors.length > 0) {
    const details = parsed.errors
      .slice(0, 5)
      .map(describeDiagnostic)
      .join("; ");
    const more =
      parsed.errors.length > 5 ? `; plus ${parsed.errors.length - 5} more` : "";
    throw new Error(
      `BibTeX parsing failed: ${details}${more}. No library data was replaced.`,
    );
  }

  if (parsed.cutShort !== undefined) {
    throw new Error(
      `BibTeX parsing failed: entry "${safeDisplayText(parsed.cutShort)}" ends inside an unclosed "{", so fields ` +
        'after that point would be lost. Check its quoted values for {"} or an unmatched "{". ' +
        "No library data was replaced.",
    );
  }
  return parsed;
}

/** Refuse a library large enough that a document-wide render would hang. */
function assertEntryCountWithinLimit(
  parsedEntries: readonly ParsedEntry[],
): void {
  if (parsedEntries.length > MAX_LIBRARY_ENTRIES) {
    throw new Error(
      `This library contains ${parsedEntries.length} entries, over the ${MAX_LIBRARY_ENTRIES} entry safety limit.`,
    );
  }
}

/** Refuse the file unless every entry key is usable and unique: a duplicate resolves to whichever entry was converted last. */
function assertKeysUniqueAndValid(parsedEntries: readonly ParsedEntry[]): void {
  const seen = new Set<string>();
  const duplicates = new Set<string>();
  for (const entry of parsedEntries) {
    const key = entry.entry_key;
    try {
      validateCitationKeys([key]);
    } catch (error) {
      throw new Error(`Invalid entry key: ${(error as Error).message}`);
    }
    if (seen.has(key)) duplicates.add(key);
    seen.add(key);
  }
  if (duplicates.size > 0) {
    const examples = Array.from(duplicates).slice(0, 10);
    const remainder =
      duplicates.size > examples.length
        ? `, plus ${duplicates.size - examples.length} more`
        : "";
    throw new Error(
      `Duplicate BibTeX key${duplicates.size === 1 ? "" : "s"}: ${examples
        .map((key) => `"${key}"`)
        .join(
          ", ",
        )}${remainder}. Rename duplicates so every citation resolves unambiguously.`,
    );
  }
}

/** A fresh validated library, paired with which fields (if any) were too long and got shortened. */
interface ParsedLibrary {
  library: Library;
  truncated: Map<string, readonly string[]>;
}

/** Build a fresh validated library, never touching the live one, which is what makes the load transactional. */
function toCslLibrary(entries: ParsedEntries): ParsedLibrary {
  const next: Library = new Map();
  const truncated = new Map<string, readonly string[]>();
  for (const [key, item] of cslItems(entries)) {
    if (!item) {
      throw new Error(
        `The BibTeX-to-CSL converter dropped entry "${key}". No library data was replaced.`,
      );
    }
    const fields = new Set<string>();
    try {
      next.set(key, validateCslItem(key, item, fields));
    } catch (error) {
      throw new Error(
        `Entry "${safeDisplayText(key)}" produced unsafe or malformed citation metadata: ` +
          `${safeDisplayText((error as Error).message)}. No library data was replaced.`,
      );
    }
    if (fields.size > 0) truncated.set(key, Array.from(fields).sort());
  }
  return { library: next, truncated };
}

/** A citation as callers hold it: its keys, and any details, are validated here again before they reach citeproc. */
export type CitationInput = CitationDetails & { keys: readonly string[] };

/** What citeproc prints, in place of an error, for a work it cannot render in the form asked of it. */
const NO_PRINTED_FORM_MARKER = /\[NO_PRINTED_FORM\]|\[CSL STYLE ERROR/;

/** CSL's way of giving a citation its form; a numeric style has only the one form, so it keeps it. */
function citationMode(
  format: CitationFormat,
  { form, pinpoint }: CitationDetails,
): CitationMode | undefined {
  if (form === undefined || format === "numeric") return undefined;
  if (form === "suppress-author") return "suppress-author";
  // MLA's narrative citation is the name alone unless a page follows it: Smith, then Smith (12).
  return format === "author" && pinpoint === undefined
    ? "author-only"
    : "composite";
}

/** `style` with the names in its macro `macro` joined by "and" where the stylesheet has "&": a processor reading it renders every citation as one reading the stylesheet does but for that word. */
function joinedWithAnd(style: string, macro: string): string {
  const start = style.indexOf(`<macro name="${macro}">`);
  const end = style.indexOf("</macro>", start);
  if (start === -1 || end === -1)
    abort(`found no macro named ${macro} in the style`);
  return (
    style.slice(0, start) +
    style.slice(start, end).replace(/and="symbol"/g, 'and="text"') +
    style.slice(end)
  );
}

/** "&" or "and" between two names. */
const NAME_JOIN = / (?:&|and) /g;

/** True when two renderings of a citation differ in nothing but the words joining names, so they agree on all else, a year's letter or an author's initial included. */
function differOnlyInJoins(a: string, b: string): boolean {
  return a.replace(NAME_JOIN, " & ") === b.replace(NAME_JOIN, " & ");
}

export class CitationEngine {
  private library: Library = new Map();
  /** Fields shortened for length on import, by key; recomputed on every load rather than persisted. */
  private truncated: Map<string, readonly string[]> = new Map();
  private style: StyleId;
  private readonly assets: AssetCache;
  /** citeproc takes up to two seconds to read a large style (APA, Chicago) and milliseconds to render with it, so a processor is kept for the next render. It is out of the cache while it renders, so one a failed render may have left half-updated is never reused. */
  private kept: {
    style: StyleId;
    library: Library;
    processor: CiteprocEngine;
    /** The processor for narrative citations in a style with a `narrativeAnd`, once a render has needed one. */
    narrative: CiteprocEngine | undefined;
  } | null = null;

  constructor(style: StyleId = "apa", loadAsset: AssetLoader = fetchText) {
    this.style = style;
    this.assets = cacheFor(loadAsset);
  }

  setStyle(style: string): void {
    if (!isStyleId(style))
      throw new Error(`Unsupported citation style: ${JSON.stringify(style)}.`);
    this.style = style;
  }

  get styleId(): StyleId {
    return this.style;
  }

  /** Parse a .bib string transactionally: any failure leaves the library. */
  load(bibtex: string): LoadReport {
    assertLoadableSource(bibtex);

    // `%` is not BibTeX's own comment character, but LaTeX users write it as one.
    const parsed = parseLibrarySource(stripPercentComments(bibtex));

    const parsedEntries = Object.values(parsed.entries);
    assertEntryCountWithinLimit(parsedEntries);
    assertKeysUniqueAndValid(parsedEntries);

    const { library: next, truncated } = toCslLibrary(parsed.entries);
    if (next.size === 0) {
      throw new Error("No BibTeX entries were found in the selected file.");
    }

    this.library = next;
    this.truncated = truncated;
    const warnings = parsed.warnings
      .filter((warning) => !isBenignFieldWarning(warning))
      .map(describeDiagnostic);
    return {
      entries: next.size,
      warnings: capped(warnings, MAX_PARSER_WARNINGS, "parser warnings"),
      fidelity: capped(
        describeTruncatedNameLists(next),
        MAX_FIDELITY_NOTES,
        "entries with the same problem",
      ),
    };
  }

  /** Load the selected CSL style and its locales before synchronous formatting. */
  async ensureReady(): Promise<void> {
    await this.assets.ensure(this.style);
  }

  /** The kept processors if they read the current style and library, else a new one, with no narrative one until a render needs it. */
  private takeProcessors(): {
    processor: CiteprocEngine;
    narrative: CiteprocEngine | undefined;
  } {
    const kept = this.kept;
    this.kept = null;
    if (kept?.style === this.style && kept.library === this.library)
      return kept;
    return { processor: this.newProcessor(), narrative: undefined };
  }

  /** A processor for the current style and library; `narrativeAnd`, the style's own, reads it with that macro's names joined by "and". */
  private newProcessor(narrativeAnd?: string): CiteprocEngine {
    const ready = this.assets.ready(this.style);
    if (ready === null) {
      throw new Error(
        "Citation engine is not ready. Load its assets before formatting.",
      );
    }

    const { locale } = builtinStyle(this.style);
    return createProcessor(
      narrativeAnd === undefined
        ? ready.style
        : joinedWithAnd(ready.style, narrativeAnd),
      locale,
      // citeproc asks for en-US and the style's own locale, both loaded; en-US was the answer to any other before there was a second.
      (lang: string) =>
        ready.locales.get(lang) ?? ready.locales.get("en-US") ?? "",
      this.library,
    );
  }

  get size(): number {
    return this.library.size;
  }

  keys(): string[] {
    return Array.from(this.library.keys()).sort((a, b) => a.localeCompare(b));
  }

  item(key: string): CslItem | undefined {
    return this.library.get(key);
  }

  /** Field labels shortened for length when this key was imported (e.g. ["abstract", "title"]); empty when none were. */
  truncatedFields(key: string): readonly string[] {
    return this.truncated.get(key) ?? [];
  }

  snapshot(keys: readonly string[] = this.keys()): Record<string, CslItem> {
    const result: Record<string, CslItem> = Object.create(null) as Record<
      string,
      CslItem
    >;
    for (const key of keys) {
      const item = this.library.get(key);
      if (!item) continue;
      const clone = cloneItem(item);
      // abstract/note excluded: unrendered by every bundled style and can carry private or copyrighted text a co-author shouldn't get invisibly; a style that needs one must fall back to the .bib, never re-add it here.
      delete clone.abstract;
      delete clone.note;
      result[key] = clone;
    }
    return result;
  }

  loadSnapshot(snapshot: Record<string, CslItem>): void {
    const next: Library = new Map();
    const truncated = new Map<string, readonly string[]>();
    const entries = Object.entries(snapshot);
    if (entries.length > MAX_LIBRARY_ENTRIES) {
      throw new Error(
        `The embedded library exceeds the ${MAX_LIBRARY_ENTRIES} entry safety limit.`,
      );
    }
    for (const [key, item] of entries) {
      validateCitationKeys([key]);
      const fields = new Set<string>();
      next.set(key, cloneItem(validateCslItem(key, item, fields)));
      if (fields.size > 0) truncated.set(key, Array.from(fields).sort());
    }
    this.library = next;
    this.truncated = truncated;
  }

  /** Formats every citation as one document-wide transaction: formatting citations independently renders every IEEE citation as [1]. */
  formatDocument(citations: readonly CitationInput[]): DocumentRenderResult {
    const validated = citations.map(({ keys, ...details }) =>
      validateCitation(keys, details),
    );
    const results: RenderResult[] = validated.map(({ keys }) => ({
      text: "",
      missing: keys.filter((key) => !this.library.has(key)),
    }));
    const style = builtinStyle(this.style);

    const resolvedCited = new Set<string>();
    const processorCitations: CiteprocCitation[] = [];
    const processorIndexToDocumentIndex: number[] = [];
    validated.forEach((citation, documentIndex) => {
      const found = citation.keys.filter((key) => this.library.has(key));
      for (const key of found) resolvedCited.add(key);
      if (found.length === 0) return;
      processorIndexToDocumentIndex.push(documentIndex);
      // Details exist only on a citation of one work (validateCitation), so they apply to that one item.
      const { pinpoint } = citation;
      const mode = citationMode(style.format, citation);
      processorCitations.push({
        citationID: `bibcite-${documentIndex}`,
        citationItems: found.map((id) =>
          pinpoint === undefined
            ? { id }
            : { id, locator: pinpoint.locator, label: pinpoint.label },
        ),
        properties:
          mode === undefined ? { noteIndex: 0 } : { noteIndex: 0, mode },
      });
    });

    let bibliographyHtml = "";
    let bibliographyXml = "";
    let indent = 0;
    if (processorCitations.length > 0) {
      const { processor, narrative: keptNarrative } = this.takeProcessors();
      let narrative = keptNarrative;
      // rebuildProcessorState keeps what an earlier render learnt, such as the given name that told two Smiths apart, so each render starts from citeproc's empty state.
      const render = (
        engine: CiteprocEngine = processor,
      ): Array<[string, number, string]> => {
        engine.restoreProcessorState([]);
        engine.setOutputFormat("text");
        return engine.rebuildProcessorState(processorCitations, "text", []);
      };
      let rendered = render();
      // A style can have no printed form for a work in the form asked: APA has none for a narrative citation of a legal case, nor MLA for a page-only citation without a page. Such a citation takes the style's standard form, as every citation does in a numeric style, rather than block every refresh of its document.
      const formless = processorCitations.filter(
        (citation, index) =>
          citation.properties.mode !== undefined &&
          NO_PRINTED_FORM_MARKER.test(rendered[index]?.[2] ?? ""),
      );
      if (formless.length > 0) {
        for (const citation of formless) citation.properties = { noteIndex: 0 };
        rendered = render();
      }

      // APA joins two names with "&" in brackets but with "and" in a sentence, Smith and Jones (2020), while its stylesheet names authors with one macro for both. So where a narrative citation has an "&", the whole document is rendered again by a processor reading that macro with "and", and each such citation is taken from it if the two renderings differ in that word alone, so the letters and initials that tell works apart stay as the reference list has them.
      const { narrativeAnd } = style;
      const ampersands =
        narrativeAnd === undefined
          ? []
          : processorCitations.flatMap((citation, index) =>
              (citation.properties.mode === "composite" ||
                citation.properties.mode === "author-only") &&
              rendered[index]?.[2]?.includes("&")
                ? [index]
                : [],
            );
      if (ampersands.length > 0) {
        if (narrative === undefined)
          narrative = this.newProcessor(narrativeAnd);
        const withAnd = render(narrative);
        for (const index of ampersands) {
          const entry = rendered[index];
          const text = withAnd[index]?.[2];
          if (
            entry !== undefined &&
            text !== undefined &&
            withAnd[index]?.[0] === entry[0] &&
            differOnlyInJoins(text, entry[2])
          )
            rendered[index] = [entry[0], entry[1], text];
        }
      }

      // Written straight over content the user already has, so a short, sparse or reordered result would silently blank or mislabel it.
      processorCitations.forEach((citation, processorIndex) => {
        const entry = rendered[processorIndex];
        const text = entry?.[2]?.trim();
        if (!entry || entry[0] !== citation.citationID || !text) {
          abort("returned an incomplete citation result");
        }
        assertSafeRenderedText(text, "a citation");
        // citeproc prints a marker, not an error, where a style cannot give a work the form asked of it; checked so that marker never reaches a document.
        if (NO_PRINTED_FORM_MARKER.test(text))
          abort("returned a citation with no printed form");
        const documentIndex = processorIndexToDocumentIndex[processorIndex];
        const target =
          documentIndex === undefined ? undefined : results[documentIndex];
        if (target === undefined)
          abort(
            "returned a citation result matching no citation in the document",
          );
        target.text = text;
      });

      // In-text citations stay plain text; only the reference list takes the style's own formatting, so the format switches to html here and only here.
      processor.setOutputFormat("html");

      // Every bundled style lists each cited work once, except a work of a type its own rules may cite in the text only, such as APA's personal communication (a @letter).
      const inTextOnlyTypes = new Set<string>(style.inTextOnly);
      const inTextOnly = new Set(
        [...resolvedCited].filter((key) =>
          inTextOnlyTypes.has(this.library.get(key)?.type ?? ""),
        ),
      );
      const bibliography = processor.makeBibliography();
      if (bibliography === false) abort("produced no bibliography");
      const entries = verifyBibliographyResult(
        bibliography,
        resolvedCited,
        inTextOnly,
      );
      bibliographyHtml = entries.join("\n");
      if (!bibliographyHtml.trim() && inTextOnly.size < resolvedCited.size)
        abort("returned an empty bibliography for resolved citations");
      if (bibliographyHtml) {
        const layout = bibliographyLayout(bibliography[0]);
        indent = entryIndent(layout) / 20;
        try {
          bibliographyXml = bibliographyOoxml(entries, layout);
        } catch (error) {
          abort(
            `returned bibliography markup Word cannot be given (${safeDisplayText((error as Error).message)})`,
          );
        }
      }
      this.kept = {
        style: this.style,
        library: this.library,
        processor,
        narrative,
      };
    }

    for (const result of results) {
      if (result.missing.length > 0) {
        const flag = missingFlag(result.missing);
        result.text = result.text ? `${result.text} ${flag}` : flag;
      }
    }

    const missing = Array.from(
      new Set(results.flatMap((result) => result.missing)),
    );
    return {
      citations: results,
      bibliography: {
        html: bibliographyHtml,
        ooxml: bibliographyXml,
        indent,
        missing,
        renderable: bibliographyHtml.length > 0,
        keysResolved: resolvedCited.size,
      },
    };
  }

  /** Format a single citation in isolation, only for the moment of insertion before the document-wide refresh assigns final numbering. */
  formatCitation(citation: CitationInput): RenderResult {
    const [result] = this.formatDocument([citation]).citations;
    if (result === undefined)
      abort("returned no result for the citation being inserted");
    return result;
  }
}
