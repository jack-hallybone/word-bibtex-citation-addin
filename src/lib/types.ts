// The .bib on disk is the single source of truth: parsed into CSL-JSON items keyed by citation key, never written back.

import { PRODUCT_NAME } from "./productName";

/** A CSL-JSON bibliography item, as produced by the BibTeX converter. */
export interface CslItem {
  /** The BibTeX citation key, used as the stable id. */
  id: string;
  type?: string;
  title?: string;
  author?: Array<{ family?: string; given?: string; literal?: string }>;
  issued?: { "date-parts"?: number[][] };
  // CSL has many more fields; the type stays open rather than modelling them all.
  [key: string]: unknown;
}

/** The parsed library: BibTeX key -> CSL item. */
export type Library = Map<string, CslItem>;

/** Tag written into each citation's content control; its keys, not the rendered text, are the authoritative data. Both are fixed literals, never derived from PRODUCT_NAME: changing either would orphan citations already written into existing .docx files. */
export const CITATION_TAG_PREFIX = "bibcite:";
export const BIBLIOGRAPHY_TAG = "bibcite-bibliography";

/** Key used to embed a portable snapshot of cited items inside the document. Same rule: a fixed literal, independent of PRODUCT_NAME. */
export const SNAPSHOT_SETTING = "bibcite:snapshot";

/** Shared so raising one of the two `.bib` bounds cannot leave the other lower. */
export const MAX_BIB_BYTES = 20 * 1024 * 1024;

// Word content-control tag limits vary across host versions; 200 UTF-16 code units leaves headroom under all of them.
const MAX_CITATION_TAG_LENGTH = 200;

// Characters with no advance width; ranges and the Unicode property overlap deliberately since the property is only as complete as the host's own tables.
const UNSAFE_DISPLAY_SOURCE =
  String.raw`[\u0000-\u001f\u007f-\u009f` + // C0 controls, DEL, C1 controls
  String.raw`\u00ad\u061c` + // soft hyphen, Arabic letter mark
  String.raw`\u200b-\u200f\u202a-\u202e` + // zero-width set, bidi marks/embeddings/overrides
  String.raw`\u2060-\u2064\u2066-\u2069` + // word joiner, invisible operators, isolates
  String.raw`\ufeff]` + // byte-order mark
  String.raw`|\p{Default_Ignorable_Code_Point}`; // and whatever Unicode adds next

/** U+2800 BRAILLE PATTERN BLANK takes width but draws nothing and is neither whitespace nor default-ignorable, so it can pad a key into a lookalike of another (lit-downloader's BLANK_GLYPH_CLASS). */
const KEY_BLANK_GLYPH_SOURCE = String.raw`[\s\u2800]`;
/** And BibTeX's own structural characters, which end the key or the entry. */
const KEY_STRUCTURAL_SOURCE = String.raw`[,{}]`;

/** For sanitising display text. Separate object: a `g` regex carries state. */
const UNSAFE_DISPLAY_CHARS = new RegExp(UNSAFE_DISPLAY_SOURCE, "gu");
const INVALID_KEY = new RegExp(
  `${KEY_STRUCTURAL_SOURCE}|${KEY_BLANK_GLYPH_SOURCE}|${UNSAFE_DISPLAY_SOURCE}`,
  "u",
);

// ZWNJ/ZWJ (Persian/Urdu/Kurdish joining) and variation selectors (emoji presentation) are genuine text, not hidden metadata; a citekey's stricter check still rejects them.
const TEXT_FIELD_EXEMPT_SOURCE = String.raw`\u200c|\u200d|[\ufe00-\ufe0f]`;
const UNSAFE_TEXT_FIELD_TEST = new RegExp(
  `(?!${TEXT_FIELD_EXEMPT_SOURCE})(?:${UNSAFE_DISPLAY_SOURCE})`,
  "u",
);

function hasUnsafeDisplayChars(value: string): boolean {
  return UNSAFE_TEXT_FIELD_TEST.test(value);
}

/** The same test, ignoring tabs and line endings: those are visible layout in a rendered bibliography, not hidden metadata. */
export function hasUnsafeNonLayoutChars(value: string): boolean {
  return hasUnsafeDisplayChars(value.replace(/[\t\r\n]/g, ""));
}

/** Make untrusted diagnostics safe to put in the task pane or a Word title. */
export function safeDisplayText(value: unknown): string {
  return String(value).replace(UNSAFE_DISPLAY_CHARS, "�");
}

/** An untrusted error's message, safe to display; Word, citeproc and the BibTeX parser none of them promise to throw an Error. */
export function describeError(error: unknown): string {
  return safeDisplayText(error instanceof Error ? error.message : error);
}

/** Clip untrusted text to a visible budget, the ellipsis included in the count. */
export function clipDisplayText(value: string, limit: number): string {
  return value.length > limit ? `${value.slice(0, limit - 1)}…` : value;
}

/** citeproc-js's own rich-text convention, which bibliojson writes into CSL-JSON text fields (e.g. a BibTeX `{TeX}` becomes `<span class="nocase">TeX</span>`); a plain preview strips it, since only the real citation/bibliography render goes through citeproc's engine, the thing that actually interprets it. */
const CSL_RICH_TEXT_TAG =
  /<\/?(?:b|i|sub|sup)>|<span class="nocase">|<span style="font-variant:small-caps;">|<\/span>/g;

export function stripCslRichText(value: string): string {
  return value.replace(CSL_RICH_TEXT_TAG, "");
}

function describeKey(value: unknown): string {
  if (typeof value !== "string") return JSON.stringify(value) ?? String(value);
  return JSON.stringify(safeDisplayText(clipDisplayText(value, 100)));
}

/** Rejects (never normalises) a key that could spoof its visible identity, since normalising would silently detach the citation from its source entry; pinned across projects by tests/citekey-conformance.json. */
export function validateCitationKeys(
  keys: readonly string[],
): [string, ...string[]] {
  if (!Array.isArray(keys) || keys.length === 0) {
    throw new Error("A citation must contain at least one BibTeX key.");
  }

  const unique = new Set<string>();
  for (const value of keys) {
    if (
      typeof value !== "string" ||
      !value ||
      value.length > MAX_CITATION_TAG_LENGTH - CITATION_TAG_PREFIX.length ||
      INVALID_KEY.test(value) ||
      value.normalize("NFC") !== value
    ) {
      throw new Error(`Unsupported BibTeX key: ${describeKey(value)}.`);
    }
    if (unique.has(value)) {
      throw new Error(`Duplicate BibTeX key in citation: "${value}".`);
    }
    unique.add(value);
  }
  // Unreachable, but it is what lets the return type promise a first key.
  const [first, ...rest] = [...unique];
  if (first === undefined)
    throw new Error("A citation must contain at least one BibTeX key.");
  return [first, ...rest];
}

/** CSL's own locator labels, in the order the pane offers them; the tag stores the name itself. */
export const LOCATOR_LABELS = [
  "page",
  "chapter",
  "section",
  "paragraph",
  "figure",
  "table",
  "line",
  "note",
  "volume",
] as const;
export type LocatorLabel = (typeof LOCATOR_LABELS)[number];

/** The forms a citation can take besides the standard one, named as CSL names them: "narrative" reads Smith (2020), "suppress-author" (2020). */
export const CITATION_FORMS = ["narrative", "suppress-author"] as const;
export type CitationForm = (typeof CITATION_FORMS)[number];

export interface Pinpoint {
  label: LocatorLabel;
  locator: string;
}

/** What a citation carries besides its keys. Both apply to a citation of one work only, since CSL places a locator and an author's name per work. */
export interface CitationDetails {
  form?: CitationForm;
  pinpoint?: Pinpoint;
}

export interface Citation extends CitationDetails {
  keys: [string, ...string[]];
}

// Room for a page range list such as "12–15, 18, 20–22" with plenty to spare, while a whole tag still fits MAX_CITATION_TAG_LENGTH.
export const MAX_LOCATOR_LENGTH = 40;

const isLocatorLabel = (value: string): value is LocatorLabel =>
  (LOCATOR_LABELS as readonly string[]).includes(value);
const isCitationForm = (value: string): value is CitationForm =>
  (CITATION_FORMS as readonly string[]).includes(value);

/** A locator as the pane takes it from the user: one line, single-spaced, NFC like a key. */
export function normaliseLocator(value: string): string {
  return value.normalize("NFC").replace(/\s+/g, " ").trim();
}

/** The tag's text after its keys: empty for a plain citation, else a space, the form if any, then the locator's label and the locator itself, which comes last because it may contain spaces and commas. Keys contain no whitespace, so the first space ends them. */
function detailsText(
  keys: readonly string[],
  { form, pinpoint }: CitationDetails,
): string {
  if (form === undefined && pinpoint === undefined) return "";
  if (keys.length !== 1) {
    throw new Error(
      "A page or a citation form applies to a citation of one reference. Cite the others separately.",
    );
  }
  const words: string[] = [];
  if (form !== undefined) {
    if (!isCitationForm(form))
      throw new Error(`Unsupported citation form: ${describeKey(form)}.`);
    words.push(form);
  }
  if (pinpoint !== undefined) {
    const { label, locator } = pinpoint;
    if (!isLocatorLabel(label))
      throw new Error(`Unsupported locator type: ${describeKey(label)}.`);
    if (typeof locator !== "string" || normaliseLocator(locator) !== locator) {
      throw new Error(
        `Unsupported locator: ${describeKey(locator)}. Enter it on one line, without leading or trailing spaces.`,
      );
    }
    if (!locator) throw new Error("A locator cannot be empty.");
    if (locator.length > MAX_LOCATOR_LENGTH) {
      throw new Error(
        `A page or other locator can be at most ${MAX_LOCATOR_LENGTH} characters long.`,
      );
    }
    if (hasUnsafeDisplayChars(locator)) {
      throw new Error(
        "A page or other locator cannot contain invisible or directional characters.",
      );
    }
    words.push(label, locator);
  }
  return ` ${words.join(" ")}`;
}

/** The citation `citationTag` would encode, copied field by field so nothing unchecked rides along; throws saying why it cannot be one. */
export function validateCitation(
  keys: readonly string[],
  { form, pinpoint }: CitationDetails = {},
): Citation {
  const citation: Citation = { keys: validateCitationKeys(keys) };
  if (form !== undefined) citation.form = form;
  if (pinpoint !== undefined) {
    citation.pinpoint = { label: pinpoint.label, locator: pinpoint.locator };
  }
  detailsText(citation.keys, citation);
  return citation;
}

/** Encode a citation for its content-control tag. A plain citation is the comma-separated list of its keys, the one form every earlier build wrote and reads; its details follow a space, which those builds refuse as a key rather than misread. */
export function citationTag(
  keys: readonly string[],
  details: CitationDetails = {},
): string {
  const citation = validateCitation(keys, details);
  const tag =
    CITATION_TAG_PREFIX +
    citation.keys.join(",") +
    detailsText(citation.keys, citation);
  if (tag.length > MAX_CITATION_TAG_LENGTH) {
    throw new Error(
      `This citation's reference tag is too long for Word (${tag.length} characters; ` +
        `limit ${MAX_CITATION_TAG_LENGTH}). Split it into two citations, or shorten its locator.`,
    );
  }
  return tag;
}

/** Refused by name (not dropped silently) since an earlier build's pinpoint tag must not be misread as data loss; the leading "{" matters because ";" is a legal key character too (`bibcite:2020;smith` is a real citation). */
const LEGACY_PINPOINT_TAG = /^(\d+);\{/;

/** Read the details after a tag's keys, accepting only exactly what `citationTag` writes, so a detail this build does not know (a newer build's, say) is refused, never dropped. */
function parseDetails(keys: readonly string[], text: string): CitationDetails {
  const details: CitationDetails = {};
  let rest = text;
  const first = rest.split(" ", 1)[0] ?? "";
  if (isCitationForm(first)) {
    details.form = first;
    rest = rest.slice(first.length + 1);
  }
  if (rest) {
    const space = rest.indexOf(" ");
    const label = space < 0 ? rest : rest.slice(0, space);
    if (isLocatorLabel(label) && space > 0) {
      details.pinpoint = { label, locator: rest.slice(space + 1) };
    }
  }
  let written: string | null = null;
  try {
    written = detailsText(keys, details);
  } catch {
    // Reported below with the rest, by what the tag holds.
  }
  if (written !== ` ${text}`) {
    throw new Error(
      `A ${PRODUCT_NAME} citation carries details this build does not read ` +
        `(${describeKey(text)}). It may have been written by a newer version of ${PRODUCT_NAME}; ` +
        "nothing has been changed.",
    );
  }
  return details;
}

/** Decode a content-control tag into its citation, or null if it is not one of this add-in's citations. */
export function parseCitationTag(tag: string): Citation | null {
  if (!tag || !tag.startsWith(CITATION_TAG_PREFIX)) return null;
  if (tag.length > MAX_CITATION_TAG_LENGTH) {
    throw new Error(
      `A ${PRODUCT_NAME} citation tag is longer than the supported Word-safe limit.`,
    );
  }
  const body = tag.slice(CITATION_TAG_PREFIX.length);
  if (!body) {
    throw new Error(`A ${PRODUCT_NAME} citation has an empty reference tag.`);
  }

  const versioned = LEGACY_PINPOINT_TAG.exec(body);
  if (versioned) {
    throw new Error(
      `This citation uses ${PRODUCT_NAME} tag format version ${versioned[1]}, which this build ` +
        "does not read: that form, written by an earlier build, carried a page or chapter " +
        "pinpoint. Delete the citation and insert it again, setting its page in the pane, or use " +
        "a build that reads it. Nothing has been changed.",
    );
  }

  const space = body.indexOf(" ");
  const keys = validateCitationKeys(
    (space < 0 ? body : body.slice(0, space)).split(","),
  );
  if (space < 0) return { keys };
  return { keys, ...parseDetails(keys, body.slice(space + 1)) };
}
