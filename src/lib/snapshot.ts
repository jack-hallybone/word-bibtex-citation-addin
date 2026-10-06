import { SafeInputError, parseJson } from "../safe-json.js";
import { PRODUCT_NAME } from "./productName";
import { builtinStyle, isStyleId, styleRevision, StyleId } from "./styles";
import {
  CslItem,
  hasUnsafeNonLayoutChars,
  safeDisplayText,
  validateCitationKeys,
} from "./types";

const SNAPSHOT_VERSION = 2;
const MAX_SNAPSHOT_BYTES = 800 * 1024;
const MAX_SNAPSHOT_ENTRIES = 20_000;
/** Guards against pathological nesting reaching citeproc or JSON.stringify. */
const MAX_ITEM_DEPTH = 8;
/** The envelope costs three levels, plus one spare, so this can never reject what `assertDepth` accepts. */
const MAX_SNAPSHOT_DEPTH = MAX_ITEM_DEPTH + 4;
/** Reviver visits, not entries: 800 KB of real records is about 85,000, and this refuses the shape the byte cap does not — a flat array of two-byte scalars. */
const MAX_SNAPSHOT_KEYS = 200_000;

/** Opens every `parseJson` refusal, so it names what the user is looking at. */
const SNAPSHOT_WHAT = "The citation data embedded in this document";

/** CSL date variables, validated because citeproc parses them structurally. */
const DATE_FIELDS = new Set([
  "accessed",
  "available-date",
  "event-date",
  "issued",
  "original-date",
  "submitted",
]);

/** CSL name variables, all sharing the same array-of-name-objects shape. */
const NAME_FIELDS = new Set([
  "author",
  "chair",
  "collection-editor",
  "compiler",
  "composer",
  "container-author",
  "contributor",
  "curator",
  "director",
  "editor",
  "editorial-director",
  "executive-producer",
  "guest",
  "host",
  "illustrator",
  "interviewer",
  "narrator",
  "original-author",
  "performer",
  "producer",
  "recipient",
  "reviewed-author",
  "script-writer",
  "series-creator",
  "translator",
]);

const ARRAY_FIELDS = new Set(["categories", "keyword"]);
/** Prototype-poisoning names, barred whatever else a field would otherwise pass. */
const RESERVED_FIELDS = ["__proto__", "prototype", "constructor"];
const NUMBER_FIELDS = new Set([
  "chapter-number",
  "citation-number",
  "collection-number",
  "edition",
  "first-reference-note-number",
  "issue",
  "number",
  "number-of-pages",
  "number-of-volumes",
  "page",
  "page-first",
  "part-number",
  "printing-number",
  "section",
  "supplement-number",
  "version",
  "volume",
  "year-suffix",
]);
const TEXT_FIELDS = new Set([
  "DOI",
  "ISBN",
  "ISSN",
  "PMCID",
  "PMID",
  "URL",
  "abstract",
  "annote",
  "archive",
  "archive-place",
  "archive_location",
  "authority",
  "call-number",
  "citation-label",
  "collection-title",
  "container-title",
  "container-title-short",
  "dimensions",
  "event",
  "event-place",
  "genre",
  "jurisdiction",
  "language",
  "medium",
  "note",
  "original-publisher",
  "original-publisher-place",
  "original-title",
  "publisher",
  "publisher-place",
  "references",
  "reviewed-genre",
  "reviewed-title",
  "scale",
  "source",
  "status",
  "title",
  "title-short",
  "type",
]);
const ITEM_FIELD = /^[A-Za-z][A-Za-z0-9_-]{0,63}$/;
/** The shape `styleRevision` writes: the style digest, then the locale's. */
const RENDERER_REVISION = /^[0-9a-f]{64}:[0-9a-f]{64}$/;

export interface PortableSnapshot {
  version: typeof SNAPSHOT_VERSION;
  styleId: StyleId;
  /** Hashes of the exact CSL style and locale used to render this snapshot. */
  rendererRevision: string;
  items: Record<string, CslItem>;
}

export type SnapshotReadResult =
  | { status: "none" }
  | { status: "valid"; snapshot: PortableSnapshot }
  | {
      status: "style-mismatch";
      snapshot: PortableSnapshot;
      error: string;
    }
  /** Written by a NEWER build: never overwrite it, or data this build cannot read is destroyed. */
  | { status: "unsupported"; error: string }
  | { status: "invalid"; error: string };

/** What the document says its own citation style is: "unreadable" means a snapshot IS embedded but unreadable, so no render may proceed on the pane's default. */
export type DocumentStyleClaim =
  | { kind: "unclaimed" }
  | { kind: "claimed"; styleId: StyleId }
  | { kind: "unreadable" };

export function documentStyleClaim(
  result: SnapshotReadResult,
): DocumentStyleClaim {
  switch (result.status) {
    case "none":
      return { kind: "unclaimed" };
    case "valid":
    case "style-mismatch":
      // The records may be held back, but the style is the document's own data and survives a revision bump.
      return { kind: "claimed", styleId: result.snapshot.styleId };
    case "unsupported":
    case "invalid":
      // The styleId is deliberately NOT recovered from the payload: parsing can fail before it's ever read, so the refusal must be unconditional.
      return { kind: "unreadable" };
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/** Reject an entry with a message reading `snapshot entry "x" has an invalid title`. */
function reject(key: string, what: string): never {
  throw new Error(`snapshot entry "${key}" has ${what}`);
}

function assertDepth(key: string, value: unknown, depth = 0): void {
  if (depth > MAX_ITEM_DEPTH)
    reject(key, `nesting deeper than ${MAX_ITEM_DEPTH} levels`);
  if (Array.isArray(value)) {
    for (const element of value) assertDepth(key, element, depth + 1);
  } else if (isRecord(value)) {
    for (const element of Object.values(value))
      assertDepth(key, element, depth + 1);
  }
}

const NAME_TEXT_PARTS = new Set([
  "family",
  "given",
  "literal",
  "suffix",
  "non-dropping-particle",
  "dropping-particle",
]);
const NAME_BOOLEAN_PARTS = new Set([
  "comma-suffix",
  "static-ordering",
  "parse-names",
]);

const DATE_TEXT_PARTS = ["raw", "literal", "season"];
const DATE_PARTS = ["date-parts", ...DATE_TEXT_PARTS, "circa"];

/** citeproc recurses per nested rich-text tag; a real title carries a handful, never near this. */
const MAX_FIELD_ANGLE_BRACKETS = 100;
/** Bounds one field's own cost separate from the whole-snapshot byte cap. Unlike the other two checks below, an unusually long field reads as a UX problem rather than an attack signal, so it is shortened rather than rejected. */
export const MAX_FIELD_LENGTH = 8 * 1024;

/** Shortens an over-length value and records `field` into `truncatedFields`, rather than rejecting it; the other two checks still reject, since a bracket bomb or hidden character is an attack signal, not merely a long value. */
function assertSafeString(
  key: string,
  field: string,
  value: string,
  truncatedFields: Set<string>,
): string {
  if (value.length > MAX_FIELD_LENGTH) {
    value = value.slice(0, MAX_FIELD_LENGTH);
    truncatedFields.add(field);
  }
  if ((value.match(/</g)?.length ?? 0) > MAX_FIELD_ANGLE_BRACKETS) {
    reject(
      key,
      `${field} with more than ${MAX_FIELD_ANGLE_BRACKETS} "<" characters`,
    );
  }
  if (hasUnsafeNonLayoutChars(value)) {
    reject(
      key,
      `${field} containing unsafe invisible or directional characters`,
    );
  }
  return value;
}

/** Each name part is checked on its own, so even a huge author list stays within MAX_FIELD_LENGTH per field. */
function validateNameList(
  key: string,
  field: string,
  value: unknown,
  truncatedFields: Set<string>,
): void {
  if (!Array.isArray(value)) reject(key, `invalid ${field} data`);
  for (const name of value) {
    if (!isRecord(name)) reject(key, `invalid ${field} data`);
    const fields = Object.keys(name);
    if (fields.length === 0) reject(key, `an empty ${field} name`);
    let named = false;
    for (const part of fields) {
      const candidate = name[part];
      if (NAME_TEXT_PARTS.has(part)) {
        if (typeof candidate !== "string" || !candidate)
          reject(key, `an invalid ${field} ${part}`);
        name[part] = assertSafeString(
          key,
          `${field} ${part}`,
          candidate,
          truncatedFields,
        );
        named = true;
      } else if (NAME_BOOLEAN_PARTS.has(part)) {
        if (typeof candidate !== "boolean")
          reject(key, `an invalid ${field} ${part}`);
      } else {
        reject(key, `unsupported ${field} name field ${JSON.stringify(part)}`);
      }
    }
    if (!named) reject(key, `a ${field} name without text`);
  }
}

function isDatePart(value: unknown): boolean {
  return typeof value === "number"
    ? Number.isFinite(value) && Number.isInteger(value)
    : typeof value === "string" && /^-?\d+$/.test(value);
}

function validateDate(
  key: string,
  field: string,
  value: unknown,
  truncatedFields: Set<string>,
): void {
  if (value === undefined) return;
  if (!isRecord(value)) reject(key, `an invalid ${field}`);

  const parts = value["date-parts"];
  if (parts !== undefined) {
    const wellFormed =
      Array.isArray(parts) &&
      parts.length >= 1 &&
      parts.length <= 2 &&
      parts.every(
        (part) =>
          Array.isArray(part) &&
          part.length >= 1 &&
          part.length <= 3 &&
          part.every(isDatePart),
      );
    if (!wellFormed) reject(key, `invalid ${field} date-parts`);
  }
  for (const text of DATE_TEXT_PARTS) {
    const candidate = value[text];
    if (
      candidate !== undefined &&
      typeof candidate !== "string" &&
      typeof candidate !== "number"
    ) {
      reject(key, `an invalid ${field} ${text}`);
    }
    if (typeof candidate === "string")
      value[text] = assertSafeString(
        key,
        `${field} ${text}`,
        candidate,
        truncatedFields,
      );
    if (typeof candidate === "number" && !Number.isFinite(candidate))
      reject(key, `an invalid ${field} ${text}`);
  }
  if (
    value.circa !== undefined &&
    typeof value.circa !== "boolean" &&
    value.circa !== 0 &&
    value.circa !== 1
  ) {
    reject(key, `an invalid ${field} circa flag`);
  }
  const unknown = Object.keys(value).filter(
    (part) => !DATE_PARTS.includes(part),
  );
  if (unknown.length > 0)
    reject(key, `unsupported ${field} field ${JSON.stringify(unknown[0])}`);
}

/** Bounds the RECORD, never the citation tag, so a document already carrying a reserved-looking key reports "unresolved" rather than crashing when the reviver drops the name. */
const RESERVED_RECORD_KEYS = new Set([
  ...Object.getOwnPropertyNames(Object.prototype),
  "prototype",
]);

/** Validates one CSL item before it can reach citeproc or a snapshot; `truncatedFields` collects the label of every field shortened for length, and a caller that only needs validation can ignore it. */
export function validateCslItem(
  key: string,
  value: unknown,
  truncatedFields: Set<string> = new Set(),
): CslItem {
  if (RESERVED_RECORD_KEYS.has(key)) {
    reject(
      key,
      "a key the JavaScript object model reserves, which neither the citation " +
        "processor nor the embedded-record reader can address",
    );
  }
  if (!isRecord(value) || value.id !== key)
    reject(key, "an invalid or mismatched id");
  assertDepth(key, value);

  for (const [field, candidate] of Object.entries(value)) {
    if (!ITEM_FIELD.test(field) || RESERVED_FIELDS.includes(field)) {
      reject(key, `an unsupported field name ${JSON.stringify(field)}`);
    }
    if (field === "id") {
      value.id = assertSafeString(
        key,
        "id",
        candidate as string,
        truncatedFields,
      );
      continue;
    }
    if (NAME_FIELDS.has(field)) {
      validateNameList(key, field, candidate, truncatedFields);
      continue;
    }
    if (DATE_FIELDS.has(field)) {
      validateDate(key, field, candidate, truncatedFields);
      continue;
    }
    if (ARRAY_FIELDS.has(field)) {
      if (
        !Array.isArray(candidate) ||
        !candidate.every((entry): entry is string => typeof entry === "string")
      ) {
        reject(key, `an invalid ${field}`);
      }
      candidate.forEach((entry, index) => {
        candidate[index] = assertSafeString(key, field, entry, truncatedFields);
      });
      continue;
    }
    if (TEXT_FIELDS.has(field)) {
      if (typeof candidate !== "string") reject(key, `an invalid ${field}`);
      value[field] = assertSafeString(key, field, candidate, truncatedFields);
      continue;
    }
    if (NUMBER_FIELDS.has(field)) {
      if (typeof candidate === "string") {
        value[field] = assertSafeString(key, field, candidate, truncatedFields);
      } else if (typeof candidate !== "number" || !Number.isFinite(candidate)) {
        reject(key, `an invalid ${field}`);
      }
      continue;
    }
    if (typeof candidate === "string") {
      value[field] = assertSafeString(key, field, candidate, truncatedFields);
    } else if (typeof candidate === "number") {
      if (!Number.isFinite(candidate)) reject(key, `an invalid ${field}`);
    } else if (typeof candidate !== "boolean") {
      // Unknown fields are forward-compatible only when scalar: a structure can be ignored by one citeproc path and crash another.
      reject(key, `an invalid structured ${field}`);
    }
  }
  return value as CslItem;
}

/** The complete field list of the one format this build reads and writes. */
const ROOT_FIELDS = ["version", "styleId", "rendererRevision", "items"];

/** Refuse root data this format does not define: dropping it on rewrite would destroy it. */
function assertKnownRootFields(decoded: Record<string, unknown>): void {
  const unknown = Object.keys(decoded).filter(
    (field) => !ROOT_FIELDS.includes(field),
  );
  if (unknown.length > 0) {
    throw new Error(
      `unsupported snapshot root field ${JSON.stringify(unknown[0])}`,
    );
  }
}

function readStyleId(decoded: Record<string, unknown>): StyleId {
  if (!isStyleId(decoded.styleId)) {
    throw new Error(
      `unsupported snapshot style ${JSON.stringify(decoded.styleId)}`,
    );
  }
  return decoded.styleId;
}

function validateItems(value: unknown): Record<string, CslItem> {
  if (!isRecord(value)) throw new Error("snapshot items are not an object");
  const entries = Object.entries(value);
  if (entries.length > MAX_SNAPSHOT_ENTRIES) {
    throw new Error(`snapshot contains too many entries (${entries.length})`);
  }

  const items: Record<string, CslItem> = Object.create(null) as Record<
    string,
    CslItem
  >;
  for (const [key, item] of entries) {
    validateCitationKeys([key]);
    items[key] = validateCslItem(key, item);
  }
  return items;
}

function utf8Size(text: string): number {
  if (text.length > MAX_SNAPSHOT_BYTES) return text.length;
  return new TextEncoder().encode(text).length;
}

// `styleId` is `string` because this is a trust boundary: typing it as StyleId would make the guard below unreachable to the compiler.
export function serializeSnapshot(
  styleId: string,
  items: Record<string, CslItem>,
  /** A caller that only REMOVES records passes the document's own, so a style-mismatch warning it is still owed is not cleared. */
  rendererRevision?: string,
): string {
  if (!isStyleId(styleId)) {
    throw new Error(
      `Cannot save unsupported citation style ${JSON.stringify(styleId)}.`,
    );
  }
  if (
    rendererRevision !== undefined &&
    !RENDERER_REVISION.test(rendererRevision)
  ) {
    throw new Error(
      `Cannot save an unrecognised renderer revision ${JSON.stringify(rendererRevision)}.`,
    );
  }
  const snapshot: PortableSnapshot = {
    version: SNAPSHOT_VERSION,
    styleId,
    rendererRevision: rendererRevision ?? styleRevision(styleId),
    items: validateItems(items),
  };
  const json = JSON.stringify(snapshot);
  const bytes = utf8Size(json);
  if (bytes > MAX_SNAPSHOT_BYTES) {
    throw new Error(
      `Cited metadata is too large to embed (${(bytes / 1024).toFixed(0)} KB; ` +
        `limit ${MAX_SNAPSHOT_BYTES / 1024} KB).`,
    );
  }
  return json;
}

export function parseSnapshot(raw: unknown): SnapshotReadResult {
  if (raw === undefined || raw === null || raw === "")
    return { status: "none" };
  if (typeof raw !== "string") {
    return {
      status: "invalid",
      error: `The embedded ${PRODUCT_NAME} snapshot is not text.`,
    };
  }
  if (utf8Size(raw) > MAX_SNAPSHOT_BYTES) {
    return {
      status: "invalid",
      error: `The embedded ${PRODUCT_NAME} snapshot exceeds the safety limit.`,
    };
  }

  try {
    // Untrusted bytes: the reader makes them safe to look at; the shape checks below say whether they're this add-in's snapshot.
    const decoded: unknown = parseJson(raw, {
      what: SNAPSHOT_WHAT,
      maxBytes: MAX_SNAPSHOT_BYTES,
      maxDepth: MAX_SNAPSHOT_DEPTH,
      maxKeys: MAX_SNAPSHOT_KEYS,
    });
    if (!isRecord(decoded)) throw new Error("snapshot root is not an object");

    // Reported distinctly so callers refuse to OVERWRITE data this build cannot read.
    if (
      typeof decoded.version === "number" &&
      Number.isInteger(decoded.version) &&
      decoded.version > SNAPSHOT_VERSION
    ) {
      return {
        status: "unsupported",
        error:
          `This document carries citation data in format version ${decoded.version}, which was ` +
          `written by a newer version of ${PRODUCT_NAME}. Update the add-in to read it.`,
      };
    }
    if (decoded.version !== SNAPSHOT_VERSION) {
      // An earlier build is named only where true (an integer below the current format), so the refusal is one a person can act on.
      const version: unknown = decoded.version;
      const older =
        typeof version === "number" &&
        Number.isInteger(version) &&
        version > 0 &&
        version < SNAPSHOT_VERSION
          ? version
          : null;
      throw new Error(
        version === undefined
          ? `it carries no ${PRODUCT_NAME} format version`
          : older !== null
            ? `it is format version ${older}, written by an earlier version of ${PRODUCT_NAME} that this build does not read`
            : `unsupported snapshot version ${JSON.stringify(version)}`,
      );
    }
    assertKnownRootFields(decoded);
    const styleId = readStyleId(decoded);
    if (
      typeof decoded.rendererRevision !== "string" ||
      !RENDERER_REVISION.test(decoded.rendererRevision)
    ) {
      throw new Error("snapshot renderer revision is invalid");
    }
    const snapshot: PortableSnapshot = {
      version: SNAPSHOT_VERSION,
      styleId,
      rendererRevision: decoded.rendererRevision,
      // No second clone pass: the reader's reviver runs bottom-up, so every object it returned was already rebuilt on a null prototype.
      items: validateItems(decoded.items),
    };
    // Records from another style or locale revision are still the document's data; only the rendering they describe is untrusted.
    if (decoded.rendererRevision !== styleRevision(styleId)) {
      return {
        status: "style-mismatch",
        snapshot,
        error:
          `This document's embedded citations use a different ${builtinStyle(styleId).label} ` +
          "style or locale revision than this build. Open the authoritative .bib file before refreshing.",
      };
    }
    return { status: "valid", snapshot };
  } catch (error) {
    // Already a sentence written for a person, and already names its subject.
    if (error instanceof SafeInputError) {
      return { status: "invalid", error: safeDisplayText(error.message) };
    }
    return {
      status: "invalid",
      error: `The embedded ${PRODUCT_NAME} snapshot is invalid: ${safeDisplayText((error as Error).message)}.`,
    };
  }
}
