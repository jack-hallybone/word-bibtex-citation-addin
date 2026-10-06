// A citation content control's tag holds the authoritative BibTeX keys; its visible text is derived data.
// No mock can drive Word: after changing this file or the pane, check by hand in a real Word host — move a citation into a footnote, refresh, reopen the pane, and confirm it reports removing nothing; insert a citation, type at once, refresh, and confirm the typing is still there, after the citation; insert a bibliography in APA and confirm every entry's second line is indented, the last entry's too, then in IEEE and confirm each number sits on its entry's first line with the text aligned beside it, in the document's own font; switch to AMA and confirm every citation number is superscript, then back to APA and confirm none is.

import { BibliographyResult, CitationEngine } from "./citationEngine";
import { PRODUCT_NAME } from "./productName";
import { builtinStyle } from "./styles";
import {
  parseSnapshot,
  serializeSnapshot,
  SnapshotReadResult,
} from "./snapshot";
import {
  BIBLIOGRAPHY_TAG,
  CITATION_TAG_PREFIX,
  Citation,
  CitationDetails,
  CslItem,
  SNAPSHOT_SETTING,
  citationTag,
  describeError,
  parseCitationTag,
} from "./types";

export interface RefreshReport {
  citationsUpdated: number;
  bibliographyUpdated: boolean;
  /** True when bibliography controls exist but nothing safe could be written, so they were deliberately left as they were. */
  bibliographySkipped: boolean;
  missingKeys: string[];
  /** Distinct cited keys the library defines, so a skipped bibliography can say whether anything resolved. */
  keysResolved: number;
}

export interface InsertCitationResult {
  missing: string[];
  /** Set when the citation was inserted but document-wide renumbering failed. */
  refreshError?: string;
}

export interface SnapshotSaveReport {
  /** Entries embedded from the currently loaded library. */
  embedded: number;
  /** Keys the current library cannot resolve, kept from the previous snapshot. */
  preserved: string[];
}

/** Every document read goes through here, NEVER `document.body.contentControls`: body alone makes a citation inside a footnote look deleted and removes its embedded record. */
const documentControls = (
  context: Word.RequestContext,
): Word.ContentControlCollection => context.document.contentControls;

/** Refuse insertion outside the document's main body, including header tables. */
async function assertMainBody(
  context: Word.RequestContext,
  range: Word.Range,
): Promise<void> {
  let body = range.parentBody;
  for (let depth = 0; depth < 8; depth++) {
    body.load("type");
    await context.sync();
    if (body.type === "MainDoc") return;
    if (body.type !== "TableCell") break;

    const parent = body.parentBodyOrNullObject;
    parent.load("isNullObject,type");
    await context.sync();
    if (parent.isNullObject) break;
    body = parent;
  }
  throw new Error(
    `${PRODUCT_NAME} citations and bibliographies can only be inserted in the main document body.`,
  );
}

async function assertNotNestedInCitationControl(
  context: Word.RequestContext,
  range: Word.Range,
): Promise<void> {
  const parent = range.parentContentControlOrNullObject;
  parent.load("isNullObject,tag");
  await context.sync();
  if (
    !parent.isNullObject &&
    typeof parent.tag === "string" &&
    (parent.tag.startsWith(CITATION_TAG_PREFIX) ||
      parent.tag === BIBLIOGRAPHY_TAG)
  ) {
    throw new Error(
      `Cursor is inside an existing ${PRODUCT_NAME} citation or bibliography. Click outside it before inserting.`,
    );
  }
}

const title = (
  kind: "Citation" | "Bibliography",
  missing: readonly string[],
): string => (missing.length > 0 ? `${kind} (unresolved)` : kind);

const locate = (index: number): string =>
  `${PRODUCT_NAME} control #${index + 1}`;

function collectCitations(controls: readonly Word.ContentControl[]): {
  controls: Word.ContentControl[];
  citations: Citation[];
  bibliographyControls: Word.ContentControl[];
} {
  const citationControls: Word.ContentControl[] = [];
  const citations: Citation[] = [];
  const bibliographyControls: Word.ContentControl[] = [];

  // Every tag is validated before the caller queues any write, so a corrupt one cannot leave the document half-updated.
  controls.forEach((control, index) => {
    if (control.tag === BIBLIOGRAPHY_TAG) {
      bibliographyControls.push(control);
      return;
    }
    let citation: Citation | null;
    try {
      citation = parseCitationTag(control.tag);
    } catch (error) {
      throw new Error(`${(error as Error).message} Found in ${locate(index)}.`);
    }
    if (citation === null) return;
    citationControls.push(control);
    citations.push(citation);
  });
  return { controls: citationControls, citations, bibliographyControls };
}

/** Checked before any write is queued: Office.js has no rollback, so a lock found only mid-batch would leave a partial change. */
function assertEditable(controls: readonly Word.ContentControl[]): void {
  const locked = controls.findIndex((control) => control.cannotEdit === true);
  if (locked >= 0) {
    throw new Error(
      `${locate(locked)} is locked against editing. Unlock it before refreshing; nothing was changed.`,
    );
  }
}

interface ControlBackup {
  id: number;
  tag: string;
  title: string;
  text: string;
  /** Null where Word reports the control's text as partly superscript. */
  superscript: boolean | null;
}

interface InsertionMarker {
  pendingTag: string;
  finalTag: string;
  id: number | null;
  created: boolean;
}

function insertionMarker(finalTag: string): InsertionMarker {
  const words = new Uint32Array(4);
  globalThis.crypto.getRandomValues(words);
  const token = Array.from(words, (word) =>
    word.toString(16).padStart(8, "0"),
  ).join("");
  return {
    pendingTag: `bibcite-pending-${token}`,
    finalTag,
    id: null,
    created: false,
  };
}

/** Resolves where a new control may go, refusing every disallowed location; this syncs, which is what makes a caller's queued `load()` readable afterwards. */
async function prepareInsertionPoint(
  context: Word.RequestContext,
): Promise<Word.Range> {
  const point = context.document
    .getSelection()
    .getRange(Word.RangeLocation.end);
  await assertMainBody(context, point);
  await assertNotNestedInCitationControl(context, point);
  return point;
}

/** A citation's own text is always plain; only the bibliography carries the style's formatting. */
type ControlContent =
  | { kind: "text"; title: string; text: string; superscript: boolean }
  | { kind: "bibliography"; title: string; bibliography: BibliographyResult };

/** Plain text cannot carry a style's superscript citations (AMA, Nature), so Word sets the whole control in it; any other style sets it back. */
function writeCitation(
  control: Word.ContentControl,
  text: string,
  superscript: boolean,
): void {
  control.insertText(text, Word.InsertLocation.replace);
  control.font.superscript = superscript;
}

/** As Office Open XML, so each entry's layout is exact (see ooxml.ts). Word does not insert a package's last paragraph mark: that paragraph's text joins the paragraph it lands in, and takes its layout, so the last entry is given the style's indent directly too. */
function writeBibliography(
  control: Word.ContentControl,
  bibliography: BibliographyResult,
): void {
  control.insertOoxml(bibliography.ooxml, Word.InsertLocation.replace);
  const last = control.paragraphs.getLast();
  last.leftIndent = bibliography.indent;
  last.firstLineIndent = -bibliography.indent;
}

/** Creates a new citation control and fills it in, under a random pending tag until Word yields an id so `removeFailedInsertion` can find it. */
async function insertTaggedControl(
  context: Word.RequestContext,
  point: Word.Range,
  marker: InsertionMarker,
  content: ControlContent,
): Promise<Word.ContentControl> {
  const control = point.insertContentControl();
  marker.created = true;
  control.tag = marker.pendingTag;
  control.title = `${PRODUCT_NAME} insertion pending`;
  control.appearance = Word.ContentControlAppearance.boundingBox;
  control.load("id");
  await context.sync();
  marker.id = control.id;

  control.tag = marker.finalTag;
  control.title = content.title;
  if (content.kind === "bibliography") {
    writeBibliography(control, content.bibliography);
  } else {
    writeCitation(control, content.text, content.superscript);
  }
  return control;
}

async function captureControls(
  context: Word.RequestContext,
  controls: readonly Word.ContentControl[],
): Promise<ControlBackup[]> {
  if (controls.length === 0) return [];
  for (const control of controls) control.load("text,title,font/superscript");
  await context.sync();
  return controls.map((control) => ({
    id: control.id,
    tag: control.tag,
    title: control.title,
    text: control.text,
    superscript: control.font.superscript,
  }));
}

/** Best-effort compensation if Word reports failure after accepting a batch. */
async function restoreControls(
  backups: readonly ControlBackup[],
): Promise<void> {
  await Word.run(async (context) => {
    const controls = backups.map((backup) =>
      documentControls(context).getByIdOrNullObject(backup.id),
    );
    for (const control of controls) control.load("isNullObject,tag,cannotEdit");
    await context.sync();

    controls.forEach((control, index) => {
      const backup = backups[index];
      // Pairing is asserted, not assumed: restoring one control's text onto another is the corruption this function exists to undo.
      if (backup === undefined)
        throw new Error("lost track of which control a backup belongs to");
      if (
        control.isNullObject ||
        control.tag !== backup.tag ||
        control.cannotEdit
      ) {
        throw new Error(
          `could not safely identify or edit control id ${backup.id}`,
        );
      }
      control.title = backup.title;
      control.insertText(backup.text, Word.InsertLocation.replace);
      if (backup.superscript !== null)
        control.font.superscript = backup.superscript;
    });
    await context.sync();
  });
}

async function throwAfterRollback(
  error: unknown,
  backups: readonly ControlBackup[],
): Promise<never> {
  const detail = describeError(error);
  try {
    await restoreControls(backups);
  } catch (rollbackError) {
    throw new Error(
      `Word rejected a ${PRODUCT_NAME} update (${detail}), and automatic recovery also failed ` +
        `(${describeError(rollbackError)}). ` +
        `The document may be partially updated; use Word Undo and review every ${PRODUCT_NAME} control.`,
    );
  }
  throw new Error(
    `Word rejected a ${PRODUCT_NAME} update (${detail}). The previous ${PRODUCT_NAME} text was restored.`,
  );
}

async function removeFailedInsertion(marker: InsertionMarker): Promise<void> {
  await Word.run(async (context) => {
    const controls = documentControls(context);
    controls.load("items/id,items/tag,items/cannotEdit");
    await context.sync();
    const matches = controls.items.filter(
      (control) =>
        control.tag === marker.pendingTag ||
        (marker.id !== null &&
          control.id === marker.id &&
          control.tag === marker.finalTag),
    );
    if (matches.length === 0) return;
    const [pending] = matches;
    if (matches.length !== 1 || pending === undefined || pending.cannotEdit) {
      throw new Error(
        "could not safely identify or remove the pending content control",
      );
    }
    pending.delete(false);
    await context.sync();
  });
}

async function throwAfterInsertionFailure(
  error: unknown,
  marker: InsertionMarker,
): Promise<never> {
  const detail = describeError(error);
  try {
    await removeFailedInsertion(marker);
  } catch (cleanupError) {
    throw new Error(
      `Word rejected a ${PRODUCT_NAME} insertion (${detail}), and automatic cleanup also failed ` +
        `(${describeError(cleanupError)}). ` +
        "A partial control may remain; use Word Undo and review the insertion point.",
    );
  }
  throw new Error(
    `Word rejected a ${PRODUCT_NAME} insertion (${detail}). Any pending content control was removed.`,
  );
}

/** Insert a citation at the end of the current selection without replacing it. */
export async function insertCitation(
  engine: CitationEngine,
  keys: readonly string[],
  details: CitationDetails = {},
): Promise<InsertCitationResult> {
  await engine.ensureReady();
  const tag = citationTag(keys, details);
  const { text, missing } = engine.formatCitation({ keys, ...details });
  const marker = insertionMarker(tag);

  try {
    await Word.run(async (context) => {
      const controls = documentControls(context);
      controls.load("items/id,items/tag,items/cannotEdit");
      const point = await prepareInsertionPoint(context);

      // The sync inside prepareInsertionPoint loaded all tags.
      const existing = collectCitations(controls.items);
      assertEditable([...existing.controls, ...existing.bibliographyControls]);

      const control = await insertTaggedControl(context, point, marker, {
        kind: "text",
        title: title("Citation", missing),
        text,
        superscript: builtinStyle(engine.styleId).superscript,
      });
      await context.sync();

      // Word can leave the cursor inside the new control, where whatever is typed next is replaced at the next refresh. Best effort, once the citation is committed: a cursor Word will not move is no reason to take the citation back out.
      try {
        control.getRange(Word.RangeLocation.after).select();
        await context.sync();
      } catch {
        // The pane already says to type outside a citation.
      }
    });
  } catch (error) {
    if (!marker.created) throw error;
    return await throwAfterInsertionFailure(error, marker);
  }

  // The citation is already committed, so a failure of the document-wide renumbering is partial success, not an insert failure.
  try {
    await refreshDocument(engine);
  } catch (error) {
    return { missing, refreshError: (error as Error).message };
  }
  return { missing };
}

/** Re-render every citation in the document as one prevalidated batch, with compensating rollback. */
export async function refreshDocument(
  engine: CitationEngine,
): Promise<RefreshReport> {
  await engine.ensureReady();
  let backups: ControlBackup[] = [];

  try {
    return await Word.run(async (context) => {
      const allControls = documentControls(context);
      allControls.load("items/id,items/tag,items/cannotEdit");
      await context.sync();

      const { controls, citations, bibliographyControls } = collectCitations(
        allControls.items,
      );
      const rendered = engine.formatDocument(citations);
      const { superscript } = builtinStyle(engine.styleId);
      const bibliographyTargets = rendered.bibliography.renderable
        ? bibliographyControls
        : [];
      const targets = [...controls, ...bibliographyTargets];
      assertEditable(targets);
      backups = await captureControls(context, targets);

      // A short or reordered result would write one citation's text into another's control, so the pairing is checked explicitly.
      if (rendered.citations.length !== controls.length) {
        throw new Error(
          "The citation processor returned the wrong number of results, so nothing was rewritten.",
        );
      }
      controls.forEach((control, index) => {
        const result = rendered.citations[index];
        if (result === undefined) {
          throw new Error(
            "The citation processor returned no result for a citation, so nothing was rewritten.",
          );
        }
        control.title = title("Citation", result.missing);
        writeCitation(control, result.text, superscript);
      });

      // An empty bibliography is a refusal, not a value: bibliographyTargets is empty unless renderable, so an existing reference list is never replaced with nothing.
      for (const control of bibliographyTargets) {
        control.title = title("Bibliography", rendered.bibliography.missing);
        writeBibliography(control, rendered.bibliography);
      }
      const bibliographyUpdated = bibliographyTargets.length > 0;
      const bibliographySkipped =
        bibliographyControls.length > 0 && !rendered.bibliography.renderable;

      await context.sync();
      return {
        citationsUpdated: controls.length,
        bibliographyUpdated,
        bibliographySkipped,
        missingKeys: rendered.bibliography.missing,
        keysResolved: rendered.bibliography.keysResolved,
      };
    });
  } catch (error) {
    if (backups.length === 0) throw error;
    return await throwAfterRollback(error, backups);
  }
}

/** Insert a bibliography, or refresh every existing bibliography copy. */
export async function insertBibliography(
  engine: CitationEngine,
): Promise<string[]> {
  await engine.ensureReady();
  let backups: ControlBackup[] = [];
  // Boxed: a plain `let` assigned inside the callback narrows to `null` by the time the catch below reads it.
  const insertionState: { marker: InsertionMarker | null } = { marker: null };

  try {
    return await Word.run(async (context) => {
      const allControls = documentControls(context);
      allControls.load("items/id,items/tag,items/cannotEdit");
      await context.sync();

      const { citations, bibliographyControls } = collectCitations(
        allControls.items,
      );
      if (citations.length === 0) {
        throw new Error(
          `Insert at least one ${PRODUCT_NAME} citation before adding a bibliography.`,
        );
      }
      const bibliography = engine.formatDocument(citations).bibliography;
      if (!bibliography.renderable) {
        // Nothing to list: no key resolved, or each one that did is to a work the style cites in the text only.
        const { label } = builtinStyle(engine.styleId);
        throw new Error(
          bibliography.keysResolved === 0
            ? "None of this document's citations resolve against the loaded library, so there is " +
                `nothing to list (unresolved: ${bibliography.missing.join(", ")}). ` +
                "Open the .bib that defines them and try again."
            : bibliography.missing.length === 0
              ? `Every citation in this document is to a work ${label} cites in the text only, such as a ` +
                "personal communication, so there is nothing to list."
              : `Every citation that resolves is to a work ${label} cites in the text only, such as a ` +
                `personal communication, so there is nothing to list (unresolved: ${bibliography.missing.join(", ")}). ` +
                "Open the .bib that defines the unresolved keys and try again.",
        );
      }

      if (bibliographyControls.length > 0) {
        assertEditable(bibliographyControls);
        backups = await captureControls(context, bibliographyControls);
        for (const control of bibliographyControls) {
          control.title = title("Bibliography", bibliography.missing);
          writeBibliography(control, bibliography);
        }
      } else {
        const point = await prepareInsertionPoint(context);
        const pending = insertionMarker(BIBLIOGRAPHY_TAG);
        insertionState.marker = pending;
        await insertTaggedControl(context, point, pending, {
          kind: "bibliography",
          title: title("Bibliography", bibliography.missing),
          bibliography,
        });
      }

      await context.sync();
      return bibliography.missing;
    });
  } catch (error) {
    if (backups.length > 0) return await throwAfterRollback(error, backups);
    const failedMarker = insertionState.marker;
    if (failedMarker?.created)
      return await throwAfterInsertionFailure(error, failedMarker);
    throw error;
  }
}

async function citedKeysInDocument(): Promise<string[]> {
  return await Word.run(async (context) => {
    const allControls = documentControls(context);
    allControls.load("items/tag");
    await context.sync();
    const { citations } = collectCitations(allControls.items);
    return Array.from(new Set(citations.flatMap((citation) => citation.keys)));
  });
}

/** Every control's tag, unparsed: parsing refuses a tag this build cannot read, and the pane only asks whether any citation is there. */
export async function documentControlTags(): Promise<string[]> {
  return await Word.run(async (context) => {
    const allControls = documentControls(context);
    allControls.load("items/tag");
    await context.sync();
    return allControls.items.map((control) => control.tag);
  });
}

/** Replaces this pane's in-memory copy of the settings with what the document holds, which another pane on it (co-authoring) may have saved since this one opened. False, with the copy left as it was, where the host has no refreshAsync (Word on Mac and iPad) or the refresh fails. */
async function refreshSettings(): Promise<boolean> {
  const settings = Office.context.document.settings;
  if (typeof settings.refreshAsync !== "function") return false;
  return await new Promise<boolean>((resolve) => {
    try {
      settings.refreshAsync((result) =>
        resolve(result.status === Office.AsyncResultStatus.Succeeded),
      );
    } catch {
      resolve(false);
    }
  });
}

/** Null-prototype, so a key out of a content-control tag cannot find an inherited member: as a `{}`, `carried["toString"]` is a function. */
function readableRecords(result: SnapshotReadResult): Record<string, CslItem> {
  return result.status === "valid" || result.status === "style-mismatch"
    ? result.snapshot.items
    : (Object.create(null) as Record<string, CslItem>);
}

/** Commit snapshot bytes into the document's own settings. */
async function writeSnapshotSetting(json: string): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    try {
      Office.context.document.settings.set(SNAPSHOT_SETTING, json);
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }
        const detail =
          result.error?.message || "Word did not provide an error message";
        reject(
          new Error(
            `Word could not save the embedded citation data: ${detail}`,
          ),
        );
      });
    } catch (error) {
      reject(error instanceof Error ? error : new Error(describeError(error)));
    }
  });
}

/** Persists metadata for cited entries only, plus the selected style, keeping the .docx portable without embedding the whole library. */
export async function saveSnapshot(
  engine: CitationEngine,
): Promise<SnapshotSaveReport> {
  // Read before the refresh, which discards what only this copy holds: a record whose save failed.
  const local = readSnapshot();
  // A co-author's pane may have saved since this one opened; a save from the stale copy would drop its records, or write over its newer-format snapshot unseen.
  const existing = (await refreshSettings()) ? readSnapshot() : local;
  // "invalid" is blocked exactly like "unsupported": either way this build can't read what's embedded, so writing over it would silently drop every record it held.
  if (existing.status === "unsupported" || existing.status === "invalid") {
    throw new Error(`${existing.error} Nothing was overwritten.`);
  }

  // The refreshed record first, so this pane's older copy of a record never replaces a co-author's newer one.
  const carried = readableRecords(existing);
  const unsaved = readableRecords(local);
  const carriedRecord = (key: string): CslItem | undefined =>
    carried[key] ?? unsaved[key];

  // The embedded record set is DERIVED from the citations the document holds, with no exception for the empty set, so a deleted citation removes its metadata too.
  const keys = await citedKeysInDocument();

  const items = engine.snapshot(keys);
  const embedded = Object.keys(items).length;
  const preserved = keys.filter(
    (key) => items[key] === undefined && carriedRecord(key) !== undefined,
  );
  for (const key of preserved) {
    const record = carriedRecord(key);
    if (record !== undefined) items[key] = record;
  }

  await writeSnapshotSetting(serializeSnapshot(engine.styleId, items));
  return { embedded, preserved };
}

/** Removes embedded records for citations the document no longer holds, re-stamping styleId/rendererRevision so a style-mismatch warning still reports on the next open. */
export async function reconcileSnapshot(): Promise<string[]> {
  const existing = readSnapshot();
  if (existing.status !== "valid" && existing.status !== "style-mismatch") {
    return [];
  }
  const { styleId, rendererRevision, items } = existing.snapshot;
  const cited = new Set(await citedKeysInDocument());
  const removed = Object.keys(items).filter((key) => !cited.has(key));
  if (removed.length === 0) return [];

  const kept: Record<string, CslItem> = Object.create(null) as Record<
    string,
    CslItem
  >;
  for (const [key, record] of Object.entries(items)) {
    if (cited.has(key)) kept[key] = record;
  }
  await writeSnapshotSetting(
    serializeSnapshot(styleId, kept, rendererRevision),
  );
  return removed;
}

export function readSnapshot(): SnapshotReadResult {
  return parseSnapshot(Office.context.document.settings.get(SNAPSHOT_SETTING));
}
