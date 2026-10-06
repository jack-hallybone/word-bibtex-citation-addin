import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readdir, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { runInNewContext } from "node:vm";

import { BibTypes } from "bibliojson";
import CSL from "citeproc";

import { parseJson } from "../src/safe-json.js";
import { readBoundedText } from "../src/lib/assets";
import { FilePickerBibSource } from "../src/lib/bibSource";
import { PRODUCT_NAME } from "../src/lib/productName";
import {
  CitationEngine,
  CitationInput,
  LoadReport,
  verifyBibliographyResult,
} from "../src/lib/citationEngine";
import { BUILTIN_STYLES, StyleId, isStyleId } from "../src/lib/styles";
import {
  DocumentStyleClaim,
  MAX_FIELD_LENGTH,
  SnapshotReadResult,
  documentStyleClaim,
  parseSnapshot,
  serializeSnapshot,
} from "../src/lib/snapshot";
import { stripPercentComments } from "../src/lib/sourceFidelity";
import {
  SNAPSHOT_SETTING,
  citationTag,
  hasUnsafeNonLayoutChars,
  parseCitationTag,
  safeDisplayText,
  stripCslRichText,
  validateCitationKeys,
} from "../src/lib/types";
import {
  insertBibliography,
  insertCitation,
  reconcileSnapshot,
  refreshDocument,
  saveSnapshot,
} from "../src/lib/wordDoc";
import * as wordDoc from "../src/lib/wordDoc";

// `path` is where the browser fetches it from, relative to dist/taskpane.html; the source lives one level up, under public/.
const loadAsset = (path: string): Promise<string> =>
  readFile(`public/${path}`, "utf8");

/** The display name as a regex fragment, off the build-time substitution the pane uses, so a rename need not reach the suite. */
const NAME = PRODUCT_NAME.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** node:test's timeout cannot interrupt synchronous code, but vm's watchdog can, so a parser that loops forever fails the test rather than hanging the suite. */
function loadWithin(engine: CitationEngine, source: string): LoadReport {
  return runInNewContext(
    "load()",
    { load: () => engine.load(source) },
    { timeout: 5_000 },
  ) as LoadReport;
}

/** `written()` is every `set()` the add-in made, not the bag's contents, so a value written and then overwritten is still observed. */
function installOffice(initial?: string): {
  stored: () => string | undefined;
  saves: () => number;
  written: () => readonly (readonly [string, string])[];
} {
  const store = new Map<string, string>();
  const writes: [string, string][] = [];
  let saves = 0;
  if (initial !== undefined) store.set("bibcite:snapshot", initial);
  (globalThis as any).Office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    context: {
      document: {
        settings: {
          get: (key: string) => store.get(key),
          set: (key: string, value: string) => {
            writes.push([key, value]);
            return store.set(key, value);
          },
          saveAsync: (callback: (result: unknown) => void) => {
            saves++;
            callback({ status: "succeeded" });
          },
        },
      },
    },
  };
  return {
    stored: () => store.get("bibcite:snapshot"),
    saves: () => saves,
    written: () => writes,
  };
}

/** Office's documented settings model, which `installOffice` leaves out: each pane on a document works on its own in-memory copy taken when it opened, `saveAsync` persists that whole copy, and `refreshAsync` re-reads what is persisted (co-authoring). */
function coauthoredDocument(initial?: string) {
  let persisted = new Map<string, string>();
  if (initial !== undefined) persisted.set(SNAPSHOT_SETTING, initial);

  const openPane = (refresh: "succeeds" | "fails" | "throws" = "succeeds") => {
    let memory = new Map(persisted);
    let failNextSave = false;
    const settings = {
      get: (key: string) => memory.get(key),
      set: (key: string, value: string) => memory.set(key, value),
      saveAsync: (callback: (result: unknown) => void) => {
        if (failNextSave) {
          failNextSave = false;
          callback({ status: "failed", error: { message: "simulated" } });
          return;
        }
        persisted = new Map(memory);
        callback({ status: "succeeded" });
      },
      refreshAsync: (callback: (result: unknown) => void) => {
        if (refresh === "throws") throw new Error("simulated refresh throw");
        if (refresh === "fails") {
          callback({ status: "failed", error: { message: "simulated" } });
          return;
        }
        memory = new Map(persisted);
        callback({ status: "succeeded" });
      },
    };
    return {
      settings,
      /** Makes this pane the one `wordDoc` talks to. */
      activate: () => {
        (globalThis as any).Office = {
          AsyncResultStatus: { Succeeded: "succeeded" },
          context: { document: { settings } },
        };
      },
      failNextSave: () => {
        failNextSave = true;
      },
    };
  };
  return { openPane, stored: () => persisted.get(SNAPSHOT_SETTING) };
}

/** The envelope comes from the writer, so `items` can carry a malformed record without a hand-rolled version or rendererRevision passing vacuously. */
function embeddedSnapshot(
  items: unknown,
  styleId: "apa" | "ieee" = "apa",
  overrides: Record<string, unknown> = {},
): string {
  const envelope = JSON.parse(serializeSnapshot(styleId, {})) as Record<
    string,
    unknown
  >;
  return JSON.stringify({ ...envelope, items, ...overrides });
}

let nextControlId = 1;

/** Word keeps a document in several stories and `document.body` is only the main one, so a control carries its story; `MainDoc` is where this add-in inserts. */
function makeWordControl(tag: string, text = "existing", story = "MainDoc") {
  return {
    id: nextControlId++,
    story,
    tag,
    title: "",
    text,
    cannotEdit: false,
    isNullObject: false,
    /** Where the last `getRange(location).select()` on this control put the cursor. */
    selected: null as string | null,
    /** The indents set on `paragraphs.getLast()`, in points. */
    lastParagraph: { leftIndent: 0, firstLineIndent: 0 },
    font: { superscript: false },
    get paragraphs() {
      const last = this.lastParagraph;
      return { getLast: () => last };
    },
    load() {},
    delete(_keepContent: boolean) {},
    getRange(location: string) {
      return {
        select: () => {
          this.selected = location;
        },
      };
    },
    insertText(next: string) {
      this.text = next;
    },
    // The mock is a content recorder, not a Word emulator: real Word lays out the package it is given, but tests here only check the Office Open XML wordDoc.ts sent.
    insertOoxml(next: string) {
      this.text = next;
    },
  };
}

/** The paragraphs of the Office Open XML package Word was given, each with its text as Word shows it: a tab as "\t", a line break as "\n". */
function ooxmlParagraphs(xml: string): { xml: string; text: string }[] {
  return (xml.match(/<w:p>.*?<\/w:p>/g) ?? []).map((paragraph) => ({
    xml: paragraph,
    text: [
      ...paragraph.matchAll(
        /<w:t xml:space="preserve">([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>/g,
      ),
    ]
      .map(([run, text]) =>
        run === "<w:tab/>" ? "\t" : run === "<w:br/>" ? "\n" : (text ?? ""),
      )
      .join("")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&"),
  }));
}

interface WordMockOptions {
  onRun?: (call: number) => void;
  onSync?: (call: number) => void;
  selectionBodyType?: "MainDoc" | "Header" | "Footer" | "TableCell";
}

/** Office.js only queues `load()`; the value arrives at the next `sync()` and reading it earlier throws PropertyNotLoaded, and once settled it stays settled. */
function createLoadGate(description: string) {
  let requested = false;
  let available = false;
  return {
    request: () => {
      requested = true;
    },
    settle: () => {
      if (requested) available = true;
    },
    read<T>(value: T): T {
      if (!available) {
        throw new Error(
          `PropertyNotLoaded: ${description} was read before load() and sync().`,
        );
      }
      return value;
    },
  };
}

function installWordItems(
  items: ReturnType<typeof makeWordControl>[],
  options: WordMockOptions = {},
) {
  let call = 0;
  let syncCall = 0;
  const gates: { settle: () => void }[] = [];
  const gate = (description: string) => {
    const created = createLoadGate(description);
    gates.push(created);
    return created;
  };

  const nullControl = () => ({
    ...makeWordControl(""),
    isNullObject: true,
  });

  const itemsGate = gate("contentControls.items");
  let backingItems = items;

  /** `document.contentControls` spans every story — body, headers, footers, textboxes — and is the collection wordDoc.ts must ask. */
  const controls = {
    load() {
      itemsGate.request();
    },
    get items() {
      return itemsGate.read(backingItems);
    },
    set items(next: ReturnType<typeof makeWordControl>[]) {
      backingItems = next;
    },
    getByIdOrNullObject(id: number) {
      // Returned unloaded in Word too; callers load() and sync() before reading.
      return backingItems.find((item) => item.id === id) ?? nullControl();
    },
  };

  /** `document.body.contentControls` is the main story alone, so a read here silently omits citations in footnotes, headers and textboxes; it is counted rather than refused so the failure stays the partial answer Word gives. */
  let bodyReads = 0;
  const mainStory = () => {
    bodyReads++;
    return backingItems.filter((item) => item.story === "MainDoc");
  };
  const bodyControls = {
    load() {
      itemsGate.request();
    },
    get items() {
      return itemsGate.read(mainStory());
    },
    getByIdOrNullObject(id: number) {
      return mainStory().find((item) => item.id === id) ?? nullControl();
    },
  };

  const attachDelete = (control: ReturnType<typeof makeWordControl>) => {
    control.delete = () => {
      backingItems = backingItems.filter((candidate) => candidate !== control);
    };
  };
  for (const control of backingItems) attachDelete(control);

  const bodyTypeGate = gate("body.type");
  const parentBodyGate = gate("parentBodyOrNullObject");
  const rootBody = {
    load() {
      bodyTypeGate.request();
    },
    get type() {
      return bodyTypeGate.read(options.selectionBodyType ?? "MainDoc");
    },
    isNullObject: false,
    parentBodyOrNullObject: {
      load() {
        parentBodyGate.request();
      },
      get type() {
        return parentBodyGate.read("MainDoc");
      },
      get isNullObject() {
        return parentBodyGate.read(options.selectionBodyType !== "TableCell");
      },
    },
  };

  const parentControlGate = gate("parentContentControlOrNullObject");
  const context: any = {
    document: {
      body: { contentControls: bodyControls },
      contentControls: controls,
      getSelection: () => ({
        getRange: () => ({
          parentBody: rootBody,
          parentContentControlOrNullObject: {
            load() {
              parentControlGate.request();
            },
            get isNullObject() {
              return parentControlGate.read(true);
            },
            get tag() {
              return parentControlGate.read("");
            },
          },
          insertContentControl: () => {
            const control = makeWordControl("");
            attachDelete(control);
            backingItems.push(control);
            return control;
          },
        }),
      }),
    },
    sync(): Promise<void> {
      try {
        options.onSync?.(syncCall++);
      } catch (error) {
        // A real host reports a failed batch as a rejected promise.
        return Promise.reject(
          error instanceof Error ? error : new Error(String(error)),
        );
      }
      for (const pending of gates) pending.settle();
      return Promise.resolve();
    },
  };
  (globalThis as any).Word = {
    run: async (callback: (value: any) => Promise<unknown>) => {
      options.onRun?.(call++);
      return await callback(context);
    },
    InsertLocation: { replace: "replace" },
    RangeLocation: { end: "end", after: "after" },
    ContentControlAppearance: { boundingBox: "boundingBox" },
  };
  return { controls, bodyControls, rootBody, bodyReads: () => bodyReads };
}

function installWord(tags: string[], onRun?: (call: number) => void): void {
  installWordItems(
    tags.map((tag) => makeWordControl(tag)),
    onRun ? { onRun } : {},
  );
}

function clearHostMocks(): void {
  delete (globalThis as any).Word;
  delete (globalThis as any).Office;
}

test("citation tags round-trip and reject ambiguous or corrupt keys", () => {
  const tag = citationTag(["einstein1905", "knuth:1984"]);
  assert.equal(
    tag,
    "bibcite:einstein1905,knuth:1984",
    "plain citations keep the compact form",
  );
  assert.deepEqual(parseCitationTag(tag), {
    keys: ["einstein1905", "knuth:1984"],
  });
  assert.equal(parseCitationTag("unrelated"), null);
  assert.throws(() => citationTag([]), /at least one/);
  assert.throws(() => citationTag(["same", "same"]), /Duplicate/);
  assert.throws(() => citationTag(["comma,key"]), /Unsupported/);
  assert.throws(() => citationTag(["safe\u202Egpj.exe"]), /Unsupported/);
  assert.throws(() => citationTag(["x".repeat(200)]), /Unsupported/);
  assert.throws(() => parseCitationTag("bibcite:good,,bad"), /Unsupported/);
});

test("a key cannot pad itself into another key's appearance", () => {
  // U+2800 BRAILLE PATTERN BLANK draws an empty cell but is neither whitespace nor default-ignorable, so it can render exactly like the key beside it.
  assert.throws(() => citationTag(["smith2020\u2800"]), /Unsupported/);
  assert.throws(() => citationTag(["\u2800"]), /Unsupported/);
  assert.doesNotThrow(() => citationTag(["smith2020"]));

  // The enumerated ranges must hold on their own: \p{Default_Ignorable_Code_Point} is only as complete as the host webview's tables, which lag Node's.
  for (const character of [
    "\u0000", // NUL
    "\u009b", // C1 control
    "\u007f", // DEL
    "\u00ad", // soft hyphen
    "\u061c", // Arabic letter mark
    "\u200b", // zero-width space
    "\u200c", // zero-width non-joiner
    "\u200d", // zero-width joiner
    "\u200f", // right-to-left mark
    "\ufe0f", // variation selector-16 (emoji presentation)
    "\u202d", // left-to-right override
    "\u2060", // word joiner
    "\u2062", // invisible times
    "\u2069", // pop directional isolate
    "\ufeff", // byte-order mark
  ]) {
    assert.throws(
      () => citationTag([`smith2020${character}`]),
      /Unsupported/,
      `must reject U+${character.codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0")}`,
    );
  }

  // Display text is deliberately the opposite rule: a blank glyph in rendered prose is layout, so only a key bars it.
  assert.equal(hasUnsafeNonLayoutChars("a\u2800b"), false);
  assert.equal(safeDisplayText("a\u2800b"), "a\u2800b");
  assert.equal(hasUnsafeNonLayoutChars("a\u200bb"), true);
  assert.equal(safeDisplayText("a\u00adb"), "a\ufffdb");

  // ZWNJ/ZWJ and variation selectors are the text check's one further exemption (Persian/Urdu/Kurdish joining, an emoji's presentation), still barred from a key just above.
  assert.equal(
    hasUnsafeNonLayoutChars("\u0645\u06cc\u200c\u062e\u0648\u0627\u0647\u0645"), // a genuine Persian word containing U+200C
    false,
  );
  assert.equal(hasUnsafeNonLayoutChars("\u2764\ufe0f"), false); // heart + VS16
  // The exemption is narrow: neighbouring invisible characters it does NOT cover still trip the check.
  assert.equal(hasUnsafeNonLayoutChars("a\u200cb\u200bc"), true);
});

/** bibliojson writes citeproc-js's own rich-text markup into CSL-JSON fields (e.g. `{TeX}` becomes a nocase span); the plain preview list must show the text, not the markup. */
test("stripCslRichText strips citeproc's own tags, not arbitrary text", () => {
  assert.equal(
    stripCslRichText('A Study of <span class="nocase">TeX</span> Systems'),
    "A Study of TeX Systems",
  );
  assert.equal(
    stripCslRichText(
      '<span style="font-variant:small-caps;">Small</span> and <b>Bold</b> and <i>Italic</i>',
    ),
    "Small and Bold and Italic",
  );
  // Not a general HTML stripper: an untagged "<" in real title text survives.
  assert.equal(stripCslRichText("A < B: a comparison"), "A < B: a comparison");
});

test("the shared citation-key contract holds", async () => {
  // tests/citekey-conformance.json is a shared contract with lit-downloader; this add-in rejects every case except reserved names, which bound the RECORD (RESERVED_RECORD_KEYS), never the TAG a citation merely points at.
  const contract = JSON.parse(
    await readFile("tests/citekey-conformance.json", "utf8"),
  ) as {
    accept: { key: string; why: string }[];
    reject: {
      key: string;
      mode: "refuse" | "terminate";
      reads?: string;
      why: string;
    }[];
  };
  assert.ok(contract.accept.length > 10 && contract.reject.length > 10);
  const reservedRecordKeys = new Set([
    ...Object.getOwnPropertyNames(Object.prototype),
    "prototype",
  ]);
  for (const { key, why } of contract.accept) {
    assert.deepEqual(
      parseCitationTag(citationTag([key])),
      { keys: [key] },
      `must accept ${JSON.stringify(key)} — ${why}`,
    );
  }
  for (const { key, why } of contract.reject) {
    if (reservedRecordKeys.has(key)) continue;
    assert.throws(
      () => citationTag([key]),
      /Unsupported/,
      `must reject ${JSON.stringify(key)} — ${why}`,
    );
  }
});

test("a citation written with a pinpoint is refused, never quietly stripped", () => {
  // Every version number is refused, not just the one an earlier build wrote: the number is the payload's claim, not a fact about it.
  for (const version of [1, 2, 9, 47]) {
    assert.throws(
      () =>
        parseCitationTag(
          `bibcite:${version};{"k":"doe2020","l":"5","b":"page"}`,
        ),
      new RegExp(
        `format version ${version}.*pinpoint.*Nothing has been changed`,
        "s",
      ),
      `accepted version ${version}`,
    );
  }
  // A versioned payload carrying no pinpoint is refused the same way rather than reduced to its key.
  assert.throws(
    () => parseCitationTag('bibcite:2;{"k":"doe2020"}'),
    /pinpoint.*Nothing has been changed/s,
  );

  // `;` is a legal citation key character, so a key that merely begins like a version marker is cited, not reported.
  assert.deepEqual(parseCitationTag("bibcite:2020;smith"), {
    keys: ["2020;smith"],
  });
});

test("a hostile citation tag is refused rather than parsed for what it holds", () => {
  const hostile: Array<[string, RegExp]> = [
    [`bibcite:${"a".repeat(250)}`, /Word-safe limit/],
    ["bibcite:", /empty reference tag/],
    ['bibcite:2;{"k":"a","l":"5"}', /pinpoint/],
    ['bibcite:{"k":"a"}', /Unsupported BibTeX key/],
    ["bibcite:a\u202Eb", /Unsupported BibTeX key/],
    ["bibcite:a,,b", /Unsupported BibTeX key/],
    ["bibcite:a,a", /Duplicate BibTeX key/],
    ["bibcite:smith2020\u2800", /Unsupported BibTeX key/],
  ];
  for (const [tag, expected] of hostile) {
    assert.throws(() => parseCitationTag(tag), expected, `accepted ${tag}`);
  }
});

test("a citation's page and form travel in its tag, after its keys", () => {
  const cases: [Parameters<typeof citationTag>[1], string][] = [
    [
      { pinpoint: { label: "page", locator: "12" } },
      "bibcite:smith2020 page 12",
    ],
    [{ form: "narrative" }, "bibcite:smith2020 narrative"],
    [
      { form: "suppress-author", pinpoint: { label: "chapter", locator: "3" } },
      "bibcite:smith2020 suppress-author chapter 3",
    ],
    // The locator comes last, so it may hold the spaces and commas a key may not.
    [
      { pinpoint: { label: "page", locator: "12–15, 18" } },
      "bibcite:smith2020 page 12–15, 18",
    ],
  ];
  for (const [details, tag] of cases) {
    assert.equal(citationTag(["smith2020"], details), tag);
    assert.deepEqual(parseCitationTag(tag), {
      keys: ["smith2020"],
      ...details,
    });
  }
  // A plain citation keeps the one form every earlier build reads.
  assert.equal(citationTag(["smith2020"], {}), "bibcite:smith2020");

  // An earlier build reads everything after the prefix as keys, and its key check (shared, unchanged) refuses a space: it stops with nothing rewritten instead of losing the page.
  assert.throws(
    () => validateCitationKeys(["smith2020 page 12"]),
    /Unsupported BibTeX key/,
  );
});

test("a page or form the tag cannot carry exactly is refused, never trimmed", () => {
  const refusedOnInsert: [Parameters<typeof citationTag>, RegExp][] = [
    [[["a", "b"], { form: "narrative" }], /one reference/],
    [
      [["a", "b"], { pinpoint: { label: "page", locator: "4" } }],
      /one reference/,
    ],
    [[["a"], { pinpoint: { label: "page", locator: "" } }], /cannot be empty/],
    [[["a"], { pinpoint: { label: "page", locator: " 12" } }], /one line/],
    [[["a"], { pinpoint: { label: "page", locator: "12\n13" } }], /one line/],
    [
      [["a"], { pinpoint: { label: "page", locator: "1".repeat(41) } }],
      /at most 40/,
    ],
    [
      [["a"], { pinpoint: { label: "page", locator: "12\u202E" } }],
      /invisible or directional/,
    ],
    [
      [["a"], { pinpoint: { label: "folio" as "page", locator: "2" } }],
      /Unsupported locator type/,
    ],
    [
      [["a"], { form: "author-only" as "narrative" }],
      /Unsupported citation form/,
    ],
    [
      [
        ["k".repeat(150)],
        { pinpoint: { label: "paragraph", locator: "9".repeat(40) } },
      ],
      /too long for Word/,
    ],
  ];
  for (const [args, expected] of refusedOnInsert) {
    assert.throws(() => citationTag(...args), expected, JSON.stringify(args));
  }

  // Read back, only what citationTag writes is accepted: a newer build's detail is refused, not dropped.
  for (const tag of [
    "bibcite:a ",
    "bibcite:a page",
    "bibcite:a page  12",
    "bibcite:a page 12 ",
    "bibcite:a narrative ",
    "bibcite:a narrative narrative",
    "bibcite:a narrative page",
    "bibcite:a 12",
    "bibcite:a,b page 12",
    "bibcite:a prefix see",
    "bibcite:a Page 12",
  ]) {
    assert.throws(
      () => parseCitationTag(tag),
      /details this build does not read.*nothing has been changed/s,
      `accepted ${JSON.stringify(tag)}`,
    );
  }
});

test("a journal article keeps its journal, volume, issue and pages", async () => {
  // The converter can resolve `@article` to CSL type "article" rather than "article-journal", dropping journal/volume/issue/pages, so this asserts through the render.
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  assert.equal(engine.item("einstein1905")?.type, "article-journal");

  const result = engine.formatDocument([{ keys: ["einstein1905"] }]);
  const entry = result.bibliography.html;
  assert.match(entry, /Annalen der [Pp]hysik/, "the journal name must survive");
  assert.match(entry, /322/, "the volume must survive");
  assert.match(entry, /\bno\. 10\b/, "the issue number must survive");
  assert.match(entry, /891/, "the page range must survive");
});

test("a journal, series or event keeps its name as written, while a work's own title is sentence-cased", async () => {
  // IEEE prints each name as stored, so a sentence-cased one read "The new england journal of medicine".
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(
    "@article{nejm, author={Doe, Jane}, title={A Trial of Things}, journal={The New England Journal of Medicine}, year={2019}}\n" +
      "@inproceedings{talk, author={Roe, Max}, title={A Talk on Things}, eventtitle={Big Conference on Things}, year={2018}}\n" +
      "@book{taocp, author={Knuth, Donald}, title={The Art of Computer Programming}, series={Addison-Wesley Series in Computer Science}, publisher={Addison-Wesley}, year={1968}}\n",
  );
  await engine.ensureReady();
  const { bibliography } = engine.formatDocument(
    ["nejm", "talk", "taocp"].map((key) => ({ keys: [key] })),
  );
  assert.deepEqual(
    ooxmlParagraphs(bibliography.ooxml).map((entry) => entry.text),
    [
      "[1]\tJ. Doe, “A trial of things,” The New England Journal of Medicine, 2019.",
      "[2]\tM. Roe, “A talk on things,” presented at the Big Conference on Things, 2018.",
      "[3]\tD. Knuth, The art of computer programming. in Addison-Wesley Series in Computer Science. Addison-Wesley, 1968.",
    ],
  );
});

test("a title's case never depends on the language of the entry before it", async () => {
  // The converter applied the last language an entry named (its langid) to every entry after it, so APA printed this English title in title case.
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(
    "@book{goethe, author={Goethe, Johann}, title={Die Leiden des jungen Werthers}, publisher={Weygand}, year={1774}, langid={german}}\n" +
      "@book{woolf, author={Woolf, Virginia}, title={A Room of One's Own}, publisher={Hogarth Press}, year={1929}}\n",
  );
  await engine.ensureReady();
  const { bibliography } = engine.formatDocument([
    { keys: ["goethe"] },
    { keys: ["woolf"] },
  ]);
  assert.deepEqual(
    ooxmlParagraphs(bibliography.ooxml).map((entry) => entry.text),
    [
      "Goethe, J. (1774). Die Leiden des jungen Werthers. Weygand.",
      "Woolf, V. (1929). A room of one’s own. Hogarth Press.",
    ],
  );
});

test("identifier fields read back as typed into the reference manager: its escapes undone, nothing else decoded", async () => {
  // Its TEX_ESCAPES (reference-manager/bib-core/src/bib-render.mjs), which it writes into these ten fields as well as into prose.
  const typed = "3 & 4 #5 $6 7% 8_9 x^2 ~10";
  const written = String.raw`3 \& 4 \#5 \$6 7\% 8\_9 x\textasciicircum{}2 \textasciitilde{}10`;
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(
    `@article{art, author={Doe, Jane}, title={T}, journal={J}, year={2020}, volume={${written}}, ` +
      `issue={${written}}, pages={${written}}, version={${written}}, part={${written}}}\n` +
      `@book{bk, author={Doe, Jane}, title={T}, publisher={P}, year={2020}, number={${written}}, ` +
      `chapter={${written}}, edition={${written}}, volumes={${written}}, pagetotal={${written}}}\n`,
  );
  // `part` has no CSL variable, so it is never shown.
  for (const [key, variables] of [
    ["art", ["volume", "issue", "page", "version"]],
    [
      "bk",
      [
        "collection-number",
        "chapter-number",
        "edition",
        "number-of-volumes",
        "number-of-pages",
      ],
    ],
  ] as const) {
    for (const variable of variables) {
      assert.equal(engine.item(key)?.[variable], typed, `${key} ${variable}`);
    }
  }

  await engine.ensureReady();
  const { html } = engine.formatDocument([
    { keys: ["art"] },
    { keys: ["bk"] },
  ]).bibliography;
  const shown = typed.replace(/&/g, "&#38;");
  assert.ok(
    html.includes(`vol. ${shown}, no. ${shown}, pp. ${shown}`),
    `the reference list shows ${JSON.stringify(typed)} as typed: ${html}`,
  );
  assert.doesNotMatch(html, /\\|textascii/);

  // Nothing else is decoded: a raw character, written before these fields were escaped or by another tool, stays as it is, and so does a -- outside a page range.
  engine.load(
    "@book{bk, title={T}, publisher={P}, year={2020}, number={WP_2020_05}, " +
      "volumes={3 & 4}, edition={x^2 $5 ~6}, volume={1--2}}",
  );
  const book = engine.item("bk");
  assert.equal(book?.["collection-number"], "WP_2020_05");
  assert.equal(book?.["number-of-volumes"], "3 & 4");
  assert.equal(book?.edition, "x^2 $5 ~6");
  assert.equal(book?.volume, "1--2");

  // A page range is still split at its dash, which the style then sets; each end reads back the same way.
  engine.load(
    String.raw`@article{art, title={T}, journal={J}, year={2020}, pages={A\_1--A_9}}`,
  );
  assert.equal(engine.item("art")?.page, "A_1-A_9");

  // Braces another tool put around a value, or part of one, are dropped, as the reference manager drops them when it shows these fields.
  engine.load(
    String.raw`@book{bk, title={T}, publisher={P}, year={2020}, volume={{II}}, number={{WP}\_2020}}`,
  );
  assert.equal(engine.item("bk")?.volume, "II");
  assert.equal(engine.item("bk")?.["collection-number"], "WP_2020");

  // A TeX command another tool wrote is still read as LaTeX.
  engine.load(
    String.raw`@book{bk, title={T}, publisher={P}, year={2020}, edition={2\textsuperscript{nd}}}`,
  );
  assert.equal(engine.item("bk")?.edition, "2<sup>nd</sup>");
});

test("BibTeX loading rejects parser errors and duplicate keys transactionally", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  const valid =
    "@article{stable, title={Stable}, author={Smith, Jane}, year={2020}}";
  assert.equal(engine.load(valid).entries, 1);
  assert.equal(engine.size, 1);

  assert.throws(
    () =>
      engine.load(
        "@article{same, title={First}, year={2020}}\n@article{same, title={Second}, year={2021}}",
      ),
    /Duplicate BibTeX key/,
  );
  assert.equal(engine.size, 1);
  assert.equal(engine.item("stable")?.title, "Stable");

  assert.throws(
    () =>
      engine.load(
        "@article{broken, title={Missing close}\n@book{next, title={Next}}",
      ),
    /BibTeX parsing failed/,
  );
  assert.equal(engine.size, 1);

  assert.throws(
    () =>
      engine.load(
        "@article{fine, title={Fine}}\n@article{culprit, title={T}, urldate={not a date}}",
      ),
    /entry "culprit" field "urldate"/,
  );
  assert.equal(engine.size, 1);
  await engine.ensureReady();
});

test("a url the parser's own check refuses, such as an intranet address, still loads and cites as written", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  const report = engine.load(
    "@online{wiki, author={Doe, Jane}, title={Team Notes}, year={2026}, url={http://wiki/page}, urldate={2026-10-05}}\n" +
      "@misc{lan, title={Lab Server}, url={http://192.168.1.10/docs}}\n" +
      "@misc{local, title={Local Notes}, url={http://localhost:8080/notes}}\n",
  );
  assert.equal(report.entries, 3);
  assert.equal(engine.item("wiki")?.URL, "http://wiki/page");
  assert.equal(engine.item("lan")?.URL, "http://192.168.1.10/docs");
  assert.equal(engine.item("local")?.URL, "http://localhost:8080/notes");
  assert.deepEqual(engine.item("wiki")?.accessed, {
    "date-parts": [[2026, 10, 5]],
  });

  await engine.ensureReady();
  const { html } = engine.formatDocument([
    { keys: ["wiki"] },
    { keys: ["lan"] },
    { keys: ["local"] },
  ]).bibliography;
  for (const url of [
    "http://wiki/page",
    "http://192.168.1.10/docs",
    "http://localhost:8080/notes",
  ]) {
    assert.ok(html.includes(url), `the reference list shows ${url}: ${html}`);
  }
  assert.doesNotMatch(html, /<a\b/);
});

test("an entry the parser cuts short inside a brace group refuses the file", () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load("@article{stable, title={Stable}, year={2020}}");

  // {"} is BibTeX's own way to put a quote in a quoted value; the parser ended the value at "The {" and dropped author, journal, year and volume without a word.
  const quotedQuote = [
    "@article{doe2020,",
    '  title   = "The {"}Big{"} Idea",',
    '  author  = "Doe, Jane and Roe, Richard",',
    '  journal = "Journal of Examples",',
    '  year    = "2020",',
    '  volume  = "12"',
    "}",
    "@article{roe2021, title={Second}, author={Roe, Richard}, journal={J}, year={2021}}",
    "",
  ].join("\n");
  assert.throws(
    () => engine.load(quotedQuote),
    /BibTeX parsing failed: entry "doe2020" ends inside an unclosed "\{".*No library data was replaced/,
  );
  assert.equal(engine.size, 1, "the refusal is transactional");
  assert.equal(engine.item("stable")?.title, "Stable");

  // In the last field nothing visibly goes missing but the rest of the value.
  assert.throws(
    () =>
      engine.load(
        '@article{a, author={Doe, J}, year={2020}, title = "He said {"}hi{"} twice"}\n@misc{b, title={B}}\n',
      ),
    /entry "a" ends inside an unclosed "\{"/,
  );

  // Braced, the same value parses whole, and a stray } in a quoted value leaves no brace open.
  engine.load('@article{a, title = {The {"}Big{"} Idea}, year={2020}}');
  assert.deepEqual(engine.item("a")?.issued, { "date-parts": [[2020]] });
  engine.load(
    '@article{k, note="unmatched } brace", title={A 50% improvement}, year={2020}}',
  );
  assert.equal(engine.item("k")?.title, "A 50% improvement");
});

test("document-wide IEEE formatting assigns stable global numbers", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  const result = engine.formatDocument([
    { keys: ["einstein1905"] },
    { keys: ["knuth1984texbook"] },
    { keys: ["einstein1905"] },
  ]);
  assert.equal(result.citations[0]?.text, "[1]");
  assert.equal(result.citations[1]?.text, "[2]");
  assert.equal(result.citations[2]?.text, "[1]");
  assert.match(result.bibliography.html, /csl-left-margin">\[1\]/);
  assert.match(result.bibliography.html, /csl-left-margin">\[2\]/);
  // Citations stay plain; only the reference list gets the style's own italics (IEEE italicises the journal/book title).
  assert.doesNotMatch(result.citations[0]?.text ?? "", /<i>/);
  assert.match(result.bibliography.html, /<i>Annalen der Physik<\/i>/);
  assert.match(result.bibliography.html, /<i>The TeXbook<\/i>/);
});

test("a page and a form render in each style's own way, and a numeric style keeps its number", async () => {
  const bib = await readFile("sample/library.bib", "utf8");
  const page = (locator: string, label: "page" | "chapter" = "page") => ({
    pinpoint: { label, locator },
  });
  const knuth = (details = {}) => ({ keys: ["knuth1984texbook"], ...details });
  const asked = [
    knuth(),
    knuth(page("12")),
    knuth(page("12-15")),
    knuth(page("3", "chapter")),
    knuth({ form: "narrative" }),
    knuth({ form: "narrative", ...page("12") }),
    knuth({ form: "suppress-author" }),
    knuth({ form: "suppress-author", ...page("12") }),
  ];
  // MLA cites no year: its narrative form is the name alone until a page follows it, and its page-only form needs a page, so without one it is the standard form. Nature and Vancouver print no page.
  const rendered: Record<StyleId, string[]> = {
    "american-medical-association": [
      "1",
      "1(p12)",
      "1(pp12-15)",
      "1(chap3)",
      "1",
      "1(p12)",
      "1",
      "1(p12)",
    ],
    apa: [
      "(Knuth, 1984)",
      "(Knuth, 1984, p. 12)",
      "(Knuth, 1984, pp. 12–15)",
      "(Knuth, 1984, Chapter 3)",
      "Knuth (1984)",
      "Knuth (1984, p. 12)",
      "(1984)",
      "(1984, p. 12)",
    ],
    "chicago-author-date": [
      "(Knuth 1984)",
      "(Knuth 1984, 12)",
      "(Knuth 1984, 12–15)",
      "(Knuth 1984, chap. 3)",
      "Knuth (1984)",
      "Knuth (1984, 12)",
      "(1984)",
      "(1984, 12)",
    ],
    "harvard-cite-them-right": [
      "(Knuth, 1984)",
      "(Knuth, 1984, p. 12)",
      "(Knuth, 1984, pp. 12–15)",
      "(Knuth, 1984, ch. 3)",
      "Knuth (1984)",
      "Knuth (1984, p. 12)",
      "(1984)",
      "(1984, p. 12)",
    ],
    ieee: [
      "[1]",
      "[1, p. 12]",
      "[1, pp. 12–15]",
      "[1, Ch. 3]",
      "[1]",
      "[1, p. 12]",
      "[1]",
      "[1, p. 12]",
    ],
    "modern-language-association": [
      "(Knuth)",
      "(Knuth 12)",
      "(Knuth 12–15)",
      "(Knuth, ch. 3)",
      "Knuth",
      "Knuth (12)",
      "(Knuth)",
      "(12)",
    ],
    nature: ["1", "1", "1", "1", "1", "1", "1", "1"],
    "nlm-citation-sequence": [
      "(1)",
      "(1)",
      "(1)",
      "(1)",
      "(1)",
      "(1)",
      "(1)",
      "(1)",
    ],
  };
  assert.deepEqual(
    BUILTIN_STYLES.map((style) => style.id),
    Object.keys(rendered),
    "every style the menu offers, and no other",
  );
  for (const style of BUILTIN_STYLES) {
    const engine = new CitationEngine(style.id, loadAsset);
    engine.load(bib);
    await engine.ensureReady();
    assert.deepEqual(
      engine.formatDocument(asked).citations.map((result) => result.text),
      rendered[style.id],
      style.id,
    );
  }

  // A document mixing citations with and without details numbers them as one: the form asks for nothing a number can show, so it is left at the number.
  const ieee = new CitationEngine("ieee", loadAsset);
  ieee.load(bib);
  await ieee.ensureReady();
  assert.deepEqual(
    ieee
      .formatDocument([
        { keys: ["einstein1905"] },
        knuth(page("12")),
        { keys: ["einstein1905"], form: "narrative" },
        knuth({ form: "suppress-author", ...page("3", "chapter") }),
      ])
      .citations.map((result) => result.text),
    ["[1]", "[2, p. 12]", "[1]", "[2, Ch. 3]"],
  );
});

test("the engine refuses details no tag could carry, however they reach it", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();
  // The tag's own check runs again here, so nothing it refuses reaches citeproc.
  assert.throws(
    () =>
      engine.formatDocument([
        {
          keys: ["einstein1905", "knuth1984texbook"],
          pinpoint: { label: "page", locator: "4" },
        },
      ]),
    /one reference/,
  );
});

test("the reference list goes to Word as a paragraph per entry, with APA's hanging indent and IEEE's number beside its text", async () => {
  const bib = await readFile("sample/library.bib", "utf8");
  const cited = [{ keys: ["einstein1905"] }, { keys: ["knuth1984texbook"] }];

  const apa = new CitationEngine("apa", loadAsset);
  apa.load(bib);
  await apa.ensureReady();
  const apaEntries = ooxmlParagraphs(
    apa.formatDocument(cited).bibliography.ooxml,
  );
  assert.deepEqual(
    apaEntries.map((entry) => entry.text),
    [
      "Einstein, A. (1905). Zur elektrodynamik bewegter körper. Annalen Der Physik, 322(10), 891–921. https://doi.org/10.1002/andp.19053221004",
      "Knuth, D. E. (1984). The TeXbook. Addison-Wesley.",
    ],
  );
  for (const entry of apaEntries) {
    // Half an inch, every line but the first.
    assert.match(
      entry.xml,
      /^<w:p><w:pPr><w:ind w:left="720" w:hanging="720"\/><\/w:pPr>/,
    );
  }

  const ieee = new CitationEngine("ieee", loadAsset);
  ieee.load(bib);
  await ieee.ensureReady();
  const ieeeEntries = ooxmlParagraphs(
    ieee.formatDocument(cited).bibliography.ooxml,
  );
  assert.equal(ieeeEntries.length, 2, "a number is not a paragraph of its own");
  for (const [index, entry] of ieeeEntries.entries()) {
    assert.ok(entry.text.startsWith(`[${index + 1}]\t`), entry.text);
    // The text's own lines start where the tab puts the first one: beside the number, not under it.
    assert.match(
      entry.xml,
      /^<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="480"\/><\/w:tabs><w:ind w:left="480" w:hanging="480"\/><\/w:pPr>/,
    );
  }
});

test("the reference list Word is given takes the document's own font and carries a hostile title as text", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(
    "@article{hostile, author={Doe, Jane}, " +
      'title={A <script>alert(1)</script> & "B" study}, journal={J \\& K}, year={2020}}',
  );
  await engine.ensureReady();
  const { ooxml } = engine.formatDocument([{ keys: ["hostile"] }]).bibliography;

  assert.match(
    ooxml,
    /^<pkg:package xmlns:pkg="http:\/\/schemas\.microsoft\.com\/office\/2006\/xmlPackage">/,
  );
  assert.doesNotMatch(ooxml, /<script|&#\d+;/);
  assert.match(
    ooxml,
    /A &lt;script&gt;alert\(1\)&lt;\/script&gt; &amp; “B” study/,
  );
  // No font, size or style of its own: an entry looks like the document around it.
  assert.doesNotMatch(ooxml, /<w:(?:rFonts|sz|pStyle|rStyle|spacing)\b/);
  // Every element it opens, it closes, in order.
  const open: string[] = [];
  for (const [, closing, name, selfClosing] of ooxml.matchAll(
    /<(\/?)([\w:]+)(?:\s[^>]*?)?(\/?)>/g,
  )) {
    if (selfClosing) continue;
    if (closing) assert.equal(open.pop(), name);
    else open.push(name!);
  }
  assert.deepEqual(open, []);
});

test("bibliography verification checks citeproc identity metadata", () => {
  const expected = new Set(["a", "b"]);
  const metadata = { entry_ids: [["a"], ["b"]], bibliography_errors: [] };
  assert.deepEqual(
    verifyBibliographyResult([metadata, [" First ", "Second\n"]], expected),
    ["First", "Second"],
  );

  for (const malformed of [
    [
      { entry_ids: [["a"], ["wrong"]], bibliography_errors: [] },
      ["First", "Second"],
    ],
    [{ entry_ids: [["a", "b"]], bibliography_errors: [] }, ["Combined"]],
    [
      { entry_ids: [["a"], ["b"]], bibliography_errors: ["failure"] },
      ["First", "Second"],
    ],
    [{ bibliography_errors: [] }, ["First", "Second"]],
  ]) {
    assert.throws(
      () => verifyBibliographyResult(malformed, expected),
      /citation processor/,
    );
  }
});

test("a work named as cited in the text only may be missing from the bibliography, and nothing else may", () => {
  const expected = new Set(["a", "b"]);
  const omitted = {
    entry_ids: [["a"], ["b"]],
    bibliography_errors: [{ index: 0, itemID: "a", error_code: 1 }],
  };
  assert.deepEqual(
    verifyBibliographyResult([omitted, ["Second"]], expected, new Set(["a"])),
    ["Second"],
  );

  const refused: [unknown, ReadonlySet<string>][] = [
    // Not named as cited in the text only.
    [[omitted, ["Second"]], new Set()],
    // Named, but citeproc reports a different failure for it.
    [
      [
        {
          ...omitted,
          bibliography_errors: [{ index: 0, itemID: "a", error_code: 2 }],
        },
        ["Second"],
      ],
      new Set(["a"]),
    ],
    // Left out of the identity metadata as well as the list.
    [[{ ...omitted, entry_ids: [["b"]] }, ["Second"]], new Set(["a"])],
    // A work that should be listed is the one missing.
    [
      [
        {
          ...omitted,
          bibliography_errors: [{ index: 1, itemID: "b", error_code: 1 }],
        },
        ["First"],
      ],
      new Set(["a"]),
    ],
  ];
  for (const [result, inTextOnly] of refused) {
    assert.throws(
      () => verifyBibliographyResult(result, expected, inTextOnly),
      /citation processor/,
    );
  }
});

test("a bibliography entry outside citeproc's known html vocabulary is refused, not inserted", () => {
  const expected = new Set(["a"]);
  const metadata = { entry_ids: [["a"]], bibliography_errors: [] };

  // citeproc's html formatter only ever emits the tags CITEPROC_HTML_TAG enumerates and escapes everything else, so an unescaped tag outside that set means unvetted output.
  assert.throws(
    () =>
      verifyBibliographyResult(
        [metadata, ["A. Author, <i>Title</i>, <script>alert(1)</script>."]],
        expected,
      ),
    /known formatting vocabulary/,
  );

  // Every tag citeproc's html format can legitimately produce is accepted.
  assert.doesNotThrow(() =>
    verifyBibliographyResult(
      [
        metadata,
        [
          '<div class="csl-entry"><div class="csl-left-margin">[1]</div>' +
            '<div class="csl-right-inline">A, B. <i>Title</i> <b>bold</b> ' +
            '<sup>x</sup><sub>y</sub> <span style="font-style:normal;">n</span>' +
            "</div></div>",
        ],
      ],
      expected,
    ),
  );
});

test("APA cites a letter in the text only, as its rules say, and IEEE lists it", async () => {
  const source =
    "@letter{doe2024, author={Doe, Jane}, title={Letter to the board}, year={2024}}\n" +
    "@book{roe2020, author={Roe, Rick}, title={A Book}, year={2020}}\n";
  const apa = new CitationEngine("apa", loadAsset);
  apa.load(source);
  await apa.ensureReady();
  assert.equal(apa.item("doe2024")?.type, "personal_communication");

  const both = apa.formatDocument([
    { keys: ["doe2024"] },
    { keys: ["roe2020"] },
  ]);
  assert.equal(
    both.citations[0]?.text,
    "(J. Doe, personal communication, 2024)",
  );
  assert.match(both.bibliography.html, /Roe, R\. \(2020\)/);
  assert.doesNotMatch(both.bibliography.html, /Doe|Letter to the board/);

  // Cited on its own, it leaves APA nothing to list: not an error, and nothing for a caller to write.
  const alone = apa.formatDocument([{ keys: ["doe2024"] }]);
  assert.equal(
    alone.citations[0]?.text,
    "(J. Doe, personal communication, 2024)",
  );
  assert.equal(alone.bibliography.renderable, false);
  assert.equal(alone.bibliography.keysResolved, 1);

  const ieee = new CitationEngine("ieee", loadAsset);
  ieee.load(source);
  await ieee.ensureReady();
  assert.match(
    ieee.formatDocument([{ keys: ["doe2024"] }]).bibliography.html,
    /Letter to the board/,
  );
});

test("every style cites and lists an article, a book, a conference paper, a @misc, a letter and a web page in its own way", async () => {
  const source =
    (await readFile("sample/library.bib", "utf8")) +
    "@misc{moe2021, author={Moe, Max}, title={Some miscellaneous thing}, year={2021}}\n" +
    "@letter{doe2024, author={Doe, Jane}, title={Letter to the board}, year={2024}}\n" +
    "@online{wiki, author={Roe, Rick}, title={Team Notes}, year={2026}, url={https://example.org/notes}, urldate={2026-10-05}}\n";
  const cited = [
    "einstein1905",
    "knuth1984texbook",
    "lamport1994latex",
    "moe2021",
    "doe2024",
    "wiki",
  ].map((key) => ({ keys: [key] }));
  // Each work's citation in turn, then the reference list as Word shows it. APA and Chicago cite a letter in the text only, by their own rules, and every other style lists it; Cite Them Right and Nature are British.
  const expected: Record<StyleId, { citations: string[]; list: string[] }> = {
    "american-medical-association": {
      citations: ["1", "2", "3", "4", "5", "6"],
      list: [
        "1.\tEinstein A. Zur elektrodynamik bewegter körper. Annalen der Physik. 1905;322(10):891-921. doi:10.1002/andp.19053221004",
        "2.\tKnuth DE. The TeXBook. Addison-Wesley; 1984.",
        "3.\tLamport L. LaTeX: a document preparation system. In: Addison-Wesley Series. 1994.",
        "4.\tMoe M. Some miscellaneous thing. Published online 2021.",
        "5.\tDoe J. Letter to the board. Published online 2024.",
        "6.\tRoe R. Team notes. 2026. Accessed October 5, 2026. https://example.org/notes",
      ],
    },
    apa: {
      citations: [
        "(Einstein, 1905)",
        "(Knuth, 1984)",
        "(Lamport, 1994)",
        "(Moe, 2021)",
        "(J. Doe, personal communication, 2024)",
        "(Roe, 2026)",
      ],
      list: [
        "Einstein, A. (1905). Zur elektrodynamik bewegter körper. Annalen Der Physik, 322(10), 891–921. https://doi.org/10.1002/andp.19053221004",
        "Knuth, D. E. (1984). The TeXbook. Addison-Wesley.",
        "Lamport, L. (1994). LaTeX: a document preparation system. Addison-Wesley Series.",
        "Moe, M. (2021). Some miscellaneous thing.",
        "Roe, R. (2026). Team notes. https://example.org/notes",
      ],
    },
    "chicago-author-date": {
      citations: [
        "(Einstein 1905)",
        "(Knuth 1984)",
        "(Lamport 1994)",
        "(Moe 2021)",
        "(Jane Doe, “Letter to the Board”)",
        "(Roe 2026)",
      ],
      list: [
        "Einstein, Albert. 1905. “Zur Elektrodynamik Bewegter Körper.” Annalen Der Physik 322 (10): 891–921. https://doi.org/10.1002/andp.19053221004.",
        "Knuth, Donald E. 1984. The TeXBook. Addison-Wesley.",
        "Lamport, Leslie. 1994. “LaTeX: A Document Preparation System.” Addison-Wesley Series.",
        "Moe, Max. 2021. “Some Miscellaneous Thing.”",
        "Roe, Rick. 2026. “Team Notes.” https://example.org/notes.",
      ],
    },
    "harvard-cite-them-right": {
      citations: [
        "(Einstein, 1905)",
        "(Knuth, 1984)",
        "(Lamport, 1994)",
        "(Moe, 2021)",
        "(Doe, 2024)",
        "(Roe, 2026)",
      ],
      list: [
        "Doe, J. (2024) ‘Letter to the board’.",
        "Einstein, A. (1905) ‘Zur elektrodynamik bewegter körper’, Annalen der Physik, 322(10), pp. 891–921. Available at: https://doi.org/10.1002/andp.19053221004.",
        "Knuth, D.E. (1984) The TeXbook. Addison-Wesley.",
        "Lamport, L. (1994) ‘LaTeX: a document preparation system’, Addison-wesley series.",
        "Moe, M. (2021) ‘Some miscellaneous thing’.",
        "Roe, R. (2026) Team notes. Available at: https://example.org/notes (Accessed: 5 October 2026).",
      ],
    },
    ieee: {
      citations: ["[1]", "[2]", "[3]", "[4]", "[5]", "[6]"],
      list: [
        "[1]\tA. Einstein, “Zur elektrodynamik bewegter körper,” Annalen der Physik, vol. 322, no. 10, pp. 891–921, 1905, doi: 10.1002/andp.19053221004.",
        "[2]\tD. E. Knuth, The TeXbook. Addison-Wesley, 1984.",
        "[3]\tL. Lamport, “LaTeX: a document preparation system,” in Addison-wesley series, 1994.",
        "[4]\tM. Moe, “Some miscellaneous thing.” 2021.",
        "[5]\tJ. Doe, “Letter to the board,” 2024.",
        "[6]\tR. Roe, “Team notes.” Accessed: Oct. 05, 2026. [Online]. Available: https://example.org/notes",
      ],
    },
    "modern-language-association": {
      citations: [
        "(Einstein)",
        "(Knuth)",
        "(Lamport)",
        "(Moe)",
        "(Doe)",
        "(Roe)",
      ],
      list: [
        "Doe, Jane. “Letter to the Board.” 2024.",
        "Einstein, Albert. “Zur Elektrodynamik Bewegter Körper.” Annalen Der Physik, vol. 322, no. 10, 1905, pp. 891–921, https://doi.org/10.1002/andp.19053221004.",
        "Knuth, Donald E. The TeXBook. Addison-Wesley, 1984.",
        "Lamport, Leslie. “LaTeX: A Document Preparation System.” Addison-Wesley Series, 1994.",
        "Moe, Max. “Some Miscellaneous Thing.” 2021.",
        "Roe, Rick. Team Notes. 2026, https://example.org/notes.",
      ],
    },
    nature: {
      citations: ["1", "2", "3", "4", "5", "6"],
      list: [
        "1.\tEinstein, A. Zur elektrodynamik bewegter körper. Annalen der Physik 322, 891–921 (1905).",
        "2.\tKnuth, D. E. The TeXBook. (Addison-Wesley, 1984).",
        "3.\tLamport, L. LaTeX: a document preparation system. in Addison-wesley series (1994).",
        "4.\tMoe, M. Some miscellaneous thing. (2021).",
        "5.\tDoe, J. Letter to the board. (2024).",
        "6.\tRoe, R. Team notes. https://example.org/notes (2026).",
      ],
    },
    "nlm-citation-sequence": {
      citations: ["(1)", "(2)", "(3)", "(4)", "(5)", "(6)"],
      list: [
        "1.\tEinstein A. Zur elektrodynamik bewegter körper. Annalen der Physik. 1905;322(10):891–921. doi:10.1002/andp.19053221004",
        "2.\tKnuth DE. The TeXbook. Addison-Wesley; 1984.",
        "3.\tLamport L. LaTeX: a document preparation system. In: Addison-wesley series. 1994.",
        "4.\tMoe M. Some miscellaneous thing. 2021.",
        "5.\tDoe J. Letter to the board. 2024.",
        "6.\tRoe R. Team notes [Internet]. 2026 [cited 2026 Oct 5]. Available from: https://example.org/notes",
      ],
    },
  };
  assert.deepEqual(
    BUILTIN_STYLES.map((style) => style.id),
    Object.keys(expected),
    "every style the menu offers, and no other",
  );
  for (const style of BUILTIN_STYLES) {
    const engine = new CitationEngine(style.id, loadAsset);
    engine.load(source);
    await engine.ensureReady();
    const { citations, bibliography } = engine.formatDocument(cited);
    assert.deepEqual(
      citations.map((result) => result.text),
      expected[style.id].citations,
      style.id,
    );
    assert.deepEqual(
      ooxmlParagraphs(bibliography.ooxml).map((entry) => entry.text),
      expected[style.id].list,
      style.id,
    );
    // A @misc is no preprint, whatever CSL type the converter gave it.
    assert.doesNotMatch(bibliography.html, /preprint/i, style.id);
  }
});

test("every style lists each kind of work the converter reads, leaving out only what its own rules cite in the text only", async () => {
  // A signed and an unsigned work of every type the converter knows, each cited once: a style leaving out anything its registry entry does not name refuses the whole render.
  const source = Object.keys(BibTypes)
    .map(
      (type, index) =>
        `@${type}{s${index}, author={Author${index}, Ann}, title={Work s${index}x}, year={2020}}\n` +
        `@${type}{u${index}, title={Work u${index}x}, year={2020}}\n`,
    )
    .join("");
  const textOnly: Partial<
    Record<StyleId, { signed: string[]; unsigned: string[] }>
  > = {
    apa: {
      signed: ["interview", "personal_communication"],
      unsigned: ["interview", "personal_communication"],
    },
    // Chicago lists a signed dictionary or encyclopaedia entry, and an unsigned one only with a link.
    "chicago-author-date": {
      signed: ["classic", "interview", "personal_communication"],
      unsigned: [
        "classic",
        "entry-dictionary",
        "entry-encyclopedia",
        "interview",
        "personal_communication",
      ],
    },
  };
  for (const style of BUILTIN_STYLES) {
    const engine = new CitationEngine(style.id, loadAsset);
    engine.load(source);
    await engine.ensureReady();
    const keys = engine.keys();
    const listed = engine
      .formatDocument(keys.map((key) => ({ keys: [key] })))
      .bibliography.html.toLowerCase();
    const leftOut = (prefix: string): string[] =>
      Array.from(
        new Set(
          keys
            .filter(
              (key) =>
                key.startsWith(prefix) && !listed.includes(`work ${key}x`),
            )
            .map((key) => engine.item(key)?.type ?? ""),
        ),
      ).sort();
    const expected = textOnly[style.id] ?? { signed: [], unsigned: [] };
    assert.deepEqual(leftOut("s"), expected.signed, style.id);
    assert.deepEqual(leftOut("u"), expected.unsigned, style.id);
    // And the registry names those types and no others.
    assert.deepEqual(
      [...style.inTextOnly].sort(),
      Array.from(new Set([...expected.signed, ...expected.unsigned])).sort(),
      style.id,
    );
  }
});

test("each style's registry entry says what its stylesheet does", async () => {
  const bib = await readFile("sample/library.bib", "utf8");
  for (const style of BUILTIN_STYLES) {
    const xml = await readFile(`public/assets/csl/${style.id}.csl`, "utf8");
    assert.equal(
      /<category citation-format="([^"]+)"\/>/.exec(xml)?.[1],
      style.format,
      style.id,
    );
    // citeproc renders a style naming its own locale in that locale, whatever it is asked for.
    const declared = /<style\b[^>]*\bdefault-locale="([^"]+)"/.exec(xml)?.[1];
    assert.ok(
      declared === undefined || declared === style.locale,
      `${style.id} declares ${declared}`,
    );
    const citationLayout = /<citation\b[\s\S]*?<layout\b([^>]*)>/.exec(
      xml,
    )?.[1];
    assert.equal(
      citationLayout?.includes('vertical-align="sup"'),
      style.superscript,
      `${style.id} superscript`,
    );

    const engine = new CitationEngine(style.id, loadAsset);
    engine.load(bib);
    await engine.ensureReady();
    const [plain, paged] = engine.formatDocument([
      { keys: ["knuth1984texbook"] },
      {
        keys: ["knuth1984texbook"],
        pinpoint: { label: "page", locator: "97" },
      },
    ]).citations;
    assert.equal(plain?.text.includes("97"), false);
    assert.equal(paged?.text.includes("97"), style.pinpoints, style.id);
  }
});

test("a style and its locale are fetched the first time a document uses that style, and only then", async () => {
  const fetched: string[] = [];
  const engine = new CitationEngine("apa", async (path) => {
    fetched.push(path);
    return await loadAsset(path);
  });
  await engine.ensureReady();
  assert.deepEqual(fetched, [
    "assets/csl/locales-en-US.xml",
    "assets/csl/apa.csl",
  ]);
  for (const style of ["harvard-cite-them-right", "nature", "apa"]) {
    engine.setStyle(style);
    await engine.ensureReady();
  }
  // Cite Them Right and Nature are British, so that locale arrives with the first of them.
  assert.deepEqual(fetched.slice(2), [
    "assets/csl/locales-en-GB.xml",
    "assets/csl/harvard-cite-them-right.csl",
    "assets/csl/nature.csl",
  ]);
});

test("a processor kept between renders gives each document what a new one would, without reading the style again", async () => {
  const Engine = CSL.Engine;
  let constructed = 0;
  CSL.Engine = class extends Engine {
    constructor(...args: ConstructorParameters<typeof Engine>) {
      super(...args);
      constructed++;
    }
  };
  try {
    const engine = new CitationEngine("ieee", loadAsset);
    engine.load(
      "@book{jane2020, author={Smith, Jane}, title={Alpha}, publisher={P}, year={2020}}\n" +
        "@book{jane2020b, author={Smith, Jane}, title={Beta}, publisher={P}, year={2020}}\n" +
        "@book{adam2020, author={Smith, Adam}, title={Gamma}, publisher={P}, year={2020}}\n",
    );
    await engine.ensureReady();
    const cite = (...keys: string[]): string[] =>
      engine
        .formatDocument(keys.map((key) => ({ keys: [key] })))
        .citations.map((result) => result.text);

    // One processor serves every render in a style, and numbers each document from 1.
    assert.deepEqual(cite("adam2020", "jane2020"), ["[1]", "[2]"]);
    assert.deepEqual(cite("jane2020"), ["[1]"]);
    assert.equal(constructed, 1);

    // A new style gets its own. Jane Smith's two works of 2020 take a letter each and an initial tells her from Adam; without Adam she needs no initial, and one work alone no letter.
    engine.setStyle("harvard-cite-them-right");
    await engine.ensureReady();
    assert.deepEqual(cite("jane2020", "jane2020b", "adam2020"), [
      "(J. Smith, 2020a)",
      "(J. Smith, 2020b)",
      "(A. Smith, 2020)",
    ]);
    assert.deepEqual(cite("jane2020", "jane2020b"), [
      "(Smith, 2020a)",
      "(Smith, 2020b)",
    ]);
    assert.deepEqual(cite("jane2020b"), ["(Smith, 2020)"]);
    assert.equal(constructed, 2);

    // So does a new library, whose works it must read.
    engine.load(
      "@book{jane2020, author={Jones, Jane}, title={Alpha}, year={2020}}",
    );
    assert.match(
      engine.formatDocument([{ keys: ["jane2020"] }]).bibliography.html,
      /Jones, J\./,
    );
    assert.equal(constructed, 3);
  } finally {
    CSL.Engine = Engine;
  }
});

test("a form a style has no printed form for falls back to its standard form, never blocking a refresh", async () => {
  const source =
    (await readFile("sample/library.bib", "utf8")) +
    "@jurisdiction{brown1954, author={Warren, Earl}, title={Brown v. Board of {Education}}, year={1954}}\n" +
    "@letter{unsigned2024, title={Letter to the board}, year={2024}}\n";
  const cases: [StyleId, CitationInput, string][] = [
    // APA has no narrative form for a legal case, which it cites by name.
    [
      "apa",
      { keys: ["brown1954"], form: "narrative" },
      "(Brown v. Board of Education, 1954)",
    ],
    // Chicago has none for a letter no one signed, which it cites by title.
    [
      "chicago-author-date",
      { keys: ["unsigned2024"], form: "narrative" },
      "(“Letter to the Board”)",
    ],
  ];
  for (const [style, citation, text] of cases) {
    const engine = new CitationEngine(style, loadAsset);
    engine.load(source);
    await engine.ensureReady();
    assert.equal(engine.formatCitation(citation).text, text, style);
  }

  // In a document, that citation takes its standard form while every other keeps the form its tag asks for.
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(source);
  await engine.ensureReady();
  try {
    const legal = makeWordControl("bibcite:brown1954 narrative", "(old)");
    const book = makeWordControl("bibcite:knuth1984texbook narrative", "(old)");
    installWordItems([legal, book]);
    await refreshDocument(engine);
    assert.deepEqual(
      [legal.text, book.text],
      ["(Brown v. Board of Education, 1954)", "Knuth (1984)"],
    );
  } finally {
    clearHostMocks();
  }
});

test("APA's narrative citations join two names with “and”, where its bracketed ones and its reference list keep “&”", async () => {
  const Engine = CSL.Engine;
  let constructed = 0;
  CSL.Engine = class extends Engine {
    constructor(...args: ConstructorParameters<typeof Engine>) {
      super(...args);
      constructed++;
    }
  };
  try {
    const engine = new CitationEngine("apa", loadAsset);
    engine.load(
      "@article{pair2020a, author={Smith, Jane and Jones, Bob}, title={One}, journal={J}, year={2020}}\n" +
        "@article{pair2020b, author={Smith, Jane and Jones, Bob}, title={Two}, journal={J}, year={2020}}\n" +
        "@article{lee2021, author={Lee, Ann and Kim, Bo and Park, Cy}, title={Three}, journal={J}, year={2021}}\n" +
        "@article{lee2021b, author={Lee, Ann and Cho, Di and Park, Cy}, title={Four}, journal={J}, year={2021}}\n" +
        "@report{firm2019, author={{Johnson \\& Johnson}}, title={Five}, institution={I}, year={2019}}\n" +
        "@report{pair2019, author={Johnson, Al and Johnson, Bea}, title={Six}, institution={I}, year={2019}}\n" +
        "@book{brown2017a, author={Brown, Alice}, title={Seven}, publisher={P}, year={2017}}\n" +
        "@book{brown2017c, author={Brown, Carl}, title={Eight}, publisher={P}, year={2017}}\n" +
        "@book{edited2016, editor={Green, Ed and White, Fay}, title={Nine}, publisher={P}, year={2016}}\n" +
        "@letter{letter2020, author={Doe, Jane and Roe, Rick}, title={Ten}, date={2020-05-01}}\n",
    );
    await engine.ensureReady();
    const cite = (citations: CitationInput[]) => {
      const { citations: results, bibliography } =
        engine.formatDocument(citations);
      return { texts: results.map((result) => result.text), bibliography };
    };

    // A citation of one author has no "&" to change, so it needs no second processor.
    assert.deepEqual(
      cite([{ keys: ["brown2017a"], form: "narrative" }]).texts,
      ["Brown (2017)"],
    );
    assert.equal(constructed, 1);

    // Each work cited both ways in one document: initials and letters tell the works apart in both forms alike, and only a person's name changes its "&".
    const keys = [
      "pair2020a",
      "pair2020b",
      "lee2021",
      "lee2021b",
      "firm2019",
      "pair2019",
      "brown2017a",
      "brown2017c",
      "edited2016",
      "letter2020",
    ];
    const bracketed: CitationInput[] = keys.map((key) => ({ keys: [key] }));
    const both = cite(
      keys.flatMap((key): CitationInput[] => [
        { keys: [key] },
        { keys: [key], form: "narrative" },
      ]),
    );
    assert.deepEqual(both.texts, [
      "(Smith & Jones, 2020a)",
      "Smith and Jones (2020a)",
      "(Smith & Jones, 2020b)",
      "Smith and Jones (2020b)",
      "(Lee, Kim, et al., 2021)",
      "Lee, Kim, et al. (2021)",
      "(Lee, Cho, et al., 2021)",
      "Lee, Cho, et al. (2021)",
      "(Johnson & Johnson, 2019)",
      "Johnson & Johnson (2019)",
      "(Johnson & Johnson, 2019)",
      "Johnson and Johnson (2019)",
      "(A. Brown, 2017)",
      "A. Brown (2017)",
      "(C. Brown, 2017)",
      "C. Brown (2017)",
      "(Green & White, 2016)",
      "Green and White (2016)",
      "(J. Doe & R. Roe, personal communication, May 1, 2020)",
      "J. Doe and R. Roe (personal communication, May 1, 2020)",
    ]);
    assert.equal(constructed, 2);
    // The reference list is the one a document of bracketed citations alone gets, with its "&".
    const list = cite(bracketed).bibliography.html;
    assert.equal(both.bibliography.html, list);
    assert.match(list, /Smith, J\., &#38; Jones, B\. \(2020a\)/);

    // A page and the year-only form keep APA's own rendering, and the processor for narrative citations is kept for the next render.
    assert.deepEqual(
      cite([
        {
          keys: ["pair2020a"],
          form: "narrative",
          pinpoint: { label: "page", locator: "12" },
        },
        { keys: ["pair2020a"], form: "suppress-author" },
      ]).texts,
      ["Smith and Jones (2020, p. 12)", "(2020)"],
    );
    assert.equal(constructed, 2);
    assert.equal(
      engine.formatCitation({ keys: ["edited2016"], form: "narrative" }).text,
      "Green and White (2016)",
    );
    assert.equal(constructed, 2);
  } finally {
    CSL.Engine = Engine;
  }
});

test("Word sets the citations of a superscript style in superscript, and any other style's back", async () => {
  const engine = new CitationEngine("american-medical-association", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  try {
    const first = makeWordControl("bibcite:einstein1905", "(old)");
    const paged = makeWordControl("bibcite:knuth1984texbook page 12", "(old)");
    const list = makeWordControl("bibcite-bibliography", "Existing");
    const host = installWordItems([first, paged, list]);
    await refreshDocument(engine);
    // The page is part of AMA's superscript; the reference list is not superscript.
    assert.deepEqual(
      [first, paged].map((control) => [control.text, control.font.superscript]),
      [
        ["1", true],
        ["2(p12)", true],
      ],
    );
    assert.equal(list.font.superscript, false);

    // So is a citation Nature inserts, from the moment it is inserted.
    engine.setStyle("nature");
    await engine.ensureReady();
    await insertCitation(engine, ["shannon1948"]);
    const inserted = host.controls.items.at(-1);
    assert.deepEqual([inserted?.text, inserted?.font.superscript], ["3", true]);

    engine.setStyle("apa");
    await engine.ensureReady();
    await refreshDocument(engine);
    assert.deepEqual(
      host.controls.items.map((control) => control.font.superscript),
      [false, false, false, false],
    );

    // A batch Word rejects after applying it is undone, superscript and all.
    const kept = makeWordControl("bibcite:einstein1905", "(Einstein, 1905)");
    installWordItems([kept], {
      onSync(call) {
        if (call === 2) throw new Error("simulated mid-batch host failure");
      },
    });
    engine.setStyle("american-medical-association");
    await engine.ensureReady();
    await assert.rejects(
      () => refreshDocument(engine),
      new RegExp(`previous ${NAME} text was restored`),
    );
    assert.deepEqual(
      [kept.text, kept.font.superscript],
      ["(Einstein, 1905)", false],
    );
  } finally {
    clearHostMocks();
  }
});

test("missing keys remain visible without changing resolved numbering", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  const result = engine.formatDocument([
    { keys: ["einstein1905", "missing-key"] },
    { keys: ["knuth1984texbook"] },
  ]);
  assert.equal(result.citations[0]?.text, "[1] [?missing-key]");
  assert.equal(result.citations[1]?.text, "[2]");
  assert.deepEqual(result.bibliography.missing, ["missing-key"]);

  const ambiguous = engine.formatDocument([{ keys: ["a][?forged"] }]);
  assert.equal(ambiguous.citations[0]?.text, "[?a%5D%5B%3Fforged]");
  assert.doesNotMatch(ambiguous.citations[0]?.text ?? "", /\]\[\?/);
});

test("document-wide APA formatting applies year-suffix disambiguation", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(
    "@article{first, author={Smith, Jane}, title={Alpha}, year={2020}}\n" +
      "@article{second, author={Smith, Jane}, title={Beta}, year={2020}}",
  );
  await engine.ensureReady();

  const result = engine.formatDocument([
    { keys: ["first"] },
    { keys: ["second"] },
  ]);
  assert.match(result.citations[0]?.text ?? "", /2020a/);
  assert.match(result.citations[1]?.text ?? "", /2020b/);
});

test("APA's reference list italicises the container title; citations stay plain in every supported style", async () => {
  const bib = await readFile("sample/library.bib", "utf8");
  for (const styleId of ["apa", "ieee"] as const) {
    const engine = new CitationEngine(styleId, loadAsset);
    engine.load(bib);
    await engine.ensureReady();
    const result = engine.formatDocument([{ keys: ["einstein1905"] }]);
    assert.doesNotMatch(
      result.citations[0]?.text ?? "",
      /[<>]/,
      `${styleId} in-text citation must stay plain text`,
    );
  }

  const apa = new CitationEngine("apa", loadAsset);
  apa.load(bib);
  await apa.ensureReady();
  assert.match(
    apa.formatDocument([{ keys: ["einstein1905"] }]).bibliography.html,
    /<i>Annalen Der Physik<\/i>/,
    "APA must italicise the journal (container title) in the reference list",
  );
});

test("Word refresh updates citations and bibliography only after all tags validate", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  const first = makeWordControl("bibcite:einstein1905");
  const second = makeWordControl("bibcite:knuth1984texbook");
  const bibliography = makeWordControl("bibcite-bibliography");
  const host = installWordItems([first, second, bibliography]);

  const report = await refreshDocument(engine);
  assert.equal(first.text, "[1]");
  assert.equal(second.text, "[2]");
  assert.deepEqual(
    ooxmlParagraphs(bibliography.text).map((paragraph) =>
      paragraph.text.slice(0, 4),
    ),
    ["[1]\t", "[2]\t"],
  );
  assert.match(
    bibliography.text,
    /<w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve">Annalen der Physik<\/w:t>/,
  );
  assert.equal(report.citationsUpdated, 2);

  first.text = "unchanged";
  host.controls.items = [
    first,
    makeWordControl("bibcite:good,,corrupt"),
    bibliography,
  ];
  // The failure names the control as well as the key, because the user has to find it.
  await assert.rejects(
    () => refreshDocument(engine),
    new RegExp(`Unsupported.*Found in ${NAME} control #2`, "s"),
  );
  assert.equal(first.text, "unchanged");
  clearHostMocks();
});

test("a document mixing plain citations and ones with a page or form refreshes each as its tag says", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  try {
    const plain = makeWordControl("bibcite:einstein1905", "(old)");
    const paged = makeWordControl("bibcite:knuth1984texbook page 12", "(old)");
    const narrative = makeWordControl(
      "bibcite:einstein1905 narrative",
      "(old)",
    );
    const host = installWordItems([plain, paged, narrative]);
    await refreshDocument(engine);
    assert.deepEqual(
      [plain, paged, narrative].map((control) => control.text),
      ["(Einstein, 1905)", "(Knuth, 1984, p. 12)", "Einstein (1905)"],
    );
    // The tag is the data, so a refresh reads it and never rewrites it.
    assert.deepEqual(
      [plain, paged, narrative].map((control) => control.tag),
      [
        "bibcite:einstein1905",
        "bibcite:knuth1984texbook page 12",
        "bibcite:einstein1905 narrative",
      ],
    );

    await insertCitation(engine, ["shannon1948"], {
      form: "narrative",
      pinpoint: { label: "page", locator: "380" },
    });
    const inserted = host.controls.items.at(-1);
    assert.equal(inserted?.tag, "bibcite:shannon1948 narrative page 380");
    assert.equal(inserted?.text, "Shannon (1948, p. 380)");
  } finally {
    clearHostMocks();
  }
});

test("a bibliography is inserted and refreshed as Office Open XML, never through Word's html import", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  try {
    const host = installWordItems([makeWordControl("bibcite:einstein1905")]);
    await insertBibliography(engine);
    const inserted = host.controls.items.find(
      (control) => control.tag === "bibcite-bibliography",
    );
    assert.match(inserted?.text ?? "", /^<pkg:package /);
    assert.equal(ooxmlParagraphs(inserted?.text ?? "").length, 1);
    // Word keeps the paragraph mark the package's last paragraph lands on, so that entry's hanging indent is set directly as well.
    assert.deepEqual(inserted?.lastParagraph, {
      leftIndent: 36,
      firstLineIndent: -36,
    });

    const existing = makeWordControl("bibcite-bibliography", "Old list");
    installWordItems([
      makeWordControl("bibcite:einstein1905"),
      makeWordControl("bibcite:shannon1948"),
      existing,
    ]);
    await refreshDocument(engine);
    assert.deepEqual(
      ooxmlParagraphs(existing.text).map((entry) => entry.text.slice(0, 9)),
      ["Einstein,", "Shannon, "],
    );
    assert.deepEqual(existing.lastParagraph, {
      leftIndent: 36,
      firstLineIndent: -36,
    });

    // IEEE's column: "[1]" is 3 characters, so (3 + 1) × 6 pt.
    engine.setStyle("ieee");
    await engine.ensureReady();
    await refreshDocument(engine);
    assert.deepEqual(existing.lastParagraph, {
      leftIndent: 24,
      firstLineIndent: -24,
    });
  } finally {
    clearHostMocks();
  }
});

test("Word writes preflight locks and restore prior text after a failed batch", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  const locked = makeWordControl("bibcite:einstein1905", "locked original");
  locked.cannotEdit = true;
  installWordItems([locked]);
  await assert.rejects(
    () => refreshDocument(engine),
    /locked against editing.*nothing was changed/s,
  );
  assert.equal(locked.text, "locked original");

  const first = makeWordControl("bibcite:einstein1905", "first original");
  const second = makeWordControl("bibcite:knuth1984texbook", "second original");
  installWordItems([first, second], {
    // Sync 2 is the write batch; the mock mutates eagerly, modelling a host that changed the document before reporting the failure.
    onSync(call) {
      if (call === 2) throw new Error("simulated mid-batch host failure");
    },
  });
  await assert.rejects(
    () => refreshDocument(engine),
    new RegExp(`previous ${NAME} text was restored`),
  );
  assert.equal(first.text, "first original");
  assert.equal(second.text, "second original");
  clearHostMocks();
});

test("Word insertion is confined to the main document body", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  const host = installWordItems([], { selectionBodyType: "Header" });
  await assert.rejects(
    () => insertCitation(engine, ["einstein1905"]),
    /main document body/,
  );
  assert.equal(
    host.controls.items.length,
    0,
    "no header content control was inserted",
  );

  const failing = installWordItems([], {
    onSync(call) {
      if (call === 3) throw new Error("simulated insertion commit failure");
    },
  });
  await assert.rejects(
    () => insertCitation(engine, ["einstein1905"]),
    /pending content control was removed/,
  );
  assert.equal(
    failing.controls.items.length,
    0,
    "a partially committed insertion is cleaned up",
  );
  clearHostMocks();
});

test("a % comment cannot reject a library, and cannot truncate a field", () => {
  const engine = new CitationEngine("apa", loadAsset);
  const entry =
    "@article{smith2020, author={Smith, J}, title={A Title}, journal={J}, year={2020}}";

  // The converter has no comment syntax, so an @ in a header line reads as the start of an entry.
  assert.equal(engine.load("% maintained by a@b.com\n" + entry).entries, 1);
  assert.equal(engine.load("% Encoding: UTF-8\n" + entry).entries, 1);
  assert.equal(engine.load("   % indented, see a@b.com\n" + entry).entries, 1);

  // Inside an entry a % is ordinary text, and brace depth alone cannot tell: BibTeX also allows @article(…), where a quoted value sits at depth 0.
  for (const source of [
    "@article{k, title={A 50% improvement}, year={2020}}",
    '@article{k, title="A 50% improvement", year="2020"}',
    "@article(k, title={A 50% improvement}, year={2020})",
    '@article(k, title="A 50% improvement", year="2020")',
    '@article{k, note="unmatched } brace", title={A 50% improvement}, year={2020}}',
  ]) {
    engine.load(source);
    assert.equal(engine.item("k")?.title, "A 50% improvement", source);
  }

  // Escaped closers inside a multiline quoted value must not make a later line-leading percent sign look like an inter-entry comment.
  for (const source of [
    '@article{k,\n title="A \\} brace\n% percent title text\ncontinued",\n year={2020}\n}',
    '@article(k,\n title="A \\) paren\n% percent title text\ncontinued",\n year={2020}\n)',
  ]) {
    for (const eol of ["\n", "\r\n", "\r"]) {
      engine.load(source.replace(/\n/g, eol));
      assert.match(String(engine.item("k")?.title), /% percent title text/);
    }
  }

  // The converter can hang on CR-only comment input, so preprocessing normalises line endings before parsing.
  assert.equal(
    engine.load(("% old\n" + entry).replace(/\n/g, "\r")).entries,
    1,
  );

  assert.deepEqual(
    engine.load(
      "% @article{old2019, author={A, B}, title={Superseded}, year={2019}}\n" +
        "@article{new2020, author={A, B}, title={Current}, year={2020}}",
    ).entries,
    1,
  );
  assert.equal(
    engine.item("old2019"),
    undefined,
    "a commented-out entry stays out",
  );
  assert.equal(
    engine.load(
      "% @article{smith2020, title={Old draft}, year={2019}}\n" +
        "@article{smith2020, title={Corrected}, year={2020}}",
    ).entries,
    1,
    "a commented-out entry cannot collide with its replacement",
  );
  assert.equal(engine.item("smith2020")?.title, "Corrected");

  // Newlines survive, so diagnostics still name the right row.
  assert.throws(
    () =>
      engine.load(
        "% one a@b.com\n% two\n@article{good, title={T}}\n@article{bad, title={T}, urldate={not a date}}",
      ),
    /line 4/,
  );
});

test("a % comment after an entry's closing brace cannot take the next entry with it", () => {
  const engine = new CitationEngine("apa", loadAsset);
  // Left for the parser, the comment lost its line break to TeX substitution (\ss, \LaTeX) and ran on through entry b.
  for (const comment of ["% checked by Gau\\ss", "% typeset in \\LaTeX"]) {
    const report = engine.load(
      `@article{a, title={A}, year={2019}} ${comment}\n` +
        "@article{b,\n  title = {B},\n  year  = {2020},\n}\n" +
        "@book{c, title={C}, year={2021}}\n",
    );
    assert.equal(report.entries, 3, comment);
    assert.deepEqual(engine.keys(), ["a", "b", "c"], comment);
    assert.equal(engine.item("b")?.title, "B", comment);
  }
});

test("a % comment between an entry's fields cannot take the next field with it", () => {
  const engine = new CitationEngine("apa", loadAsset);
  // The same lost line break: left for the parser, the comment ran on through the field after it.
  for (const comment of ["% checked by Gau\\ss", "% pages to add \\ldots"]) {
    for (const [open, close] of [
      ["{", "}"],
      ["(", ")"],
    ]) {
      engine.load(
        `@article${open}a,\n  title = {A},\n  ${comment}\n  journal = {J},\n  year = {2019}\n${close}\n`,
      );
      assert.equal(engine.item("a")?.["container-title"], "J", comment);
    }
    engine.load(`@article{a,\n  title = {A}, ${comment}\n  year = {2019}\n}\n`);
    assert.deepEqual(
      engine.item("a")?.issued,
      { "date-parts": [[2019]] },
      `${comment}, after a field on its line`,
    );
  }

  // Inside a value a `%` after a space is text, straight after text it can belong to a key, and an @comment block is text throughout.
  engine.load(
    '@article{50%off, title={A 50 % gain}, note="B 50 % gain", year={2020}}\n' +
      "@comment{jabref-meta: fileDirectory:/data/100 % done;}\n" +
      "@book{b, title={B}, year={2021}}\n",
  );
  assert.deepEqual(engine.keys(), ["50%off", "b"]);
  assert.equal(engine.item("50%off")?.title, "A 50 % gain");
  assert.equal(engine.item("50%off")?.note, "B 50 % gain");

  assert.throws(
    () =>
      engine.load(
        "@article{good,\n  title={T},\n  % note\n  year={2020}}\n@article{bad, title={T}, urldate={not a date}}",
      ),
    /line 5/,
    "the comment's line break survives, so diagnostics still name the right row",
  );
});

test("a backslash run in a .bib cannot stall the comment scanner", () => {
  // 250,000 backslashes take milliseconds under a linear scan and about 90 seconds under the quadratic one this replaced.
  const run = "\\".repeat(250_000);
  const source = `% header\n@article{a, note={${run}}}\n`;

  const started = Date.now();
  const stripped = stripPercentComments(source);
  const elapsed = Date.now() - started;
  assert.ok(
    elapsed < 5_000,
    `comment scanning took ${elapsed}ms; it must stay linear`,
  );

  assert.equal(stripped, `\n@article{a, note={${run}}}\n`);
});

test("a % comment the parser reaches can neither hang nor overflow it", () => {
  const engine = new CitationEngine("apa", loadAsset);
  const entry = "@article{a, title={T}, year={2020}}";

  // The parser looked for a comment's line break past the end of the input, forever.
  assert.equal(loadWithin(engine, `${entry} % note`).entries, 1);
  // TeX substitution eats the line break after \ldots, so a final newline is no guarantee of one.
  assert.equal(
    loadWithin(engine, `${entry} % pages still to add \\ldots\n`).entries,
    1,
  );
  // Straight after a comma a comment is left for the parser, so an unclosed last entry ending in one reaches the same scan; it is refused, not looped on.
  assert.throws(
    () => loadWithin(engine, "@article{a, title={T},% year={2020}"),
    /BibTeX parsing failed/,
  );

  // It also recursed once per comment line, which overflowed the stack well before 100,000. The stripper takes such lines out first, but not these: it counts the `{` in the first comment, which the parser skips, as opening a value until the `}` in the last.
  const commented = `@article{a,\n title={T},%{\n${"% c\n".repeat(100_000)}%}\n year={2020}}\n`;
  assert.equal(loadWithin(engine, commented).entries, 1);
  assert.deepEqual(engine.item("a")?.issued, { "date-parts": [[2020]] });
});

test("fields other tools write do not bury the warnings that matter", () => {
  const engine = new CitationEngine("apa", loadAsset);
  const many = Array.from(
    { length: 25 },
    (_, i) =>
      `@article{s${i}, author={S, J}, title={P${i}}, journal={J}, year={2020}, file={s${i}.pdf}, groups={inbox}}`,
  ).join("\n");
  assert.deepEqual(
    engine.load(many).warnings,
    [],
    "a library written by the capture tool is quiet",
  );
  assert.equal(
    engine.load("@article{x, title={T}, year={2020}, notarealfield={y}}")
      .warnings.length,
    1,
  );

  const noisyFields = Array.from(
    { length: 75 },
    (_, index) => `unknown${index}={x}`,
  ).join(",");
  const warnings = engine.load(
    `@article{noisy, title={T}, year={2020}, ${noisyFields}}`,
  ).warnings;
  assert.equal(warnings.length, 51, "50 details plus one bounded summary");
  assert.match(warnings.at(-1) ?? "", /plus 25 more parser warnings/);
});

test("unsafe metadata is rejected before it can reach citeproc or Word", () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load("@article{stable, title={Stable}, year={2020}}");
  assert.throws(
    () =>
      engine.load(
        "@article{spoof, title={Quarterly report\u202Efdp.exe}, year={2020}}",
      ),
    /unsafe invisible or directional characters/,
  );
  assert.equal(
    engine.item("stable")?.title,
    "Stable",
    "the prior library remains active",
  );

  const embedded = embeddedSnapshot({
    spoof: { id: "spoof", title: "Quarterly report\u202Efdp.exe" },
  });
  assert.equal(parseSnapshot(embedded).status, "invalid");
});

test("a library with genuine Persian/Urdu/Kurdish ZWNJ or an emoji's variation selector loads, not just the rest of the library around it", () => {
  const engine = new CitationEngine("apa", loadAsset);
  // "می‌خواهم" contains a real ZWNJ (joining behaviour, not hidden metadata); the emoji title carries VS16.
  // Titles are already sentence case, which bibliojson's import-time sentence-casing leaves as written, so the assertions test only the character check.
  engine.load(
    "@article{persian, title={می‌خواهم}, author={Ahmadi, Sara}, year={2020}}\n" +
      "@article{emoji, title={Cardiac monitoring with a ❤️}, year={2021}}",
  );
  assert.equal(engine.item("persian")?.title, "می‌خواهم");
  assert.equal(engine.item("emoji")?.title, "Cardiac monitoring with a ❤️");

  const embedded = embeddedSnapshot({
    persian: { id: "persian", title: "می‌خواهم" },
    emoji: { id: "emoji", title: "Cardiac monitoring with a ❤️" },
  });
  assert.equal(parseSnapshot(embedded).status, "valid");
});

test("pathological rich-text nesting is rejected before it can reach citeproc", () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load("@article{stable, title={Stable}, year={2020}}");

  // 101 raw "<" characters, one past the 100 the fix allows; citeproc's flip-flopper recurses per nested tag and overflows the stack (RangeError at ~5,000 in review).
  const nested = "<i>".repeat(101);
  assert.throws(
    () => engine.load(`@article{nestedspoof, title={${nested}}, year={2020}}`),
    /with more than 100 "<" characters/,
  );
  assert.equal(
    engine.item("stable")?.title,
    "Stable",
    "the prior library remains active",
  );
  assert.equal(
    parseSnapshot(
      embeddedSnapshot({ nestedspoof: { id: "nestedspoof", title: nested } }),
    ).status,
    "invalid",
  );
});

/** Unlike the checks above, an over-length field is a UX problem, not an attack signal, so it is shortened and flagged rather than failing the whole import. */
test("an over-length field is shortened and flagged, not rejected", () => {
  const engine = new CitationEngine("apa", loadAsset);
  const overlong = "x".repeat(MAX_FIELD_LENGTH + 1);

  const report = engine.load(
    `@article{overlong, title={${overlong}}, journal={J}, year={2020}}`,
  );
  assert.equal(report.entries, 1, "the entry is kept, not dropped");
  assert.equal(engine.item("overlong")?.title?.length, MAX_FIELD_LENGTH);
  assert.deepEqual(engine.truncatedFields("overlong"), ["title"]);

  // A field nowhere near the cap is not flagged.
  engine.load("@article{fine, title={Fine}, journal={J}, year={2020}}");
  assert.deepEqual(engine.truncatedFields("fine"), []);

  // The embedded-snapshot path (CitationEngine.loadSnapshot) tracks it independently of the .bib load path exercised above.
  const snapshotEngine = new CitationEngine("apa", loadAsset);
  snapshotEngine.loadSnapshot({
    overlong: { id: "overlong", title: overlong },
  });
  assert.equal(
    snapshotEngine.item("overlong")?.title?.length,
    MAX_FIELD_LENGTH,
  );
  assert.deepEqual(snapshotEngine.truncatedFields("overlong"), ["title"]);

  // Reading the raw embedded snapshot (e.g. written by an older build, before this cap was lowered from 32 KB) shortens rather than refusing the whole document's citation data.
  const read = parseSnapshot(
    embeddedSnapshot({ overlong: { id: "overlong", title: overlong } }),
  );
  assert.equal(read.status, "valid");
  if (read.status === "valid") {
    assert.equal(
      read.snapshot.items.overlong?.title,
      overlong.slice(0, MAX_FIELD_LENGTH),
    );
  }
});

/** Import/edit only ever produces plain CSL-JSON, never a Word rich-text run, so a hostile title is neither executed nor rewritten at load time — citeproc's own html-escaping at render time (verified by assertSafeBibliographyHtml) is what keeps it inert. */
test("a hostile title is stored as inert data and renders escaped, never as a live tag", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(
    "@article{hostile, author={Doe, Jane}, " +
      "title={A <script>alert(1)</script> Study}, journal={J}, year={2020}}",
  );
  // Loaded as data: nothing at import strips or executes it (bibliojson's own sentence-case conversion of the title is unrelated and harmless).
  assert.equal(
    engine.item("hostile")?.title,
    "A <script>alert(1)</script> study",
  );
  await engine.ensureReady();

  const result = engine.formatDocument([{ keys: ["hostile"] }]);
  assert.doesNotMatch(result.bibliography.html, /<script/);
  assert.match(result.bibliography.html, /&#60;script&#62;/);
});

/** The nesting-limit defence (assertSafeString's MAX_FIELD_ANGLE_BRACKETS) protects the html render too: both it and the plain-text citation render read the one Library through the one import-time check. */
test("the reference-list html render is bound by the same nesting limit import already enforces", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  const atLimit = "<i>".repeat(100); // exactly the cap import allows through
  engine.load(
    `@article{atlimit, title={${atLimit}}, journal={J}, year={2020}}`,
  );
  await engine.ensureReady();

  const result = engine.formatDocument([{ keys: ["atlimit"] }]);
  assert.ok(result.bibliography.html.length > 0);

  // Nothing above the cap ever reaches this render: import refuses it first.
  assert.throws(
    () =>
      engine.load(
        `@article{overlimit, title={${atLimit}<i>}, journal={J}, year={2020}}`,
      ),
    /more than 100 "<" characters/,
  );
});

test("the embedded snapshot excludes abstract and note", () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(
    "@article{k, title={T}, year={2020}, abstract={Secret abstract.}, note={Private note.}}",
  );
  assert.equal(
    engine.item("k")?.abstract,
    "Secret abstract.",
    "the live library, backed by the .bib, still holds it",
  );
  const snap = engine.snapshot(["k"]);
  assert.equal(snap.k?.abstract, undefined);
  assert.equal(snap.k?.note, undefined);
  assert.equal(snap.k?.title, "T", "unrelated fields still embed");
});

/** A File-shaped stand-in, so the size gate can be exercised without the bytes. */
function fileOf(
  name: string,
  bytes: Uint8Array,
  size = bytes.byteLength,
): File {
  return {
    name,
    size,
    arrayBuffer: () =>
      Promise.resolve(
        bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ),
      ),
  } as unknown as File;
}

test("the file source enforces the gates the security model advertises", async () => {
  const utf8 = new TextEncoder().encode(
    "@article{a, title={Caf\u00e9}, year={2020}}",
  );
  assert.match(
    await new FilePickerBibSource(fileOf("ok.bib", utf8)).read(),
    /Caf\u00e9/,
  );

  // File.text() would replace invalid UTF-8 with U+FFFD and silently mutate a citation key, so the decode is fatal.
  const invalid = new Uint8Array([0x40, 0x61, 0xff, 0xfe, 0x62]);
  await assert.rejects(
    () => new FilePickerBibSource(fileOf("bad.bib", invalid)).read(),
    /not valid UTF-8/,
  );

  const withNul = new TextEncoder().encode("@article{a, title={T}}\u0000");
  await assert.rejects(
    () => new FilePickerBibSource(fileOf("nul.bib", withNul)).read(),
    /NUL bytes/,
  );

  // The size gate reads File.size, so a huge selection fails before any read.
  const oversize = fileOf("big.bib", utf8, 21 * 1024 * 1024);
  await assert.rejects(
    () => new FilePickerBibSource(oversize).read(),
    /over the 20 MB safety limit/,
  );
  assert.equal(
    new FilePickerBibSource(fileOf("edge.bib", utf8, 20 * 1024 * 1024)).label,
    "edge.bib",
  );
});

test("a hostile filename cannot misrepresent itself in the status line", () => {
  const bytes = new TextEncoder().encode("@article{a, title={T}}");
  const spoofed = new FilePickerBibSource(fileOf("safe\u202Egpj.exe", bytes))
    .label;
  assert.doesNotMatch(
    spoofed,
    /\u202E/,
    "the bidi override is replaced, not passed through",
  );
  assert.match(spoofed, /\uFFFD/);

  const long = new FilePickerBibSource(fileOf(`${"n".repeat(400)}.bib`, bytes))
    .label;
  assert.ok(long.length <= 160, `label clipped to ${long.length}`);
  assert.match(long, /\u2026$/);
});

test("every offered style has a stylesheet, and only offered styles are reachable", async () => {
  // isStyleId is what stops an embedded style id selecting an arbitrary same-origin file.
  for (const style of BUILTIN_STYLES) {
    assert.ok(isStyleId(style.id));
    const xml = await readFile(`public/assets/csl/${style.id}.csl`, "utf8");
    assert.match(
      xml,
      /<style[\s>]/,
      `public/assets/csl/${style.id}.csl must be a CSL stylesheet`,
    );
  }
  for (const value of [
    "../../etc/passwd",
    "apa.csl",
    "APA",
    "",
    "locales-en-US",
    null,
    1,
    {},
  ]) {
    assert.equal(
      isStyleId(value),
      false,
      `must reject ${JSON.stringify(value)}`,
    );
  }
});

test("citation assets are streamed within a byte limit and checked for integrity", async () => {
  assert.equal(
    await readBoundedText(new Response("valid UTF-8"), "test.xml"),
    "valid UTF-8",
  );
  await assert.rejects(
    () =>
      readBoundedText(new Response(new Uint8Array([0xff, 0xfe])), "bad.xml"),
    /not valid UTF-8/,
  );

  let pulls = 0;
  let cancelled = false;
  const oversized = new Response(
    new ReadableStream<Uint8Array>({
      pull(controller) {
        pulls++;
        controller.enqueue(new Uint8Array(1024 * 1024));
      },
      cancel() {
        cancelled = true;
      },
    }),
  );
  await assert.rejects(
    () => readBoundedText(oversized, "large.xml"),
    /unexpectedly large/,
  );
  assert.equal(
    cancelled,
    true,
    "the oversized stream is cancelled immediately",
  );
  assert.ok(pulls <= 7, `read ${pulls} MiB before enforcing the 5 MiB limit`);

  const tampered = new CitationEngine(
    "apa",
    async (path) => `${await loadAsset(path)} `,
  );
  tampered.load("@article{a, title={T}, year={2020}}");
  await assert.rejects(() => tampered.ensureReady(), /SHA-256 integrity check/);
});

test("versioned snapshots preserve style and reject tampering", async () => {
  const items = {
    stable: {
      id: "stable",
      type: "article-journal",
      title: "Stable",
    },
  };
  const encoded = serializeSnapshot("ieee", items);
  const result = parseSnapshot(encoded);
  assert.equal(result.status, "valid");
  if (result.status === "valid") {
    assert.equal(result.snapshot.styleId, "ieee");
    assert.equal(result.snapshot.items.stable?.title, "Stable");
    assert.match(
      result.snapshot.rendererRevision,
      /^[0-9a-f]{64}:[0-9a-f]{64}$/,
    );
  }

  const stale = JSON.parse(encoded) as { rendererRevision: string };
  stale.rendererRevision = `${"0".repeat(64)}:${"1".repeat(64)}`;
  const staleResult = parseSnapshot(JSON.stringify(stale));
  assert.equal(staleResult.status, "style-mismatch");
  if (staleResult.status === "style-mismatch") {
    assert.equal(staleResult.snapshot.items.stable?.title, "Stable");
  }

  // A style records the locale it renders in, and a mismatch names the style as the style menu does.
  const british = JSON.parse(
    serializeSnapshot("harvard-cite-them-right", items),
  ) as { rendererRevision: string };
  const britishLocale = await readFile("public/assets/csl/locales-en-GB.xml");
  assert.ok(
    british.rendererRevision.endsWith(
      createHash("sha256").update(britishLocale).digest("hex"),
    ),
  );
  british.rendererRevision = stale.rendererRevision;
  const britishResult = parseSnapshot(JSON.stringify(british));
  assert.equal(britishResult.status, "style-mismatch");
  if (britishResult.status === "style-mismatch") {
    assert.match(
      britishResult.error,
      /a different Harvard \(Cite Them Right\) style or locale revision/,
    );
  }

  const tampered = parseSnapshot(
    embeddedSnapshot({ stable: { id: "different" } }, "ieee"),
  );
  assert.equal(tampered.status, "invalid");
  assert.equal(
    parseSnapshot(
      embeddedSnapshot(items, "ieee", { futureCritical: "preserve-me" }),
    ).status,
    "invalid",
    "unknown current-version root data must not be silently discarded",
  );
});

test("a snapshot this build cannot read never chooses a document's style", () => {
  const items = { a: { id: "a", title: "T" } };
  const valid = serializeSnapshot("ieee", items);
  const mismatched = JSON.parse(valid) as { rendererRevision: string };
  mismatched.rendererRevision = `${"0".repeat(64)}:${"1".repeat(64)}`;

  const statuses = new Set<SnapshotReadResult["status"]>();
  const claimFor = (raw: unknown): DocumentStyleClaim => {
    const result = parseSnapshot(raw);
    statuses.add(result.status);
    return documentStyleClaim(result);
  };

  assert.deepEqual(
    claimFor(undefined),
    { kind: "unclaimed" },
    "no snapshot is no claim — an ordinary first use, unless the document already cites, which the pane checks for itself (hasCitationTag)",
  );
  assert.deepEqual(claimFor(valid), { kind: "claimed", styleId: "ieee" });
  assert.deepEqual(
    claimFor(JSON.stringify(mismatched)),
    { kind: "claimed", styleId: "ieee" },
    "a moved CSL revision does not unsay which style the document is written in",
  );

  /* Both refused payloads carry a styleId this build reads perfectly well; they are refused whole, so that style does not decide the render. */
  assert.deepEqual(
    claimFor(embeddedSnapshot(items, "ieee", { version: 3 })),
    { kind: "unreadable" },
    "a snapshot from a newer build names a style, and this build must not act on it",
  );
  assert.deepEqual(
    claimFor('{"version":2,"styleId":"ieee",'),
    { kind: "unreadable" },
    "bytes that did not parse name a style too, in the part that did",
  );

  assert.equal(
    statuses.size,
    5,
    `these payloads reached ${statuses.size} of the snapshot reader's statuses (${[...statuses].join(", ")}); every one it can answer with must be placed here`,
  );
});

test("the embedded snapshot is read through parseJson, not JSON.parse", () => {
  // 13 is one past MAX_SNAPSHOT_DEPTH, written as a literal rather than derived from the bound, so the bound itself is what is pinned.
  const justOverTheBound = "[".repeat(13) + "]".repeat(13);
  const overBound = parseSnapshot(justOverTheBound);
  assert.equal(overBound.status, "invalid");
  if (overBound.status === "invalid") {
    assert.match(
      overBound.error,
      /citation data embedded in this document is nested too deeply/,
    );
  }
  // The bomb: 200 levels are counted in the text and refused before JSON.parse builds anything.
  const bomb = "[".repeat(200) + "]".repeat(200);
  const nested = parseSnapshot(bomb);
  assert.equal(nested.status, "invalid");
  if (nested.status === "invalid") {
    assert.match(
      nested.error,
      /citation data embedded in this document is nested too deeply/,
    );
  }

  // 900 KB is a literal past the 800 KB bound, so the size cap is pinned by a payload rather than by the constant that sets it.
  const oversized = `{"version":2,"pad":"${"p".repeat(900_000)}"}`;
  const tooBig = parseSnapshot(oversized);
  assert.equal(tooBig.status, "invalid");
  if (tooBig.status === "invalid") {
    assert.match(tooBig.error, /snapshot exceeds the safety limit/);
  }

  const broken = parseSnapshot('{"version":2,');
  assert.equal(broken.status, "invalid");
  if (broken.status === "invalid") {
    assert.match(
      broken.error,
      /citation data embedded in this document is not valid JSON/,
    );
  }

  // The reviver drops the prototype-chain names on the way in; the per-item RESERVED_FIELDS check still guards items built from a .bib on the way out, which never pass through the reader.
  const polluted = parseSnapshot(
    embeddedSnapshot(
      JSON.parse('{"a":{"id":"a","title":"T","__proto__":{"pwned":true}}}'),
    ),
  );
  assert.equal(polluted.status, "valid");
  assert.equal(({} as Record<string, unknown>).pwned, undefined);
  if (polluted.status === "valid") {
    assert.equal(polluted.snapshot.items.a?.pwned, undefined);
    // These records become the live library, and the reviver runs bottom-up rebuilding each object, which is what leaves them on null prototypes in place of a second clone pass.
    assert.equal(Object.getPrototypeOf(polluted.snapshot.items), null);
    assert.equal(Object.getPrototypeOf(polluted.snapshot.items.a), null);
    assert.equal("toString" in (polluted.snapshot.items.a as object), false);
  }

  // 200,001 array elements is one past MAX_SNAPSHOT_KEYS and 400 KB of text, half the byte cap, so the entry cap is what has to refuse it.
  const wide = `{"version":2,"items":[${"0,".repeat(200_000)}0]}`;
  assert.ok(wide.length < 800 * 1024, "the wide payload is over the byte cap");
  const tooMany = parseSnapshot(wide);
  assert.equal(tooMany.status, "invalid");
  if (tooMany.status === "invalid") {
    assert.match(tooMany.error, /has more than 200000 entries/);
  }

  // The options object is hand-built rather than taken from parseSnapshot, because what is pinned is the reader's own contract: a missing bound is a refusal, not a silent no-cap.
  assert.throws(
    () =>
      (parseJson as (t: string, o: unknown) => unknown)("{}", {
        maxBytes: 10,
        maxDepth: 2,
      }),
    /parseJson needs a positive maxKeys/,
  );

  // Nine levels is one past MAX_ITEM_DEPTH but inside the reader's own bound, so the tighter check with the better message is the one that answers.
  const deep: Record<string, unknown> = { id: "a", title: "T" };
  let cursor = deep;
  for (let i = 0; i < 9; i++) {
    const child: Record<string, unknown> = {};
    cursor.child = child;
    cursor = child;
  }
  const overItemDepth = parseSnapshot(embeddedSnapshot({ a: deep }));
  assert.equal(overItemDepth.status, "invalid");
  if (overItemDepth.status === "invalid") {
    assert.match(overItemDepth.error, /nesting deeper than 8 levels/);
  }
});

test("there is one snapshot format and no reader for any other", () => {
  const items = { a: { id: "a", title: "T" } };

  const bare = parseSnapshot(JSON.stringify(items));
  assert.equal(bare.status, "invalid", "a bare item record is not a snapshot");
  if (bare.status === "invalid")
    assert.match(bare.error, new RegExp(`carries no ${NAME} format version`));

  const v1 = parseSnapshot(
    JSON.stringify({ version: 1, styleId: "apa", items }),
  );
  assert.equal(v1.status, "invalid");
  if (v1.status === "invalid") {
    assert.match(v1.error, /format version 1/);
    assert.match(
      v1.error,
      new RegExp(`written by an earlier version of ${NAME}`),
    );
  }

  // A version no build ever wrote is not an earlier build's data, so the message must not claim it was.
  const nonsense = parseSnapshot(
    JSON.stringify({ version: "two", styleId: "apa", items }),
  );
  assert.equal(nonsense.status, "invalid");
  if (nonsense.status === "invalid") {
    assert.doesNotMatch(
      nonsense.error,
      new RegExp(`earlier version of ${NAME}`),
    );
  }

  assert.equal(
    parseSnapshot(JSON.stringify({ version: 99, styleId: "apa", items }))
      .status,
    "unsupported",
  );
});

test("saving a snapshot never drops metadata for a still-cited key", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(
    "@article{einstein1905, author={Einstein, Albert}, title={Zur}, year={1905}}",
  );

  const previous = serializeSnapshot("apa", {
    einstein1905: { id: "einstein1905", title: "Zur" },
    shannon1948: {
      id: "shannon1948",
      title: "A Mathematical Theory of Communication",
    },
  });
  const office = installOffice(previous);
  installWord(["bibcite:einstein1905", "bibcite:shannon1948"]);

  try {
    const report = await saveSnapshot(engine);
    assert.equal(report.embedded, 1);
    assert.deepEqual(report.preserved, ["shannon1948"]);

    const written = parseSnapshot(office.stored());
    assert.equal(written.status, "valid");
    if (written.status === "valid") {
      assert.equal(
        written.snapshot.items.shannon1948?.title,
        "A Mathematical Theory of Communication",
      );
      assert.equal(written.snapshot.items.einstein1905?.title, "Zur");
    }
  } finally {
    clearHostMocks();
  }
});

// Derived over every inherited name rather than spot-checking `toString`: the hole was the prototype, not one word on it.
test("an unresolved key that names an inherited property cannot block the save", async () => {
  for (const key of [...Object.getOwnPropertyNames(Object.prototype)]) {
    const engine = new CitationEngine("apa", loadAsset);
    engine.load(
      "@article{einstein1905, author={Einstein, Albert}, title={Zur}, year={1905}}",
    );
    const office = installOffice();
    installWord(["bibcite:einstein1905", `bibcite:${key}`]);
    try {
      const report = await saveSnapshot(engine);
      assert.equal(report.embedded, 1, `${key} must not block the save`);
      assert.deepEqual(report.preserved, []);
      const written = parseSnapshot(office.stored());
      assert.equal(written.status, "valid");
      if (written.status === "valid") {
        assert.equal(written.snapshot.items.einstein1905?.title, "Zur");
      }
    } finally {
      clearHostMocks();
    }
  }
});

// The set is derived here as it is in the code, so a name a future runtime adds to Object.prototype is covered by both.
test("a record key the object model reserves is refused, never crashed on", () => {
  const reserved = [
    ...Object.getOwnPropertyNames(Object.prototype),
    "prototype",
  ];
  assert.ok(reserved.length > 5);

  for (const key of reserved) {
    const engine = new CitationEngine("apa", loadAsset);
    assert.throws(
      () => engine.load(`@article{${key}, title={T}, year={2020}}`),
      new RegExp(`Entry "${key.replace(/[$]/g, "\\$&")}".*reserves`, "s"),
      `a .bib entry keyed ${key} must be refused, not handed to citeproc`,
    );
    assert.equal(engine.size, 0, "and the refusal stays transactional");
    assert.throws(
      () => serializeSnapshot("apa", { [key]: { id: key, title: "T" } }),
      /reserves/,
      `a snapshot record keyed ${key} must never be written`,
    );
    assert.throws(
      () => engine.loadSnapshot({ [key]: { id: key, title: "T" } }),
      /reserves/,
    );
  }

  // Why nothing may write one: the reader hands back a snapshot it calls valid with the record simply gone, which nothing downstream can notice.
  const smuggled = embeddedSnapshot({
    constructor: { id: "constructor", title: "Vanishes" },
  });
  const read = parseSnapshot(smuggled);
  assert.equal(read.status, "valid");
  if (read.status === "valid") {
    assert.deepEqual(Object.keys(read.snapshot.items), []);
  }

  // A key that merely looks awkward is untouched; this bounds records, not keys.
  const fine = new CitationEngine("apa", loadAsset);
  fine.load("@article{prototype2020, title={T}, year={2020}}");
  assert.deepEqual(fine.keys(), ["prototype2020"]);
  assert.deepEqual(parseCitationTag("bibcite:toString"), {
    keys: ["toString"],
  });
});

// What moved is the style/locale revision the records were rendered against, not the format.
test("re-saving a renderer-mismatched snapshot preserves its embedded records", async () => {
  const stale = JSON.parse(
    serializeSnapshot("apa", {
      carried: { id: "carried", title: "Carried metadata" },
    }),
  ) as { rendererRevision: string };
  stale.rendererRevision = `${"0".repeat(64)}:${"1".repeat(64)}`;

  const engine = new CitationEngine("apa", loadAsset);
  engine.load("@article{current, title={Current}, year={2020}}");
  const office = installOffice(JSON.stringify(stale));
  installWord(["bibcite:carried"]);

  try {
    const report = await saveSnapshot(engine);
    assert.deepEqual(report.preserved, ["carried"]);
    const resaved = parseSnapshot(office.stored());
    assert.equal(resaved.status, "valid");
    if (resaved.status === "valid") {
      assert.equal(resaved.snapshot.items.carried?.title, "Carried metadata");
    }
  } finally {
    clearHostMocks();
  }
});

test("a failed post-insert refresh reports partial success instead of an insert failure", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  installOffice();
  // Word.run 1 is the document-wide refresh that follows the insert; failing it models a host hiccup mid-renumber.
  installWord(["bibcite:einstein1905"], (call) => {
    if (call === 1) throw new Error("host refresh unavailable");
  });

  try {
    const result = await insertCitation(engine, ["einstein1905"]);
    assert.deepEqual(result.missing, []);
    assert.match(result.refreshError ?? "", /host refresh unavailable/);
  } finally {
    clearHostMocks();
  }
});

// Only what the add-in asks of Word is checkable here; where real Word leaves the cursor is the manual check at the top of wordDoc.ts.
test("inserting a citation moves the cursor past it, never at the citation's cost", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  try {
    const host = installWordItems([]);
    await insertCitation(engine, ["einstein1905"]);
    const [inserted] = host.controls.items;
    assert.equal(inserted?.selected, "after");

    // Sync 4 is the cursor move, after the citation is committed; a host refusing it must not cost the citation.
    const refusing = installWordItems([], {
      onSync(call) {
        if (call === 4) throw new Error("simulated selection refusal");
      },
    });
    const result = await insertCitation(engine, ["einstein1905"]);
    assert.equal(result.refreshError, undefined);
    assert.deepEqual(
      refusing.controls.items.map((control) => control.text),
      ["[1]"],
    );
  } finally {
    clearHostMocks();
  }
});

test("a bibliography is never overwritten with nothing", async () => {
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(
    "@article{unrelated, author={Z, Y}, title={U}, journal={J}, year={2010}}",
  );
  await engine.ensureReady();

  const citation = makeWordControl("bibcite:shannon1948", "[1]");
  const bibliography = makeWordControl(
    "bibcite-bibliography",
    "[1] C. E. Shannon, …",
  );
  installWordItems([citation, bibliography]);
  let report = await refreshDocument(engine);
  assert.equal(
    citation.text,
    "[?shannon1948]",
    "the citation still degrades visibly",
  );
  assert.equal(
    bibliography.text,
    "[1] C. E. Shannon, …",
    "the bibliography survives untouched",
  );
  assert.equal(report.bibliographyUpdated, false);
  assert.equal(report.bibliographySkipped, true);

  const orphan = makeWordControl("bibcite-bibliography", "[1] Existing entry");
  installWordItems([orphan]);
  report = await refreshDocument(engine);
  assert.equal(orphan.text, "[1] Existing entry");
  assert.equal(report.citationsUpdated, 0);
  assert.equal(report.bibliographySkipped, true);

  installWordItems([makeWordControl("bibcite:shannon1948", "[1]")]);
  await assert.rejects(() => insertBibliography(engine), /nothing to list/);

  clearHostMocks();
});

test("a document citing only letters in APA keeps its bibliography and says why", async () => {
  const engine = new CitationEngine("apa", loadAsset);
  engine.load(
    "@letter{doe2024, author={Doe, Jane}, title={Letter to the board}, year={2024}}\n" +
      "@book{roe2020, author={Roe, Rick}, title={A Book}, year={2020}}\n",
  );
  await engine.ensureReady();

  try {
    const citation = makeWordControl("bibcite:doe2024", "(old)");
    const bibliography = makeWordControl("bibcite-bibliography", "Existing");
    installWordItems([citation, bibliography]);
    const report = await refreshDocument(engine);
    assert.equal(citation.text, "(J. Doe, personal communication, 2024)");
    assert.equal(bibliography.text, "Existing");
    assert.equal(report.bibliographySkipped, true);
    assert.deepEqual(report.missingKeys, []);

    installWordItems([makeWordControl("bibcite:doe2024", "(old)")]);
    await assert.rejects(
      () => insertBibliography(engine),
      /Every citation in this document is to a work APA cites in the text only, such as a personal communication, so there is nothing to list\./,
    );

    // Beside a work APA does list, the letter is cited and the list carries only the book.
    const letter = makeWordControl("bibcite:doe2024", "(old)");
    const book = makeWordControl("bibcite:roe2020", "(old)");
    const list = makeWordControl("bibcite-bibliography", "Existing");
    installWordItems([letter, book, list]);
    const mixed = await refreshDocument(engine);
    assert.equal(mixed.bibliographyUpdated, true);
    assert.equal(letter.text, "(J. Doe, personal communication, 2024)");
    assert.match(list.text, /Roe/);
    assert.doesNotMatch(list.text, /Letter to the board/);

    // Beside a key the library lacks, the letter still resolved, so neither message may say nothing did.
    const cited = makeWordControl("bibcite:doe2024", "(old)");
    const unknown = makeWordControl("bibcite:nobody1999", "(old)");
    const kept = makeWordControl("bibcite-bibliography", "Existing");
    installWordItems([cited, unknown, kept]);
    const partial = await refreshDocument(engine);
    assert.equal(kept.text, "Existing");
    assert.equal(partial.bibliographySkipped, true);
    assert.equal(partial.keysResolved, 1);
    assert.deepEqual(partial.missingKeys, ["nobody1999"]);

    installWordItems([
      makeWordControl("bibcite:doe2024", "(old)"),
      makeWordControl("bibcite:nobody1999", "(old)"),
    ]);
    await assert.rejects(
      () => insertBibliography(engine),
      /Every citation that resolves is to a work APA cites in the text only.*unresolved: nobody1999/,
    );

    // Chicago cites a letter in the text only too, and the reason names it, not APA.
    engine.setStyle("chicago-author-date");
    await engine.ensureReady();
    installWordItems([makeWordControl("bibcite:doe2024", "(old)")]);
    await assert.rejects(
      () => insertBibliography(engine),
      /Every citation in this document is to a work Chicago \(author-date\) cites in the text only/,
    );
  } finally {
    clearHostMocks();
  }
});

test("the Word mock refuses a read that precedes load() and sync()", async () => {
  const host = installWordItems([makeWordControl("bibcite:einstein1905")]);
  assert.throws(
    () => host.controls.items,
    /PropertyNotLoaded/,
    "read before load() must throw",
  );

  host.controls.load();
  assert.throws(
    () => host.controls.items,
    /PropertyNotLoaded/,
    "load() alone is not enough",
  );

  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();
  await refreshDocument(engine); // syncs, so the property becomes readable
  assert.equal(host.controls.items.length, 1);
  clearHostMocks();
});

/** `Word.BodyType` in `@types/office-js` is the only machine-readable statement of which stories exist, so the list is derived from it rather than spelled out and going stale as Word adds stories. */
async function wordStoryTypes(): Promise<string[]> {
  const declaration = "node_modules/@types/office-js/index.d.ts";
  const source = await readFile(declaration, "utf8");
  const found = [...source.matchAll(/\benum BodyType \{\n([\s\S]*?)\n {4}\}/g)];
  if (found.length !== 1) {
    throw new Error(
      `expected exactly one Word.BodyType declaration in ${declaration}, found ${found.length}`,
    );
  }
  const stories = [...found[0]![1]!.matchAll(/^ +\w+ = "(\w+)",?$/gm)].map(
    (match) => match[1]!,
  );
  if (stories.length < 2 || !stories.includes("MainDoc")) {
    throw new Error(
      `read ${JSON.stringify(stories)} out of Word.BodyType; this gate needs the main story ` +
        "Word names MainDoc — which is what makeWordControl defaults to — and at least one other",
    );
  }
  return stories;
}

/** Both sets are derived — stories from `Word.BodyType`, operations from `wordDoc`'s exports — but no mock can settle whether real Word's `document.contentControls` reaches a footnote, so that stays a manual check (see wordDoc.ts). */
test("every document read spans every story Word has, not just the main body", async () => {
  const stories = await wordStoryTypes();
  const engine = new CitationEngine("ieee", loadAsset);
  engine.load(await readFile("sample/library.bib", "utf8"));
  await engine.ensureReady();

  // `ghost` is cited by nothing and must still be pruned, so the prune is asserted alongside the record it must not touch.
  const embedded = serializeSnapshot("apa", {
    einstein1905: { id: "einstein1905", title: "TITLE-einstein" },
    ghost: { id: "ghost", title: "TITLE-ghost-c0nfidential" },
  });

  interface Probe {
    story: string;
    citation: ReturnType<typeof makeWordControl>;
    bibliography: ReturnType<typeof makeWordControl>;
    office: ReturnType<typeof installOffice>;
  }

  const invisible = (probe: Probe, what: string) =>
    `${what} — a ${PRODUCT_NAME} citation in the ${probe.story} story was invisible to it`;

  const probes: Record<keyof typeof wordDoc, (probe: Probe) => Promise<void>> =
    {
      documentControlTags: async (probe) => {
        assert.ok(
          (await wordDoc.documentControlTags()).includes(
            "bibcite:einstein1905",
          ),
          invisible(
            probe,
            "a document without a snapshot would open with no style hold",
          ),
        );
      },

      readSnapshot: (probe) => {
        // This one reads settings and no content control, so the story axis cannot bite; it is driven anyway because the operation set is derived from the module's exports rather than chosen.
        assert.equal(wordDoc.readSnapshot().status, "valid");
        assert.equal(probe.office.saves(), 0);
        return Promise.resolve();
      },

      reconcileSnapshot: async (probe) => {
        const removed = await wordDoc.reconcileSnapshot();
        assert.deepEqual(
          removed,
          ["ghost"],
          invisible(
            probe,
            "opening the pane destroyed a record the document still cites",
          ),
        );
        const after = probe.office.stored();
        assert.ok(after !== undefined);
        assert.ok(
          after.includes("TITLE-einstein"),
          invisible(probe, "the cited record was pruned out of the document"),
        );
        assert.ok(
          !after.includes("TITLE-ghost"),
          "the uncited record survived the prune",
        );
      },

      saveSnapshot: async (probe) => {
        const report = await wordDoc.saveSnapshot(engine);
        assert.equal(
          report.embedded,
          1,
          invisible(probe, "the snapshot write embedded no record"),
        );
        const after = probe.office.stored();
        assert.ok(after !== undefined && after.includes("einstein1905"));
        assert.ok(
          !after.includes("TITLE-ghost"),
          "the uncited record survived the write",
        );
      },

      refreshDocument: async (probe) => {
        const report = await wordDoc.refreshDocument(engine);
        assert.equal(
          report.citationsUpdated,
          1,
          invisible(probe, "the refresh rewrote no citation"),
        );
        assert.equal(
          probe.citation.text,
          "[1]",
          invisible(probe, "the citation kept its stale text"),
        );
        assert.match(
          probe.bibliography.text,
          /Einstein/,
          invisible(probe, "the bibliography was left without its entry"),
        );
      },

      insertBibliography: async (probe) => {
        assert.deepEqual(await wordDoc.insertBibliography(engine), []);
        assert.match(
          probe.bibliography.text,
          /Einstein/,
          invisible(probe, "the bibliography did not list the work it cites"),
        );
      },

      insertCitation: async (probe) => {
        const result = await wordDoc.insertCitation(engine, [
          "knuth1984texbook",
        ]);
        assert.deepEqual(result.missing, []);
        assert.equal(result.refreshError, undefined);
        assert.equal(
          probe.citation.text,
          "[1]",
          invisible(
            probe,
            "the document-wide renumber skipped an existing citation",
          ),
        );
      },
    };

  assert.deepEqual(
    Object.keys(probes).sort(),
    Object.keys(wordDoc).sort(),
    "every operation wordDoc exports has to be driven here",
  );

  // Every pair runs and every failure is kept, so a scope widened story by story reports every story it still misses at once.
  const failures: string[] = [];
  for (const story of stories) {
    for (const [operation, drive] of Object.entries(probes)) {
      const citation = makeWordControl("bibcite:einstein1905", "STALE", story);
      const bibliography = makeWordControl(
        "bibcite-bibliography",
        "STALE LIST",
        story,
      );
      const office = installOffice(embedded);
      const host = installWordItems([citation, bibliography]);
      try {
        await drive({ story, citation, bibliography, office });
        assert.equal(
          host.bodyReads(),
          0,
          "it asked document.body.contentControls, which is the main story alone",
        );
      } catch (error) {
        failures.push(
          `${operation}, citation in the ${story} story: ${(error as Error).message.split("\n")[0]}`,
        );
      } finally {
        clearHostMocks();
      }
    }
  }
  assert.deepEqual(
    failures,
    [],
    `${failures.length} of ${stories.length * Object.keys(probes).length} story/operation pairs:\n  ${failures.join("\n  ")}`,
  );
});

test("a name list BibTeX can express and CSL cannot is reported, not dropped", () => {
  // "and others" is preserved rather than dropped, because citeproc counts names to choose between "Doe et al." and "Doe & Roe".
  const engine = new CitationEngine("apa", loadAsset);
  const truncated = engine.load(
    "@article{o, author={Doe, Jane and others}, title={T}, year={2000}}",
  );
  assert.equal(truncated.fidelity.length, 1);
  assert.match(truncated.fidelity[0] ?? "", /and others/);
  assert.equal(
    (engine.item("o")?.author ?? []).length,
    2,
    "the marker is preserved, not discarded",
  );

  assert.deepEqual(
    engine.load("@article{ok, author={Doe, Jane}, title={T}, year={2000}}")
      .fidelity,
    [],
  );
});

test("a snapshot from a newer build is never overwritten", async () => {
  const future = JSON.stringify({
    version: 3,
    styleId: "apa",
    items: {},
    extra: "future data",
  });
  const result = parseSnapshot(future);
  assert.equal(result.status, "unsupported");

  const engine = new CitationEngine("apa", loadAsset);
  engine.load("@article{a, author={Doe, J}, title={T}, year={2000}}");
  const office = installOffice(future);
  installWord(["bibcite:a"]);
  try {
    await assert.rejects(
      () => saveSnapshot(engine),
      new RegExp(`newer version of ${NAME}`),
    );
    assert.equal(
      office.stored(),
      future,
      "the newer snapshot is left byte-for-byte intact",
    );
  } finally {
    clearHostMocks();
  }
});

test("an unreadable ('invalid') snapshot blocks the save instead of silently wiping it", async () => {
  // One malformed record makes the WHOLE snapshot unreadable — the shape a corrupted or hand-edited .docx can produce; a wrong-typed title (not a reserved key) is used so the prototype-pollution guard doesn't drop the field before validateCslItem sees it.
  const corrupted = embeddedSnapshot({
    alpha: { id: "alpha", title: "Alpha" },
    beta: { id: "beta", title: 12345 },
  });
  assert.equal(parseSnapshot(corrupted).status, "invalid");

  const engine = new CitationEngine("apa", loadAsset);
  // Nothing in the live library resolves either key (e.g. no .bib is open) — exactly the state that let the old code treat the unreadable snapshot as empty and overwrite it.
  const office = installOffice(corrupted);
  installWord(["bibcite:alpha"]);

  try {
    await assert.rejects(() => saveSnapshot(engine), /invalid/i);
    assert.equal(
      office.stored(),
      corrupted,
      "the unreadable snapshot must be left byte-for-byte intact, not wiped",
    );
    assert.equal(office.saves(), 0);
  } finally {
    clearHostMocks();
  }
});

test("a save starts from what a co-author's pane saved, not from this pane's stale copy", async () => {
  const engineFor = (bib: string) => {
    const engine = new CitationEngine("apa", loadAsset);
    engine.load(bib);
    return engine;
  };
  const withX = engineFor("@article{x, title={From A}, year={2020}}");
  const withY = engineFor("@article{y, title={From B}, year={2021}}");

  // B's pane opened before A's first save, so its own copy holds no snapshot.
  const shared = coauthoredDocument();
  const paneB = shared.openPane();
  const paneA = shared.openPane();
  try {
    installWord(["bibcite:x"]);
    paneA.activate();
    await saveSnapshot(withX);

    // B cites y from a .bib without x; A's citation of x is in the shared document.
    installWord(["bibcite:x", "bibcite:y"]);
    paneB.activate();
    const report = await saveSnapshot(withY);
    assert.deepEqual(report.preserved, ["x"]);
    const written = parseSnapshot(shared.stored());
    assert.equal(written.status, "valid");
    if (written.status === "valid") {
      assert.deepEqual(Object.keys(written.snapshot.items).sort(), ["x", "y"]);
      assert.equal(written.snapshot.items.x?.title, withX.item("x")?.title);
    }
  } finally {
    clearHostMocks();
  }

  // A snapshot from a newer build, saved after this pane opened, is refused as one already there would be.
  const future = JSON.stringify({ version: 3, styleId: "apa", items: {} });
  const later = coauthoredDocument();
  const pane = later.openPane();
  const newer = later.openPane();
  newer.settings.set(SNAPSHOT_SETTING, future);
  newer.settings.saveAsync(() => {});
  try {
    installWord(["bibcite:y"]);
    pane.activate();
    await assert.rejects(
      () => saveSnapshot(withY),
      new RegExp(`newer version of ${NAME}`),
    );
    assert.equal(later.stored(), future, "the newer snapshot was overwritten");
  } finally {
    clearHostMocks();
  }
});

test("the refresh before a save cannot cost a record held only in memory, nor block the save", async () => {
  const engineFor = (bib: string) => {
    const engine = new CitationEngine("apa", loadAsset);
    engine.load(bib);
    return engine;
  };
  const withZ = engineFor("@article{z, title={Zeta}, year={2019}}");
  const withY = engineFor("@article{y, title={Why}, year={2021}}");

  // A failed save leaves its record in this pane's copy alone, which the refresh replaces.
  const shared = coauthoredDocument();
  const pane = shared.openPane();
  try {
    installWord(["bibcite:z"]);
    pane.activate();
    pane.failNextSave();
    await assert.rejects(() => saveSnapshot(withZ), /could not save/);
    const report = await saveSnapshot(withY);
    assert.deepEqual(report.preserved, ["z"]);
    const written = parseSnapshot(shared.stored());
    assert.equal(written.status, "valid");
    if (written.status === "valid") {
      assert.deepEqual(Object.keys(written.snapshot.items), ["z"]);
    }
  } finally {
    clearHostMocks();
  }

  // A refresh that fails leaves the save to go ahead from this pane's copy, as it would on a host with no refresh at all.
  for (const refresh of ["fails", "throws"] as const) {
    const kept = coauthoredDocument(
      serializeSnapshot("apa", { z: { id: "z", title: "Zeta" } }),
    );
    kept.openPane(refresh).activate();
    installWord(["bibcite:z", "bibcite:y"]);
    try {
      const report = await saveSnapshot(withY);
      assert.deepEqual(report.preserved, ["z"], refresh);
      const written = parseSnapshot(kept.stored());
      assert.equal(written.status, "valid", refresh);
      if (written.status === "valid") {
        assert.deepEqual(
          Object.keys(written.snapshot.items).sort(),
          ["y", "z"],
          refresh,
        );
      }
    } finally {
      clearHostMocks();
    }
  }
});

// The empty cited set is a member of the space, not an exception to it.
test("a document embeds records for what it cites and for nothing else", async () => {
  const universe = ["alpha", "beta", "gamma"];
  const titleOf = (key: string) => `TITLE-${key}-c0nfidential`;
  const entry = (key: string) =>
    `@article{${key}, author={Doe, J}, title={${titleOf(key)}}, journal={J}, year={2000}}`;
  const records = Object.fromEntries(
    universe.map((key) => [key, { id: key, title: titleOf(key) }]),
  );
  const previous = serializeSnapshot("apa", records);

  const subsets = universe.reduce<string[][]>(
    (all, key) => [...all, ...all.map((subset) => [...subset, key])],
    [[]],
  );
  assert.equal(subsets.length, 8);

  // Written with citationTag rather than by hand, so a tag form the encoder gains is exercised by the same call the pane makes.
  const layouts = (cited: string[]): { name: string; tags: string[] }[] => [
    {
      name: "one control each",
      tags: cited.map((key) => citationTag([key])),
    },
    {
      name: "one cluster",
      tags: cited.length > 0 ? [citationTag(cited)] : [],
    },
    {
      name: "each cited twice",
      tags: cited.flatMap((key) => [citationTag([key]), citationTag([key])]),
    },
  ];

  for (const cited of subsets) {
    for (const resolvable of subsets) {
      for (const layout of layouts(cited)) {
        const label = `cited [${cited.join(",")}] resolvable [${resolvable.join(",")}] as ${layout.name}`;
        const engine = new CitationEngine("apa", loadAsset);
        // No .bib open at all is a state a user reaches by reopening the document without picking a file.
        if (resolvable.length > 0)
          engine.load(resolvable.map(entry).join("\n"));
        const office = installOffice(previous);
        installWord(layout.tags);
        try {
          await saveSnapshot(engine);
          const raw = office.stored() ?? "";
          const written = parseSnapshot(raw);
          assert.equal(written.status, "valid", label);
          if (written.status === "valid") {
            assert.deepEqual(
              Object.keys(written.snapshot.items).sort(),
              [...cited].sort(),
              label,
            );
          }
          for (const key of universe) {
            assert.equal(
              raw.includes(titleOf(key)),
              cited.includes(key),
              `${label}: ${titleOf(key)} in the bytes written to the .docx`,
            );
          }
        } finally {
          clearHostMocks();
        }
      }
    }
  }
});

// An unreadable tag is a different answer from a document that cites nothing, which is what makes the empty-cited-set save above safe; each tag here is first checked to be one the reader refuses.
test("a citation tag this build cannot read refuses the save, keeping every record", async () => {
  const previous = serializeSnapshot("apa", {
    kept: { id: "kept", title: "TITLE-kept" },
  });
  const refused = [
    "bibcite:",
    'bibcite:2;{"k":"kept","l":"5"}',
    "bibcite:has space",
    "bibcite:kept,kept",
    `bibcite:${"x".repeat(300)}`,
  ];

  for (const tag of refused) {
    assert.throws(
      () => parseCitationTag(tag),
      `the tag reader accepted ${tag}, so this case proves nothing`,
    );
    const engine = new CitationEngine("apa", loadAsset);
    engine.load(
      "@article{kept, author={Doe, J}, title={TITLE-kept}, journal={J}, year={2000}}",
    );
    const office = installOffice(previous);
    installWord([tag]);
    try {
      await assert.rejects(() => saveSnapshot(engine));
      assert.equal(
        office.stored(),
        previous,
        `${tag} cleared the embedded records`,
      );
      assert.equal(office.saves(), 0, `${tag} wrote to the document`);
    } finally {
      clearHostMocks();
    }
  }
});

// reconcileSnapshot runs at pane open, before the user has asked for anything, so what it may touch at all is the point.
test("opening the pane takes out records for citations the document lost", async () => {
  const kept = { id: "kept", title: "TITLE-kept" };
  const deleted = { id: "deleted", title: "TITLE-deleted-c0nfidential" };
  const current = serializeSnapshot("apa", { kept, deleted });
  const otherRevision = `${"a".repeat(64)}:${"b".repeat(64)}`;
  const mismatched = JSON.stringify({
    ...(JSON.parse(current) as Record<string, unknown>),
    rendererRevision: otherRevision,
  });

  const untouched =
    (stored: string | undefined) =>
    (removed: string[], after: string | undefined, saves: number) => {
      assert.deepEqual(removed, []);
      assert.equal(
        after,
        stored,
        "bytes this build must not touch were rewritten",
      );
      assert.equal(
        saves,
        0,
        "a snapshot this build must not touch was written",
      );
    };
  const pruned =
    (expectStatus: "valid" | "style-mismatch") =>
    (removed: string[], after: string | undefined, saves: number) => {
      assert.deepEqual(removed, ["deleted"]);
      assert.equal(saves, 1);
      assert.ok(after !== undefined);
      assert.ok(
        !after.includes(deleted.title),
        "the deleted citation's metadata is still in the document",
      );
      assert.ok(after.includes(kept.title), "a still-cited record was dropped");
      const reread = parseSnapshot(after);
      // The envelope stays the document's own: pruning re-derives no record, so re-stamping the revision would clear a mismatch warning it is owed.
      assert.equal(reread.status, expectStatus);
      if (reread.status === "valid" || reread.status === "style-mismatch") {
        assert.equal(reread.snapshot.styleId, "apa");
      }
    };

  const scenarios: Record<
    SnapshotReadResult["status"],
    {
      stored: string | undefined;
      expect: (
        removed: string[],
        after: string | undefined,
        saves: number,
      ) => void;
    }
  > = {
    none: { stored: undefined, expect: untouched(undefined) },
    valid: { stored: current, expect: pruned("valid") },
    "style-mismatch": { stored: mismatched, expect: pruned("style-mismatch") },
    unsupported: {
      stored: JSON.stringify({ version: 99, styleId: "apa", items: {} }),
      expect: (removed, after, saves) =>
        untouched(JSON.stringify({ version: 99, styleId: "apa", items: {} }))(
          removed,
          after,
          saves,
        ),
    },
    invalid: {
      stored: "{ not a snapshot",
      expect: (removed, after, saves) =>
        untouched("{ not a snapshot")(removed, after, saves),
    },
  };

  for (const [status, scenario] of Object.entries(scenarios)) {
    const office = installOffice(scenario.stored);
    installWord([citationTag(["kept"])]);
    try {
      const removed = await reconcileSnapshot();
      scenario.expect(removed, office.stored(), office.saves());
    } catch (error) {
      throw new Error(`${status}: ${(error as Error).message}`);
    } finally {
      clearHostMocks();
    }
  }

  // A document whose records already match its citations is not written to at all, or opening the pane would mark it as modified.
  const matching = serializeSnapshot("apa", { kept });
  const office = installOffice(matching);
  installWord([citationTag(["kept"])]);
  try {
    assert.deepEqual(await reconcileSnapshot(), []);
    assert.equal(office.stored(), matching);
    assert.equal(office.saves(), 0, "an unchanged document was written to");
  } finally {
    clearHostMocks();
  }

  const office2 = installOffice(current);
  installWord(["bibcite:"]);
  try {
    await assert.rejects(() => reconcileSnapshot());
    assert.equal(office2.stored(), current);
    assert.equal(office2.saves(), 0);
  } finally {
    clearHostMocks();
  }
});

/** `ESCAPE_ROUTES` in `tests/artefact.test.ts` measures the shipped chunk for a second way out; this measures what goes down the one there is, over every key handed to `settings.set` and every operation `wordDoc` exports. */
test("nothing this add-in writes carries a state the document used to be in", async () => {
  const PRIOR = {
    "the text a control held before the batch": "PRIOR-TEXT-c0nfidential",
    "the title a control held before the batch": "PRIOR-TITLE-c0nfidential",
    "the record for a citation the document no longer holds":
      "PRIOR-RECORD-c0nfidential",
    "an entry of the library a later load replaced":
      "PRIOR-LIBRARY-c0nfidential",
  };
  const CURRENT = "CURRENT-RECORD-live";
  const entry = (key: string, title: string) =>
    `@article{${key}, author={Doe, J}, title={${title}}, journal={J}, year={2000}}`;

  const embedded = serializeSnapshot("apa", {
    live: { id: "live", title: CURRENT },
    deleted: {
      id: "deleted",
      title: PRIOR["the record for a citation the document no longer holds"],
    },
  });

  const drive: Record<
    keyof typeof wordDoc,
    (engine: CitationEngine) => Promise<unknown>
  > = {
    documentControlTags: () => wordDoc.documentControlTags(),
    readSnapshot: () => Promise.resolve(wordDoc.readSnapshot()),
    reconcileSnapshot: () => wordDoc.reconcileSnapshot(),
    saveSnapshot: (engine) => wordDoc.saveSnapshot(engine),
    refreshDocument: (engine) => wordDoc.refreshDocument(engine),
    insertBibliography: (engine) => wordDoc.insertBibliography(engine),
    insertCitation: (engine) => wordDoc.insertCitation(engine, ["live"]),
  };
  assert.deepEqual(
    Object.keys(drive).sort(),
    Object.keys(wordDoc).sort(),
    "every operation wordDoc exports has to be driven here",
  );

  let sawWrite = false;
  let sawRecord = false;
  for (const [operation, run] of Object.entries(drive)) {
    const engine = new CitationEngine("apa", loadAsset);
    // Loaded and then replaced: the first library is a prior state of the pane, as a discarded undo entry would be.
    engine.load(
      entry("dropped", PRIOR["an entry of the library a later load replaced"]),
    );
    engine.load(entry("live", CURRENT));
    await engine.ensureReady();

    const citation = makeWordControl(
      "bibcite:live",
      PRIOR["the text a control held before the batch"],
    );
    citation.title = PRIOR["the title a control held before the batch"];
    const bibliography = makeWordControl(
      "bibcite-bibliography",
      `${PRIOR["the text a control held before the batch"]} list`,
    );
    bibliography.title = `${PRIOR["the title a control held before the batch"]} list`;

    const office = installOffice(embedded);
    installWordItems([citation, bibliography]);
    try {
      await run(engine);
      // The snapshot write the pane really does follow each operation with; without it the four that write nothing would pass vacuously.
      await saveSnapshot(engine);

      const wrote = office.written();
      assert.ok(wrote.length > 0, `${operation}: nothing was written at all`);
      for (const [key, value] of wrote) {
        assert.equal(
          key,
          SNAPSHOT_SETTING,
          `${operation}: wrote ${JSON.stringify(key)}, a key other than the snapshot, which this gate has never inspected`,
        );
        for (const [what, marker] of Object.entries(PRIOR)) {
          assert.ok(
            !value.includes(marker),
            `${operation}: ${what} reached the bytes this add-in writes into the .docx`,
          );
        }
        sawRecord ||= value.includes(CURRENT);
      }
      sawWrite = true;
    } finally {
      clearHostMocks();
    }
  }

  assert.ok(sawWrite, "no operation wrote anything, so this proves nothing");
  assert.ok(
    sawRecord,
    "nothing written carried the record for the key the document cites, so this proves nothing",
  );
});

test("snapshot validation covers the date fields citeproc parses structurally", () => {
  const ok = embeddedSnapshot({
    a: { id: "a", title: "T", issued: { "date-parts": [[2020, 1, 1]] } },
  });
  assert.equal(parseSnapshot(ok).status, "valid");

  for (const issued of [
    { "date-parts": [["not-a-number"]] },
    { "date-parts": "2020" },
    { "date-parts": [[2020, 1, 1, 1]] },
    "2020",
  ]) {
    const raw = embeddedSnapshot({ a: { id: "a", issued } });
    assert.equal(
      parseSnapshot(raw).status,
      "invalid",
      `accepted ${JSON.stringify(issued)}`,
    );
  }

  const deep: Record<string, unknown> = { id: "a", title: "T" };
  let cursor = deep;
  for (let i = 0; i < 40; i++) {
    const child: Record<string, unknown> = {};
    cursor.child = child;
    cursor = child;
  }
  assert.throws(
    () => serializeSnapshot("apa", { a: deep as never }),
    /nesting deeper than/,
  );
});

test("snapshot validation rejects every malformed CSL structural value", () => {
  const resultFor = (item: Record<string, unknown>) =>
    parseSnapshot(embeddedSnapshot({ a: { id: "a", ...item } }));

  for (const item of [
    { volume: ["12"] },
    { volume: true },
    { publisher: { name: "Publisher" } },
    { publisher: false },
    { type: 42 },
    { title: 42 },
    { author: [{ family: "Doe", future: { nested: true } }] },
    { categories: ["safe", { nested: true }] },
    { issued: { "date-parts": [] } },
  ]) {
    assert.equal(
      resultFor(item).status,
      "invalid",
      `accepted ${JSON.stringify(item)}`,
    );
  }

  assert.equal(
    resultFor({
      volume: "12",
      publisher: "Publisher",
      author: [{ family: "Doe", given: "Jane", "static-ordering": false }],
      categories: ["safe"],
    }).status,
    "valid",
  );
});

test("the build refuses to produce a manifest that cannot be published", async () => {
  // Built into a directory of its own because `dist/` is what `tests/artefact.test.ts` asserts about and the deploy workflow uploads, and this builds repeatedly with deliberately wrong environments.
  const outDir = await mkdtemp(join(tmpdir(), "bibcite-build-"));
  const manifest = join(outDir, "manifest.xml");

  // Cleared rather than inherited, or a developer with it exported would be testing their own environment.
  const run = (env: Record<string, string>) => {
    const inherited = { ...process.env };
    delete inherited.VITE_SITE_URL;
    return spawnSync(
      process.execPath,
      [
        "node_modules/vite/bin/vite.js",
        "build",
        "--outDir",
        outDir,
        "--emptyOutDir",
      ],
      {
        cwd: process.cwd(),
        env: { ...inherited, ...env },
        encoding: "utf8",
      },
    );
  };

  try {
    const good = run({ VITE_SITE_URL: "https://cite.example.org" });
    assert.equal(good.status, 0, good.stderr);
    const shipped = await readFile(manifest, "utf8");
    assert.doesNotMatch(shipped, /localhost|127\.0\.0\.1/);
    assert.match(shipped, /https:\/\/cite\.example\.org\/taskpane\.html/);

    // Certificate material is checked in the build output rather than by grepping the config, which would pass on a name and fail on a rename.
    const built = await readdir(outDir, { recursive: true });
    assert.deepEqual(
      built.filter((name) => /\.(pem|key|p12|pfx)$/i.test(name)),
      [],
      "a production build must not ship certificate or key material",
    );

    // A GitHub Pages project site lives under a path, so substituting only the origin would point every manifest URL at the account root.
    assert.equal(
      run({
        VITE_SITE_URL: "https://user.github.io/word-bibtex-citation-addin",
      }).status,
      0,
    );
    assert.match(
      await readFile(manifest, "utf8"),
      /DefaultValue="https:\/\/user\.github\.io\/word-bibtex-citation-addin\/taskpane\.html"/,
      "the deployment base path must survive substitution",
    );

    assert.notEqual(
      run({ VITE_SITE_URL: "http://cite.example.org" }).status,
      0,
      "http is rejected",
    );
    assert.notEqual(
      run({ VITE_SITE_URL: "not a url" }).status,
      0,
      "a malformed origin is rejected",
    );

    const silent = run({});
    assert.notEqual(silent.status, 0, "an unset VITE_SITE_URL still built");
    assert.match(silent.stderr, /VITE_SITE_URL is not set/);

    // A site URL is judged by the class its host falls in: 127.0.0.2 is as much the reader's own machine as 127.0.0.1, and .internal as much their own network.
    for (const host of [
      "https://localhost:3000",
      "https://127.0.0.2:3000",
      "https://[::1]:3000",
      "https://0x7f000001",
      "https://192.168.1.5:3000",
      "https://dev.internal",
      "https://buildbox:3000",
    ]) {
      const refused = run({ VITE_SITE_URL: host });
      assert.notEqual(refused.status, 0, `${host} was accepted as a site URL`);
      assert.match(refused.stderr, /cannot be published/);
    }
  } finally {
    await rm(outDir, { recursive: true, force: true });
  }
});
