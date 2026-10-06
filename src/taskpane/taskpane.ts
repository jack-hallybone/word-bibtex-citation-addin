import { BibSource, FilePickerBibSource } from "../lib/bibSource";
import { CitationEngine } from "../lib/citationEngine";
import { PRODUCT_NAME } from "../lib/productName";
import { MAX_FIELD_LENGTH, documentStyleClaim } from "../lib/snapshot";
import { BUILTIN_STYLES, CitationFormat, builtinStyle } from "../lib/styles";
import {
  RefreshReport,
  SnapshotSaveReport,
  documentControlTags,
  insertBibliography,
  insertCitation,
  readSnapshot,
  reconcileSnapshot,
  refreshDocument,
  saveSnapshot,
} from "../lib/wordDoc";
import {
  CITATION_TAG_PREFIX,
  CitationDetails,
  CitationForm,
  CslItem,
  LOCATOR_LABELS,
  LocatorLabel,
  MAX_LOCATOR_LENGTH,
  clipDisplayText,
  describeError,
  normaliseLocator,
  safeDisplayText,
  stripCslRichText,
} from "../lib/types";
import { shortAuthors } from "../names.js";

let engine = new CitationEngine();
const selected = new Set<string>();
let busy = false;
let pendingRender = 0;

/** Holds EVERY document write until the user says which style this document uses, since otherwise an unreadable snapshot, or citations with no snapshot at all, silently reformat an IEEE document into APA; only `onStyleChange` lifts it. */
let styleChoicePending = false;

/** The placeholder offered while `styleChoicePending`; not a style id, so `isStyleId` refuses it. */
const CHOOSE_STYLE = "";

const INTERACTIVE_IDS = [
  "open-bib",
  "style-select",
  "filter",
  "locator-label",
  "locator",
  "citation-form",
  "insert-citation",
  "clear-selection",
  "insert-bibliography",
  "refresh-doc",
] as const;

function element<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (!found) throw new Error(`Missing required element #${id}.`);
  return found as T;
}

// Display budgets for untrusted library text.
const MAX_ENTRY_LABEL = 280;
const MAX_ENTRY_NAME = 100;
const MAX_ENTRY_TITLE = 200;
const MAX_FIDELITY_DETAIL = 600;
const MAX_RENDERED_ENTRIES = 500;

// One action posts a summary plus up to seven detail notes, and must not evict its own summary.
const MAX_MESSAGES = 8;

// The label carries the status to a reader who cannot see the tone's colour.
const MESSAGE_TONES = {
  ok: { tone: "success", label: "Done" },
  warn: { tone: "warning", label: "Note" },
  err: { tone: "danger", label: "Problem" },
} as const;

function message(text: string, kind: keyof typeof MESSAGE_TONES): void {
  const box = element<HTMLElement>("messages");
  const { tone, label } = MESSAGE_TONES[kind];
  const notice = document.createElement("div");
  notice.className = `panel ${tone}`;
  const status = document.createElement("strong");
  status.className = "banner-title";
  status.textContent = label;
  notice.append(status, ` ${safeDisplayText(text)}`);
  if (kind === "err") notice.setAttribute("role", "alert");
  box.prepend(notice);
  while (box.children.length > MAX_MESSAGES) box.lastElementChild?.remove();
}

/** Exported for `tests/taskpane.test.ts`: pure, so it is tested directly rather than through a DOM this project has no mock for. */
export function entryLabel(item: CslItem): string {
  const who = shortAuthors(
    (item.author ?? []).map((author) =>
      clipDisplayText(
        stripCslRichText(author.family ?? author.literal ?? author.given ?? ""),
        MAX_ENTRY_NAME,
      ),
    ),
  );
  const year = item.issued?.["date-parts"]?.[0]?.[0];
  const title = clipDisplayText(
    stripCslRichText(item.title ?? ""),
    MAX_ENTRY_TITLE,
  );
  return clipDisplayText(
    [who, year, title].filter(Boolean).join(" · "),
    MAX_ENTRY_LABEL,
  );
}

/** Tooltip for the "Shortened" badge, naming which field(s) on this entry were cut to size. */
function truncationTooltip(fields: readonly string[]): string {
  return `${fields.join(", ")} longer than ${MAX_FIELD_LENGTH} characters, so shortened to fit.`;
}

function renderStyleOptions(): void {
  const select = element<HTMLSelectElement>("style-select");
  const styles = BUILTIN_STYLES.map(
    (style) => new Option(style.label, style.id),
  );
  if (styleChoicePending) {
    // A menu showing "APA" would claim a style the pane doesn't know, and would make the matching choice fire no change event.
    select.replaceChildren(
      new Option("Choose this document’s style…", CHOOSE_STYLE),
      ...styles,
    );
    select.value = CHOOSE_STYLE;
    return;
  }
  select.replaceChildren(...styles);
  select.value = engine.styleId;
}

export interface ActionControlState {
  insertCitationDisabled: boolean;
  clearSelectionDisabled: boolean;
  /** The page and form controls: a citation of several references takes neither, so they wait until one alone is selected. */
  detailsDisabled: boolean;
  /** Shared by "insert-bibliography" and "refresh-doc": both write the whole document, so neither needs a selection, but both wait on the style choice. */
  otherActionsDisabled: boolean;
}

/** Exported for `tests/taskpane.test.ts`: pure, so the enablement rule behind every document-writing button is tested directly rather than through a DOM this project has no mock for. */
export function actionControlState(
  selectedCount: number,
  busy: boolean,
  styleChoicePending: boolean,
): ActionControlState {
  return {
    insertCitationDisabled: busy || selectedCount === 0 || styleChoicePending,
    // Clearing the selection writes nothing, so it does not wait on the style choice.
    clearSelectionDisabled: busy || selectedCount === 0,
    // Nor does choosing a page or form, which may come before the reference it is for.
    detailsDisabled: busy || selectedCount > 1,
    otherActionsDisabled: busy || styleChoicePending,
  };
}

/** Exported for `tests/taskpane.test.ts`: pure, so the rule that holds the style choice for a document with no snapshot is tested directly rather than through a Word host. The prefix alone decides: parsing would throw on an earlier build's pinpoint tag, and a bibliography or a failed insertion's pending control is no citation. */
export function hasCitationTag(tags: readonly string[]): boolean {
  return tags.some((tag) => tag.startsWith(CITATION_TAG_PREFIX));
}

/** The state of every control whose press writes to the document, plus the selection count beside them. */
function updateActionControls(): void {
  const count = selected.size;
  element<HTMLElement>("selection-status").textContent = `${count} selected`;
  const state = actionControlState(count, busy, styleChoicePending);
  element<HTMLButtonElement>("insert-citation").disabled =
    state.insertCitationDisabled;
  element<HTMLButtonElement>("clear-selection").disabled =
    state.clearSelectionDisabled;
  for (const id of ["locator-label", "locator", "citation-form"] as const) {
    element<HTMLInputElement | HTMLSelectElement>(id).disabled =
      state.detailsDisabled;
  }
  element<HTMLElement>("details-note").hidden =
    count < 2 ||
    (element<HTMLElement>("locator-field").hidden &&
      element<HTMLElement>("form-field").hidden);
  for (const id of ["insert-bibliography", "refresh-doc"] as const) {
    element<HTMLButtonElement>(id).disabled = state.otherActionsDisabled;
  }
}

/** The form menu for a style that cites this way, each choice beside an example; "" is the standard form. Exported for `tests/taskpane.test.ts`: pure, so what each kind of style offers is tested directly. */
export function formOptions(
  format: CitationFormat,
): readonly (readonly [CitationForm | "", string])[] {
  switch (format) {
    case "author-date":
      return [
        ["", "Standard: (Smith, 2020)"],
        ["narrative", "Narrative: Smith (2020)"],
        ["suppress-author", "Year only: (2020)"],
      ];
    case "author":
      // MLA cites by name and page, so the page stands where another style's year would.
      return [
        ["", "Standard: (Smith 12)"],
        ["narrative", "Narrative: Smith (12)"],
        ["suppress-author", "Page only: (12)"],
      ];
    case "numeric":
      // A number is the whole citation, so there is no other form to offer.
      return [];
  }
}

/** Offer a form only where the current style has more than one, and a page only where it prints one. */
function renderCitationDetails(): void {
  const style = builtinStyle(engine.styleId);
  const options = formOptions(style.format);
  element<HTMLSelectElement>("citation-form").replaceChildren(
    ...options.map(([value, label]) => new Option(label, value)),
  );
  element<HTMLElement>("form-field").hidden = options.length === 0;
  element<HTMLElement>("locator-field").hidden = !style.pinpoints;
  updateActionControls();
}

/** The page and form the pane is set to, leaving out whatever the current style does not offer. */
function chosenDetails(): CitationDetails {
  const details: CitationDetails = {};
  const form = element<HTMLSelectElement>("citation-form").value;
  if (!element<HTMLElement>("form-field").hidden && form) {
    // Only ever one of formOptions' own values; the tag check refuses anything else.
    details.form = form as CitationForm;
  }
  const locator = normaliseLocator(element<HTMLInputElement>("locator").value);
  if (!element<HTMLElement>("locator-field").hidden && locator) {
    details.pinpoint = {
      label: element<HTMLSelectElement>("locator-label").value as LocatorLabel,
      locator,
    };
  }
  return details;
}

/** Exported for `tests/taskpane.test.ts`: pure, so the search rule (case-insensitive, matched against the key and its rendered label) is tested directly rather than through a DOM this project has no mock for. */
export function filterEntryKeys(
  engine: CitationEngine,
  rawQuery: string,
): string[] {
  const query = rawQuery.trim().toLocaleLowerCase();
  return engine.keys().filter((key) => {
    if (!query) return true;
    const item = engine.item(key);
    return (
      item !== undefined &&
      `${key} ${entryLabel(item)}`.toLocaleLowerCase().includes(query)
    );
  });
}

function renderEntryList(): void {
  pendingRender = 0;
  const list = element<HTMLUListElement>("entry-list");
  list.replaceChildren();

  const query = element<HTMLInputElement>("filter").value.trim();
  const matches = filterEntryKeys(engine, query);

  for (const key of matches.slice(0, MAX_RENDERED_ENTRIES)) {
    const item = engine.item(key);
    if (!item) continue;

    const row = document.createElement("li");
    row.classList.toggle("selected", selected.has(key));

    const label = document.createElement("label");
    label.className = "entry-option";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = selected.has(key);
    checkbox.setAttribute("aria-label", `Select ${key}`);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) selected.add(key);
      else selected.delete(key);
      row.classList.toggle("selected", checkbox.checked);
      updateActionControls();
    });

    const text = document.createElement("span");
    const keyText = document.createElement("span");
    keyText.className = "entry-key text-mono text-sm";
    keyText.textContent = `@${key}`;
    const metadata = document.createElement("span");
    metadata.className = "entry-meta text-muted text-sm";
    metadata.textContent = entryLabel(item) || "No author, year, or title";
    text.append(keyText, metadata);

    const truncatedFields = engine.truncatedFields(key);
    if (truncatedFields.length > 0) {
      const badge = document.createElement("span");
      badge.className = "badge warning";
      badge.textContent = "Shortened";
      badge.title = truncationTooltip(truncatedFields);
      text.append(badge);
    }

    label.append(checkbox, text);
    row.appendChild(label);
    list.appendChild(row);
  }

  if (matches.length === 0) {
    const note = document.createElement("li");
    note.className = "list-note text-muted text-sm";
    note.textContent = query
      ? "No references match this filter."
      : "No references to display.";
    list.appendChild(note);
  } else if (matches.length > MAX_RENDERED_ENTRIES) {
    const note = document.createElement("li");
    note.className = "list-note text-muted text-sm";
    note.textContent =
      `Showing the first ${MAX_RENDERED_ENTRIES} of ${matches.length} matches. ` +
      "Narrow the filter to find the rest.";
    list.appendChild(note);
  }
}

function scheduleEntryRender(): void {
  if (pendingRender) cancelAnimationFrame(pendingRender);
  pendingRender = requestAnimationFrame(renderEntryList);
}

function setBusy(nextBusy: boolean, activeId?: string): void {
  busy = nextBusy;
  element<HTMLElement>("app").setAttribute("aria-busy", String(nextBusy));
  for (const id of INTERACTIVE_IDS) {
    const control = element<
      HTMLInputElement | HTMLButtonElement | HTMLSelectElement
    >(id);
    control.disabled = nextBusy;
    const showSpinner =
      nextBusy && id === activeId && control instanceof HTMLButtonElement;
    control.classList.toggle("is-busy", showSpinner);
    // The theme's .spinner is a real element, not a pseudo-element: insert one to show progress, and take it out again rather than merely hide it.
    const spinner = control.querySelector(":scope > .spinner");
    if (showSpinner && !spinner) {
      const el = document.createElement("span");
      el.className = "spinner";
      el.setAttribute("aria-hidden", "true");
      control.append(el);
    } else if (!showSpinner && spinner) {
      spinner.remove();
    }
  }
  if (!nextBusy) updateActionControls();
}

/** Bind an async handler: `addEventListener` discards the promise a handler returns, so an escaping rejection would leave the pane silent. */
function onAsync(
  id: string,
  event: "click" | "change",
  handler: () => Promise<void>,
): void {
  element<HTMLElement>(id).addEventListener(event, () => {
    void handler().catch((error: unknown) => {
      message(
        `${PRODUCT_NAME} hit an unexpected problem: ${describeError(error)}`,
        "err",
      );
    });
  });
}

/** Runs one UI action under the global busy state, reporting any failure; the catch sits inside the try so a throw from `setBusy` cannot lose that report. */
async function withBusy(
  activeId: string,
  failure: string,
  action: () => Promise<void>,
): Promise<void> {
  if (busy) return;
  setBusy(true, activeId);
  try {
    await action();
  } catch (error) {
    message(`${failure}: ${describeError(error)}`, "err");
  } finally {
    setBusy(false);
  }
}

/** What a refresh in the style `styleLabel` names warns about after its summary line. Exported for `tests/taskpane.test.ts`, like keptRecordsNote below. */
export function refreshWarnings(
  report: RefreshReport,
  styleLabel: string,
): string[] {
  const warnings: string[] = [];
  if (report.citationsUpdated === 0) {
    // Nothing was rewritten, the bibliography included, so no note about the bibliography follows.
    warnings.push(
      `No ${PRODUCT_NAME} citations were found in this document, so nothing was rewritten. ` +
        "If you expected citations here, they may have been removed or converted to plain text.",
    );
  } else if (report.bibliographySkipped) {
    warnings.push(skippedBibliographyNote(report, styleLabel));
  }
  if (report.missingKeys.length > 0) {
    warnings.push(`Unresolved keys: ${report.missingKeys.join(", ")}.`);
  }
  return warnings;
}

/** Why a refresh left an existing bibliography as it was: nothing resolved, or each key that did is to a work the style cites in the text only. Exported for `tests/taskpane.test.ts`, like keptRecordsNote below. */
export function skippedBibliographyNote(
  report: RefreshReport,
  styleLabel: string,
): string {
  if (report.keysResolved === 0) {
    return (
      "The bibliography was left unchanged because no citation resolved against the current " +
      "library. Open the .bib that defines these keys, then refresh again."
    );
  }
  if (report.missingKeys.length === 0) {
    return (
      `The bibliography was left unchanged because every citation is to a work ${styleLabel} ` +
      "cites in the text only, such as a personal communication."
    );
  }
  return (
    "The bibliography was left unchanged because every citation that resolved is to a work " +
    `${styleLabel} cites in the text only, such as a personal communication. Open the .bib that ` +
    "defines the unresolved keys, then refresh again."
  );
}

/** Exported for `tests/taskpane.test.ts`: pure, so the wording is tested directly rather than through a DOM this project has no mock for. The kept records are never rendered from, since the .bib is the source of truth; the document only carries them. */
export function keptRecordsNote(keys: readonly string[]): string {
  return (
    `Kept this document’s embedded data for keys the current library does not define: ${keys.join(", ")}. ` +
    "Their citations show [?key] until you open a .bib that defines them, or close and reopen " +
    "this pane and choose “Refresh all citations” before opening a .bib."
  );
}

async function persistPortableSnapshot(): Promise<void> {
  let report: SnapshotSaveReport;
  try {
    report = await saveSnapshot(engine);
  } catch (error) {
    message(
      `Could not update the document’s portable citation data (${describeError(error)}). ` +
        "Save or share it only with the .bib file available.",
      "warn",
    );
    return;
  }
  if (report.preserved.length > 0)
    message(keptRecordsNote(report.preserved), "warn");
}

/** Take the document's embedded records for deleted citations out at pane open, and say so. */
async function removeRecordsForDeletedCitations(): Promise<void> {
  try {
    const removed = await reconcileSnapshot();
    if (removed.length > 0) {
      message(
        `Removed embedded metadata for ${removed.length} citation(s) this document no longer ` +
          "contains. Save the document to keep it out of the file.",
        "warn",
      );
    }
  } catch (error) {
    message(
      `Could not check this document’s embedded citation data (${describeError(error)}). ` +
        "It may still hold metadata for citations that were deleted.",
      "warn",
    );
  }
}

/** Report a refresh, escalating when nothing was updated or a write was skipped. */
function reportRefresh(report: RefreshReport, prefix: string): void {
  let text = `${prefix} ${report.citationsUpdated} citation(s)`;
  if (report.bibliographyUpdated) text += " and the bibliography";
  const notable = report.citationsUpdated === 0 || report.bibliographySkipped;
  message(`${text}.`, notable || report.missingKeys.length > 0 ? "warn" : "ok");

  for (const warning of refreshWarnings(
    report,
    builtinStyle(engine.styleId).label,
  ))
    message(warning, "warn");
}

/** Surface source/render differences as ONE message, so many affected entries cannot push the refresh summary out of the list. */
function reportFidelity(notes: readonly string[]): void {
  if (notes.length === 0) return;
  const shown = notes.slice(0, 3).join(" ");
  const rest =
    notes.length > 3 ? ` Plus ${notes.length - 3} more like this.` : "";
  message(
    `Some of this .bib will not render exactly as written. ${clipDisplayText(shown, MAX_FIDELITY_DETAIL)}${rest}`,
    "warn",
  );
}

/** Reveal the library and document panels, with `text` as the source status. */
function showLibrary(text: string): void {
  const status = element<HTMLElement>("source-status");
  status.textContent = text;
  status.classList.remove("text-muted");
  element<HTMLElement>("library-panel").hidden = false;
  element<HTMLElement>("doc-panel").hidden = false;
  renderEntryList();
}

async function loadFromSource(next: BibSource): Promise<void> {
  const text = await next.read();
  const candidate = new CitationEngine(engine.styleId);
  const report = candidate.load(text);
  await candidate.ensureReady();
  // The load is safe whatever the document says; the refresh renders in the pane's default, so it waits on the user's answer.
  const refreshReport = styleChoicePending
    ? null
    : await refreshDocument(candidate);

  engine = candidate;
  selected.clear();
  element<HTMLInputElement>("filter").value = "";
  // Another tool may append to the .bib while the pane is open, so the load time makes that staleness visible.
  const loadedAt = new Date().toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
  });
  showLibrary(`${next.label} — ${engine.size} entries, loaded ${loadedAt}`);
  updateActionControls();

  if (refreshReport === null) {
    message(
      `Loaded ${engine.size} entries from ${next.label}. No citation in this document was ` +
        "rewritten, because which style it uses could not be read — choose the citation style " +
        "above and this document will be refreshed into it.",
      "warn",
    );
  } else {
    await persistPortableSnapshot();
    reportRefresh(
      refreshReport,
      `Loaded ${engine.size} entries from ${next.label} and refreshed`,
    );
  }
  reportFidelity(report.fidelity);
  if (report.warnings.length > 0) {
    message(
      `The parser reported ${report.warnings.length} warning(s): ${report.warnings.slice(0, 3).join("; ")}`,
      "warn",
    );
  }
  message(
    "This is a point-in-time file snapshot. Reopen the .bib to pick up later edits, " +
      "including entries added by another tool.",
    "warn",
  );
}

async function onOpenBib(): Promise<void> {
  await withBusy("open-bib", "Could not load the file", async () => {
    const picked = await FilePickerBibSource.pick();
    if (picked) await loadFromSource(picked);
  });
}

async function onInsertCitation(): Promise<void> {
  const keys = Array.from(selected);
  // The details controls are disabled for several references, so whatever they still show is not for this citation.
  const details = keys.length === 1 ? chosenDetails() : {};
  await withBusy("insert-citation", "Insert failed", async () => {
    // `keys` only ever holds keys engine.keys() rendered a checkbox for, so insertCitation's `missing` can never be non-empty here.
    const { refreshError } = await insertCitation(engine, keys, details);
    // A page belongs to the citation just inserted; carried over, it would land unasked on the next one.
    element<HTMLInputElement>("locator").value = "";
    element<HTMLSelectElement>("citation-form").value = "";
    await persistPortableSnapshot();
    if (refreshError) {
      message(
        `Citation inserted, but document-wide renumbering did not finish (${refreshError}) ` +
          "Use “Refresh all citations” before relying on numbering.",
        "warn",
      );
    } else {
      message(`Inserted citation for ${keys.join(", ")}.`, "ok");
    }
  });
}

async function onInsertBibliography(): Promise<void> {
  await withBusy("insert-bibliography", "Bibliography failed", async () => {
    const missing = await insertBibliography(engine);
    await persistPortableSnapshot();
    if (missing.length > 0)
      message(
        `Bibliography updated; unresolved: ${missing.join(", ")}.`,
        "warn",
      );
    else message("Bibliography updated.", "ok");
  });
}

async function onRefreshDocument(): Promise<void> {
  await withBusy("refresh-doc", "Refresh failed", async () => {
    const report = await refreshDocument(engine);
    await persistPortableSnapshot();
    reportRefresh(report, "Refreshed");
  });
}

async function onStyleChange(): Promise<void> {
  const select = element<HTMLSelectElement>("style-select");
  const previousStyle = engine.styleId;
  const requestedStyle = select.value;

  // withBusy would drop the request silently, leaving the menu showing a style the engine never adopted.
  if (busy) {
    renderStyleOptions();
    return;
  }

  await withBusy("style-select", "Could not apply the style", async () => {
    const name = select.selectedOptions[0]?.textContent ?? requestedStyle;
    const wasPending = styleChoicePending;
    try {
      engine.setStyle(requestedStyle);
      await engine.ensureReady();
      // The user has now said which style this document is written in.
      styleChoicePending = false;
      renderStyleOptions();
      renderCitationDetails();
      if (engine.size > 0) {
        const report = await refreshDocument(engine);
        await persistPortableSnapshot();
        reportRefresh(report, `Applied ${name} to`);
      } else {
        message(`Style set to ${name}.`, "ok");
      }
    } catch (error) {
      // A style that could not be applied is not a style the user chose, so the hold goes back on.
      styleChoicePending = wasPending;
      engine.setStyle(previousStyle);
      renderStyleOptions();
      renderCitationDetails();
      throw error;
    }
  });
}

function clearSelection(): void {
  selected.clear();
  renderEntryList();
  updateActionControls();
}

function showUnsupportedHost(): void {
  element<HTMLElement>("source-status").textContent =
    `${PRODUCT_NAME} runs inside Microsoft Word.`;
}

// `Office.onReady()` never settles when Office.js is blocked, offline or refused by a host's CSP.
const OFFICE_READY_TIMEOUT_MS = 20_000;

async function readyHostInfo(): Promise<{ host: Office.HostType | null }> {
  if (typeof Office === "undefined" || typeof Office.onReady !== "function") {
    throw new Error(
      "the Office.js runtime did not load. Check the network connection to " +
        "appsforoffice.microsoft.com and reload the pane.",
    );
  }
  let timer = 0;
  try {
    return await Promise.race([
      Office.onReady(),
      new Promise<never>((_, reject) => {
        timer = window.setTimeout(
          () =>
            reject(
              new Error(
                "Word did not finish connecting to the add-in. Reload the pane and try again.",
              ),
            ),
          OFFICE_READY_TIMEOUT_MS,
        );
      }),
    ]);
  } finally {
    window.clearTimeout(timer);
  }
}

async function initialize(): Promise<void> {
  const info = await readyHostInfo();
  if (info.host !== Office.HostType.Word) {
    showUnsupportedHost();
    return;
  }

  // Runs before anything reads the embedded records, so a deleted citation's metadata reaches neither the pane nor a later save.
  await removeRecordsForDeletedCitations();
  const snapshotResult = readSnapshot();
  const claim = documentStyleClaim(snapshotResult);
  if (claim.kind === "claimed") {
    // A style-mismatch snapshot's records are held back, but its styleId is the document's own data and must be adopted or the next .bib load refreshes every citation into the pane's default.
    engine = new CitationEngine(claim.styleId);
    if (snapshotResult.status === "valid") {
      engine.loadSnapshot(snapshotResult.snapshot.items);
    }
  }
  // No snapshot is not always a first use: citations outlive a failed snapshot save, or arrive pasted from another document, with nothing recording their style.
  styleChoicePending =
    claim.kind === "unreadable" ||
    (claim.kind === "unclaimed" && hasCitationTag(await documentControlTags()));
  renderStyleOptions();
  renderCitationDetails();
  element<HTMLSelectElement>("locator-label").replaceChildren(
    ...LOCATOR_LABELS.map(
      (label) =>
        new Option(label.charAt(0).toUpperCase() + label.slice(1), label),
    ),
  );
  element<HTMLInputElement>("locator").maxLength = MAX_LOCATOR_LENGTH;

  element<HTMLButtonElement>("open-bib").disabled = false;
  element<HTMLSelectElement>("style-select").disabled = false;
  element<HTMLElement>("source-status").textContent = "No file loaded.";

  onAsync("open-bib", "click", onOpenBib);
  onAsync("insert-citation", "click", onInsertCitation);
  element<HTMLButtonElement>("clear-selection").addEventListener(
    "click",
    clearSelection,
  );
  onAsync("insert-bibliography", "click", onInsertBibliography);
  onAsync("refresh-doc", "click", onRefreshDocument);
  onAsync("style-select", "change", onStyleChange);
  element<HTMLInputElement>("filter").addEventListener(
    "input",
    scheduleEntryRender,
  );
  // Enter in the page box inserts, as it would submit a form; click() does nothing while the button is disabled.
  element<HTMLInputElement>("locator").addEventListener("keydown", (event) => {
    if (event.key === "Enter")
      element<HTMLButtonElement>("insert-citation").click();
  });

  if (snapshotResult.status === "unsupported") {
    message(
      `${snapshotResult.error} Until then ${PRODUCT_NAME} will not modify the embedded data, so nothing ` +
        "this document carries will be lost.",
      "err",
    );
  } else if (snapshotResult.status === "style-mismatch") {
    message(
      `${snapshotResult.error} Its embedded records are preserved when the snapshot is next saved.`,
      "err",
    );
  } else if (snapshotResult.status === "invalid") {
    message(
      `${snapshotResult.error} Open the authoritative .bib file before refreshing.`,
      "err",
    );
  } else if (snapshotResult.status === "valid" && engine.size > 0) {
    showLibrary(`Embedded snapshot — ${engine.size} cited entries`);
    message(
      "Loaded cited metadata embedded in this document. Open the .bib to sync with the current source file.",
      "warn",
    );
  }
  if (styleChoicePending) {
    message(
      `${PRODUCT_NAME} cannot tell which citation style this document is written in, so it will not ` +
        "rewrite any of its citations. They are unchanged, and their keys are still there. Choose " +
        "the citation style above to work on this document.",
      "warn",
    );
  }
  updateActionControls();
}

// Guarded so tests/taskpane.test.ts can import this file's pure exports without running the pane itself; taskpane.html is the only real caller and always has a document.
if (typeof document !== "undefined") {
  initialize().catch((error) => {
    // Last-resort handler: the expected elements may not exist, so it must never throw.
    try {
      element<HTMLElement>("source-status").textContent =
        "Word integration did not start.";
      message(
        `${PRODUCT_NAME} could not initialize: ${describeError(error)}`,
        "err",
      );
    } catch {
      document.title = `${PRODUCT_NAME} — not connected`;
    }
  });
}
