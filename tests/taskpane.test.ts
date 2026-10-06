import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";

import { CitationEngine } from "../src/lib/citationEngine";
import { BIBLIOGRAPHY_TAG, CslItem, citationTag } from "../src/lib/types";
import {
  actionControlState,
  entryLabel,
  filterEntryKeys,
  formOptions,
  hasCitationTag,
  keptRecordsNote,
  refreshWarnings,
  skippedBibliographyNote,
} from "../src/taskpane/taskpane";

/** A minimal, otherwise-valid item, so each case overrides only what it tests. */
function item(fields: Partial<CslItem> = {}): CslItem {
  return { id: "k", ...fields };
}

test("entryLabel joins author, year and title with the pane's own separator", () => {
  assert.equal(
    entryLabel(
      item({
        author: [{ family: "Doe" }],
        issued: { "date-parts": [[2020]] },
        title: "A Title",
      }),
    ),
    "Doe · 2020 · A Title",
  );
});

test("entryLabel uses bib-core's short author form: 'A', 'A and B', then 'A et al.'", () => {
  const many = item({
    author: [{ family: "Doe" }, { family: "Roe" }, { family: "Poe" }],
    title: "T",
  });
  assert.equal(entryLabel(many), "Doe et al. · T");

  const two = item({
    author: [{ family: "Doe" }, { family: "Roe" }],
    title: "T",
  });
  assert.equal(entryLabel(two), "Doe and Roe · T");

  const one = item({ author: [{ family: "Doe" }], title: "T" });
  assert.equal(entryLabel(one), "Doe · T");
});

test("entryLabel reads BibTeX's preserved 'and others' marker as et al.", () => {
  const engine = new CitationEngine();
  engine.load(
    "@article{o, author={Doe, Jane and others}, title={T}, year={2000}}",
  );
  const loaded = engine.item("o");
  assert.ok(loaded);
  assert.equal(entryLabel(loaded), "Doe et al. · 2000 · T");
});

test("entryLabel falls back through family, then literal, then given", () => {
  assert.equal(
    entryLabel(item({ author: [{ literal: "Doe Co." }], title: "T" })),
    "Doe Co. · T",
  );
  assert.equal(
    entryLabel(item({ author: [{ given: "Jane" }], title: "T" })),
    "Jane · T",
  );
});

test("entryLabel omits whatever the item does not carry, without stray separators", () => {
  assert.equal(entryLabel(item({ title: "Only A Title" })), "Only A Title");
  assert.equal(
    entryLabel(item({ author: [{ family: "Solo" }] })),
    "Solo",
    "an author with neither year nor title must not trail a separator",
  );
  assert.equal(
    entryLabel(item({ issued: { "date-parts": [[1999]] } })),
    "1999",
  );
  assert.equal(
    entryLabel(item({})),
    "",
    "an item with no author, year or title renders as empty, not as bare separators",
  );
});

test("entryLabel strips citeproc's own rich-text markup rather than showing it", () => {
  assert.equal(
    entryLabel(
      item({ title: 'A Study of <span class="nocase">TeX</span> Systems' }),
    ),
    "A Study of TeX Systems",
  );
});

test("entryLabel clips an assembled label that runs past the pane's display budget", () => {
  const longTitle = "x".repeat(400);
  const label = entryLabel(item({ title: longTitle }));
  assert.ok(
    label.length <= 280,
    `label ran to ${label.length} characters, past the 280 budget`,
  );
  assert.ok(label.length > 0);
});

test("filterEntryKeys returns every key, in the engine's own order, for a blank or whitespace-only query", () => {
  const engine = new CitationEngine();
  engine.load(
    "@article{doe2020, author={Doe, Jane}, title={Alpha}, year={2020}}\n" +
      "@article{roe2021, author={Roe, John}, title={Beta}, year={2021}}",
  );
  assert.deepEqual(filterEntryKeys(engine, ""), ["doe2020", "roe2021"]);
  assert.deepEqual(
    filterEntryKeys(engine, "   "),
    ["doe2020", "roe2021"],
    "a query that is only whitespace must behave like no filter at all",
  );
});

test("filterEntryKeys matches case-insensitively against the key and the rendered label", () => {
  const engine = new CitationEngine();
  engine.load(
    "@article{doe2020, author={Doe, Jane}, title={Widget Robustness}, year={2020}}\n" +
      "@article{roe2021, author={Roe, John}, title={Something Else}, year={2021}}",
  );
  assert.deepEqual(
    filterEntryKeys(engine, "DOE2020"),
    ["doe2020"],
    "must match on the key, case-insensitively",
  );
  assert.deepEqual(
    filterEntryKeys(engine, "widget"),
    ["doe2020"],
    "must match on the rendered label's title",
  );
  assert.deepEqual(
    filterEntryKeys(engine, "2021"),
    ["roe2021"],
    "must match on the rendered label's year",
  );
  assert.deepEqual(filterEntryKeys(engine, "nothing matches this"), []);
});

test("actionControlState: with a selection and nothing pending, every action is available", () => {
  assert.deepEqual(actionControlState(1, false, false), {
    insertCitationDisabled: false,
    clearSelectionDisabled: false,
    detailsDisabled: false,
    otherActionsDisabled: false,
  });
});

test("actionControlState: an empty selection blocks inserting and clearing, but not document-wide actions", () => {
  assert.deepEqual(actionControlState(0, false, false), {
    insertCitationDisabled: true,
    clearSelectionDisabled: true,
    detailsDisabled: false,
    otherActionsDisabled: false,
  });
});

test("actionControlState: an unresolved style choice blocks every write except clearing the selection", () => {
  // Clearing the selection writes nothing, so it alone must survive a pending style choice.
  assert.deepEqual(actionControlState(1, false, true), {
    insertCitationDisabled: true,
    clearSelectionDisabled: false,
    detailsDisabled: false,
    otherActionsDisabled: true,
  });
});

test("actionControlState: busy overrides selection and style state for every control", () => {
  assert.deepEqual(actionControlState(1, true, false), {
    insertCitationDisabled: true,
    clearSelectionDisabled: true,
    detailsDisabled: true,
    otherActionsDisabled: true,
  });
});

test("actionControlState: a page and form wait for one reference alone, since a citation of several takes neither", () => {
  assert.equal(actionControlState(1, false, false).detailsDisabled, false);
  assert.equal(
    actionControlState(0, false, false).detailsDisabled,
    false,
    "a page may be typed before its reference is ticked",
  );
  assert.equal(actionControlState(2, false, false).detailsDisabled, true);
  assert.equal(
    actionControlState(2, false, false).insertCitationDisabled,
    false,
    "several references still insert, as a plain citation",
  );
});

test("formOptions offers a narrative and a year-only form in an author-date style, a page-only one in MLA, and none in a numeric one", () => {
  assert.deepEqual(formOptions("author-date"), [
    ["", "Standard: (Smith, 2020)"],
    ["narrative", "Narrative: Smith (2020)"],
    ["suppress-author", "Year only: (2020)"],
  ]);
  // MLA cites no year, so its forms are the ones it renders (see core.test.ts).
  assert.deepEqual(formOptions("author"), [
    ["", "Standard: (Smith 12)"],
    ["narrative", "Narrative: Smith (12)"],
    ["suppress-author", "Page only: (12)"],
  ]);
  assert.deepEqual(formOptions("numeric"), []);
});

test("hasCitationTag: a document with no snapshot holds the style choice only when it already cites", () => {
  assert.equal(hasCitationTag([]), false, "a new document is not held");
  assert.equal(hasCitationTag([citationTag(["doe2020"])]), true);
  assert.equal(
    hasCitationTag(["other-add-in", citationTag(["doe2020", "roe2021"])]),
    true,
  );
  assert.equal(
    hasCitationTag([BIBLIOGRAPHY_TAG]),
    false,
    "a bibliography alone is no citation",
  );
  assert.equal(
    hasCitationTag(["bibcite-pending-00112233445566778899aabbccddeeff"]),
    false,
    "nor is a failed insertion's leftover",
  );
  // An earlier build's pinpoint tag is refused by the tag parser, but it is still a citation written in some style.
  assert.equal(hasCitationTag(['bibcite:2;{"k":"doe2020","l":"5"}']), true);
});

test("skippedBibliographyNote says nothing resolved only when nothing did, and names the style that cites the rest in the text only", () => {
  const skipped = {
    citationsUpdated: 2,
    bibliographyUpdated: false,
    bibliographySkipped: true,
  };
  const style = "Chicago (author-date)";
  assert.match(
    skippedBibliographyNote(
      { ...skipped, keysResolved: 0, missingKeys: ["a", "b"] },
      style,
    ),
    /no citation resolved/,
  );
  assert.match(
    skippedBibliographyNote(
      { ...skipped, keysResolved: 2, missingKeys: [] },
      style,
    ),
    /every citation is to a work Chicago \(author-date\) cites in the text only, such as a personal communication\.$/,
  );
  const mixed = skippedBibliographyNote(
    { ...skipped, keysResolved: 1, missingKeys: ["b"] },
    style,
  );
  assert.match(
    mixed,
    /every citation that resolved is to a work Chicago \(author-date\) cites in the text only/,
  );
  assert.doesNotMatch(mixed, /no citation resolved|APA/);
});

test("refreshWarnings says only that nothing was rewritten when a document has no citations", () => {
  const none = refreshWarnings(
    {
      citationsUpdated: 0,
      bibliographyUpdated: false,
      bibliographySkipped: true,
      missingKeys: [],
      keysResolved: 0,
    },
    "APA",
  );
  assert.equal(none.length, 1);
  assert.match(
    none[0] ?? "",
    /citations were found in this document, so nothing was rewritten/,
  );

  // With citations, a skipped bibliography gets its note, then the unresolved keys.
  const unresolved = {
    citationsUpdated: 2,
    bibliographyUpdated: false,
    bibliographySkipped: true,
    missingKeys: ["a", "b"],
    keysResolved: 0,
  };
  assert.deepEqual(refreshWarnings(unresolved, "APA"), [
    skippedBibliographyNote(unresolved, "APA"),
    "Unresolved keys: a, b.",
  ]);
});

test("keptRecordsNote names every kept key and both ways to restore their citations", () => {
  const note = keptRecordsNote(["smith2020", "doe2019"]);
  assert.match(note, /does not define: smith2020, doe2019\./);
  assert.match(note, /open a \.bib that defines them/);
  assert.match(
    note,
    /reopen this pane and choose “Refresh all citations” before opening a \.bib/,
  );
});

test("the pane says a refresh replaces text typed inside a citation, and where a page goes instead", async () => {
  // Nothing locks a citation or the bibliography, so a page number typed inside one is lost at the next refresh; this note is the pane's only warning.
  const pane = (await readFile("src/taskpane/taskpane.html", "utf8")).replace(
    /\s+/g,
    " ",
  );
  assert.ok(pane.includes("replacing anything typed inside them."));
  assert.ok(
    pane.includes(
      "page and form when you insert it; to change them, delete the citation and insert it again. " +
        "Type anything else outside the citation.",
    ),
  );
  assert.ok(pane.includes('<label class="field-label" for="locator"'));
});
