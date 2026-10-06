import assert from "node:assert/strict";
import { test } from "node:test";

import { bibliographyOoxml } from "../src/lib/ooxml";

const NONE = { hangingIndent: false, numberColumn: 0 };

/** The runs of the first paragraph in a package, each as its properties and its text. */
function runs(xml: string): [string, string][] {
  const paragraph = /<w:p>.*?<\/w:p>/.exec(xml)?.[0] ?? "";
  return [
    ...paragraph.matchAll(
      /<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?(?:<w:t xml:space="preserve">([^<]*)<\/w:t>|(<w:tab\/>|<w:br\/>))<\/w:r>/g,
    ),
  ].map(([, properties, text, mark]) => [properties ?? "", text ?? mark ?? ""]);
}

test("each entry becomes one paragraph, laid out as the style's metadata asks", () => {
  const entry = '<div class="csl-entry">Doe, J. (2020). Title.</div>';
  const paragraphs = (xml: string) => xml.match(/<w:p>.*?<\/w:p>/g) ?? [];

  const hanging = paragraphs(
    bibliographyOoxml([entry, entry], { hangingIndent: true, numberColumn: 0 }),
  );
  assert.equal(hanging.length, 2);
  assert.match(hanging[0] ?? "", /<w:ind w:left="720" w:hanging="720"\/>/);

  // Harvard's metadata asks for neither, so its entries take the document's own indentation.
  assert.doesNotMatch(
    paragraphs(bibliographyOoxml([entry], NONE))[0] ?? "",
    /<w:pPr>/,
  );

  // The column fits the widest number citeproc reports ("[10]" is 4 characters), whatever the hanging indent says.
  const numbered = bibliographyOoxml(
    [
      '<div class="csl-entry">\n    <div class="csl-left-margin">[10]</div><div class="csl-right-inline">Doe, J.</div>\n  </div>',
    ],
    { hangingIndent: true, numberColumn: 4 },
  );
  assert.match(
    numbered,
    /<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="600"\/><\/w:tabs><w:ind w:left="600" w:hanging="600"\/><\/w:pPr>/,
  );
  assert.deepEqual(runs(numbered), [
    ["", "[10]"],
    ["", "<w:tab/>"],
    ["", "Doe, J."],
  ]);
});

test("citeproc's formatting becomes Word's, nested as citeproc nests it", () => {
  const xml = bibliographyOoxml(
    [
      '<div class="csl-entry"><i>Italic <span style="font-style:normal;">upright</span> again</i> ' +
        '<b>12</b><span style="font-weight:normal;"> plain</span> <sup>2</sup><sub>x</sub> ' +
        '<span style="font-variant:small-caps;">Caps<span style="font-variant:normal;"> low</span></span> ' +
        '<span style="text-decoration:underline;">under<span style="text-decoration:none;"> not</span></span> ' +
        '<sup>up <span style="baseline">base</span></sup> <em>oblique</em></div>',
    ],
    NONE,
  );
  assert.deepEqual(runs(xml), [
    ["<w:i/>", "Italic "],
    ["", "upright"],
    ["<w:i/>", " again"],
    ["", " "],
    ["<w:b/>", "12"],
    ["", " plain"],
    ["", " "],
    ['<w:vertAlign w:val="superscript"/>', "2"],
    ['<w:vertAlign w:val="subscript"/>', "x"],
    ["", " "],
    ["<w:smallCaps/>", "Caps"],
    ["", " low"],
    ["", " "],
    ['<w:u w:val="single"/>', "under"],
    ["", " not"],
    ["", " "],
    ['<w:vertAlign w:val="superscript"/>', "up "],
    ["", "base"],
    ["", " "],
    ["<w:i/>", "oblique"],
  ]);
});

test("text decodes citeproc's escapes once and is escaped again for XML", () => {
  // citeproc writes a title's literal "&#60;" as "&#38;#60;", which must come back as those five characters, not as "<".
  const xml = bibliographyOoxml(
    ['<div class="csl-entry">A &#60;b&#62; &#38; c &#38;#60; d\u00a0 e</div>'],
    NONE,
  );
  assert.deepEqual(runs(xml), [
    ["", "A &lt;b&gt; &amp; c &amp;#60; d\u00a0 e"],
  ]);
});

test("a character XML cannot carry becomes a replacement character, not a refused insert", () => {
  // A field cut to its length limit can end in half a surrogate pair.
  const xml = bibliographyOoxml(
    [
      '<div class="csl-entry">Cut \ud83d|\ude00| \uffff|kept \ud83d\ude00</div>',
    ],
    NONE,
  );
  assert.deepEqual(runs(xml), [
    ["", "Cut \ufffd|\ufffd| \ufffd|kept \ud83d\ude00"],
  ]);
});

test("a block starts a line of its own within its entry", () => {
  const xml = bibliographyOoxml(
    [
      '<div class="csl-entry">Before\n\n    <div class="csl-block">Block</div>\nAfter</div>',
    ],
    NONE,
  );
  assert.deepEqual(runs(xml), [
    ["", "Before"],
    ["", "<w:br/>"],
    ["", "Block"],
    ["", "<w:br/>"],
    ["", "After"],
  ]);
});

test("markup outside citeproc's own structure is refused, not passed to Word", () => {
  for (const entry of [
    '<div class="csl-entry"><a href="x">link</a></div>',
    '<div class="csl-entry"><i>unclosed</div>',
    '<div class="csl-entry">closed twice</i></div>',
    '<div class="csl-entry">unclosed',
    'text first<div class="csl-entry">entry</div>',
    '<div class="csl-entry">one</div><div class="csl-entry">two</div>',
    '<div class="csl-left-margin">[1]</div>',
    '<i>outside</i><div class="csl-entry">entry</div>',
    '<div class="csl-entry">a > b</div>',
    '<div class="csl-entry">a < b</div>',
  ]) {
    assert.throws(() => bibliographyOoxml([entry], NONE), Error, entry);
  }
});
