// Turns citeproc's reference-list html into the Office Open XML that Word inserts, so each entry's layout is set exactly: Word's html import ignores citeproc's classes, putting an IEEE number on a line of its own and dropping APA's hanging indent.

/** How a style lays out its reference list, from citeproc's bibliography metadata. */
export interface BibliographyLayout {
  /** Every line of an entry but its first is indented (APA, Chicago, MLA). */
  hangingIndent: boolean;
  /** Characters in the widest first field a style sets apart in a column, such as "[10]" (IEEE) or "10." (Vancouver), with the rest of the entry aligned beside it; 0 for none. */
  numberColumn: number;
}

/** Half an inch, the hanging indent APA, Chicago and MLA prescribe, in Word's twentieths of a point. */
const HANGING_INDENT_TWIPS = 720;
/** A digit's width at 12 pt, so the number column fits its widest number with a character to spare. */
const COLUMN_CHARACTER_TWIPS = 120;

interface RunFormat {
  italic: boolean;
  bold: boolean;
  smallCaps: boolean;
  underline: boolean;
  vertAlign: "superscript" | "subscript" | null;
}

interface FormatTag {
  tag: string;
  format: Partial<RunFormat>;
}

/** Each formatting tag citeproc's html output format can emit (citationEngine.ts refuses any other before this runs), and what it does to the text inside it. */
const FORMAT_TAGS: Readonly<Record<string, FormatTag>> = {
  "<i>": { tag: "i", format: { italic: true } },
  "<em>": { tag: "em", format: { italic: true } },
  "<b>": { tag: "b", format: { bold: true } },
  "<sup>": { tag: "sup", format: { vertAlign: "superscript" } },
  "<sub>": { tag: "sub", format: { vertAlign: "subscript" } },
  '<span style="font-style:normal;">': {
    tag: "span",
    format: { italic: false },
  },
  '<span style="font-variant:small-caps;">': {
    tag: "span",
    format: { smallCaps: true },
  },
  '<span style="font-variant:normal;">': {
    tag: "span",
    format: { smallCaps: false },
  },
  '<span style="font-weight:normal;">': {
    tag: "span",
    format: { bold: false },
  },
  '<span style="text-decoration:none;">': {
    tag: "span",
    format: { underline: false },
  },
  '<span style="text-decoration:underline;">': {
    tag: "span",
    format: { underline: true },
  },
  '<span style="baseline">': { tag: "span", format: { vertAlign: null } },
};

const DISPLAY_DIV =
  /^<div class="csl-(entry|left-margin|right-inline|block|indent)">$/;

/** citeproc escapes exactly these three in text, and nothing else, so decoding them alone, in one pass, gives back the text. */
const ENTITY = /&#(38|60|62);/g;
const ENTITY_TEXT: Readonly<Record<string, string>> = {
  "38": "&",
  "60": "<",
  "62": ">",
};

/** XML 1.0 cannot carry an unpaired surrogate (a field cut to length can split a pair) or U+FFFE/U+FFFF, and Word refuses a whole insert over one. */
function xmlText(text: string): string {
  return text
    .replace(
      /[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDFFF\uFFFE\uFFFF]/g,
      (unit) => (unit.length === 2 ? unit : "\uFFFD"),
    )
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function runProperties(format: RunFormat): string {
  // In the order the schema requires.
  const properties = [
    format.bold ? "<w:b/>" : "",
    format.italic ? "<w:i/>" : "",
    format.smallCaps ? "<w:smallCaps/>" : "",
    format.underline ? '<w:u w:val="single"/>' : "",
    format.vertAlign ? `<w:vertAlign w:val="${format.vertAlign}"/>` : "",
  ].join("");
  return properties ? `<w:rPr>${properties}</w:rPr>` : "";
}

/** One entry as one paragraph; throws on anything outside citeproc's own structure, which citationEngine.ts reports as a refused render. */
function entryParagraph(html: string, layout: BibliographyLayout): string {
  const plain: RunFormat = {
    italic: false,
    bold: false,
    smallCaps: false,
    underline: false,
    vertAlign: null,
  };
  const formats: RunFormat[] = [plain];
  const open: (FormatTag | { tag: "div"; display: string })[] = [];
  const runs: string[] = [];
  let entries = 0;
  let lineBreakDue = false;

  // A line break and the spaces around it, next to a div, are how citeproc indents its own html, not text.
  const tokens = html
    .replace(/\s*\n\s*(?=<\/?div)/g, "")
    .replace(/(<\/?div[^>]*>)\s*\n\s*/g, "$1")
    .split(/(<[^>]*>)/);
  for (const token of tokens) {
    if (!token) continue;
    const current = formats[formats.length - 1] ?? plain;
    if (!token.startsWith("<")) {
      if (/[<>]/.test(token)) throw new Error("unescaped markup in text");
      if (open.length === 0) throw new Error("text outside the entry");
      if (lineBreakDue) runs.push("<w:r><w:br/></w:r>");
      lineBreakDue = false;
      const text = token
        .replace(/[\t\r\n]+/g, " ")
        .replace(ENTITY, (_, code: string) => ENTITY_TEXT[code] ?? "");
      runs.push(
        `<w:r>${runProperties(current)}<w:t xml:space="preserve">${xmlText(text)}</w:t></w:r>`,
      );
      continue;
    }

    const display = DISPLAY_DIV.exec(token)?.[1];
    const formatTag = FORMAT_TAGS[token];
    if (display !== undefined) {
      if (
        display === "entry"
          ? open.length > 0 || ++entries > 1
          : open.length === 0
      ) {
        throw new Error("content outside its one entry");
      }
      if (display === "right-inline") runs.push("<w:r><w:tab/></w:r>");
      // A block starts a line of its own, as html displays it.
      if ((display === "block" || display === "indent") && runs.length > 0)
        lineBreakDue = true;
      open.push({ tag: "div", display });
    } else if (formatTag !== undefined) {
      if (open.length === 0) throw new Error("formatting outside the entry");
      open.push(formatTag);
      formats.push({ ...current, ...formatTag.format });
    } else {
      const closing = /^<\/(\w+)>$/.exec(token)?.[1];
      const opener = open.pop();
      if (closing === undefined || opener?.tag !== closing) {
        throw new Error(`unexpected markup ${JSON.stringify(token)}`);
      }
      if ("format" in opener) formats.pop();
      else if (opener.display === "block" || opener.display === "indent")
        lineBreakDue = true;
    }
  }
  if (open.length > 0) throw new Error("an entry left unclosed");

  const indent = entryIndent(layout);
  // The tab after a number stops at the indent, where the entry's later lines start.
  const tabs =
    layout.numberColumn > 0
      ? `<w:tabs><w:tab w:val="left" w:pos="${indent}"/></w:tabs>`
      : "";
  const properties =
    indent > 0
      ? `<w:pPr>${tabs}<w:ind w:left="${indent}" w:hanging="${indent}"/></w:pPr>`
      : "";
  return `<w:p>${properties}${runs.join("")}</w:p>`;
}

/** How far every line of an entry but its first is indented, in Word's twentieths of a point; the first line hangs back by as much. */
export function entryIndent(layout: BibliographyLayout): number {
  if (layout.numberColumn > 0)
    return (layout.numberColumn + 1) * COLUMN_CHARACTER_TWIPS;
  return layout.hangingIndent ? HANGING_INDENT_TWIPS : 0;
}

/** The smallest package Word's insertOoxml takes: its relationships and the document part. No styles part, so every paragraph and run takes the document's own font and spacing. */
export function bibliographyOoxml(
  entries: readonly string[],
  layout: BibliographyLayout,
): string {
  const paragraphs = entries
    .map((entry) => entryParagraph(entry, layout))
    .join("");
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
    '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships></pkg:xmlData></pkg:part>" +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
    '<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${paragraphs}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`
  );
}
