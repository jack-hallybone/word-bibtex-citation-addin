/* Shared verbatim with the vendored copy in word-bibtex-citation-addin (its scripts/sync-bib-core.mjs re-copies this file) — edit here, never the copy. No imports and ES2019 syntax, since the add-in bundles it to that floor. BibTeX name lists read the way a citation reads them: "Doe, Jane and Roe, Richard" -> "Doe and Roe", three or more -> "Doe et al.". */

const MAX_NAMES_LEN = 20000;
const AND = /\s+and\s+/iy;

/** A raw (still TeX-encoded) name list split on "and" at brace depth 0, as BibTeX splits it: `{Barnes and Noble}` stays one name. */
export function splitNames(value) {
  if (value == null) return [];
  const text = String(value).slice(0, MAX_NAMES_LEN);
  const names = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "{") depth++;
    else if (c === "}") depth = Math.max(0, depth - 1);
    else if (
      depth === 0 &&
      /\s/.test(c) &&
      (i === 0 || !/\s/.test(text[i - 1]))
    ) {
      // Tried only where a run of blanks starts: a match from inside the run is the same match, and trying at every blank re-read the rest of a long run each time.
      AND.lastIndex = i;
      if (AND.test(text)) {
        names.push(text.slice(start, i));
        start = AND.lastIndex;
        i = start - 1;
      }
    }
  }
  names.push(text.slice(start));
  return names.map((name) => name.replace(/\s+/g, " ").trim()).filter(Boolean);
}

/** One raw name's words and commas at brace depth 0; a braced group is one word. */
function topLevelParts(name) {
  const parts = [];
  let word = "";
  let depth = 0;
  for (const c of name) {
    if (c === "{") depth++;
    else if (c === "}") depth = Math.max(0, depth - 1);
    if (depth === 0 && (c === "," || /\s/.test(c))) {
      if (word) parts.push(word);
      if (c === ",") parts.push(",");
      word = "";
    } else word += c;
  }
  if (word) parts.push(word);
  return parts;
}

/** A raw name's family part, still TeX-encoded: "von Last, First" and "First von Last" both give "von Last"; a lone braced group is returned whole. */
export function familyName(name) {
  const parts = topLevelParts(String(name == null ? "" : name).trim());
  const comma = parts.indexOf(",");
  if (comma !== -1) return parts.slice(0, comma).join(" ");
  if (parts.length <= 1) return parts.length ? parts[0] : "";
  // BibTeX's "von" rule: the family part starts at the first lowercase word after the first, else it is the last word.
  const von = parts.findIndex(
    (word, i) => i > 0 && i < parts.length - 1 && /^\p{Ll}/u.test(word),
  );
  return parts.slice(von === -1 ? parts.length - 1 : von).join(" ");
}

/** The short author form from family names already decoded for display: one "A", two "A and B", three or more (or a trailing "others") "A et al.". */
export function shortAuthors(families) {
  const names = (families || [])
    .map((family) => String(family == null ? "" : family).trim())
    .filter(Boolean);
  const others = names.length > 0 && /^others$/i.test(names[names.length - 1]);
  if (others) names.pop();
  if (names.length === 0) return "";
  if (names.length === 1) return others ? `${names[0]} et al.` : names[0];
  if (names.length === 2 && !others) return `${names[0]} and ${names[1]}`;
  return `${names[0]} et al.`;
}
