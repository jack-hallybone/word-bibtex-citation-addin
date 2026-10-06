// Reports where the .bib source and the rendered citation will disagree; the .bib is never modified, so nothing here repairs anything.

import { Library } from "./types";

/** Fields other tools sharing the library folder write; warning about them would bury the warnings that matter. */
const BENIGN_UNKNOWN_FIELDS = new Set([
  "file",
  "groups",
  "owner",
  "timestamp",
  "priority",
  "readstatus",
  "mendeley-tags",
  "bdsk-file-1",
  // Written by every entry the sibling reference-manager project saves, in JabRef's own timestamp form.
  "creationdate",
  "modificationdate",
]);

function isLetter(character: string | undefined): boolean {
  return (
    character !== undefined &&
    ((character >= "a" && character <= "z") ||
      (character >= "A" && character <= "Z"))
  );
}

/** The parser's own whitespace, less the CR the scan normalises away. */
function isBlank(character: string | undefined): boolean {
  return character === " " || character === "\t" || character === "\n";
}

/** Remove `%` line comments between entries, and between an entry's fields: BibTeX has no comment syntax, so a commented-out entry would otherwise parse, while inside a value a `%` is text ("50%"). Malformed input errs toward being inside a value. */
export function stripPercentComments(input: string): string {
  // CR-only input can loop the downstream BibTeX parser forever; normalise line endings before scanning.
  const source = input.replace(/\r\n?/g, "\n");
  if (!source.includes("%")) return source;

  let out = "";
  let pendingOpener = -1;
  let closer: "}" | ")" | null = null;
  let commentBlock = false;
  let outerDepth = 0;
  let braceDepth = 0;
  let quoted = false;
  // Counted forward, not by walking back (which made the scan quadratic on a run of backslashes); an odd count means the character is escaped.
  let precedingBackslashes = 0;

  const findOpener = (index: number): number => {
    let cursor = index + 1;
    const start = cursor;
    while (isLetter(source[cursor])) cursor++;
    if (cursor === start) return -1;
    while (isBlank(source[cursor])) cursor++;
    return source[cursor] === "{" || source[cursor] === "(" ? cursor : -1;
  };

  /** Where the parser reads a `%` as a comment too: after whitespace, outside every value. Straight after text it can belong to a key ("50%off"), so those are left to the parser, as is an @comment block, which it reads as text. */
  const betweenFields = (index: number): boolean =>
    !commentBlock &&
    !quoted &&
    outerDepth === 1 &&
    braceDepth === 0 &&
    isBlank(source[index - 1]);

  for (let index = 0; index < source.length; index++) {
    const character = source[index];

    // Not only at the start of a line: left for the parser, `} % checked by Gau\ss` loses its line break to TeX substitution and the comment swallows the next entry, or between fields the next field.
    if (
      character === "%" &&
      (closer === null ? precedingBackslashes % 2 === 0 : betweenFields(index))
    ) {
      while (index < source.length && source[index] !== "\n") index++;
      if (index < source.length) out += "\n";
      pendingOpener = -1;
      // `index` rests on the newline that ended the comment, so no backslash precedes the next character.
      precedingBackslashes = 0;
      continue;
    }

    if (closer === null) {
      if (character === "@") pendingOpener = findOpener(index);
      if (index === pendingOpener) {
        closer = character === "{" ? "}" : ")";
        commentBlock =
          source
            .slice(source.lastIndexOf("@", index) + 1, index)
            .trim()
            .toLowerCase() === "comment";
        outerDepth = 1;
        braceDepth = 0;
        quoted = false;
        pendingOpener = -1;
      }
    } else if (precedingBackslashes % 2 === 0) {
      if (
        character === '"' &&
        (quoted || (closer === "}" ? outerDepth === 1 : braceDepth === 0))
      ) {
        quoted = !quoted;
      } else if (!quoted && closer === "}") {
        if (character === "{") outerDepth++;
        else if (character === "}" && --outerDepth === 0) closer = null;
      } else if (!quoted && closer === ")") {
        if (character === "{") braceDepth++;
        else if (character === "}" && braceDepth > 0) braceDepth--;
        else if (braceDepth === 0 && character === "(") outerDepth++;
        else if (braceDepth === 0 && character === ")" && --outerDepth === 0)
          closer = null;
      }
    }

    out += character;
    precedingBackslashes = character === "\\" ? precedingBackslashes + 1 : 0;
  }
  return out;
}

/** True for a warning about a field another tool in the folder writes by design. */
export function isBenignFieldWarning(value: unknown): boolean {
  if (!value || typeof value !== "object") return false;
  const diagnostic = value as Record<string, unknown>;
  return (
    typeof diagnostic.type === "string" &&
    diagnostic.type.includes("field") &&
    typeof diagnostic.field_name === "string" &&
    BENIGN_UNKNOWN_FIELDS.has(diagnostic.field_name.toLowerCase())
  );
}

/** Entries using BibTeX's `and others`, which CSL-JSON cannot express: dropping it would make the list look complete and would also change the in-text citation, since citeproc counts names. */
export function describeTruncatedNameLists(library: Library): string[] {
  const notes: string[] = [];
  for (const [key, item] of library) {
    for (const field of ["author", "editor"] as const) {
      const list = item[field];
      if (!Array.isArray(list)) continue;
      if (
        !list.some(
          (name) =>
            (name as Record<string, unknown> | null)?.literal === "others",
        )
      )
        continue;
      notes.push(
        `entry "${key}": BibTeX "and others" has no CSL equivalent, so this ${field} list ` +
          `renders as "… and others" rather than "et al.".`,
      );
    }
  }
  return notes;
}
