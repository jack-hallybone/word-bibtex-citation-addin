/* Types for names.js, vendored from bib-core. */

/** A raw (still TeX-encoded) BibTeX name list split on "and" at brace depth 0. */
export declare function splitNames(value: string | null | undefined): string[];

/** A raw name's family part, still TeX-encoded. */
export declare function familyName(name: string | null | undefined): string;

/** One "A", two "A and B", three or more (or a trailing "others") "A et al.". */
export declare function shortAuthors(
  families: ReadonlyArray<string | null | undefined>,
): string;
