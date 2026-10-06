/* Types for csl-engine.js, vendored from bib-core. */

import type { Engine } from "citeproc";
import type { BiblatexParseResult, CSLEntry } from "bibliojson";

/** citeproc's whole HTML vocabulary, as a global pattern: what is left of its output once these are removed holds no "<". */
export declare const CITEPROC_HTML_TAG: RegExp;

export interface ParsedBibtex extends BiblatexParseResult {
  /** The key of the first entry bibliojson stopped reading partway through without an error, if one did. */
  cutShort: string | undefined;
}

/** `bibtex` as bibliojson reads it, with the fixes this add-in needs. Neither failure throws: each must be refused before the entries are used. */
export declare function parseBibtex(bibtex: string): ParsedBibtex;

/** Each entry as a CSL-JSON item with its key as its id, in the entries' order, or undefined for one bibliojson's exporter dropped. */
export declare function cslItems(
  entries: BiblatexParseResult["entries"],
): Array<[string, CSLEntry | undefined]>;

/** A citeproc processor for one stylesheet, citing `items` by key; an item of CSL's old catch-all type "article" reaches it as "document". */
export declare function createProcessor(
  style: string,
  locale: string,
  retrieveLocale: (lang: string) => string,
  items: ReadonlyMap<string, unknown>,
): Engine;
