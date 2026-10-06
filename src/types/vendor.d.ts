// Minimal ambient declarations for citeproc (which ships no types) covering exactly our call sites; bibliojson ships real types, so it isn't declared here.

declare module "citeproc" {
  export interface CiteprocSys {
    retrieveLocale(lang: string): string;
    /** MUST NOT return undefined — the engine only ever asks for ids we hold. */
    retrieveItem(id: string): unknown;
  }
  /** Declared to exactly the properties this add-in passes; `label` is one of CSL's locator terms ("page", "chapter", …). */
  export interface CiteItem {
    id: string;
    locator?: string;
    label?: string;
  }
  /** A `mode` names the author, so it must never reach a numeric style, which has no author to emit: IEEE renders the literal text "[NO_PRINTED_FORM]". */
  export type CitationMode = "composite" | "author-only" | "suppress-author";
  export interface Citation {
    citationID?: string;
    citationItems: CiteItem[];
    properties: { noteIndex: number; mode?: CitationMode };
  }
  export class Engine {
    constructor(sys: CiteprocSys, style: string, lang?: string);
    setOutputFormat(format: "text" | "html" | "rtf"): void;
    makeBibliography():
      | [
          {
            entry_ids: string[][];
            bibliography_errors: unknown[];
            [key: string]: unknown;
          },
          string[],
        ]
      | false;
    /** With no citations, returns the processor to the empty state it was constructed in. */
    restoreProcessorState(citations: []): unknown;
    /** Returns [citationID, noteIndex, text] positionally, in document order. */
    rebuildProcessorState(
      citations: Citation[],
      mode: "text" | "html" | "rtf",
      uncitedItemIDs?: string[],
    ): Array<[string, number, string]>;
  }
  const CSL: { Engine: typeof Engine; [key: string]: unknown };
  export default CSL;
}
