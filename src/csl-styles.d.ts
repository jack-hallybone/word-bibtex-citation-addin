/* Types for csl-styles.js, vendored from bib-core: a style added there is added to CitationStyleId here. */

/** The CSL locales in public/assets/csl/, by the name citeproc asks for. */
export type CslLocaleId = "en-US" | "en-GB";

/** Each locale's file and its SHA-256. citeproc reads en-US's terms beneath any other locale's, so en-US is always loaded. */
export declare const CSL_LOCALES: {
  readonly [locale in CslLocaleId]: {
    readonly file: string;
    readonly sha256: string;
  };
};

/** The upstream file's name, which a document records as its style, so it never changes. */
export type CitationStyleId =
  | "american-medical-association"
  | "apa"
  | "chicago-author-date"
  | "harvard-cite-them-right"
  | "ieee"
  | "modern-language-association"
  | "nature"
  | "nlm-citation-sequence";

export interface CitationStyle {
  readonly id: CitationStyleId;
  readonly label: string;
  /** The style's own default locale where it names one (Nature), else that of the guide it follows: Cite Them Right is British, the rest American. */
  readonly locale: CslLocaleId;
}

/** The styles, in the order the style menu lists them. */
export declare const CITATION_STYLES: readonly CitationStyle[];

/** How a style cites in the text, which decides what a citation's form can do there: "author-date" cites by name and year, "author" by name and page (MLA), and "numeric" by number alone, so it has no other form. */
export type CitationFormat = "author-date" | "author" | "numeric";

export interface StyleDetails {
  /** Of the stylesheet, which each fetch of it is checked against. */
  readonly sha256: string;
  readonly format: CitationFormat;
  /** The style sets its citations in superscript, which their plain text cannot carry, so Word is told to. */
  readonly superscript: boolean;
  /** False where the style prints no page in a citation (Nature, Vancouver), so the pane offers none. */
  readonly pinpoints: boolean;
  /** CSL types the style's own rules may cite in the text only, leaving them out of its reference list, as its bibliography layout says. */
  readonly inTextOnly: readonly string[];
  /** The macro that names the authors in the style's citations, where it joins two with "&" but its guide wants "and" in a narrative citation, as APA's does: narrative citations are rendered with a copy of that macro saying "and". */
  readonly narrativeAnd?: string;
}

/** What the add-in needs of each style beyond CITATION_STYLES, by id. */
export declare const STYLE_DETAILS: {
  readonly [id in CitationStyleId]: StyleDetails;
};
