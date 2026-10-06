/* Shared verbatim with the vendored copy in word-bibtex-citation-addin (its scripts/sync-bib-core.mjs re-copies this file, and every file in ../csl/ into its public/assets/csl/) — edit here, never the copy. No imports and ES2019 syntax, since the add-in bundles it to that floor. The citation styles the reference manager copies references in and the Word add-in cites in, and the files in ../csl/ they render with: the Citation Style Language project's own, unmodified (CC BY-SA 3.0), each pinned by its SHA-256. Adding a style is adding its file there and listing it in both CITATION_STYLES and STYLE_DETAILS. */

/** The CSL locales in csl/, by the name citeproc asks for, each with its file's SHA-256. citeproc reads en-US's terms beneath any other locale's, so en-US is always loaded. */
export const CSL_LOCALES = {
  "en-US": {
    file: "locales-en-US.xml",
    sha256: "ac864c7c21166b4390d82c31792cdc509400727fa0060b43d8aa17e07f9cb079",
  },
  "en-GB": {
    file: "locales-en-GB.xml",
    sha256: "e4a1fed9cec04b5fed5e9027b4055925cd4d8f7c9fe233eb0ceca6d4be3999f6",
  },
};

/** The styles, in the order a style menu lists them, each in csl/<id>.csl. An `id` is the upstream file's name, which a Word document records as its style, so it never changes. `locale` is the style's own default where it names one (Nature), else that of the guide it follows: Cite Them Right is British, the rest American. */
export const CITATION_STYLES = /** @type {const} */ ([
  {
    id: "american-medical-association",
    label: "AMA (numeric)",
    locale: "en-US",
  },
  { id: "apa", label: "APA", locale: "en-US" },
  {
    id: "chicago-author-date",
    label: "Chicago (author-date)",
    locale: "en-US",
  },
  {
    id: "harvard-cite-them-right",
    label: "Harvard (Cite Them Right)",
    locale: "en-GB",
  },
  { id: "ieee", label: "IEEE (numeric)", locale: "en-US" },
  { id: "modern-language-association", label: "MLA", locale: "en-US" },
  { id: "nature", label: "Nature (numeric)", locale: "en-GB" },
  {
    // The NLM's Citing Medicine, which is the Vancouver style.
    id: "nlm-citation-sequence",
    label: "Vancouver (numeric)",
    locale: "en-US",
  },
]);

/**
 * @typedef {object} StyleDetails
 * @property {string} sha256  of the stylesheet, which the add-in checks each fetch of it against
 * @property {"author-date"|"author"|"numeric"} format  how the style cites in the text, which decides what a citation's form can do there: "author-date" cites by name and year, "author" by name and page (MLA), and "numeric" by number alone, so it has no other form
 * @property {boolean} superscript  the style sets its citations in superscript, which their plain text cannot carry, so Word is told to
 * @property {boolean} pinpoints  false where the style prints no page in a citation (Nature, Vancouver), so the add-in offers none
 * @property {string[]} inTextOnly  CSL types the style's own rules may cite in the text only, leaving them out of its reference list, as its bibliography layout says
 * @property {string} [narrativeAnd]  the macro that names the authors in the style's citations, where it joins two with "&" but its guide wants "and" in a narrative citation: APA's (Smith & Jones, 2020) reads Smith and Jones (2020) in a sentence, which a stylesheet with one macro for both can't say, so the add-in renders narrative citations with a copy of that macro saying "and"
 */

/** What the Word add-in, which fetches each stylesheet and inserts citations in the text, needs of each style beyond CITATION_STYLES, by id. Kept apart from that list so an app that only lists references, as the reference manager does, loads none of it with its style menu. @type {Record<string, StyleDetails>} */
export const STYLE_DETAILS = {
  "american-medical-association": {
    sha256: "d70f843d93c2b1e8c59c0e26e2396bea4ed99fb4a1aac69b6dddc249a10fadf7",
    format: "numeric",
    superscript: true,
    pinpoints: true,
    inTextOnly: [],
  },
  apa: {
    sha256: "1ece4fb3c295e66d04b4394e295aa58a87741ceeef1658192437eb9953c2f13e",
    format: "author-date",
    superscript: false,
    pinpoints: true,
    // A personal communication (a @letter) or interview a reader cannot retrieve (APA 8.9).
    inTextOnly: ["interview", "personal_communication"],
    narrativeAnd: "author-short",
  },
  "chicago-author-date": {
    sha256: "002fade78d7e4fe9d42936a16b43a8066b097013f6255df40b1bfba6631eff9b",
    format: "author-date",
    superscript: false,
    pinpoints: true,
    // As APA, plus a classic cited with no edition named, and an unsigned dictionary or encyclopaedia entry with no link.
    inTextOnly: [
      "classic",
      "entry-dictionary",
      "entry-encyclopedia",
      "interview",
      "personal_communication",
    ],
  },
  "harvard-cite-them-right": {
    sha256: "0ca3d2c7c881cd98cdd81098b5349f5921edf430498aa9a5a41918621cf7cc45",
    format: "author-date",
    superscript: false,
    pinpoints: true,
    inTextOnly: [],
  },
  ieee: {
    sha256: "b4c7619fc16c45a31e4cc3271eab94ffe83192d3b4c7fc729470a3b459448de3",
    format: "numeric",
    superscript: false,
    pinpoints: true,
    inTextOnly: [],
  },
  "modern-language-association": {
    sha256: "e1432698ef28938d044b06bd45ddcc57d6b974cec25f71007075646a11f99484",
    format: "author",
    superscript: false,
    pinpoints: true,
    inTextOnly: [],
  },
  nature: {
    sha256: "15b1272a3c360168e0b51ab9257534b08ced67190ed57be7a1c32cea8fc87be8",
    format: "numeric",
    superscript: true,
    pinpoints: false,
    inTextOnly: [],
  },
  "nlm-citation-sequence": {
    sha256: "492f7935df42054b33d515f7275122f419780df990ea35925dcbbff0fecf1137",
    format: "numeric",
    superscript: false,
    pinpoints: false,
    inTextOnly: [],
  },
};
