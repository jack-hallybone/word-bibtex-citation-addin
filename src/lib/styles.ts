// The styles the pane cites in: bib-core's (src/csl-styles.js, vendored from reference-manager/bib-core), which the reference manager copies references in too, with the same labels and locales, each with what this add-in needs of it.

import {
  CITATION_STYLES,
  CSL_LOCALES,
  STYLE_DETAILS,
  type CitationFormat,
  type CitationStyle,
  type StyleDetails,
} from "../csl-styles.js";

export { CSL_LOCALES, type CitationFormat };

export type LocaleId = keyof typeof CSL_LOCALES;

/** CSL styles bundled with the add-in, in the order the style menu lists them. */
export const BUILTIN_STYLES: ReadonlyArray<CitationStyle & StyleDetails> =
  CITATION_STYLES.map((style) => ({ ...style, ...STYLE_DETAILS[style.id] }));

export type StyleId = (typeof BUILTIN_STYLES)[number]["id"];
export type BuiltinStyle = (typeof BUILTIN_STYLES)[number];

export function isStyleId(value: unknown): value is StyleId {
  return (
    typeof value === "string" &&
    BUILTIN_STYLES.some((style) => style.id === value)
  );
}

export function builtinStyle(styleId: StyleId): BuiltinStyle {
  const style = BUILTIN_STYLES.find((candidate) => candidate.id === styleId);
  if (style === undefined)
    throw new Error(`No citation style is registered as ${styleId}.`);
  return style;
}

export function styleRevision(styleId: StyleId): string {
  const style = builtinStyle(styleId);
  return `${style.sha256}:${CSL_LOCALES[style.locale].sha256}`;
}
