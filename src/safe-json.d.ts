/* Types for safe-json.js: every bound is required with no default, or an omitted cap silently reads as unbounded (++keys > undefined is always false). */

export declare class SafeInputError extends Error {
  constructor(message: string);
  name: "SafeInputError";
}

export interface ParseJsonOptions {
  /** Byte cap, checked before the parse. */
  maxBytes: number;
  /** Deepest `{`/`[` nesting allowed, measured on the text before the parse. */
  maxDepth: number;
  /** Cap on the values the reviver visits. */
  maxKeys: number;
  /** What the person is looking at (e.g. "This document's citation data"): opens every refusal message. */
  what?: string;
}

/** Parses untrusted JSON; every object comes back on a null prototype (arrays stay arrays), and anything other than an object/array — including null — is refused. */
export declare function parseJson(
  text: string,
  options: ParseJsonOptions,
): object;
