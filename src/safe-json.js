/* Shared verbatim with the vendored copies in literature-pdf-bibtex-downloader and word-bibtex-citation-addin (each project's scripts/sync-bib-core.mjs re-copies this file) — coordinate changes across all three, don't edit any one copy in isolation. Reading JSON this app did not write: the browser's parser, plus the three bounds it does not have and a reviver that drops prototype keys. */

export class SafeInputError extends Error {
  constructor(message) {
    super(message);
    this.name = "SafeInputError";
  }
}

// `constructor` and `prototype` are not vectors the way `__proto__` is, but a record carrying them reaches code expecting a value, and this app stores neither.
const FORBIDDEN_KEYS = new Set(["__proto__", "constructor", "prototype"]);

const withThousands = (count) =>
  String(count).replace(/\B(?=(\d{3})+(?!\d))/g, ",");

/* Only brackets outside a string count, so a title containing "[[[" is not mistaken for nesting. */
function measureDepth(text, limit) {
  let depth = 0;
  let deepest = 0;
  let inString = false;
  let escaped = false;
  for (const c of text) {
    if (inString) {
      if (escaped) escaped = false;
      else if (c === "\\") escaped = true;
      else if (c === '"') inString = false;
      continue;
    }
    if (c === '"') inString = true;
    else if (c === "{" || c === "[") {
      depth += 1;
      if (depth > deepest) deepest = depth;
      if (deepest > limit) return deepest;
    } else if (c === "}" || c === "]") depth -= 1;
  }
  return deepest;
}

export function parseJson(
  text,
  { what = "This file", maxBytes, maxDepth, maxKeys } = {},
) {
  /* No defaults: a caller that omits `maxKeys` would get `++keys > undefined`, which is always false, and be unbounded without knowing it. */
  for (const [name, value] of [
    ["maxBytes", maxBytes],
    ["maxDepth", maxDepth],
    ["maxKeys", maxKeys],
  ]) {
    if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
      throw new TypeError(`parseJson needs a positive ${name}; got ${value}`);
    }
  }
  if (typeof text !== "string") {
    throw new SafeInputError(`${what} could not be read as text.`);
  }
  /* Length before bytes: encoding allocates the whole input a second time, which is what the cap exists to refuse. UTF-8 never spends fewer bytes than a string has UTF-16 code units, so a length over the cap proves the byte count is, and the message says "at least" because that is what is known without encoding. */
  if (text.length > maxBytes) {
    throw new SafeInputError(
      `${what} is at least ${withThousands(text.length)} bytes, over the ` +
        `${withThousands(maxBytes)}-byte limit.`,
    );
  }
  const bytes = new TextEncoder().encode(text).length;
  if (bytes > maxBytes) {
    throw new SafeInputError(
      `${what} is ${withThousands(bytes)} bytes, over the ` +
        `${withThousands(maxBytes)}-byte limit.`,
    );
  }
  if (text.trim() === "") throw new SafeInputError(`${what} is empty.`);

  /* The cap is for serialisation, not the parser: structuredClone throws RangeError from about 5,000 deep, so an uncapped document parses and then cannot be sent anywhere. */
  if (measureDepth(text, maxDepth) > maxDepth) {
    throw new SafeInputError(
      `${what} is nested too deeply (over ${maxDepth} levels).`,
    );
  }

  let keys = 0;
  let parsed;
  try {
    parsed = JSON.parse(text, function reviver(key, value) {
      if (FORBIDDEN_KEYS.has(key)) return undefined;
      if (++keys > maxKeys) {
        throw new SafeInputError(`${what} has more than ${maxKeys} entries.`);
      }
      if (
        value !== null &&
        typeof value === "object" &&
        !Array.isArray(value)
      ) {
        return Object.assign(Object.create(null), value);
      }
      return value;
    });
  } catch (error) {
    if (error instanceof SafeInputError) throw error;
    throw new SafeInputError(`${what} is not valid JSON.`);
  }

  if (parsed === null || typeof parsed !== "object") {
    throw new SafeInputError(`${what} is not a JSON object or array.`);
  }
  return parsed;
}
