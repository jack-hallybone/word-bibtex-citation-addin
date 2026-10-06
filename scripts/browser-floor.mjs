/** The one reader of the browser floor tsconfig.json's `target` declares, so vite.config.ts (builds to it) and tests/artefact.test.ts (parses the shipped chunk at it) cannot drift apart. */
import ts from "typescript";

// Resolved against this file: vite.config.ts and the suite don't agree about the working directory, and the suite is bundled elsewhere before it runs.
const TSCONFIG = new URL("../tsconfig.json", import.meta.url).pathname;

/** Read by tsc's own parser, so a comment or trailing comma reads as tsc reads it. */
function compilerOptions() {
  const { config, error } = ts.readConfigFile(TSCONFIG, ts.sys.readFile);
  if (error) {
    throw new Error(
      `tsconfig.json does not parse (${ts.flattenDiagnosticMessageText(error.messageText, " ")}), so the browser floor cannot be read from it`,
    );
  }
  const options = config?.compilerOptions;
  if (options === null || typeof options !== "object") {
    throw new Error(
      "tsconfig.json declares no `compilerOptions`, so nothing says what the oldest supported Word webview can parse",
    );
  }
  return options;
}

/** `ES<year>` as a number, or `undefined` for any other spelling. */
const esYear = (value, allowSublibrary) =>
  typeof value === "string"
    ? Number(
        (allowSublibrary ? /^ES(\d{4})(?:\.[\w.]+)?$/i : /^ES(\d{4})$/i).exec(
          value,
        )?.[1],
      ) || undefined
    : undefined;

/** The syntax floor, in the esbuild and acorn spellings; a moving target (ESNext) is refused rather than defaulted, since accepting one would reduce the artefact check to "is this JavaScript". */
export function browserFloor() {
  const { target } = compilerOptions();
  if (target === undefined) {
    throw new Error(
      "tsconfig.json declares no `target`, so nothing says what the oldest supported Word webview can parse",
    );
  }
  const year = esYear(target, false);
  if (year === undefined) {
    throw new Error(
      `tsconfig.json's target is ${JSON.stringify(target)}; the browser floor must be an ES<year> ` +
        "so the build can downlevel to it and the shipped chunk can be parsed at it",
    );
  }
  return { target: target.toLowerCase(), ecmaVersion: year };
}

/** The API floor: the newest ES surface `lib` opens (a lib list is a union), since downlevelling rewrites syntax like `?.` but doesn't add APIs like `Promise.allSettled`. */
export function apiFloor() {
  const { lib } = compilerOptions();
  if (!Array.isArray(lib)) {
    throw new Error(
      "tsconfig.json declares no `lib` array, so nothing says which APIs the pane may reach for",
    );
  }
  let floor;
  for (const entry of lib) {
    if (typeof entry !== "string" || !/^es/i.test(entry)) continue;
    const year = esYear(entry, true);
    if (year === undefined) {
      throw new Error(
        `tsconfig.json's lib names ${JSON.stringify(entry)}; an ES entry must be an ES<year>, optionally with a ` +
          "sub-library suffix, so it can be compared with the target",
      );
    }
    if (floor === undefined || year > floor.ecmaVersion) {
      floor = { lib: entry, ecmaVersion: year };
    }
  }
  if (floor === undefined) {
    throw new Error(
      `tsconfig.json's lib is ${JSON.stringify(lib)}, which names no ES library, so the API floor cannot be compared with the target`,
    );
  }
  return floor;
}
