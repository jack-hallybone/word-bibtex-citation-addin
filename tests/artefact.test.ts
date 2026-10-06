import assert from "node:assert/strict";
import { readdir, readFile, stat } from "node:fs/promises";
import { test } from "node:test";

import * as acorn from "acorn";

import { browserFloor } from "../scripts/browser-floor.mjs";
import { manifestEndpointFailures } from "../scripts/manifest-endpoints.mjs";

/** `dist/` is an input to this file, never an output: the deploy workflow runs this suite on the same folder it then uploads. */
const DIST = "dist";

/** The upstream styles as committed, which is where their authorship metadata is authoritative. */
const CSL_SOURCE = "public/assets/csl";

const EMAIL_PATTERN = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;

try {
  if (!(await stat(DIST)).isDirectory()) throw new Error("not a directory");
} catch {
  throw new Error(
    `No ${DIST}/ to check. Every assertion in this file is about the served ` +
      "folder: run `npm run build` first (it needs VITE_SITE_URL, which " +
      "`npm run verify` supplies).",
  );
}

/** Report every violation, not the first: these read whole folders. */
function noneOf(problems: readonly string[]): void {
  assert.equal(problems.length, 0, `\n  ${problems.join("\n  ")}`);
}

async function readIfPresent(path: string): Promise<Buffer | null> {
  try {
    return await readFile(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw error;
  }
}

/** Every JavaScript chunk in the served folder, refusing an empty answer. */
async function bundleChunks(): Promise<
  { path: string; code: string }[] | null
> {
  const dir = `${DIST}/assets`;
  let names: string[];
  try {
    names = (await readdir(dir)).filter((name) => name.endsWith(".js"));
  } catch {
    return null;
  }
  if (names.length === 0) return null;
  return Promise.all(
    names.map(async (name) => ({
      path: `${dir}/${name}`,
      code: await readFile(`${dir}/${name}`, "utf8"),
    })),
  );
}

async function chunksOrFail(): Promise<{ path: string; code: string }[]> {
  const chunks = await bundleChunks();
  assert.notEqual(
    chunks,
    null,
    `${DIST}/assets emitted no JavaScript to check`,
  );
  return chunks ?? [];
}

/** `no-console` lints this project's source only, never the dependencies bundled into it, so this reads the shipped chunk directly for a call that would hand the user's bibliography to a host log the pane's CSP cannot reach. */
test("no console call survives into the shipped chunk", async () => {
  const problems: string[] = [];
  // A member call, not the bare identifier: citeproc keeps a `typeof console` guard around a fallback it never reaches in a browser.
  const call = /console\s*\.\s*[A-Za-z$_]/g;
  for (const { path, code } of await chunksOrFail()) {
    const hits = code.match(call);
    if (hits) {
      problems.push(
        `${path}: ${hits.length} console call(s) survive minification (${[...new Set(hits)].join(", ")}); bibliography content must not reach the host console`,
      );
    }
  }
  noneOf(problems);
});

/** The pane's origin is shared with every other project on `<user>.github.io`, so browser storage is a leak into them; checked against the shipped chunk rather than a list of source directories, since a route added anywhere a grep doesn't look would pass unnoticed. */
const ESCAPE_ROUTES: ReadonlyArray<readonly [string, string]> = [
  [
    "localStorage",
    "browser storage shared with every other site on this origin",
  ],
  [
    "sessionStorage",
    "browser storage shared with every other site on this origin",
  ],
  ["indexedDB", "browser storage shared with every other site on this origin"],
  ["caches", "the Cache API"],
  ["serviceWorker", "a service worker, which outlives the pane"],
  ["createObjectURL", "a URL the user can open or save"],
  ["msSaveBlob", "a file save"],
  ["showSaveFilePicker", "a file save"],
  ["clipboard", "the clipboard"],
  ["\\.download\\b", "a download link"],
];

test("the shipped chunk has no route out but the document's own settings bag", async () => {
  const problems: string[] = [];
  const chunks = await chunksOrFail();
  for (const { path, code } of chunks) {
    for (const [needle, what] of ESCAPE_ROUTES) {
      if (new RegExp(needle).test(code)) {
        problems.push(
          `${path}: reaches ${what} (${needle}); the .docx's own settings bag must stay the only thing this add-in writes or hands out`,
        );
      }
    }
  }
  // And the one route that IS allowed has to still be there, or this passes on a build that embeds nothing at all.
  if (!chunks.some(({ code }) => /settings\.set\b/.test(code))) {
    problems.push(
      `${DIST}/assets: no chunk writes the document's settings bag, so either the embedded snapshot has gone or this test is looking at the wrong bytes`,
    );
  }
  noneOf(problems);
});

/** Only the bundler can hold the output to the floor tsconfig.json declares, checked by parsing rather than scanning so `?.` inside a string is not mistaken for optional chaining. */
test("the shipped chunk parses at the browser floor tsconfig.json declares", async () => {
  const problems: string[] = [];
  const floor = browserFloor();

  const parses = (code: string, ecmaVersion: number): Error | null => {
    try {
      acorn.parse(code, {
        ecmaVersion: ecmaVersion as acorn.ecmaVersion,
        sourceType: "module",
      });
      return null;
    } catch (error) {
      return error as Error;
    }
  };

  for (const { path, code } of await chunksOrFail()) {
    const error = parses(code, floor.ecmaVersion);
    if (error === null) continue;
    // The lowest level it DOES parse at, so the report says how far past the floor the build went.
    let needs = "a level newer than this test tries";
    for (let year = floor.ecmaVersion + 1; year <= 2026; year += 1) {
      if (parses(code, year) === null) {
        needs = `ES${year}`;
        break;
      }
    }
    problems.push(
      `${path}: does not parse as ${floor.target.toUpperCase()}, the floor tsconfig.json declares: ${error.message}. ` +
        `It needs ${needs}. The bundle is what the webview parses, so \`build.target\` in vite.config.ts is what has to carry the floor.`,
    );
  }
  noneOf(problems);
});

/** The CSP that keeps the bibliography and the document off the wire is a property of the page the host loads, not the template it was built from, since Vite rewrites that file; whitespace is collapsed because the directives wrap across lines. */
test("the served pane keeps the CSP that holds the bibliography on this origin", async () => {
  const problems: string[] = [];
  const path = `${DIST}/taskpane.html`;
  const bytes = await readIfPresent(path);
  assert.notEqual(bytes, null, `${path}: the build emitted no task pane`);

  const html = (bytes as Buffer).toString("utf8").replace(/\s+/g, " ");
  assert.match(
    html,
    /<meta http-equiv="Content-Security-Policy"/,
    `${path}: the served pane carries no Content-Security-Policy`,
  );
  // Anchored on the terminating `;`, so "connect-src 'self' https://anywhere" fails rather than matching a prefix.
  for (const [pattern, detail] of [
    [/connect-src 'self' ?;/, "connect-src must stay at exactly 'self'"],
    [/form-action 'none' ?;/, "form-action must be 'none'"],
    [/object-src 'none' ?;/, "object-src must be 'none'"],
    [/style-src 'self' ?;/, "style-src must be 'self'"],
  ] as const) {
    if (!pattern.test(html))
      problems.push(`${path}: ${detail} in the served pane`);
  }
  const widened = /unsafe-inline|unsafe-eval/.exec(html);
  if (widened) problems.push(`${path}: the served CSP allows ${widened[0]}`);
  noneOf(problems);
});

/** The only email addresses the served folder may carry are the CSL styles' own authorship metadata, which CC-BY-SA requires kept intact — a bound on what LEAVES the build, so a new address anywhere (template, comment, dependency banner) fails rather than being published. */
test("the served folder carries no email address but the CSL metadata", async () => {
  const problems: string[] = [];
  const pinnedStyles = /^dist\/assets\/csl\/[\w.-]+\.(csl|xml)$/;

  // Read off the upstream files rather than listed, so no third party's address is written into this repository's own source.
  const allowed = new Set<string>();
  for (const name of await readdir(CSL_SOURCE)) {
    const text = await readFile(`${CSL_SOURCE}/${name}`, "utf8");
    for (const [address] of text.matchAll(EMAIL_PATTERN)) allowed.add(address);
  }
  assert.notEqual(
    allowed.size,
    0,
    `${CSL_SOURCE} yielded no address, so this test no longer knows what upstream carries`,
  );

  async function* walk(dir: string): AsyncGenerator<string> {
    for (const entry of await readdir(dir, { withFileTypes: true })) {
      const path = `${dir}/${entry.name}`;
      if (entry.isDirectory()) yield* walk(path);
      else yield path;
    }
  }

  for await (const path of walk(DIST)) {
    if (/\.(png|ico|woff2?)$/.test(path)) continue;
    const text = await readFile(path, "utf8");
    for (const [address] of text.matchAll(EMAIL_PATTERN)) {
      if (!allowed.has(address)) {
        problems.push(
          `${path}: ships an unexpected email address (${address})`,
        );
      } else if (!pinnedStyles.test(path)) {
        problems.push(
          `${path}: ships ${address} outside the pinned CSL styles, where it is not covered by their attribution`,
        );
      }
    }
  }
  noneOf(problems);
});

/** Comments are stripped before scanning since the theme's prose names its own custom properties, and an unstripped reader would match that English text as CSS. */
/** Every stylesheet the build emitted, joined, with CSS comments stripped so a `;` inside one cannot be read as the end of a declaration. */
async function servedStyles(): Promise<{ names: string[]; css: string }> {
  const names = (await readdir(`${DIST}/assets`)).filter((name) =>
    name.endsWith(".css"),
  );
  const css = (
    await Promise.all(
      names.map((name) => readFile(`${DIST}/assets/${name}`, "utf8")),
    )
  )
    .join("\n")
    .replace(/\/\*[\s\S]*?\*\//g, " ");
  return { names, css };
}

test("the served pane's stylesheet is linked, and declares every token it uses", async () => {
  const problems: string[] = [];
  const { names: sheets, css } = await servedStyles();
  const documents = (await readdir(DIST)).filter((name) =>
    name.endsWith(".html"),
  );
  assert.notEqual(
    sheets.length,
    0,
    `${DIST}/assets: the build emitted no stylesheet`,
  );
  assert.notEqual(
    documents.length,
    0,
    `${DIST}: the build emitted no document to link a stylesheet from`,
  );

  const markup = (
    await Promise.all(
      documents.map((name) => readFile(`${DIST}/${name}`, "utf8")),
    )
  ).join("\n");
  for (const name of sheets) {
    if (!markup.includes(name)) {
      problems.push(
        `${DIST}/assets/${name}: is emitted but no served document references it, so nothing loads it`,
      );
    }
  }

  const declared = new Set(
    [...css.matchAll(/(--[\w-]+)\s*:/g)].map((match) => match[1]),
  );
  assert.notEqual(
    declared.size,
    0,
    `${DIST}/assets: the served stylesheet declares no custom property; the theme did not reach it`,
  );
  // A trailing `,` after the name is a fallback (e.g. `var(--x, 1px)`), which needs no declaration.
  for (const [, name, fallback] of css.matchAll(/var\(\s*(--[\w-]+)\s*(,)?/g)) {
    if (!fallback && !declared.has(name)) {
      problems.push(
        `${DIST}/assets: a served rule reaches for ${name} and no served rule declares it`,
      );
    }
  }
  noneOf(problems);
});

test("the minifier leaves the theme's colour syntax alone", async () => {
  // Lightning CSS rewrites `light-dark()` for older targets into a `prefers-color-scheme` pair that answers the OS; `build.cssTarget` in vite.config.ts stops that.
  const { names, css } = await servedStyles();
  assert.notEqual(
    names.length,
    0,
    `${DIST}/assets: the build emitted no stylesheet`,
  );
  assert.match(
    css,
    /light-dark\(/u,
    `${DIST}/assets: no served rule uses light-dark(), so the syntax the vendored theme declares its colours in did not survive the build`,
  );
  assert.doesNotMatch(
    css,
    /--lightningcss-/u,
    `${DIST}/assets: the minifier polyfilled light-dark() into --lightningcss-* custom properties, which answer the operating system rather than the element's own color-scheme; build.cssTarget in vite.config.ts is what stops it`,
  );
});

/** `taskpane.html` and `manifest.xml` are templates carrying a `__PRODUCT_NAME__` token, so a dropped substitution ships a pane titled with the token and a manifest Word shows the same way in its ribbon. */
test("the served pane and manifest carry the display name, not the token", async () => {
  const problems: string[] = [];
  const token = "__PRODUCT_NAME__";
  const { productName } = JSON.parse(
    await readFile("package.json", "utf8"),
  ) as { productName?: unknown };
  assert.equal(
    typeof productName === "string" && productName !== "",
    true,
    "package.json declares no productName to have derived from",
  );

  for (const path of [`${DIST}/taskpane.html`, `${DIST}/manifest.xml`]) {
    const bytes = await readIfPresent(path);
    if (bytes === null) {
      problems.push(`${path}: the build did not emit it`);
      continue;
    }
    const text = bytes.toString("utf8");
    if (text.includes(token)) {
      problems.push(
        `${path}: still carries an unsubstituted ${token}; the build did not fill the display name in`,
      );
    }
    if (!text.includes(productName as string)) {
      problems.push(
        `${path}: does not name ${JSON.stringify(productName)}, the declared productName; the derivation dropped it`,
      );
    }
  }
  noneOf(problems);
});

/** Sideloaded, a manifest whose SourceLocation, icons and bt:Urls point at `https://localhost:3000` grants ReadWriteDocument over the reader's Word document to whatever answers on the reader's own machine, so this checks the shipped bytes directly rather than trusting the build's own check, which is meant to stop it. */
test("the shipped manifest points only at public https endpoints", async () => {
  const path = `${DIST}/manifest.xml`;
  const bytes = await readIfPresent(path);
  assert.notEqual(bytes, null, `${path}: the build did not emit it`);
  noneOf(
    manifestEndpointFailures((bytes as Buffer).toString("utf8")).map(
      (refusal) => `${path}: ${refusal}`,
    ),
  );
});
