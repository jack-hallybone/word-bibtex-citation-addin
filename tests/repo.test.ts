import assert from "node:assert/strict";
import { readdir, readFile } from "node:fs/promises";
import { test } from "node:test";

import { apiFloor, browserFloor } from "../scripts/browser-floor.mjs";

/** What ships broken in files the bundle never contains — the compiler's own floors, the manifest template, and the workflow holding the Pages credential — none of which live in `dist/`, so no other suite can reach them. */

/** Report every violation, not the first — the same reason as in `tests/artefact.test.ts`. */
function noneOf(problems: readonly string[]): void {
  assert.equal(problems.length, 0, `\n  ${problems.join("\n  ")}`);
}

/** A `lib` above `target` type-checks a method the oldest supported webview does not have and ships it unpolyfilled, uncaught by tsc (which is what's fooled) or by the acorn syntax pass over the chunk, since a missing method is not a syntax error. */
test("the API floor does not run ahead of the syntax floor", () => {
  const syntax = browserFloor();
  const api = apiFloor();
  assert.ok(
    api.ecmaVersion <= syntax.ecmaVersion,
    `tsconfig.json: lib opens ${api.lib} while target is ${syntax.target.toUpperCase()}: downlevelling rewrites ` +
      `syntax and polyfills no API, so an ES${api.ecmaVersion} method would type-check clean and ship to a ` +
      "webview without it",
  );
});

/** `OfficeApp`'s children are an `xsd:sequence` — out of order the host silently rejects the manifest until someone sideloads it — checked here since `npm run validate` needs Microsoft's online validator and can't run in CI; top-level children are read as lines indented by exactly two spaces. */
test("the manifest's top-level elements are in the schema's sequence", async () => {
  const problems: string[] = [];
  const file = "manifest.xml";
  // The task-pane order from the add-in-only manifest schema, in full, so an element added later is placed against the schema rather than by eye.
  const schemaOrder = [
    "Id",
    "AlternateId",
    "Version",
    "ProviderName",
    "DefaultLocale",
    "DisplayName",
    "Description",
    "IconUrl",
    "HighResolutionIconUrl",
    "SupportUrl",
    "AppDomains",
    "Hosts",
    "Requirements",
    "DefaultSettings",
    "Permissions",
    "Dictionary",
    "VersionOverrides",
    "ExtendedOverrides",
  ];
  const manifest = await readFile(file, "utf8");
  const found = [...manifest.matchAll(/^ {2}<([A-Za-z]+)[\s>/]/gm)].map(
    (match) => match[1] ?? "",
  );

  assert.ok(
    found.length >= 10,
    `${file}: found only ${found.length} top-level elements; the two-space indentation this test reads must have changed`,
  );

  let previous = -1;
  for (const name of found) {
    const position = schemaOrder.indexOf(name);
    if (position === -1) {
      problems.push(
        `${file}: <${name}> is not an OfficeApp child in this schema`,
      );
      continue;
    }
    if (position < previous) {
      problems.push(
        `${file}: <${name}> comes after <${schemaOrder[previous]}>, but the schema orders it before`,
      );
    }
    previous = Math.max(previous, position);
  }
  noneOf(problems);
});

const WORKFLOWS = ".github/workflows";

/** Every workflow file, refusing an empty answer so nothing here passes vacuously. */
async function workflows(): Promise<{ path: string; yaml: string }[]> {
  const names = (await readdir(WORKFLOWS)).filter((name) =>
    /\.ya?ml$/.test(name),
  );
  assert.notEqual(names.length, 0, `${WORKFLOWS}: no workflows found to check`);
  return Promise.all(
    names.map(async (name) => ({
      path: `${WORKFLOWS}/${name}`,
      yaml: await readFile(`${WORKFLOWS}/${name}`, "utf8"),
    })),
  );
}

/** A workflow's step blocks: a `- ` list item and everything under it. */
function stepBlocks(yaml: string): string[] {
  const blocks: string[] = [];
  let current: string[] | null = null;
  for (const line of yaml.split("\n")) {
    if (/^\s*-\s/.test(line)) {
      if (current) blocks.push(current.join("\n"));
      current = [line];
    } else if (current) {
      current.push(line);
    }
  }
  if (current) blocks.push(current.join("\n"));
  return blocks;
}

/** A step's `name:`, for a failure message that points at the right step. */
function stepLabel(block: string): string {
  return /\bname:\s*(.+)/.exec(block)?.[1]?.trim() ?? "unnamed step";
}

/** The body of each job in `jobs:`, keyed by job name. */
function jobBodies(yaml: string): Map<string, string> {
  const lines = yaml.split("\n");
  const start = lines.findIndex((line) => /^jobs:[ \t]*$/.test(line));
  const jobs = new Map<string, string>();
  if (start === -1) return jobs;
  let name: string | null = null;
  let body: string[] = [];
  for (const line of lines.slice(start + 1)) {
    if (/^\S/.test(line)) break; // a top-level key ends the mapping
    const head = /^ {2}([A-Za-z0-9_-]+):[ \t]*$/.exec(line);
    if (head) {
      if (name !== null) jobs.set(name, body.join("\n"));
      name = head[1] ?? "";
      body = [];
    } else if (name !== null) {
      body.push(line);
    }
  }
  if (name !== null) jobs.set(name, body.join("\n"));
  return jobs;
}

/** The lines nested under `<indent spaces><key>:` in a block, to its dedent. */
function nestedUnder(body: string, key: string, indent: number): string[] {
  const lines = body.split("\n");
  const header = new RegExp(`^ {${indent}}${key}:[ \\t]*$`);
  const start = lines.findIndex((line) => header.test(line));
  if (start === -1) return [];
  const nested: string[] = [];
  for (const line of lines.slice(start + 1)) {
    if (line.trim() === "") continue;
    if (line.length - line.trimStart().length <= indent) break;
    nested.push(line);
  }
  return nested;
}

/** The `deploy` job holds the credential that publishes the site, so it must never check the repository out or run a script since a pull request could change what that step executes; privilege is derived from each job's declared permissions, not its name, so a second privileged job cannot appear unnoticed. */
test("no job that holds a write credential runs this repository's code", async () => {
  const problems: string[] = [];
  let privileged = 0;

  for (const { path, yaml } of await workflows()) {
    const jobs = jobBodies(yaml);
    assert.notEqual(jobs.size, 0, `${path}: declares no jobs`);

    for (const [name, body] of jobs) {
      const writes = nestedUnder(body, "permissions", 4)
        .map((line) => /^\s*([a-z-]+):\s*(\S+)\s*$/.exec(line))
        .filter((entry) => entry?.[2] === "write")
        .map((entry) => entry?.[1] ?? "");
      if (writes.length === 0) continue;
      privileged += 1;

      const steps = stepBlocks(nestedUnder(body, "steps", 4).join("\n"));
      assert.notEqual(
        steps.length,
        0,
        `${path}: job ${name} holds ${writes.join(", ")} but declares no steps, so this test read the file wrongly`,
      );
      for (const step of steps) {
        const label = stepLabel(step);
        if (/^\s*-?\s*run:/m.test(step)) {
          problems.push(
            `${path}: job ${name} holds ${writes.join(", ")} and its step "${label}" runs a script; a change to that script reaches the credential`,
          );
        }
        if (/uses:\s*actions\/checkout@/.test(step)) {
          problems.push(
            `${path}: job ${name} holds ${writes.join(", ")} and its step "${label}" checks the repository out; the credential must not meet repository code`,
          );
        }
      }
    }
  }

  assert.notEqual(
    privileged,
    0,
    "no job in any workflow declares a write permission, so this test proved nothing — the deploy job must have been renamed or removed",
  );
  noneOf(problems);
});

/** A checkout that keeps its credential leaves the job's `GITHUB_TOKEN` in `.git/config`, where anything the job then runs (a build script, a transitive install hook) can read it and push to `main`, so this checks per step rather than by name, since a second checkout step elsewhere must not be vouched for by one hardened step. */
test("no checkout keeps the job's credential in the work tree", async () => {
  const problems: string[] = [];
  let checkouts = 0;

  for (const { path, yaml } of await workflows()) {
    const steps = stepBlocks(yaml).filter((block) =>
      /uses:\s*actions\/checkout@/.test(block),
    );
    checkouts += steps.length;
    for (const [index, block] of steps.entries()) {
      if (!/persist-credentials:\s*false/.test(block)) {
        const label = stepLabel(block);
        problems.push(
          `${path}: checkout step ${index + 1} of ${steps.length} (${label}) must set persist-credentials: false`,
        );
      }
    }
  }

  assert.notEqual(
    checkouts,
    0,
    "no workflow checks the repository out, so this test proved nothing",
  );
  noneOf(problems);
});
