import { mkdtempSync, readFileSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { spawnSync } from "node:child_process";
import * as esbuild from "esbuild";

/** `scripts/*.mjs` stay external, not bundled: `browser-floor.mjs` resolves `tsconfig.json` relative to its own file, and a bundled copy under /tmp would have no correct anchor for that. */
const realScripts = {
  name: "real-scripts",
  setup(build) {
    build.onResolve({ filter: /^\.\.\/scripts\/[\w.-]+\.mjs$/ }, (args) => ({
      path: pathToFileURL(resolve(args.resolveDir, args.path)).href,
      external: true,
    }));
  },
};

// Discovered, not listed: a suite file added later must not be able to go unrun.
const suites = readdirSync("tests")
  .filter((name) => name.endsWith(".test.ts"))
  .sort()
  .map((name) => resolve("tests", name));
if (suites.length === 0) throw new Error("no tests/*.test.ts suites found");

const directory = mkdtempSync(join(tmpdir(), "bibcite-tests-"));
try {
  // The display name reaches the code through one build-time substitution, so the suite makes the same one the bundler does (`define`, off the same field) — nothing gets to hard-code a fallback.
  const { productName } = JSON.parse(readFileSync("package.json", "utf8"));
  if (typeof productName !== "string" || productName.trim() === "") {
    throw new Error("package.json declares no productName to build against");
  }
  await esbuild.build({
    entryPoints: suites,
    bundle: true,
    outdir: directory,
    outExtension: { ".js": ".mjs" },
    platform: "node",
    format: "esm",
    target: "node24",
    define: { __PRODUCT_NAME__: JSON.stringify(productName) },
    plugins: [realScripts],
    logLevel: "warning",
  });
  const result = spawnSync(
    process.execPath,
    [
      "--test",
      ...suites.map((suite) =>
        join(directory, `${basename(suite, ".ts")}.mjs`),
      ),
    ],
    { cwd: process.cwd(), stdio: "inherit" },
  );
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} finally {
  rmSync(directory, { recursive: true, force: true });
}
