// The file list sync-bib-core.mjs walks: the contract fixture, the safety module, the short author form, and the citation styles with the engine that reads BibTeX for them, since everything else in src/lib is this add-in's own reimplementation.

import { readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";

export const BIB_CORE_DIR = fileURLToPath(
  new URL("../../reference-manager/bib-core", import.meta.url),
);

/** local: this project's vendored copy. upstream: the bib-core file it must stay byte-identical to. Every file in bib-core's csl/ is a stylesheet or locale this add-in serves from public/assets/csl/, pinned in csl-styles. */
export const VENDORED_FILES = [
  { local: "src/safe-json.js", upstream: "src/safe-json.mjs" },
  { local: "src/names.js", upstream: "src/names.mjs" },
  { local: "src/csl-styles.js", upstream: "src/csl-styles.mjs" },
  { local: "src/csl-engine.js", upstream: "src/csl-engine.mjs" },
  {
    local: "tests/citekey-conformance.json",
    upstream: "citekey-conformance.json",
  },
  ...readdirSync(`${BIB_CORE_DIR}/csl`)
    .sort()
    .map((name) => ({
      local: `public/assets/csl/${name}`,
      upstream: `csl/${name}`,
    })),
];
