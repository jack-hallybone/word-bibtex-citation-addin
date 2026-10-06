// Run after deliberately deciding to pull a bib-core change: copies every VENDORED_FILES entry byte-for-byte. Never hand-edit a vendored copy; change bib-core and re-run this instead.

import { copyFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { BIB_CORE_DIR, VENDORED_FILES } from "./bib-core-vendor.mjs";

const ROOT = fileURLToPath(new URL("..", import.meta.url));

for (const { local, upstream } of VENDORED_FILES) {
  const from = join(BIB_CORE_DIR, upstream);
  const to = join(ROOT, local);
  mkdirSync(dirname(to), { recursive: true });
  copyFileSync(from, to);
}
console.log(`Synced ${VENDORED_FILES.length} files from bib-core.`);
