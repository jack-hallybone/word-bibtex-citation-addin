import { readFileSync } from "node:fs";
import process from "node:process";

// Licences compatible with this project being all rights reserved.
const ALLOWED = new Set([
  "0BSD",
  "Apache-2.0",
  "BlueOak-1.0.0",
  "BSD-2-Clause",
  "BSD-3-Clause",
  "CC-BY-4.0",
  "CC0-1.0",
  "ISC",
  "MIT",
  "MIT-0",
  "MPL-2.0",
  "Python-2.0",
  "Zlib",
]);

// Knowingly accepted copyleft exceptions, pinned by version so an upgrade is reviewed again.
const EXCEPTIONS = new Set(["bibliojson@4.0.4", "citeproc@2.4.63"]);

// "A OR B" needs one allowed side and "A AND B" needs both; anything else (WITH, nested brackets, no licence at all) fails closed.
const allowed = (licence) => {
  const expr = licence.replace(/^\((.*)\)$/, "$1");
  return (
    !/[()]/.test(expr) &&
    expr
      .split(" OR ")
      .some((term) => term.split(" AND ").every((id) => ALLOWED.has(id)))
  );
};

const lock = JSON.parse(
  readFileSync(new URL("../package-lock.json", import.meta.url), "utf8"),
);
const packages = Object.entries(lock.packages)
  .filter(([path]) => path)
  .map(([path, { version, license }]) => ({
    name: path.replace(/^.*node_modules\//, ""),
    version,
    license,
  }));
const refused = packages.filter(
  ({ name, version, license }) =>
    !EXCEPTIONS.has(`${name}@${version}`) &&
    (typeof license !== "string" || !allowed(license)),
);

for (const { name, version, license } of refused)
  console.error(`${name}@${version}: ${license ?? "no licence declared"}`);
if (refused.length) process.exit(1);
console.log(`license:check: all ${packages.length} packages allowed`);
