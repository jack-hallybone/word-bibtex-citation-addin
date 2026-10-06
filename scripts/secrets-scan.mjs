import { readdirSync, readFileSync } from "node:fs";
import { join, relative } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { TextDecoder } from "node:util";

const ROOT = fileURLToPath(new URL("..", import.meta.url));
// Build output is scanned too, since it ships: verify runs this after the build.
const SKIP = new Set(["node_modules", ".git"]);
const PATTERNS = [
  ["private key", /-----BEGIN (?:[A-Z0-9]+ )*PRIVATE KEY(?: BLOCK)?-----/],
  ["AWS access key", /\b(?:AKIA|ASIA)[0-9A-Z]{16}\b/],
  ["GitHub token", /\b(?:gh[pousr]_[A-Za-z0-9]{36,}|github_pat_\w{60,})/],
  ["GitLab token", /\bglpat-[\w-]{20,}/],
  ["npm token", /\bnpm_[A-Za-z0-9]{36}\b/],
  [
    "npm registry credential",
    /:_(?:authToken|auth|password)[ \t]*=[ \t]*(?!\$\{)\S{8,}/,
  ],
  ["Slack token", /\bxox[abposr]-[A-Za-z0-9-]{10,}/],
  ["Slack webhook", /hooks\.slack\.com\/services\/T\w+\/B\w+\/\w+/],
  ["Google API key", /\bAIza[\w-]{35}\b/],
  ["Stripe live key", /\b[rs]k_live_[0-9A-Za-z]{24,}/],
  ["OpenAI or Anthropic key", /\bsk-(?:ant|proj)-[\w-]{32,}/],
  ["SendGrid key", /\bSG\.[\w-]{22}\.[\w-]{43}\b/],
  ["JSON Web Token", /\beyJ[\w-]{10,}\.[\w-]{10,}\.[\w-]{10,}/],
  [
    "secret in an env-style assignment",
    /^[ \t]*(?:export[ \t]+|-[ \t]+)?[A-Z0-9_]*(?:SECRET|PASSWORD|PASSWD|TOKEN|API_?KEY)[A-Z0-9_]*[ \t]*=[ \t]*\S{8,}/m,
  ],
  // Any case, but only a bare value running to the end of the line, so code such as `password = input.value;`, `token: read(),` or `resetToken={token}` doesn't count.
  [
    "secret in a config-style assignment",
    /^[ \t]*(?:export[ \t]+|-[ \t]+)?[\w.-]*(?:secret|passw(?:or)?d|token|api[_-]?key)[\w.-]*[ \t]*[:=][ \t]*(["']?)[^\s"'(){},;]{8,}\1[ \t]*$/im,
  ],
  [
    "secret in a quoted key",
    /["'][\w.-]*(?:secret|passw(?:or)?d|token|api[_-]?key)[\w.-]*["'][ \t]*:[ \t]*["'][^"'\s]{8,}["']/i,
  ],
];

// TextDecoder, not Buffer's toString("utf16le"): on Node 24 that corrupts memory for an odd-length slice at an odd offset, and aborted this scan on a build's binaries.
const utf16 = new TextDecoder("utf-16le");
// A NUL byte means UTF-16 text or a binary file. Neither is skipped: UTF-16 is read as the text it is, in either byte order, and a binary for any ASCII inside it, such as a key pasted into a media file's metadata.
const decode = (bytes) =>
  bytes.includes(0)
    ? [
        bytes.toString("latin1"),
        utf16.decode(bytes),
        utf16.decode(bytes.subarray(1)),
      ]
    : [bytes.toString("utf8")];

const findings = [];
const scan = (path) => {
  for (const decoded of decode(readFileSync(path))) {
    // A byte-order mark would stop a match on the first line.
    const text = decoded.replace(/^\uFEFF/, "");
    for (const [kind, pattern] of PATTERNS) {
      const match = pattern.exec(text);
      if (match)
        findings.push(
          `${relative(ROOT, path)}:${text.slice(0, match.index).split("\n").length}: ${kind}`,
        );
    }
  }
};
const walk = (dir) => {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory() && !SKIP.has(entry.name)) walk(path);
    else if (entry.isFile()) scan(path);
  }
};

walk(ROOT);
for (const finding of findings) console.error(finding);
if (findings.length) process.exit(1);
console.log("secrets:scan: no credentials found");
