import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";

import {
  classifyEndpoint,
  manifestEndpointFailures,
} from "../scripts/manifest-endpoints.mjs";

/** These cases are the class a manifest URL might fall in, not a couple of spellings — the regex this replaced missed every one of them. */
const REFUSED: ReadonlyArray<readonly [string, string]> = [
  // Loopback, by range and in every spelling the URL parser folds together.
  ["https://localhost:3000/taskpane.html", "the name the old regex caught"],
  ["https://LocalHost:3000/taskpane.html", "case"],
  ["https://localhost./taskpane.html", "with the root label"],
  ["https://127.0.0.1/taskpane.html", "the address the old regex caught"],
  ["https://127.0.0.2/taskpane.html", "the rest of 127/8"],
  ["https://127.1/taskpane.html", "the short form"],
  ["https://0x7f000001/taskpane.html", "hexadecimal"],
  ["https://2130706433/taskpane.html", "decimal"],
  ["https://0177.0.0.1/taskpane.html", "octal"],
  ["https://[::1]/taskpane.html", "IPv6 loopback"],
  ["https://[::ffff:127.0.0.1]/taskpane.html", "IPv4-mapped loopback"],
  ["https://[::ffff:7f00:1]/taskpane.html", "IPv4-mapped, hex spelling"],
  ["https://[2002:7f00:1::]/taskpane.html", "6to4 of loopback"],
  ["https://[::]/taskpane.html", "the unspecified address"],
  ["https://0.0.0.0/taskpane.html", "this network"],

  // The rest of the private and reserved address space.
  ["https://10.1.2.3/taskpane.html", "private use"],
  ["https://172.16.0.1/taskpane.html", "private use"],
  ["https://172.31.255.254/taskpane.html", "the far end of 172.16/12"],
  ["https://192.168.1.5/taskpane.html", "private use"],
  ["https://169.254.169.254/taskpane.html", "link local"],
  ["https://100.64.0.1/taskpane.html", "carrier-grade NAT"],
  ["https://198.18.0.1/taskpane.html", "benchmarking"],
  ["https://192.0.2.1/taskpane.html", "documentation"],
  ["https://203.0.113.9/taskpane.html", "documentation"],
  ["https://239.1.2.3/taskpane.html", "multicast"],
  ["https://255.255.255.255/taskpane.html", "broadcast"],
  ["https://[fe80::1]/taskpane.html", "IPv6 link local"],
  ["https://[fd00:1234::5]/taskpane.html", "IPv6 unique local"],
  ["https://[2001:db8::1]/taskpane.html", "IPv6 documentation"],

  // Namespaces that cannot resolve on the public internet.
  ["https://my-laptop.local/taskpane.html", "mDNS"],
  ["https://pane.localhost/taskpane.html", "the .localhost namespace"],
  ["https://dev.internal/taskpane.html", "reserved private-use TLD"],
  ["https://addin.test/taskpane.html", "reserved for testing"],
  ["https://addin.invalid/taskpane.html", "reserved as unresolvable"],
  ["https://addin.example/taskpane.html", "reserved for documentation"],
  ["https://server.corp/taskpane.html", "never delegated"],
  ["https://nas.lan/taskpane.html", "never delegated"],
  ["https://intranet/taskpane.html", "a single label"],
  ["https://buildbox/taskpane.html", "a bare host name"],

  // And the shape of the URL itself.
  ["http://cite.example.org/taskpane.html", "not https"],
  ["ftp://cite.example.org/taskpane.html", "not https"],
  ["https://user:pw@cite.example.org/taskpane.html", "credentials in the URL"],
  ["data:text/html,<script></script>", "not an https endpoint at all"],
  ["assets/icon-16.png", "relative"],
];

const ACCEPTED: readonly string[] = [
  "https://cite.example.org/taskpane.html",
  "https://user.github.io/word-bibtex-citation-addin/taskpane.html",
  "https://cite.example.org:8443/taskpane.html",
  "https://8.8.8.8/taskpane.html",
  "https://[2606:4700:4700::1111]/taskpane.html",
];

test("a manifest URL is judged by the class it falls in, not by its spelling", () => {
  for (const [value, why] of REFUSED) {
    const verdict = classifyEndpoint(value);
    assert.equal(verdict.ok, false, `${value} was accepted (${why})`);
    if (verdict.ok) continue;
    assert.notEqual(
      verdict.reason,
      "",
      `${value} was refused without a reason`,
    );
  }
  for (const value of ACCEPTED) {
    const verdict = classifyEndpoint(value);
    assert.equal(
      verdict.ok,
      true,
      `${value} was refused: ${verdict.ok ? "" : verdict.reason}`,
    );
  }
});

/** The template with `__PRODUCT_NAME__` filled and one origin substituted in. */
async function manifestFor(origin: string): Promise<string> {
  const template = await readFile("manifest.xml", "utf8");
  return template
    .split("__PRODUCT_NAME__")
    .join("Test")
    .split("https://localhost:3000")
    .join(origin);
}

test("the shipped manifest's endpoints are all public, and all one origin", async () => {
  const published = await manifestFor("https://cite.example.org");
  assert.deepEqual(manifestEndpointFailures(published), []);

  // The XML namespaces are http:// and must stay: rewriting them would break the schema, and refusing them would force this check off.
  assert.match(published, /xmlns="http:\/\/schemas\.microsoft\.com/);

  // What the build shipped before this gate existed.
  const development = await manifestFor("https://localhost:3000");
  const refusals = manifestEndpointFailures(development);
  assert.ok(
    refusals.length > 0,
    "a manifest pointing at localhost was accepted",
  );
  assert.ok(
    refusals.every((line) => /loopback|single-label/.test(line)),
    `refusals name the class: ${refusals.join(" | ")}`,
  );

  // A part-done substitution leaves two origins, whichever they are.
  const halfDone = published.replace(
    'DefaultValue="https://cite.example.org/assets/icon-16.png"',
    'DefaultValue="https://elsewhere.example.net/assets/icon-16.png"',
  );
  assert.notEqual(halfDone, published, "the icon URL to divert was not found");
  assert.ok(
    manifestEndpointFailures(halfDone).some((line) =>
      /different origins/.test(line),
    ),
    "two origins in one manifest were accepted",
  );

  // Guard the guard: a manifest with nothing to fetch must not pass vacuously.
  assert.ok(
    manifestEndpointFailures(
      "<OfficeApp><Permissions>ReadWriteDocument" +
        "</Permissions></OfficeApp>",
    ).length > 0,
    "a manifest declaring no endpoint at all was accepted",
  );

  // A relative SourceLocation is not a URL Word can resolve, and it would slip past a sweep that only looks at values already shaped like URLs.
  const relative = published.replace(
    'DefaultValue="https://cite.example.org/taskpane.html"',
    'DefaultValue="taskpane.html"',
  );
  assert.ok(
    manifestEndpointFailures(relative).some((line) =>
      /not an absolute URL/.test(line),
    ),
    "a relative SourceLocation was accepted",
  );
});
