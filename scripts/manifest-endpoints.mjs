/** Word grants ReadWriteDocument to whatever SourceLocation names, so this classifies each URL by derived IP/DNS range rather than resolving it, since a build gate must work offline. */
import { isIP } from "node:net";

/** IPv4 prefixes IANA's Special-Purpose Address Registry marks non-global, plus the multicast/reserved blocks it omits: ranges, not addresses, since 127.0.0.2 is as much loopback as 127.0.0.1. */
const IPV4_NON_GLOBAL = [
  ["0.0.0.0", 8, "this network"],
  ["10.0.0.0", 8, "private use"],
  ["100.64.0.0", 10, "shared address space"],
  ["127.0.0.0", 8, "loopback"],
  ["169.254.0.0", 16, "link local"],
  ["172.16.0.0", 12, "private use"],
  ["192.0.0.0", 24, "IETF protocol assignments"],
  ["192.0.2.0", 24, "documentation (TEST-NET-1)"],
  ["192.88.99.0", 24, "deprecated 6to4 relay anycast"],
  ["192.168.0.0", 16, "private use"],
  ["198.18.0.0", 15, "benchmarking"],
  ["198.51.100.0", 24, "documentation (TEST-NET-2)"],
  ["203.0.113.0", 24, "documentation (TEST-NET-3)"],
  ["224.0.0.0", 4, "multicast"],
  ["240.0.0.0", 4, "reserved"],
];

/** The IPv6 equivalent, with the three ranges that EMBED an IPv4 address marked so the embedded address's own class applies: [::ffff:127.0.0.1] is loopback however it's spelled. */
const IPV6_NON_GLOBAL = [
  ["::", 128, "unspecified"],
  ["::1", 128, "loopback"],
  ["64:ff9b:1::", 48, "local-use IPv4/IPv6 translation"],
  ["100::", 64, "discard only"],
  ["2001::", 23, "IETF protocol assignments"],
  ["2001:db8::", 32, "documentation"],
  ["5f00::", 16, "segment routing"],
  ["fc00::", 7, "unique local"],
  ["fe80::", 10, "link local"],
  ["ff00::", 8, "multicast"],
];

/** IPv6 prefixes that carry an IPv4 address, and where in the 16 bytes it sits. */
const IPV6_EMBEDDED_IPV4 = [
  ["::ffff:0:0", 96, 12, "IPv4-mapped"],
  ["2002::", 16, 2, "6to4"],
];

/** Top-level names that cannot resolve publicly (IANA's Special-Use Domain Names registry plus ICANN-withheld/conventionally-private strings): the namespace equivalent of the address ranges above. */
const NON_PUBLIC_TLDS = new Set([
  "alt",
  "arpa",
  "corp",
  "domain",
  "example",
  "home",
  "host",
  "internal",
  "intranet",
  "invalid",
  "lan",
  "local",
  "localdomain",
  "localhost",
  "onion",
  "private",
  "test",
]);

/** The schema elements whose `DefaultValue` Word fetches. */
const URL_BEARING =
  /<(SourceLocation|IconUrl|HighResolutionIconUrl|SupportUrl|bt:Url|bt:Image)\b[^>]*?\bDefaultValue="([^"]*)"/g;

const ipv4ToBytes = (text) => Uint8Array.from(text.split(".").map(Number));

/** An already-valid IPv6 literal (the URL parser vetted it) as its 16 bytes. */
function ipv6ToBytes(text) {
  let head = text;
  // A dotted tail is the v4-in-v6 form; fold it into the two groups it means.
  const lastColon = head.lastIndexOf(":");
  const tail = head.slice(lastColon + 1);
  if (tail.includes(".")) {
    const octets = tail.split(".").map(Number);
    const high = ((octets[0] << 8) | octets[1]).toString(16);
    const low = ((octets[2] << 8) | octets[3]).toString(16);
    head = `${head.slice(0, lastColon + 1)}${high}:${low}`;
  }

  const [left, right] = head.split("::");
  const leading = left === "" ? [] : left.split(":");
  const trailing = right === undefined || right === "" ? [] : right.split(":");
  const groups =
    right === undefined
      ? leading
      : [
          ...leading,
          ...Array(8 - leading.length - trailing.length).fill("0"),
          ...trailing,
        ];
  if (groups.length !== 8) return null;

  const bytes = new Uint8Array(16);
  for (const [index, group] of groups.entries()) {
    const value = Number.parseInt(group, 16);
    if (!Number.isInteger(value) || value < 0 || value > 0xffff) return null;
    bytes[index * 2] = value >> 8;
    bytes[index * 2 + 1] = value & 0xff;
  }
  return bytes;
}

/** Do `bytes` fall inside the prefix `prefix`/`bits`? */
function withinPrefix(bytes, prefix, bits) {
  for (let index = 0; index * 8 < bits; index += 1) {
    const remaining = bits - index * 8;
    const mask = remaining >= 8 ? 0xff : (0xff << (8 - remaining)) & 0xff;
    if ((bytes[index] & mask) !== (prefix[index] & mask)) return false;
  }
  return true;
}

/** The range an IPv4 address falls in, or null when it is globally reachable. */
function ipv4Range(bytes) {
  for (const [prefix, bits, label] of IPV4_NON_GLOBAL) {
    if (withinPrefix(bytes, ipv4ToBytes(prefix), bits)) return label;
  }
  return null;
}

function ipv6Range(bytes) {
  for (const [prefix, bits, offset, label] of IPV6_EMBEDDED_IPV4) {
    if (!withinPrefix(bytes, ipv6ToBytes(prefix), bits)) continue;
    const embedded = bytes.slice(offset, offset + 4);
    const range = ipv4Range(embedded);
    if (range === null) return null;
    return `${label} ${embedded.join(".")}, which is ${range}`;
  }
  // The v4-compatible form, ::a.b.c.d, once the two ranges above are excluded.
  if (
    withinPrefix(bytes, ipv6ToBytes("::"), 96) &&
    !withinPrefix(bytes, ipv6ToBytes("::"), 126)
  ) {
    const embedded = bytes.slice(12, 16);
    const range = ipv4Range(embedded);
    return range === null
      ? "deprecated IPv4-compatible"
      : `IPv4-compatible ${embedded.join(".")}, which is ${range}`;
  }
  for (const [prefix, bits, label] of IPV6_NON_GLOBAL) {
    if (withinPrefix(bytes, ipv6ToBytes(prefix), bits)) return label;
  }
  return null;
}

/** Classifies one URL as `{ ok: true, url }` or `{ ok: false, reason }`, the reason naming the derived class so a refusal reads as "loopback", not "on the list". */
export function classifyEndpoint(value) {
  const refuse = (reason) => ({ ok: false, reason });

  let url;
  try {
    url = new URL(value);
  } catch {
    return refuse("is not an absolute URL, so the host cannot fetch it");
  }
  if (url.protocol !== "https:") {
    return refuse(
      `uses ${url.protocol}//; Word loads an add-in over https and nothing else`,
    );
  }
  if (url.username !== "" || url.password !== "") {
    return refuse("carries credentials in the URL");
  }

  const host = url.hostname;
  const literal = host.startsWith("[") ? host.slice(1, -1) : host;
  const family = isIP(literal);

  if (family === 4) {
    const range = ipv4Range(ipv4ToBytes(literal));
    return range === null
      ? { ok: true, url }
      : refuse(`is the IPv4 address ${literal}, which is ${range}`);
  }
  if (family === 6) {
    const bytes = ipv6ToBytes(literal);
    if (bytes === null)
      return refuse(`is an unreadable IPv6 literal [${literal}]`);
    const range = ipv6Range(bytes);
    return range === null
      ? { ok: true, url }
      : refuse(`is the IPv6 address [${literal}], which is ${range}`);
  }

  // A DNS name, where one trailing dot is the root label and says nothing about it.
  const name = host.endsWith(".") ? host.slice(0, -1) : host;
  const labels = name.split(".");
  if (
    name === "" ||
    labels.some((label) => label === "" || label.length > 63)
  ) {
    return refuse(`is not a well-formed host name (${host})`);
  }
  if (labels.length < 2) {
    return refuse(
      `is the single-label name ${name}; a public host is under a public domain, and a bare name resolves to whatever the reader's own resolver decides it means`,
    );
  }
  const tld = labels[labels.length - 1];
  if (NON_PUBLIC_TLDS.has(tld)) {
    return refuse(
      `is under .${tld}, a namespace reserved from the public DNS, so it names a machine on the reader's own network`,
    );
  }
  return { ok: true, url };
}

/** XML comments, dropped: they are documentation, not endpoints. */
const withoutComments = (xml) => xml.replaceAll(/<!--[\s\S]*?-->/g, "");

/** Every failure a manifest's URLs have, as sentences (empty means it may be distributed); called by both the build and `tests/artefact.test.ts` so the two cannot drift apart. */
export function manifestEndpointFailures(xml) {
  const failures = [];
  const body = withoutComments(xml);

  // Every attribute's absolute-URL value is checked except xmlns/xmlns:*, which is never fetched; one entry per URL keeps every place it was found.
  const candidates = new Map();
  const found = (value, where) => {
    const places = candidates.get(value) ?? new Set();
    places.add(where);
    candidates.set(value, places);
  };
  for (const [, attribute, value] of body.matchAll(
    /\s([A-Za-z_][\w:.-]*)\s*=\s*"([^"]*)"/g,
  )) {
    if (attribute === "xmlns" || attribute.startsWith("xmlns:")) continue;
    if (/^[a-z][a-z0-9+.-]*:\/\//i.test(value)) found(value, attribute);
  }
  // And any absolute URL in element text, which no attribute sweep would see.
  for (const [, text] of body.matchAll(/>([^<]+)</g)) {
    const trimmed = text.trim();
    if (/^[a-z][a-z0-9+.-]*:\/\//i.test(trimmed))
      found(trimmed, "element text");
  }

  // The elements Word actually fetches, added whatever shape their value has: a substitution that produced a relative or data: URL must be refused, not skipped for not looking like one.
  const bearing = [...body.matchAll(URL_BEARING)];
  for (const [, element, value] of bearing) {
    found(value, `<${element}> DefaultValue`);
  }
  // Guard the guard: a manifest with nothing to fetch would pass vacuously.
  if (bearing.length === 0) {
    failures.push(
      "declares no SourceLocation, icon or bt:Url endpoint at all; this check would pass over it vacuously",
    );
  }

  const origins = new Set();
  for (const [value, places] of candidates) {
    const verdict = classifyEndpoint(value);
    if (!verdict.ok) {
      // Named by element where one names it, and by the bare attribute only when nothing else does — the element name is what actually points at a line.
      const byElement = [...places].filter((place) => place.startsWith("<"));
      const where = (byElement.length > 0 ? byElement : [...places]).sort();
      failures.push(
        `${where.join(", ")}: ${JSON.stringify(value)} ${verdict.reason}`,
      );
      continue;
    }
    origins.add(verdict.url.origin);
  }

  // One origin, derived: a manifest is built by substituting a deployment base into a template, so disagreeing origins mean a substitution only partly happened.
  if (origins.size > 1) {
    failures.push(
      `names ${origins.size} different origins (${[...origins].sort().join(", ")}); every URL in a manifest comes from the one deployment base, so this one was only part-substituted`,
    );
  }

  return failures;
}
