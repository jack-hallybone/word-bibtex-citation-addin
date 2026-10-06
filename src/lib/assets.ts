// The only network calls the add-in makes, in one place so they can be audited.

import {
  BUILTIN_STYLES,
  CSL_LOCALES,
  LocaleId,
  StyleId,
  builtinStyle,
} from "./styles";

const ASSET_BASE = "assets/csl/";
const MAX_ASSET_BYTES = 5 * 1024 * 1024;
const ASSET_TIMEOUT_MS = 10_000;

const stylePath = (style: StyleId): string => `${ASSET_BASE}${style}.csl`;
const localePath = (locale: LocaleId): string =>
  `${ASSET_BASE}${CSL_LOCALES[locale].file}`;

// Binds the bundle to the exact CSL/locale files it was reviewed with, so a stale or replaced server response fails closed.
const EXPECTED_SHA256: Readonly<Record<string, string>> = {
  ...Object.fromEntries(
    BUILTIN_STYLES.map((style) => [stylePath(style.id), style.sha256]),
  ),
  ...Object.fromEntries(
    Object.values(CSL_LOCALES).map((locale) => [
      `${ASSET_BASE}${locale.file}`,
      locale.sha256,
    ]),
  ),
};

/** The locales a style renders with: en-US, which citeproc reads beneath any other, and the style's own. */
function localesFor(style: StyleId): LocaleId[] {
  const { locale } = builtinStyle(style);
  return locale === "en-US" ? ["en-US"] : ["en-US", locale];
}

export type AssetLoader = (path: string) => Promise<string>;

/** Read and decode a response without ever buffering past the asset limit. */
export async function readBoundedText(
  response: Response,
  path: string,
): Promise<string> {
  const declared = response.headers.get("content-length");
  if (declared !== null) {
    const declaredBytes = Number(declared);
    if (Number.isFinite(declaredBytes) && declaredBytes > MAX_ASSET_BYTES) {
      throw new Error(`citation asset ${path} is unexpectedly large`);
    }
  }

  if (response.body === null) {
    throw new Error(
      `citation asset ${path} could not be read as a bounded stream`,
    );
  }

  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let bytes = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      bytes += value.byteLength;
      if (bytes > MAX_ASSET_BYTES) {
        await reader.cancel();
        throw new Error(`citation asset ${path} is unexpectedly large`);
      }
      chunks.push(value);
    }
  } catch (error) {
    try {
      await reader.cancel();
    } catch {
      // Preserve the original network/size failure.
    }
    throw error;
  }

  const combined = new Uint8Array(bytes);
  let offset = 0;
  for (const chunk of chunks) {
    combined.set(chunk, offset);
    offset += chunk.byteLength;
  }
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(combined);
  } catch {
    throw new Error(`citation asset ${path} is not valid UTF-8`);
  }
}

async function sha256(value: string): Promise<string> {
  if (!globalThis.crypto?.subtle)
    throw new Error(
      "Web Crypto is unavailable; citation assets cannot be verified",
    );
  const digest = await globalThis.crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(value),
  );
  return Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
}

async function loadVerified(load: AssetLoader, path: string): Promise<string> {
  const expected = EXPECTED_SHA256[path];
  if (expected === undefined)
    throw new Error(
      `No integrity value is registered for citation asset ${path}`,
    );
  const text = await load(path);
  const actual = await sha256(text);
  if (actual !== expected) {
    throw new Error(
      `Citation asset ${path} failed its SHA-256 integrity check`,
    );
  }
  return text;
}

/** Same-origin asserted rather than assumed, redirects refused, and the response bounded by both the declared and the actual size. */
export async function fetchText(path: string): Promise<string> {
  const url = new URL(path, location.href);
  if (url.origin !== location.origin) {
    throw new Error(
      `Refused to load a cross-origin citation asset: ${url.origin}`,
    );
  }

  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), ASSET_TIMEOUT_MS);
  try {
    const response = await fetch(url.toString(), {
      cache: "no-cache",
      credentials: "same-origin",
      redirect: "error",
      signal: controller.signal,
    });
    if (!response.ok)
      throw new Error(`could not load ${path} (${response.status})`);

    return await readBoundedText(response, path);
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      throw new Error(`timed out loading citation asset ${path}`);
    }
    throw error;
  } finally {
    window.clearTimeout(timeout);
  }
}

/** Cached per loader, not per engine: `style`'s StyleId allowlist — not the same-origin check — is what stops a hostile embedded style id from picking an arbitrary same-origin file. */
export class AssetCache {
  private readonly locales = new Map<LocaleId, string>();
  private readonly styles = new Map<StyleId, string>();

  constructor(private readonly load: AssetLoader) {}

  /** Each file is fetched the first time a style needs it, so a style never chosen costs nothing. */
  async ensure(style: StyleId): Promise<void> {
    for (const locale of localesFor(style)) {
      if (!this.locales.has(locale)) {
        this.locales.set(
          locale,
          await loadVerified(this.load, localePath(locale)),
        );
      }
    }
    if (!this.styles.has(style)) {
      this.styles.set(style, await loadVerified(this.load, stylePath(style)));
    }
  }

  ready(
    style: StyleId,
  ): { locales: ReadonlyMap<string, string>; style: string } | null {
    const styleXml = this.styles.get(style);
    if (
      styleXml === undefined ||
      localesFor(style).some((locale) => !this.locales.has(locale))
    )
      return null;
    return { locales: this.locales, style: styleXml };
  }
}

const assetCaches = new WeakMap<AssetLoader, AssetCache>();

export function cacheFor(load: AssetLoader): AssetCache {
  let cache = assetCaches.get(load);
  if (cache === undefined) {
    cache = new AssetCache(load);
    assetCaches.set(load, cache);
  }
  return cache;
}
