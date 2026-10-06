import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { defineConfig, type Plugin } from "vite";
import { browserFloor } from "./scripts/browser-floor.mjs";
import {
  classifyEndpoint,
  manifestEndpointFailures,
} from "./scripts/manifest-endpoints.mjs";

const abs = (path: string): string =>
  fileURLToPath(new URL(path, import.meta.url));

/** The origin manifest.xml's URLs are written with; every build swaps in VITE_SITE_URL's. */
const TEMPLATE_ORIGIN = "https://localhost:3000";

/** One token in three languages: substituted into TypeScript, HTML and XML below. */
const PRODUCT_NAME_TOKEN = "__PRODUCT_NAME__";

function productName(): string {
  const { productName: declared } = JSON.parse(
    readFileSync(abs("package.json"), "utf8"),
  ) as { productName?: unknown };
  if (typeof declared !== "string" || declared.trim() === "") {
    throw new Error("package.json declares no productName to build from.");
  }
  return declared;
}

/** Fill the pane's `__PRODUCT_NAME__` tokens, in `vite dev` as well as in a build. */
function productNameInHtml(name: string): Plugin {
  return {
    name: "bibcite:product-name",
    // Before Vite's own `%VAR%` env pass, so what it sees is already resolved.
    transformIndexHtml: {
      order: "pre",
      handler: (html) => html.split(PRODUCT_NAME_TOKEN).join(name),
    },
  };
}

/** An unset origin is a failure, not a default: a manifest left at the template's origin would grant ReadWriteDocument to whatever answers there. */
function siteBase(siteUrl: string | undefined): string {
  if (siteUrl === undefined || siteUrl === "") {
    throw new Error(
      "VITE_SITE_URL is not set, so this build has no origin to write into the manifest. " +
        "Set VITE_SITE_URL=https://your.host/path, the address the add-in is published at.",
    );
  }

  let parsed: URL;
  try {
    parsed = new URL(siteUrl);
  } catch {
    throw new Error(`VITE_SITE_URL is not a valid URL: ${siteUrl}`);
  }
  if (parsed.protocol !== "https:") {
    throw new Error(`VITE_SITE_URL must use https, got ${parsed.protocol}//`);
  }
  const verdict = classifyEndpoint(parsed.origin + "/");
  if (!verdict.ok) {
    throw new Error(
      `VITE_SITE_URL cannot be published: its host ${verdict.reason}.`,
    );
  }
  return (parsed.origin + parsed.pathname).replace(/\/+$/, "");
}

function manifestXml(base: string, name: string): string {
  const source = readFileSync(abs("manifest.xml"), "utf8");
  const xml = source
    .split(PRODUCT_NAME_TOKEN)
    .join(name)
    .split(TEMPLATE_ORIGIN)
    .join(base);

  // Same check the artefact gate applies to the shipped file, so a build that can't produce a publishable manifest fails here too.
  const refusals = manifestEndpointFailures(xml);
  if (refusals.length > 0) {
    throw new Error(
      `manifest.xml cannot be published after substitution:\n  ${refusals.join("\n  ")}`,
    );
  }
  return xml;
}

/** The notices are a licence obligation for the bundled CPAL/AGPL and LGPL deps. */
function emitDistributionFiles(base: string, name: string): Plugin {
  return {
    name: "bibcite:distribution-files",
    apply: "build",
    generateBundle() {
      this.emitFile({
        type: "asset",
        fileName: "manifest.xml",
        source: manifestXml(base, name),
      });
      this.emitFile({
        type: "asset",
        fileName: "THIRD-PARTY-NOTICES.md",
        source: readFileSync(abs("THIRD-PARTY-NOTICES.md"), "utf8"),
      });
    },
  };
}

export default defineConfig(({ command }) => {
  const name = productName();
  // Only a build emits a manifest, so only a build needs VITE_SITE_URL.
  const distribution =
    command === "build"
      ? [emitDistributionFiles(siteBase(process.env.VITE_SITE_URL), name)]
      : [];
  // tsc emits nothing, so this is the only place the browser floor reaches the bundle (tests/artefact.test.ts checks the emitted chunk against it).
  const floor = browserFloor();

  return {
    root: abs("src/taskpane"),
    // Relative URLs, so one bundle works at a domain root or under a project path.
    base: "./",
    // publicDir points at the repo-root public/, not src/taskpane's own, since root is src/taskpane here.
    publicDir: abs("public"),
    plugins: [productNameInHtml(name), ...distribution],
    define: { [PRODUCT_NAME_TOKEN]: JSON.stringify(name) },
    build: {
      modulePreload: { polyfill: false },
      // Relative to `root`; publicDir resolves the same way.
      outDir: "../../dist",
      emptyOutDir: true,
      rolldownOptions: {
        // Not index.html, so the entry must be named explicitly.
        input: abs("src/taskpane/taskpane.html"),
        output: {
          // dropConsole strips citeproc/bibliojson's own console.warn calls, which would otherwise leak .bib content to a host log; mangle: true restores the default this object replaces.
          minify: { compress: { dropConsole: true }, mangle: true },
        },
      },
      target: floor.target,
      // Targets browsers with native light-dark() so the CSS minifier doesn't rewrite it into a prefers-color-scheme polyfill that ignores the element's own color-scheme.
      cssTarget: ["chrome123", "firefox128", "safari17.5"],
      sourcemap: false,
      minify: true,
      // Raised above the real bundle size since citeproc/bibliojson dominate it and don't split usefully, so the warning still means something.
      chunkSizeWarningLimit: 1000,
    },
    // Plain HTTP, so like preview below it's a browser look, never something Word can sideload.
    server: {
      port: 3000,
      strictPort: true,
      // Loopback always.
      host: "127.0.0.1",
    },
    // Vite's own default preview port; deliberately no certificate since this is a plain-HTTP look, never something Word can sideload.
    preview: {
      port: 4173,
      strictPort: true,
      host: "127.0.0.1",
    },
  };
});
