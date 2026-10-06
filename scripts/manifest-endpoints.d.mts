/** Types for `manifest-endpoints.mjs`, kept in step with it by hand since the module is plain Node ESM but `tests/manifest-endpoints.test.ts` is still type-checked against it. */

/** A URL a manifest may ship, already parsed. */
export interface PublicEndpoint {
  ok: true;
  url: URL;
}

/** A URL it may not, and the derived class it falls in. */
export interface RefusedEndpoint {
  ok: false;
  reason: string;
}

/** Is this a URL a distributable manifest may point at: https, no credentials, and a host that is a globally reachable address or a name under a public domain? */
export declare function classifyEndpoint(
  value: string,
): PublicEndpoint | RefusedEndpoint;

/** Every reason the manifest in `xml` may not be distributed, as sentences; empty means its URLs are fit to publish. */
export declare function manifestEndpointFailures(xml: string): string[];
