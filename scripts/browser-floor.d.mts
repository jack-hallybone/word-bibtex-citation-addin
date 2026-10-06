/** Types for `browser-floor.mjs`, kept in step with it by hand since the module is plain Node ESM but this declaration lets the TypeScript suite type-check against it. */

/** The syntax floor: the esbuild/Vite spelling, and the year acorn parses at. */
export interface SyntaxFloor {
  target: string;
  ecmaVersion: number;
}

/** The API floor: the `lib` entry that opens the newest ES surface. */
export interface ApiFloor {
  lib: string;
  ecmaVersion: number;
}

/** The oldest JavaScript the pane may be written in, from `target`. */
export declare function browserFloor(): SyntaxFloor;

/** The newest ES surface `lib` opens. */
export declare function apiFloor(): ApiFloor;
