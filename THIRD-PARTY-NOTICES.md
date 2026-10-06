# Third-party notices

The production bundle (`dist/assets/taskpane-*.js`) statically includes these
packages. `licenses/` holds two kinds of file: each package's own LICENSE,
vendored verbatim, and the complete text of every licence those notices only
refer to. Both kinds are served with the bundle.

- citeproc-js — CPAL-1.0-or-later OR AGPL-3.0-or-later — `licenses/citeproc.txt`
  (authoritative: the package's SPDX string says `AGPL-1.0`, its licence text
  says otherwise). That file is the package's copyright notice; it grants no
  terms by itself and directs the reader to the licences, so both are vendored
  in full: `licenses/cpal-1.0.txt` and `licenses/agpl-3.0.txt`.
- bibliojson — LGPL-3.0 — `licenses/bibliojson.txt` (the renamed successor to
  biblatex-csl-converter, same author and licence). LGPL-3.0 is a set of
  additional permissions over GPL-3.0 — it "incorporates the terms and
  conditions of version 3 of the GNU General Public License" — so the base text
  is vendored beside it as `licenses/gpl-3.0.txt`.
- Its transitives XRegExp — MIT — `licenses/xregexp.txt`,
  `@babel/runtime-corejs3` — MIT — `licenses/babel-runtime-corejs3.txt`, and
  core-js-pure — MIT — `licenses/core-js-pure.txt`. An MIT notice is the whole
  licence, so these stand alone.

The build copies `licenses/` and this file into `dist/`, which is where the
obligation lands: what is published is the copy the reader receives.

The styles and locales in `assets/csl/` are redistributed **unmodified** under
[CC-BY-SA 3.0](https://creativecommons.org/licenses/by-sa/3.0/), with their
embedded authorship and `<rights>` metadata intact. Keep them byte-identical to
upstream — a modified file becomes an adaptation under BY-SA. Add a style as a
new file, never by editing one of these. They are the Citation Style Language
project's `american-medical-association`, `apa`, `chicago-author-date`,
`harvard-cite-them-right`, `ieee`, `modern-language-association`, `nature` and
`nlm-citation-sequence` (offered as Vancouver) styles, from
<https://github.com/citation-style-language/styles>, and its `en-US` and `en-GB`
locales, from <https://github.com/citation-style-language/locales>.
`npm run sync:bib-core` copies them in from `reference-manager/bib-core/csl/`,
so a style is added there.

That metadata includes nine maintainer email addresses, seven of them distinct
(one in `american-medical-association.csl`, one in `apa.csl`, four in
`ieee.csl`, two in `nature.csl` and one in `nlm-citation-sequence.csl`). They
are the CSL project's own authorship records, not this project's data: BY-SA
requires the attribution be kept intact, and the bytes are SHA-256 pinned in
`src/csl-styles.js`, so editing them would both strip the attribution and fail
the runtime integrity check. The decision is to ship them as upstream wrote
them. The suite reads that set off the styles themselves, so any _other_ address
appearing anywhere in the served folder fails the build.

The task-pane icon is Lucide's `book-marked` from
[Lucide](https://lucide.dev) `lucide-static` **1.41.0** (ISC), unmodified but
for scale; naming the version is what makes that checkable. The attribution
lives in `assets/icon-source.svg`, which ships with the bundle, and the four
`assets/icon-*.png` files are that file rasterised for the Office manifest. ISC
asks that its permission notice appear in all copies, and a bundle handed to
Word is a copy, so the full text ships as `licenses/lucide.txt`.
Office.js is loaded from Microsoft's CDN at runtime and is not redistributed
here.

## Distributing a build

citeproc-js is taken under its CPAL branch, not AGPL. CPAL and LGPL attach
obligations to any hosted or shipped bundle, independent of how this project's
own source is licensed:

- Serve this file and `licenses/` alongside the bundle; the build copies both
  into `dist/`.
- Preserve the citeproc-js attribution (the task-pane footer and this file).
- Be able to point users at corresponding source and build instructions for the
  LGPL converter so they can rebuild or substitute it. Publishing this
  repository satisfies that.

Publishing source only, with no hosted bundle, carries none of these. This is a
summary, not legal advice: confirm against the texts in `licenses/` before
distributing commercially or through AppSource.
