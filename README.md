# The Welty Project

Open-source source code and content for **[theweltyproject.com](https://theweltyproject.com)** — a
multi-year, evidence-based genealogy of the Welty / Weldy / Wälti family, tracing the
direct paternal line from present-day America back to its German and Swiss origins.

The site is a set of static HTML pages (family trees, a geographic timeline, a people
map, and a family-papers archive) generated from a private research log by the Python
scripts in `scripts/`.

## What's in this repo

```
site/       The published static website (HTML, CSS/JS, SVG crests, icons, manifest).
scripts/    Python generators that build the pages in site/ from the research data.
```

The **private research log** (a spreadsheet containing notes on living people, unpublished
leads, and other private material) is deliberately **not** included here. The generator
scripts read from it locally; without it they won't regenerate the pages, but the finished
site in `site/` is fully self-contained and can be hosted as-is.

## Licensing

This project is split into two layers, each licensed separately:

- **Code** — everything in `scripts/`, plus the page templates, CSS, and JavaScript that
  make up the site — is licensed under the **GNU Affero General Public License v3.0**
  (see [`LICENSE`](LICENSE)). AGPL is copyleft: you're free to use, modify, and even sell
  it, **but** any version you distribute *or run as a network service* must also be made
  available under the AGPL. It cannot be turned into a closed, proprietary product.

- **Content** — the written genealogical narrative, research write-ups, and the underlying
  factual data — is dedicated to the public domain under **CC0 1.0** (see
  [`LICENSE-CONTENT`](LICENSE-CONTENT)). Genealogical facts aren't copyrightable in the
  first place, so the content is simply free for anyone to reuse, quote, or build on.

In plain terms: **the software that runs this site can never be locked up behind a
paywall, and the family history itself belongs to everyone.**

If you republish or build on the content, a credit and a link back to
theweltyproject.com is appreciated but not required.

## Hosting the site

The `site/` folder is a plain static site — no build step, no server code. You can host it
by dropping the folder onto any static host (Netlify, GitHub Pages, Cloudflare Pages, etc.)
or serve it locally:

```bash
cd site
python3 -m http.server 8000
# then open http://localhost:8000
```

## Regenerating the pages (maintainers)

The scripts in `scripts/` expect a local copy of the private research log. They are shared
here for transparency and so the site is reproducible by anyone who holds the data — not as
a turnkey pipeline. Key scripts:

- `generate_chart.py` — builds the family-tree HTML pages
- `generate_people_map.py` — builds the interactive people map
- `generate_web_assets.py` — generates favicons / social-share images from the crests
- `inject_seo.py` — injects meta tags, the shared nav, and footer into each page

## A note on accuracy

This is living research. Conclusions are sourced and revised as new evidence appears;
rejected family lore is labelled as such. If you spot an error or hold a record that
bears on the line, corrections are welcome via the contact link on the site.
