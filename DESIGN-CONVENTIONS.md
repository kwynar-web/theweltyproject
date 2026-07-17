# Design Conventions — theweltyproject.com

**The rule: `site/welty.css` is the single source of truth. Any session that touches a
page, a generator script, or adds a family group MUST run `python3 scripts/check_design.py`
before handing the work back. If it fails, the work isn't done.**

## The system (settled 12–13 Jul 2026, reaffirmed 17 Jul 2026)

| Element | Value |
|---|---|
| Body / running text | EB Garamond (`--font-body`), 19px, line-height 1.55 |
| Headings & page titles | Cormorant SC (`--font-display`) |
| The interactive tree ONLY | Iowan Old Style, scoped to `#chart` (`--font-tree`) — the one exception, do not use anywhere else |
| Page background | `#f5efe2` parchment + the radial gradient (`radial-gradient(ellipse at 50% -10%, #fdf9ee 0%, #f5efe2 55%, #ede4cf 100%)`) on every page |
| Sub-page titles | ink `#1c1a17`, centred, **2.15rem**, letter-spacing .5px (crimson titles are reserved for the home page) |
| Accent meanings | gold = wayfinding · crimson = the argument/thesis · navy = call-to-action · tan = quiet asides (see `welty.css` `:root`) |

## Page chrome — every page gets all of it

Every page in `site/` (including 404.html) carries: the `.wsite-hdr` masthead with the
4-link nav (Journal / Family Tree / Timeline / Family Papers), `welty.css`, both Google
fonts (EB Garamond + Cormorant SC), and — except the 404 — the `.wsite-ftr` footer,
`#toTop`, and the contact line. Exception on the tree page only: `.wsite-hdr` is
`position:static` there because its search/filter bar is what pins to the top.

## Family colours — one palette, three places, must be identical

The family-group colours are defined in **three places** and `check_design.py` verifies
they never disagree:

1. `scripts/generate_chart.py` → the tree page CSS vars (`--eden`, `--manch`, `--swiss`, `--md`, `--r1a`, `--yrk`, `--cum`, `--gva`, `--san`)
2. `scripts/generate_people_map.py` → `FAMS` (timeline maps, legend, stats)
3. `site/timeline.html` → the hand-written `.chip.c-*` / `.card.b-*` rules

Current palette: Eden `#b71c1c` (trunk `#8c1d1d`, John Jacob `#e0705a`) · Manchester
`#3a4a5e` · Swiss `#5e3a5e` · Maryland `#2f6f4f` · R1a `#9a6a15` · Conewago `#0f6b6b` ·
**Cumberland `#4a3a8c`** · Goochland `#a34a2a` · Saanen `#2a4d7a`.

## When a family group is added to the tree (this is what broke on 17 Jul)

Adding a group to the roster / `generate_chart.py` is **not finished** until:

1. `generate_people_map.py` knows it: `FAMS`, `FAM_OF_ROSTER`, a gazetteer node,
   `CHAPTER` anchor, `STAT_LABELS`, and the `legend_html` order.
2. `site/timeline.html` gets its chip (`.chip.c-*`), card colour (`.card.b-*`), a
   profile card in the `#others` chapter (`id="fam-*"`), and the family counts in the
   subtitle/intro prose are updated.
3. `python3 scripts/generate_people_map.py` has been re-run.
4. `python3 scripts/check_design.py` passes.

## Publishing

Edit files in the repo `site/` folder (never only the root copies), then Kwyn pushes
via GitHub Desktop. Back up any file you touch to `Backups/` first
(`BACKUP-pre-<topic>-<timestamp>-<file>`).
