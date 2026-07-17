#!/usr/bin/env python3
"""check_design.py — the design-conventions gate for theweltyproject.com.

Run after ANY change to site pages, generator scripts, or family groups:

    python3 scripts/check_design.py

Verifies (see DESIGN-CONVENTIONS.md):
  1. every page links welty.css + both display fonts and carries the site chrome
  2. the family colour palette agrees across tree / people-map script / timeline chips
  3. every non-Eden roster family on the tree has a chip, a card colour, and a
     profile card on the timeline, and is known to generate_people_map.py
Exits non-zero on failure so it can gate a deploy.
"""
import os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SITE = os.path.join(ROOT, "site")
fails, warns = [], []

def read(p):
    with open(p, encoding="utf-8") as f:
        return f.read()

# ---------------------------------------------------------------- 1. chrome
PAGES = ["index.html", "all-families.html", "timeline.html", "family-papers.html", "404.html"]
FULL_CHROME = {"index.html", "all-families.html", "timeline.html", "family-papers.html"}
NAV_LINKS = ["Journal", "Family Tree", "Timeline", "Family Papers"]

for page in PAGES:
    path = os.path.join(SITE, page)
    if not os.path.exists(path):
        fails.append(f"{page}: MISSING"); continue
    h = read(path)
    if "/welty.css" not in h:
        fails.append(f"{page}: does not link /welty.css")
    if "EB+Garamond" not in h or "Cormorant+SC" not in h:
        fails.append(f"{page}: missing EB Garamond / Cormorant SC font link")
    if 'class="wsite-hdr"' not in h:
        fails.append(f"{page}: missing .wsite-hdr masthead")
    else:
        hdr = h[h.index('class="wsite-hdr"'):]
        hdr = hdr[:hdr.index("</header>")]
        for l in NAV_LINKS:
            if f">{l}<" not in hdr:
                fails.append(f"{page}: masthead nav missing '{l}'")
    if page in FULL_CHROME:
        for probe, what in [('class="wsite-ftr"', "footer"), ('id="toTop"', "back-to-top"),
                            ('wsite-contact', "contact line")]:
            if probe not in h:
                fails.append(f"{page}: missing {what}")
    # per-page body font-size overrides drift away from the 19px base
    for mm in re.finditer(r"(?<![\w.#-])body\s*\{[^}]*font-size\s*:\s*([\d.]+px)", h):
        if mm.group(1) != "19px":
            warns.append(f"{page}: body font-size override {mm.group(1)} (base is 19px in welty.css)")

# ---------------------------------------------------------------- 2. colours
chart = read(os.path.join(ROOT, "scripts", "generate_chart.py"))
pmap  = read(os.path.join(ROOT, "scripts", "generate_people_map.py"))
tline = read(os.path.join(SITE, "timeline.html"))
tree_html = read(os.path.join(SITE, "all-families.html"))

tree_colors = dict(re.findall(r"--(eden|manch|swiss|md|r1a|yrk|cum|gva|san):\s*(#[0-9a-fA-F]{6})", tree_html))
fams_block = pmap[pmap.index("FAMS = {"):pmap.index("FAM_OF_ROSTER")]
map_colors = dict(re.findall(r'"(\w+)":\s*\("[^"]*",\s*"(#[0-9a-fA-F]{6})"\)', fams_block))
chip_colors = dict(re.findall(r"\.chip\.c-(\w+)\s*\{\s*background:\s*(#[0-9a-fA-F]{6})", tline))
card_colors = dict(re.findall(r"\.card\.b-(\w+)\s*\{\s*border-left:\s*5px solid\s*(#[0-9a-fA-F]{6})", tline))

PAIRS = {  # tree var -> people-map/chip key
    "eden": "pj", "manch": "manch", "swiss": "swiss", "md": "md",
    "r1a": "r1a", "yrk": "yrk", "cum": "cum", "gva": "gva", "san": "san",
}
for tvar, key in PAIRS.items():
    tc = tree_colors.get(tvar)
    if not tc:
        fails.append(f"tree page: no CSS var --{tvar}"); continue
    mc = map_colors.get(key)
    if mc and mc.lower() != tc.lower():
        fails.append(f"colour mismatch '{key}': tree {tc} vs generate_people_map FAMS {mc}")
    if key not in ("pj",):  # eden reds appear on timeline as c-main/vars, checked via FAMS
        for src, d in (("timeline chip", chip_colors), ("timeline card", card_colors)):
            c = d.get(key)
            if c and c.lower() != tc.lower():
                fails.append(f"colour mismatch '{key}': tree {tc} vs {src} {c}")

# ---------------------------------------------------------------- 3. family coverage
m = re.search(r"FAM_OF_ROSTER\s*=\s*\{([^}]*)\}", pmap)
roster_map = dict(re.findall(r'"(\w+)":"(\w+)"', m.group(1))) if m else {}
tree_fams = set(re.findall(r'data-fam="(\w+)"', tree_html)) - {"Eden"}
for fam in sorted(tree_fams):
    key = roster_map.get(fam)
    if not key:
        fails.append(f"family '{fam}' is on the tree but unknown to generate_people_map.py (FAM_OF_ROSTER)")
        continue
    if key not in map_colors:
        fails.append(f"family '{fam}': no FAMS colour in generate_people_map.py")
    if key not in chip_colors:
        fails.append(f"family '{fam}': no .chip.c-{key} on the timeline")
    if key not in card_colors:
        fails.append(f"family '{fam}': no .card.b-{key} on the timeline")
    if f'id="fam-{key}"' not in tline:
        fails.append(f"family '{fam}': no profile card id=\"fam-{key}\" in the timeline #others chapter")

# people-map fragment freshness: legend must mention every FAMS colour
for key, col in map_colors.items():
    if col not in tline:
        fails.append(f"timeline people-map fragment is stale: colour {col} ('{key}') absent — re-run generate_people_map.py")

# ---------------------------------------------------------------- report
for w in warns:
    print("  ~ warn:", w)
if fails:
    print("DESIGN CHECK FAILED:")
    for f_ in fails:
        print("  ✗", f_)
    sys.exit(1)
print(f"Design check passed — {len(PAGES)} pages, {len(tree_fams)} non-Eden families, palette in sync.")
