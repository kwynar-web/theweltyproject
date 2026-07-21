#!/usr/bin/env python3
"""Push the freshly generated, cleaned family-tree DATA into the deployed pages
WITHOUT re-running inject_seo (so each page keeps its existing SEO / nav / footer).

EN pages  -> take the fresh generator DATA verbatim (English families + clean people)
          -> patch the one render line so notes_html (organized + linked) is shown.

The German-LANGUAGE tree was fully retired 6 Jul 2026 (never went live, only
wasted regen cycles). This script now touches English pages only.
"""
import re, json, os

# ROOT derived from this script's own location (scripts live in "Other stuff/scripts/"
# under the project root) — fixes the per-session hardcoded mount path that broke re-runs.
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
FRESH = os.path.join(ROOT, "Welty Family Tree - All Families.html")

OLD_RENDER = '  if(p.notes) h+=`<div class="notes">${esc(p.notes)}</div>`;'
NEW_RENDER = ('  if(p.notes_html) h+=`<div class="notes">${p.notes_html}</div>`;\n'
              '  else if(p.notes) h+=`<div class="notes">${esc(p.notes)}</div>`;')

DATA_LINE = re.compile(r'(?m)^const DATA = .*$')

# The visible header count ("N people · Edenkoben (German) family N · Swiss N")
# is baked into the page shell, which this script deliberately preserves — so it
# went stale (stuck at 227) while DATA kept updating. Patch it from the fresh file
# too (added 10 Jul 2026).
COUNT_RE = re.compile(r"\d+ people · Edenkoben \(German\) family \d+ · York \d+")
# ...and the intro paragraph carries a second baked-in count ("<b>N</b> people
# tracked so far") — same fix (added 10 Jul 2026, spotted by Kwyn).
TRACKED_RE = re.compile(r"<b>\d+</b> people tracked so far")

def get_data_line(text):
    m = DATA_LINE.search(text)
    return m.group(0)

def get_data(text):
    line = get_data_line(text)
    return json.loads(line[len("const DATA = "):].rstrip(";"))

EN_TARGETS = [
    "theweltyproject-site/site/all-families.html",
]

fresh_text = open(FRESH, encoding="utf-8").read()
fresh_line = get_data_line(fresh_text)
fresh_data = get_data(fresh_text)
print("fresh people:", len(fresh_data["people"]),
      "| have notes_html:", sum(1 for p in fresh_data["people"].values() if p.get("notes_html")))

# ---------------------------------------------------------------------------
# Record-image chips (added 11 Jul 2026). The chip DATA rides in via the DATA
# line above; the *rendering* code (CSS, JS helpers, lightbox DOM, and the one
# nodeHTML render line) lives in the page shell, which this script preserves.
# So we transplant those marker-wrapped blocks from the FRESH generator output
# into each target, idempotently (skip if already present).
def _block(text, start, end):
    i = text.find(start); j = text.find(end)
    return text[i:j+len(end)] if i != -1 and j != -1 else ""

REC_CSS    = _block(fresh_text, "/*RECORDS-CSS-START*/", "/*RECORDS-CSS-END*/")
REC_JS     = _block(fresh_text, "/*RECORDS-JS-START*/",  "/*RECORDS-JS-END*/")
REC_LB     = _block(fresh_text, "<!--RECORDS-LB-START-->", "<!--RECORDS-LB-END-->")
REC_RENDER = '  if(p.records&&p.records.length) h+=recStrip(p.records); /*RECORDS-RENDER*/'
SRC_LINE   = '  if(p.source_html) h+=`<div class="src"><b>Source:</b> ${p.source_html}</div>`;'

def _put(t, start, end, fresh, anchor):
    """Replace the marked block if present (keeps the live page in sync with the
    generator), else insert `fresh` before the first `anchor`. Idempotent."""
    if not fresh:
        return t, None
    i = t.find(start)
    if i != -1:
        j = t.find(end, i)
        if j != -1:
            j += len(end)
            if t[i:j] == fresh:
                return t, "unchanged"
            return t[:i] + fresh + t[j:], "replaced"
    if anchor in t:
        return t.replace(anchor, fresh + "\n" + anchor, 1), "inserted"
    return t, "no-anchor"

def inject_records(t):
    added = {}
    t, added["css"] = _put(t, "/*RECORDS-CSS-START*/", "/*RECORDS-CSS-END*/", REC_CSS, "</style>")
    t, added["js"]  = _put(t, "/*RECORDS-JS-START*/",  "/*RECORDS-JS-END*/",  REC_JS,  "function nodeHTML(")
    t, added["lb"]  = _put(t, "<!--RECORDS-LB-START-->", "<!--RECORDS-LB-END-->", REC_LB, "</body>")
    if "/*RECORDS-RENDER*/" not in t and SRC_LINE in t:
        t = t.replace(SRC_LINE, SRC_LINE + "\n" + REC_RENDER, 1); added["render"] = "inserted"
    return t, added

def patch_file(path, data_line, count_str, tracked_str):
    fp = os.path.join(ROOT, path)
    if not os.path.exists(fp):
        print("  SKIP (missing):", path); return
    t = open(fp, encoding="utf-8").read()
    n_data = len(DATA_LINE.findall(t))
    t2 = DATA_LINE.sub(lambda m: data_line, t, count=1)
    n_render = t2.count(OLD_RENDER)
    t2 = t2.replace(OLD_RENDER, NEW_RENDER, 1)
    t2, rec_added = inject_records(t2)
    n_count = len(COUNT_RE.findall(t2))
    if count_str:
        t2 = COUNT_RE.sub(lambda m: count_str, t2)
    n_tracked = len(TRACKED_RE.findall(t2))
    if tracked_str:
        t2 = TRACKED_RE.sub(lambda m: tracked_str, t2)
    open(fp, "w", encoding="utf-8").write(t2)
    print(f"  patched {path}  (DATA lines seen={n_data}, render line replaced={n_render}, count labels updated={n_count}, tracked-so-far updated={n_tracked}, records injected={rec_added or 'already present'})")

m = COUNT_RE.search(fresh_text)
fresh_count = m.group(0) if m else None
print("fresh count label:", fresh_count)
m2 = TRACKED_RE.search(fresh_text)
fresh_tracked = m2.group(0) if m2 else None
print("fresh tracked label:", fresh_tracked)

print("EN:")
for p in EN_TARGETS:
    patch_file(p, fresh_line, fresh_count, fresh_tracked)

print("done.")
