#!/usr/bin/env python3
"""Generate the Dover Township schematic property plat (SVG, site style).

A RECONSTRUCTION, not a survey: colonial PA used metes-and-bounds, so no exact
parcel coordinates survive. Each household is drawn as a parcel whose AREA is
proportional to its real taxed acreage, arranged by the documented tax-roll
adjacency (1779 & 1782-83 rolls, where order is real; the 1787 roll is
alphabetised so its order is meaningless). Sources: PA Archives 3rd ser. vol XXI
Dover tax lists 1779-1783 (M2/M3/M14/M15/PL1), Ancestry coll.2497 (1787, P26),
Strayer's/Salem Reformed baptisms (SR1-11), First Trinity Reformed York marriages.

Out:  ../../Welty Dover Township Plat.html   (project root, standalone)
Run:  python3 generate_dover_plat.py
"""
import os, math, re

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(ROOT, "Welty Dover Township Plat.html")

S = 6.2  # px per sqrt(acre)

# name, acres, short-acre label, subnote (or ""), cx, cy, kind, direct, dashed
# kind: pj (Edenkoben/Philip-Jacob line) | gw (Manchester) | inlaw | other
PARCELS = [
    ("Georg Wolfgang", 120, "120 ac", "", 190, 250, "gw", False, False),
    ("the Widow",       80, "80 ac · 1780", "likely Anna Maria", 190, 415, "gw", False, True),
    ("Philip Jacob",   100, "100 → 80 ac", "", 375, 250, "pj", True, False),
    ("Jacob",           80, "80 ac", "", 525, 250, "pj", False, False),
    ("Michael",         44, "44 ac", "", 375, 415, "pj", True, False),
    ("John",            80, "80 ac", "", 525, 415, "pj", False, False),
    ("John Ruthrauff", 195, "195 ac", "Christina’s family · Michael m. in, 1784", 745, 330, "inlaw", False, False),
    ("a Swiss Welty",   36, "unrelated (I2b)", "", 955, 235, "other", False, False),
    ("Paul Wilt",      250, "250 ac", "separate family (Wilt)", 960, 470, "other", False, False),
]

CHURCHES = [
    (375, 150, "Strayer’s (Salem) Reformed", "in-township · the baptisms"),
    (360, 665, "First Trinity Reformed, York", "~6 mi SE · the marriages"),
    (790, 150, "Quickel’s · Conewago", "NE · Christina’s baptism"),
]

MARRIAGES = [
    ("Elizabeth Welty", "George Gauf", "26 Mar 1775", "hypo"),
    ("John Welty", "Margaret Ilgenfritz", "17 Aug 1783", "hypo"),
    ("Christina Welty", "Peter Messerle", "13 Apr 1784", "hypo"),
    ("Michael Welty", "Christina Ruthrauff", "18 May 1784", "proven"),
    ("Catharina Welty", "Jacob Boehm", "10 May 1785", "hypo"),
    ("Jacob Welty", "Anna Maria Miller", "10 May 1785", "open"),
]

COL = {"pj":"#8a5a2b", "gw":"#46618a", "inlaw":"#5b7a4a", "other":"#9b9285"}
TINT = {"pj":"#efe1cf", "gw":"#dde5f0", "inlaw":"#e2ecd9", "other":"#e9e5dd"}

def side(acres): return math.sqrt(acres) * S

def parcels_svg():
    out = []
    for name, ac, aclab, sub, cx, cy, kind, direct, dashed in PARCELS:
        s = side(ac); x = cx - s/2; y = cy - s/2
        col = COL[kind]; tint = TINT[kind]
        dash = ' stroke-dasharray="6 5"' if dashed else ''
        star = f'<text x="{cx+s/2-9:.0f}" y="{y+15:.0f}" class="star" text-anchor="middle">&#9733;</text>' if direct else ''
        subtxt = f'<text x="{cx}" y="{y+s+27:.0f}" class="psub" text-anchor="middle">{sub}</text>' if sub else ''
        out.append(f'''  <g class="parcel">
    <title>{name} — {aclab} {("· "+sub) if sub else ""}</title>
    <rect x="{x:.0f}" y="{y:.0f}" width="{s:.0f}" height="{s:.0f}" rx="3"
          fill="{tint}" stroke="{col}" stroke-width="2.4"{dash}/>
    {star}
    <text x="{cx}" y="{y-6:.0f}" class="pname" text-anchor="middle" fill="{col}">{name}</text>
    <text x="{cx}" y="{y+s+14:.0f}" class="pac" text-anchor="middle">{aclab}</text>
    {subtxt}
  </g>''')
    return "\n".join(out)

def church_svg():
    out = []
    for x, y, name, sub in CHURCHES:
        out.append(f'''  <g class="church">
    <title>{name} — {sub}</title>
    <path d="M{x} {y-13} L{x+9} {y-4} L{x-9} {y-4} Z" fill="#7a2e2e"/>
    <rect x="{x-7}" y="{y-4}" width="14" height="12" fill="#7a2e2e"/>
    <rect x="{x-1.2}" y="{y-11}" width="2.4" height="6" fill="#fffdf8"/>
    <text x="{x}" y="{y+22}" class="cname" text-anchor="middle">{name}</text>
    <text x="{x}" y="{y+34}" class="csub" text-anchor="middle">{sub}</text>
  </g>''')
    return "\n".join(out)

# dashed bracket under Philip + his sons (the divided 100 acres)
CARVE = '''  <path d="M330 470 L330 478 L570 478 L570 470" fill="none" stroke="#8a5a2b" stroke-width="1.5"/>
  <text x="450" y="496" class="carve" text-anchor="middle">the sons &mdash; carved from Philip Jacob&rsquo;s 100 acres (tax rolls, 1782&ndash;83)</text>'''

def marriages_html():
    badge = {"proven":"pm-proven","hypo":"pm-hypo","open":"pm-open"}
    label = {"proven":"proven","hypo":"hypothesis","open":"unassigned"}
    rows = []
    for w, sp, date, grade in MARRIAGES:
        rows.append(f'<li><b>{w}</b> &#9901; {sp} <span class="mdate">{date}</span> '
                    f'<span class="pm-badge {badge[grade]}">{label[grade]}</span></li>')
    return "\n        ".join(rows)

def build():
    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>The Weltys of Dover Township &mdash; a Property Plat</title>
<link href="https://fonts.googleapis.com/css2?family=EB+Garamond:ital@0;1&display=swap" rel="stylesheet">
<style>
  body{{margin:0;background:#f5efe3;color:#2b2620;font-family:'EB Garamond',Georgia,serif;line-height:1.55;}}
  .wrap{{max-width:1040px;margin:0 auto;padding:34px 18px 80px;}}
  h1{{text-align:center;font-size:1.95em;margin:0 0 4px;letter-spacing:.5px;}}
  .subtitle{{text-align:center;color:#6b6156;font-style:italic;margin-bottom:6px;}}
  .updated{{text-align:center;font-size:.82em;color:#7a2e2e;font-weight:bold;margin-bottom:18px;}}
  a{{color:#7a2e2e;}}
  .intro{{max-width:800px;margin:0 auto 14px;color:#4a453d;font-size:.92em;text-align:center;}}
  .legend{{display:flex;flex-wrap:wrap;justify-content:center;gap:8px 18px;font-size:.84em;margin:0 auto 6px;max-width:940px;}}
  .legend span{{display:inline-flex;align-items:center;gap:6px;}}
  .sw{{width:15px;height:12px;border-radius:2px;display:inline-block;border:2px solid;}}
  .plat{{background:#fffdf8;border:1px solid #e3d9c4;border-radius:10px;padding:8px;margin:8px 0 16px;box-shadow:0 1px 3px rgba(60,45,20,.08);}}
  svg{{width:100%;height:auto;display:block;}}
  .parcel rect{{transition:stroke-width .12s;}}
  .parcel:hover rect{{stroke-width:4;}}
  .pname{{font-size:13.5px;font-weight:600;font-family:Georgia,serif;paint-order:stroke;stroke:#fffdf8;stroke-width:3px;}}
  .pac{{font-size:11px;fill:#5a5348;font-family:Georgia,serif;paint-order:stroke;stroke:#fffdf8;stroke-width:2.5px;}}
  .psub{{font-size:10px;fill:#7a6f5c;font-style:italic;font-family:Georgia,serif;paint-order:stroke;stroke:#fffdf8;stroke-width:2.5px;}}
  .star{{fill:#c69b3a;font-size:15px;}}
  .cname{{font-size:11.5px;fill:#7a2e2e;font-weight:600;font-family:Georgia,serif;paint-order:stroke;stroke:#fffdf8;stroke-width:3px;}}
  .csub{{font-size:9.5px;fill:#8a7a58;font-style:italic;font-family:Georgia,serif;paint-order:stroke;stroke:#fffdf8;stroke-width:2.5px;}}
  .carve{{font-size:11px;fill:#8a5a2b;font-style:italic;font-family:Georgia,serif;paint-order:stroke;stroke:#fffdf8;stroke-width:3px;}}
  .frame-label{{font-size:15px;fill:#8a7a52;font-style:italic;letter-spacing:1px;font-family:Georgia,serif;}}
  .edge{{font-size:11px;fill:#9a8b66;font-style:italic;font-family:Georgia,serif;}}
  .panel{{background:#fbf5e8;border:1px solid #e3d9c4;border-left:5px solid #5b7a4a;border-radius:8px;padding:14px 20px;margin:0 auto 16px;max-width:840px;}}
  .panel h2{{font-size:1.05em;color:#4a6238;margin:0 0 8px;}}
  .panel ul{{margin:0;padding-left:18px;columns:2;column-gap:28px;font-size:.92em;}}
  .panel li{{margin:3px 0;break-inside:avoid;}}
  .mdate{{color:#6b6156;font-size:.85em;}}
  .pm-badge{{font-size:.7em;padding:1px 6px;border-radius:9px;border:1px solid;margin-left:2px;white-space:nowrap;}}
  .pm-proven{{color:#3d6b3d;border-color:#3d6b3d;background:#eef5ee;}}
  .pm-hypo{{color:#9a6a1e;border-color:#9a6a1e;background:#f7f0e2;}}
  .pm-open{{color:#7a6a5a;border-color:#7a6a5a;background:#efeae2;}}
  .note{{max-width:840px;margin:0 auto;font-size:.82em;color:#6b6156;text-align:center;font-style:italic;}}
</style>
</head>
<body>
<div class="wrap">
  <h1>The Weltys of Dover Township</h1>
  <p class="subtitle">Who held what &mdash; and who married whom &mdash; on the York County frontier, c. 1779&ndash;1789</p>
  <p class="updated">A reconstruction from the tax rolls &middot; 4 July 2026</p>

  <p class="intro">By 1779 the immigrant brothers&rsquo; households sat side by side in Dover. The tax rolls then show a <b>father settling his sons</b>: Philip Jacob&rsquo;s 100 acres shrink to 80 as a new &ldquo;Jacob, 80 ac&rdquo; and &ldquo;Michael, 44 ac&rdquo; appear beside him. And the family married into the neighbours &mdash; most tellingly the <b>Ruthrauffs</b> on the next big farm.</p>

  <div class="legend">
    <span><i class="sw" style="background:#efe1cf;border-color:#8a5a2b"></i> Welty &mdash; Edenkoben (Philip Jacob) line</span>
    <span><i class="sw" style="background:#dde5f0;border-color:#46618a"></i> Welty &mdash; Manchester (Georg Wolfgang)</span>
    <span><i class="sw" style="background:#e2ecd9;border-color:#5b7a4a"></i> in-law family (married in)</span>
    <span><i class="sw" style="background:#e9e5dd;border-color:#9b9285"></i> separate family (not our line)</span>
    <span><span class="star" style="color:#c69b3a;font-size:1.2em">&#9733;</span> direct line to today</span>
  </div>

  <div class="plat">
  <svg viewBox="0 0 1200 720" xmlns="http://www.w3.org/2000/svg">
    <rect x="30" y="70" width="1140" height="600" rx="14" fill="#f3ecd8" stroke="#b8a06a" stroke-width="2" stroke-dasharray="3 6"/>
    <text x="600" y="55" class="frame-label" text-anchor="middle">DOVER TOWNSHIP &mdash; York County, Pennsylvania</text>
    <g stroke="#b8a06a" stroke-width="1.4" fill="none"><path d="M1105 105 L1105 140 M1105 105 L1099 117 M1105 105 L1111 117"/></g>
    <text x="1105" y="156" class="edge" text-anchor="middle">N</text>
    <text x="1120" y="650" class="edge" text-anchor="end">&rarr; to York borough (SE)</text>
{CARVE}
{parcels_svg()}
{church_svg()}
    <g class="church">
      <rect x="632" y="545" width="18" height="14" fill="#8a7a58"/>
      <rect x="654" y="549" width="12" height="10" fill="#8a7a58"/>
      <text x="648" y="577" class="csub" text-anchor="middle" style="font-size:10.5px;fill:#6b6156;">Dover town &mdash; Philip&rsquo;s house + 2 lots (1787)</text>
    </g>
  </svg>
  </div>

  <div class="panel">
    <h2>Married into the family &mdash; First Trinity Reformed, York, 1775&ndash;1785</h2>
    <ul>
        {marriages_html()}
    </ul>
    <p style="margin:10px 0 0;font-size:.85em;color:#6b6156;">Eight Welty marriages cluster at Trinity in these years. The three daughters marrying Gauf, Messerle &amp; Boehm are a <b>working hypothesis</b> as Philip Jacob&rsquo;s daughters (a sibling set); Jacob&rsquo;s 1785 marriage is still unassigned. Sponsors at Strayer&rsquo;s tie in further neighbours &mdash; Paulus &amp; Christine Wild, Daniel Messerle, Jacob Miller.</p>
  </div>

  <p class="note">A schematic reconstruction. Parcels are sized to their real <b>taxed acreage</b> and arranged by the documented tax-roll adjacency (1779 and 1782&ndash;83, where order is meaningful); colonial Pennsylvania recorded land by metes-and-bounds, so exact parcel shapes and positions do not survive. Paul Wilt (250 ac) was long mistaken for a Welty &mdash; his 1803 will proves the Wilt family, a separate line.</p>
</div>
</body>
</html>'''

def fix_svg_entities(html):
    named = {"&mdash;":"&#8212;", "&ndash;":"&#8211;", "&rarr;":"&#8594;",
             "&larr;":"&#8592;", "&rsquo;":"&#8217;", "&lsquo;":"&#8216;",
             "&middot;":"&#183;", "&hellip;":"&#8230;"}
    def repl(m):
        s = m.group(0)
        for k,v in named.items():
            s = s.replace(k, v)
        return s
    return re.sub(r"<svg.*?</svg>", repl, html, flags=re.S)

html = fix_svg_entities(build())
with open(OUT, "w") as f:
    f.write(html)
print("wrote", OUT)
