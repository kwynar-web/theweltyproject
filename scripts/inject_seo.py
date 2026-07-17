#!/usr/bin/env python3
"""Inject SEO / Open-Graph / Twitter / favicon meta into the published pages.

Idempotent: re-running replaces the previous block (delimited by markers), so it
is safe to run again after regenerating a tree page. Run from the "Netlify Upload"
folder.  Usage:  python3 ../Other\ stuff/scripts/inject_seo.py
"""
import re, sys

DOMAIN = "https://theweltyproject.com"
START, END = "<!--welty-seo:start-->", "<!--welty-seo:end-->"
HSTART, HEND = "<!--welty-hdr:start-->", "<!--welty-hdr:end-->"
FSTART, FEND = "<!--welty-ftr:start-->", "<!--welty-ftr:end-->"
ASTART, AEND = "<!--welty-analytics:start-->", "<!--welty-analytics:end-->"

# --- Cloudflare Web Analytics ---------------------------------------------
# Privacy-friendly, cookieless page-view stats for theweltyproject.com.
# To ACTIVATE: sign in at https://dash.cloudflare.com -> Analytics & Logs ->
# Web Analytics -> Add a site (theweltyproject.com). Cloudflare shows a snippet
# containing a token like '{"token":"abc123..."}'. Paste ONLY that token string
# below, then re-run this script and push. While the token is empty the block is
# injected as an inert HTML comment (no beacon loads), so the site is unaffected.
CF_BEACON_TOKEN = ""  # <-- paste your Cloudflare Web Analytics token here

def analytics_block():
    if not CF_BEACON_TOKEN:
        return (ASTART +
                "<!-- Cloudflare Web Analytics: paste token in inject_seo.py "
                "(CF_BEACON_TOKEN) to activate -->" + AEND)
    return (ASTART +
            "<script defer src=\"https://static.cloudflareinsights.com/beacon.min.js\" "
            "data-cf-beacon='{\"token\": \"%s\"}'></script>" % CF_BEACON_TOKEN +
            AEND)

# obfuscated contact address (assembled in JS so scrapers don't harvest the raw string)
CONTACT_JS = ('<script>(function(){var u="outreach",d="theweltyproject.com",'
              'a=document.getElementById("wsite-mail");if(a){a.href="mailto:"+u+"@"+d+'
              '"?subject=The%20Welty%20Project";a.textContent=u+"@"+d;}})();</script>')

# One consistent footer is injected on EVERY page (4 Jul 2026). Structure is identical
# everywhere and always in the same spot, right before </body>:
#   1. a page-specific NOTE (the only part that varies; set per page in PAGES["note"])
#   2. the shared contact line
#   3. the shared open-source attribution block (licence + repo + third-party credits)
# Because inject_seo now owns the WHOLE footer, each page's HTML must NOT carry its own
# hand-written <footer>; if one is re-added it will double up. The map credit
# (Leaflet + OpenStreetMap) is added only where PAGES["maps"] is true, since OSM
# requires attribution wherever its tiles appear.
REPO = "https://github.com/kwynar-web/theweltyproject"

FOOTER_CSS = """
<style>
.wsite-ftr{margin-top:42px;padding:26px 20px 30px;border-top:1px solid #cbb98d;text-align:center;
  font-family:'EB Garamond',Georgia,serif;color:#6e6353;background:#faf6ec;
  font-size:1rem;line-height:1.6;font-style:normal;letter-spacing:normal}
.wsite-ftr a{color:#7a1f1f;font-weight:600;text-decoration:none;border-bottom:1px solid #cbb98d}
.wsite-ftr a:hover{color:#1c1a17}
.wsite-note{max-width:680px;margin:0 auto 14px;font-size:.95rem;color:#55493a}
.wsite-note b{font-family:'Cormorant SC',Georgia,serif;color:#7a1f1f;letter-spacing:.03em;
  display:block;margin-bottom:4px;font-size:1.05rem}
.wsite-note .q{display:block;font-style:italic;margin:2px 0}
.wsite-note .upd{display:block;font-size:.84rem;margin-top:5px}
.wsite-note .repos{display:block;font-size:.84rem;margin-top:6px}
.wsite-contact{margin:8px 0 2px}
.wsite-oss{max-width:680px;margin:16px auto 0;padding-top:13px;border-top:1px solid #e3d6b3;
  font-size:.8rem;line-height:1.6;color:#8a7d68}
.wsite-oss a{font-weight:600}
</style>"""

MAP_CREDIT = (' Maps &copy; <a href="https://www.openstreetmap.org/copyright" target="_blank" '
              'rel="noopener">OpenStreetMap</a> contributors, rendered with '
              '<a href="https://leafletjs.com/" target="_blank" rel="noopener">Leaflet</a> '
              '(BSD-2-Clause).')

def oss_block(maps=False):
    return ('<div class="wsite-oss">'
        'Open source &mdash; the site&rsquo;s code is licensed '
        '<a href="%s/blob/main/LICENSE" target="_blank" rel="noopener">AGPL-3.0</a> '
        'and its content <a href="%s/blob/main/LICENSE-CONTENT" target="_blank" rel="noopener">CC0&nbsp;1.0</a>; '
        'the full source lives <a href="%s" target="_blank" rel="noopener">on GitHub</a>.<br>'
        'Free to keep, always &mdash; never sold or gated. The research is public domain (CC0) '
        'so anyone can copy, correct, or build on it, and the code stays open (AGPL-3.0).<br>'
        'Set in EB&nbsp;Garamond &amp; Cormorant&nbsp;SC via Google Fonts (SIL Open Font License).'
        '%s'
        ' Municipal arms of Bischheim &amp; Edenkoben and the U.S. state flags are '
        '<a href="https://commons.wikimedia.org/" target="_blank" rel="noopener">Wikimedia Commons</a> '
        'files &mdash; the ancestral villages&rsquo; civic arms, not personal family arms.'
        '</div>' % (REPO, REPO, REPO, MAP_CREDIT if maps else ""))

def footer(cfg):
    note = cfg.get("note", "")
    note_html = ('<div class="wsite-note">%s</div>' % note) if note else ""
    return FSTART + FOOTER_CSS + """
<footer class="wsite-ftr">""" + note_html + """
<div class="wsite-contact">Questions or comments? <a id="wsite-mail" href="#">get in touch</a></div>
""" + oss_block(cfg.get("maps", False)) + """
</footer>
""" + CONTACT_JS + """
""" + FEND

# ---- shared site header + nav (self-contained; scoped .wsite-* classes) --------
# NOTE (4 Jul 2026): Family Papers MUST stay in this list. This nav is rewritten
# onto every page each run, so anything missing here silently disappears from the
# whole site's navigation. On 4 Jul the Family Papers link vanished because a regen
# ran a copy of this script whose NAV had reverted to 3 items. Do NOT drop a page
# from NAV/PAGES without removing the page itself. See memory: family-papers-page.
NAV = [("index.html", "Journal", "index"),
       ("all-families.html", "Family Tree", "all"),
       ("timeline.html", "Timeline", "timeline"),
       ("family-papers.html", "Family Papers", "papers")]

def header(current):
    links = "".join(
        '<a href="%s"%s>%s</a>' % (href, ' aria-current="page"' if key == current else "", label)
        for href, label, key in NAV)
    return HSTART + """
<style>
.wsite-hdr{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;
  padding:11px 22px;background:#faf5e9;border-bottom:2px solid #b8912f;
  font-family:'EB Garamond',Georgia,serif;box-shadow:0 1px 5px rgba(60,45,20,.09);}
.wsite-hdr *{box-sizing:border-box}
.wsite-brand{display:flex;align-items:center;gap:11px;text-decoration:none;border:0}
.wsite-brand img{width:34px;height:40px;display:block}
.wsite-word{font-family:'Cormorant SC','EB Garamond',Georgia,serif;font-weight:700;font-size:1.3rem;
  letter-spacing:.5px;color:#1c1a17;line-height:1.05}
.wsite-word small{display:block;font-family:'EB Garamond',Georgia,serif;font-weight:400;font-size:.6rem;
  letter-spacing:2.5px;text-transform:uppercase;color:#8a7a52;margin-top:3px}
.wsite-nav{display:flex;gap:5px;flex-wrap:wrap}
.wsite-nav a{text-decoration:none;color:#5a4f3c;font-size:1.02rem;padding:6px 14px;border-radius:5px;
  border:1px solid transparent;transition:.15s}
.wsite-nav a:hover{background:#f0e6cf;color:#1c1a17}
.wsite-nav a[aria-current=page]{color:#7a1f1f;border-color:#cbb98d;background:#fffdf6;font-weight:600}
@media(max-width:520px){.wsite-word{font-size:1.1rem}.wsite-nav a{padding:5px 10px;font-size:.94rem}}
</style>
<header class="wsite-hdr">
  <a class="wsite-brand" href="index.html">
    <img src="favicon.svg" alt="Welty family crest">
    <span class="wsite-word">The Welty Family<small>Research Journal</small></span>
  </a>
  <nav class="wsite-nav">""" + links + """</nav>
</header>
""" + HEND

PAGES = {
    "index.html": dict(
        nav="index",
        path="/",
        ogtitle="The Welty Project — A Living Genealogy Journal",
        desc="A living, openly-sourced genealogy of the Welty (Wäldi/Welde) family of "
             "Edenkoben in the Palatinate — traced by DNA and primary records from "
             "c. 1680 to today.",
        jsonld=True,
        note='<b>The live research journal of Kwyn Welty.</b>',
    ),
    "all-families.html": dict(
        nav="all",
        path="/all-families.html",
        ogtitle="The Welty Family Tree — All Lines",
        desc="Explore the interactive Welty family tree: the German Edenkoben "
             "family of the Palatinate and the genuinely separate Swiss Emmental "
             "Wälti — built from primary records and Y-DNA.",
        jsonld=False,
        note='The complete interactive Welty tree &mdash; the German Edenkoben family of the '
             'Palatinate and the genuinely separate Swiss W&auml;lti line, built from primary '
             'records and Y-DNA.',
    ),
    "timeline.html": dict(
        nav="timeline",
        path="/timeline.html",
        ogtitle="The Welty Families, Place by Place",
        desc="A place-by-place history of every American Welty family — the German "
             "Edenkoben line, the Swiss Wälti, the Maryland Weltys, the Virginia "
             "Weldys and more, from their European villages to Pennsylvania and beyond.",
        jsonld=False,
        maps=True,
        note='<b>Ten generations, two continents, one river valley to another.</b>'
             '<span class="upd">Compiled from the Welty Ancestry Research Log &mdash; '
             'last updated 12 July 2026.</span>'
             '<span class="repos">Key repositories: '
             '<a href="https://www.archion.de" target="_blank" rel="noopener">Archion</a> &middot; '
             '<a href="https://www.familysearch.org/search/full-text" target="_blank" rel="noopener">FamilySearch full-text</a> &middot; '
             '<a href="https://www.familytreedna.com/public/Welty" target="_blank" rel="noopener">FTDNA Welty project</a> &middot; '
             '<a href="https://www.yorkcountyarchives.org/" target="_blank" rel="noopener">York County Archives</a> &middot; '
             '<a href="https://services.dar.org/Public/DAR_Research/search_adb/" target="_blank" rel="noopener">DAR GRC</a></span>',
    ),
    "family-papers.html": dict(
        nav="papers",
        path="/family-papers.html",
        ogtitle="Family Papers — The Marian Welty Family Archive",
        desc="The Marian Welty Family Archive — hand-drawn lineage charts, an early "
             "typed family genealogy, and record copies from the Welty family's own "
             "papers.",
        jsonld=False,
        note='<span class="q">More of Marian&rsquo;s collection will be added here as it is '
             'photographed and catalogued.</span>'
             '<span>Family papers shown by courtesy of the family&rsquo;s own archive; '
             'traditions preserved on these documents are noted where later research has '
             'revised them.</span>',
    ),
}

def esc(s):
    return s.replace("&", "&amp;").replace('"', "&quot;")

def block(cfg):
    url = DOMAIN + cfg["path"]
    desc = esc(cfg["desc"])
    ogt = esc(cfg["ogtitle"])
    img = DOMAIN + "/og-image.png"
    parts = [START,
        '<meta name="description" content="%s">' % desc,
        '<meta name="author" content="The Welty Project">',
        '<meta name="robots" content="index, follow">',
        '<meta name="theme-color" content="#7a1f1f">',
        '<link rel="canonical" href="%s">' % url,
        '',
        # Favicon set — matches the last config that actually rendered in Safari
        # (the v4 backup: ico + svg + png). Both the .ico AND the svg matter:
        # Safari's macOS tab renders the crest SVG (proven — it's the in-page
        # logo too); the earlier "Safari can't do svg favicons" note was wrong
        # and dropping the svg is what regressed the tab icon. Do NOT strip the
        # .ico or the .svg line. Fresh -v8 filenames defeat Safari's favicon
        # cache. Chrome/Firefox use the svg/png. (3 Jul 2026)
        '<link rel="icon" href="/favicon-v8.ico" sizes="any">',
        '<link rel="icon" type="image/svg+xml" href="/favicon-v8.svg">',
        '<link rel="icon" type="image/png" sizes="32x32" href="/favicon-32-v8.png">',
        '<link rel="icon" type="image/png" sizes="16x16" href="/favicon-16-v8.png">',
        '<link rel="apple-touch-icon" href="/apple-touch-v8.png">',
        # Manifest = site-v8.webmanifest, whose icons are OPAQUE crest tiles.
        # Backstory: the old /site.webmanifest had transparent PNG icons + start_url
        # "/", so Safari applied it to the homepage tab, failed to render the
        # transparent icon, and fell back to a short_name ("Welty" -> "W") monogram.
        # Fixed by (a) opaque crest icons so Safari renders the crest, and (b) a NEW
        # manifest URL so Safari re-reads instead of using its cached copy. Keep the
        # manifest icons opaque; don't point back to /site.webmanifest. (3 Jul 2026)
        '<link rel="manifest" href="/site-v8.webmanifest">',
        '',
        '<link rel="preconnect" href="https://fonts.googleapis.com">',
        '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>',
        '<link href="https://fonts.googleapis.com/css2?family=EB+Garamond:ital,wght@0,400;0,500;0,600;0,700;1,400&family=Cormorant+SC:wght@500;600;700&display=swap" rel="stylesheet">',
        # Shared design system — fonts, colours, page background, heading style.
        # Single source of truth so pages can\'t drift apart. See site/welty.css.
        '<link rel="stylesheet" href="/welty.css">',
        '',
        '<meta property="og:type" content="website">',
        '<meta property="og:site_name" content="The Welty Project">',
        '<meta property="og:locale" content="en_US">',
        '<meta property="og:title" content="%s">' % ogt,
        '<meta property="og:description" content="%s">' % desc,
        '<meta property="og:url" content="%s">' % url,
        '<meta property="og:image" content="%s">' % img,
        '<meta property="og:image:width" content="1200">',
        '<meta property="og:image:height" content="630">',
        '<meta property="og:image:alt" content="The Welty Project — the Edenkoben family crest">',
        '',
        '<meta name="twitter:card" content="summary_large_image">',
        '<meta name="twitter:title" content="%s">' % ogt,
        '<meta name="twitter:description" content="%s">' % desc,
        '<meta name="twitter:image" content="%s">' % img,
    ]
    if cfg["jsonld"]:
        parts += [
        '',
        '<script type="application/ld+json">',
        '{"@context":"https://schema.org","@type":"WebSite",'
        '"name":"The Welty Project",'
        '"alternateName":"The Welty Family Research Journal",'
        '"url":"%s/",' % DOMAIN +
        '"inLanguage":"en",'
        '"description":"%s"}' % cfg["desc"].replace('"', '\\"'),
        '</script>',
        ]
    parts.append(END)
    return "\n".join(parts)

def main():
    for fn, cfg in PAGES.items():
        try:
            t = open(fn, encoding="utf-8").read()
        except FileNotFoundError:
            print("skip (not found):", fn); continue
        # --- SEO/OG block in <head>, right after </title> ---
        t = re.sub(re.escape(START) + r".*?" + re.escape(END) + r"\n?", "", t, flags=re.S)
        # Strip any page-template's OWN Google-Fonts preconnect/stylesheet links so the
        # only font <link> left is the one this script injects (line below in block()).
        # This keeps every page on one canonical font request and kills the duplicate
        # <link> that the tree/papers templates carried (audit, 12 Jul 2026).
        t = re.sub(r'[ \t]*<link[^>]*fonts\.(?:googleapis|gstatic)\.com[^>]*>\n?', "", t, flags=re.I)
        m = re.search(r"</title>", t)
        if not m:
            print("no <title> in", fn); continue
        i = m.end()
        t = t[:i] + "\n" + block(cfg) + t[i:]

        # --- Cloudflare Web Analytics beacon in <head>, right after the SEO block ---
        t = re.sub(re.escape(ASTART) + r".*?" + re.escape(AEND) + r"\n?", "", t, flags=re.S)
        me = re.search(re.escape(END), t)
        if me:
            e = me.end()
            t = t[:e] + "\n" + analytics_block() + t[e:]

        # --- shared header, right after <body ...> ---
        t = re.sub(re.escape(HSTART) + r".*?" + re.escape(HEND) + r"\n?", "", t, flags=re.S)
        mb = re.search(r"<body[^>]*>", t)
        if mb:
            j = mb.end()
            t = t[:j] + "\n" + header(cfg["nav"]) + t[j:]
        else:
            print("no <body> in", fn)

        # --- shared footer, right before </body> (every page; index included) ---
        t = re.sub(re.escape(FSTART) + r".*?" + re.escape(FEND) + r"\n?", "", t, flags=re.S)
        mc = re.search(r"</body>", t)
        if mc:
            k = mc.start()
            t = t[:k] + footer(cfg) + "\n" + t[k:]

        open(fn, "w", encoding="utf-8").write(t)
        print("injected:", fn)

    # Pages outside PAGES that should still get the analytics beacon (no SEO/
    # header/footer changes) — e.g. the static 404 page.
    for fn in ("404.html",):
        try:
            t = open(fn, encoding="utf-8").read()
        except FileNotFoundError:
            print("skip (not found):", fn); continue
        t = re.sub(re.escape(ASTART) + r".*?" + re.escape(AEND) + r"\n?", "", t, flags=re.S)
        mt = re.search(r"</title>", t)
        if not mt:
            print("no <title> in", fn); continue
        i = mt.end()
        t = t[:i] + "\n" + analytics_block() + t[i:]
        open(fn, "w", encoding="utf-8").write(t)
        print("analytics injected:", fn)

if __name__ == "__main__":
    main()
