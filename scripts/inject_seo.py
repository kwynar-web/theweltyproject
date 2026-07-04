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

# obfuscated contact address (assembled in JS so scrapers don't harvest the raw string)
CONTACT_JS = ('<script>(function(){var u="outreach",d="theweltyproject.com",'
              'a=document.getElementById("wsite-mail");if(a){a.href="mailto:"+u+"@"+d+'
              '"?subject=The%20Welty%20Project";a.textContent=u+"@"+d;}})();</script>')

# shared contact footer (index.html has its own footer + contact card, so it is skipped)
FOOTER_ON = ("all-families.html", "timeline.html", "family-papers.html")

def footer():
    return FSTART + """
<style>
.wsite-ftr{margin-top:42px;padding:22px 20px;border-top:1px solid #cbb98d;text-align:center;
  font-family:'EB Garamond',Georgia,'Times New Roman',serif;color:#6e6353;background:#faf6ec;font-size:1rem}
.wsite-ftr a{color:#7a1f1f;font-weight:600;text-decoration:none;border-bottom:1px solid #cbb98d}
.wsite-ftr a:hover{color:#1c1a17}
</style>
<footer class="wsite-ftr">Questions or comments? <a id="wsite-mail" href="#">get in touch</a>
""" + CONTACT_JS + """
</footer>
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
  font-family:'EB Garamond',Georgia,'Times New Roman',serif;box-shadow:0 1px 5px rgba(60,45,20,.09);}
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
    ),
    "all-families.html": dict(
        nav="all",
        path="/all-families.html",
        ogtitle="The Welty Family Tree — All Lines",
        desc="Explore the interactive Welty family tree: the German Edenkoben "
             "family of the Palatinate and the genuinely separate Swiss Emmental "
             "Wälti — built from primary records and Y-DNA.",
        jsonld=False,
    ),
    "timeline.html": dict(
        nav="timeline",
        path="/timeline.html",
        ogtitle="The Welty Family, Place by Place",
        desc="A place-by-place history of the Welty family — from Bischheim and "
             "Edenkoben in the Palatinate across the Atlantic to Pennsylvania, "
             "Ohio and beyond.",
        jsonld=False,
    ),
    "family-papers.html": dict(
        nav="papers",
        path="/family-papers.html",
        ogtitle="Family Papers — The Marian Welty Family Archive",
        desc="The Marian Welty Family Archive — hand-drawn lineage charts, an early "
             "typed family genealogy, and record copies from the Welty family's own "
             "papers.",
        jsonld=False,
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
        m = re.search(r"</title>", t)
        if not m:
            print("no <title> in", fn); continue
        i = m.end()
        t = t[:i] + "\n" + block(cfg) + t[i:]

        # --- shared header, right after <body ...> ---
        t = re.sub(re.escape(HSTART) + r".*?" + re.escape(HEND) + r"\n?", "", t, flags=re.S)
        mb = re.search(r"<body[^>]*>", t)
        if mb:
            j = mb.end()
            t = t[:j] + "\n" + header(cfg["nav"]) + t[j:]
        else:
            print("no <body> in", fn)

        # --- shared contact footer, right before </body> (skip index: it has its own) ---
        t = re.sub(re.escape(FSTART) + r".*?" + re.escape(FEND) + r"\n?", "", t, flags=re.S)
        if fn in FOOTER_ON:
            mc = re.search(r"</body>", t)
            if mc:
                k = mc.start()
                t = t[:k] + footer() + "\n" + t[k:]

        open(fn, "w", encoding="utf-8").write(t)
        print("injected:", fn)

if __name__ == "__main__":
    main()
