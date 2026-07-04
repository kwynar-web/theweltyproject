#!/usr/bin/env python3
"""Generate favicons + social-preview image for The Welty Project site.

Source art:  crest-edenkoben.svg  (the family's origin-town arms).
Outputs (into the Netlify Upload/ folder, next to the HTML):
    favicon.svg            – vector icon (the crest)
    favicon.ico            – 16/32/48 multi-size, for legacy browsers
    favicon-16.png, favicon-32.png
    apple-touch-icon.png   – 180×180, crest on a parchment tile (iOS)
    icon-192.png, icon-512.png – PWA / Android (manifest)
    og-image.png           – 1200×630 social share card

Re-run after changing the crest or the wording. Requires: cairosvg, Pillow.
Run from inside the "Netlify Upload" folder.
"""
import cairosvg
from PIL import Image, ImageDraw, ImageFont

CREST = "crest-edenkoben.svg"

# ---- theme colours (from index.html) ----
INK        = (28, 26, 23)
PARCH      = (245, 239, 226)
PARCH_DEEP = (233, 224, 202)
CARD       = (251, 247, 236)
GOLD       = (184, 145, 47)
GOLD_BR    = (212, 175, 55)
CRIMSON    = (122, 31, 31)
RULE       = (203, 185, 141)
MUTED      = (110, 99, 83)

LSERIF = "/usr/share/fonts/truetype/liberation/LiberationSerif-{}.ttf"
def font(style, size):
    return ImageFont.truetype(LSERIF.format(style), size)

def render_crest(px):
    """Rasterise the crest SVG to a transparent RGBA image `px` tall."""
    cairosvg.svg2png(url=CREST, write_to="/tmp/_crest.png", output_height=px)
    return Image.open("/tmp/_crest.png").convert("RGBA")

def rounded_tile(size, radius, bg):
    tile = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(tile)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=bg + (255,))
    return tile

# ---------------------------------------------------------------- favicons
def make_favicons():
    # transparent PNG icons straight from the crest
    for s in (16, 32, 48, 64, 192, 512):
        c = render_crest(s)
        # letterbox to square so the shield sits centred
        w, h = c.size
        side = max(w, h)
        canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        canvas.paste(c, ((side - w) // 2, (side - h) // 2), c)
        canvas = canvas.resize((s, s), Image.LANCZOS)
        if s in (16, 32):
            canvas.save(f"favicon-{s}.png")
        if s == 192:
            canvas.save("icon-192.png")
        if s == 512:
            canvas.save("icon-512.png")
    # multi-size .ico
    ico_src = render_crest(64)
    w, h = ico_src.size; side = max(w, h)
    sq = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    sq.paste(ico_src, ((side - w) // 2, (side - h) // 2), ico_src)
    sq.save("favicon.ico", sizes=[(16, 16), (32, 32), (48, 48)])

    # apple-touch: crest on a parchment rounded tile (iOS shows opaque bg)
    S = 180
    tile = rounded_tile(S, 40, PARCH)
    crest = render_crest(int(S * 0.74))
    cw, ch = crest.size
    tile.paste(crest, ((S - cw) // 2, (S - ch) // 2), crest)
    tile.convert("RGB").save("apple-touch-icon.png")

# ---------------------------------------------------------------- OG image
def make_og():
    W, H = 1200, 630
    img = Image.new("RGB", (W, H), PARCH)
    d = ImageDraw.Draw(img)

    # subtle vertical parchment gradient
    for y in range(H):
        t = y / H
        r = int(PARCH[0] + (PARCH_DEEP[0] - PARCH[0]) * t)
        g = int(PARCH[1] + (PARCH_DEEP[1] - PARCH[1]) * t)
        b = int(PARCH[2] + (PARCH_DEEP[2] - PARCH[2]) * t)
        d.line([(0, y), (W, y)], fill=(r, g, b))

    # double gold frame
    d.rectangle([26, 26, W - 27, H - 27], outline=GOLD, width=3)
    d.rectangle([38, 38, W - 39, H - 39], outline=GOLD_BR, width=1)

    # crest on the left
    crest = render_crest(430)
    cw, ch = crest.size
    cx, cy = 92, (H - ch) // 2
    img.paste(crest, (cx, cy), crest)

    # ---- text block on the right ----
    tx = cx + cw + 70
    # eyebrow
    eb = font("Regular", 30)
    eyebrow = "A  L I V I N G   G E N E A L O G Y   J O U R N A L"
    d.text((tx, 150), eyebrow, font=eb, fill=GOLD)

    # title (two lines, small-caps feel via uppercase + spacing)
    tf = font("Bold", 96)
    d.text((tx - 2, 196), "THE WELTY", font=tf, fill=INK)
    d.text((tx - 2, 292), "PROJECT", font=tf, fill=CRIMSON)

    # gold rule
    d.rectangle([tx, 410, tx + 300, 414], fill=GOLD_BR)

    # subtitle
    sf = font("Italic", 37)
    d.text((tx, 436), "The Edenkoben family of the Palatinate,", font=sf, fill=MUTED)
    d.text((tx, 480), "traced from c. 1680 to today.", font=sf, fill=MUTED)

    # domain footer
    df = font("Regular", 29)
    d.text((tx, 540), "theweltyproject.com", font=df, fill=GOLD)

    img.save("og-image.png")

if __name__ == "__main__":
    make_favicons()
    make_og()
    print("web assets written")
