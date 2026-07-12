#!/usr/bin/env python3
"""Generate the Welty family chart from the one People Roster sheet.
Reads  'People Roster (chart source)'  in  Welty Ancestry Research Log.xlsx
Writes 'Welty Family Tree - All Families.html'  (interactive master tree, all 3 families)
Writes 'Welty Family Tree - By Generation (German).html'  (generation grid,
       Edenkoben + Manchester only, each card shows Proof record + Source)
Re-run this any time you add/edit people in the roster sheet — ONE run keeps
both live HTMLs in sync with the roster automatically.

RETIRED (kept in code, not called): the old 3-family 'By Generation' grid
(render_gen_grid) and the German-Lines graphical node-and-line chart
(render_graph) — Kwyn prefers the proof-annotated By-Generation grid (1 Jul 2026).
"""
import json, html, re, os, openpyxl

XLSX    = "Welty Ancestry Research Log.xlsx"
SHEET   = "People Roster (chart source)"
OUT_ALL = "Welty Family Tree - All Families.html"
OUT_GEN = "Welty Family Tree - By Generation.html"

# Record-image chips: built by build_records.py -> site/records/records.json.
# Maps PersonID -> [ {slug, caption, repo, url, confidence}, ... ]. Optional:
# if the file is absent (e.g. a fresh clone before build_records runs), the tree
# renders exactly as before with no chips.
_RECORDS_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "..", "site", "records", "records.json")
def _load_records():
    try:
        with open(_RECORDS_JSON, encoding="utf-8") as fh:
            return json.load(fh).get("by_person", {})
    except (FileNotFoundError, ValueError):
        return {}
RECORDS_BY_PERSON = _load_records()

COLS = ["PersonID","Family","Gen","Name","Sex","Birth","Death","Place",
        "Spouse","FatherID","Proof","DNAkit","Direct","Notes","ProofRec","Source"]

FAMILIES = [
    ("Eden",  "The German family · Edenkoben", "R1b",
     "One German family: the Wäldi/Welty household of Edenkoben (Palatinate). Immigrant brothers Philip Jacob (1750, our R1b spine), John Jacob (Weltytown line) and Georg Wolfgang — whose Manchester branch carries an I1 Y-line (a paternity break inside the family, not a separate clan) and appears under his node below.", True),
    ("Manch", "Manchester branch", "I1",
     "The Manchester branch of the German family (Dover / Manchester Twp, York Co PA) — I1 Y-line via a paternity break at/above Georg Wolfgang. Shares the names Philip Jacob / Henry / Catherine — the main source of the old confusion.", False),
    ("Swiss", "Swiss Emmental Wälti", "I2b",
     "The Swiss Emmental line — the Rüderswil / Emmental Wälti of the “Welty Family Chronicle” (Bacon 2002); source of the “Weltys are Swiss” story.", False),
    ("Md", "Maryland · Taneytown / Emmitsburg", "untested",
     "A genuinely separate, Roman-Catholic Welty family from Eppingen in Baden (Kraichgau) — <b>not</b> the Edenkoben line. The immigrant John Welty (b.1722) crossed on the ship <i>Neptune</i> in 1751, lived in York Co PA, then settled the Piney Creek / Taneytown district of Maryland and died near Emmitsburg at 94. Because Emmitsburg sits ~10 miles south of Gettysburg, this well-documented clan is the best real-world seed of the family lore's “Gettysburg / Maryland brother.” No Y-DNA sample exists for the line.", False),
    ("R1a", "Greene County, TN · John Welty", "R1a",
     "A separate, unrelated York patriline — Y-DNA <b>R1a (R-M198)</b>, which cannot share a paternal ancestor with our R1b Edenkoben line, the Manchester I1 branch, or the Swiss I2b Wälti. Its progenitor, <b>John (Nicolaus) Welty Sr.</b> (b.~1744 in the Palatinate), was a German-born Continental Army soldier — German Regiment, enlisted at Baltimore in 1778 (<b>DAR Patriot A203144</b>, pension S39882). In 1783 he married Margaret Ilgenfritz, a Dover girl and thus one of the Edenkoben family's own neighbours, at First Trinity Reformed in York — the same register and decade as our own Michael — then followed the Great Wagon Road to Shenandoah Co, Virginia and on to Greene Co, Tennessee. His son George (b.1797 VA) and grandson Peter Hughes Welty (b.1827 TN) carry the line down to a living FTDNA tester (kit #43635). He stands here as the clearest case of the wider truth: the colonial York Weltys were <i>several</i> unrelated families who lived door to door, worshipped together, and assumed a shared bloodline the DNA does not confirm.", False),
    ("Yrk", "York, PA · George Welty", "R1b",
     "A York County Welty line carried by FTDNA kit <b>#19175</b> — a <i>different</i> R1b subclade (R-L151) from our Edenkoben cluster, with no common ancestor in a genealogical timeframe. Just the two Pennsylvania members are shown — <b>George Welty</b> (b.~1797, Conewago) and his son <b>John</b> (b.1827, York Co) — before the family left for Ohio and Michigan. Above George is a brick wall: the record that would name his parents is John's baptism in Quickel's Conewago register, still access-locked. One of the several York Welty families that get tangled together in the county records.", False),
    ("Gva", "Goochland Co, VA · William Weldy", "R1b (R-DF49)",
     "A Virginia Weldy line — Y-DNA <b>R1b (R-DF49)</b> — of 18th-century Goochland County, sharing no ancestor with the Palatinate Edenkoben or any other Welty line in a genealogical timeframe. Its documented anchor is <b>George Weldy</b> (b.~1752, d.1821 Goochland), a Revolutionary War soldier of the Virginia line (pension S39883) whose 1821 will survives. Two FTDNA participants — kits <b>SMGF#1</b> and <b>#367336</b> — both trace on paper to George through two of his sons, Henry (b.1790) and William (b.1791), yet their Y-DNA diverges by a genetic distance of about 8 at 25 markers — far more than true brothers' paternal lines could. A <b>non-paternity event</b> therefore sits in one of the two branches: at most one is George Weldy's biological male line. Shown as one documented Goochland family with that discrepancy flagged, not silently merged.", False),
    ("San", "Saanen → Upper Sandusky · Wälti (hg G)", "G2a3b",
     "A wholly separate Wälti clan — Y-DNA <b>haplogroup G (G2a3b)</b> — with <b>no</b> kinship to the R1b Edenkoben or Goochland lines, the I1 Manchester branch, or the I2b Swiss-Emmental Wälti; haplogroup G split from all of them in deep prehistory, an independent adoption of the Wälti surname. Traced by FTDNA kit <b>#161609</b> (Big Y) up an unbroken Simmental parish pedigree to <b>Anthoni Wälti</b> (christened 1568, Lenk im Simmental, Bern). The American immigrant is <b>Johann Gottlieb Wälti</b> (b.1861 Saanen), who settled at Upper Sandusky, Wyandot Co, Ohio. The deep Swiss rungs are graded as compiled parish lore; Gottlieb's American children await a census pass.", False),
]

# ============================ CITATION FORMATTER ============================
# Turns the roster's terse internal source shorthand into reference-quality,
# outward-facing citations: strips the private lead codes ([P40], [M4], …),
# expands abbreviations, and linkifies stable repository IDs so an outside
# researcher can click straight through to the record.
def _fmt_link(url, text):
    return f'<a href="{url}" target="_blank" rel="noopener">{text}</a>'

def _fmt_segment(seg):
    # 1) drop internal lead codes like [P40], [M10/M18], [FB3/FB9], [SR1-SR3/FB21]
    seg = re.sub(r'\s*\[[^\]]*\]', '', seg).strip(' ;,')
    # 1a) drop internal event references ("ev.7965", "(ev.7965)") — the roster's
    #     private event-row pointers that were leaking into Proof/Source lines.
    seg = re.sub(r'\(\s*ev\.?\s*\d+\s*\)', '', seg)
    seg = re.sub(r'[;,]?\s*\bev\.?\s*\d+\b', '', seg)
    # 1b) also drop PARENTHETICAL lead codes that leak into proof/source cells:
    #     "(P75)", "(P75, H/W caveat)", "(D7)". Only the parenthetical form is
    #     stripped so real citations that look code-like (e.g. census microfilm
    #     "M637", "Vol I p.433") are left untouched.
    _codes = r'(?:P|M|D|FB|FA|SR|US|H|FT|PS|DL|TL)-?\d+[a-z]?'
    seg = re.sub(r'\(\s*' + _codes + r'(?:\s*[,;]\s*)?', '(', seg)  # "(P75, x" -> "(x" ; "(P75)" -> "()"
    seg = re.sub(r'\(\s*\)', '', seg)                              # empty parens
    seg = re.sub(r'\s+\)', ')', seg)
    # normalize internal DNA-line label "SMGF#2 line" -> readable (keeps the real
    # sample citation SMGF-2152399, which uses a hyphen and is left untouched)
    seg = re.sub(r'\bSMGF#\d+\s*line\b', 'SMGF-tested line', seg)
    seg = re.sub(r'\bSMGF#\d+', 'SMGF', seg)
    seg = re.sub(r'\s{2,}', ' ', seg).strip(' ;,')
    if not seg:
        return ''
    # 2) normalise a few known published sources to full citations (plain text)
    seg = seg.replace('Strassburger/Hinke',
                      'Strassburger & Hinke, Pennsylvania German Pioneers (1934),')
    seg = re.sub(r'Biog\.\s*Memoirs of Wyandot Co(?:\s+OH)?\s*1902',
                 'Biographical Memoirs of Wyandot County, Ohio (1902)', seg)
    seg = re.sub(r'\bBowen 1902\b',
                 'Bowen, Biographical Memoirs of Wyandot County, Ohio (1902)', seg)
    seg = re.sub(r'\bBeers 1899\b', 'Beers, Biographical Record (1899)', seg)
    seg = seg.replace('held by Kwyn', 'held by the compiler')
    seg = re.sub(r'\b[Pp]er Kwyn\b', "per compiler's records", seg)

    # 3) linkify stable repository IDs; stash each anchor behind a placeholder so
    #    later HTML-escaping and abbreviation expansion cannot corrupt the markup.
    links = []
    def stash(frag):
        links.append(frag)
        return '\x00%d\x00' % (len(links) - 1)

    seg = re.sub(r'\b(?:FindAGrave|FaG)\s*#\s*(\d+)',
        lambda m: stash(_fmt_link(f'https://www.findagrave.com/memorial/{m.group(1)}',
                                  f'FindAGrave&nbsp;#{m.group(1)}')), seg)
    seg = re.sub(r'\b(?:FindAGrave|FaG)\s+cem\.?\s*(\d+)',
        lambda m: stash(_fmt_link(f'https://www.findagrave.com/cemetery/{m.group(1)}',
                                  f'FindAGrave cemetery&nbsp;{m.group(1)}')), seg)
    seg = re.sub(r'\b(?:Ancestry\s+)?coll\.?\s*(\d+)\s*ev\.?\s*(\d+)',
        lambda m: stash(_fmt_link(
            f'https://www.ancestry.com/discoveryui-content/view/{m.group(2)}:{m.group(1)}',
            f'Ancestry collection&nbsp;{m.group(1)} (record&nbsp;{m.group(2)})')), seg)
    seg = re.sub(r'\b(?:Ancestry\s+)?coll\.?\s*(\d+)',
        lambda m: stash(_fmt_link(f'https://www.ancestry.com/search/collections/{m.group(1)}/',
                                  f'Ancestry collection&nbsp;{m.group(1)}')), seg)
    seg = re.sub(r'\bFS ark\s*([0-9A-Za-z-]+)',
        lambda m: stash(_fmt_link(f'https://www.familysearch.org/ark:/61903/3:1:{m.group(1)}',
                                  f'FamilySearch image (ark&nbsp;{m.group(1)})')), seg)
    seg = re.sub(r'\bFS-DL\s*(\d+)',
        lambda m: stash(_fmt_link(
            f'https://www.familysearch.org/library/books/records/item/{m.group(1)}',
            f'FamilySearch Digital Library item&nbsp;{m.group(1)}')), seg)
    seg = re.sub(r'\bfilm\s*(\d{5,})',
        lambda m: stash('FamilySearch film ' +
                        _fmt_link(f'https://www.familysearch.org/search/film/{m.group(1)}',
                                  m.group(1))), seg)
    seg = re.sub(r'\bWikiTree\s+([A-Za-z]+-\d+)',
        lambda m: stash(_fmt_link(f'https://www.wikitree.com/wiki/{m.group(1)}',
                                  f'WikiTree {m.group(1)}')), seg)
    seg = re.sub(r'\bGAMEO\b',
        lambda m: stash(_fmt_link('https://gameo.org',
                                  'GAMEO (Global Anabaptist Mennonite Encyclopedia Online)')), seg)
    # archive.org item IDs: capture the FULL identifier (may contain _ . - and
    # upper-case, e.g. 'bub_gb_5doyAQAAMAAJ') — the old [a-z0-9]+ truncated at the
    # first underscore, producing dead links like /details/bub. Only linkify tokens
    # that look like real IA identifiers (contain a digit); this leaves descriptive
    # prose such as 'archive.org viewer' as plain text instead of a bogus link.
    def _ia_link(m):
        tok = m.group(1).rstrip('.,);:')
        if not re.search(r'\d', tok):          # 'viewer', 'stream' etc. -> not an ID
            return m.group(0)
        return stash('archive.org ' +
                     _fmt_link(f'https://archive.org/details/{tok}', tok))
    seg = re.sub(r'\barchive\.org\s+([A-Za-z0-9][A-Za-z0-9_.-]*)', _ia_link, seg)
    seg = re.sub(r'\bbrian-hamman\.com/(\S+)',
        lambda m: stash(_fmt_link(f'https://brian-hamman.com/{m.group(1)}',
                                  'brian-hamman.com/' + m.group(1))), seg)
    seg = re.sub(r'\bbrian-hamman\.com\b',
        lambda m: stash(_fmt_link('https://brian-hamman.com', 'brian-hamman.com')), seg)
    seg = re.sub(r'\barchive\.org\b',
        lambda m: stash(_fmt_link('https://archive.org', 'archive.org')), seg)
    seg = re.sub(r'\bArchion\b',
        lambda m: stash(_fmt_link('https://www.archion.de', 'Archion')), seg)

    # 4) HTML-escape the remaining plain text
    seg = html.escape(seg, quote=True)

    # 5) expand abbreviations (links are safely stashed as placeholders)
    seg = re.sub(r'\bFS Full[- ]?Text\b', 'FamilySearch Full-Text Search', seg, flags=re.I)
    seg = re.sub(r'\bFS\b', 'FamilySearch', seg)
    seg = re.sub(r'\bFaG\b', 'FindAGrave', seg)
    seg = re.sub(r'\bAdmin Rec\b', 'Administration Records', seg)
    seg = re.sub(r'\breg\b', 'register', seg)
    seg = re.sub(r'\bimgs\b', 'images', seg)
    seg = re.sub(r'\bimg\b', 'image', seg)
    seg = re.sub(r'\btranscr\.', 'transcription', seg)
    seg = re.sub(r'\bobit\b', 'obituary', seg)
    seg = re.sub(r'\bcert\b', 'certificate', seg)
    seg = re.sub(r'\bCem\b', 'Cemetery', seg)
    seg = re.sub(r'\(ref\.\)', '(Reformed)', seg)

    # 6) restore the stashed anchors
    seg = re.sub('\x00(\\d+)\x00', lambda m: links[int(m.group(1))], seg)
    return seg

def format_source_html(raw):
    """Public-facing, reference-quality HTML rendering of a roster Source/Proof cell."""
    if not raw:
        return ''
    parts = [_fmt_segment(s) for s in str(raw).split(';')]
    return '; '.join(p for p in parts if p)

# Strip private lead codes (P60, TL-34, FB15 …) that leak into the compact display
# fields (Birth/Death/Place/Spouse) while preserving the surrounding meaningful text.
_DISPLAY_CODE = r'(?:P|M|D|FB|FA|SR|US|H|FT|PS|DL|TL)-?\d+[a-z]?'
def scrub_display(s):
    if not s:
        return s
    s = str(s)
    s = re.sub(r'\s*\[[^\]]*\]', '', s)                       # [..] codes
    s = re.sub(r',?\s*\bsee\s+' + _DISPLAY_CODE + r'\b', '', s)  # ", see FB39"
    s = re.sub(r'[;,]?\s*\b' + _DISPLAY_CODE + r'\b', '', s)   # standalone codes
    s = re.sub(r'\(\s*[;:,]\s*', '(', s)
    s = re.sub(r'[;:,]\s*\)', ')', s)
    s = re.sub(r'\(\s*\)', '', s)
    s = re.sub(r'\s+([;,.)])', r'\1', s)
    s = re.sub(r'\(\s+', '(', s)
    s = re.sub(r'\s{2,}', ' ', s)
    return s.strip().strip(' ,;')


# Public-facing scrub for the free-text Notes field. The Notes carry the internal
# research narrative (lead codes, login handles, working shorthand); this strips
# that before the notes are published in the interactive tree, WITHOUT touching the
# working roster. (Added 4 Jul 2026 — the display fields were already scrubbed, but
# Notes was published raw, leaking TL-/M/P codes, "PLACED", login handles, etc.)
def scrub_notes_public(s):
    if not s:
        return s
    s = scrub_display(s)                              # [..] + standalone lead codes
    s = re.sub(r'\s*\|\|\s*', ' — ', s)               # internal "||" separators
    s = re.sub(r'\bPLACED\b[:\s]*', '', s)            # "PLACED 1 Jul 2026:" shorthand
    s = re.sub(r'\bKwyn\d+\b', '', s)                 # login handles (Kwyn28)
    s = re.sub(r'\bSMGF#\d+\b', 'SMGF', s)            # internal "SMGF#2 line" label
    s = re.sub(r'\bcoll\.?\s*(\d+)', r'collection \1', s)  # "coll 4940" -> "collection 4940"
    s = re.sub(r'\bDEATH FORK\b\s*', '', s)            # internal all-caps working label
    s = re.sub(r'\s*/\s*\)', ')', s)                  # leftover slash from stripped "P60/M73"
    s = re.sub(r'\s{2,}', ' ', s)
    s = re.sub(r'\s+([;,.)])', r'\1', s)
    return s.strip().strip(' ,;—')


def pub_dna_label(s):
    """Normalize the public DNA badge: internal 'SMGF#2 line' -> readable, keep the
    real sample id SMGF-2152399 (hyphen form) untouched."""
    if not s:
        return s
    s = re.sub(r'\bSMGF#\d+\s*line\b', 'SMGF-tested line', s)
    s = re.sub(r'\bSMGF#\d+', 'SMGF', s)
    return s


# ============================ PUBLIC NOTES ORGANIZER ============================
# The roster Notes field is the internal research DIARY: it carries family
# codenames (Family A/B), roster node-IDs (B-george, E-...), research-session
# dates ("2 Jul 2026"), and process-status narration (UPGRADED, MERGED, "the
# compiler decision", "own-eyes read"). scrub_notes_public() (above) only removed
# lead codes, so the tree still published that diary voice as a wall of text.
# This organizer breaks the note at its real diary joints into a few clean,
# reader-facing paragraphs and strips the internal artifacts — while KEEPING every
# record citation (reg/img/ark/kit/coll/WikiTree/FindAGrave), which notes_to_html()
# then renders as clean links. Runs at publish time only; the roster is untouched.
_N_MONTH = r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?'
_N_DATE  = r'\d{1,2}\s+' + _N_MONTH + r'\s+20\d\d'      # a research-session date
_N_NODE  = r'[BEA]-[a-z][A-Za-z0-9]*'                   # roster PersonID token
_N_CAPS  = (r'UPGRADED|DEMOTED|MERGED|RESOLVED|PLACED|CONFIRMATION\s+FOUND|CONFIRMED|'
            r'CHILD-?LIST\s+VALIDATED|VALIDATED|IMAGE-?VERIF(?:Y|IED)|IMAGE-?READ|'
            r'RE-?VERIFIED|SHIP|SWEEP|DEATH\s+FORK')
_N_LEADS = r'Reg\s+\d|Marriage\s+found|Descendant\s+line|Loc\s+tentative'
_N_JOINT = re.compile(r'\s+—\s+(?=\(?\s*(?:' + _N_DATE + r'|' + _N_CAPS + r'|' + _N_LEADS + r')\b)')
_N_HEADER = re.compile(
    r'^\s*(?:'
    r'(?:' + _N_CAPS + r')\b[^:.—]{0,80}?:'
    r'|(?:' + _N_CAPS + r')\b[^.—]{0,40}?[→>][^.—]{0,20}?(?=\s|$)'
    r'|\(?\s*' + _N_DATE + r'[^):.—]*\)?\s*[:,]?'
    r')\s*', re.I)

def _n_deartifact(s):
    # fix hyphen artifacts left by earlier lead-code stripping ("with-Philip")
    s = re.sub(r'\b(with|is|to|as|vs|of|by|for|than|and|the|our|his|her)-(?=[A-Z])', r'\1 ', s)
    s = re.sub(r'\b(the|not|now|a|an|his|her|our|their)-(?=[a-z])', r'\1 ', s)
    return s

def _n_strip_internal(s):
    s = str(s)
    s = re.sub(r'\(\s*Family\s*[AB]\s+head\s*\)', '', s, flags=re.I)   # "(Family B head)"
    s = re.sub(r'\bFamily\s*B\b', 'the Manchester branch', s)
    s = re.sub(r'\bFamily\s*A\b', 'the Swiss line', s)
    s = re.sub(r'\bFamily\s*C\b', '', s)
    s = re.sub(r'[,;]?\s*\(?\s*(?:see|cf\.?|per|retired dup|dup|under|vs\.?)\s+' + _N_NODE + r'\)?', '', s, flags=re.I)
    s = re.sub(r'\(\s*' + _N_NODE + r'\s*\)', '', s)
    s = re.sub(r'\b' + _N_NODE + r'\b', '', s)
    s = re.sub(r'\(\s*the compiler decision\s*\)', '', s, flags=re.I)
    s = re.sub(r'\bthe compiler decision\b', '', s, flags=re.I)
    s = re.sub(r'\(\s*(?:our|my) reading\s*\)', '', s, flags=re.I)
    s = re.sub(r'\bown-eyes read\b', '', s, flags=re.I)
    s = re.sub(r'\b(?:the compiler|Claude/Chrome)[^),.]*?(?:Archion|FamilySearch|FS)\s+(?:login|sign-?in)\b', '', s, flags=re.I)
    s = re.sub(r'\bthis (?:row|node) is now the PRIMARY identity[;,.]?', '', s, flags=re.I)
    s = re.sub(r'[;,]?\s*demoted to tentative-duplicate\b', '', s, flags=re.I)
    s = re.sub(r'\bCAVEAT UNCHANGED\b\s*:?', 'Caveat:', s)
    s = re.sub(r'\bKwyn\d+\b', '', s)
    s = re.sub(r'\(\s*' + _N_DATE + r'[^)]*\)', '', s)   # "(2 Jul 2026, …)"
    s = re.sub(r'\b' + _N_DATE + r'\b', '', s)
    s = re.sub(r'\b20\d\d\b', '', s)                     # leftover bare "2026"
    s = _n_deartifact(s)
    s = re.sub(r'\bour\b', 'the', s)                     # soften internal first person
    return s

def _n_tidy(seg):
    seg = re.sub(r'\(\s*[/=;:,.]*\s*\)', '', seg)
    seg = re.sub(r'\[\s*\]', '', seg)
    seg = re.sub(r'\s{2,}', ' ', seg)
    seg = re.sub(r'\bSee\.\s*', '', seg)
    seg = re.sub(r'\s*\.(\s*\.)+', '.', seg)
    seg = re.sub(r'\(\s*;?\s*', '(', seg)
    seg = re.sub(r'[,;:]\s*\)', ')', seg)      # orphaned comma/semicolon before ) left when a stripped date was the paren's tail
    seg = re.sub(r'\(\s*[,;:]\s*', '(', seg)   # ...or the paren's head
    seg = re.sub(r'\s+([;,.)])', r'\1', seg)
    seg = re.sub(r'^[\s;,:.—>-]+', '', seg)
    seg = re.sub(r'[\s;,:—>-]+$', '', seg)
    return seg.strip()

def clean_note_segments(raw):
    """Split one roster Notes cell into cleaned, reader-facing paragraphs."""
    if not raw:
        return []
    s = re.sub(r'\s*\|\|\s*', ' — ', str(raw))     # normalize internal joints
    s = scrub_display(s)                            # strip lead codes (P75/FB19/TL-34/"see FB19")
    # Yearless working-addendum stamps ("— 4 Jul :", "— 4 Jul (addendum):") are
    # internal session markers. The year-form dates are handled by _N_HEADER /
    # _n_strip_internal, but these carry no year and slipped through. Convert to a
    # sentence break + capitalize the next word so the addendum publishes clean.
    s = re.sub(r'[\s.]*—\s*\d{1,2}\s+' + _N_MONTH + r'(?:\s*\([^)]*\))?\s*:\s*(\w)',
               lambda m: '. ' + m.group(1).upper(), s)
    s = re.sub(r'\bResearch placeholder\.?\s*', '', s, flags=re.I)  # internal status label
    out = []
    for chunk in _N_JOINT.split(s):                # split on RAW (markers intact)
        chunk = _N_HEADER.sub('', chunk.strip())   # strip leading dated/caps header
        chunk = _n_strip_internal(chunk)           # remove residual internal cruft
        chunk = _n_tidy(chunk)
        if len(re.sub(r'[^A-Za-z0-9]', '', chunk)) < 3:
            continue
        if chunk[-1] not in '.!?':
            chunk += '.'
        out.append(chunk)
    return out

def _n_linkify_ark(seg):
    # bare "ark 3:1:3QHV-V382" -> FamilySearch record link (the "FS ark" form is
    # handled inside _fmt_segment; this catches the un-prefixed form in notes)
    return re.sub(r'\bark\s+(\d:\d:[0-9A-Za-z-]+)',
        lambda m: f'<a href="https://www.familysearch.org/ark:/61903/{m.group(1)}"'
                  f' target="_blank" rel="noopener">FamilySearch record</a>', seg)

def notes_to_plain(raw):
    """Cleaned notes as plain text (used for the search index)."""
    return ' '.join(clean_note_segments(raw))

def notes_to_html(raw):
    """Cleaned notes as organized HTML paragraphs with linked citations."""
    segs = clean_note_segments(raw)
    return '<br>'.join('&bull;&nbsp;' + _n_linkify_ark(_fmt_segment(s)) for s in segs)


# ---------------------------- LIVING-PERSON PRIVACY ----------------------------
# Genealogy standard: never publish the given name, dates, places, spouse, notes,
# or DNA-kit numbers of a living individual. We keep only the birth SURNAME plus a
# "Living" tag so the tree structure still reads. A person is treated as living if
# the roster marks them "living" OR they were born within the last 100 years with
# no recorded death. This runs at publish time only — the working roster keeps the
# full private detail; re-running the generator re-applies the redaction every time.
CURRENT_YEAR = 2026
# People with no usable birth date who are nonetheless living — living DNA-test
# participants whose kits are actively compared in the project. The "born <100yr"
# rule can't catch them, so they are listed explicitly. (Standard: an unconfirmed-
# death DNA tester is treated as living.) Remove an ID here if proven deceased.
KNOWN_LIVING = set()   # (was {"B-merleSMGF"}) Merle William Welty PROVEN DECEASED
# 12 Jul 2026: Merle William Welty is publicly memorialized 1925-2012 (FindAGrave
# #7483539 child-link; named as Miller Wayne's son in the Troy Daily News obit,
# 2 Mar 1978) — de-redacted per this block's own "remove if proven deceased" rule.

# ---------------------------- UNPROVEN PARENT LINK ----------------------------
# Nodes whose descent from the parent above is NOT yet proven. The person is
# solidly documented in their own right, but the *link* to their father rests on
# indirect evidence (the Y-DNA pedigree) rather than a direct record. These render
# with a dashed incoming edge + a "Father link unproven" badge + a one-line note,
# so the tree tells its own crux story instead of painting the rung solid green.
# The node's own Proof badge is left untouched (the person stays "proven"); only
# the edge to the parent is flagged. (6 Jul 2026)
#   Policy (12 Jul 2026): ALL of Dover Philip Jacob's children render with an
#   unproven parent-link (dashed) until a PRIMARY record names the father. Each
#   child is solidly documented in their own right, but none of their descents
#   from Philip Jacob rests on a direct record yet — Michael's on the Y-DNA
#   pedigree + an unlocated ~1757 baptism; the others (Jacob Sr., John of Dover,
#   and the three Trinity-York daughters) on circumstantial/naming evidence. Add
#   any newly-found child of Philip Jacob here until their parentage is proven.
LINK_UNPROVEN = {"E-michael", "E-jacobsr", "E-johndover",
                 "E-elizabethgauf", "E-christinamesserle", "E-catharinaboehm"}

def _year_in(s):
    m = re.search(r'\b(1[5-9]\d\d|20\d\d)\b', s or '')
    return int(m.group(1)) if m else None

def is_living(p):
    if p.get("PersonID") in KNOWN_LIVING:
        return True
    if 'living' in (p.get('Proof') or '').lower():
        return True
    if (p.get('Death') or '').strip():
        return False
    by = _year_in(p.get('Birth'))
    return by is not None and by > CURRENT_YEAR - 100

def _surname(name):
    # birth surname in parens wins: "Carol Lynn (Welty) Fox" -> "Welty"
    m = re.search(r'\(([^)]+)\)', name or '')
    if m:
        return m.group(1).strip()
    toks = [t for t in (name or '').split() if t]
    return toks[-1] if toks else 'Welty'

def privatize(p):
    """Redact a living person's record in place, keeping only surname + Living tag."""
    p['Name']    = 'Living ' + _surname(p['Name'])
    p['Birth']   = ''
    p['Death']   = ''
    p['Place']   = ''
    p['Spouse']  = ''
    p['Notes']   = ''
    p['DNAkit']  = ''
    p['ProofRec'] = ''
    p['Source']  = ''
    p['Proof']   = 'living'
    return p

# Living people are also *named* inside other (deceased) relatives' free-text
# Notes — e.g. "kit held today by living descendant Aaron Welty". Scrub those
# references too, mapping the compiler's own name to "the compiler" and other
# distinctive living given names to a neutral role. (Ancestor names that collide
# — Philip, Michael — are deliberately NOT in this list.)
def scrub_living_notes(s):
    if not s:
        return s
    s = re.sub(r'\bliving descendant\s+Aaron\s+Welty\b', 'a living descendant', s)
    s = re.sub(r'\bAaron\s+Welty\b', 'a living descendant', s)
    s = re.sub(r"\bKwyn(?:'s|’s)?\b", 'the compiler', s)
    # living DNA testers named in evidence text. The negative lookahead protects
    # "Merle Ilgenfritz" (a FindAGrave source — a *different*, non-target person)
    # while still catching the tester "Merle [William] [Welty]" and bare "Merle".
    s = re.sub(r'\bMerle(?:\s+William)?(?:\s+Welty)?\b(?!\s+Ilgenfritz)',
               'a living relative', s, flags=re.I)
    s = re.sub(r'\bSteven(?:\s+Welty)?\b', 'a living relative', s, flags=re.I)
    # living married-in spouse (Alan Sprow's widow, m.1960, no death recorded)
    s = re.sub(r'\bJoy(?:\s+Ann)?\s+Bindbeutel\b', 'a living spouse', s)
    s = re.sub(r'\bBindbeutel\b', 'a living spouse', s)
    s = re.sub(r'\bJoy\b', 'a living spouse', s)
    s = re.sub(r'\b(Kris|Kurt|Beth|Sheri|Lori|Tyler|Sheila|Carol(?:\s+Lynn)?|Philip\s+Karl)\b',
               'a living relative', s)
    s = re.sub(r'\s{2,}', ' ', s)
    return s.strip()


def read_people():
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb[SHEET]
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    people = []
    for r in rows:
        if r is None:
            continue
        rec = {COLS[i]: ("" if i >= len(r) or r[i] is None else str(r[i]).strip())
               for i in range(len(COLS))}
        if not rec["PersonID"]:
            continue
        # scrub private lead codes out of the compact display fields
        for _f in ("Birth", "Death", "Place", "Spouse"):
            rec[_f] = scrub_display(rec[_f])
        people.append(rec)
    # Generation is COMPUTED from FatherID depth, not read from the typed "Gen"
    # column: the oldest known ancestor of each line (no/absent FatherID) is Gen 1,
    # each child one deeper. Push the lineage back = add a row with its FatherID and
    # re-run; every descendant renumbers automatically, no hand-editing. (3 Jul 2026)
    _fathers = {p["PersonID"]: p["FatherID"] for p in people}
    def _depth(pid):
        seen, d, cur = set(), 1, pid
        while True:
            f = _fathers.get(cur, "")
            if not f or f not in _fathers or f in seen or d > 50:
                return d
            seen.add(f); d += 1; cur = f
    for p in people:
        p["GenNum"] = _depth(p["PersonID"])
    return people

def main():
    people = read_people()

    # ---------- LIVING-PERSON PRIVACY, pass 1: anonymise IDs ----------
    # PersonIDs like "E-kris"/"E-kwyn" embed a living person's given name and would
    # leak into the published page source (JSON keys, kids arrays). Remap the living
    # people's IDs to neutral tokens BEFORE the tree is built, rewriting FatherID too
    # so parent/child links survive. (pass 2 = privatize() blanks their detail below.)
    living_recs = [p for p in people if is_living(p)]
    _idmap = {p["PersonID"]: "LV%02d" % (i + 1) for i, p in enumerate(living_recs)}
    for p in people:
        if p["PersonID"] in _idmap:
            p["PersonID"] = _idmap[p["PersonID"]]
        if p["FatherID"] in _idmap:
            p["FatherID"] = _idmap[p["FatherID"]]

    by_id = {p["PersonID"]: p for p in people}
    # children map
    for p in people:
        p["children"] = []
    roots_by_fam = {f[0]: [] for f in FAMILIES}
    for p in people:
        fid = p["FatherID"]
        if fid and fid in by_id:
            by_id[fid]["children"].append(p["PersonID"])
        else:
            roots_by_fam.setdefault(p["Family"], []).append(p["PersonID"])

    # sort children by generation, then AGE ORDER (oldest first) where a birth
    # date/year is known; undated children keep roster order and go last.
    # Handles the roster's messy birth strings: "~1757", "bp. 2 Feb 1713",
    # "before 1755 (M73)", "FT: 1786", "after 1940", "1969/70", "Sep 1853", etc.
    _MONTHS = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
               "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
    _roster_idx = {p["PersonID"]: i for i, p in enumerate(people)}
    def _birth_sort(pid):
        import re as _re
        p = by_id[pid]
        s = (p["Birth"] or "").lower()
        m = _re.search(r"\b(1[4-9]\d\d|20\d\d)\b", s)
        if not m:
            return (9999.0, 0, 0)
        year = float(m.group(1))
        pre = s[:m.start()]
        if "before" in pre or "<" in pre:
            year -= 0.4          # "before 1755" sorts ahead of exact 1755
        elif "after" in pre:
            year += 0.4          # "after 1940" sorts behind exact 1940
        month = day = 0
        mm = _re.search(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*", s)
        if mm and mm.start() < m.start():   # month belongs to the birth date, not a note
            month = _MONTHS[mm.group(1)]
            dm = _re.search(r"\b([0-3]?\d)\s+" + mm.group(1), s)
            if dm:
                day = int(dm.group(1))
        return (year, month, day)
    def sort_key(pid):
        p = by_id[pid]
        return (p["GenNum"] if p["GenNum"] is not None else 999,
                _birth_sort(pid), _roster_idx.get(pid, 9999))
    for p in people:
        p["children"].sort(key=sort_key)
    for fam in roots_by_fam:
        roots_by_fam[fam].sort(key=sort_key)

    # ---------- Manchester branch placement (merged 4 Jul 2026) ----------
    # The Manchester (I1) branch hangs directly off Georg Wolfgang
    # (E-georgwolfgang-disp), his documented father-node in the Wäldi household.
    # The old duplicate root B-george was merged into that node in the roster, so
    # no render-time graft is needed — the children point straight at him.

    # Fold the Manchester branch into the Edenkoben (German) family for display:
    # the public tree shows just two families — Edenkoben and Swiss. The Manchester
    # people keep their I1 evidence in their own notes/DNA fields; they simply live
    # under the German family now (the paternity break sits at Georg Wolfgang, noted
    # on his node). This also retires the old cross-family "copy" badge entirely.
    for p in people:
        if p["Family"] == "Manch":
            p["Family"] = "Eden"

    def person_dict(p, kids):
        return {
            "id": p["PersonID"], "fam": p["Family"], "gen": p["GenNum"],
            "name": p["Name"], "birth": p["Birth"], "death": p["Death"],
            "place": p["Place"], "spouse": p["Spouse"], "proof": p["Proof"],
            "dna": pub_dna_label(p["DNAkit"]), "direct": p["Direct"],
            "linkunproven": (p["PersonID"] in LINK_UNPROVEN),
            "notes": notes_to_plain(p["Notes"]),          # plain text -> search index
            "notes_html": notes_to_html(p["Notes"]),      # organized, linked -> display
            "proof_html": format_source_html(p["ProofRec"]),
            "source_html": format_source_html(p["Source"]),
            # record-image chips, but never for living people (privacy)
            "records": ([] if str(p.get("Proof","")).lower()=="living"
                        else RECORDS_BY_PERSON.get(p["PersonID"], [])),
            "kids": kids,
        }

    # ---------- LIVING-PERSON PRIVACY, pass 2: blank detail + scrub mentions ----------
    for p in living_recs:
        privatize(p)                       # own name/dates/etc. -> "Living <Surname>"
    for p in people:                       # scrub living names out of EVERYONE's free text
        for _f in ("Notes", "ProofRec", "Source", "DNAkit", "Spouse", "Place"):
            p[_f] = scrub_living_notes(p[_f])
    print("Redacted %d living people: %s" % (len(living_recs),
          ", ".join(_idmap.keys())))

    counts = {f[0]: sum(1 for p in people if p["Family"] == f[0]) for f in FAMILIES}
    total = len(people)

    # ---------- 1) ALL-FAMILIES payload (everyone, full tree) ----------
    payload_all = {
        "families": [
            {"key": k, "title": t, "hap": h, "desc": d, "ours": ours,
             "roots": roots_by_fam.get(k, [])}
            for (k, t, h, d, ours) in FAMILIES
            if k != "Manch"   # one German family: the Manchester branch renders
                              # once, inside the Eden tree under Georg Wolfgang
        ],
        "people": {p["PersonID"]: person_dict(p, p["children"]) for p in people},
    }
    fam_controls = (
        '<div class="grp">'
        '<label class="chk eden"><input type="checkbox" data-fam="Eden" checked> Edenkoben</label>'
        '<label class="chk swiss"><input type="checkbox" data-fam="Swiss" checked> Swiss</label>'
        '<label class="chk md"><input type="checkbox" data-fam="Md" checked> Maryland</label>'
        '<label class="chk r1a"><input type="checkbox" data-fam="R1a" checked> Greene Co</label>'
        '<label class="chk yrk"><input type="checkbox" data-fam="Yrk" checked> York</label>'
        '<label class="chk gva"><input type="checkbox" data-fam="Gva" checked> Goochland</label>'
        '<label class="chk san"><input type="checkbox" data-fam="San" checked> Saanen</label>'
        '</div>')
    german = total - counts.get('Swiss', 0) - counts.get('Md', 0) - counts.get('R1a', 0) - counts.get('Yrk', 0) - counts.get('Gva', 0) - counts.get('San', 0)
    render(OUT_ALL, payload_all,
           h1="The Welty Families &mdash; interactive tree",
           sub=("Every documented Welty, in one place: the <b>Edenkoben</b> German family &mdash; "
                "the Wäldi household of the Palatinate, including the branch whose I1 Y-line marks a "
                "paternity break inside the family (shown under Georg Wolfgang) &mdash; and the "
                "genuinely separate <b>Swiss Emmental</b> family, and the separate Roman-Catholic "
                "<b>Maryland</b> family of Taneytown / Emmitsburg (the real-world seed of the "
                "“Gettysburg brother” lore). Click a &#9656; to expand a person's "
                "children; use the search and filters to find anyone. Built automatically from the "
                f"<b>People Roster</b> sheet of the research log. <b>{total}</b> people tracked so far."),
           fam_controls=fam_controls,
           count_label=(f"{total} people · Edenkoben (German) family {german} · "
                        f"Swiss {counts.get('Swiss',0)} · Maryland {counts.get('Md',0)} · "
                        f"Greene Co {counts.get('R1a',0)} · York {counts.get('Yrk',0)} · "
                        f"Goochland {counts.get('Gva',0)} · Saanen {counts.get('San',0)}"))
    print(f"wrote {OUT_ALL}  ({total} people: Edenkoben {german}, Swiss {counts.get('Swiss',0)}, Maryland {counts.get('Md',0)}, Greene Co {counts.get('R1a',0)}, York-George {counts.get('Yrk',0)}, Goochland {counts.get('Gva',0)}, Saanen {counts.get('San',0)})")

    # ---------- 1b) GERMAN-LINES graphical chart — RETIRED 1 Jul 2026 (Kwyn prefers the
    # By-Generation grid). render_graph()/GRAPH_TEMPLATE kept below but no longer called.
    german_fams = ("Eden", "Manch")

    # ---------- 1c) GERMAN-LINES by-generation grid (with proof record + source per person) ----------
    OUT_GEN_DE = "Welty Family Tree - By Generation (German).html"
    gpeople = [p for p in people if p["Family"] in german_fams]
    render_gen_german(OUT_GEN_DE, gpeople, by_id, counts)
    print(f"wrote {OUT_GEN_DE}  ({len(gpeople)} people, generation grid with proofs)")

    # ---------- 2) BY-GENERATION grid — RETIRED 1 Jul 2026 (single-tree policy) ----------
    # render_gen_grid(OUT_GEN, people, by_id, counts, total)
    # print(f"wrote {OUT_GEN}  (generation grid: {total} people)")


# ============================ BY-GENERATION GRID ============================
def esc(s):
    return html.escape(s or "", quote=True)

def render_gen_grid(outfile, people, by_id, counts, total):
    fam_keys = [f[0] for f in FAMILIES]
    fam_meta = {f[0]: {"title": f[1], "hap": f[2]} for f in FAMILIES}

    # bucket people by generation, then family
    gens = sorted({p["GenNum"] for p in people if p["GenNum"] is not None})
    buckets = {g: {k: [] for k in fam_keys} for g in gens}
    for p in people:
        if p["GenNum"] is None:
            continue
        buckets[p["GenNum"]][p["Family"]].append(p)
    # sort within a cell: by father (to group siblings), then name
    for g in gens:
        for k in fam_keys:
            buckets[g][k].sort(key=lambda p: (p["FatherID"], p["Name"]))

    # representative earliest year per generation (for the pre-immigration Swiss rows)
    import re as _re
    def first_year(p):
        for fld in (p["Birth"], p["Death"]):
            m = _re.search(r"\d{4}", fld or "")
            if m: return int(m.group())
        return None
    rep_year = {}
    for g in gens:
        yrs = [first_year(p) for k in fam_keys for p in buckets[g][k]]
        yrs = [y for y in yrs if y]
        rep_year[g] = min(yrs) if yrs else None

    def gen_label(g):
        if g >= 1:  return f"Gen&nbsp;{g}"
        if g == 0:  return "Gen&nbsp;0"
        return f"c.{rep_year[g]}" if rep_year.get(g) else "Origin"
    def gen_era(g):
        if g < 1:  return "Swiss<br>origins"
        return ""

    def proof_tag(pr):
        m = {"proven":("t-proven","Proven"),"documented":("t-documented","Doc."),
             "hypo":("t-hypo","Hypo."),"lore":("t-lore","Lore"),"living":("t-living","Living"),
             "disputed":("t-disputed","Disp.⁉"),
             "dna pedigree":("t-dnaped","DNA ped."),"hypo (pedigree)":("t-dnaped","DNA ped.")}
        t = m.get((pr or "").lower())
        return f'<span class="tag {t[0]}">{t[1]}</span>' if t else ""

    def card(p):
        cls = "pcard"
        if p["Direct"] == "yes": cls += " direct"
        elif p["Direct"] == "kit": cls += " kit"
        bits = []
        if p["Birth"] or p["Death"]:
            bits.append("b."+esc(p["Birth"] or "?") + (" – d."+esc(p["Death"]) if p["Death"] else ""))
        meta = " · ".join(bits)
        meta_html = f'<div class="pm">{meta}</div>' if meta else ""
        father = by_id.get(p["FatherID"])
        parent = f'<div class="pp">child of {esc(father["Name"])}</div>' if father else ""
        badges = proof_tag(p["Proof"])
        if p["DNAkit"]: badges += f'<span class="tag t-dna">{esc(p["DNAkit"])}</span>'
        if p["Direct"] == "yes": badges += '<span class="tag t-direct">Direct</span>'
        if p["Direct"] == "kit": badges += '<span class="tag t-kit">Kit</span>'
        return (f'<div class="{cls}"><div class="pn">{esc(p["Name"])}</div>'
                f'{meta_html}{parent}<div class="pb">{badges}</div></div>')

    rows_html = []
    for g in gens:
        era = gen_era(g)
        era_html = f'<span class="gera">{era}</span>' if era else ""
        cells = [f'<div class="axis"><span class="gnum">{gen_label(g)}</span>{era_html}</div>']
        for k in fam_keys:
            ppl = buckets[g][k]
            inner = "".join(card(p) for p in ppl) if ppl else '<div class="empty">—</div>'
            cells.append(f'<div class="cell {k.lower()}">{inner}</div>')
        rows_html.append('<div class="genrow">' + "".join(cells) + '</div>')

    heads = ['<div class="axishead"></div>']
    for k in fam_keys:
        heads.append(f'<div class="colhead {k.lower()}"><h2>{fam_meta[k]["title"]}'
                     f'<span class="hap {k.lower()}">Y-DNA {fam_meta[k]["hap"]}</span></h2></div>')
    head_html = '<div class="ghead">' + "".join(heads) + '</div>'

    out = (GEN_TEMPLATE
           .replace("__HEAD__", head_html)
           .replace("__ROWS__", "\n".join(rows_html))
           .replace("__TOTAL__", str(total))
           .replace("__CEDEN__", str(counts.get("Eden",0)))
           .replace("__CMANCH__", str(counts.get("Manch",0)))
           .replace("__CSWISS__", str(counts.get("Swiss",0))))
    with open(outfile, "w", encoding="utf-8") as fh:
        fh.write(out)


def render(outfile, payload, h1, sub, fam_controls, count_label):
    data_json = json.dumps(payload, ensure_ascii=False)
    out = (TEMPLATE
           .replace("/*DATA*/", data_json)
           .replace("__H1__", h1)
           .replace("__SUB__", sub)
           .replace("__FAMCONTROLS__", fam_controls)
           .replace("__COUNTLABEL__", count_label))
    with open(outfile, "w", encoding="utf-8") as fh:
        fh.write(out)

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>The Welty Families — Interactive Tree (All Lines)</title>
<style>
  :root{
    --ink:#1c1a17; --muted:#6b6459; --line:#c9bfae; --bg:#f6f1e7; --card:#fffdf8;
    --proven:#2e7d32; --documented:#1565c0; --hypo:#b26a00; --lore:#8e24aa;
    --living:#00695c; --direct:#b71c1c; --disputed:#6d4c41;
    --eden:#b71c1c; --manch:#3a4a5e; --swiss:#5e3a5e; --md:#2f6f4f; --r1a:#9a6a15;
    --eden-mid:#dd8b84; --eden-soft:#fbe8e6;
    --manch-mid:#7d93ad; --manch-soft:#e8eef6;
    --swiss-mid:#a878a8; --swiss-soft:#f3eaf3;
    --md-mid:#67a785; --md-soft:#e6f3ec;
    --r1a-mid:#c9a24e; --r1a-soft:#f5ecd6;
    --yrk:#0f6b6b; --yrk-mid:#4fa3a3; --yrk-soft:#dcefef;
    --gva:#a34a2a; --gva-mid:#cf9179; --gva-soft:#f6e7e0;
    --san:#2a4d7a; --san-mid:#7d9bc4; --san-soft:#e4ecf6;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);
    font-family:"Iowan Old Style","Palatino Linotype",Palatino,Georgia,serif;
    line-height:1.4;padding:0}
  .wrap{max-width:1500px;margin:0 auto;padding:24px 20px 90px}
  h1{font-size:28px;margin:0 0 4px;letter-spacing:.2px;text-align:center}
  .sub{color:var(--muted);margin:0 auto 14px;font-size:14px;max-width:1000px;text-align:center}

  /* controls */
  .controls{position:sticky;top:0;z-index:20;background:var(--bg);
    border-bottom:1px solid var(--line);padding:10px 0 12px;margin-bottom:8px;
    display:flex;flex-wrap:wrap;gap:10px 18px;align-items:center}
  .controls .grp{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
  .search{font:inherit;font-size:14px;padding:7px 12px;border:1px solid var(--line);
    border-radius:8px;min-width:230px;background:var(--card);color:var(--ink)}
  .chk{font-size:13px;display:inline-flex;gap:6px;align-items:center;cursor:pointer;
    padding:4px 11px 4px 9px;border:1.5px solid var(--line);border-radius:20px;background:var(--card);
    user-select:none;font-weight:600}
  /* colored dot swatch so each family reads at a glance in the control bar */
  .chk::before{content:"";width:10px;height:10px;border-radius:50%;background:var(--line);flex:0 0 auto}
  .chk input{accent-color:#7a6a4a}
  .chk.eden{border-color:var(--eden-mid);color:var(--eden)} .chk.eden::before{background:var(--eden)} .chk.eden input{accent-color:var(--eden)}
  .chk.manch{border-color:var(--manch-mid);color:var(--manch)} .chk.manch::before{background:var(--manch)} .chk.manch input{accent-color:var(--manch)}
  .chk.swiss{border-color:var(--swiss-mid);color:var(--swiss)} .chk.swiss::before{background:var(--swiss)} .chk.swiss input{accent-color:var(--swiss)}
  .chk.md{border-color:var(--md-mid);color:var(--md)} .chk.md::before{background:var(--md)} .chk.md input{accent-color:var(--md)}
  .chk.r1a{border-color:var(--r1a-mid);color:var(--r1a)} .chk.r1a::before{background:var(--r1a)} .chk.r1a input{accent-color:var(--r1a)}
  .chk.yrk{border-color:var(--yrk-mid);color:var(--yrk)} .chk.yrk::before{background:var(--yrk)} .chk.yrk input{accent-color:var(--yrk)}
  .chk.gva{border-color:var(--gva-mid);color:var(--gva)} .chk.gva::before{background:var(--gva)} .chk.gva input{accent-color:var(--gva)}
  .chk.san{border-color:var(--san-mid);color:var(--san)} .chk.san::before{background:var(--san)} .chk.san input{accent-color:var(--san)}
  select.gen{font:inherit;font-size:13px;padding:6px 9px;border:1px solid var(--line);border-radius:8px;background:var(--card)}
  .btn{font:inherit;font-size:13px;padding:6px 12px;border:1px solid var(--line);
    border-radius:8px;background:var(--card);cursor:pointer;color:#4a4238}
  .btn:hover{background:#efe8db}
  .btn.on{background:#4a4238;color:#fff;border-color:#4a4238}
  .count{font-size:12.5px;color:var(--muted)}

  .legend{background:var(--card);border:1px solid var(--line);border-radius:10px;
    padding:9px 14px;margin:8px 0 18px;font-size:12.5px;display:flex;flex-wrap:wrap;gap:6px 15px;align-items:center}
  .tag{display:inline-block;font-size:9.5px;font-weight:700;letter-spacing:.3px;
    text-transform:uppercase;padding:1px 6px;border-radius:20px;color:#fff;vertical-align:middle}
  .t-proven{background:var(--proven)} .t-documented{background:var(--documented)}
  .t-hypo{background:var(--hypo)} .t-lore{background:var(--lore)}
  .t-living{background:var(--living)} .t-dna{background:#37474f} .t-dnaped{background:#5d4037}
  .t-disputed{background:var(--disputed)}
  .genchip{display:inline-block;font-size:10px;font-weight:700;color:#4a4238;background:#eae2d2;
    border-radius:20px;padding:1px 7px;letter-spacing:.3px}

  /* source key */
  .sourcekey{background:var(--card);border:1px solid var(--line);border-radius:10px;
    padding:6px 14px;margin:0 0 14px;font-size:12.5px}
  .sourcekey summary{cursor:pointer;font-weight:600;color:#5a4632;padding:3px 0}
  .sourcekey .skbody{padding:4px 0 6px;max-width:1050px;color:#4a4238;line-height:1.5}
  .sourcekey .skbody p{margin:6px 0}
  .sourcekey .repos{font-size:12px;color:var(--muted)}
  .sourcekey a{color:#1565c0;text-decoration:none;border-bottom:1px dotted #9db8de}
  .sourcekey a:hover{background:#eef3fb}

  /* family blocks — each line gets its own color so the four families read
     apart at a glance: a bold accent bar on top, a saturated border, a tinted
     gradient header, and colored connector lines/toggles down the whole tree. */
  .fam{margin-top:26px;border-radius:12px;border:1px solid var(--line);
    border-top:5px solid var(--line);overflow:hidden;box-shadow:0 2px 12px rgba(90,70,30,.10)}
  .fam.eden{border-color:var(--eden-mid);border-top-color:var(--eden)}
  .fam.manch{border-color:var(--manch-mid);border-top-color:var(--manch)}
  .fam.swiss{border-color:var(--swiss-mid);border-top-color:var(--swiss)}
  .fam.md{border-color:var(--md-mid);border-top-color:var(--md)}
  .fam.r1a{border-color:var(--r1a-mid);border-top-color:var(--r1a)}
  .fam.yrk{border-color:var(--yrk-mid);border-top-color:var(--yrk)}
  .fam.gva{border-color:var(--gva-mid);border-top-color:var(--gva)}
  .fam.san{border-color:var(--san-mid);border-top-color:var(--san)}
  .famhd{padding:13px 16px 12px;border-bottom:2px solid var(--line)}
  .fam.eden .famhd{background:linear-gradient(180deg,var(--eden-soft),var(--card));border-bottom-color:var(--eden-mid)}
  .fam.manch .famhd{background:linear-gradient(180deg,var(--manch-soft),var(--card));border-bottom-color:var(--manch-mid)}
  .fam.swiss .famhd{background:linear-gradient(180deg,var(--swiss-soft),var(--card));border-bottom-color:var(--swiss-mid)}
  .fam.md .famhd{background:linear-gradient(180deg,var(--md-soft),var(--card));border-bottom-color:var(--md-mid)}
  .fam.r1a .famhd{background:linear-gradient(180deg,var(--r1a-soft),var(--card));border-bottom-color:var(--r1a-mid)}
  .fam.yrk .famhd{background:linear-gradient(180deg,var(--yrk-soft),var(--card));border-bottom-color:var(--yrk-mid)}
  .fam.gva .famhd{background:linear-gradient(180deg,var(--gva-soft),var(--card));border-bottom-color:var(--gva-mid)}
  .fam.san .famhd{background:linear-gradient(180deg,var(--san-soft),var(--card));border-bottom-color:var(--san-mid)}
  .famhd h2{margin:0;font-size:19px;letter-spacing:.2px}
  .fam.eden h2{color:var(--eden)} .fam.manch h2{color:var(--manch)} .fam.swiss h2{color:var(--swiss)} .fam.md h2{color:var(--md)} .fam.r1a h2{color:var(--r1a)} .fam.yrk h2{color:var(--yrk)} .fam.gva h2{color:var(--gva)} .fam.san h2{color:var(--san)}
  .hap{display:inline-block;font-size:10px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;
    padding:2px 8px;border-radius:20px;margin-left:8px;vertical-align:middle;color:#fff}
  .hap.eden{background:var(--eden)} .hap.manch{background:var(--manch)} .hap.swiss{background:var(--swiss)} .hap.md{background:var(--md)} .hap.r1a{background:var(--r1a)} .hap.yrk{background:var(--yrk)} .hap.gva{background:var(--gva)} .hap.san{background:var(--san)}
  .famdesc{font-size:12.5px;color:#5a564f;margin-top:5px;max-width:1000px}
  .famct{font-size:11.5px;color:var(--muted);margin-top:3px}
  .tree{padding:10px 16px 16px}
  /* family-colored tree structure: connector lines, toggles, and a soft left
     accent stripe on each person card (inset shadow so it never fights the
     disputed / unproven-link border states) */
  .fam.eden .kids{border-left-color:var(--eden-mid)}
  .fam.manch .kids{border-left-color:var(--manch-mid)}
  .fam.swiss .kids{border-left-color:var(--swiss-mid)}
  .fam.md .kids{border-left-color:var(--md-mid)}
  .fam.r1a .kids{border-left-color:var(--r1a-mid)}
  .fam.yrk .kids{border-left-color:var(--yrk-mid)}
  .fam.gva .kids{border-left-color:var(--gva-mid)}
  .fam.san .kids{border-left-color:var(--san-mid)}
  .fam.eden .tog{border-color:var(--eden-mid);color:var(--eden)}
  .fam.manch .tog{border-color:var(--manch-mid);color:var(--manch)}
  .fam.swiss .tog{border-color:var(--swiss-mid);color:var(--swiss)}
  .fam.md .tog{border-color:var(--md-mid);color:var(--md)}
  .fam.r1a .tog{border-color:var(--r1a-mid);color:var(--r1a)}
  .fam.yrk .tog{border-color:var(--yrk-mid);color:var(--yrk)}
  .fam.gva .tog{border-color:var(--gva-mid);color:var(--gva)}
  .fam.san .tog{border-color:var(--san-mid);color:var(--san)}
  .fam.eden .person{box-shadow:inset 3px 0 0 var(--eden-mid)}
  .fam.manch .person{box-shadow:inset 3px 0 0 var(--manch-mid)}
  .fam.swiss .person{box-shadow:inset 3px 0 0 var(--swiss-mid)}
  .fam.md .person{box-shadow:inset 3px 0 0 var(--md-mid)}
  .fam.r1a .person{box-shadow:inset 3px 0 0 var(--r1a-mid)}
  .fam.yrk .person{box-shadow:inset 3px 0 0 var(--yrk-mid)}
  .fam.gva .person{box-shadow:inset 3px 0 0 var(--gva-mid)}
  .fam.san .person{box-shadow:inset 3px 0 0 var(--san-mid)}

  /* nodes */
  .node{margin:4px 0}
  .kids{margin-left:20px;padding-left:14px;border-left:2px solid var(--line);
    display:none}
  .node.open>.kids{display:block}
  .row{display:flex;align-items:flex-start;gap:6px}
  .tog{flex:0 0 auto;width:18px;height:18px;margin-top:6px;border:1px solid var(--line);border-radius:5px;
    background:var(--card);cursor:pointer;font-size:12px;line-height:16px;text-align:center;color:#6b6459;user-select:none}
  .tog.leaf{visibility:hidden}
  .person{flex:1 1 auto;background:var(--card);border:1px solid var(--line);border-radius:9px;
    padding:6px 11px;margin:1px 0}
  .person .top{display:flex;flex-wrap:wrap;gap:5px;align-items:baseline}
  .name{font-weight:700;font-size:14.5px}
  .person.disputed{border-style:dashed;border-width:1.5px;border-color:var(--disputed);background:#f7f3ef}
  .person.disputed .name{color:var(--disputed)}
  .node.disputed>.row>.tog{border-style:dashed}
  /* unproven PARENT link (person is proven; the edge to the father is not) */
  .t-linksoft{background:#8d6e63}
  .person.linksoft{border-left:3px dashed var(--disputed)}
  .node.linksoft>.row>.tog{border-style:dashed;border-color:var(--disputed)}
  .linknote{font-size:10.5px;color:var(--disputed);margin-top:4px;font-style:italic;
    border-top:1px dotted var(--line);padding-top:3px;line-height:1.35}
  .meta{color:var(--muted);font-size:11.5px;margin-top:2px}
  .meta b{color:#5a4632}
  .notes{font-size:11px;color:var(--muted);font-style:italic;margin-top:2px}
  .prf{font-size:10.5px;color:#3d3528;margin-top:4px;border-top:1px dotted var(--line);
    padding-top:3px;line-height:1.35}
  .prf.none{color:#b3a98f}
  .src{font-size:10.5px;color:var(--muted);margin-top:2px;line-height:1.35}
  .prf b,.src b{color:#5a4632;font-size:9.5px;text-transform:uppercase;letter-spacing:.4px}
  .prf a,.src a{color:#1565c0;text-decoration:none;border-bottom:1px dotted #9db8de}
  .prf a:hover,.src a:hover{background:#eef3fb;border-bottom-color:#1565c0}
  .kidcount{font-size:10.5px;color:#9a9384}
  mark{background:#ffe9a8;color:inherit;padding:0 1px;border-radius:2px}
  .hidden{display:none !important}

  .foot{margin-top:26px;font-size:12px;color:var(--muted)}
  .foot b{color:#5a4632}
/*RECORDS-CSS-START*/
  .reclabel{font-size:11.5px;color:var(--muted);margin:6px 0 3px}
  .reclabel b{color:#5a4632;font-weight:600}
  .recstrip{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 2px}
  .recchip{position:relative;width:54px;height:54px;padding:0;border:1px solid var(--line);border-radius:6px;
    overflow:hidden;background:#efe7d3;line-height:0;cursor:zoom-in;
    box-shadow:0 1px 3px rgba(90,70,30,.18);transition:transform .12s ease,box-shadow .12s ease,border-color .12s}
  .recchip:hover,.recchip:focus{transform:translateY(-1px);border-color:#b8912f;
    box-shadow:0 3px 9px rgba(90,70,30,.32);outline:none}
  .recchip img{width:100%;height:100%;object-fit:cover;display:block}
  .recn{position:absolute;top:2px;right:2px;min-width:15px;height:15px;padding:0 3px;
    background:rgba(122,31,31,.92);color:#fff;font:600 10px/15px system-ui,sans-serif;
    text-align:center;border-radius:8px;box-shadow:0 1px 2px rgba(0,0,0,.35)}
  #reclb .recprev,#reclb .recnext{position:absolute;top:50%;transform:translateY(-50%);
    background:rgba(0,0,0,.35);border:0;color:#f6f1e7;font-size:40px;line-height:1;
    width:54px;height:80px;cursor:pointer;border-radius:6px}
  #reclb .recprev{left:12px} #reclb .recnext{right:12px}
  #reclb .recprev:hover,#reclb .recnext:hover{background:rgba(0,0,0,.6)}
  .recpg{display:inline-block;background:rgba(255,255,255,.15);border:1px solid #ffd777;
    color:#ffd777;font-size:12px;padding:1px 7px;border-radius:10px;margin-right:6px}
  #reclb{position:fixed;inset:0;z-index:200;display:none;background:rgba(20,16,10,.87);
    align-items:center;justify-content:center;flex-direction:column;padding:24px}
  #reclb.open{display:flex}
  #reclb img{max-width:94vw;max-height:80vh;object-fit:contain;border:3px solid #f6f1e7;
    box-shadow:0 6px 34px rgba(0,0,0,.55)}
  #reclb .cap{color:#f6f1e7;max-width:900px;margin-top:12px;font-size:14px;text-align:center;line-height:1.45}
  #reclb .cap a{color:#ffd777}
  #reclb .x{position:absolute;top:12px;right:20px;background:none;border:0;color:#f6f1e7;
    font-size:32px;line-height:1;cursor:pointer}
  @media(max-width:640px){.recchip{width:46px;height:46px}}
/*RECORDS-CSS-END*/
  @media print{
    body{background:#fff;padding:0}.controls{display:none}
    .kids{display:block !important}.tog{display:none}
    .person{box-shadow:none}.node{break-inside:avoid}
    .recstrip,.reclabel{display:none}
  }
</style>
</head>
<body>
<div class="wrap">
<h1>__H1__</h1>
<p class="sub">__SUB__</p>

<div class="controls">
  <div class="grp">
    <input id="q" class="search" type="search" placeholder="Search name, place, spouse, notes…" autocomplete="off">
  </div>
  __FAMCONTROLS__
  <div class="grp">
    <select id="gen" class="gen"><option value="">All generations</option></select>
  </div>
  <div class="grp">
    <button id="expandAll" class="btn">Expand all</button>
    <button id="collapseAll" class="btn">Collapse all</button>
  </div>
  <span id="count" class="count"></span>
</div>

<div class="legend">
  <span style="font-weight:600">Key:</span>
  <span><span class="tag t-proven">Proven</span> primary records</span>
  <span><span class="tag t-documented">Documented</span></span>
  <span><span class="tag t-hypo">Hypothesized</span> unproven rung</span>
  <span><span class="tag t-lore">Lore</span></span>
  <span><span class="tag t-living">Living</span></span>
  <span><span class="tag t-dna">DNA</span> Y-line tester</span>
  <span><span class="tag t-disputed">Disputed link ⁉</span> same household, different Y-DNA (non-paternity)</span>
  <span><span class="tag t-linksoft">Father link unproven ⁉</span> the person is proven; their descent from the parent above is not</span>
</div>

<details class="sourcekey">
  <summary>How to read the sources &amp; citations</summary>
  <div class="skbody">
    <p>Every documented person carries two lines: <b>Proof</b> &mdash; the specific record that
    establishes them (a baptism, will, tax entry, census, gravestone) &mdash; and <b>Source</b> &mdash;
    exactly where that record can be found. Blue links open the record or repository in a new tab.
    Where no primary record has been located yet, the person is tagged <span class="tag t-hypo">Hypothesized</span>
    or <span class="tag t-lore">Lore</span> and the proof line is left open.</p>
    <p class="repos"><b>Repositories used:</b>
    <a href="https://www.archion.de" target="_blank" rel="noopener">Archion</a> (German Protestant church registers &mdash; the Edenkoben Reformed books; <i>subscription</i>) ·
    <a href="https://www.familysearch.org" target="_blank" rel="noopener">FamilySearch</a> (church &amp; court records, Full-Text Search, Digital Library; <i>free account</i>) ·
    <a href="https://www.ancestry.com" target="_blank" rel="noopener">Ancestry</a> (PA tax, exoneration &amp; church collections; <i>subscription</i>) ·
    <a href="https://www.findagrave.com" target="_blank" rel="noopener">FindAGrave</a> (memorials &amp; cemeteries) ·
    <a href="https://www.familytreedna.com/groups/welty" target="_blank" rel="noopener">Welty Y-DNA project</a> (FTDNA; patrilineage evidence) ·
    plus published works (Strassburger &amp; Hinke <i>Pennsylvania German Pioneers</i>; Bowen; Beers), <a href="https://archive.org" target="_blank" rel="noopener">archive.org</a> book scans, <a href="https://gameo.org" target="_blank" rel="noopener">GAMEO</a>, <a href="https://www.wikitree.com" target="_blank" rel="noopener">WikiTree</a>, and the compiler's family papers.</p>
  </div>
</details>

<div id="chart"></div>

<div class="foot">
  <p style="margin-top:10px;color:#9a9384">Auto-generated from <b>Welty Ancestry Research Log.xlsx</b> &rarr; sheet &ldquo;People Roster (chart source)&rdquo; via generate_chart.py. Edit the sheet and re-run to update.</p>
</div>
</div>

<script>
const DATA = /*DATA*/;
const P = DATA.people;

function proofTag(pr){
  const m={proven:['t-proven','Proven'],documented:['t-documented','Documented'],
           hypo:['t-hypo','Hypothesized'],lore:['t-lore','Lore'],living:['t-living','Living'],
           disputed:['t-disputed','Disputed link ⁉']};
  const t=m[(pr||'').toLowerCase()]; return t?`<span class="tag ${t[0]}">${t[1]}</span>`:'';
}
function esc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function genLabel(g,fam){ if(g===null||g===undefined||g==='') return ''; if(g<1) return fam==='Swiss'?'Swiss origin':'Ancestor'; return 'Gen '+g; }

function metaLine(p){
  const bits=[];
  if(p.birth||p.death){ bits.push('b. '+(p.birth||'?')+(p.death?(' · d. '+p.death):'')); }
  if(p.place) bits.push(esc(p.place));
  if(p.spouse) bits.push('m. '+esc(p.spouse));
  return bits.join(' · ');
}

/*RECORDS-JS-START*/
// Each entry in `groups` is ONE document (a record) that may have several page
// images; the chip shows one thumbnail with a page-count badge, and the lightbox
// pages through them. `groups` is p.records from the manifest.
function recStrip(groups){
  let s=`<div class="reclabel">&#128247; <b>Record${groups.length>1?'s':''}</b> · ${groups.length} &mdash; primary manuscript${groups.length>1?'s':''}, click to enlarge</div><div class="recstrip">`;
  s+=groups.map((g,i)=>{
    const n=g.pages.length;
    const badge=n>1?`<span class="recn" title="${n} pages">${n}</span>`:'';
    const payload=encodeURIComponent(JSON.stringify({pages:g.pages,url:g.url||'',repo:g.repo||''}));
    return `<button type="button" class="recchip" data-rec="${payload}" onclick="recOpen(this)" aria-label="Open record: ${esc(g.caption)}"><img loading="lazy" src="records/thumb/${g.slug}.jpg" alt="${esc(g.caption)}">${badge}</button>`;
  }).join('');
  return s+`</div>`;
}
function recOpen(btn){
  const c=JSON.parse(decodeURIComponent(btn.dataset.rec));
  window._recCtx={pages:c.pages,url:c.url,repo:c.repo,i:0};
  recShow(); const lb=document.getElementById('reclb');
  lb.classList.add('open'); document.body.style.overflow='hidden';
}
function recShow(){
  const c=window._recCtx; if(!c) return;
  const lb=document.getElementById('reclb'), p=c.pages[c.i], multi=c.pages.length>1;
  lb.querySelector('img').src='records/full/'+p.slug+'.jpg';
  let cap=(multi?`<span class="recpg">Page ${c.i+1} / ${c.pages.length}</span> `:'')+esc(p.caption);
  if(c.url) cap+=` <a href="${c.url}" target="_blank" rel="noopener">&mdash; view at ${esc(c.repo)||'source'}</a>`;
  lb.querySelector('.cap').innerHTML=cap;
  lb.querySelector('.recprev').style.display=multi?'':'none';
  lb.querySelector('.recnext').style.display=multi?'':'none';
}
function recStep(d,ev){ if(ev) ev.stopPropagation(); const c=window._recCtx; if(!c) return;
  c.i=(c.i+d+c.pages.length)%c.pages.length; recShow(); }
function recClose(){const lb=document.getElementById('reclb');lb.classList.remove('open');lb.querySelector('img').src='';document.body.style.overflow='';window._recCtx=null;}
document.addEventListener('keydown',e=>{
  if(!document.getElementById('reclb')||!document.getElementById('reclb').classList.contains('open'))return;
  if(e.key==='Escape')recClose();
  else if(e.key==='ArrowRight')recStep(1);
  else if(e.key==='ArrowLeft')recStep(-1);
});
/*RECORDS-JS-END*/
function nodeHTML(id,famKey){
  const p=P[id]; const kids=p.kids||[];
  let cls='person';
  if((p.proof||'').toLowerCase()==='disputed') cls+=' disputed';
  if(p.linkunproven) cls+=' linksoft';
  let badges=proofTag(p.proof);
  if(p.dna) badges+=`<span class="tag t-dna">${esc(p.dna)}</span>`;
  if(p.linkunproven) badges+='<span class="tag t-linksoft">Father link unproven ⁉</span>';
  const g=genLabel(p.gen, p.fam);
  const genchip=g?`<span class="genchip">${g}</span>`:'';
  const meta=metaLine(p);
  const kc=kids.length?`<span class="kidcount">(${kids.length})</span>`:'';
  let h=`<div class="node${(p.proof||'').toLowerCase()==='disputed'?' disputed':''}${p.linkunproven?' linksoft':''}" data-id="${id}" data-fam="${p.fam}" data-gen="${p.gen===null?'':p.gen}" data-direct="${p.direct||''}">`;
  h+=`<div class="row">`;
  h+=`<div class="tog${kids.length?'':' leaf'}">${kids.length?'▸':''}</div>`;
  h+=`<div class="${cls}"><div class="top"><span class="name">${esc(p.name)}</span>${genchip}${badges}${kc}</div>`;
  if(meta) h+=`<div class="meta">${meta}</div>`;
  if(p.notes_html) h+=`<div class="notes">${p.notes_html}</div>`;
  else if(p.notes) h+=`<div class="notes">${esc(p.notes)}</div>`;
  if(p.proof_html) h+=`<div class="prf"><b>Proof:</b> ${p.proof_html}</div>`;
  if(p.source_html) h+=`<div class="src"><b>Source:</b> ${p.source_html}</div>`;
  if(p.records&&p.records.length) h+=recStrip(p.records); /*RECORDS-RENDER*/
  if(p.linkunproven) h+=`<div class="linknote">&#8265; Descent from the father above is not yet proven by a primary record naming the parent &mdash; it rests on indirect evidence (Y-DNA and/or circumstantial records). A baptism or record naming the father would confirm it.</div>`;
  h+=`</div></div>`;
  if(kids.length){ h+=`<div class="kids">`+kids.map(k=>nodeHTML(k,famKey)).join('')+`</div>`; }
  h+=`</div>`;
  return h;
}

// render families
const chart=document.getElementById('chart');
DATA.families.forEach(f=>{
  const div=document.createElement('div');
  div.className='fam '+f.key.toLowerCase(); div.dataset.fam=f.key;
  const ct=(f.roots||[]).length;
  div.innerHTML=`<div class="famhd"><h2>${f.title}<span class="hap ${f.key.toLowerCase()}">Y-DNA ${f.hap}</span></h2>`+
                `<div class="famdesc">${f.desc}</div></div>`+
                `<div class="tree">`+ (f.roots||[]).map(r=>nodeHTML(r,f.key)).join('') +`</div>`;
  chart.appendChild(div);
});

// generation dropdown options
const gens=[...new Set(Object.values(P).map(p=>p.gen).filter(g=>g!==null))].sort((a,b)=>a-b);
const genSel=document.getElementById('gen');
gens.forEach(g=>{const o=document.createElement('option');o.value=g;o.textContent=(g<1?('Swiss origin ('+g+')'):('Generation '+g));genSel.appendChild(o);});

// toggle behaviour
function setOpen(node,open){ node.classList.toggle('open',open); const t=node.querySelector(':scope>.row>.tog'); if(t&&!t.classList.contains('leaf')) t.textContent=open?'▾':'▸'; }
document.addEventListener('click',e=>{
  const tog=e.target.closest('.tog'); if(!tog||tog.classList.contains('leaf'))return;
  const node=tog.closest('.node'); setOpen(node,!node.classList.contains('open'));
});
function allNodes(){return [...document.querySelectorAll('.node')];}
function openAll(v){allNodes().forEach(n=>{ if(n.querySelector(':scope>.kids')) setOpen(n,v); });}
document.getElementById('expandAll').onclick=()=>openAll(true);
document.getElementById('collapseAll').onclick=()=>openAll(false);

// default: open the top two generations of every family (neutral overview)
function openDefault(){
  document.querySelectorAll('.fam>.tree>.node').forEach(root=>{
    setOpen(root,true);
    root.querySelectorAll(':scope>.kids>.node').forEach(n=>setOpen(n,true));
  });
}
openDefault();

// filtering
const q=document.getElementById('q');
const famChecks=[...document.querySelectorAll('.chk input')];
famChecks.forEach(c=>c.addEventListener('change',applyFilter));
genSel.addEventListener('change',applyFilter);
q.addEventListener('input',applyFilter);

function matchText(p,term){
  if(!term)return true;
  return (p.name+' '+p.place+' '+p.spouse+' '+p.notes+' '+p.dna).toLowerCase().includes(term);
}
function highlight(node,term){
  const nameEl=node.querySelector(':scope>.row .name');
  if(!nameEl)return;
  const raw=P[node.dataset.id].name;
  if(term && raw.toLowerCase().includes(term)){
    const i=raw.toLowerCase().indexOf(term);
    nameEl.innerHTML=esc(raw.slice(0,i))+'<mark>'+esc(raw.slice(i,i+term.length))+'</mark>'+esc(raw.slice(i+term.length));
  } else { nameEl.textContent=raw; }
}

function applyFilter(){
  const term=q.value.trim().toLowerCase();
  const famOn={};
  if(famChecks.length){ famChecks.forEach(c=>famOn[c.dataset.fam]=c.checked); }
  else { DATA.families.forEach(f=>famOn[f.key]=true); }
  const genv=genSel.value;
  let shown=0;
  // family blocks
  DATA.families.forEach(f=>{
    document.querySelector('.fam.'+f.key).classList.toggle('hidden',!famOn[f.key]);
  });
  // recursive visibility: a node is visible if it matches AND (no gen filter or gen matches) etc,
  // OR one of its descendants is visible. Ancestors of matches stay visible for context.
  function visit(node){
    const p=P[node.dataset.id];
    let selfMatch = matchText(p,term)
      && (genv==='' || String(p.gen)===genv);
    let childVisible=false;
    const kidNodes=[...node.querySelectorAll(':scope>.kids>.node')];
    kidNodes.forEach(k=>{ if(visit(k)) childVisible=true; });
    const visible = selfMatch || childVisible;
    node.classList.toggle('hidden',!visible);
    if(visible){ shown++; highlight(node,term);
      if(childVisible && (term||genv)) setOpen(node,true);
    }
    return visible;
  }
  DATA.families.forEach(f=>{
    if(!famOn[f.key])return;
    [...document.querySelectorAll('.fam.'+f.key+'>.tree>.node')].forEach(visit);
  });
  const active = term||genv;
  document.getElementById('count').textContent =
    active ? (shown+' shown') : ('__COUNTLABEL__');
  if(!active) openDefault();
}
applyFilter();
</script>
<!--RECORDS-LB-START--><div id="reclb" onclick="if(event.target.id==='reclb')recClose()"><button type="button" class="x" onclick="recClose()" aria-label="Close">&times;</button><button type="button" class="recprev" onclick="recStep(-1,event)" aria-label="Previous page">&#8249;</button><img alt="Record image"><button type="button" class="recnext" onclick="recStep(1,event)" aria-label="Next page">&#8250;</button><div class="cap"></div></div><!--RECORDS-LB-END-->
</body>
</html>
"""

# ---------------- BY-GENERATION grid template (static, offline, no JS needed) ----------------
GEN_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>The Welty Families — By Generation (all three lines side by side)</title>
<style>
  :root{
    --ink:#1c1a17; --muted:#6b6459; --line:#c9bfae; --bg:#f6f1e7; --card:#fffdf8;
    --proven:#2e7d32; --documented:#1565c0; --hypo:#b26a00; --lore:#8e24aa;
    --living:#00695c; --direct:#b71c1c; --disputed:#6d4c41;
    --eden:#b71c1c; --manch:#3a4a5e; --swiss:#5e3a5e; --md:#2f6f4f; --r1a:#9a6a15;
    --eden-mid:#dd8b84; --eden-soft:#fbe8e6;
    --manch-mid:#7d93ad; --manch-soft:#e8eef6;
    --swiss-mid:#a878a8; --swiss-soft:#f3eaf3;
    --md-mid:#67a785; --md-soft:#e6f3ec;
    --r1a-mid:#c9a24e; --r1a-soft:#f5ecd6;
    --yrk:#0f6b6b; --yrk-mid:#4fa3a3; --yrk-soft:#dcefef;
    --gva:#a34a2a; --gva-mid:#cf9179; --gva-soft:#f6e7e0;
    --san:#2a4d7a; --san-mid:#7d9bc4; --san-soft:#e4ecf6;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);
    font-family:"Iowan Old Style","Palatino Linotype",Palatino,Georgia,serif;
    line-height:1.35;padding:24px 18px 90px}
  .wrap{max-width:1600px;margin:0 auto}
  h1{font-size:27px;margin:0 0 4px;letter-spacing:.2px}
  .sub{color:var(--muted);margin:0 0 12px;font-size:14px;max-width:1050px}
  .legend{background:var(--card);border:1px solid var(--line);border-radius:10px;
    padding:9px 14px;margin:8px 0 14px;font-size:12px;display:flex;flex-wrap:wrap;gap:6px 14px;align-items:center}
  .tag{display:inline-block;font-size:9px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;
    padding:1px 6px;border-radius:20px;color:#fff;vertical-align:middle}
  .t-proven{background:var(--proven)} .t-documented{background:var(--documented)}
  .t-hypo{background:var(--hypo)} .t-lore{background:var(--lore)}
  .t-living{background:var(--living)} .t-dna{background:#37474f} .t-dnaped{background:#5d4037}
  .t-direct{background:var(--direct)} .t-kit{background:#b8860b} .t-disputed{background:var(--disputed)}

  /* grid: axis + 3 equal family columns; every row shares the template so columns line up */
  .ghead,.genrow{display:grid;grid-template-columns:78px 1fr 1fr 1fr;gap:10px;align-items:stretch}
  .ghead{position:sticky;top:0;z-index:5;background:var(--bg);padding:4px 0 8px}
  .colhead{border-radius:10px 10px 0 0;padding:9px 13px;border:1px solid var(--line);border-bottom:none}
  .colhead.eden{background:#fff1ef;border-color:#e6b7b1}
  .colhead.manch{background:#eef2f7;border-color:#b9c6d6}
  .colhead.swiss{background:#f5eef5;border-color:#cdb6cd}
  .colhead h2{margin:0;font-size:16px}
  .colhead.eden h2{color:var(--eden)} .colhead.manch h2{color:var(--manch)} .colhead.swiss h2{color:var(--swiss)}
  .hap{display:inline-block;font-size:9.5px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;
    padding:2px 7px;border-radius:20px;margin-left:7px;vertical-align:middle;color:#fff}
  .hap.eden{background:var(--eden)} .hap.manch{background:var(--manch)} .hap.swiss{background:var(--swiss)}

  .genrow{margin-bottom:8px}
  .genrow:nth-child(even) .cell{background:#00000005}
  .axis{display:flex;flex-direction:column;align-items:center;justify-content:flex-start;padding-top:10px;text-align:center}
  .axis .gnum{font-size:13px;font-weight:700;color:#4a4238;letter-spacing:.4px}
  .axis .gera{font-size:9.5px;color:var(--muted);margin-top:3px;line-height:1.15}
  .cell{border-left:3px solid transparent;border-radius:6px;padding:4px 4px 4px 8px;min-height:26px}
  .cell.eden{border-color:#e6b7b1} .cell.manch{border-color:#b9c6d6} .cell.swiss{border-color:#cdb6cd}
  .empty{color:#c9bfae;font-size:16px;padding:2px 4px}

  .pcard{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:5px 9px;margin:4px 0}
  .pcard.direct{border-color:var(--direct);border-width:1.5px;background:#fffaf8}
  .pcard.kit{border-color:#b8860b;border-width:2px;background:#fffbe9;box-shadow:0 0 0 2px #f0dd9e}
  .pn{font-weight:700;font-size:13.5px}
  .pcard.direct .pn{color:var(--direct)} .pcard.kit .pn{color:#8a6100}
  .pm{font-size:11px;color:var(--muted);margin-top:1px}
  .pp{font-size:10px;color:#8a8375;font-style:italic;margin-top:1px}
  .pb{margin-top:3px;display:flex;flex-wrap:wrap;gap:3px}

  .foot{margin-top:26px;font-size:12px;color:var(--muted)}
  .foot b{color:#5a4632}
  @media print{body{background:#fff;padding:0}.ghead{position:static}.genrow{break-inside:avoid}}
</style>
</head>
<body>
<div class="wrap">
<h1>The Welty Families &mdash; by generation</h1>
<p class="sub">All three genetically distinct Welty lines side by side, aligned by generation. Read <b>down</b> a column to follow one family; read <b>across</b> a row to compare who sits in the same generation across all three. Generation&nbsp;1 = each line's American founder (the Swiss line reaches further back, shown as pre-immigration rows). Companion to the interactive master tree; both are built from the same <b>People Roster</b>. <b>__TOTAL__</b> people &mdash; Edenkoben __CEDEN__ · Manchester __CMANCH__ · Swiss __CSWISS__.</p>

<div class="legend">
  <span style="font-weight:600">Key:</span>
  <span><span class="tag t-proven">Proven</span></span>
  <span><span class="tag t-documented">Doc.</span></span>
  <span><span class="tag t-hypo">Hypo.</span></span>
  <span><span class="tag t-lore">Lore</span></span>
  <span><span class="tag t-living">Living</span></span>
  <span><span class="tag t-dna">DNA</span></span>
  <span style="color:var(--direct);font-weight:600">Red = the proven direct paternal Welty line</span>
  <span><span class="tag t-kit">Kit</span> Y-DNA kit owner</span>
</div>

__HEAD__
__ROWS__

<div class="foot">
  <p><b>Why three families:</b> the surname Welty (Welti / W&auml;lti / Weldy) arose independently in several German- and Swiss-speaking places; three of those families settled the same York County, PA townships and reused the same given names. Y-DNA is the clean separator — <b>R1b</b> = Edenkoben, <b>I1</b> = Manchester, <b>I2b</b> = Swiss Emmental.</p>
  <p style="margin-top:8px;color:#9a9384">Auto-generated from Welty Ancestry Research Log.xlsx → “People Roster (chart source)” via generate_chart.py. Edit the roster and re-run to update.</p>
</div>
</div>
</body>
</html>
"""

# ---------------- GERMAN-LINES BY-GENERATION GRID (proof record + source per card) ----------------
def render_gen_german(outfile, people, by_id, counts):
    fam_keys = ["Eden", "Manch"]
    fam_meta = {f[0]: {"title": f[1], "hap": f[2]} for f in FAMILIES}

    gens = sorted({p["GenNum"] for p in people if p["GenNum"] is not None})
    buckets = {g: {k: [] for k in fam_keys} for g in gens}
    for p in people:
        if p["GenNum"] is None:
            continue
        buckets[p["GenNum"]][p["Family"]].append(p)
    for g in gens:
        for k in fam_keys:
            buckets[g][k].sort(key=lambda p: (p["FatherID"], p["Name"]))

    def gen_label(g):
        if g >= 1:  return f"Gen&nbsp;{g}"
        if g == 0:  return "Gen&nbsp;0"
        return "Apex"
    def gen_era(g):
        return "German<br>origins" if g < 1 else ""

    def proof_tag(pr):
        m = {"proven":("t-proven","Proven"),"documented":("t-documented","Doc."),
             "hypo":("t-hypo","Hypo."),"lore":("t-lore","Lore"),"living":("t-living","Living"),
             "disputed":("t-disputed","Disp.&#8265;")}
        t = m.get((pr or "").lower())
        return f'<span class="tag {t[0]}">{t[1]}</span>' if t else ""

    def card(p):
        cls = "pcard"
        if p["Direct"] == "yes": cls += " direct"
        elif p["Direct"] == "kit": cls += " kit"
        if (p["Proof"] or "").lower() == "disputed": cls += " disputed"
        bits = []
        if p["Birth"] or p["Death"]:
            bits.append("b."+esc(p["Birth"] or "?") + (" &ndash; d."+esc(p["Death"]) if p["Death"] else ""))
        meta = " · ".join(bits)
        meta_html = f'<div class="pm">{meta}</div>' if meta else ""
        father = by_id.get(p["FatherID"])
        parent = f'<div class="pp">child of {esc(father["Name"])}</div>' if father else ""
        badges = proof_tag(p["Proof"])
        if p["DNAkit"]: badges += f'<span class="tag t-dna">{esc(p["DNAkit"])}</span>'
        if p["Direct"] == "yes": badges += '<span class="tag t-direct">Direct</span>'
        if p["Direct"] == "kit": badges += '<span class="tag t-kit">Kit</span>'
        prf = f'<div class="prf"><b>Proof:</b> {format_source_html(p["ProofRec"])}</div>' if p["ProofRec"] else '<div class="prf none"><b>Proof:</b> &mdash;</div>'
        src = f'<div class="src"><b>Source:</b> {format_source_html(p["Source"])}</div>' if p["Source"] else ""
        return (f'<div class="{cls}"><div class="pn">{esc(p["Name"])}</div>'
                f'{meta_html}{parent}<div class="pb">{badges}</div>{prf}{src}</div>')

    rows_html = []
    for g in gens:
        era = gen_era(g)
        era_html = f'<span class="gera">{era}</span>' if era else ""
        cells = [f'<div class="axis"><span class="gnum">{gen_label(g)}</span>{era_html}</div>']
        for k in fam_keys:
            ppl = buckets[g][k]
            inner = "".join(card(p) for p in ppl) if ppl else '<div class="empty">&mdash;</div>'
            cells.append(f'<div class="cell {k.lower()}">{inner}</div>')
        rows_html.append('<div class="genrow">' + "".join(cells) + '</div>')

    heads = ['<div class="axishead"></div>']
    for k in fam_keys:
        heads.append(f'<div class="colhead {k.lower()}"><h2>{fam_meta[k]["title"]}'
                     f'<span class="hap {k.lower()}">Y-DNA {fam_meta[k]["hap"]}</span></h2></div>')
    head_html = '<div class="ghead">' + "".join(heads) + '</div>'

    total = len(people)
    out = (GEN_DE_TEMPLATE
           .replace("__HEAD__", head_html)
           .replace("__ROWS__", "\n".join(rows_html))
           .replace("__TOTAL__", str(total))
           .replace("__CEDEN__", str(counts.get("Eden", 0)))
           .replace("__CMANCH__", str(counts.get("Manch", 0))))
    with open(outfile, "w", encoding="utf-8") as fh:
        fh.write(out)

GEN_DE_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>The German Welty Lines — By Generation (with proofs)</title>
<style>
  :root{
    --ink:#1c1a17; --muted:#6b6459; --line:#c9bfae; --bg:#f6f1e7; --card:#fffdf8;
    --proven:#2e7d32; --documented:#1565c0; --hypo:#b26a00; --lore:#8e24aa;
    --living:#00695c; --direct:#b71c1c; --disputed:#6d4c41;
    --eden:#b71c1c; --manch:#3a4a5e;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);
    font-family:"Iowan Old Style","Palatino Linotype",Palatino,Georgia,serif;
    line-height:1.35;padding:24px 18px 90px}
  .wrap{max-width:1500px;margin:0 auto}
  h1{font-size:27px;margin:0 0 4px;letter-spacing:.2px}
  .sub{color:var(--muted);margin:0 0 12px;font-size:14px;max-width:1050px}
  .legend{background:var(--card);border:1px solid var(--line);border-radius:10px;
    padding:9px 14px;margin:8px 0 14px;font-size:12px;display:flex;flex-wrap:wrap;gap:6px 14px;align-items:center}
  .tag{display:inline-block;font-size:9px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;
    padding:1px 6px;border-radius:20px;color:#fff;vertical-align:middle}
  .t-proven{background:var(--proven)} .t-documented{background:var(--documented)}
  .t-hypo{background:var(--hypo)} .t-lore{background:var(--lore)}
  .t-living{background:var(--living)} .t-dna{background:#37474f} .t-dnaped{background:#5d4037}
  .t-direct{background:var(--direct)} .t-kit{background:#b8860b} .t-disputed{background:var(--disputed)}

  .ghead,.genrow{display:grid;grid-template-columns:78px 1.25fr 1fr;gap:10px;align-items:stretch}
  .ghead{position:sticky;top:0;z-index:5;background:var(--bg);padding:4px 0 8px}
  .colhead{border-radius:10px 10px 0 0;padding:9px 13px;border:1px solid var(--line);border-bottom:none}
  .colhead.eden{background:#fff1ef;border-color:#e6b7b1}
  .colhead.manch{background:#eef2f7;border-color:#b9c6d6}
  .colhead h2{margin:0;font-size:16px}
  .colhead.eden h2{color:var(--eden)} .colhead.manch h2{color:var(--manch)}
  .hap{display:inline-block;font-size:9.5px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;
    padding:2px 7px;border-radius:20px;margin-left:7px;vertical-align:middle;color:#fff}
  .hap.eden{background:var(--eden)} .hap.manch{background:var(--manch)}

  .genrow{margin-bottom:8px}
  .genrow:nth-child(even) .cell{background:#00000005}
  .axis{display:flex;flex-direction:column;align-items:center;justify-content:flex-start;padding-top:10px;text-align:center}
  .axis .gnum{font-size:13px;font-weight:700;color:#4a4238;letter-spacing:.4px}
  .axis .gera{font-size:9.5px;color:var(--muted);margin-top:3px;line-height:1.15}
  .cell{border-left:3px solid transparent;border-radius:6px;padding:4px 4px 4px 8px;min-height:26px;
    display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:6px;align-content:start}
  .cell.eden{border-color:#e6b7b1} .cell.manch{border-color:#b9c6d6}
  .empty{color:#c9bfae;font-size:16px;padding:2px 4px}

  .pcard{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:6px 10px}
  .pcard.direct{border-color:var(--direct);border-width:1.5px;background:#fffaf8}
  .pcard.kit{border-color:#b8860b;border-width:2px;background:#fffbe9;box-shadow:0 0 0 2px #f0dd9e}
  .pcard.disputed{border-style:dashed;border-width:1.5px;border-color:var(--disputed);background:#f7f3ef}
  .pn{font-weight:700;font-size:13.5px}
  .pcard.direct .pn{color:var(--direct)} .pcard.kit .pn{color:#8a6100} .pcard.disputed .pn{color:var(--disputed)}
  .pm{font-size:11px;color:var(--muted);margin-top:1px}
  .pp{font-size:10px;color:#8a8375;font-style:italic;margin-top:1px}
  .pb{margin-top:3px;display:flex;flex-wrap:wrap;gap:3px}
  .prf{font-size:10.5px;color:#3d3528;margin-top:4px;border-top:1px dotted var(--line);padding-top:3px;line-height:1.3}
  .prf b,.src b{color:#5a4632;font-size:9.5px;text-transform:uppercase;letter-spacing:.4px}
  .prf.none{color:#b3a98f}
  .src{font-size:10.5px;color:var(--muted);margin-top:2px;line-height:1.3}
  .prf a,.src a{color:#1565c0;text-decoration:none;border-bottom:1px dotted #9db8de}
  .prf a:hover,.src a:hover{background:#eef3fb;border-bottom-color:#1565c0}

  .sourcekey{background:var(--card);border:1px solid var(--line);border-radius:10px;
    padding:6px 14px;margin:0 0 14px;font-size:12.5px}
  .sourcekey summary{cursor:pointer;font-weight:600;color:#5a4632;padding:3px 0}
  .sourcekey .skbody{padding:4px 0 6px;max-width:1050px;color:#4a4238;line-height:1.5}
  .sourcekey .skbody p{margin:6px 0}
  .sourcekey .repos{font-size:12px;color:var(--muted)}
  .sourcekey a{color:#1565c0;text-decoration:none;border-bottom:1px dotted #9db8de}
  .sourcekey a:hover{background:#eef3fb}

  .foot{margin-top:26px;font-size:12px;color:var(--muted)}
  .foot b{color:#5a4632}
  @media print{body{background:#fff;padding:0}.ghead{position:static}.genrow{break-inside:avoid}.pcard{break-inside:avoid}}
</style>
</head>
<body>
<div class="wrap">
<h1>The German Welty family &mdash; by generation, with proofs</h1>
<p class="sub">One large German family, two Y-lines, side by side and aligned by generation &mdash;
the <b style="color:var(--eden)">Edenkoben R1b spine</b> and the <b style="color:var(--manch)">Manchester branch (I1)</b>,
whose Y-line marks a paternity break inside the family at/above Georg Wolfgang (bapt. ~1716 in the same household), not a separate clan.
The genuinely unrelated Swiss family is excluded. Read <b>down</b> a column to follow one line; read <b>across</b> to compare generations.
Every card carries the <b>official record</b> that proves the person and <b>where that record lives</b>.
<b>__TOTAL__</b> people &mdash; Edenkoben __CEDEN__ · Manchester branch __CMANCH__.</p>

<div class="legend">
  <span style="font-weight:600">Key:</span>
  <span><span class="tag t-proven">Proven</span> primary records</span>
  <span><span class="tag t-documented">Doc.</span></span>
  <span><span class="tag t-hypo">Hypo.</span> unproven rung</span>
  <span><span class="tag t-lore">Lore</span></span>
  <span><span class="tag t-living">Living</span></span>
  <span><span class="tag t-dna">DNA</span></span>
  <span style="color:var(--direct);font-weight:600">Red = the proven direct paternal Welty line</span>
  <span><span class="tag t-kit">Kit</span> Y-DNA kit owner</span>
  <span><span class="tag t-disputed">Disp.&#8265;</span> dashed = non-paternity</span>
</div>

<details class="sourcekey">
  <summary>How to read the sources &amp; citations</summary>
  <div class="skbody">
    <p>Every card carries two lines: <b>Proof</b> &mdash; the specific record that establishes the
    person (baptism, will, tax entry, census, gravestone) &mdash; and <b>Source</b> &mdash; exactly where
    that record can be found. Blue links open the record or repository in a new tab.</p>
    <p class="repos"><b>Repositories used:</b>
    <a href="https://www.archion.de" target="_blank" rel="noopener">Archion</a> (Edenkoben Reformed church registers; <i>subscription</i>) ·
    <a href="https://www.familysearch.org" target="_blank" rel="noopener">FamilySearch</a> (church &amp; court records, Full-Text Search, Digital Library; <i>free account</i>) ·
    <a href="https://www.ancestry.com" target="_blank" rel="noopener">Ancestry</a> (PA tax &amp; church collections; <i>subscription</i>) ·
    <a href="https://www.findagrave.com" target="_blank" rel="noopener">FindAGrave</a> ·
    <a href="https://brian-hamman.com" target="_blank" rel="noopener">Welty Y-DNA project</a> (Brian Hamman, admin) ·
    plus published works (Strassburger &amp; Hinke <i>Pennsylvania German Pioneers</i>; Bowen; Beers), <a href="https://archive.org" target="_blank" rel="noopener">archive.org</a> book scans, and the compiler's family papers.</p>
  </div>
</details>

__HEAD__
__ROWS__

<div class="foot">
  <p><b>Why two families here:</b> both lines trace to the Palatinate (Edenkoben) but carry different Y-DNA &mdash; <b>R1b</b> = the Edenkoben line, <b>I1</b> = the Manchester line (raised in the same W&auml;lti household; non-paternity event). Same name, different fathers.</p>
  <p style="margin-top:8px;color:#9a9384">Auto-generated from Welty Ancestry Research Log.xlsx &rarr; &ldquo;People Roster (chart source)&rdquo; (Proof record + Source columns) via generate_chart.py. Edit the roster and re-run to update.</p>
</div>
</div>
</body>
</html>
"""

# ---------------- GERMAN-LINES GRAPHICAL CHART (node-and-line, pan/zoom) ----------------
def render_graph(outfile, payload, count_label):
    data_json = json.dumps(payload, ensure_ascii=False)
    out = (GRAPH_TEMPLATE
           .replace("/*DATA*/", data_json)
           .replace("__COUNTLABEL__", count_label))
    with open(outfile, "w", encoding="utf-8") as fh:
        fh.write(out)

GRAPH_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>The German Welty Lines — Graphical Family Tree</title>
<style>
  :root{--ink:#1c1a17;--muted:#6b6459;--line:#c9bfae;--bg:#f6f1e7;--card:#fffdf8;
    --proven:#2e7d32;--documented:#1565c0;--hypo:#b26a00;--lore:#8e24aa;--living:#00695c;
    --direct:#b71c1c;--disputed:#6d4c41;--eden:#b71c1c;--manch:#3a4a5e}
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);
    font-family:"Iowan Old Style","Palatino Linotype",Palatino,Georgia,serif;line-height:1.35;padding:18px 18px 24px}
  h1{font-size:24px;margin:0 0 3px}
  .sub{color:var(--muted);font-size:13px;margin:0 0 10px;max-width:1050px}
  .controls{display:flex;flex-wrap:wrap;gap:8px 14px;align-items:center;margin-bottom:8px}
  .search{font:inherit;font-size:13.5px;padding:6px 11px;border:1px solid var(--line);border-radius:8px;min-width:240px;background:var(--card);color:var(--ink)}
  .btn{font:inherit;font-size:13px;padding:5px 11px;border:1px solid var(--line);border-radius:8px;background:var(--card);cursor:pointer;color:#4a4238}
  .btn:hover{background:#efe8db}.btn.on{background:#4a4238;color:#fff;border-color:#4a4238}
  .chk{font-size:13px;display:inline-flex;gap:5px;align-items:center;cursor:pointer;padding:4px 10px;border:1px solid var(--line);border-radius:20px;background:var(--card);user-select:none}
  .chk input{accent-color:#7a6a4a}
  .chk.eden{border-color:#e0a9a2}.chk.manch{border-color:#a9bbd0}
  .count{font-size:12px;color:var(--muted)}
  .legend{font-size:11.5px;color:var(--muted);display:flex;flex-wrap:wrap;gap:5px 13px;align-items:center;margin-bottom:8px}
  .tag{display:inline-block;font-size:9px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;padding:1px 6px;border-radius:20px;color:#fff}
  .t-proven{background:var(--proven)}.t-documented{background:var(--documented)}.t-hypo{background:var(--hypo)}
  .t-lore{background:var(--lore)}.t-living{background:var(--living)}.t-dna{background:#37474f}
  .t-direct{background:var(--direct)}.t-kit{background:#b8860b}.t-disputed{background:var(--disputed)}

  #viewport{position:relative;height:78vh;border:1px solid var(--line);border-radius:12px;
    background:var(--card);overflow:hidden;cursor:grab;touch-action:none}
  #viewport.panning{cursor:grabbing}
  #world{position:absolute;left:0;top:0;transform-origin:0 0}
  #edges{position:absolute;left:0;top:0;overflow:visible}
  #nodes{position:absolute;left:0;top:0}
  .gnode{position:absolute;width:190px;height:106px;overflow:hidden;background:var(--card);
    border:1.5px solid var(--line);border-radius:10px;padding:7px 10px 6px;cursor:pointer;box-shadow:0 1px 3px #0002}
  .gnode:hover{box-shadow:0 2px 8px #0004}
  .gnode.eden{border-color:#d99d96}.gnode.manch{border-color:#9fb2c9}
  .gnode.direct{border-color:var(--direct);border-width:2px;background:#fffaf8}
  .gnode.kit{border-color:#b8860b;border-width:2.5px;background:#fffbe9;box-shadow:0 0 0 3px #f0dd9e}
  .gnode.disputed{border-style:dashed;border-color:var(--disputed)}
  .gnode.sel{outline:3px solid #7a6a4a}
  .gn{font-weight:700;font-size:13px;line-height:1.2;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
  .gnode.direct .gn{color:var(--direct)}.gnode.kit .gn{color:#8a6100}
  .gy{font-size:11px;color:#5a4632;margin-top:2px}
  .gp{font-size:10px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:1px}
  .gb{margin-top:4px;display:flex;flex-wrap:wrap;gap:3px;align-items:center}
  .genchip{display:inline-block;font-size:9px;font-weight:700;color:#4a4238;background:#eae2d2;border-radius:20px;padding:0 6px}
  body.searching .gnode{opacity:.15}body.searching .gnode.hit{opacity:1}
  body.directonly .gnode:not(.direct):not(.kit){opacity:.13}
  body.directonly #edges path:not(.dpath){opacity:.14}
  .bandhead{position:absolute;font-weight:700;font-size:15px;letter-spacing:.3px;padding:4px 14px;border-radius:20px;color:#fff;white-space:nowrap}
  .bandhead.eden{background:var(--eden)}.bandhead.manch{background:var(--manch)}

  #panel{position:fixed;right:26px;top:130px;width:335px;max-height:68vh;overflow:auto;background:var(--card);
    border:1px solid var(--line);border-radius:12px;box-shadow:0 6px 24px #0003;padding:14px 16px;z-index:50;font-size:13px}
  #panel h3{margin:0 24px 4px 0;font-size:17px}
  #panel .pm{color:#4a4238;font-size:12.5px;margin:3px 0}
  #panel .pm b{color:#5a4632}
  #panel .pn-notes{font-style:italic;color:#4a4238;font-size:12px;margin-top:8px;border-top:1px dashed var(--line);padding-top:7px}
  #panel .plinks{margin-top:8px;font-size:12px;line-height:1.6}
  #panel a{color:#1565c0;cursor:pointer;text-decoration:underline}
  #pclose{position:absolute;right:8px;top:6px;border:none;background:none;font-size:18px;cursor:pointer;color:var(--muted)}
  .hidden{display:none !important}
  .foot{margin-top:10px;font-size:11.5px;color:var(--muted)}
</style>
</head>
<body>
<h1>The German Welty lines &mdash; graphical tree</h1>
<p class="sub">Node-and-branch chart of the two Palatinate-connected Welty families &mdash;
<b style="color:var(--eden)">Edenkoben (R1b)</b> and <b style="color:var(--manch)">Manchester (I1)</b>.
The Swiss Emmental line is deliberately excluded. <b>Drag</b> to pan, <b>scroll</b> to zoom, <b>click</b> any person for full details.
Solid lines = proven/documented parentage; long dashes = hypothesized; short dashes = disputed (non-paternity).</p>
<div class="controls">
  <input id="q" class="search" type="search" placeholder="Find a person… (Enter = next match)" autocomplete="off">
  <label class="chk eden"><input type="checkbox" data-fam="Eden" checked> Edenkoben</label>
  <label class="chk manch"><input type="checkbox" data-fam="Manch" checked> Manchester branch (I1)</label>
  <button id="directBtn" class="btn">Direct line only</button>
  <button id="zin" class="btn">+</button>
  <button id="zout" class="btn">&minus;</button>
  <button id="fitBtn" class="btn">Fit all</button>
  <span id="count" class="count">__COUNTLABEL__</span>
</div>
<div class="legend">
  <span style="font-weight:600">Key:</span>
  <span><span class="tag t-proven">Proven</span></span><span><span class="tag t-documented">Documented</span></span>
  <span><span class="tag t-hypo">Hypothesized</span></span><span><span class="tag t-lore">Lore</span></span>
  <span><span class="tag t-living">Living</span></span><span><span class="tag t-dna">DNA</span> Y-tester</span>
  <span style="color:var(--direct);font-weight:600">Red = the proven direct paternal Welty line</span>
  <span><span class="tag t-kit">Kit</span> living kit owner</span>
  <span><span class="tag t-disputed">Disputed &#8265;</span> dashed border = non-paternity</span>
</div>
<div id="viewport"><div id="world"><svg id="edges" xmlns="http://www.w3.org/2000/svg"></svg><div id="nodes"></div></div></div>
<div id="panel" class="hidden"><button id="pclose">&times;</button><div id="pbody"></div></div>
<div class="foot">Auto-generated from <b>Welty Ancestry Research Log.xlsx</b> &rarr; sheet &ldquo;People Roster (chart source)&rdquo; via generate_chart.py. Edit the roster and re-run to update.</div>

<script>
const DATA = /*DATA*/;
const P = DATA.people;
const NODE_W=190,NODE_H=106,HGAP=24,VGAP=64,PITCH=NODE_W+HGAP,ROWP=NODE_H+VGAP,FAMGAP=180,TOP=64;
const vp=document.getElementById('viewport'),world=document.getElementById('world'),
      svg=document.getElementById('edges'),nodesEl=document.getElementById('nodes');
const famOn={Eden:true,Manch:true};
let bounds={w:400,h:400},selId=null;

function esc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function proofTag(pr){const m={proven:['t-proven','Proven'],documented:['t-documented','Documented'],
  hypo:['t-hypo','Hypothesized'],lore:['t-lore','Lore'],living:['t-living','Living'],disputed:['t-disputed','Disputed ⁉']};
  const t=m[(pr||'').toLowerCase()];return t?`<span class="tag ${t[0]}">${t[1]}</span>`:'';}

function build(){
  nodesEl.innerHTML='';svg.innerHTML='';
  const gens=Object.values(P).filter(p=>famOn[p.fam]).map(p=>p.gen).filter(g=>g!==null&&g!==undefined);
  const minGen=gens.length?Math.min(...gens):0;
  let cursor=40,maxRow=0;
  const vis=[];
  function layout(id,prow){
    const p=P[id];
    const row=(p.gen!==null&&p.gen!==undefined)?(p.gen-minGen):(prow+1);
    p._row=row;maxRow=Math.max(maxRow,row);
    const kids=p.kids||[];
    if(!kids.length){p._x=cursor;cursor+=PITCH;}
    else{kids.forEach(k=>layout(k,row));p._x=(P[kids[0]]._x+P[kids[kids.length-1]]._x)/2;}
    vis.push(id);
  }
  const bands=[];
  DATA.families.forEach(f=>{
    if(!famOn[f.key])return;
    const x0=cursor;
    (f.roots||[]).forEach(r=>layout(r,-1));
    bands.push({key:f.key,title:f.title,hap:f.hap,x0:x0,x1:cursor-HGAP});
    cursor+=FAMGAP;
  });
  bounds={w:Math.max(cursor-FAMGAP+40,400),h:TOP+(maxRow+1)*ROWP+40};
  svg.setAttribute('width',bounds.w);svg.setAttribute('height',bounds.h);
  world.style.width=bounds.w+'px';world.style.height=bounds.h+'px';
  bands.forEach(b=>{
    const d=document.createElement('div');d.className='bandhead '+b.key.toLowerCase();
    d.style.left=Math.max(b.x0,(b.x0+b.x1)/2-130)+'px';d.style.top='12px';
    d.textContent=b.title+'  ·  Y-DNA '+b.hap;nodesEl.appendChild(d);
  });
  const NS='http://www.w3.org/2000/svg';
  vis.forEach(id=>{
    const p=P[id];
    if(!p.dad||!P[p.dad]||!famOn[P[p.dad].fam])return;
    const f=P[p.dad];
    const x1=f._x+NODE_W/2,y1=TOP+f._row*ROWP+NODE_H,x2=p._x+NODE_W/2,y2=TOP+p._row*ROWP;
    const my=(y1+y2)/2;
    const path=document.createElementNS(NS,'path');
    path.setAttribute('d','M'+x1+','+y1+' V'+my+' H'+x2+' V'+y2);
    path.setAttribute('fill','none');
    const direct=!!(f.direct&&p.direct);
    path.setAttribute('stroke',direct?'#b71c1c':'#b6ab93');
    path.setAttribute('stroke-width',direct?3:1.6);
    const pr=(p.proof||'').toLowerCase();
    if(pr==='disputed')path.setAttribute('stroke-dasharray','3 5');
    else if(pr==='hypo')path.setAttribute('stroke-dasharray','9 6');
    if(direct)path.setAttribute('class','dpath');
    svg.appendChild(path);
  });
  vis.forEach(id=>{
    const p=P[id];
    const d=document.createElement('div');
    let cls='gnode '+p.fam.toLowerCase();
    if(p.direct==='yes')cls+=' direct';if(p.direct==='kit')cls+=' kit';
    if((p.proof||'').toLowerCase()==='disputed')cls+=' disputed';
    d.className=cls;d.dataset.id=id;
    d.style.left=p._x+'px';d.style.top=(TOP+p._row*ROWP)+'px';
    const yrs=(p.birth||p.death)?((p.birth||'?')+' – '+(p.death||'')):'';
    let badges='';
    if(p.gen!==null&&p.gen!==undefined)badges+='<span class="genchip">Gen '+p.gen+'</span>';
    badges+=proofTag(p.proof);
    if(p.dna)badges+='<span class="tag '+(p.direct==='kit'?'t-kit':'t-dna')+'">DNA</span>';
    d.innerHTML='<div class="gn">'+esc(p.name)+'</div>'
      +(yrs?'<div class="gy">'+esc(yrs)+'</div>':'')
      +(p.place?'<div class="gp" title="'+esc(p.place)+'">'+esc(p.place)+'</div>':'')
      +(p.spouse?'<div class="gp" title="'+esc(p.spouse)+'">m. '+esc(p.spouse)+'</div>':'')
      +'<div class="gb">'+badges+'</div>';
    nodesEl.appendChild(d);
  });
  applySearch();
}

/* ---- pan & zoom ---- */
let view={x:20,y:10,s:1},drag=null,moved=false;
function applyView(){world.style.transform='translate('+view.x+'px,'+view.y+'px) scale('+view.s+')';}
vp.addEventListener('pointerdown',e=>{drag={x:e.clientX,y:e.clientY,vx:view.x,vy:view.y};moved=false;
  vp.classList.add('panning');vp.setPointerCapture(e.pointerId);});
vp.addEventListener('pointermove',e=>{if(!drag)return;
  const dx=e.clientX-drag.x,dy=e.clientY-drag.y;
  if(Math.abs(dx)+Math.abs(dy)>4)moved=true;
  view.x=drag.vx+dx;view.y=drag.vy+dy;applyView();});
vp.addEventListener('pointerup',()=>{vp.classList.remove('panning');drag=null;});
vp.addEventListener('wheel',e=>{e.preventDefault();zoomAt(e.clientX,e.clientY,e.deltaY<0?1.15:1/1.15);},{passive:false});
function zoomAt(cx,cy,f){const r=vp.getBoundingClientRect();
  const px=cx-r.left,py=cy-r.top;
  const ns=Math.min(Math.max(view.s*f,0.03),3);f=ns/view.s;
  view.x=px-(px-view.x)*f;view.y=py-(py-view.y)*f;view.s=ns;applyView();}
document.getElementById('zin').onclick=()=>{const r=vp.getBoundingClientRect();zoomAt(r.left+r.width/2,r.top+r.height/2,1.3);};
document.getElementById('zout').onclick=()=>{const r=vp.getBoundingClientRect();zoomAt(r.left+r.width/2,r.top+r.height/2,1/1.3);};
function fit(){const r=vp.getBoundingClientRect();
  const s=Math.min(r.width/bounds.w,r.height/bounds.h);
  view.s=Math.min(Math.max(s,0.03),1.2);
  view.x=(r.width-bounds.w*view.s)/2;view.y=Math.max((r.height-bounds.h*view.s)/2,8);applyView();}
document.getElementById('fitBtn').onclick=fit;
function centerOn(id){const p=P[id];if(p._x===undefined)return;
  const r=vp.getBoundingClientRect();const s=Math.max(view.s,0.85);view.s=s;
  view.x=r.width/2-(p._x+NODE_W/2)*s;view.y=r.height/2-(TOP+p._row*ROWP+NODE_H/2)*s;applyView();}

/* ---- select & detail panel ---- */
const panel=document.getElementById('panel'),pbody=document.getElementById('pbody');
vp.addEventListener('click',e=>{if(moved)return;const n=e.target.closest('.gnode');if(n)select(n.dataset.id);});
document.getElementById('pclose').onclick=()=>{panel.classList.add('hidden');
  document.querySelectorAll('.gnode.sel').forEach(n=>n.classList.remove('sel'));selId=null;};
function select(id){
  selId=id;
  document.querySelectorAll('.gnode.sel').forEach(n=>n.classList.remove('sel'));
  const el=document.querySelector('.gnode[data-id="'+id+'"]');if(el)el.classList.add('sel');
  const p=P[id];
  const rows=[];
  if(p.birth)rows.push('<div class="pm"><b>Born:</b> '+esc(p.birth)+'</div>');
  if(p.death)rows.push('<div class="pm"><b>Died:</b> '+esc(p.death)+'</div>');
  if(p.place)rows.push('<div class="pm"><b>Place:</b> '+esc(p.place)+'</div>');
  if(p.spouse)rows.push('<div class="pm"><b>Spouse:</b> '+esc(p.spouse)+'</div>');
  if(p.dna)rows.push('<div class="pm"><b>DNA:</b> '+esc(p.dna)+'</div>');
  if(p.proofrec)rows.push('<div class="pm"><b>Proof record:</b> '+esc(p.proofrec)+'</div>');
  if(p.source)rows.push('<div class="pm"><b>Source:</b> '+esc(p.source)+'</div>');
  let badges=proofTag(p.proof);
  if(p.direct==='yes')badges+='<span class="tag t-direct">Direct line</span>';
  if(p.direct==='kit')badges+='<span class="tag t-kit">Kit owner</span>';
  const fam=DATA.families.find(f=>f.key===p.fam);
  let links='';
  if(p.dad&&P[p.dad])links+='<div><b>Father:</b> <a data-go="'+p.dad+'">'+esc(P[p.dad].name)+'</a></div>';
  if((p.kids||[]).length)links+='<div><b>Children:</b> '+p.kids.map(k=>'<a data-go="'+k+'">'+esc(P[k].name)+'</a>').join(' · ')+'</div>';
  pbody.innerHTML='<h3>'+esc(p.name)+'</h3>'
    +'<div class="pm">'+(fam?esc(fam.title):'')+((p.gen!==null&&p.gen!==undefined)?' · Gen '+p.gen:'')+'</div>'
    +'<div class="gb" style="margin:5px 0">'+badges+'</div>'+rows.join('')
    +(p.notes?'<div class="pn-notes">'+esc(p.notes)+'</div>':'')
    +(links?'<div class="plinks">'+links+'</div>':'');
  panel.classList.remove('hidden');
}
pbody.addEventListener('click',e=>{const a=e.target.closest('a[data-go]');
  if(a){select(a.dataset.go);centerOn(a.dataset.go);}});

/* ---- search ---- */
const q=document.getElementById('q');let matches=[],mi=-1;
function applySearch(){
  const term=q.value.trim().toLowerCase();
  matches=[];mi=-1;
  document.body.classList.toggle('searching',!!term);
  document.querySelectorAll('.gnode').forEach(n=>{
    const p=P[n.dataset.id];if(!p)return;
    const hit=!!(term&&((p.name+' '+(p.place||'')+' '+(p.spouse||'')+' '+(p.notes||'')+' '+(p.dna||'')).toLowerCase().includes(term)));
    n.classList.toggle('hit',hit);
    if(hit)matches.push(n.dataset.id);
  });
  document.getElementById('count').textContent =
    term?(matches.length+' match'+(matches.length===1?'':'es')+' — Enter to step through'):'__COUNTLABEL__';
}
q.addEventListener('input',applySearch);
q.addEventListener('keydown',e=>{
  if(e.key==='Enter'&&matches.length){
    mi=(mi+1)%matches.length;const id=matches[mi];
    centerOn(id);select(id);
    document.getElementById('count').textContent=(mi+1)+' / '+matches.length+' matches';
  }});

/* ---- toggles ---- */
const directBtn=document.getElementById('directBtn');
directBtn.onclick=()=>{document.body.classList.toggle('directonly');directBtn.classList.toggle('on');};
document.querySelectorAll('.chk input').forEach(c=>c.addEventListener('change',()=>{
  famOn[c.dataset.fam]=c.checked;
  panel.classList.add('hidden');
  build();fit();
}));

build();fit();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
