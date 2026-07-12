#!/usr/bin/env python3
"""
build_records.py — wire the Record Images/ backlog onto the family tree.

For every file in  Record Images/  this script:
  1. converts it to a web JPG (PDFs via pdftoppm) at two sizes ->
        site/records/full/<slug>.jpg   (max 1600px)
        site/records/thumb/<slug>.jpg  (max 360px)
  2. reads the curated per-file MAPPING below (which person[s] on the tree
     the record belongs to, a caption, a confidence flag)
  3. pulls the repository + external URL from  Record Images/Capture Manifest.xlsx
  4. writes  site/records/records.json   -> { PersonID: [ {slug,caption,repo,url,logids,confidence}, ... ] }
     plus an "_evidence" bucket for records that document a decoy / negative
     sweep / separate cluster and MUST NOT be pinned to a direct-line person.
  5. mirrors the whole map into a readable "Tree Image Map" sheet in the
     Capture Manifest workbook so the curation is durable and reviewable.

generate_chart.py reads records.json at publish time and renders a thumbnail
chip in each person's Source line.

Curation rule (Welty conventions): decoys, negative-sweep calibration pages,
and records that belong to a DIFFERENT family (the Newberry D7 Philip, the
Manheim John Welty cluster, the Donegal Michael) are marked show=False and
routed to _evidence, never silently attached to a direct ancestor.
"""
import os, re, json, subprocess, glob
from PIL import Image
import openpyxl

ROOT   = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
RECDIR = os.path.join(ROOT, "Record Images")
SITE   = os.path.join(os.path.dirname(__file__), "..", "site")
OUTF   = os.path.join(SITE, "records", "full")
OUTT   = os.path.join(SITE, "records", "thumb")
MANXL  = os.path.join(RECDIR, "Capture Manifest.xlsx")
FULL_MAX, THUMB_MAX = 1600, 360

# ---------------------------------------------------------------------------
# CURATED PER-FILE MAP  (basename -> persons / caption / confidence)
#   persons: roster PersonIDs the record is pinned to. [] => evidence-only.
#   conf:    high | med | review   (review = attribution wants Kwyn's eye)
# ---------------------------------------------------------------------------
M = {
 # --- Edenkoben Reformed registers (Archion) ------------------------------
 "FB57 — Georg Wolffgang Welde marriage entry (Archion reg 54927 Edenkoben Taufen Bild 56).pdf":
   (["E-georgwolfgang-disp"], "Georg Wolfgang Welde's marriage entry, Edenkoben Reformed (Bild 56).", "high"),
 "FB57 — Georg Wolffgang Welde entry (Archion reg 54927 Edenkoben Taufen Bild 62).pdf":
   (["B-johanjacob"], "The 14 Jun 1750 Edenkoben baptism naming Georg Wolfgang Welde as father (Bild 62).", "high"),
 "P75 — Anna Elisabetha Welde baptism 31 Jan 1740 (Archion reg 54927 Edenkoben Taufen Bild 11).pdf":
   (["E-jjc-annaelis1740"], "Anna Elisabetha Welde's baptism, 31 Jan 1740, Edenkoben Reformed (Bild 11).", "high"),
 "P75 — Welde child baptism (Archion reg 54927 Edenkoben Taufen Bild 28).pdf":
   (["E-johnjacob1710"], "Baptism of a Welde child of John Jacob Welde (b.1710), Edenkoben Reformed (Bild 28).", "med"),
 "P75 — Welde child baptism 1745-46 (Archion reg 54927 Edenkoben Taufen Bild 32).pdf":
   (["E-johnjacob1710"], "Baptism of a Welde child, 1745–46, Edenkoben Reformed (Bild 32).", "med"),
 "P75 — Welde child baptism (Archion reg 54927 Edenkoben Taufen Bild 43).pdf":
   (["E-johnjacob1710"], "Baptism of a Welde child of John Jacob Welde (b.1710), Edenkoben Reformed (Bild 43).", "med"),
 "P74 + FB56 — Joh Jacob Welde x Anna Catharina Croissant marriage (Archion reg 54933 Edenkoben4 Bild 79).pdf":
   (["E-johnjacob1710"], "Marriage of Joh. Jacob Welde and Anna Catharina Croissant, Edenkoben Reformed (Bild 79).", "high"),
 "P64 — Edenkoben marriage register, bride's father read in full (Archion reg 54933 Edenkoben4 Bild 63).pdf":
   (["E-johnjacob1710"], "Edenkoben marriage register — the bride's father read in full (Bild 63).", "med"),
 "P80 — Hanß Jacob Croissant x A.M. Müllerin marriage (Archion reg 54933 Edenkoben4 Bild 66).pdf":
   (["E-johnjacob1710"], "Marriage of Hanß Jacob Croissant & A. M. Müllerin — the Croissant in-laws (Bild 66).", "med"),
 "P116 — Doll x Neu marriage (Archion reg 54933 Edenkoben4 Bild 93).pdf":
   ([], "Doll × Neu marriage — methodology / graft-source calibration record (Bild 93).", "high"),
 "P82 — Edenkoben marriage sweep (Archion reg 54933 Edenkoben4 Bild 80).pdf":
   ([], "Calibration page for the full Edenkoben marriage-register sweep (Bild 80).", "high"),
 # --- Archive.org (free) --------------------------------------------------
 "D15 — Welty Philip 40ac Northampton 1785 decoy (Archive.org 3rdPAarch v19 p141).png":
   ([], "DECOY — a Philip Welty, 40 ac Northampton Co. 1785; ruled out, not the Dover Philip.", "high"),
 "P13 + P147 — York Co 1780 tax return p203 (Welty Widow + Philip Welty + George Welty) (archive.org 3rdPAarch v21 leaf221).jpg":
   (["E-philipjacob"], "York Co. 1780 tax return — Welty Widow, Philip Welty and George Welty (PA Arch. 3rd ser. v21, p.203).", "high"),
 "P70 — Lydia 1749 shiplist p420 (List 142C start, Ship Lydia) (archive.org pennsylvaniagerm42stra leaf502).jpg":
   (["E-johnjacob1710"], "Ship Lydia 1749 passenger list, p.420 — the Croissant in-law crossing (List 142C).", "med"),
 "P70 — Lydia 1749 shiplist p421 (Johan Jacob Messer) (archive.org pennsylvaniagerm42stra leaf503).jpg":
   (["E-johnjacob1710"], "Ship Lydia 1749 passenger list, p.421.", "med"),
 "P70 — Lydia 1749 shiplist p422 (Jeorg Crassan = Croissant) (archive.org pennsylvaniagerm42stra leaf504).jpg":
   (["E-johnjacob1710"], "Ship Lydia 1749 list, p.422 — 'Jeorg Crassan' (Croissant), the in-law surname.", "med"),
 "US14 — Johannes Welte, Neptune 1751 shiplist p469 (archive.org pennsylvaniagerm42stra leaf553).jpg":
   ([], "UNSORTED — Johannes Welte, ship Neptune 1751; not yet tied to the line.", "high"),
 # --- FamilySearch --------------------------------------------------------
 "FA32 — Joseph Welty ~19 son of John Welty late of Manheim (FamilySearch ark 3QS7-899B-27YW).jpg":
   ([], "Separate cluster — Joseph Welty (~19), son of John Welty late of Manheim.", "high"),
 "FA33 — John Welty deceased testate - exec Christian Rubel (FamilySearch ark 3QS7-L99B-27Q1).jpg":
   ([], "Separate cluster — John Welty dec'd (testate), exec. Christian Rubel.", "high"),
 "FA34 — John Welty Sr will - wife Eve + 12-child heir list (FamilySearch ark 3QS7-899B-K845).jpg":
   ([], "Separate cluster — John Welty Sr. will, wife Eve + 12-child heir list.", "high"),
 "FB37 — Henry Welty x Coleman Manchester deed 1839 (FamilySearch York Deed Bks 1819-45 img76).png":
   (["B-henrySMGF"], "Henry Welty & Coleman, Manchester Twp deed, 1839 (York Deed Bks 1819–45, img 76).", "high"),
 "M141 — Daniel Welty household 1850 census, Marion Twp Owen Co IN (FamilySearch img291, ark S3HY-69YS-CP3).pdf":
   ([], "REVIEW — Daniel Welty household, 1850 census, Marion Twp, Owen Co. IN; which Daniel to confirm.", "review"),
 "M141 — Daniel Welty x Louisa Whaling marriage 5 Apr 1839, Tuscarawas Co OH (FamilySearch marriage reg entry 2748, ark 33S7-95BW-VCC).pdf":
   ([], "REVIEW — Daniel Welty × Louisa Whaling, 5 Apr 1839, Tuscarawas Co. OH; which Daniel to confirm.", "review"),
 "M20 — Michael Welty 1828 partition, heirs named (FamilySearch Tuscarawas Court Recs img96).png":
   (["E-michael"], "1828 partition of Michael Welty's estate, Tuscarawas Co. OH — the heirs named.", "high"),
 "M97 — Daniel Prosser to Michael Weldy Donegal deed 1797 (FamilySearch Westmoreland Deed Bks img156).png":
   ([], "REVIEW — Daniel Prosser to Michael Weldy, Donegal Twp deed 1797 (Westmoreland Co.).", "review"),
 "P110 — OC docket general index W-section Books D-K (zero Welty) (FamilySearch ark 3QSQ-G99B-VG4K).jpg":
   ([], "Negative sweep — Orphans' Court docket index, W-section, Books D–K: zero Welty.", "high"),
 "P111 — Gauff settlement motion (want of notice to guardians) (FamilySearch ark 3QS7-L99B-278P).jpg":
   (["E-elizabethgauf"], "Gauf estate settlement motion (want of notice to guardians).", "high"),
 "P111+P112 — George Gauff guardianships - Elizabeth + Margaret (FamilySearch ark 3QSQ-G99B-2W6X).jpg":
   (["E-elizabethgauf"], "George Gauf guardianships — daughters Elizabeth & Margaret.", "high"),
 "P111+P112 — OC docket general index G-section (Gauff lines) (FamilySearch ark 3QS7-L99B-VLVJ).jpg":
   (["E-elizabethgauf"], "Orphans' Court docket index, G-section — the Gauff lines.", "med"),
 "P112 — George Gauf admin account - Jacob Nigle + Catharine (FamilySearch ark 3QS7-L99B-2731).jpg":
   (["E-elizabethgauf"], "George Gauf administration account (Jacob Nigle + Catharine).", "med"),
 "P113 — Letters of administration memo - Catharine Gauff + Jacob Weigel (FamilySearch ark 3QS7-L99B-KZR3).jpg":
   (["E-elizabethgauf"], "Letters of administration — Catharine Gauff + Jacob Weigel.", "med"),
 "P114 — Jacob Lower guardianship account for Elizabeth Goff (FamilySearch ark 3QS7-L99B-2796).jpg":
   (["E-elizabethgauf"], "Jacob Lower's guardianship account for Elizabeth Goff (Gauf).", "med"),
 "P120 — QS continuance p142 Catharine bound 40 pounds (FamilySearch ark 3Q9M-CSVF-LTZL).jpg":
   (["E-elizabethgauf"], "Quarter Sessions continuance p.142 — Catharine bound £40 (Gauf matter).", "med"),
 "P120 — QS disposition case No.10 State v Catharine Gauff (FamilySearch ark 3Q9M-CSVF-LTNX).jpg":
   (["E-elizabethgauf"], "Quarter Sessions disposition, case No.10 — State v. Catharine Gauff.", "med"),
 "P120 — QS recognizances p133 vs Catharine Gauff (FamilySearch ark 3Q9M-CSVF-LTVG).jpg":
   (["E-elizabethgauf"], "Quarter Sessions recognizances p.133 — vs Catharine Gauff.", "med"),
 "P120 — QS cancelled duplicate entry (FamilySearch ark 3Q9M-CSVF-LTLK).jpg":
   ([], "Quarter Sessions cancelled duplicate entry — housekeeping image.", "high"),
 "P121 — Philip Gauf of Dover estate 868 pounds (FamilySearch ark 3QS7-L99B-2LD).jpg":
   (["E-elizabethgauf"], "Philip Gauf of Dover — estate inventory (£868); the elder Gauf household.", "med"),
 "P122 — 1774 warrant Dundohr 80ac Dover Twp (FamilySearch ark 3QHV-J385-19F7-V).jpg":
   (["E-philipjacob"], "1774 land warrant, 80 ac Dover Twp — adjoining-owner context for Philip Jacob Welty.", "med"),
 "P127 — Dover 1771 assessment continuation D-Z + Manheim warrant (FamilySearch ark 3Q9M-CSVF-Y9VN-B).jpg":
   (["E-philipjacob"], "Dover Twp 1771 assessment (D–Z) + Manheim warrant — Dover tax context.", "med"),
 "P127+P145 — Dover 1771-for-1772 assessment (verified zero-Welty + Peter warrant) (FamilySearch ark 3Q9M-CSVF-Y9VF-6).jpg":
   (["E-philipjacob"], "Dover Twp 1771-for-1772 assessment — verified zero-Welty calibration + Peter warrant.", "med"),
 "P143 — Dover 1775 Provincial Tax D-Z, Welty Philip + George (FamilySearch York tax film 008122326 img148, ark 3Q9M-CSVF-Y9VD-Y).jpg":
   (["E-philipjacob"], "Dover Twp 1775 Provincial Tax (D–Z) — Welty Philip & George.", "high"),
 "P143 — Dover 1775 Provincial Tax warrant 8 Dec 1774 + A-D (FamilySearch ark 3Q9M-CSVF-Y9VH-5).jpg":
   (["E-philipjacob"], "Dover Twp 1775 Provincial Tax — warrant of 8 Dec 1774 + A–D.", "med"),
 "P38 — Jacob Welday Sr 10 children, First Families of Ohio (FamilySearch fullText img1327).png":
   (["E-jacobsr"], "Jacob Welday Sr.'s ten children (First Families of Ohio record).", "high"),
 "P4 — Manchester Twp 1771-for-1772 assessment, W-column (FamilySearch film 008122326, ark 3Q9M-CSVF-Y9VF-J).jpg":
   (["E-georgwolfgang-disp"], "Manchester Twp 1771-for-1772 assessment — the W-column.", "high"),
 "P4 — Manchester assessment wrapper 1771-for-1772 (FamilySearch film 008122326, ark 3Q9M-CSVF-Y9VN-H).jpg":
   (["E-georgwolfgang-disp"], "Manchester Twp 1771-for-1772 assessment — cover wrapper.", "med"),
 "P4 — Manchester collectors warrant (To John Ferry Collector) (FamilySearch ark 3Q9M-CSVF-Y9VN-5).jpg":
   (["E-georgwolfgang-disp"], "Manchester Twp collector's warrant (to John Ferry, collector).", "med"),
 "P6 — Philip Welty of Newberry to Jacob Bare 1797, D7 firewall (FamilySearch York Deed Bks img200).png":
   ([], "FIREWALL — Philip Welty of Newberry Twp (the D7 Philip), a DIFFERENT man from the Dover Philip.", "high"),
 "TL55 - Welty Black Book pg3 (John Jacob 1710 x Christina Broff family register; FS memory 9VSM-X8K, uploader genehisthome).png":
   (["E-johnjacob1710"], "The 'Welty Black Book' family-register page for John Jacob Welde (b.1710). Note: the 'Broff' wife-name it carries is a debunked name-match.", "med"),
 "US19 — Obediah Weldy decoy, Chester tax 1768-71 (FamilySearch img109).png":
   ([], "DECOY — Obediah Weldy, Chester Co. tax 1768–71; not the line.", "high"),
}

# ---------------------------------------------------------------------------
# VITAL-ONLY POLICY (Kwyn, 11 Jul 2026): a record earns a chip on a person only
# if it's a vital event FOR THEM (baptism/birth, marriage, death/burial, will,
# partition/guardianship) or the keystone record that NAMES them where little
# else does. Everything else — in-laws, elder-Gauf & Catharine court records,
# adjoining-owner warrants, negative/calibration pages, and second views of a
# fact already shown — stays in the Record Images backlog as evidence-only.
# To promote a record onto a person, add its filename here.
KEEP_ON_TREE = {
 "P74 + FB56 — Joh Jacob Welde x Anna Catharina Croissant marriage (Archion reg 54933 Edenkoben4 Bild 79).pdf",           # marriage
 "FB57 — Georg Wolffgang Welde marriage entry (Archion reg 54927 Edenkoben Taufen Bild 56).pdf",                          # marriage
 "FB57 — Georg Wolffgang Welde entry (Archion reg 54927 Edenkoben Taufen Bild 62).pdf",                                   # 1750 baptism of Johan Jacob
 "P75 — Anna Elisabetha Welde baptism 31 Jan 1740 (Archion reg 54927 Edenkoben Taufen Bild 11).pdf",                      # baptism
 "M20 — Michael Welty 1828 partition, heirs named (FamilySearch Tuscarawas Court Recs img96).png",                        # partition
 "P111+P112 — George Gauff guardianships - Elizabeth + Margaret (FamilySearch ark 3QSQ-G99B-2W6X).jpg",                   # guardianship names her line
 "P38 — Jacob Welday Sr 10 children, First Families of Ohio (FamilySearch fullText img1327).png",                         # names his children
 "P13 + P147 — York Co 1780 tax return p203 (Welty Widow + Philip Welty + George Welty) (archive.org 3rdPAarch v21 leaf221).jpg",  # names Philip; resolves the 1780 Widow
 "FB37 — Henry Welty x Coleman Manchester deed 1839 (FamilySearch York Deed Bks 1819-45 img76).png",                      # keystone deed naming this Henry
}

def slugify(base):
    stem = os.path.splitext(base)[0]
    stem = stem.replace("+"," ").replace("×","x")
    stem = re.sub(r"\(.*?\)", "", stem)          # drop parenthetical repo refs
    stem = re.sub(r"[^A-Za-z0-9]+","-", stem).strip("-").lower()
    return stem[:70].strip("-")

def unique_slugs(files):
    """Deterministic slugs; siblings that collapse to the same base get -2/-3.
    First (sorted) file keeps the clean base name, so re-runs overwrite in place."""
    out={}; seen={}
    for f in files:
        b=slugify(os.path.basename(f))
        seen[b]=seen.get(b,0)+1
        out[f]= b if seen[b]==1 else f"{b}-{seen[b]}"
    return out

def load_manifest_urls():
    urls={}
    if not os.path.exists(MANXL): return urls
    wb=openpyxl.load_workbook(MANXL, read_only=True, data_only=True)
    ws=wb["Capture Manifest"]; rows=list(ws.iter_rows(values_only=True))
    h={c:i for i,c in enumerate(rows[0])}
    for r in rows[1:]:
        lid=r[h['Log ID']]
        if lid: urls[str(lid).strip()]=(r[h['Repository']], r[h['URL']])
    return urls

def to_jpg(src, slug):
    """Return a PIL image (first page for PDFs)."""
    if src.lower().endswith(".pdf"):
        import tempfile
        tmp=os.path.join(tempfile.gettempdir(), "welty_rec_tmp")
        subprocess.run(["pdftoppm","-jpeg","-r","150","-f","1","-l","1",src,tmp],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        cand=sorted(glob.glob(tmp+"*"))
        img=Image.open(cand[0]).convert("RGB")
        loaded=img.copy()
        img.close()
        for c in cand:
            try: os.remove(c)
            except OSError: pass
        return loaded
    return Image.open(src).convert("RGB")

def resized(img, mx):
    w,h=img.size
    if max(w,h)>mx:
        s=mx/max(w,h); img=img.resize((int(w*s),int(h*s)), Image.LANCZOS)
    return img

def main():
    os.makedirs(OUTF, exist_ok=True); os.makedirs(OUTT, exist_ok=True)
    urls=load_manifest_urls()
    files=sorted(f for f in glob.glob(os.path.join(RECDIR,"**","*"), recursive=True)
                 if os.path.isfile(f) and f.lower().endswith((".jpg",".jpeg",".png",".pdf")))
    slugmap=unique_slugs([f for f in files if os.path.basename(f) in M])
    by_person={}; flat_by_person={}; evidence=[]; rows_out=[]; unmapped=[]
    for f in files:
        base=os.path.basename(f)
        if base not in M:
            unmapped.append(base); continue
        persons, caption, conf = M[base]
        if base not in KEEP_ON_TREE:      # vital-only policy: demote to evidence
            persons = []
        logids=re.findall(r'(?:P|M|D|FB|FA|US)\d+|TL-?\d+', base)
        slug=slugmap[f]
        img=to_jpg(f, slug)
        resized(img, FULL_MAX).save(os.path.join(OUTF,slug+".jpg"),"JPEG",quality=82)
        resized(img, THUMB_MAX).save(os.path.join(OUTT,slug+".jpg"),"JPEG",quality=80)
        repo,url = urls.get(logids[0], (None,None)) if logids else (None,None)
        plid = logids[0] if logids else slug   # grouping key: pages of one document share a log-ID
        rec={"slug":slug,"caption":caption,"logids":logids,"repo":repo,"url":url,
             "confidence":conf,"_plid":plid}
        rows_out.append((base, ";".join(persons), ";".join(logids), caption, "yes" if persons else "no", conf, repo or "", url or ""))
        if persons:
            for p in persons: flat_by_person.setdefault(p,[]).append(rec)
        else:
            evidence.append({k:v for k,v in rec.items() if k!="_plid"} | {"file":base})

    # Collapse multiple pages of ONE document (same primary log-ID) into a single
    # chip per person, so the tree shows one thumbnail per record instead of a row
    # of near-identical manuscript pages. The extra pages ride along in `pages[]`
    # and the lightbox pages through them.
    for pid, recs in flat_by_person.items():
        groups={}
        for r in recs:
            groups.setdefault(r["_plid"], []).append(r)
        out_list=[]
        for k, items in groups.items():
            items=sorted(items, key=lambda r:r["slug"])
            head=items[0]
            out_list.append({
                "slug":head["slug"], "caption":head["caption"],
                "repo":head["repo"], "url":head["url"],
                "confidence":head["confidence"], "logids":head["logids"],
                "pages":[{"slug":r["slug"],"caption":r["caption"]} for r in items],
            })
        by_person[pid]=out_list
    chips=sum(len(v) for v in by_person.values())
    imgs_on_people=sum(len(g["pages"]) for v in by_person.values() for g in v)
    out={"by_person":by_person,"_evidence":evidence,
         "_meta":{"total_files":len(files),"chips":chips,"images_on_people":imgs_on_people,
                  "evidence_only":len(evidence),"unmapped":unmapped}}
    with open(os.path.join(SITE,"records","records.json"),"w",encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    # mirror into Capture Manifest workbook (durable, reviewable)
    wb=openpyxl.load_workbook(MANXL)
    if "Tree Image Map" in wb.sheetnames: del wb["Tree Image Map"]
    ws=wb.create_sheet("Tree Image Map")
    ws.append(["File","PersonIDs","LogIDs","Caption","Show on tree","Confidence","Repository","URL"])
    for row in sorted(rows_out): ws.append(list(row))
    wb.save(MANXL)
    print(f"files={len(files)} mapped={len(rows_out)} unmapped={unmapped}")
    print(f"chips={chips} (from {imgs_on_people} images) on {len(by_person)} people; evidence_only={len(evidence)}")

if __name__=="__main__":
    main()
