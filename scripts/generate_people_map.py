#!/usr/bin/env python3
"""Generate the 'Everyone, Gen 1-6' Welty people map (Leaflet, site style).

Reads the People Roster (chart source) sheet of the master log, plots every
placed Eden-family person in Gen 1-6 at their primary node, colored by branch
and styled by proof grade. Household members with a blank Place inherit their
father's node (flagged 'residence with parent'). Two maps: the Palatinate
(German stayers) and America (everyone who crossed + their descendants).

Outputs:
  1. A standalone page:  ../../Welty Gen 1-6 People Map.html
  2. Injected section at the TOP of the timeline (geography) page, between
     <!-- PEOPLEMAP:START --> ... <!-- PEOPLEMAP:END --> markers, in both
     the project copy and the Netlify Upload copy.

No personal names appear on the page (the direct line is labelled generically).
Run:  python3 generate_people_map.py
"""
import openpyxl, json, os, shutil
import re

# Strip internal lead/source codes (P60, TL-34, M141, FB15 …) from the display
# fields before they ship in the embedded people JSON. Mirrors scrub_display() in
# generate_chart.py, plus em-dash + "verify CODE" cleanup for annotated fields.
_DISPLAY_CODE = r'(?:P|M|D|FB|FA|SR|US|H|FT|PS|DL|TL)-?\d+[a-z]?'
def scrub_display(s):
    if not s:
        return s
    s = str(s)
    s = re.sub(r'\s*\[[^\]]*\]', '', s)                                   # [..] codes
    s = re.sub(r',?\s*\b(?:see|verify|recheck|confirm|check|cf\.?|per)\s+' + _DISPLAY_CODE + r'\b', '', s, flags=re.I)
    s = re.sub(r'[;,]?\s*(?:—\s*)?\b' + _DISPLAY_CODE + r'\b', '', s)  # standalone codes (incl leading em-dash)
    s = re.sub(r'\s*—\s*([;,)])', r'\1', s)                           # "7 Sep — ;" -> "7 Sep;"
    s = re.sub(r'\(\s*—\s*', '(', s)
    s = re.sub(r'\(\s*[;:,]\s*', '(', s)
    s = re.sub(r'[;:,]\s*\)', ')', s)
    s = re.sub(r'\(\s*\)', '', s)
    s = re.sub(r'\s+([;,.)])', r'\1', s)
    s = re.sub(r'\(\s+', '(', s)
    s = re.sub(r'\s{2,}', ' ', s)
    return s.strip().strip(' ,;—')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
LOG = os.path.join(ROOT, "Welty Ancestry Research Log.xlsx")
OUT = os.path.join(ROOT, "Welty Gen 1-6 People Map.html")
TIMELINE = os.path.join(ROOT, "Welty Family Timeline (Geography).html")
TIMELINE_NETLIFY = os.path.join(ROOT, "Netlify Upload", "timeline.html")

# ---------------------------------------------------------------- gazetteer
# node-key -> (lat, lon, label, continent)  continent: DE or US
NODES = {
    "bischheim":   (49.688, 8.046,  "Bischheim (Donnersberg)", "DE"),
    "edenkoben":   (49.283, 8.128,  "Edenkoben", "DE"),
    "philadelphia":(39.953,-75.165, "Philadelphia — the landings", "US"),
    "chester":     (40.150,-75.650, "Chester Co (East Vincent / Brownback’s)", "US"),
    "lancaster":   (40.038,-76.305, "Chester–Lancaster border", "US"),
    "dover":       (40.001,-76.850, "Dover Twp, York Co", "US"),
    "manchester":  (40.030,-76.745, "Manchester Twp, York Co", "US"),
    "york":        (39.963,-76.727, "York borough (Trinity Reformed)", "US"),
    "conewago":    (40.052,-76.870, "Conewago Twp, York Co (Quickel’s)", "US"),
    "bedford":     (40.018,-78.503, "Milford Twp, Bedford Co", "US"),
    "somerset":    (39.970,-79.100, "Somerset Co (Milford Twp)", "US"),
    "fayette":     (40.017,-79.590, "Fayette Co (Connellsville)", "US"),
    "westmoreland":(40.175,-79.545, "Mt. Pleasant Twp, Westmoreland Co", "US"),
    "greensburg":  (40.301,-79.539, "Greensburg, Westmoreland Co", "US"),
    "crosscreek":  (40.325,-80.655, "Cross Creek, Jefferson Co OH", "US"),
    "steubenville":(40.370,-80.633, "Steubenville, Jefferson Co OH", "US"),
    "crookedrun":  (40.520,-81.474, "Crooked Run, Tuscarawas Co OH", "US"),
    "waynetusc":   (40.455,-81.470, "Wayne Twp, Tuscarawas Co OH", "US"),
    "sugarcreek":  (40.500,-81.640, "Sugarcreek Twp, Tuscarawas Co OH", "US"),
    "wilmot":      (40.657,-81.634, "Wilmot / Sugar Creek, Stark Co OH", "US"),
    "starkco":     (40.700,-81.550, "Stark Co OH", "US"),
    "massillon":   (40.797,-81.521, "Massillon area, Stark Co OH", "US"),
    "navarre":     (40.724,-81.523, "Navarre / Justus, Stark Co OH", "US"),
    "wyandot":     (40.819,-83.130, "Nevada, Wyandot Co OH", "US"),
    "hicksville":  (41.293,-84.762, "Hicksville, Defiance Co OH", "US"),
    "bridgeport":  (40.068,-80.740, "Bridgeport, Belmont Co OH", "US"),
    "fairfield_ia":(41.007,-91.962, "Fairfield, Jefferson Co IA", "US"),
    "crawfordsville":(40.041,-86.874,"Crawfordsville, IN", "US"),
    "escondido":   (33.119,-117.086,"Escondido, San Diego Co CA", "US"),
    "elpaso_co":   (38.831,-104.821,"El Paso Co CO (Monument)", "US"),
}

def node_for(pid, place, father_node):
    if not place:
        return father_node, True
    p = place.lower()
    tests = [
        ("escondido", "escondido"), ("oak hill", "escondido"),
        ("hicksville", "hicksville"),
        ("nevada, wyandot", "wyandot"), ("wyandot", "wyandot"),
        ("bridgeport", "bridgeport"),
        ("fairfield, jefferson co ia", "fairfield_ia"),
        ("crawfordsville", "crawfordsville"),
        ("el paso co co", "elpaso_co"), ("colorado springs", "elpaso_co"),
        ("navarre", "navarre"), ("justus", "navarre"),
        ("massillon", "massillon"),
        ("wilmot", "wilmot"),
        ("sugarcreek twp oh", "sugarcreek"),
        ("wayne twp, tuscarawas", "waynetusc"),
        ("crooked run", "crookedrun"),
        ("tuscarawas co oh", "crookedrun"),
        ("steubenville", "steubenville"),
        ("cross creek", "crosscreek"), ("jefferson co oh", "crosscreek"),
        ("greensburg", "greensburg"),
        ("mt. pleasant twp, westmoreland", "westmoreland"),
        ("weltytown", "westmoreland"),
        ("stark co oh", "starkco"),
        ("milford, bedford", "bedford"),
        ("conewago", "conewago"),
        ("dover/manchester", "manchester"), ("manchester twp", "manchester"),
        ("dover twp", "dover"),
        ("chester/lancaster", "lancaster"), ("lancaster co", "lancaster"),
        ("chester co pa", "chester"),
        ("york co pa", "york"),
        ("edenkoben", "edenkoben"),
        ("bischheim", "bischheim"),
    ]
    for needle, node in tests:
        if needle in p:
            return node, False
    return None, False

# ---------------------------------------------------------------- branches
BRANCHES = {
    "trunk":  ("Old Country trunk (apex → the three brothers)", "#a07a35"),
    "pj":     ("Philip Jacob → Michael, Cross Creek, Dover", "#8a5a2b"),
    "jj":     ("John Jacob → Weltytown (Westmoreland)", "#5b6e3a"),
    "gw":     ("Georg Wolfgang → Manchester", "#46618a"),
}
DIRECT_IDS = set()

# node-key -> chapter anchor id in the timeline page (for the "read the chapter" link)
CHAPTER = {
    "bischheim":"pfalz", "edenkoben":"pfalz",
    "philadelphia":"crossing",
    "chester":"chester", "lancaster":"chester",
    "dover":"york", "manchester":"york", "york":"york", "conewago":"york",
    "bedford":"westpa", "somerset":"westpa", "fayette":"westpa",
    "westmoreland":"westpa", "greensburg":"westpa",
    "crosscreek":"ohio", "steubenville":"ohio", "crookedrun":"ohio",
    "waynetusc":"ohio", "sugarcreek":"ohio", "wilmot":"ohio", "starkco":"ohio",
    "massillon":"ohio", "navarre":"ohio", "wyandot":"ohio", "hicksville":"ohio",
    "bridgeport":"ohio",
    "fairfield_ia":"offshoots", "crawfordsville":"offshoots",
    "escondido":"offshoots", "elpaso_co":"offshoots",
}

def load():
    wb = openpyxl.load_workbook(LOG, read_only=True)
    ws = wb["People Roster (chart source)"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[1]; idx = {h:i for i,h in enumerate(hdr)}
    people = {}
    for r in rows[2:]:
        if not r[0]:
            continue
        try: g = int(r[idx["Gen"]])
        except (TypeError, ValueError): continue
        if g > 6 or r[idx["Family"]] != "Eden":
            continue
        people[r[idx["PersonID"]]] = {
            "id": r[idx["PersonID"]], "gen": g, "name": r[idx["Name"]],
            "sex": r[idx["Sex"]], "birth": r[idx["Birth"]], "death": r[idx["Death"]],
            "place": r[idx["Place"]], "spouse": r[idx["Spouse"]],
            "father": r[idx["FatherID"]], "proof": (r[idx["Proof"]] or "").strip(),
            "dna": r[idx["DNAkit"]], "direct": (r[idx["Direct"]] or "").strip(),
        }
    return people

def branch_of(pid, people):
    seen = set(); cur = pid; chain = []
    while cur and cur in people and cur not in seen:
        seen.add(cur); chain.append(cur)
        cur = people[cur]["father"]
    if pid in ("E-johanngeorgwaldi","E-hansjacobwaldi","E-hansphilipp1693",
               "E-johannes1713","E-annabarbara1714"):
        return "trunk"
    if "E-johnjacob1710" in chain:
        return "jj"
    if "E-georgwolfgang-disp" in chain:
        return "gw"
    if "E-philipjacob" in chain:
        return "pj"
    return "trunk"

def build():
    people = load()
    for pid,p in people.items():
        if p["direct"] in ("yes","kit"):
            cur = pid
            while cur and cur in people:
                DIRECT_IDS.add(cur); cur = people[cur]["father"]

    node_of = {}
    for pid,p in people.items():
        if p["place"]:
            n,_ = node_for(pid, p["place"], None); node_of[pid] = n
    changed = True
    while changed:
        changed = False
        for pid,p in people.items():
            if node_of.get(pid) is None:
                fn = node_of.get(p["father"])
                if fn: node_of[pid] = fn; changed = True

    de_nodes, us_nodes = {}, {}
    placed, skipped = [], []
    for pid,p in people.items():
        n = node_of.get(pid)
        if not n or n not in NODES:
            skipped.append(pid); continue
        entry = {
            "id": pid, "name": p["name"], "gen": p["gen"], "branch": branch_of(pid, people),
            "proof": p["proof"], "birth": scrub_display(p["birth"]), "death": scrub_display(p["death"]),
            "spouse": scrub_display(p["spouse"]), "place": scrub_display(p["place"]), "dna": p["dna"],
            "inherited": (not p["place"]), "direct": pid in DIRECT_IDS,
        }
        cont = NODES[n][3]
        (de_nodes if cont == "DE" else us_nodes).setdefault(n, []).append(entry)
        placed.append(pid)
        if cont == "US" and p["place"] and "edenkoben" in p["place"].lower():
            echo = dict(entry); echo["emig"] = True
            de_nodes.setdefault("edenkoben", []).append(echo)

    def pack(nodedict):
        out = {}
        for n, ppl in nodedict.items():
            lat,lon,label,_ = NODES[n]
            ppl_sorted = sorted(ppl, key=lambda e:(e.get("emig",False), e["gen"], e["name"] or ""))
            out[n] = {"lat":lat,"lon":lon,"label":label,"chapter":CHAPTER.get(n),
                      "people":ppl_sorted}
        return out

    return {"de":pack(de_nodes),"us":pack(us_nodes)}, placed, skipped

def counts(data):
    from collections import Counter
    br = Counter(); pr = Counter(); tot = 0
    for cont in ("de","us"):
        for n,v in data[cont].items():
            for p in v["people"]:
                if p.get("emig"): continue
                br[p["branch"]] += 1; pr[p["proof"] or "?"] += 1; tot += 1
    return tot, br, pr

DATA, PLACED, SKIPPED = build()

# ---------------------------------------------------------------- fragment
FRAG_STYLE = """<style>
  /* guard Leaflet's vector overlay from any global `svg{width:100%}` on the host page */
  .leaflet-container svg, .leaflet-pane svg, .leaflet-overlay-pane svg{width:auto;height:auto;max-width:none;display:block;}
  .pm-wrap{--pm-trunk:#a07a35;--pm-pj:#8a5a2b;--pm-jj:#5b6e3a;--pm-gw:#46618a;--pm-star:#c69b3a;}
  .pm-intro{max-width:760px;margin:0 auto 12px;color:#6b6156;text-align:center;font-size:.82em;line-height:1.45;}
  .pm-intro b{color:var(--pm-pj);}
  .pm-chlink{display:block;margin-top:7px;font-size:.9em;}
  .pm-stats{display:flex;flex-wrap:wrap;justify-content:center;gap:7px 9px;margin:0 auto 14px;max-width:860px;}
  .pm-stat{background:#fffdf8;border:1px solid #e3d9c4;border-radius:20px;padding:4px 13px;font-size:.83em;}
  .pm-stat b{font-size:1.12em;}
  .pm-legend{display:flex;flex-wrap:wrap;justify-content:center;gap:8px 18px;font-size:.84em;margin:0 auto 6px;max-width:920px;}
  .pm-legend span{display:inline-flex;align-items:center;gap:6px;}
  .pm-dot{width:13px;height:13px;border-radius:50%;display:inline-block;border:2px solid #fff;box-shadow:0 0 0 1px #bbb;}
  .pm-legend2{display:flex;flex-wrap:wrap;justify-content:center;gap:6px 16px;font-size:.78em;color:#6b6156;margin:0 auto 16px;max-width:920px;}
  .pm-lmap{height:500px;border-radius:8px;}
  .pm-lmap.de{height:360px;}
  .pm-badge{font-size:.7em;padding:1px 6px;border-radius:9px;border:1px solid;margin-left:3px;white-space:nowrap;}
  .pm-proven{color:#3d6b3d;border-color:#3d6b3d;background:#eef5ee;}
  .pm-documented{color:#4a5c8a;border-color:#4a5c8a;background:#eef0f7;}
  .pm-hypo{color:#9a6a1e;border-color:#9a6a1e;background:#f7f0e2;}
  .pm-lore{color:#9a4a4a;border-color:#9a4a4a;background:#f7ecec;}
  .pm-star{color:var(--pm-star);}
  .pm-ph{font-weight:bold;color:#7a2e2e;font-size:1.05em;border-bottom:1px solid #c9b99a;padding-bottom:4px;margin-bottom:6px;}
  .pm-row{display:flex;align-items:baseline;gap:7px;margin:3px 0;font-size:.93em;line-height:1.35;}
  .pm-pdot{flex:0 0 auto;width:10px;height:10px;border-radius:50%;margin-top:4px;border:1.5px solid #fff;box-shadow:0 0 0 1px #ccc;}
  .pm-name{font-weight:600;}
  .pm-meta{color:#6b6156;font-size:.9em;}
  .pm-emig{color:#6b6156;font-style:italic;font-size:.86em;}
  .leaflet-popup-content .pm-ph, .leaflet-popup-content .pm-row{font-family:'EB Garamond',Georgia,serif;}
  .leaflet-popup-content{max-height:330px;overflow:auto;}
</style>"""

FRAG_MARKUP = """<div class="pm-wrap">
  <div class="mapbox" style="padding:16px 18px 6px;background:#fbf5e8;">
    <h2 style="margin-top:0;">Everyone we&rsquo;ve placed &mdash; the first six generations</h2>
    <div class="pm-intro">
      Every person <b>placed</b> in generations 1&ndash;6 of the Edenkoben line. Click a place for everyone who lived there. Marker <b>colour</b> = branch &middot; <b>style</b> = proof &middot; <span class="pm-star">&#9733;</span> = the direct line.
    </div>
    <div class="pm-stats">__STATS__</div>
    <div class="pm-legend">
      <span><i class="pm-dot" style="background:var(--pm-trunk)"></i> Old Country trunk</span>
      <span><i class="pm-dot" style="background:var(--pm-pj)"></i> Philip Jacob &rarr; Michael</span>
      <span><i class="pm-dot" style="background:var(--pm-jj)"></i> John Jacob &rarr; Weltytown</span>
      <span><i class="pm-dot" style="background:var(--pm-gw)"></i> Georg Wolfgang &rarr; Manchester</span>
      <span><span class="pm-star" style="font-size:1.2em">&#9733;</span> the direct line</span>
    </div>
    <div class="pm-legend2">
      <span>Proof: <span class="pm-badge pm-proven">proven</span> <span class="pm-badge pm-documented">documented</span> <span class="pm-badge pm-hypo">hypothesis</span> <span class="pm-badge pm-lore">family lore</span></span>
      <span>&mdash; larger circle = more people at that place</span>
    </div>
    <p class="hint" style="margin-left:2px;">The Old Country &mdash; Bischheim &amp; the Edenkoben household (names marked &ldquo;emigrated&rdquo; sailed and reappear on the American map):</p>
    <div id="pmMapDe" class="pm-lmap de"></div>
    <p class="hint" style="margin:8px 0 0 2px;">America &mdash; Philadelphia landings &rarr; York County &rarr; the westward push. Faint lines trace each brother&rsquo;s branch:</p>
    <div id="pmMapUs" class="pm-lmap"></div>
    <p class="tilenote" style="font-size:.72em;color:#6b6156;margin:6px 2px 2px;">Map data &copy; OpenStreetMap contributors.</p>
  </div>
</div>"""

FRAG_SCRIPT = """<script>
(function(){
  var DATA = __DATA__;
  var CHLINKS = __CHLINKS__;
  var COL = {trunk:'#a07a35', pj:'#8a5a2b', jj:'#5b6e3a', gw:'#46618a'};
  var OSM='https://tile.openstreetmap.org/{z}/{x}/{y}.png';
  var ATTR='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors';
  function esc(s){return (s==null?'':String(s)).replace(/[&<>]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];});}
  function mkMap(id,bounds){var m=L.map(id,{scrollWheelZoom:false});
    L.tileLayer(OSM,{maxZoom:17,attribution:ATTR}).addTo(m);m.fitBounds(bounds);return m;}
  function route(m,pts,color,dash,w){L.polyline(pts,{color:color,weight:w||2.5,opacity:.5,dashArray:dash||null}).addTo(m);}
  function popupHtml(node){
    var h='<div class="pm-ph">'+esc(node.label)+'</div>';
    node.people.forEach(function(p){
      var c=COL[p.branch]||'#888';
      var star=p.direct?'<span class="pm-star" title="the direct line">&#9733;</span> ':'';
      var dates=[p.birth,p.death].filter(Boolean).join(' &ndash; ');
      var badge='<span class="pm-badge pm-'+esc(p.proof||'living')+'">'+esc(p.proof||'?')+'</span>';
      var emig=p.emig?' <span class="pm-emig">&#8599; emigrated to America</span>':'';
      var withp=(p.inherited&&!p.emig)?' <span class="pm-emig">(with parent)</span>':'';
      h+='<div class="pm-row"><span class="pm-pdot" style="background:'+c+'"></span><span>'
        +star+'<span class="pm-name">'+esc(p.name)+'</span> '+badge
        +'<span class="pm-meta"> &middot; Gen '+p.gen+(dates?' &middot; '+dates:'')+'</span>'+emig+withp+'</span></div>';
    });
    if(CHLINKS && node.chapter){h+='<a class="pm-chlink" href="#'+node.chapter+'">Read the chapter &darr;</a>';}
    return h;
  }
  function dominant(node){var c={},best=null,bn=-1;
    node.people.forEach(function(p){if(!p.emig)c[p.branch]=(c[p.branch]||0)+1;});
    for(var k in c){if(c[k]>bn){bn=c[k];best=k;}} return best||node.people[0].branch;}
  function addNodes(m,nodes){
    for(var key in nodes){
      var nd=nodes[key];
      var real=nd.people.filter(function(p){return !p.emig;});
      if(nd.people.length===0) continue;
      var col=COL[dominant(nd)]||'#888';
      var r=Math.min(9+Math.sqrt(real.length||1)*3.4,26);
      if(real.some(function(p){return p.direct;})){
        L.circleMarker([nd.lat,nd.lon],{radius:r+4,color:'#c69b3a',weight:2.5,opacity:.9,fill:false}).addTo(m);
      }
      var mk=L.circleMarker([nd.lat,nd.lon],{radius:r,color:'#fff',weight:2.2,fillColor:col,fillOpacity:.92}).addTo(m);
      mk.bindPopup(popupHtml(nd),{maxWidth:360,minWidth:250});
      var short=nd.label.split(/[,\\u2014(]/)[0].trim();
      mk.bindTooltip(short+' &middot; '+(real.length||nd.people.length),
        {permanent:true,direction:'right',className:'wtip',offset:[r-2,0]});
    }
  }
  function boot(){
    if(typeof L==='undefined'){return setTimeout(boot,120);}
    if(document.getElementById('pmMapDe')._leaflet_id) return;
    var md=mkMap('pmMapDe',[[49.15,7.85],[49.80,8.30]]);
    route(md,[[49.688,8.046],[49.283,8.128]],COL.trunk,'6 6',3);
    addNodes(md,DATA.de);
    var mu=mkMap('pmMapUs',[[32.4,-118.5],[42.4,-79.0]]);
    route(mu,[[39.953,-75.165],[40.001,-76.85],[40.018,-78.503],[39.97,-79.10],[40.017,-79.59],[40.52,-81.474],[40.657,-81.634],[41.293,-84.762]],COL.pj,null,3);
    route(mu,[[40.52,-81.474],[38.831,-104.821]],COL.pj,'2 8',1.8);
    route(mu,[[40.657,-81.634],[33.119,-117.086]],COL.pj,'2 8',1.8);
    route(mu,[[40.325,-80.655],[40.041,-86.874]],COL.pj,'2 8',1.6);
    route(mu,[[40.657,-81.634],[41.007,-91.962]],COL.pj,'2 8',1.6);
    route(mu,[[40.001,-76.85],[40.325,-80.655]],COL.pj,'5 6',2);
    route(mu,[[40.038,-76.305],[40.175,-79.545],[40.301,-79.539]],COL.jj,null,2.6);
    route(mu,[[39.953,-75.165],[40.030,-76.745]],COL.gw,null,2.6);
    route(mu,[[39.953,-75.165],[40.150,-75.650],[40.001,-76.85]],COL.trunk,'6 6',2.4);
    addNodes(mu,DATA.us);
  }
  if(document.readyState==='complete'){boot();}
  else{window.addEventListener('load',boot);}
})();
</script>"""

def stats_html():
    total, br, pr = counts(DATA)
    s = [f'<span class="pm-stat"><b>{total}</b> people</span>']
    for k,lbl in [("trunk","trunk"),("pj","Philip Jacob"),("jj","John Jacob"),("gw","Georg Wolfgang")]:
        if br.get(k): s.append(f'<span class="pm-stat"><b>{br[k]}</b> {lbl}</span>')
    for k in ("proven","documented","hypo","lore"):
        if pr.get(k): s.append(f'<span class="pm-stat">{pr[k]} {k}</span>')
    return "\n      ".join(s), total

def fragment(chlinks):
    stats, total = stats_html()
    markup = FRAG_MARKUP.replace("__STATS__", stats)
    script = (FRAG_SCRIPT
        .replace("__DATA__", json.dumps(DATA, ensure_ascii=False))
        .replace("__CHLINKS__", "true" if chlinks else "false"))
    return FRAG_STYLE + "\n" + markup + "\n" + script, total

# ---------------------------------------------------------------- outputs
STANDALONE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>The Welty Family, Everyone in Gen 1&ndash;6 &mdash; a Map</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css">
<link href="https://fonts.googleapis.com/css2?family=EB+Garamond:ital@0;1&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"></script>
<style>
  body{margin:0;background:#f5efe3;color:#2b2620;font-family:'EB Garamond',Georgia,serif;line-height:1.55;}
  .wrap{max-width:1000px;margin:0 auto;padding:34px 16px 80px;}
  h1{text-align:center;font-size:1.95em;margin:0 0 4px;letter-spacing:.5px;}
  h2{color:#7a2e2e;font-size:1.1em;}
  .subtitle{text-align:center;color:#6b6156;font-style:italic;margin-bottom:6px;}
  .updated{text-align:center;font-size:.82em;color:#7a2e2e;font-weight:bold;margin-bottom:18px;}
  .mapbox{background:#fffdf8;border:1px solid #e3d9c4;border-radius:10px;padding:14px;box-shadow:0 1px 3px rgba(60,45,20,.08);}
  .hint{font-size:.8em;color:#6b6156;font-style:italic;}
  .wtip{background:rgba(255,253,248,.92);border:1px solid #d8c9a8;border-radius:5px;font-family:'EB Garamond',Georgia,serif;font-size:11.5px;color:#2b2620;padding:1px 5px;box-shadow:none;}
  .wtip:before{display:none;}
</style>
__FRAGSTYLE__
</head>
<body>
<div class="wrap">
  <h1>Everyone in the First Six Generations</h1>
  <p class="subtitle">The proven Edenkoben Welty line, person by person, placed on the ground.</p>
  <p class="updated">__TOTAL__ people mapped &middot; updated 4 Jul 2026</p>
__FRAGBODY__
</div>
</body>
</html>"""

def write_standalone(fragstyle_and_body, total):
    # split fragment back into style vs body+script for the shell
    style = FRAG_STYLE
    body = fragstyle_and_body[len(FRAG_STYLE)+1:]
    html_out = (STANDALONE
        .replace("__FRAGSTYLE__", style)
        .replace("__FRAGBODY__", body)
        .replace("__TOTAL__", str(total)))
    with open(OUT, "w") as f:
        f.write(html_out)

MARK_A = "<!-- PEOPLEMAP:START -->"
MARK_B = "<!-- PEOPLEMAP:END -->"
ANCHOR = "<!-- ================= WHAT IS THE PFALZ ================= -->"

def inject(path, frag):
    if not os.path.exists(path):
        print("  (skip, missing:", os.path.basename(path), ")"); return
    with open(path) as f: html = f.read()
    block = MARK_A + "\n" + frag + "\n" + MARK_B + "\n\n"
    if MARK_A in html and MARK_B in html:
        pre = html[:html.index(MARK_A)]
        post = html[html.index(MARK_B)+len(MARK_B):].lstrip("\n")
        html = pre + block + post
    elif ANCHOR in html:
        html = html.replace(ANCHOR, block + "  " + ANCHOR, 1)
    else:
        print("  (no anchor in", os.path.basename(path), "- not injected)"); return
    with open(path, "w") as f: f.write(html)
    print("  injected ->", os.path.basename(path))

frag_solo, total = fragment(chlinks=False)   # standalone: no chapters to link
frag_embed, _   = fragment(chlinks=True)     # timeline: link into chapters
write_standalone(frag_solo, total)
inject(TIMELINE, frag_embed)
inject(TIMELINE_NETLIFY, frag_embed)
print(f"placed={len(PLACED)}  skipped={len(SKIPPED)} -> {SKIPPED}")
print(f"standalone -> {OUT}  ({total} people)")
