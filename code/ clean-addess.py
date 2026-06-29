"""
TH44 Address Cleaning — Offline Postal-Anchored Consensus Engine
================================================================
No external gazetteer (network blocked) -> gazetteer is built from the
dataset's own consensus + an encoded 77-province table. Fuzzy = difflib.

Strategy: don't "fix" messy text — REGENERATE geography from the clean
POSTAL CODE anchor, use messy text only to pick subdistrict + street detail.

Output: colour-coded status per row, 2-language assembled address,
per-field 'Fixed' provenance + confidence. Non-destructive (raw kept).
LLM residual = wired but gated on ANTHROPIC_API_KEY (inactive here).
"""

import os, re, difflib
from pathlib import Path
from collections import Counter, defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

INPUT_PATH  = Path("/mnt/user-data/uploads/24062026_TH44_all_CUSTOMERS_excl_flag_del.xlsx")
OUTPUT_PATH = Path("/mnt/user-data/outputs/TH44_Addresses_Cleaned.xlsx")

# ── 77-province table (Thai -> RTGS English) for the bilingual bridge ──────────
PROV_TH2EN = {
 "กรุงเทพมหานคร":"Bangkok","สมุทรปราการ":"Samut Prakan","นนทบุรี":"Nonthaburi",
 "ปทุมธานี":"Pathum Thani","พระนครศรีอยุธยา":"Phra Nakhon Si Ayutthaya","อ่างทอง":"Ang Thong",
 "ลพบุรี":"Lopburi","สิงห์บุรี":"Sing Buri","ชัยนาท":"Chai Nat","สระบุรี":"Saraburi",
 "ชลบุรี":"Chonburi","ระยอง":"Rayong","จันทบุรี":"Chanthaburi","ตราด":"Trat",
 "ฉะเชิงเทรา":"Chachoengsao","ปราจีนบุรี":"Prachinburi","นครนายก":"Nakhon Nayok","สระแก้ว":"Sa Kaeo",
 "นครราชสีมา":"Nakhon Ratchasima","บุรีรัมย์":"Buriram","สุรินทร์":"Surin","ศรีสะเกษ":"Sisaket",
 "อุบลราชธานี":"Ubon Ratchathani","ยโสธร":"Yasothon","ชัยภูมิ":"Chaiyaphum","อำนาจเจริญ":"Amnat Charoen",
 "บึงกาฬ":"Bueng Kan","หนองบัวลำภู":"Nong Bua Lamphu","ขอนแก่น":"Khon Kaen","อุดรธานี":"Udon Thani",
 "เลย":"Loei","หนองคาย":"Nong Khai","มหาสารคาม":"Maha Sarakham","ร้อยเอ็ด":"Roi Et",
 "กาฬสินธุ์":"Kalasin","สกลนคร":"Sakon Nakhon","นครพนม":"Nakhon Phanom","มุกดาหาร":"Mukdahan",
 "เชียงใหม่":"Chiang Mai","ลำพูน":"Lamphun","ลำปาง":"Lampang","อุตรดิตถ์":"Uttaradit",
 "แพร่":"Phrae","น่าน":"Nan","พะเยา":"Phayao","เชียงราย":"Chiang Rai","แม่ฮ่องสอน":"Mae Hong Son",
 "นครสวรรค์":"Nakhon Sawan","อุทัยธานี":"Uthai Thani","กำแพงเพชร":"Kamphaeng Phet","ตาก":"Tak",
 "สุโขทัย":"Sukhothai","พิษณุโลก":"Phitsanulok","พิจิตร":"Phichit","เพชรบูรณ์":"Phetchabun",
 "ราชบุรี":"Ratchaburi","กาญจนบุรี":"Kanchanaburi","สุพรรณบุรี":"Suphan Buri","นครปฐม":"Nakhon Pathom",
 "สมุทรสาคร":"Samut Sakhon","สมุทรสงคราม":"Samut Songkhram","เพชรบุรี":"Phetchaburi",
 "ประจวบคีรีขันธ์":"Prachuap Khiri Khan","นครศรีธรรมราช":"Nakhon Si Thammarat","กระบี่":"Krabi",
 "พังงา":"Phang Nga","ภูเก็ต":"Phuket","สุราษฎร์ธานี":"Surat Thani","ระนอง":"Ranong","ชุมพร":"Chumphon",
 "สงขลา":"Songkhla","สตูล":"Satun","ตรัง":"Trang","พัทลุง":"Phatthalung","ปัตตานี":"Pattani",
 "ยะลา":"Yala","นราธิวาส":"Narathiwat",
}
PROV_EN2TH = {v.lower(): k for k, v in PROV_TH2EN.items()}
PROV_TH_SET = set(PROV_TH2EN)
BKK_ALIASES = {"กรุงเทพ","กรุงเทพฯ","กทม","กทม.","กรุงเทพมหานคร","bangkok","krung thep","bkk","กรุงเทพมหา นคร"}

# ── status colour scheme (designed per user request) ───────────────────────────
STATUS_FILL = {
 "VERIFIED":     PatternFill("solid", start_color="C8E6C9"),  # green
 "AUTO_FIXED":   PatternFill("solid", start_color="BBDEFB"),  # blue
 "FUZZY_FIXED":  PatternFill("solid", start_color="FFF9C4"),  # yellow
 "LLM_RESOLVED": PatternFill("solid", start_color="E1BEE7"),  # purple
 "NEEDS_REVIEW": PatternFill("solid", start_color="FFCDD2"),  # red
 "FOREIGN":      PatternFill("solid", start_color="E0E0E0"),  # gray
}
STATUS_RANK = {"VERIFIED":0,"AUTO_FIXED":1,"FUZZY_FIXED":2,"LLM_RESOLVED":3,"NEEDS_REVIEW":4,"FOREIGN":5}

# ── script + text helpers ──────────────────────────────────────────────────────
TH_RE = re.compile(r"[\u0E00-\u0E7F]")
def script_of(s):
    if not s: return None
    t = TH_RE.search(s); l = re.search(r"[A-Za-z]", s)
    if t and not l: return "TH"
    if l and not t: return "EN"
    if t and l:     return "TH" if len(TH_RE.findall(s)) >= len(re.findall(r"[A-Za-z]", s)) else "EN"
    return None

def clean_ws(s):
    if pd.isna(s): return None
    s = re.sub(r"\s{2,}", " ", str(s).strip()).strip(" ,")
    return s or None

PROV_EN_NORM = {re.sub(r"\s+", "", v.lower()): k for k, v in PROV_TH2EN.items()}
PROV_ALIAS_EN = {"ayutthaya":"พระนครศรีอยุธยา","ayudhya":"พระนครศรีอยุธยา","korat":"นครราชสีมา",
                 "khorat":"นครราชสีมา","bangkok":"กรุงเทพมหานคร","krungthep":"กรุงเทพมหานคร",
                 "krungthepmahanakhon":"กรุงเทพมหานคร","bkk":"กรุงเทพมหานคร","sriracha":"ชลบุรี"}
TH_PROV_LIST = list(PROV_TH2EN.keys())

def norm_prov(s):
    """Canonicalise a province surface form (any script, abbrev, typo) -> Thai key."""
    s = clean_ws(s)
    if not s: return None
    s = re.sub(r"\s*\bprovince\b\s*", "", s, flags=re.I).strip()
    s = re.sub(r"^(จังหวัด|จ\.|จ\s)\s*", "", s).strip()
    low = re.sub(r"\s+", "", s.lower()).replace("ฯ", "")
    if low in {"กรุงเทพ","กรุงเทพมหานคร","กทม","กทม."} or low in PROV_ALIAS_EN:
        return PROV_ALIAS_EN.get(low, "กรุงเทพมหานคร")
    if s in PROV_TH2EN: return s                       # Thai exact
    if low in PROV_EN_NORM: return PROV_EN_NORM[low]    # English exact (space-insensitive)
    if TH_RE.search(s):                                # Thai typo -> fuzzy
        m = difflib.get_close_matches(s, TH_PROV_LIST, n=1, cutoff=0.82)
        if m: return m[0]
    else:                                              # English typo -> fuzzy
        m = difflib.get_close_matches(low, list(PROV_EN_NORM), n=1, cutoff=0.82)
        if m: return PROV_EN_NORM[m[0]]
    return s                                           # unresolved (kept as-is)

def strip_dist(s):
    s = clean_ws(s)
    if not s: return None
    return re.sub(r"^(อำเภอ|อ\.|เขต|อ\s)\s*", "", s).strip() or None

def strip_sub(s):
    s = clean_ws(s)
    if not s: return None
    return re.sub(r"^(ตำบล|ต\.|แขวง|ต\s)\s*", "", s).strip() or None

def norm_road(s):
    s = clean_ws(s)
    if not s: return None
    s = re.sub(r"^ถ\.\s*", "ถนน", s)
    s = re.sub(r"^ซ\.\s*", "ซอย", s)
    return s

def parse_moo(s):
    s = clean_ws(s)
    if not s: return None
    m = re.search(r"(?:หมู่ที่|หมู่|ม\.)\s*(\d+)", s)
    return m.group(1) if m else (s if re.fullmatch(r"\d+", s) else None)

def parse_house(s):
    """Return (house_no, building_name_if_wrong_column)."""
    s = clean_ws(s)
    if not s: return None, None
    s2 = re.sub(r"^เลขที่\s*", "", s).strip()
    if re.search(r"\d", s2):
        m = re.match(r"^([\d/\-\s]+)", s2)
        return (m.group(1).strip() if m else s2), None
    return None, s2   # no digit -> it's a building/place name (wrong column)

# ── fuzzy canonicaliser ────────────────────────────────────────────────────────
def build_canon(values, cutoff=0.86):
    """Most-frequent spelling wins; rare near-duplicates map to it."""
    cnt = Counter(v for v in values if isinstance(v, str) and v.strip())
    mapping, accepted = {}, []
    for w, _ in cnt.most_common():
        m = difflib.get_close_matches(w, accepted, n=1, cutoff=cutoff)
        if m: mapping[w] = m[0]
        else: mapping[w] = w; accepted.append(w)
    return mapping, set(accepted)

# ══════════════════════════════════════════════════════════════════════════════
print("Loading …")
df = pd.read_excel(INPUT_PATH, dtype=str)
N = len(df)
is_th = df["COUNTRY"] == "Thailand"

# Pre-normalise the three admin columns + split by script
work = pd.DataFrame(index=df.index)
work["prov_raw"] = df["City"].apply(norm_prov)
work["dist_raw"] = df["Other city"].apply(strip_dist)
work["sub_raw"]  = df["Dist"].apply(strip_sub)
work["postal"]   = df["POSTAL CODE"].apply(lambda s: clean_ws(s))
work["script"]   = df.apply(lambda r: script_of(" ".join(
                      [str(r["Other city"] or ""), str(r["City"] or ""), str(r["Dist"] or "")])) or "TH", axis=1)

# ── Build consensus maps from Thailand rows ────────────────────────────────────
print("Building consensus gazetteer from data …")
thmask = is_th & work["postal"].notna()

# postal -> province consensus (script-agnostic Thai key; only count recognised provinces)
postal_prov = defaultdict(Counter)
for i in df.index[thmask]:
    p = work.at[i, "postal"]; pr = work.at[i, "prov_raw"]
    if isinstance(pr, str) and pr in PROV_TH2EN:
        postal_prov[p][pr] += 1

def prov_consensus(postal):
    c = postal_prov.get(postal)
    return c.most_common(1)[0][0] if c else None

# province (Thai key) -> district canon  /  province -> subdistrict canon (per script)
dist_by_prov = defaultdict(lambda: defaultdict(list))
sub_by_prov  = defaultdict(lambda: defaultdict(list))
postal_dist  = defaultdict(lambda: defaultdict(Counter))
for i in df.index[thmask]:
    pr = prov_consensus(work.at[i, "postal"]) or work.at[i, "prov_raw"]
    sc = work.at[i, "script"]
    if work.at[i, "dist_raw"]:
        dist_by_prov[pr][sc].append(work.at[i, "dist_raw"])
        postal_dist[work.at[i, "postal"]][sc][work.at[i, "dist_raw"]] += 1
    if work.at[i, "sub_raw"]:
        sub_by_prov[pr][sc].append(work.at[i, "sub_raw"])

dist_canon = {pr: {sc: build_canon(v) for sc, v in d.items()} for pr, d in dist_by_prov.items()}
sub_canon  = {pr: {sc: build_canon(v) for sc, v in d.items()} for pr, d in sub_by_prov.items()}

def resolve(raw, canon_for_prov, script):
    """Return (value, how) where how in exact/fuzzy/unmatched/empty."""
    if not isinstance(raw, str) or not raw.strip(): return None, "empty"
    pack = canon_for_prov.get(script)
    if not pack: return raw, "unmatched"
    mapping, canon_set = pack
    if raw in canon_set: return raw, "exact"
    if raw in mapping and mapping[raw] != raw: return mapping[raw], "fuzzy"
    m = difflib.get_close_matches(raw, list(canon_set), n=1, cutoff=0.86)
    if m: return m[0], "fuzzy"
    return raw, "unmatched"

# ── LLM residual module (wired; activates only with ANTHROPIC_API_KEY) ──────────
def llm_resolve_residual(res_df, src_df):
    """Resolve NEEDS_REVIEW rows with an LLM, with province/postal LOCKED from the
    postal anchor so the model cannot hallucinate the high-level geography.
    Inactive unless ANTHROPIC_API_KEY is set (and outbound network is allowed).
    Returns count resolved. Safe to call — no-ops cleanly when key is absent."""
    key = os.getenv("ANTHROPIC_API_KEY")
    targets = res_df.index[res_df["status"] == "NEEDS_REVIEW"].tolist()
    if not key:
        print(f"LLM residual: SKIPPED (no ANTHROPIC_API_KEY) — "
              f"{len(targets)} rows kept as NEEDS_REVIEW for manual/LLM review.")
        return 0
    import json, urllib.request
    resolved = 0
    for i in targets:
        locked_prov = res_df.at[i, "prov"]; postal = res_df.at[i, "postal"]
        raw = {c: src_df.at[i, c] for c in
               ["STREET","STREET2","STREET3","STREET4","STREET5","Dist","Other city","City"]}
        prompt = ("You are normalising a Thai address. The PROVINCE and POSTAL CODE below are "
                  "verified — do NOT change them. Using the raw fragments, return ONLY a JSON object "
                  '{"house_no","moo","soi","road","subdistrict","district"} in the SAME language as '
                  f"the source.\nVERIFIED province: {locked_prov}\nVERIFIED postal: {postal}\n"
                  f"Raw fragments: {json.dumps(raw, ensure_ascii=False)}")
        body = json.dumps({"model":"claude-sonnet-4-6","max_tokens":400,
                           "messages":[{"role":"user","content":prompt}]}).encode()
        req = urllib.request.Request("https://api.anthropic.com/v1/messages", data=body,
              headers={"x-api-key":key,"anthropic-version":"2023-06-01","content-type":"application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                txt = json.loads(r.read())["content"][0]["text"]
            obj = json.loads(txt[txt.find("{"):txt.rfind("}")+1])
            for k_, col in [("subdistrict","sub"),("district","dist"),("road","road"),
                            ("moo","moo"),("soi","soi"),("house_no","house")]:
                if obj.get(k_): res_df.at[i, col] = obj[k_]
            res_df.at[i, "status"] = "LLM_RESOLVED"; res_df.at[i, "conf"] = 70
            res_df.at[i, "flags"] = (str(res_df.at[i,"flags"]) + "; llm_resolved").strip("; ")
            resolved += 1
        except Exception as e:
            print(f"  LLM row {i} failed: {repr(e)[:80]}")
    print(f"LLM residual: resolved {resolved}/{len(targets)} rows.")
    return resolved

# ── Per-row resolution ─────────────────────────────────────────────────────────
print("Resolving addresses …")
rows = []
for i in df.index:
    sc = work.at[i, "script"]
    postal = work.at[i, "postal"]
    postal = postal if isinstance(postal, str) and postal.strip() else None
    flags, statuses = [], []

    if not is_th[i]:
        # FOREIGN — assemble from cleaned raw, no Thai logic
        parts = [clean_ws(df.at[i, c]) for c in
                 ["STREET","STREET2","STREET3","STREET4","STREET5","Dist","Other city","City"]]
        native = ", ".join([p for p in parts if p])
        rows.append(dict(house="",building="",moo="",soi="",road="",sub="",dist="",
                         prov=clean_ws(df.at[i,"City"]) or "", postal=postal or "",
                         country=df.at[i,"COUNTRY"], native=native, en=native,
                         status="FOREIGN", conf=60, lang=sc or "EN",
                         flags="foreign address — assembled as-is"))
        continue

    # ---------- PROVINCE: postal anchor vs City ----------
    p_postal = prov_consensus(postal) if postal else None
    city_key = norm_prov(df.at[i, "City"])
    city_is_prov = isinstance(city_key, str) and city_key in PROV_TH2EN
    prov_conflict = False
    if p_postal and city_is_prov:
        prov_th = p_postal
        if city_key != p_postal:
            prov_conflict = True
            flags.append(f"province:conflict(city={city_key}->postal={p_postal})")
    elif p_postal:
        prov_th = p_postal
        if city_key is None: flags.append("province:from_postal(city_blank)")
        else: flags.append(f"province:from_postal(city_unrecognized={city_key})")
    elif city_is_prov:
        prov_th = city_key
    else:
        prov_th = None

    is_bkk = prov_th == "กรุงเทพมหานคร"
    sub_pfx, dist_pfx = ("แขวง", "เขต") if is_bkk else ("ตำบล", "อำเภอ")

    if prov_th is None:
        prov_native = clean_ws(df.at[i, "City"]) or ""; prov_en = prov_native
    elif sc == "TH":
        prov_native = "กรุงเทพมหานคร" if is_bkk else f"จังหวัด{prov_th}"
        prov_en = PROV_TH2EN.get(prov_th, prov_th)
    else:
        prov_en = PROV_TH2EN.get(prov_th, prov_th); prov_native = prov_en

    if prov_conflict or prov_th is None:
        prov_ct = "REVIEW"
    elif clean_ws(df.at[i, "City"]) == prov_native:
        prov_ct = "UNCHANGED"
    else:
        prov_ct = "AUTO"
        if not any(f.startswith("province:") for f in flags):
            flags.append(f"province:standardized({clean_ws(df.at[i,'City'])}->{prov_native})")

    # ---------- DISTRICT ----------
    dist_val, how = resolve(work.at[i, "dist_raw"], dist_canon.get(prov_th, {}), sc)
    dist_disp, dist_ct = None, None
    if how != "empty":
        dist_disp = (f"{dist_pfx}{dist_val}" if sc == "TH" else dist_val)
        if how == "fuzzy":
            dist_ct = "FUZZY"; flags.append(f"district:typo_fixed({work.at[i,'dist_raw']}->{dist_val})")
        elif how == "unmatched":
            dist_ct = "AUTO"; flags.append(f"district:kept_unverified({dist_val})")
        else:
            if clean_ws(df.at[i, "Other city"]) == dist_disp:
                dist_ct = "UNCHANGED"
            else:
                dist_ct = "AUTO"; flags.append(f"district:standardized({clean_ws(df.at[i,'Other city'])}->{dist_disp})")

    # ---------- SUBDISTRICT ----------
    sub_val, how = resolve(work.at[i, "sub_raw"], sub_canon.get(prov_th, {}), sc)
    sub_disp, sub_ct = None, None
    if how != "empty":
        sub_disp = (f"{sub_pfx}{sub_val}" if sc == "TH" else sub_val)
        if how == "fuzzy":
            sub_ct = "FUZZY"; flags.append(f"subdistrict:typo_fixed({work.at[i,'sub_raw']}->{sub_val})")
        elif how == "unmatched":
            sub_ct = "AUTO"; flags.append(f"subdistrict:kept_unverified({sub_val})")
        else:
            if clean_ws(df.at[i, "Dist"]) == sub_disp:
                sub_ct = "UNCHANGED"
            else:
                sub_ct = "AUTO"; flags.append(f"subdistrict:standardized({clean_ws(df.at[i,'Dist'])}->{sub_disp})")

    # ---------- STREET DETAIL ----------
    house, bld = parse_house(df.at[i, "STREET2"])
    if bld: flags.append(f"wrong_column:STREET2_name({bld})")
    moo = parse_moo(df.at[i, "STREET3"])
    road = norm_road(df.at[i, "STREET5"])
    if road and road != clean_ws(df.at[i, "STREET5"]): flags.append("road:abbrev_expanded")
    soi_raw = clean_ws(df.at[i, "STREET4"])
    soi = soi_raw if (soi_raw and re.search(r"ซอย|ซ\.|soi", soi_raw, re.I)) else None

    # ---------- ASSEMBLE NATIVE ----------
    if sc == "TH":
        seg = []
        if house: seg.append(f"เลขที่ {house}")
        if bld: seg.append(bld)
        if moo: seg.append(f"หมู่ที่ {moo}")
        if soi: seg.append(soi)
        if road: seg.append(road)
        if sub_disp: seg.append(sub_disp)
        if dist_disp: seg.append(dist_disp)
        if prov_native: seg.append(prov_native)
        if postal: seg.append(postal)
        native = " ".join(seg)
    else:
        head = [x for x in [house or bld, soi, road] if x]
        tail = [x for x in [sub_val, dist_val, prov_native, postal] if x]
        native = ", ".join(([" ".join(head)] if head else []) + tail)

    # ---------- ASSEMBLE EN ----------
    head_en = [x for x in [house or bld, soi, (road if sc == "EN" else None)] if x]
    tail_en = [x for x in [(sub_val if sc == "EN" else None),
                           (dist_val if sc == "EN" else None), prov_en, postal] if x]
    addr_en = ", ".join(([" ".join(head_en)] if head_en else []) + tail_en + ["Thailand"])
    if sc == "TH" and (sub_val or dist_val):
        flags.append("en:partial(thai_names_not_romanized)")

    # ---------- STATUS + CONFIDENCE ----------
    cts = [c for c in (prov_ct, dist_ct, sub_ct) if c]
    if "REVIEW" in cts:  status = "NEEDS_REVIEW"
    elif "FUZZY" in cts: status = "FUZZY_FIXED"
    elif "AUTO" in cts:  status = "AUTO_FIXED"
    else:                status = "VERIFIED"
    # completeness: a Thai address with neither district nor subdistrict is not "clean"
    if not dist_val and not sub_val and status in ("VERIFIED", "AUTO_FIXED"):
        status = "NEEDS_REVIEW"; flags.append("incomplete:no_district_or_subdistrict")
    change_flags = [f for f in flags if not f.startswith("en:")]
    conf = {"VERIFIED":97,"AUTO_FIXED":88,"FUZZY_FIXED":78,"NEEDS_REVIEW":45}[status] \
           - min(max(len(change_flags) - 1, 0) * 3, 12)

    rows.append(dict(house=house or "", building=bld or "", moo=moo or "", soi=soi or "",
                     road=road or "", sub=sub_val or "", dist=dist_val or "",
                     prov=prov_native or "", postal=postal or "", country="Thailand",
                     native=native, en=addr_en, status=status, conf=conf, lang=sc,
                     flags="; ".join(flags) if flags else "no_changes_needed"))

res = pd.DataFrame(rows, index=df.index)

# ── Build output frame: context + raw + cleaned + provenance ───────────────────
out = pd.DataFrame({
    "BP Number": df["BP Number"], "Cuscode": df["Cuscode"], "NAME1": df["NAME1"],
    "COUNTRY": df["COUNTRY"],
    "raw_STREET2": df["STREET2"], "raw_STREET3": df["STREET3"], "raw_STREET4": df["STREET4"],
    "raw_STREET5": df["STREET5"], "raw_Dist(subdist)": df["Dist"],
    "raw_Othercity(dist)": df["Other city"], "raw_City(prov)": df["City"], "raw_POSTAL": df["POSTAL CODE"],
    "addr_house_no": res["house"], "addr_building": res["building"], "addr_moo": res["moo"],
    "addr_soi": res["soi"], "addr_road": res["road"], "addr_subdistrict": res["sub"],
    "addr_district": res["dist"], "addr_province": res["prov"], "addr_postal": res["postal"],
    "addr_country": res["country"],
    "addr_full_native": res["native"], "addr_full_en": res["en"],
    "addr_status": res["status"], "addr_confidence": res["conf"], "addr_lang": res["lang"],
    "addr_flags": res["flags"],
})

ADDED = [c for c in out.columns if c.startswith("addr_")]

# ── Write workbook with colour-coded status ────────────────────────────────────
print("Writing workbook …")
def prep(frame):
    o = frame.copy()
    for c in o.columns: o[c] = [v if pd.notna(v) else None for v in o[c]]
    return o

wb = Workbook(); del wb["Sheet"]
ws = wb.create_sheet("Addresses")
hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
hdr_fill = PatternFill("solid", start_color="37474F")
add_fill = PatternFill("solid", start_color="F57C00")
rws = list(dataframe_to_rows(prep(out), index=False, header=True))
status_col = list(out.columns).index("addr_status") + 1
native_col = list(out.columns).index("addr_full_native") + 1
added_idx = {list(out.columns).index(c)+1 for c in ADDED}
for r_i, row in enumerate(rws, 1):
    for c_i, val in enumerate(row, 1):
        cell = ws.cell(r_i, c_i, val)
        if r_i == 1:
            cell.font = hdr_font; cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.fill = add_fill if c_i in added_idx else hdr_fill
        else:
            cell.font = Font(name="Arial", size=10)
            if c_i == status_col and val in STATUS_FILL: cell.fill = STATUS_FILL[val]
            elif c_i == native_col and row[status_col-1] in STATUS_FILL:
                cell.fill = STATUS_FILL[row[status_col-1]]
ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
for c_i in range(1, len(rws[0])+1):
    w = max(len(str(r[c_i-1] or "")) for r in rws[:60])
    ws.column_dimensions[ws.cell(1, c_i).column_letter].width = min(w+2, 50)

# Status summary sheet
summ = Counter(res["status"])
ssum = pd.DataFrame([(k, summ.get(k,0), f"{100*summ.get(k,0)/N:.1f}%") for k in STATUS_RANK],
                    columns=["addr_status","count","pct"])
ws2 = wb.create_sheet("Status_Summary")
for r_i, row in enumerate(list(dataframe_to_rows(ssum, index=False, header=True)), 1):
    for c_i, val in enumerate(row, 1):
        cell = ws2.cell(r_i, c_i, val)
        if r_i == 1: cell.font = Font(bold=True)
        elif c_i == 1 and val in STATUS_FILL: cell.fill = STATUS_FILL[val]
for c_i in (1,2,3): ws2.column_dimensions[ws2.cell(1,c_i).column_letter].width = 18

wb.save(OUTPUT_PATH)

# ── Report ─────────────────────────────────────────────────────────────────────
print("\n" + "="*60)
print("  ADDRESS CLEANING COMPLETE")
print(f"  Rows processed:   {N:,}")
for k in STATUS_RANK:
    if summ.get(k): print(f"    {k:13} {summ[k]:>6,}  ({100*summ[k]/N:4.1f}%)")
print(f"\n  Output: {OUTPUT_PATH.name}")
print("="*60)
print("\nSample cleaned addresses:")
for st in ["VERIFIED","FUZZY_FIXED","NEEDS_REVIEW"]:
    ex = res[res["status"]==st].head(1)
    if len(ex): print(f"  [{st}] {ex.iloc[0]['native'][:90]}")