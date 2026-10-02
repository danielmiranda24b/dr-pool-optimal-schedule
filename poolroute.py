#!/usr/bin/env python3
"""
poolroute - weekly pool-route optimizer. Zero AI tokens: just run it.

  python poolroute.py SCHEDULE.xlsx CONTACTS.xls [--depot "address"] [--service-min 25] [--only MIKA]

1. Reads each tech's stops per day from the schedule workbook.
2. Matches every stop to the contacts database (fuzzy name match) to get its address.
3. Geocodes addresses (US Census, then OpenStreetMap fallback). Cached forever in cache/geocode.json,
   so only NEW addresses are ever looked up.
4. Gets real road distance + drive time between stops (OSRM). Falls back to straight-line x1.35 if offline.
5. Re-orders each tech's day twice: SHORTEST MILES and FASTEST TIME. Writes one Excel file.
"""
import argparse, glob, io, itertools, json, math, os, random, re, sys, time, urllib.parse, urllib.request
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "cache"); os.makedirs(CACHE, exist_ok=True)
UA = {"User-Agent": "poolroute/1.0 (route planning script)"}
BBOX = (25.3, 26.9, -80.9, -79.9)   # South Florida sanity box (lat_min, lat_max, lon_min, lon_max)
DEFAULT_DEPOT = "2850 Glades Circle, Bay #4, Weston, FL 33327"
DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT"]

# ---------------------------------------------------------------- inputs
def find_latest(prefix, exts):
    """Newest file (by modified time) named PREFIX*.ext in the current folder or next to this script.
    Lets the weekly file keep changing its name, e.g. SCHEDULE_10.11.2026_-10.17.2026.xlsx."""
    found = {}
    for folder in (os.getcwd(), HERE):
        for f in os.listdir(folder):
            if f.lower().startswith(prefix.lower()) and f.lower().endswith(exts) and not f.startswith(("~$", "optimized_")):
                p = os.path.join(folder, f); found[os.path.abspath(p)] = os.path.getmtime(p)
    if not found: sys.exit(f"No {prefix}*{exts[0]} file found in {os.getcwd()} - put it there or pass its path.")
    return max(found, key=found.get)

def load_schedule(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    starts = [r for r in range(1, ws.max_row + 1)
              if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().upper().startswith("ROUTE")]
    rows = []
    for i, s in enumerate(starts):
        end = starts[i + 1] - 1 if i + 1 < len(starts) else ws.max_row
        title = ws.cell(s, 1).value
        tech = re.sub(r"ROUTE\s*#\s*\d+\s*", "", title.strip(), flags=re.I).strip()
        rnum = re.search(r"#\s*(\d+)", title).group(1)
        for r in range(s + 3, end + 1):
            seq = ws.cell(r, 1).value
            if seq is None or not str(seq).strip().isdigit():
                continue
            for d in range(6):
                nm = ws.cell(r, 2 + 2 * d).value
                if nm is None or not str(nm).strip():
                    continue
                rows.append(dict(route=rnum, tech=tech, day=DAYS[d], seq=int(str(seq).strip()),
                                 name=str(nm).strip(), price=ws.cell(r, 3 + 2 * d).value,
                                 sub=str(ws.cell(r + 1, 2 + 2 * d).value or "").strip(),
                                 svc=str(ws.cell(r + 1, 3 + 2 * d).value or "").strip()))
    return pd.DataFrame(rows)

def load_contacts(path):
    import xlrd
    if path.lower().endswith(".xls"):
        try:
            book = xlrd.open_workbook(path)
        except xlrd.biffh.XLRDError as e:              # file is "encrypted" with Excel's default read-only key
            if "encrypted" not in str(e).lower(): raise
            import msoffcrypto
            buf = io.BytesIO(); f = msoffcrypto.OfficeFile(open(path, "rb"))
            f.load_key(password="VelvetSweatshop"); f.decrypt(buf); buf.seek(0)
            book = xlrd.open_workbook(file_contents=buf.read())
        ws = book.sheet_by_name("Ativos") if "Ativos" in book.sheet_names() else book.sheet_by_index(0)
        data = [ws.row_values(i)[:8] for i in range(1, ws.nrows)]
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["Ativos"] if "Ativos" in wb.sheetnames else wb.worksheets[0]
        data = [[c for c in r[:8]] for r in ws.iter_rows(min_row=2, values_only=True)]
    df = pd.DataFrame(data, columns=["acct", "client", "address", "city", "community", "subdivision", "service", "obs"])
    df = df[df.client.astype(str).str.strip() != ""].reset_index(drop=True)
    return df.fillna("")

# ---------------------------------------------------------------- matching
def norm(x):
    x = str(x).lower().replace("&", " and ")
    x = re.sub(r"\b(llc|inc|the|at|of)\b", " ", x)
    x = re.sub(r"-\s*[a-z]{1,2}\b", " ", x)          # "- C", "- T", "- BT" tags in schedule names
    x = re.sub(r"\b\d+\b", " ", x)                    # house numbers / ids appended to names
    x = re.sub(r"[^a-z ]", " ", x)
    return re.sub(r"\s+", " ", x).strip()

def match(sched, contacts, thr=82):
    from rapidfuzz import process, fuzz
    cand = {}
    for i, r in contacts.iterrows():
        cand.setdefault(norm(r.client), []).append(i)
    keys = list(cand)
    idxs, scores = [], []
    for _, r in sched.iterrows():
        k = norm(r["name"])
        hit = process.extractOne(k, keys, scorer=lambda a, b, **kw: max(fuzz.token_sort_ratio(a, b), fuzz.token_set_ratio(a, b) - 8))
        if hit and hit[1] >= thr:
            ids = cand[hit[0]]
            if len(ids) > 1:
                ctx = norm(r["sub"] + " " + r["svc"])
                ids = [max(ids, key=lambda i: fuzz.partial_ratio(ctx, norm(contacts.loc[i, "city"] + " " + contacts.loc[i, "community"])))]
            idxs.append(ids[0]); scores.append(round(hit[1]))
        else:
            idxs.append(None); scores.append(0)
    out = sched.copy(); out["cidx"] = idxs; out["match_score"] = scores
    return out

# ---------------------------------------------------------------- geocoding
def _jload(name):
    p = os.path.join(CACHE, name)
    return json.load(open(p)) if os.path.exists(p) else {}
def _jsave(name, obj):
    json.dump(obj, open(os.path.join(CACHE, name), "w"))

def clean_addr(a):
    a = re.sub(r"\b(apt|unit|ste|suite)\.?\s*\S+|#\s*\S+", "", str(a), flags=re.I)
    a = re.sub(r"\bCirc\b", "Circle", a, flags=re.I)
    return re.sub(r"\s+", " ", a).strip(" ,")

def in_box(lat, lon):
    return BBOX[0] <= lat <= BBOX[1] and BBOX[2] <= lon <= BBOX[3]

def census_batch(items):
    """items: {key: (street, city)} -> {key: (lat, lon)}  (Census batch endpoint, free, no key)"""
    import requests
    res = {}; keys = list(items)
    for i in range(0, len(keys), 1000):
        chunk = keys[i:i + 1000]
        csv = "\n".join(f'{j},"{items[k][0]}","{items[k][1]}",FL,' for j, k in zip(range(len(chunk)), chunk))
        try:
            r = requests.post("https://geocoding.geo.census.gov/geocoder/locations/addressbatch",
                              files={"addressFile": ("a.csv", csv)}, data={"benchmark": "Public_AR_Current"}, timeout=180, headers=UA)
            r.raise_for_status()
        except Exception as e:
            print("  census batch unavailable:", str(e)[:80]); return res
        for line in r.text.splitlines():
            parts = next(__import__("csv").reader([line]))
            if len(parts) >= 6 and parts[2] == "Match":
                lon, lat = map(float, parts[5].split(","))
                if in_box(lat, lon): res[chunk[int(parts[0])]] = (lat, lon)
    return res

def nominatim(q):
    url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode({"q": q, "format": "json", "limit": 1, "countrycodes": "us"})
    try:
        d = json.load(urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30))
        time.sleep(1.1)                                   # OSM usage policy: max 1 request/second
        if d and in_box(float(d[0]["lat"]), float(d[0]["lon"])): return float(d[0]["lat"]), float(d[0]["lon"])
    except Exception as e:
        print("  nominatim error:", str(e)[:80])
    return None

def geocode_all(stops, dry=False):
    """stops has columns gkey, gstreet, gcity, gfallback. Adds lat/lon."""
    cache = _jload("geocode.json")
    need = stops.drop_duplicates("gkey")
    need = need[~need.gkey.isin(cache)]
    if len(need):
        print(f"Geocoding {len(need)} new addresses (cached: {len(cache)})...")
        if dry:
            for _, r in need.iterrows():
                h = random.Random(r.gkey); cache[r.gkey] = [26.10 + h.uniform(-.06, .06), -80.36 + h.uniform(-.08, .08)]
        else:
            got = census_batch({r.gkey: (r.gstreet, r.gcity) for _, r in need.iterrows() if r.gstreet})
            for k, v in got.items(): cache[k] = list(v)
            miss = need[~need.gkey.isin(cache)]
            if len(miss): print(f"  Census matched {len(got)}; trying OpenStreetMap for {len(miss)}...")
            for _, r in miss.iterrows():
                v = nominatim(f"{r.gstreet}, {r.gcity}, FL") if r.gstreet else None
                v = v or nominatim(f"{r.gfallback}, FL")
                if v: cache[r.gkey] = list(v)
                else: print("  could not locate:", r.gkey)
        _jsave("geocode.json", cache)
    stops = stops.copy()
    stops["lat"] = stops.gkey.map(lambda k: cache.get(k, [None, None])[0])
    stops["lon"] = stops.gkey.map(lambda k: cache.get(k, [None, None])[1])
    return stops

# ---------------------------------------------------------------- distances
def haversine(a, b):
    R = 6371000; p1, p2 = math.radians(a[0]), math.radians(b[0]); dl = math.radians(b[1] - a[1])
    h = math.sin((p2 - p1) / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.asin(math.sqrt(h))

def matrices(pts, dry=False):
    """returns (meters, seconds, source). Real roads via OSRM when reachable."""
    n = len(pts); ck = "|".join(f"{p[0]:.5f},{p[1]:.5f}" for p in pts)
    mc = _jload("matrix.json")
    if ck in mc: return mc[ck]["m"], mc[ck]["s"], mc[ck]["src"]
    m = s = None; src = "road"
    if not dry:
        try:
            coords = ";".join(f"{p[1]:.6f},{p[0]:.6f}" for p in pts)
            d = json.load(urllib.request.urlopen(urllib.request.Request(
                f"https://router.project-osrm.org/table/v1/driving/{coords}?annotations=distance,duration", headers=UA), timeout=60))
            if d.get("code") == "Ok": m, s = d["distances"], d["durations"]; time.sleep(0.3)
        except Exception as e:
            print("  OSRM unavailable, using straight-line estimate:", str(e)[:60])
    if m is None:
        src = "estimate"
        m = [[haversine(a, b) * 1.35 for b in pts] for a in pts]; s = [[x / 11.2 for x in row] for row in m]   # ~25 mph
    if src == "road": mc[ck] = {"m": m, "s": s, "src": src}; _jsave("matrix.json", mc)
    return m, s, src

# ---------------------------------------------------------------- optimizer
def tour_cost(t, D): return sum(D[t[i]][t[(i + 1) % len(t)]] for i in range(len(t)))

def solve_tour(D, restarts=60, seed=1):
    """closed TSP. exact (Held-Karp) up to 12 nodes, else multi-start 2-opt + or-opt."""
    n = len(D)
    if n <= 3: return list(range(n))
    if n <= 12:
        INF = float("inf"); full = 1 << (n - 1)
        dp = [[INF] * (n - 1) for _ in range(full)]; par = [[-1] * (n - 1) for _ in range(full)]
        for j in range(n - 1): dp[1 << j][j] = D[n - 1][j]
        for mask in range(1, full):
            for j in range(n - 1):
                if not mask >> j & 1 or dp[mask][j] == INF: continue
                for k in range(n - 1):
                    if mask >> k & 1: continue
                    nm = mask | 1 << k; c = dp[mask][j] + D[j][k]
                    if c < dp[nm][k]: dp[nm][k] = c; par[nm][k] = j
        j = min(range(n - 1), key=lambda j: dp[full - 1][j] + D[j][n - 1]); mask = full - 1; path = []
        while j != -1: path.append(j); pj = par[mask][j]; mask ^= 1 << j; j = pj
        return [n - 1] + path
    rng = random.Random(seed); best = None; bc = float("inf")
    for r in range(restarts):
        t = list(range(n)); rng.shuffle(t) if r else None
        if r == 0:                                         # nearest-neighbour seed
            t = [0]; left = set(range(1, n))
            while left: x = min(left, key=lambda y: D[t[-1]][y]); t.append(x); left.remove(x)
        imp = True
        while imp:
            imp = False
            for i in range(1, n - 1):                      # 2-opt
                for j in range(i + 1, n):
                    a, b, c, d = t[i - 1], t[i], t[j], t[(j + 1) % n]
                    if D[a][c] + D[b][d] < D[a][b] + D[c][d] - 1e-9: t[i:j + 1] = reversed(t[i:j + 1]); imp = True
            for L in (1, 2, 3):                            # or-opt
                for i in range(1, n - L + 1):
                    seg = t[i:i + L]; rest = t[:i] + t[i + L:]; base = tour_cost(t, D)
                    for p in range(1, len(rest) + 1):
                        for sg in (seg, seg[::-1]):
                            cand = rest[:p] + sg + rest[p:]
                            if tour_cost(cand, D) < base - 1e-9: t = cand; imp = True; break
                        if imp: break
                    if imp: break
                if imp: break
        c = tour_cost(t, D)
        if c < bc: bc, best = c, t[:]
    return best

def optimize(D, depot=False):
    """D includes depot at index 0 if depot else not. Returns visiting order of stop indices (0-based over stops)."""
    n = len(D)
    if not depot:   # open path: add a zero-cost dummy node so the closed solver gives a free start and end
        D2 = [row + [0] for row in D] + [[0] * (n + 1)]
    else:
        D2 = D
    t = solve_tour(D2); z = (n if not depot else 0)
    k = t.index(z); t = t[k:] + t[:k]; t = t[1:]
    return [x - 1 for x in t] if depot else t

def evaluate(order, M, S, depot):
    seq = ([0] if depot else []) + ([o + 1 for o in order] if depot else order) + ([0] if depot else [])
    legs = [(M[a][b] / 1609.344, S[a][b] / 60) for a, b in zip(seq, seq[1:])]
    return legs

# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("schedule", nargs="?", help="schedule workbook; omit to use the newest SCHEDULE*.xlsx in this folder")
    ap.add_argument("contacts", nargs="?", help="contacts workbook; omit to use the newest CONTACTS*.xls/.xlsx in this folder")
    ap.add_argument("--depot", default=DEFAULT_DEPOT, help=f'start/end address (default: "{DEFAULT_DEPOT}"). Use --depot none for a free start/end')
    ap.add_argument("--service-min", type=float, default=25, help="minutes spent at each pool (default 25)")
    ap.add_argument("--only", help="only this tech (e.g. MIKA)"); ap.add_argument("--day", help="only this day (MON..SAT)")
    ap.add_argument("--out", default=None); ap.add_argument("--dry-run", action="store_true", help="TEST ONLY: fake coordinates")
    a = ap.parse_args()

    if a.depot and a.depot.lower() == "none": a.depot = None
    a.schedule = a.schedule or find_latest("SCHEDULE", (".xlsx", ".xlsm"))
    a.contacts = a.contacts or find_latest("CONTACTS", (".xls", ".xlsx"))
    print("schedule:", os.path.basename(a.schedule), "| contacts:", os.path.basename(a.contacts), "| depot:", a.depot or "none")
    sched = load_schedule(a.schedule); contacts = load_contacts(a.contacts)
    if a.only: sched = sched[sched.tech.str.upper().str.contains(a.only.upper())]
    if a.day: sched = sched[sched.day == a.day.upper()]
    print(f"{len(sched)} stops, {sched.tech.nunique()} techs, {len(contacts)} contacts")
    m = match(sched, contacts)
    def info(r, col): return contacts.loc[int(r.cidx), col] if pd.notna(r.cidx) else ""
    m["address"] = m.apply(lambda r: clean_addr(info(r, "address")), axis=1)
    m["city"] = m.apply(lambda r: str(info(r, "city")).strip() or "", axis=1)
    m["obs"] = m.apply(lambda r: str(info(r, "obs")).strip(), axis=1)
    m["contact"] = m.apply(lambda r: info(r, "client"), axis=1)
    m["gstreet"] = m.address; m["gcity"] = m.city
    m["gkey"] = m.apply(lambda r: f"{r.address}|{r.city}".lower() if r.address else f"{r['name']}|{r['sub']}".lower(), axis=1)
    m["gfallback"] = m.apply(lambda r: f"{r['name']}, {r['sub']}", axis=1)
    m = geocode_all(m, a.dry_run)

    depot_pt = None
    if a.depot:
        d = m.iloc[[0]].copy(); d["gkey"] = a.depot.lower(); d["gstreet"] = a.depot; d["gcity"] = ""; d["gfallback"] = a.depot
        depot_pt = geocode_all(d, a.dry_run).iloc[0]
        if pd.isna(depot_pt.lat): sys.exit("Could not locate the depot address - check it or your internet connection.")
        depot_pt = (depot_pt.lat, depot_pt.lon)

    results = {"MILES": [], "TIME": []}; summary = []; check = []
    for (tech, day), g in m.groupby(["tech", "day"], sort=False):
        g = g.sort_values("seq"); ok = g[g.lat.notna()].reset_index(drop=True); bad = g[g.lat.isna()]
        for _, r in bad.iterrows(): check.append((tech, day, r["name"], "NO LOCATION FOUND - add address to contacts"))
        for _, r in g[(g.cidx.notna()) & (g.match_score < 92)].iterrows():
            check.append((tech, day, r["name"], f"low-confidence match -> '{r.contact}' ({r.match_score}%) - verify"))
        for _, r in g[g.cidx.isna()].iterrows(): check.append((tech, day, r["name"], "not in contacts; located by name/community only"))
        if len(ok) < 2:      # nothing to reorder: list as-is
            for obj in results:
                for _, s in g.iterrows():
                    results[obj].append(dict(Tech=tech, Day=day, Order=1, Client=s["name"], Address=s.address or "(by community)", City=s.city,
                        Community=s["sub"], Service=s.svc, Price=s.price, Notes=s.obs, LegMiles=0, LegMin=0, ClockMin=0, PrevOrder=int(s.seq)))
            continue
        pts = ([depot_pt] if depot_pt else []) + list(zip(ok.lat, ok.lon))
        M, S, src = matrices(pts, a.dry_run); dep = depot_pt is not None
        cur = list(range(len(ok))); cl = evaluate(cur, M, S, dep)
        row = dict(tech=tech, day=day, stops=len(g), located=len(ok), dist_src=src,
                   current_mi=sum(x[0] for x in cl), current_drive_min=sum(x[1] for x in cl))
        for obj, mat in (("MILES", M), ("TIME", S)):
            order = optimize(mat, dep)      # with a depot, index 0 of the matrix is the depot
            legs = evaluate(order, M, S, dep)
            ll = legs if dep else [(0, 0)] + legs      # leg INTO each stop (first stop has none without a depot)
            row[f"{obj.lower()}_mi"] = sum(x[0] for x in legs); row[f"{obj.lower()}_drive_min"] = sum(x[1] for x in legs)
            cum = 0
            for n_, (oi, lg) in enumerate(zip(order, ll), 1):
                s = ok.iloc[oi]; cum += lg[1] + a.service_min
                results[obj].append(dict(Tech=tech, Day=day, Order=n_, Client=s["name"], Address=s.address or "(by community)", City=s.city,
                    Community=s["sub"], Service=s.svc, Price=s.price, Notes=s.obs, LegMiles=round(lg[0], 1), LegMin=round(lg[1]),
                    ClockMin=round(cum), PrevOrder=int(s.seq)))
            for _, s in bad.iterrows():
                results[obj].append(dict(Tech=tech, Day=day, Order="?", Client=s["name"], Address="NO LOCATION", City="", Community=s["sub"],
                    Service=s.svc, Price=s.price, Notes="", LegMiles=None, LegMin=None, ClockMin=None, PrevOrder=int(s.seq)))
        summary.append(row)
        print(f"  {tech:14s}{day}  {len(ok):2d} stops  now {row['current_mi']:5.1f} mi -> {row['miles_mi']:5.1f} mi (min-miles) | "
              f"{row['current_drive_min']:4.0f} -> {row['time_drive_min']:4.0f} min (min-time)")

    S_ = pd.DataFrame(summary)
    if not S_.empty:
        S_["miles_saved"] = S_.current_mi - S_.miles_mi; S_["min_saved"] = S_.current_drive_min - S_.time_drive_min
        S_ = S_.round(1)
        tot = S_[["current_mi", "miles_mi", "time_mi", "current_drive_min", "miles_drive_min", "time_drive_min"]].sum()
        print(f"\nWEEK TOTAL: {tot.current_mi:.0f} mi now -> {tot.miles_mi:.0f} mi (min-miles)   |   {tot.current_drive_min/60:.1f} h now -> {tot.time_drive_min/60:.1f} h (min-time)")
    out = a.out or f"optimized_routes_{time.strftime('%Y%m%d')}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        S_.to_excel(xw, sheet_name="Summary", index=False)
        for obj, nm in (("MILES", "Shortest Miles"), ("TIME", "Fastest Time")):
            df = pd.DataFrame(results[obj]); df.to_excel(xw, sheet_name=nm, index=False)
            if not df.empty:   # one Google Maps link per tech/day
                links = []
                for (t, d), g in df[df.Order != "?"].groupby(["Tech", "Day"], sort=False):
                    path = "/".join(urllib.parse.quote(f"{r.Address}, {r.City} FL") for r in g.itertuples())
                    links.append(dict(Tech=t, Day=d, MapsLink="https://www.google.com/maps/dir/" + path))
                pd.DataFrame(links).to_excel(xw, sheet_name=nm + " (maps)", index=False)
        pd.DataFrame(check, columns=["Tech", "Day", "Stop", "Issue"]).to_excel(xw, sheet_name="Check These", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = min(40, max(10, max(len(str(c.value or "")) for c in col[:60]) + 2))
    print("wrote", out, f"({len(check)} items to check)")
    if a.dry_run: print("*** DRY RUN: coordinates are FAKE. Do not use these results. ***")

if __name__ == "__main__":
    main()
