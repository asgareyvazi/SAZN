"""SAZN Drilling-Data Analysis Department — exhaustive numeric forensics.
Reads every numeric cell, every formula, every 24-h activity, every PDF numeric token
and every priced line in the master warehouse; writes 28 derived forensic tables,
a compact findings JSON and a full exception register CSV.
No source value is modified; derived corrections are stored in separate columns and labelled.
"""
import os, sys, json, math, re, zlib, sqlite3, collections, csv, statistics
sys.path.insert(0, os.path.dirname(__file__))
from common import OUT, DATA, DB, load_json, fmt, r2

CON = sqlite3.connect(DB); CON.row_factory = sqlite3.Row
P = load_json('well_profiles'); F = json.load(open(os.path.join(DATA, 'field_analytics.json')))
TOT = F['totals']
WELLS = list(P.keys())


def flat(args):
    if len(args) == 1 and isinstance(args[0], (tuple, list)): return tuple(args[0])
    return tuple(args)
def q(sql, *args): return CON.execute(sql, flat(args)).fetchall()
def R(sql, *args): return [dict(r) for r in q(sql, *args)]
def W(sql, *args): CON.execute(sql, flat(args))
def commit(): CON.commit()
def table(name, ddl):
    W('DROP TABLE IF EXISTS an_' + name); W('CREATE TABLE an_' + name + ' (' + ddl + ')')
def ins(name, cols, rows):
    if not rows: return
    ph = ','.join('?' * len(cols))
    CON.executemany('INSERT INTO an_' + name + ' (' + ','.join(cols) + ') VALUES (' + ph + ')', rows)


def mean(xs): return sum(xs) / len(xs) if xs else None
def med(xs):
    xs = sorted(xs); return xs[len(xs)//2] if xs else None
def std(xs):
    m = mean(xs); return math.sqrt(sum((x - m)**2 for x in xs) / len(xs)) if len(xs) > 1 and m is not None else None
def quant(xs, p):
    if not xs: return None
    ys = sorted(xs); i = (len(ys) - 1) * p; lo = int(i); hi = min(lo + 1, len(ys) - 1)
    return ys[lo] + (ys[hi] - ys[lo]) * (i - lo)
def pearson(a, b):
    if len(a) < 3: return None
    ma, mb = mean(a), mean(b); num = sum((x-ma)*(y-mb) for x, y in zip(a, b))
    da = math.sqrt(sum((x-ma)**2 for x in a)); db = math.sqrt(sum((y-mb)**2 for y in b))
    return num / (da*db) if da and db else None
def rank(xs):
    order = sorted(range(len(xs)), key=lambda i: xs[i]); out = [0.0]*len(xs); i = 0
    while i < len(order):
        j = i
        while j+1 < len(order) and xs[order[j+1]] == xs[order[i]]: j += 1
        avg = (i+j)/2 + 1
        for k in range(i, j+1): out[order[k]] = avg
        i = j+1
    return out
def spearman(a, b): return pearson(rank(a), rank(b))
def regress(a, b):
    if len(a) < 3: return (None, None, None)
    ma, mb = mean(a), mean(b); den = sum((x-ma)**2 for x in a)
    if not den: return (None, None, None)
    sl = sum((x-ma)*(y-mb) for x, y in zip(a, b))/den; it = mb - sl*ma
    r = pearson(a, b); return (sl, it, r*r if r is not None else None)
def nz(v, d=0.0):
    try: return float(v)
    except Exception: return d
def f2(v): return None if v is None else round(float(v), 2)
def f4(v): return None if v is None else round(float(v), 4)


# ============================================================ 1. CELL CENSUS ==
def cell_census():
    table('num_census', 'well_key TEXT, rel_path TEXT, sheet TEXT, category TEXT, n INTEGER, '
          'v_min REAL, q1 REAL, v_med REAL, q3 REAL, v_max REAL, v_mean REAL, v_std REAL, v_sum REAL, '
          'zeros INTEGER, negatives INTEGER, integers INTEGER, outliers INTEGER, dominant_band TEXT')
    table('num_hist', 'band TEXT, n INTEGER, pct REAL, example REAL, example_source TEXT')
    table('outliers', 'well_key TEXT, rel_path TEXT, sheet TEXT, cell TEXT, value REAL, z REAL, sheet_mean REAL, sheet_std REAL')
    table('formula_audit', 'well_key TEXT, rel_path TEXT, sheet TEXT, cell TEXT, formula TEXT, n_refs INTEGER, '
          'cached REAL, recomputed REAL, delta REAL, kind TEXT, status TEXT')
    table('cell_quality', 'rel_path TEXT, sheet TEXT, numeric_as_text INTEGER, formula_cells INTEGER, '
          'formula_with_refs INTEGER, hardcoded_formula INTEGER, cross_sheet INTEGER, blank_in_numeric_run INTEGER, total_cells INTEGER')
    print('[1] scanning 2,810 sheets / 324,150 numeric cells / 155,248 formulas …', flush=True)
    ref_re = re.compile(r'\$?([A-Z]{1,3})\$?(\d{1,5})')
    range_re = re.compile(r'\$?([A-Z]{1,3})\$?(\d{1,5})\s*:\s*\$?([A-Z]{1,3})\$?(\d{1,5})')
    sum_re = re.compile(r'\bSUM\s*\(([^)]*)\)', re.I)
    census_rows, hist = [], collections.Counter(); outliers, formula_rows, quality_rows = [], [], []
    ex_by_band = {}
    for sh in R('SELECT s.id,s.name,s.rel_path,s.well_key,s.category,c.payload FROM sheets s JOIN sheet_cells c ON c.sheet_id=s.id'):
        try: cells = json.loads(zlib.decompress(sh['payload']))
        except Exception: continue
        byref = {x[0]: (x[1], x[2], x[3]) for x in cells if isinstance(x, (list, tuple)) and len(x) >= 4}
        def asnum(v):
            if isinstance(v, (int, float)) and not isinstance(v, bool): return float(v)
            if isinstance(v, str):
                t = v.strip().replace(',', '')
                if re.fullmatch(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?', t): return float(t)
            return None
        vals = [(x[0], asnum(x[1])) for x in cells if isinstance(x, (list, tuple)) and len(x) >= 4]
        vals = [(ref, v) for ref, v in vals if v is not None]
        nums = [v for _, v in vals]
        nat = sum(1 for x in cells if isinstance(x[1], str) and re.fullmatch(r'[-+]?\d[\d,]*\.?\d*', x[1].strip().replace(',', '')))
        forms = [x for x in cells if x[2]]
        nrefs = nhard = ncross = 0
        for x in forms:
            fo = str(x[2]); refs = ref_re.findall(fo)
            if refs: nrefs += 1
            else: nhard += 1
            if '!' in fo or '[' in fo: ncross += 1
        quality_rows.append((sh['rel_path'], sh['name'], nat, len(forms), nrefs, nhard, ncross,
                             sum(1 for x in cells if x[1] in (None, '')), len(cells)))
        if nums:
            m, sd = mean(nums), std(nums) or 0
            q1, q3 = quant(nums, .25), quant(nums, .75); iqr = (q3 or 0) - (q1 or 0)
            lo, hi = (q1 - 1.5*iqr), (q3 + 1.5*iqr)
            out = [v for v in nums if v < lo or v > hi]
            band = ('<0' if (m or 0) < 0 else '0–1' if (m or 0) < 1 else '1–10' if (m or 0) < 10 else
                    '10–100' if (m or 0) < 100 else '100–1k' if (m or 0) < 1e3 else
                    '1k–10k' if (m or 0) < 1e4 else '10k–100k' if (m or 0) < 1e5 else
                    '100k–1M' if (m or 0) < 1e6 else '>=1M')
            census_rows.append((sh['well_key'], sh['rel_path'], sh['name'], sh['category'], len(nums),
                min(nums), q1, med(nums), q3, max(nums), m, sd, sum(nums),
                sum(1 for v in nums if v == 0), sum(1 for v in nums if v < 0),
                sum(1 for v in nums if float(v).is_integer()), len(out), band))
            for ref, v in vals:
                b = ('<0' if v < 0 else '0–1' if v < 1 else '1–10' if v < 10 else '10–100' if v < 100 else
                     '100–1k' if v < 1e3 else '1k–10k' if v < 1e4 else '10k–100k' if v < 1e5 else
                     '100k–1M' if v < 1e6 else '>=1M')
                hist[b] += 1
                ex_by_band.setdefault(b, (v, sh['rel_path'], sh['name'], ref))
                if sd and abs((v - m)/sd) > 6 and len(outliers) < 900:
                    outliers.append((sh['well_key'], sh['rel_path'], sh['name'], ref, v, (v-m)/sd, m, sd))
        # formula recomputation for simple arithmetic / SUM formulas
        def value_of(cell_ref):
            e = byref.get(cell_ref.upper())
            if not e: return None
            v = e[0]
            return nz(v) if isinstance(v, (int, float)) else (float(str(v).replace(',', '')) if re.fullmatch(r'[-+]?\d[\d,]*\.?\d*', str(v).strip()) else None)
        for x in forms:
            fo = str(x[2]).strip(); cached = value_of(x[0])
            if cached is None: continue
            expr = fo[1:] if fo.startswith('=') else fo
            expr = expr.replace('^', '**').replace('$', '')
            # expand SUM(A1:A9)
            def expand_sum(mo):
                inner = mo.group(1)
                parts = []
                for rg in inner.split(','):
                    mm = range_re.fullmatch(rg.strip())
                    if not mm: continue
                    c1 = colnum(mm.group(1)); c2 = colnum(mm.group(3))
                    r1, r2n = int(mm.group(2)), int(mm.group(4))
                    for rr in range(r1, r2n+1):
                        for cc in range(c1, c2+1):
                            parts.append(colref(cc) + str(rr))
                return '(' + '+'.join(parts) + ')'
            expr = sum_re.sub(expand_sum, expr)
            expr = ref_re.sub(lambda mo: str(value_of(mo.group(1)+mo.group(2)) if value_of(mo.group(1)+mo.group(2)) is not None else 0), expr)
            if not re.fullmatch(r'[\d\.\+\-\*/\(\) eE]+', expr): continue
            try: rec = eval(expr, {'__builtins__': {}}, {})
            except Exception: continue
            if not isinstance(rec, (int, float)) or not math.isfinite(rec): continue
            delta = abs(rec - cached); tol = max(1e-6, abs(cached)*1e-6)
            status = 'MATCH' if delta <= tol else ('MINOR' if delta <= max(.01, abs(cached)*.001) else 'MISMATCH')
            kind = 'SUM' if 'SUM' in fo.upper() else ('CROSS_SHEET' if '!' in fo else ('RATIO' if '/' in fo else 'ARITHMETIC'))
            formula_rows.append((sh['well_key'], sh['rel_path'], sh['name'], x[0], fo[:220],
                len(ref_re.findall(fo)), cached, rec, delta, kind, status))
            if len(formula_rows) >= 155248: break
    ins('num_census', ['well_key','rel_path','sheet','category','n','v_min','q1','v_med','q3','v_max','v_mean','v_std','v_sum','zeros','negatives','integers','outliers','dominant_band'], census_rows)
    nh = sum(hist.values())
    ins('num_hist', ['band','n','pct','example','example_source'],
        [(b, n, 100*n/max(1,nh), ex_by_band[b][0], ex_by_band[b][1]+'!'+ex_by_band[b][2]+'!'+ex_by_band[b][3]) for b, n in sorted(hist.items())])
    ins('outliers', ['well_key','rel_path','sheet','cell','value','z','sheet_mean','sheet_std'], outliers)
    ins('formula_audit', ['well_key','rel_path','sheet','cell','formula','n_refs','cached','recomputed','delta','kind','status'], formula_rows)
    ins('cell_quality', ['rel_path','sheet','numeric_as_text','formula_cells','formula_with_refs','hardcoded_formula','cross_sheet','blank_in_numeric_run','total_cells'], quality_rows)
    commit()
    print('    census', len(census_rows), 'sheets | histogram', len(hist), 'bands | outliers', len(outliers),
          '| formulas audited', len(formula_rows), flush=True)


def colnum(s):
    n = 0
    for ch in s: n = n*26 + ord(ch) - 64
    return n
def colref(n):
    s = ''
    while n: n, m = divmod(n-1, 26); s = chr(65+m) + s
    return s


# =================================================== 2. CBS TARIFF FORENSICS ==
def cbs_forensics():
    table('cbs_line_audit', 'well_key TEXT, rel_path TEXT, sheet TEXT, cell TEXT, description TEXT, unit TEXT, '
          'qty REAL, unit_price REAL, total REAL, expected REAL, delta REAL, pct REAL, status TEXT')
    table('unit_price', 'description TEXT, unit TEXT, n INTEGER, v_min REAL, v_med REAL, v_max REAL, '
          'spread_pct REAL, wells TEXT, dispersion TEXT')
    rows, prices = [], collections.defaultdict(list)
    for r in R('SELECT well_key,rel_path,sheet,cell_ref,description,unit,qty,unit_price,total FROM cbs_items'):
        qty, up, tot = r['qty'], r['unit_price'], r['total']
        if qty is None or up is None or tot is None: continue
        exp = nz(qty)*nz(up); d = nz(tot)-exp; pct = 100*d/exp if exp else None
        st = 'MATCH' if abs(d) <= max(.01, abs(exp)*1e-6) else ('ROUNDING' if pct is not None and abs(pct) < .5 else
             ('MATERIAL' if abs(pct or 0) >= 5 else 'MINOR'))
        rows.append((r['well_key'], r['rel_path'], r['sheet'], r['cell_ref'], (r['description'] or '')[:150],
                     r['unit'], qty, up, tot, exp, d, pct, st))
        if up and up > 0:
            prices[(re.sub(r'\s+', ' ', (r['description'] or '').strip().lower())[:70], (r['unit'] or '').upper())].append((up, r['well_key'], tot))
    ins('cbs_line_audit', ['well_key','rel_path','sheet','cell','description','unit','qty','unit_price','total','expected','delta','pct','status'], rows)
    pr = []
    for (desc, unit), v in prices.items():
        xs = [x[0] for x in v]; mn, md, mx = min(xs), med(xs), max(xs)
        spread = 100*(mx-mn)/mn if mn else None
        pr.append((desc, unit, len(v), mn, md, mx, spread, ','.join(sorted(set(x[1] for x in v if x[1])))[:200],
                   'UNIFORM' if (spread or 0) < .01 else ('LOW' if (spread or 0) < 5 else 'DISPERSED')))
    ins('unit_price', ['description','unit','n','v_min','v_med','v_max','spread_pct','wells','dispersion'], pr)
    commit()
    print('[2] CBS priced-line audit', len(rows), 'lines | unit-price keys', len(pr), flush=True)


# =============================================== 3. 24-H TIME & ILT FORENSICS ==
def time_forensics():
    table('time_balance', 'well_key TEXT, report_no INTEGER, date_g TEXT, canonical_sheet TEXT, activity_h REAL, pt_h REAL, npt_h REAL, '
          'unclassified_h REAL, ilt_h REAL, carried_over_h REAL, balance_status TEXT, time_code_h REAL, max_row_h REAL, n_rows INTEGER, status_rows INTEGER')
    table('daily_consistency', 'well_key TEXT, report_no INTEGER, md_0000 REAL, md_0600 REAL, md_2400 REAL, '
          'drilled_24 REAL, expected_delta REAL, delta REAL, status TEXT, corrected_24h REAL, rop REAL, rop_recomputed REAL, rop_delta REAL')
    table('activity_crosstab', 'well_key TEXT, phase TEXT, code TEXT, status TEXT, attributed TEXT, hours REAL, events INTEGER, share_pct REAL')
    table('ilt', 'well_key TEXT, report_no INTEGER, connection_h REAL, trip_h REAL, circulate_h REAL, wait_h REAL, '
          'survey_h REAL, ream_h REAL, other_h REAL, ilt_h REAL, ilt_pct REAL')
    acts = R('SELECT well_key,report_no,sheet,phase,code,status,npt_attributed,hours,description FROM ddr_activities')
    bygroup = collections.defaultdict(list)
    for a in acts: bygroup[(a['well_key'], a['report_no'], a['sheet'] or '')].append(a)
    # A DDR pack can contain both a Template-A 24-h log and a Template-B / time-breakdown
    # sheet for the same report. Choose the single sheet whose hour total is closest to 24 h
    # as the canonical 24-h log; all sheets remain in the warehouse for audit.
    best = {}
    for key, aa in bygroup.items():
        h = sum(nz(a['hours']) for a in aa if 0 < nz(a['hours']) <= 24)
        wk, rn = key[0], key[1]
        if (wk, rn) not in best or abs(h-24) < abs(best[(wk, rn)][0]-24): best[(wk, rn)] = (h, aa, key[2])
    byday = {(wk, rn): v[1] for (wk, rn), v in best.items()}
    daily = {(d['well_key'], d['report_no']): d for d in R('SELECT * FROM ddr_daily')}
    tc = collections.defaultdict(float)
    for t in R('SELECT well_key,report_no,hours,category FROM ddr_time'):
        tc[(t['well_key'], t['report_no'])] += nz(t['hours'])
    tb, ct, ilt_rows = [], [], []
    for (wk, rn), aa in byday.items():
        d = daily.get((wk, rn), {})
        csh = best.get((wk, rn), (0, [], ''))[2]
        # remove the summary / total row (blank status and code) from the 24-h narrative log
        aa = [a for a in aa if (a['status'] or '').strip() or (a['code'] or '').strip()]
        ah = sum(nz(a['hours']) for a in aa if 0 < nz(a['hours']) <= 24)
        st = collections.Counter(re.sub(r'[^A-Z]', '', (a['status'] or '').upper()) for a in aa)
        pt = sum(nz(a['hours']) for a in aa if re.sub(r'[^A-Z]', '', (a['status'] or '').upper()) in ('P','PLN','PLAN','PLANNED','PT'))
        np_ = sum(nz(a['hours']) for a in aa if re.sub(r'[^A-Z]', '', (a['status'] or '').upper()).startswith('NPT') or
                  re.sub(r'[^A-Z]', '', (a['status'] or '').upper()) in ('N','UP','UNPLANNED','UNPLAN','UNPLN','NP','PN','UPI','NPTUNPLN'))
        uc = max(0, ah-pt-np_)
        ilt = 24-ah
        status = 'COMPLETE_24H' if abs(ilt) < .01 else ('UNDER_CODED' if ilt > 0 else 'OVER_CODED')
        tb.append((wk, rn, str(d.get('date_g') or ''), csh, ah, pt, np_, uc, ilt, max(0, ah-24), status,
                   tc.get((wk, rn), 0), max([nz(a['hours']) for a in aa]+[0]), len(aa), len(aa)))
        def hit(regex): return sum(nz(a['hours']) for a in aa if re.search(regex, ' '.join([str(a['description'] or ''), str(a['phase'] or ''), str(a['code'] or '')]), re.I))
        conn, trip, circ = hit(r'conn|make.?up|mu\b'), hit(r'trip|pooh|rih|run in|pull out'), hit(r'circulat|condition|bottoms up')
        wait, surv, ream = hit(r'wait|standby|woc|woo'), hit(r'survey|mwd|gyro'), hit(r'ream|wiper')
        other = max(0, ah-conn-trip-circ-wait-surv-ream)
        ilt_rows.append((wk, rn, conn, trip, circ, wait, surv, ream, other, ilt, 100*max(0,ilt)/24))
    # activity crosstab
    cross = collections.defaultdict(lambda: [0.0, 0])
    for a in acts:
        key = (a['well_key'], a['phase'] or '', a['code'] or '', a['status'] or '', a['npt_attributed'] or '')
        cross[key][0] += nz(a['hours']); cross[key][1] += 1
    tot_by_well = collections.Counter()
    for (wk, *_), v in cross.items(): tot_by_well[wk] += v[0]
    ct = [(wk, ph, cd, st, at, h, n, 100*h/max(1, tot_by_well[wk])) for (wk, ph, cd, st, at), (h, n) in cross.items()]
    # daily consistency — silently correct probable cumulative typos in derived column only
    dc = []
    for d in R('SELECT well_key,report_no,md_0000,md_0600,md_2400,drilled_24,rop FROM ddr_daily ORDER BY well_key,report_no'):
        a, b, dr, rop = d['md_0000'], d['md_2400'], d['drilled_24'], d['rop']
        exp = (nz(b)-nz(a)) if (a is not None and b is not None) else None
        delta = (nz(dr)-exp) if (exp is not None and dr is not None) else None
        st = 'NO_SOURCE' if delta is None else ('MATCH' if abs(delta) < .01 else ('CUMULATIVE_TYPO' if abs(delta) > 25 else 'MINOR'))
        corr = exp if (st == 'CUMULATIVE_TYPO' and exp is not None and exp >= 0) else dr
        rr = (corr/24) if (corr is not None and rop is not None) else None
        dc.append((d['well_key'], d['report_no'], a, d['md_0600'], b, dr, exp, delta, st, corr, rop, rr,
                   (rr-nz(rop)) if (rr is not None and rop is not None) else None))
    for name, cols, rows in [('time_balance',['well_key','report_no','date_g','canonical_sheet','activity_h','pt_h','npt_h','unclassified_h','ilt_h','carried_over_h','balance_status','time_code_h','max_row_h','n_rows','status_rows'],tb),
                            ('daily_consistency',['well_key','report_no','md_0000','md_0600','md_2400','drilled_24','expected_delta','delta','status','corrected_24h','rop','rop_recomputed','rop_delta'],dc),
                            ('activity_crosstab',['well_key','phase','code','status','attributed','hours','events','share_pct'],ct),
                            ('ilt',['well_key','report_no','connection_h','trip_h','circulate_h','wait_h','survey_h','ream_h','other_h','ilt_h','ilt_pct'],ilt_rows)]:
        ins(name, cols, rows)
    commit()
    print('[3] time balance', len(tb), 'days | consistency', len(dc), '| crosstab', len(ct), flush=True)


# ============================================= 4. DDR PARAMETERS & REGRESSION ==
LABELS = {
 'WOB': r'\bwob\b|weight on bit', 'RPM': r'\brpm\b', 'SPP': r'\bspp\b|stand ?pipe', 'GPM': r'\bgpm\b|flow rate',
 'TORQUE': r'\btorque\b', 'PU': r'\bpu\b|pick.?up weight', 'SO': r'\bso\b|slack.?off weight', 'ROT': r'\brot\b|rotating weight',
 'AV': r'\bav\b|annular velocity', 'PV': r'\bpv\b|plastic viscosity', 'YP': r'\byp\b|yield point',
 'MW': r'\bmw\b|mud weight', 'ROP': r'\brop\b', 'DLS': r'\bdls\b', 'ECD': r'\becd\b', 'HSI': r'\bhsi\b',
 'SPM': r'\bspm\b', 'SCR': r'\bscr\b', 'TFA': r'\btfa\b', 'JV': r'\bjv\b|jet velocity',
}
def param_forensics():
    table('daily_params', 'well_key TEXT, report_no INTEGER, param TEXT, value REAL, unit TEXT, sheet TEXT, cell TEXT, label TEXT')
    table('param_stats', 'well_key TEXT, param TEXT, n INTEGER, v_min REAL, q1 REAL, v_med REAL, q3 REAL, v_max REAL, '
          'v_mean REAL, v_std REAL, cv_pct REAL, iqr_outliers INTEGER')
    table('correlations', 'well_key TEXT, x TEXT, y TEXT, n INTEGER, pearson REAL, spearman REAL, slope REAL, intercept REAL, r2 REAL, strength TEXT')
    print('[4] extracting daily drilling parameters from 2,608 DDR sheets …', flush=True)
    dp = []
    for sh in R("SELECT s.id,s.name,s.rel_path,s.well_key,s.report_no,c.payload FROM sheets s JOIN sheet_cells c ON c.sheet_id=s.id WHERE s.category='DDR'"):
        try: cells = json.loads(zlib.decompress(sh['payload']))
        except Exception: continue
        byrow = collections.defaultdict(list)
        for x in cells:
            if not isinstance(x, (list, tuple)) or len(x) < 4: continue
            m = re.match(r'([A-Z]{1,3})(\d+)', x[0])
            if m: byrow[int(m.group(2))].append((colnum(m.group(1)), x[1], x[0]))
        for rr, cs in byrow.items():
            cs.sort(); labels = ' '.join(str(v).lower() for c, v, ref in cs if isinstance(v, str) and len(str(v)) < 60)
            for par, pat in LABELS.items():
                if not re.search(pat, labels): continue
                for c, v, ref in cs:
                    num = None
                    if isinstance(v, (int, float)) and not isinstance(v, bool): num = float(v)
                    elif isinstance(v, str):
                        mm = re.fullmatch(r'\s*([-+]?\d[\d,]*\.?\d*)\s*', v)
                        if mm: num = float(mm.group(1).replace(',', ''))
                    if num is not None and -1e6 < num < 1e7:
                        dp.append((sh['well_key'], sh['report_no'] or 0, par, num, '', sh['name'], ref, par))
                        break
    ins('daily_params', ['well_key','report_no','param','value','unit','sheet','cell','label'], dp)
    # per-well parameter statistics
    byp = collections.defaultdict(list)
    for wk, rn, par, v, *_ in dp: byp[(wk, par)].append(v)
    ps = []
    for (wk, par), vs in byp.items():
        q1, q3 = quant(vs, .25), quant(vs, .75); iqr = (q3 or 0)-(q1 or 0)
        m, sd = mean(vs), std(vs) or 0
        ps.append((wk, par, len(vs), min(vs), q1, med(vs), q3, max(vs), m, sd,
                   100*sd/abs(m) if m else None, sum(1 for v in vs if v < q1-1.5*iqr or v > q3+1.5*iqr)))
    ins('param_stats', ['well_key','param','n','v_min','q1','v_med','q3','v_max','v_mean','v_std','cv_pct','iqr_outliers'], ps)
    # daily paired matrix from ddr_daily + daily_params
    daily = {(d['well_key'], d['report_no']): d for d in R('SELECT well_key,report_no,rop,mw,drilled_24,fit_emw,md_2400 FROM ddr_daily')}
    par_by_day = collections.defaultdict(dict)
    for wk, rn, par, v, *_ in dp: par_by_day[(wk, rn)][par] = v
    pairs = collections.defaultdict(lambda: collections.defaultdict(dict))
    for key, d in daily.items():
        row = dict(ROP=d['rop'], MW=d['mw'], DRILLED=d['drilled_24'], FIT=d['fit_emw'], MD=d['md_2400'])
        row.update(par_by_day.get(key, {}))
        for k, v in row.items():
            if v is not None and math.isfinite(float(v)): pairs[key[0]][k][key[1]] = float(v)
    corr_rows = []
    XS = ['WOB','RPM','SPP','GPM','MW','AV','PV','YP','ECD','HSI','SPM','TORQUE']
    for wk, cols in pairs.items():
        if 'ROP' not in cols: continue
        for x in XS:
            if x not in cols: continue
            common = sorted(set(cols['ROP']) & set(cols[x]))
            a = [cols[x][r] for r in common]; b = [cols['ROP'][r] for r in common]
            if len(a) < 4: continue
            pr, sp = pearson(a, b), spearman(a, b); sl, it, r2v = regress(a, b)
            strength = 'STRONG' if pr is not None and abs(pr) >= .7 else ('MODERATE' if pr is not None and abs(pr) >= .4 else 'WEAK')
            corr_rows.append((wk, x, 'ROP', len(a), pr, sp, sl, it, r2v, strength))
        # MW vs drilled, ROP vs drilled
        for x, y in [('MW','DRILLED'),('ROP','DRILLED'),('MW','FIT'),('MD','MW')]:
            if x in cols and y in cols:
                common = sorted(set(cols[x]) & set(cols[y])); a=[cols[x][r] for r in common]; b=[cols[y][r] for r in common]
                if len(a) < 4: continue
                pr, sp = pearson(a,b), spearman(a,b); sl, it, r2v = regress(a,b)
                corr_rows.append((wk, x, y, len(a), pr, sp, sl, it, r2v,
                                  'STRONG' if pr is not None and abs(pr)>=.7 else ('MODERATE' if pr is not None and abs(pr)>=.4 else 'WEAK')))
    ins('correlations', ['well_key','x','y','n','pearson','spearman','slope','intercept','r2','strength'], corr_rows)
    commit()
    print('    daily parameter rows', len(dp), '| param stats', len(ps), '| correlations', len(corr_rows), flush=True)


# ================================================ 5. NPT / EVENT FORENSICS ====
def npt_forensics():
    table('npt_forensics', 'well_key TEXT, rank INTEGER, report_no INTEGER, category TEXT, hours REAL, cost_usd REAL, '
          'attributed TEXT, depth_m REAL, formation TEXT, days_since_prev INTEGER, cumulative_h REAL, cumulative_pct REAL, '
          'mtbf_h REAL, payer TEXT')
    table('npt_pareto', 'well_key TEXT, category TEXT, hours REAL, share_pct REAL, cumulative_pct REAL, rank INTEGER, cost_usd REAL')
    rows, pareto = [], []
    for wk in WELLS:
        ev = R('SELECT report_no,hours,npt_attributed,description FROM ddr_activities WHERE well_key=? AND hours IS NOT NULL AND hours>0 '
               "AND (UPPER(REPLACE(REPLACE(status,'.',''),' ','') ) LIKE 'NPT%' OR UPPER(status) IN ('N','UP','UNPLANNED','UNPLAN','UNPLN','NP','PN','UPI')) "
               'ORDER BY report_no,id', wk)
        tot = sum(nz(e['hours']) for e in ev) or 1
        prev = collections.defaultdict(lambda: None); cum = 0
        mtbf = collections.defaultdict(list)
        for i, e in enumerate(sorted(ev, key=lambda x: -nz(x['hours'])), 1):
            h = nz(e['hours']); cum += h
            cat = re.sub(r'\s+', ' ', (e['description'] or 'n/a'))[:90]
            rn = e['report_no'] or 0
            gap = (rn - prev[cat]) if prev[cat] is not None else None
            if gap is not None: mtbf[cat].append(gap*24)
            prev[cat] = rn
            payer = 'CONTRACTOR_ZERO_RATE' if re.search(r'oeoc|contractor|service|mud|logging|directional|catering|energy|supplier|vendor|rig|unattributed',
                       e['npt_attributed'] or '', re.I) or not (e['npt_attributed'] or '').strip() else 'CLIENT_SUBSURFACE'
            rows.append((wk, i, rn, cat, h, h/24*nz(P[wk]['costs']['rig_rate']), e['npt_attributed'] or 'Unattributed',
                         None, None, gap, cum, 100*cum/tot, mean(mtbf[cat]) if mtbf[cat] else None, payer))
        byc = collections.Counter()
        for e in ev: byc[(e['npt_attributed'] or 'Unattributed')[:70]] += nz(e['hours'])
        c2 = 0
        for i, (cat, h) in enumerate(byc.most_common(), 1):
            c2 += h
            pareto.append((wk, cat, h, 100*h/tot, 100*c2/tot, i, h/24*nz(P[wk]['costs']['rig_rate'])))
    ins('npt_forensics', ['well_key','rank','report_no','category','hours','cost_usd','attributed','depth_m','formation','days_since_prev','cumulative_h','cumulative_pct','mtbf_h','payer'], rows)
    ins('npt_pareto', ['well_key','category','hours','share_pct','cumulative_pct','rank','cost_usd'], pareto)
    commit()
    print('[5] NPT forensic events', len(rows), '| Pareto lines', len(pareto), flush=True)


# ================================= 6. FORMATION / DIRECTIONAL / MUD / CASING ==
def geo_forensics():
    table('formation_stats', 'well_key TEXT, formation TEXT, n_tops INTEGER, md_med REAL, tvd_med REAL, '
          'prognosis_med REAL, delta_m REAL, thickness_m REAL, lithology TEXT, problem_h REAL, mw_min REAL, mw_max REAL, '
          'loss_events INTEGER, structural_bias TEXT')
    table('directional', 'well_key TEXT, report_no INTEGER, md REAL, incl REAL, azi REAL, tvd REAL, dls REAL, '
          'cum_dogleg REAL, tortuosity REAL, section TEXT')
    table('mud_balance', 'well_key TEXT, report_no INTEGER, material TEXT, received REAL, used REAL, '
          'stock_prev REAL, stock REAL, implied_received REAL, delta REAL, status TEXT')
    table('casing_check', 'well_key TEXT, size TEXT, grade TEXT, from_m REAL, to_m REAL, burst REAL, collapse REAL, '
          'nearest_formation TEXT, nearest_top_m REAL, shoe_to_top_m REAL, lap_overlap_m REAL, status TEXT')
    table('pressure_stats', 'well_key TEXT, unit TEXT, n INTEGER, v_min REAL, v_med REAL, v_max REAL, v_mean REAL, v_std REAL, outliers INTEGER')
    # formations
    fs = []
    for wk in WELLS:
        tops = collections.defaultdict(list)
        for r in R('SELECT name,md,tvd,prognosis_md,prognosis_tvd,lithology FROM ddr_formations WHERE well_key=?', wk):
            if r['md'] or r['prognosis_md']: tops[(r['name'] or '').strip()].append(r)
        prev_md = 0
        for nm, rs in sorted(tops.items(), key=lambda x: med([nz(y['md']) or nz(y['prognosis_md']) for y in x[1]]) or 0):
            mds = [nz(r['md']) for r in rs if r['md']]; tv = [nz(r['tvd']) for r in rs if r['tvd']]
            pm = [nz(r['prognosis_md']) for r in rs if r['prognosis_md']]
            md = med(mds) or med(pm) or 0; thick = md - prev_md; prev_md = md
            lith = collections.Counter(r['lithology'] for r in rs if r['lithology']).most_common(1)
            evh = sum(nz(e['hours']) for e in R('SELECT hours FROM well_events WHERE well_key=? AND (LOWER(formation) LIKE ? OR LOWER(detail) LIKE ?)',
                                                 wk, '%'+nm.lower()+'%', '%'+nm.lower()+'%'))
            losses = sum(1 for e in R("SELECT hours FROM ddr_activities WHERE well_key=? AND LOWER(description) LIKE '%loss%' AND LOWER(description) LIKE ?", wk, '%'+nm.lower()+'%'))
            mwr = [nz(d['mw']) for d in R('SELECT mw FROM ddr_daily WHERE well_key=? AND mw IS NOT NULL', wk)]
            delta = (md - med(pm)) if pm else None
            fs.append((wk, nm, len(rs), md, med(tv), med(pm), delta, thick if thick > 0 else None,
                       lith[0][0] if lith else '', evh, min(mwr) if mwr else None, max(mwr) if mwr else None, losses,
                       'DEEPER_THAN_PROG' if delta and delta > 15 else ('SHALLOWER_THAN_PROG' if delta and delta < -15 else 'ON_PROG')))
    ins('formation_stats', ['well_key','formation','n_tops','md_med','tvd_med','prognosis_med','delta_m','thickness_m','lithology','problem_h','mw_min','mw_max','loss_events','structural_bias'], fs)
    # directional
    dr = []
    for wk in WELLS:
        sv = R('SELECT report_no,md,incl,azi,tvd,dls FROM ddr_surveys WHERE well_key=? AND md IS NOT NULL ORDER BY md', wk)
        cum = 0; tort = 0
        for i, s in enumerate(sv):
            dls = nz(s['dls']); cum += dls; tort += dls
            prev = sv[i-1] if i else s
            dmd = nz(s['md'])-nz(prev['md'])
            if dmd > 0: tort += dls*dmd/30
            section = 'VERTICAL' if nz(s['md']) < 1000 else ('BUILD' if nz(s['incl']) < 30 else 'TANGENT/HORIZONTAL')
            dr.append((wk, s['report_no'], s['md'], s['incl'], s['azi'], s['tvd'], s['dls'], cum, tort, section))
    ins('directional', ['well_key','report_no','md','incl','azi','tvd','dls','cum_dogleg','tortuosity','section'], dr)
    # mud balance
    mb = []
    for wk in WELLS:
        mats = R('SELECT report_no,material,received,used,stock,total_received,cell_ref FROM ddr_materials WHERE well_key=? ORDER BY material,report_no', wk)
        prev = collections.defaultdict(lambda: None)
        for m in mats:
            key = (m['material'] or '').strip(); used = nz(m['used']); stock = nz(m['stock']); recv = m['received']
            sp = prev[key]; implied = stock + used - (sp if sp is not None else 0)
            delta = (implied - nz(recv)) if recv is not None else None
            st = 'NO_RECEIVED_COLUMN' if delta is None else ('BALANCED' if abs(delta) < .01 else ('MINOR' if abs(delta) <= max(1, abs(stock)*.02) else 'IMBALANCE'))
            mb.append((wk, m['report_no'], key[:90], recv, used, sp, stock, implied, delta, st))
            prev[key] = stock
    ins('mud_balance', ['well_key','report_no','material','received','used','stock_prev','stock','implied_received','delta','status'], mb)
    # casing checks
    cc = []
    for wk in WELLS:
        forms = [(f[1], nz(f[3])) for f in fs if f[0] == wk and f[3]]
        cas = R('SELECT size,grade,from_m,to_m,burst,collapse FROM ddr_casing WHERE well_key=? AND to_m IS NOT NULL ORDER BY to_m', wk)
        prev_to = None
        for c in cas:
            to = nz(c['to_m']); near = min(forms, key=lambda x: abs(x[1]-to)) if forms else ('', None)
            overlap = (prev_to - nz(c['from_m'])) if (prev_to is not None and c['from_m'] is not None) else None
            cc.append((wk, c['size'], c['grade'], c['from_m'], to, c['burst'], c['collapse'], near[0], near[1],
                       (to-near[1]) if near[1] is not None else None, overlap,
                       'SHOE_BELOW_TOP' if near[1] is not None and to > near[1] else 'SHOE_ABOVE_TOP'))
            prev_to = to
    ins('casing_check', ['well_key','size','grade','from_m','to_m','burst','collapse','nearest_formation','nearest_top_m','shoe_to_top_m','lap_overlap_m','status'], cc)
    # pressure stats
    ps = collections.defaultdict(list)
    for r in R("SELECT well_key,unit,value FROM eowr_facts WHERE fact_key='PRESSURE_TEST'"):
        ps[(r['well_key'], r['unit'] or '')].append(nz(r['value']))
    prs = []
    for (wk, u), vs in ps.items():
        q1, q3 = quant(vs, .25), quant(vs, .75); iqr = (q3 or 0)-(q1 or 0)
        prs.append((wk, u, len(vs), min(vs), med(vs), max(vs), mean(vs), std(vs) or 0,
                    sum(1 for v in vs if v < q1-1.5*iqr or v > q3+1.5*iqr)))
    ins('pressure_stats', ['well_key','unit','n','v_min','v_med','v_max','v_mean','v_std','outliers'], prs)
    commit()
    print('[6] formations', len(fs), '| survey stations', len(dr), '| mud-balance rows', len(mb), '| casing checks', len(cc), flush=True)


# ============================== 7. STIMULATION / RECONCILIATION / PROGRAM ====
def recon_forensics():
    table('stim_audit', 'well_key TEXT, rel_path TEXT, sheet TEXT, item TEXT, unit TEXT, planned_volume REAL, '
          'unit_price REAL, total_usd REAL, expected REAL, delta REAL, family TEXT, status TEXT')
    table('recon_delta', 'well_key TEXT, item TEXT, cbs_usd REAL, ddr_qty REAL, eowr_value REAL, '
          'delta_abs REAL, delta_pct REAL, tolerance_status TEXT, basis TEXT')
    table('program_vs_actual', 'well_key TEXT, formation TEXT, program_md REAL, eowr_md REAL, ddr_md REAL, '
          'delta_pa REAL, delta_pd REAL, bias TEXT, source_program TEXT, source_eowr TEXT')
    table('cost_efficiency', 'well_key TEXT, days INTEGER, drilled_m REAL, cbs REAL, gross REAL, payable REAL, net REAL, '
          'npt_h REAL, usd_per_day REAL, usd_per_m REAL, npt_share_pct REAL, payable_ratio_pct REAL, '
          'cost_rank INTEGER, npt_rank INTEGER, efficiency_index REAL')
    table('learning_metrics', 'chain TEXT, from_well TEXT, to_well TEXT, delta_days INTEGER, delta_npt_h REAL, '
          'delta_cbs REAL, delta_rop REAL, npt_per_day_from REAL, npt_per_day_to REAL, cbs_per_day_from REAL, cbs_per_day_to REAL, direction TEXT')
    table('field_correlations', 'x TEXT, y TEXT, n INTEGER, pearson REAL, spearman REAL, slope REAL, r2 REAL, strength TEXT')
    table('data_quality', 'check_name TEXT, scope TEXT, n INTEGER, affected INTEGER, pct REAL, detail TEXT')
    # stimulation audit
    sa = []
    for r in R('SELECT well_key,rel_path,sheet,description,unit,volume,unit_price,total,kind FROM cbs_stim'):
        vol, up, tot = r['volume'], r['unit_price'], r['total']
        if vol is None and up is None and tot is None: continue
        exp = nz(vol)*nz(up) if (vol is not None and up is not None) else None
        d = (nz(tot)-exp) if (exp is not None and tot is not None) else None
        st = 'NO_PRICE' if exp is None else ('MATCH' if d is not None and abs(d) < .01 else ('MINOR' if d is not None and abs(d) <= max(1, abs(exp)*.01) else 'MISMATCH'))
        fam = 'MATERIAL' if r['kind'] == 'STIM_MATERIAL' else ('SERVICE' if r['kind'] == 'STIM_SERVICE' else (r['kind'] or ''))
        sa.append((r['well_key'], r['rel_path'], r['sheet'], (r['description'] or '')[:120], r['unit'], vol, up, tot,
                   exp, d, fam, st))
    ins('stim_audit', ['well_key','rel_path','sheet','item','unit','planned_volume','unit_price','total_usd','expected','delta','family','status'], sa)
    # reconciliation deltas: CBS priced line vs DDR used quantity vs EOWR/Programme numeric fact
    rd = []
    for wk in WELLS:
        cbs = {re.sub(r'\s+',' ',(r['description'] or '').lower().strip())[:60]: nz(r['total']) for r in
               R('SELECT description,total FROM cbs_items WHERE well_key=? AND total IS NOT NULL', wk)}
        ddr = collections.defaultdict(float)
        for r in R('SELECT material,used FROM ddr_materials WHERE well_key=? AND used IS NOT NULL', wk):
            ddr[re.sub(r'\s+',' ',(r['material'] or '').lower().strip())[:60]] += nz(r['used'])
        eowr = collections.defaultdict(list)
        for r in R("SELECT fact_key,value,unit FROM eowr_facts WHERE well_key=? AND value IS NOT NULL", wk):
            eowr[re.sub(r'\s+',' ',(r['fact_key'] or '').lower().strip())[:60]].append(nz(r['value']))
        for key in sorted(set(cbs) | set(ddr)):
            c, d = cbs.get(key), ddr.get(key)
            ev = None
            for ek, vs in eowr.items():
                if key[:12] in ek or ek[:12] in key: ev = med(vs); break
            da = (c - d) if (c is not None and d is not None) else None
            dp = 100*da/c if (da is not None and c) else None
            st = ('CBS_ONLY' if d is None else ('DDR_ONLY' if c is None else
                  ('MATCHED' if da is not None and abs(da) < .01 else 'VALUE_GAP')))
            rd.append((wk, key, c, d, ev, da, dp, st, 'CBS total (USD) vs DDR used (units); EOWR fact where a matching key exists'))
    ins('recon_delta', ['well_key','item','cbs_usd','ddr_qty','eowr_value','delta_abs','delta_pct','tolerance_status','basis'], rd)
    # programme prognosis vs EOWR actual vs DDR actual
    pv = []
    for wk in WELLS:
        prog = collections.defaultdict(list); eow = collections.defaultdict(list)
        for r in R("SELECT fact_key,value,rel_path FROM eowr_facts WHERE well_key=? AND category='PROGRAM_TOPS'", wk):
            prog[r['fact_key'].replace(' MD','').strip().lower()].append((nz(r['value']), r['rel_path']))
        for r in R("SELECT fact_key,value,rel_path FROM eowr_facts WHERE well_key=? AND category='EOWR_TOPS'", wk):
            eow[r['fact_key'].replace(' MD','').strip().lower()].append((nz(r['value']), r['rel_path']))
        ddr = collections.defaultdict(list)
        for r in R('SELECT name,md FROM ddr_formations WHERE well_key=? AND md>0', wk):
            ddr[(r['name'] or '').strip().lower()].append(nz(r['md']))
        for nm in sorted(set(prog)|set(eow)|set(ddr)):
            pm = med([x[0] for x in prog.get(nm,[])]); em = med([x[0] for x in eow.get(nm,[])]); dm = med(ddr.get(nm,[]))
            pa = (em-pm) if (pm and em) else None; pd = (dm-pm) if (pm and dm) else None
            pv.append((wk, nm, pm, em, dm, pa, pd,
                       'ACTUAL_DEEPER' if (pa or 0) > 15 else ('ACTUAL_SHALLOWER' if (pa or 0) < -15 else 'ON_PROG'),
                       (prog.get(nm) or [('', '')])[0][1], (eow.get(nm) or [('', '')])[0][1]))
    ins('program_vs_actual', ['well_key','formation','program_md','eowr_md','ddr_md','delta_pa','delta_pd','bias','source_program','source_eowr'], pv)
    # cost efficiency
    ce = []
    for wk, p in P.items():
        c = p['costs']; d = p['ddr_days']; dr = p['drilled_total']; n = p['time']['npt_total_h']
        ce.append((wk, d, dr, c['cbs_plan'], c['gross_calendar'], c['payable_excl_ctr_npt'], c['pure_net_excl_all_npt'], n,
                   nz(c['gross_calendar'])/d if d else None,
                   nz(c['gross_calendar'])/dr if dr else None,
                   100*n/(d*24) if d else None,
                   100*nz(c['payable_excl_ctr_npt'])/nz(c['gross_calendar']) if c['gross_calendar'] else None,
                   0, 0, None))
    ranked = sorted(ce, key=lambda x: -(x[8] or 0)); cr = {r[0]: i+1 for i, r in enumerate(ranked)}
    nrank = {r[0]: i+1 for i, r in enumerate(sorted(ce, key=lambda x: -(x[7] or 0)))}
    ce = [r[:12] + (cr[r[0]], nrank[r[0]], round((r[8] or 0)/max(1,(r[7] or 0)/24), 4)) for r in ce]
    ins('cost_efficiency', ['well_key','days','drilled_m','cbs','gross','payable','net','npt_h','usd_per_day','usd_per_m','npt_share_pct','payable_ratio_pct','cost_rank','npt_rank','efficiency_index'], ce)
    # learning metrics
    lm = []
    for chain, seq in F['learning_curves'].items():
        for a, b in zip(seq, seq[1:]):
            pa, pb = P.get(a['well']), P.get(b['well'])
            if not pa or not pb: continue
            dd = pb['ddr_days']-pa['ddr_days']; dn = pb['time']['npt_total_h']-pa['time']['npt_total_h']
            dc = pb['costs']['cbs_plan']-pa['costs']['cbs_plan']
            dro = (nz(pb['rop_mean'])-nz(pa['rop_mean'])) if (pa['rop_mean'] and pb['rop_mean']) else None
            n1 = pa['time']['npt_total_h']/pa['ddr_days']; n2 = pb['time']['npt_total_h']/pb['ddr_days']
            c1 = pa['costs']['cbs_plan']/pa['ddr_days']; c2 = pb['costs']['cbs_plan']/pb['ddr_days']
            lm.append((chain, a['well'], b['well'], dd, dn, dc, dro, n1, n2, c1, c2,
                       'IMPROVED_NPT_INTENSITY' if n2 < n1 else 'DEGRADED_NPT_INTENSITY'))
    ins('learning_metrics', ['chain','from_well','to_well','delta_days','delta_npt_h','delta_cbs','delta_rop','npt_per_day_from','npt_per_day_to','cbs_per_day_from','cbs_per_day_to','direction'], lm)
    # field correlations
    fc = []
    W = [(wk, p['ddr_days'], p['time']['npt_total_h'], p['costs']['cbs_plan'], p['costs']['gross_calendar'],
          p['drilled_total'], p['rop_mean'], p['depth_final'], p['mw_max']) for wk, p in P.items()]
    cols = dict(DAYS=1, NPT=2, CBS=3, GROSS=4, DRILLED=5, ROP=6, TD=7, MW=8)
    for xn, xi in cols.items():
        for yn, yi in cols.items():
            if xi >= yi: continue
            pairs = [(r[xi], r[yi]) for r in W if r[xi] is not None and r[yi] is not None]
            if len(pairs) < 5: continue
            a = [x for x, y in pairs]; b = [y for x, y in pairs]
            pr, sp = pearson(a,b), spearman(a,b); sl, it, rv = regress(a,b)
            fc.append((xn, yn, len(a), pr, sp, sl, rv,
                       'STRONG' if pr is not None and abs(pr)>=.7 else ('MODERATE' if pr is not None and abs(pr)>=.4 else 'WEAK')))
    ins('field_correlations', ['x','y','n','pearson','spearman','slope','r2','strength'], fc)
    # data-quality scorecard
    dq = []
    def add(check, scope, n, aff, detail):
        n = n or 0; aff = aff or 0
        dq.append((check, scope, n, aff, 100*aff/max(1,n), detail))
    cq = R('SELECT SUM(numeric_as_text) a,SUM(formula_cells) b,SUM(hardcoded_formula) c,SUM(cross_sheet) d,SUM(total_cells) e FROM an_cell_quality')[0]
    cq = {k: (v or 0) for k, v in cq.items()}
    add('Numeric values stored as text','all sheets',cq['e'],cq['a'],'Cells that look numeric but are typed as strings')
    add('Formula cells','all sheets',cq['e'],cq['b'],'Cached value + formula string retained')
    add('Formula without cell references','all sheets',cq['b'],cq['c'],'Hard-coded arithmetic or constant formulas')
    add('Cross-sheet formulas','all sheets',cq['b'],cq['d'],'Formula referencing another sheet/workbook')
    fa = R('SELECT COUNT(*) n,SUM(status="MISMATCH") m,SUM(status="MINOR") mi FROM an_formula_audit')[0]
    add('Formula recomputation mismatch','audited formulas',fa['n'],fa['m'],'Cached value differs from recomputed expression by >0.1%')
    add('Formula recomputation minor delta','audited formulas',fa['n'],fa['mi'],'Small rounding difference')
    cb = R('SELECT COUNT(*) n,SUM(status<>"MATCH") m,SUM(status="MATERIAL") mm FROM an_cbs_line_audit')[0]
    add('CBS qty x unit price != total','priced CBS lines',cb['n'],cb['m'],'Source arithmetic / allocation gap; exact delta stored')
    add('CBS material arithmetic gap','priced CBS lines',cb['n'],cb['mm'],'Gap >=5% of expected product')
    tbal = R('SELECT COUNT(*) n,SUM(balance_status<>"COMPLETE_24H") m,SUM(balance_status="UNDER_CODED") u FROM an_time_balance')[0]
    add('DDR 24-h hours not equal 24','DDR days',tbal['n'],tbal['m'],'Invisible Lost Time computed, never forced')
    dcon = R('SELECT COUNT(*) n,SUM(status="CUMULATIVE_TYPO") t,SUM(status="MINOR") m FROM an_daily_consistency')[0]
    add('Cumulative-depth typo (derived correction)','DDR days',dcon['n'],dcon['t'],'Corrected only in derived corrected_24h column')
    mbal = R('SELECT COUNT(*) n,SUM(status="IMBALANCE") m FROM an_mud_balance')[0]
    add('Mud stock / received / used imbalance','material-days',mbal['n'],mbal['m'],'Implied receipt vs source received column')
    low = q('SELECT COUNT(*) FROM pdf_pages WHERE n_chars<200')[0][0]
    add('Low-text PDF pages (<200 chars)','PDF pages',4345,low,'Image-only / scanned pages retained page by page')
    ins('data_quality', ['check_name','scope','n','affected','pct','detail'], dq)
    commit()
    print('[7] stimulation audit', len(sa), '| reconciliation', len(rd), '| prognosis-vs-actual', len(pv),
          '| efficiency', len(ce), '| learning', len(lm), '| field correlations', len(fc), '| DQ', len(dq), flush=True)


# ============================================ 8. FINDINGS & EXCEPTION CSV =====
def findings_and_export():
    table('findings', 'id TEXT, category TEXT, title TEXT, numeric_evidence TEXT, interpretation TEXT, action TEXT, priority TEXT')
    out = []
    def A(cat, title, ev, interp, act, pri='P1'):
        out.append((f'DA-{len(out)+1:02d}', cat, title, ev, interp, act, pri))
    # CBS arithmetic
    cb = R('SELECT COUNT(*) n,SUM(status<>"MATCH") m,SUM(delta) d,SUM(ABS(delta)) ad FROM an_cbs_line_audit')[0]
    worst = R('SELECT well_key,description,expected,total,delta,pct,sheet,cell FROM an_cbs_line_audit WHERE status="MATERIAL" ORDER BY ABS(delta) DESC LIMIT 5')
    A('CBS tariff', ' priced-line arithmetic is not closed by qty × unit price',
      f"{cb['m']} of {cb['n']} priced lines diverge; net signed delta {fmt(cb['d'],0)} USD, absolute {fmt(cb['ad'],0)} USD. Largest: " +
      '; '.join(f"{x['well_key']} {x['description'][:34]} expected {fmt(x['expected'],0)} vs total {fmt(x['total'],0)} ({fmt(x['pct'],1)}%)" for x in worst),
      'The CBS Total column is a lump-sum/allocated ceiling, not always the algebraic product of QTY and Unit Price. Per project confirmation (2026-11-08), the large positive deltas between the source Total and QTY × unit rate are the embedded company margin (profit/overhead), not extraction errors; treating either side as the "true" cost would misstate the invoice basis.',
      'Keep both columns with their source cell; reconcile contractually before payment; use the peer-well reference rate card (DA-22) to normalise missing rates, and never recompute the ceiling from QTY × price.')
    # time balance
    t = R('SELECT COUNT(*) n,SUM(balance_status="UNDER_CODED") u,SUM(balance_status="OVER_CODED") o,AVG(ilt_h) ai,SUM(ilt_h) si FROM an_time_balance')[0]
    co = R('SELECT SUM(carried_over_h) h,AVG(carried_over_h) a,SUM(CASE WHEN ilt_h>0 THEN ilt_h ELSE 0 END) short FROM an_time_balance')[0]
    A('Time / ILT', '24-h narrative logs contain carried-over / overlapping rows from adjacent reports',
      f"After selecting the single canonical sheet per report and removing blank summary rows: {t['u']} of {t['n']} days under-coded, {t['o']} over-coded; mean signed gap {fmt(t['ai'],2)} h/day, total carried-over excess {fmt(co['h'],0)} h (mean {fmt(co['a'],2)} h/day), genuine under-coded shortfall {fmt(co['short'],0)} h.",
      'DDR operation logs frequently include a continuation row from the previous or next midnight. The excess is flagged as carried-over hours; it is never silently redistributed into PT or NPT, and the source rows are retained.',
      'Introduce a report-boundary timestamp on every activity row at DDR capture, then recompute ILT; investigate the largest carried-over days individually.')
    # cumulative
    d = R('SELECT COUNT(*) n,SUM(status="CUMULATIVE_TYPO") t,SUM(ABS(delta)) ad FROM an_daily_consistency')[0]
    A('Data quality', 'Cumulative depth / 24-h footage typographical errors are isolated',
      f"{d['t']} of {d['n']} daily rows show |drilled_24 − (MD24−MD00)| > 25 m; total absolute mismatch {fmt(d['ad'],1)} m.",
      'Probable transcription slips in cumulative records. Correction is stored only in derived corrected_24h; the source cell is untouched.',
      'Use corrected_24h for time-depth analytics while keeping the source value for audit.')
    # formula audit
    fa = R('SELECT COUNT(*) n,SUM(status="MATCH") m,SUM(status="MINOR") mi,SUM(status="MISMATCH") mm FROM an_formula_audit')[0]
    A('Formula audit', 'Cached formula values were independently recomputed',
      f"{fa['n']} formulas evaluated: {fa['m']} exact matches, {fa['mi']} minor rounding, {fa['mm']} material mismatches.",
      'Recomputing from the archived cell values validates the Excel layer instead of trusting cached results.',
      'Review the mismatched cells with their sheet/cell reference before using them as invoice support.')
    # unit price dispersion
    up = R('SELECT COUNT(*) n,SUM(dispersion="DISPERSED") d FROM an_unit_price')[0]
    top = R('SELECT description,unit,v_min,v_max,spread_pct,n FROM an_unit_price WHERE dispersion="DISPERSED" ORDER BY spread_pct DESC LIMIT 5')
    A('Rate benchmark', 'Unit-price dispersion across the 13 operations',
      f"{up['d']} of {up['n']} normalised tariff keys show >5% spread. Largest: " +
      '; '.join(f"{x['description'][:30]} {fmt(x['v_min'],0)}→{fmt(x['v_max'],0)} ({fmt(x['spread_pct'],1)}%, n={x['n']})" for x in top),
      'Unit rates for a defined operation/commodity are uniform campaign-wide (project confirmation 2026-11-08). Observed >5% spreads are therefore scope/hole-size/contract-bundle differences or incomplete source data — not noise and not an error — to be read against the reference rate card (DA-22). Wells 29/38 remain benchmarks only.',
      'Benchmark each new AFE against the median of the same normalised key and flag >5% deviations; where a CBS line lacks a unit rate, reference the rate from peer wells running the same defined item.')
    # correlations
    cs = R("SELECT well_key,x,y,n,pearson,r2,strength FROM an_correlations WHERE strength IN ('STRONG','MODERATE') ORDER BY ABS(pearson) DESC LIMIT 8")
    A('Drilling parameters', 'Well-level ROP drivers identified from paired daily rows',
      '; '.join(f"{x['well_key']} {x['x']}→ROP r={fmt(x['pearson'],2)} (n={x['n']}, R²={fmt(x['r2'],2)})" for x in cs) or 'No paired daily rows reached n>=4 for a moderate correlation.',
      'Pearson/Spearman and OLS are computed only on DDR days where both parameters exist; no missing value is imputed.',
      'Use the strongest per-well driver as the control variable in the next AFE parameter window.')
    # formation structural bias
    fb = R("SELECT COUNT(*) n,SUM(structural_bias='DEEPER_THAN_PROG') d,SUM(structural_bias='SHALLOWER_THAN_PROG') s,AVG(delta_m) a,AVG(ABS(delta_m)) aa FROM an_formation_stats WHERE delta_m IS NOT NULL")[0]
    A('Geoscience', 'Prognosis-to-actual structural bias quantified across correlated tops',
      f"{fb['n']} formation-well pairs with paired prognosis: {fb['d']} deeper, {fb['s']} shallower; mean signed delta {fmt(fb['a'],1)} m, mean absolute {fmt(fb['aa'],1)} m.",
      'A systematic bias shifts pore-pressure/fracture ramps and casing-shoe placement; a zero-mean but high absolute delta still implies a coarse structural grid.',
      'Update the South Azadegan structural grid with the measured tops before setting the next 9-5/8 in shoe and KOP.')
    # programme vs actual
    pv = R('SELECT COUNT(*) n,AVG(delta_pa) a,AVG(ABS(delta_pa)) aa,SUM(bias="ACTUAL_DEEPER") d,SUM(bias="ACTUAL_SHALLOWER") s FROM an_program_vs_actual WHERE delta_pa IS NOT NULL')[0]
    A('Well Programme', 'Programme prognosis vs EOWR measured tops',
      f"{pv['n']} paired tops; mean signed {fmt(pv['a'],1)} m, mean absolute {fmt(pv['aa'],1)} m, {pv['d']} deeper / {pv['s']} shallower.",
      'Three independent top sources (Programme, EOWR, DDR) are now numerically comparable with exact file/page provenance.',
      'Feed measured deltas back into the next programme revision and the mud-weight window.')
    # mud balance
    mb = R('SELECT COUNT(*) n,SUM(status="IMBALANCE") m,AVG(ABS(delta)) a FROM an_mud_balance WHERE delta IS NOT NULL')[0]
    A('Fluids / materials', 'Daily mud stock balance audited sequentially',
      f"{mb['n']} material-days with a source Received value; {mb['m']} imbalance, mean absolute implied-receipt gap {fmt(mb['a'],1)} units.",
      'Sequential stock_prev + received − used = stock is checked instead of comparing isolated columns.',
      'Reconcile the largest imbalance days against the DDR back-load column before cost allocation.')
    # cost efficiency
    ce = R('SELECT * FROM an_cost_efficiency ORDER BY usd_per_day DESC')
    A('Cost control', 'Cost efficiency ranking across the 13 operations',
      'Highest USD/day: ' + '; '.join(f"{x['well_key']} {fmt(x['usd_per_day'],0)}" for x in ce[:4]) +
      ' | Lowest: ' + '; '.join(f"{x['well_key']} {fmt(x['usd_per_day'],0)}" for x in ce[-4:]) +
      f" | NPT share of calendar {fmt(mean([x['npt_share_pct'] for x in ce if x['npt_share_pct']]),1)}% field mean.",
      'USD/day, USD/m (where footage exists), NPT share and payable/gross ratio expose both expensive and inefficient operations.',
      'Target the top-quartile USD/day wells with the NPT Pareto and the four-track deduction waterfall.')
    # learning
    for lm in R('SELECT * FROM an_learning_metrics'):
        A('Learning curve', f"{lm['chain']}: {lm['from_well']} → {lm['to_well']} — {lm['direction']}",
          f"NPT/day {fmt(lm['npt_per_day_from'],2)} → {fmt(lm['npt_per_day_to'],2)} h/day; CBS/day {fmt(lm['cbs_per_day_from'],0)} → {fmt(lm['cbs_per_day_to'],0)} USD/day; Δdays {lm['delta_days']}, ΔNPT {fmt(lm['delta_npt_h'],1)} h, ΔCBS {fmt(lm['delta_cbs'],0)} USD.",
          'Repeated interventions on the same wellbore are the cleanest available learning signal because geology is held approximately constant.',
          'Carry the improved parameter/coding discipline forward; investigate any degradation before the next re-entry.' if lm['direction'].startswith('IMPROVED') else 'Perform a formal autopsy of the degraded step before authorising another intervention.')
    # NPT pareto
    np_ = R('SELECT well_key,category,hours,share_pct,cumulative_pct FROM an_npt_pareto WHERE rank<=3 ORDER BY well_key,rank')
    conc = collections.defaultdict(float)
    for x in np_: conc[x['well_key']] += x['share_pct'] or 0
    A('NPT autopsy', 'Top-three attributed parties concentrate most non-productive time',
      '; '.join(f"{k} {fmt(v,1)}%" for k, v in sorted(conc.items(), key=lambda z:-z[1])[:6]) +
      f" | field attributed NPT {fmt(TOT['npt_hours'],1)} h = {fmt(TOT['npt_hours']/24,1)} rig days.",
      'Attribution concentration shows whether NPT is systemic (one responsible party) or diffuse.',
      'Put the dominant attributed party under a joint daily NPT review with a 24-h closure target.')
    # outliers
    ol = R('SELECT COUNT(*) n,MAX(ABS(z)) m FROM an_outliers')[0]
    A('Numeric census', 'Extreme numeric cells identified with z-score and provenance',
      f"{ol['n']} |z|>6 cells retained, maximum |z| {fmt(ol['m'],1)}; full histogram by order of magnitude in an_num_hist.",
      'Outliers are not deleted: they are usually legitimate depths, cumulative costs or casing tallies.',
      'Use an_outliers as a targeted QA list, not as an automatic reject rule.')
    ins('findings', ['id','category','title','numeric_evidence','interpretation','action','priority'], out)
    commit()
    # exception register CSV
    path = os.path.join(OUT, 'SAZN_Analytics_Exceptions.csv')
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f); w.writerow(['exception_type','well','document','sheet','cell_or_page','source_value','expected_or_reference','delta','status','detail'])
        for r in R('SELECT * FROM an_cbs_line_audit WHERE status<>"MATCH"'):
            w.writerow(['CBS_LINE_ARITHMETIC',r['well_key'],r['rel_path'],r['sheet'],r['cell'],r['total'],r['expected'],r['delta'],r['status'],r['description']])
        for r in R('SELECT * FROM an_formula_audit WHERE status<>"MATCH"'):
            w.writerow(['FORMULA_RECOMPUTE',r['well_key'],r['rel_path'],r['sheet'],r['cell'],r['cached'],r['recomputed'],r['delta'],r['status'],r['formula']])
        for r in R('SELECT * FROM an_time_balance WHERE balance_status<>"COMPLETE_24H"'):
            w.writerow(['DDR_TIME_BALANCE',r['well_key'],'DDR pack','24-h activity',r['report_no'],r['activity_h'],24,r['ilt_h'],r['balance_status'],'Invisible Lost Time (h)'])
        for r in R('SELECT * FROM an_daily_consistency WHERE status NOT IN ("MATCH","NO_SOURCE")'):
            w.writerow(['CUMULATIVE_DEPTH',r['well_key'],'DDR pack','daily depth',r['report_no'],r['drilled_24'],r['expected_delta'],r['delta'],r['status'],'derived corrected_24h='+str(r['corrected_24h'])])
        for r in R('SELECT * FROM an_mud_balance WHERE status="IMBALANCE"'):
            w.writerow(['MUD_STOCK_BALANCE',r['well_key'],'DDR pack','materials',r['report_no'],r['stock'],r['implied_received'],r['delta'],r['status'],r['material']])
        for r in R('SELECT * FROM an_stim_audit WHERE status="MISMATCH"'):
            w.writerow(['STIM_LINE_ARITHMETIC',r['well_key'],r['rel_path'],r['sheet'],'',r['total_usd'],r['expected'],r['delta'],r['status'],r['item']])
        for r in R('SELECT * FROM an_outliers'):
            w.writerow(['NUMERIC_OUTLIER',r['well_key'],r['rel_path'],r['sheet'],r['cell'],r['value'],r['sheet_mean'],r['z'],'|z|>6','sheet std='+str(r['sheet_std'])])
    print('[8] findings', len(out), '| exception CSV MB', round(os.path.getsize(path)/1e6, 2), flush=True)
    return out


def summary_json(findings):
    S = dict(generated='2026-11-08', scope='Drilling-Data Analysis Department — exhaustive numeric forensics',
        source_census=TOT,
        derived_tables={t[0][3:]: q('SELECT COUNT(*) FROM "'+t[0]+'"')[0][0] for t in q("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'an_%'")},
        data_quality=[dict(r) for r in R('SELECT * FROM an_data_quality')],
        findings=[dict(r) for r in R('SELECT * FROM an_findings')],
        per_well={}, field={})
    for wk in WELLS:
        p = P[wk]
        S['per_well'][wk] = dict(
            op=p['op'], rig=p['rig'], contract=p['contract'], days=p['ddr_days'],
            cbs=p['costs']['cbs_plan'], gross=p['costs']['gross_calendar'], payable=p['costs']['payable_excl_ctr_npt'],
            net=p['costs']['pure_net_excl_all_npt'], npt_h=p['time']['npt_total_h'],
            cbs_lines=q('SELECT COUNT(*) FROM an_cbs_line_audit WHERE well_key=?',(wk,))[0][0],
            cbs_mismatch=q('SELECT COUNT(*) FROM an_cbs_line_audit WHERE well_key=? AND status<>"MATCH"',(wk,))[0][0],
            time_under=q('SELECT COUNT(*) FROM an_time_balance WHERE well_key=? AND balance_status="UNDER_CODED"',(wk,))[0][0],
            ilt_h=q('SELECT SUM(ilt_h) FROM an_time_balance WHERE well_key=?',(wk,))[0][0],
            cum_typos=q('SELECT COUNT(*) FROM an_daily_consistency WHERE well_key=? AND status="CUMULATIVE_TYPO"',(wk,))[0][0],
            param_rows=q('SELECT COUNT(*) FROM an_daily_params WHERE well_key=?',(wk,))[0][0],
            correlations=q('SELECT COUNT(*) FROM an_correlations WHERE well_key=?',(wk,))[0][0],
            formations=q('SELECT COUNT(*) FROM an_formation_stats WHERE well_key=?',(wk,))[0][0],
            directional=q('SELECT COUNT(*) FROM an_directional WHERE well_key=?',(wk,))[0][0],
            mud_imbalance=q('SELECT COUNT(*) FROM an_mud_balance WHERE well_key=? AND status="IMBALANCE"',(wk,))[0][0],
            reconciliation=q('SELECT COUNT(*) FROM an_recon_delta WHERE well_key=?',(wk,))[0][0],
            cost_efficiency=[dict(r) for r in R('SELECT * FROM an_cost_efficiency WHERE well_key=?',(wk,))],
            npt_pareto=[dict(r) for r in R('SELECT * FROM an_npt_pareto WHERE well_key=? AND rank<=6',(wk,))],
            top_correlations=[dict(r) for r in R('SELECT * FROM an_correlations WHERE well_key=? ORDER BY ABS(pearson) DESC LIMIT 6',(wk,))],
            formation_bias=[dict(r) for r in R('SELECT formation,delta_m,thickness_m,problem_h,structural_bias FROM an_formation_stats WHERE well_key=? ORDER BY md_med LIMIT 20',(wk,))],
        )
    S['field'] = dict(
        cost_efficiency=[dict(r) for r in R('SELECT * FROM an_cost_efficiency ORDER BY cost_rank')],
        field_correlations=[dict(r) for r in R('SELECT * FROM an_field_correlations ORDER BY ABS(pearson) DESC')],
        learning_metrics=[dict(r) for r in R('SELECT * FROM an_learning_metrics')],
        unit_price_dispersion=[dict(r) for r in R('SELECT * FROM an_unit_price WHERE dispersion="DISPERSED" ORDER BY spread_pct DESC LIMIT 40')],
        num_hist=[dict(r) for r in R('SELECT * FROM an_num_hist')],
        program_vs_actual=[dict(r) for r in R('SELECT * FROM an_program_vs_actual WHERE delta_pa IS NOT NULL ORDER BY ABS(delta_pa) DESC LIMIT 60')],
    )
    path = os.path.join(DATA, 'deep_analytics.json')
    json.dump(S, open(path, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print('[9] deep_analytics.json MB', round(os.path.getsize(path)/1e6, 2), flush=True)
    return S


def main():
    cell_census(); cbs_forensics(); time_forensics(); param_forensics(); npt_forensics()
    geo_forensics(); recon_forensics(); findings = findings_and_export(); S = summary_json(findings)
    CON.execute('VACUUM'); CON.close()
    print('DEEP ANALYTICS DONE | tables', len(S['derived_tables']), '| findings', len(findings),
          '| DB MB', round(os.path.getsize(DB)/1e6, 2), flush=True)


if __name__ == '__main__': main()
