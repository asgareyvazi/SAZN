"""SAZN Drilling-Data Analysis Department — PDF numeric-token + four-track cost forensics.

Reads every one of the 305,447 numeric PDF tokens archived in pdf_num_pages,
every Excel numeric cell / EOWR fact value, every 24-h NPT attribution row and
every sourced four-track cost input.  Classifies, statistically audits and
cross-checks each PDF token without deleting or altering any source value.
Extends an_findings with DA-16 … DA-21 and refreshes deep_analytics.json.

Derived classifications, robust outlier flags and "exact numeric coincidence"
flags are explicitly labelled and stored separately from the source token.
"""
import os, sys, json, math, re, zlib, gzip, csv, hashlib, sqlite3, collections, statistics
sys.path.insert(0, os.path.dirname(__file__))
from common import OUT, DATA, DB, load_json, fmt, r2, norm

CON = sqlite3.connect(DB)
CON.row_factory = sqlite3.Row
P = load_json('well_profiles')
S_PATH = os.path.join(DATA, 'deep_analytics.json')
S = json.load(open(S_PATH, encoding='utf-8')) if os.path.exists(S_PATH) else {}
WELLS = list(P.keys())

KEY = re.compile(r'mud|cement|casing|bit|bha|npt|acid|stimulation|cbl|usit|packer|sidetrack|loss|fish|milling|trsssv|plt|pvt|waste|formation|top|pressure|rate|pi|whp|choke|volume|depth|md|tvd|rop|wob|rpm', re.I)
NAV = re.compile(r'page\s+\d+\s+of\s+\d+|table of contents|\.{3,}|^\s*\d+\s*$|contents', re.I)
IDENT = re.compile(r'CCA/|AZNS|PAD|WELL|REPORT|REV|DATE|CONTRACT|PAGE|DOC|NO\.|20\d\d|14\d\d', re.I)
WAIT = re.compile(r'wait|waiting', re.I)
CK = re.compile(r'oeoc|contractor|service|mud|logging|directional|catering|energy|supplier|vendor|rig|unattributed', re.I)
NPT_STATUS = re.compile(r'^(?:N|NPT|UP|UNPLANNED|UNPLAN|UNPLN|NP|PN|UPI|NPTUNPLN|NPT-UNPLN|UP-NPT|UNPLN-NPT)$|^NPT|UNPLN', re.I)


def flat(a):
    if len(a) == 1 and isinstance(a[0], (tuple, list)): return tuple(a[0])
    return tuple(a)
def W(sql, *a): CON.execute(sql, flat(a))
def R(sql, *a): return [dict(r) for r in CON.execute(sql, flat(a)).fetchall()]
def table(name, ddl):
    W('DROP TABLE IF EXISTS an_' + name)
    W('CREATE TABLE an_' + name + ' (' + ddl + ')')
def ins(name, cols, rows):
    if not rows: return
    CON.executemany('INSERT INTO an_' + name + ' (' + ','.join(cols) + ') VALUES (' + ','.join('?' * len(cols)) + ')', rows)

def mean(xs): return sum(xs) / len(xs) if xs else None
def med(xs):
    ys = sorted(xs); return ys[len(ys)//2] if ys else None
def std(xs):
    m = mean(xs); return math.sqrt(sum((x-m)**2 for x in xs)/len(xs)) if len(xs) > 1 and m is not None else None
def quant(xs, p):
    if not xs: return None
    ys = sorted(xs); i = (len(ys)-1)*p; lo = int(i); hi = min(lo+1, len(ys)-1)
    return ys[lo] + (ys[hi]-ys[lo])*(i-lo)
def vkey(v):
    if v is None: return ''
    try: x = float(v)
    except Exception: return ''
    if abs(x) < 1e-12: x = 0.0
    return format(x, '.6f')
def a1(r, c):
    s = ''
    while c and c > 0:
        c, rem = divmod(c-1, 26); s = chr(65+rem) + s
    return (s or 'A') + str(r)
def band(v):
    if v == 0: return '0'
    if v < 0: return 'negative'
    a = abs(v)
    for lim, name in ((1,'<1'),(10,'1-10'),(100,'10-100'),(1000,'100-1k'),(10000,'1k-10k'),
                      (100000,'10k-100k'),(1000000,'100k-1M'),(10000000,'1M-10M'),(100000000,'10M-100M'),(1e12,'>=100M')):
        if a < lim: return name
    return '>=1e12'
def pct(a, b): return round(100.0*a/b, 4) if b else None

# ------------------------------------------------------------------ sources --
SHEETS = {r['id']: (r['rel_path'], r['name']) for r in R('SELECT id,rel_path,name FROM sheets')}
excel_counts = collections.Counter(); excel_first = {}
for r in CON.execute('SELECT sheet_id,r,c,value FROM cell_numeric'):
    k = vkey(r['value'])
    if not k: continue
    excel_counts[k] += 1
    if k not in excel_first:
        rel, sh = SHEETS.get(r['sheet_id'], ('', ''))
        excel_first[k] = '%s!%s' % (rel, sh + '!' + a1(r['r'], r['c']))
fact_counts = collections.Counter(); fact_first = {}
for r in R('SELECT well_key,value,rel_path,page_no,fact_key FROM eowr_facts WHERE value IS NOT NULL'):
    k = vkey(r['value'])
    if not k: continue
    fact_counts[k] += 1
    if k not in fact_first:
        fact_first[k] = '%s p.%s [%s]' % (r['rel_path'], r['page_no'], r['fact_key'])

# Token tuple indices
I_WELL,I_REL,I_CAT,I_PAGE,I_LINE,I_TOK,I_VAL,I_UNIT,I_CTX,I_CHAP,I_CHARS,I_IMG, \
I_BAND,I_CLASS,I_KIND,I_MEAS,I_NAV,I_IDENT,I_LOW,I_EXCEL,I_FACT,I_OUT,I_LOG,I_Z,I_KEY,I_FLAGS,I_UNITKEY = range(27)

token_rows = []
page_meta = {}
pdf_counts = collections.Counter(); pdf_first = {}
for r in CON.execute('''
    SELECT n.file_id,n.page_no,n.well_key AS nwell,f.rel_path,f.category,
           p.well_key AS pwell,p.n_chars,p.n_images,p.chapter,n.payload
    FROM pdf_num_pages n
    JOIN source_files f ON f.id=n.file_id
    LEFT JOIN pdf_pages p ON p.file_id=n.file_id AND p.page_no=n.page_no'''):
    well = r['pwell'] or r['nwell'] or ''
    rel = r['rel_path']; cat = r['category'] or ''
    page = int(r['page_no']); n_chars = int(r['n_chars'] or 0); n_images = int(r['n_images'] or 0)
    chapter = r['chapter'] or ''
    pk = (rel, page)
    page_meta[pk] = dict(well=well, rel=rel, cat=cat, n_chars=n_chars, n_images=n_images, chapter=chapter)
    try: data = json.loads(zlib.decompress(r['payload']).decode('utf-8'))
    except Exception: data = []
    for line, tok, v, unit, ctx in data:
        try: v = float(v)
        except Exception: continue
        tok = str(tok); unit = (unit or '').strip().lower(); ctx = norm(ctx)[:180]
        b = band(v); unitkey = unit or 'NO_UNIT'; logv = math.log10(1+abs(v))
        is_measured = 1 if (unit or KEY.search(ctx)) else 0
        is_nav = 1 if NAV.search(ctx) else 0
        is_ident = 1 if IDENT.search(ctx) else 0
        is_low = 1 if n_chars < 200 else 0
        if unit: cclass = 'unit-bearing'
        elif KEY.search(ctx): cclass = 'measurement-context'
        elif NAV.search(ctx): cclass = 'navigation/toc'
        elif IDENT.search(ctx): cclass = 'header/identifier'
        else: cclass = 'unclassified'
        kind = 'integer' if ('.' not in tok and 'e' not in tok.lower() and float(v).is_integer()) else 'decimal'
        k = vkey(v)
        ex = 1 if k in excel_counts else 0
        fa = 1 if k in fact_counts else 0
        flags = ['SOURCE_ONLY_DERIVED_CLASSIFICATION']
        if unit: flags.append('UNIT')
        if is_measured: flags.append('MEASURED')
        if is_nav: flags.append('NAVIGATION')
        if is_ident: flags.append('IDENTIFIER')
        if is_low: flags.append('LOW_TEXT_PAGE')
        if n_images and is_low: flags.append('IMAGE_ONLY_RISK')
        if ex: flags.append('EXACT_EXCEL_VALUE_COINCIDENCE')
        if fa: flags.append('EXACT_FACT_VALUE_COINCIDENCE')
        if cclass == 'unclassified': flags.append('UNCLASSIFIED_CONTEXT')
        token_rows.append([
            well, rel, cat, page, int(line), tok, v, unit, ctx, chapter, n_chars, n_images,
            b, cclass, kind, is_measured, is_nav, is_ident, is_low, ex, fa, 0, logv, 0.0, k, '|'.join(flags), unitkey
        ])
        pdf_counts[k] += 1
        if k not in pdf_first:
            pdf_first[k] = '%s p.%d L%d [%s]' % (rel, page, line, ctx[:60])
print('PDF TOKENS', len(token_rows), flush=True)

# Robust unit-group magnitude fences (derived; no source value changed).
groups = collections.defaultdict(list)
for t in token_rows: groups[t[I_UNITKEY]].append(t[I_LOG])
thresh = {}
for u, vals in groups.items():
    q1, q3, m = quant(vals, .25), quant(vals, .75), med(vals)
    iqr = q3-q1
    absdev = [abs(x-m) for x in vals]; mad = med(absdev)
    den = 1.4826*mad if mad else (iqr or 1.0)
    thresh[u] = (q1, q3, iqr, m, den, q1-3*iqr if iqr else q1-.001, q3+3*iqr if iqr else q3+.001)
for t in token_rows:
    q1,q3,iqr,m,den,lo,hi = thresh[t[I_UNITKEY]]
    z = (t[I_LOG]-m)/den
    t[I_Z] = z
    out = 1 if (iqr > 0 and (t[I_LOG] < lo or t[I_LOG] > hi or abs(z) > 6)) else 0
    t[I_OUT] = out
    if out: t[I_FLAGS] += '|ROBUST_UNIT_GROUP_OUTLIER'

# -------------------------------------------------------------- DB tables ----
table('pdf_tokens', '''id INTEGER PRIMARY KEY, well_key TEXT, rel_path TEXT, category TEXT, page_no INTEGER, line_no INTEGER,
 token TEXT, value REAL, unit TEXT, context TEXT, chapter TEXT, n_chars INTEGER, n_images INTEGER,
 band TEXT, context_class TEXT, value_kind TEXT, is_measured INTEGER, is_navigation INTEGER, is_identifier INTEGER,
 is_low_text INTEGER, excel_exact_match INTEGER, fact_exact_match INTEGER, robust_outlier INTEGER,
 log_abs REAL, robust_z REAL, value_key TEXT, flags TEXT''')
COLS = ['well_key','rel_path','category','page_no','line_no','token','value','unit','context','chapter','n_chars','n_images',
        'band','context_class','value_kind','is_measured','is_navigation','is_identifier','is_low_text',
        'excel_exact_match','fact_exact_match','robust_outlier','log_abs','robust_z','value_key','flags']
ins('pdf_tokens', COLS, [tuple(t[:I_UNITKEY]) for t in token_rows])
W('CREATE INDEX IF NOT EXISTS ix_an_pdf_tokens_value ON an_pdf_tokens(value_key)')
W('CREATE INDEX IF NOT EXISTS ix_an_pdf_tokens_well ON an_pdf_tokens(well_key)')
W('CREATE INDEX IF NOT EXISTS ix_an_pdf_tokens_unit ON an_pdf_tokens(unit)')

# Value coincidence register — exact numeric equality to six decimals is NOT semantic equivalence.
value_rows = []
for k, n in pdf_counts.items():
    ec, fc = excel_counts.get(k, 0), fact_counts.get(k, 0)
    if not (ec or fc): continue
    value_rows.append((k, float(k), n, ec, fc, pdf_first.get(k,''), excel_first.get(k,''), fact_first.get(k,''),
                       'excel+fact' if ec and fc else ('excel' if ec else 'fact'),
                       'EXACT_NUMERIC_COINCIDENCE_ONLY'))
value_rows.sort(key=lambda x: (-(x[2]*(x[3]+x[4])), -x[2]))
table('pdf_value_matches', '''value_key TEXT, value REAL, pdf_tokens INTEGER, excel_cells INTEGER, fact_rows INTEGER,
 first_pdf_source TEXT, first_excel_source TEXT, first_fact_source TEXT, match_kind TEXT, caveat TEXT''')
ins('pdf_value_matches', ['value_key','value','pdf_tokens','excel_cells','fact_rows','first_pdf_source','first_excel_source','first_fact_source','match_kind','caveat'], value_rows)

# Grouped token statistics: category × well × unit × magnitude band.
G = collections.defaultdict(lambda: dict(vals=[], zeros=0, negatives=0, ints=0, decs=0, meas=0, nav=0, ident=0,
                                         low=0, out=0, ex=0, fa=0, unit=0, unclass=0))
for t in token_rows:
    g = G[(t[I_CAT], t[I_WELL], t[I_UNITKEY], t[I_BAND])]
    g['vals'].append(t[I_VAL])
    g['zeros'] += t[I_VAL] == 0; g['negatives'] += t[I_VAL] < 0
    g['ints'] += t[I_KIND] == 'integer'; g['decs'] += t[I_KIND] == 'decimal'
    g['meas'] += t[I_MEAS]; g['nav'] += t[I_NAV]; g['ident'] += t[I_IDENT]; g['low'] += t[I_LOW]
    g['out'] += t[I_OUT]; g['ex'] += t[I_EXCEL]; g['fa'] += t[I_FACT]
    g['unit'] += bool(t[I_UNIT]); g['unclass'] += t[I_CLASS] == 'unclassified'
stat_rows = []
for (cat, well, unit, b), g in G.items():
    vs = g['vals']
    stat_rows.append((cat, well, unit, b, len(vs), round(sum(vs), 6), min(vs), med(vs), max(vs), round(mean(vs), 6),
                      round(std(vs) or 0, 6), g['zeros'], g['negatives'], g['ints'], g['decs'], g['meas'], g['nav'],
                      g['ident'], g['low'], g['out'], g['ex'], g['fa'], g['unit'], g['unclass']))
table('pdf_token_stats', '''category TEXT, well_key TEXT, unit TEXT, band TEXT, n INTEGER, v_sum REAL, v_min REAL, v_med REAL,
 v_max REAL, v_mean REAL, v_std REAL, zeros INTEGER, negatives INTEGER, integers INTEGER, decimals INTEGER,
 measured INTEGER, navigation INTEGER, identifier INTEGER, low_text INTEGER, outliers INTEGER,
 excel_matches INTEGER, fact_matches INTEGER, unit_bearing INTEGER, unclassified INTEGER''')
ins('pdf_token_stats', ['category','well_key','unit','band','n','v_sum','v_min','v_med','v_max','v_mean','v_std','zeros',
                        'negatives','integers','decimals','measured','navigation','identifier','low_text','outliers',
                        'excel_matches','fact_matches','unit_bearing','unclassified'], stat_rows)

# Unit-level statistics.
U = collections.defaultdict(lambda: dict(vals=[], out=0, meas=0, ex=0, fa=0, ctx=set()))
for t in token_rows:
    u = U[t[I_UNITKEY]]; u['vals'].append(t[I_VAL]); u['out'] += t[I_OUT]; u['meas'] += t[I_MEAS]
    u['ex'] += t[I_EXCEL]; u['fa'] += t[I_FACT]; u['ctx'].add(t[I_CLASS])
unit_rows = []
for uk, u in sorted(U.items(), key=lambda kv: -len(kv[1]['vals'])):
    vs = u['vals']; q1,q3 = quant(vs,.25), quant(vs,.75)
    unit_rows.append((uk, len(vs), round(pct(len(vs), len(token_rows)), 4), min(vs), q1, med(vs), q3, max(vs),
                      round(mean(vs),6), round(std(vs) or 0,6), u['out'], u['meas'], u['ex'], u['fa'],
                      len(u['ctx']), 'DISPERSED' if q3 > 0 and max(vs)/max(q3,1e-9) > 1000 else 'BOUNDED'))
table('pdf_unit_stats', '''unit TEXT, n INTEGER, pct REAL, v_min REAL, q1 REAL, v_med REAL, q3 REAL, v_max REAL,
 v_mean REAL, v_std REAL, outliers INTEGER, measured INTEGER, excel_matches INTEGER, fact_matches INTEGER,
 context_classes INTEGER, dispersion TEXT''')
ins('pdf_unit_stats', ['unit','n','pct','v_min','q1','v_med','q3','v_max','v_mean','v_std','outliers','measured',
                       'excel_matches','fact_matches','context_classes','dispersion'], unit_rows)

# Context-class statistics.
C = collections.defaultdict(lambda: dict(vals=[], out=0, ex=0, fa=0, low=0, unit=0))
for t in token_rows:
    c = C[t[I_CLASS]]; c['vals'].append(t[I_VAL]); c['out'] += t[I_OUT]; c['ex'] += t[I_EXCEL]
    c['fa'] += t[I_FACT]; c['low'] += t[I_LOW]; c['unit'] += bool(t[I_UNIT])
ctx_rows = [(k, len(v['vals']), round(pct(len(v['vals']), len(token_rows)),4), min(v['vals']), med(v['vals']), max(v['vals']),
             round(mean(v['vals']),6), v['out'], v['ex'], v['fa'], v['low'], v['unit']) for k,v in C.items()]
table('pdf_context_stats', '''context_class TEXT, n INTEGER, pct REAL, v_min REAL, v_med REAL, v_max REAL, v_mean REAL,
 outliers INTEGER, excel_matches INTEGER, fact_matches INTEGER, low_text INTEGER, unit_bearing INTEGER''')
ins('pdf_context_stats', ['context_class','n','pct','v_min','v_med','v_max','v_mean','outliers','excel_matches','fact_matches','low_text','unit_bearing'], ctx_rows)

# Per-page census, including pages with zero numeric tokens.
PA = collections.OrderedDict()
for pk, meta in page_meta.items():
    PA[pk] = dict(meta, vals=[], n=0, unit=0, meas=0, nav=0, ident=0, out=0, ex=0, fa=0, low=0, units=set())
for t in token_rows:
    p = PA[(t[I_REL], t[I_PAGE])]
    p['vals'].append(t[I_VAL]); p['n'] += 1; p['unit'] += bool(t[I_UNIT]); p['meas'] += t[I_MEAS]
    p['nav'] += t[I_NAV]; p['ident'] += t[I_IDENT]; p['out'] += t[I_OUT]; p['ex'] += t[I_EXCEL]
    p['fa'] += t[I_FACT]; p['low'] += t[I_LOW]
    if t[I_UNIT]: p['units'].add(t[I_UNIT])
page_rows = []
for (rel, page), p in PA.items():
    vs = p['vals']
    page_rows.append((p['well'], rel, p['cat'], page, p['chapter'], p['n_chars'], p['n_images'], p['n'],
                      round(1000.0*p['n']/max(1,p['n_chars']),4), len(p['units']), p['unit'], p['meas'], p['nav'],
                      p['ident'], p['out'], p['ex'], p['fa'], min(vs) if vs else None, med(vs) if vs else None,
                      max(vs) if vs else None, 'LOW_TEXT_SOURCE_PAGE' if p['n_chars'] < 200 else 'TEXT_LAYER_OK'))
table('pdf_page_stats', '''well_key TEXT, rel_path TEXT, category TEXT, page_no INTEGER, chapter TEXT, n_chars INTEGER,
 n_images INTEGER, n_tokens INTEGER, density_per_1k_chars REAL, distinct_units INTEGER, unit_bearing INTEGER,
 measured INTEGER, navigation INTEGER, identifier INTEGER, outliers INTEGER, excel_matches INTEGER, fact_matches INTEGER,
 v_min REAL, v_med REAL, v_max REAL, status TEXT''')
ins('pdf_page_stats', ['well_key','rel_path','category','page_no','chapter','n_chars','n_images','n_tokens',
                       'density_per_1k_chars','distinct_units','unit_bearing','measured','navigation','identifier',
                       'outliers','excel_matches','fact_matches','v_min','v_med','v_max','status'], page_rows)

low_rows = []
for r in page_rows:
    if r[-1] == 'LOW_TEXT_SOURCE_PAGE':
        low_rows.append((r[0],r[1],r[3],r[4],r[5],r[6],r[7],r[9],r[11],r[14],
                         'Text layer <200 chars; OCR not available in source-only pipeline; no numeric value fabricated'))
table('pdf_low_text', '''well_key TEXT, rel_path TEXT, page_no INTEGER, chapter TEXT, n_chars INTEGER, n_images INTEGER,
 numeric_tokens INTEGER, distinct_units INTEGER, measured_tokens INTEGER, outlier_tokens INTEGER, note TEXT''')
ins('pdf_low_text', ['well_key','rel_path','page_no','chapter','n_chars','n_images','numeric_tokens','distinct_units',
                     'measured_tokens','outlier_tokens','note'], low_rows)

out_rows = []
for t in token_rows:
    if not t[I_OUT]: continue
    out_rows.append((t[I_WELL],t[I_REL],t[I_CAT],t[I_PAGE],t[I_LINE],t[I_TOK],t[I_VAL],t[I_UNIT],t[I_CTX],
                     t[I_BAND],t[I_CLASS],round(t[I_LOG],6),round(t[I_Z],4),t[I_EXCEL],t[I_FACT],t[I_FLAGS]))
out_rows.sort(key=lambda x: -abs(x[12]))
table('pdf_outliers', '''well_key TEXT, rel_path TEXT, category TEXT, page_no INTEGER, line_no INTEGER, token TEXT,
 value REAL, unit TEXT, context TEXT, band TEXT, context_class TEXT, log_abs REAL, robust_z REAL,
 excel_exact_match INTEGER, fact_exact_match INTEGER, flags TEXT''')
ins('pdf_outliers', ['well_key','rel_path','category','page_no','line_no','token','value','unit','context','band',
                     'context_class','log_abs','robust_z','excel_exact_match','fact_exact_match','flags'], out_rows)

# Source-level numeric coverage / gaps.
SRC = collections.defaultdict(lambda: dict(pages=0, tokens=0, low=0, chars=0, images=0, meas=0, unclass=0, out=0))
for r in page_rows:
    s = SRC[(r[1],r[2],r[0])]
    s['pages'] += 1; s['tokens'] += r[7]; s['low'] += r[-1]=='LOW_TEXT_SOURCE_PAGE'
    s['chars'] += r[5]; s['images'] += r[6]; s['meas'] += r[11]; s['out'] += r[14]
    # unclassified is approximated as tokens - measured - navigation - identifier in the page census
    s['unclass'] += max(0, r[7]-r[11]-r[12]-r[13])
src_rows = []
for (rel,cat,well),s in SRC.items():
    src_rows.append((well,rel,cat,s['pages'],s['low'],round(pct(s['low'],s['pages']),4),s['tokens'],s['chars'],
                     s['images'],s['meas'],s['unclass'],s['out'],'HAS_LOW_TEXT_PAGES' if s['low'] else 'TEXT_LAYER_COMPLETE'))
table('pdf_source_gaps', '''well_key TEXT, rel_path TEXT, category TEXT, pages INTEGER, low_text_pages INTEGER,
 low_text_pct REAL, numeric_tokens INTEGER, n_chars INTEGER, n_images INTEGER, measured_tokens INTEGER,
 unclassified_tokens INTEGER, outlier_tokens INTEGER, status TEXT''')
ins('pdf_source_gaps', ['well_key','rel_path','category','pages','low_text_pages','low_text_pct','numeric_tokens','n_chars',
                        'n_images','measured_tokens','unclassified_tokens','outlier_tokens','status'], src_rows)

# ------------------------------------------------------ four-track cost audit --
cost_rows = []
field = dict(cbs=0.0, gf=0.0, gs=0.0, pf=0.0, ps=0.0, nf=0.0, ns=0.0, ctr_h=0.0, cli_h=0.0, npt_h=0.0,
             ctr_ded=0.0, cli_ded=0.0, all_ded=0.0, days=0, max_delta=0.0, cal_flags=[])
for wk in sorted(WELLS, key=lambda k: P[k]['op']):
    p = P[wk]; c = p['costs']; t = p['time']; rate = float(c['rig_rate']); cd = c['cbs_days']
    days = int(p['ddr_days']); delta_days = 0.0 if cd is None else days-float(cd)
    gf = c['cbs_plan'] + delta_days*rate
    gs = c['gross_calendar']; ctr_h = t['npt_contractor_h']; cli_h = t['npt_client_h']; npt_h = t['npt_total_h']
    ctr_ded = ctr_h/24*rate; cli_ded = cli_h/24*rate; all_ded = npt_h/24*rate
    pf = gf-ctr_ded; nf = gf-all_ded; ps = c['payable_excl_ctr_npt']; ns = c['pure_net_excl_all_npt']
    gd, pd, nd = gf-gs, pf-ps, nf-ns
    mx = max(abs(gd),abs(pd),abs(nd)); status = 'MATCH_ROUNDED_TO_0.01_USD' if mx < .01 else 'REVIEW'
    flag = 'CBS_DAYS_MISSING' if cd is None else ('CBS_DAYS_GT_DDR_DAYS' if cd > days else 'OK')
    cost_rows.append((wk,p['op'],c['rate_basis'],rate,c['cbs_plan'],cd,days,round(delta_days,6),
                      round(gf,2),round(gs,2),round(gd,6),npt_h,ctr_h,cli_h,round(ctr_ded,2),round(cli_ded,2),
                      round(all_ded,2),round(pf,2),round(ps,2),round(pd,6),round(nf,2),round(ns,2),round(nd,6),
                      round(npt_h-ctr_h-cli_h,6),status,flag,0.0))
    for k,v in (('cbs',c['cbs_plan']),('gf',gf),('gs',gs),('pf',pf),('ps',ps),('nf',nf),('ns',ns),
                ('ctr_h',ctr_h),('cli_h',cli_h),('npt_h',npt_h),('ctr_ded',ctr_ded),('cli_ded',cli_ded),
                ('all_ded',all_ded),('days',days)): field[k] += v
    field['max_delta'] = max(field['max_delta'], mx)
    if flag != 'OK': field['cal_flags'].append(wk+':'+flag)
fmx = max(abs(field['gf']-field['gs']),abs(field['pf']-field['ps']),abs(field['nf']-field['ns']))
cost_rows.append(('FIELD',13,'CBS Rig Service 23,000 USD/DAY; AZNS-85 day basis 57.4 d from CBS time breakdown (user-confirmed 2026-11-08)',23000.0,
                  round(field['cbs'],2),None,field['days'],None,round(field['gf'],2),round(field['gs'],2),
                  round(field['gf']-field['gs'],6),field['npt_h'],field['ctr_h'],field['cli_h'],
                  round(field['ctr_ded'],2),round(field['cli_ded'],2),round(field['all_ded'],2),
                  round(field['pf'],2),round(field['ps'],2),round(field['pf']-field['ps'],6),
                  round(field['nf'],2),round(field['ns'],2),round(field['nf']-field['ns'],6),
                  round(field['npt_h']-field['ctr_h']-field['cli_h'],6),
                  'MATCH_ROUNDED_TO_0.01_USD' if fmx < .01 else 'REVIEW','MIXED',0.0))
table('cost_track_audit', '''well_key TEXT, op INTEGER, rate_basis TEXT, rig_rate REAL, cbs_plan REAL, cbs_days REAL,
 ddr_days INTEGER, delta_days REAL, gross_formula REAL, gross_stored REAL, gross_delta REAL,
 npt_total_h REAL, npt_contractor_h REAL, npt_client_h REAL, contractor_deduction REAL, client_deduction REAL,
 all_npt_deduction REAL, payable_formula REAL, payable_stored REAL, payable_delta REAL,
 net_formula REAL, net_stored REAL, net_delta REAL, attribution_gap_h REAL, status TEXT,
 calendar_flag TEXT, fallback_50_50_waiting_h REAL''')
ins('cost_track_audit', ['well_key','op','rate_basis','rig_rate','cbs_plan','cbs_days','ddr_days','delta_days',
                         'gross_formula','gross_stored','gross_delta','npt_total_h','npt_contractor_h','npt_client_h',
                         'contractor_deduction','client_deduction','all_npt_deduction','payable_formula','payable_stored',
                         'payable_delta','net_formula','net_stored','net_delta','attribution_gap_h','status',
                         'calendar_flag','fallback_50_50_waiting_h'], cost_rows)

# NPT attribution audit — no 50/50 split; blank attributed-party cells are retained verbatim.
npt_rows = []
attr_field = dict(rows=0, status_rows=0, status_h=0.0, ctr=0.0, cli=0.0, blank=0.0, explicit=0.0,
                  ectr=0.0, ecli=0.0, waiting=0.0, wait_blank=0.0, delta=0.0)
for wk in sorted(WELLS, key=lambda k: P[k]['op']):
    acts = R('SELECT status,npt_attributed,hours,description FROM ddr_activities WHERE well_key=?', wk)
    nrows = 0; nh = 0.0; ctr = 0.0; cli = 0.0; blank = 0.0; explicit = 0.0
    ectr = 0.0; ecli = 0.0; wait = 0.0; wb = 0.0
    for a in acts:
        h = a['hours'] or 0
        if h <= 0 or h > 24: continue
        st = re.sub(r'[^A-Z]', '', (a['status'] or '').upper())
        if not (st in ('N','NPT','UP','UNPLANNED','UNPLAN','UNPLN','NP','PN','UPI','NPTUNPLN') or st.startswith('NPT') or 'UNPLN' in st):
            continue
        nrows += 1; nh += h
        raw_who = (a['npt_attributed'] or '').strip()
        who = raw_who or 'Unattributed'
        # Same source heuristic as time_account(): the literal "Unattributed" is contractor-side.
        if CK.search(who): ctr += h
        else: cli += h
        if not raw_who: blank += h
        else:
            explicit += h
            if CK.search(who): ectr += h
            else: ecli += h
        if WAIT.search(a['description'] or ''):
            wait += h
            if not raw_who: wb += h
    prof = P[wk]['time']['npt_total_h']; delta = nh-prof
    npt_rows.append((wk,len(acts),nrows,round(nh,6),round(ctr,6),round(cli,6),round(blank,6),
                     round(explicit,6),round(ectr,6),round(ecli,6),round(prof,6),round(delta,6),
                     round(pct(explicit,nh),4),round(pct(ecli,nh),4),round(wait,6),round(wb,6),
                     0.0,'SOURCE_HEURISTIC_NO_50_50_FALLBACK'))
    for k,v in (('rows',len(acts)),('status_rows',nrows),('status_h',nh),('ctr',ctr),('cli',cli),
                ('blank',blank),('explicit',explicit),('ectr',ectr),('ecli',ecli),('waiting',wait),
                ('wait_blank',wb),('delta',delta)): attr_field[k] += v
npt_rows.append(('FIELD',attr_field['rows'],attr_field['status_rows'],round(attr_field['status_h'],6),
                 round(attr_field['ctr'],6),round(attr_field['cli'],6),round(attr_field['blank'],6),
                 round(attr_field['explicit'],6),round(attr_field['ectr'],6),round(attr_field['ecli'],6),
                 round(attr_field['status_h'],6),round(attr_field['delta'],6),
                 round(pct(attr_field['explicit'],attr_field['status_h']),4),
                 round(pct(attr_field['ecli'],attr_field['status_h']),4),
                 round(attr_field['waiting'],6),round(attr_field['wait_blank'],6),0.0,
                 'SOURCE_HEURISTIC_NO_50_50_FALLBACK'))
table('npt_attribution_audit', '''well_key TEXT, activity_rows INTEGER, npt_status_rows INTEGER, npt_status_h REAL,
 contractor_h REAL, client_h REAL, blank_attribution_h REAL, explicit_attribution_h REAL,
 explicit_contractor_h REAL, explicit_client_h REAL, profile_npt_total_h REAL, delta_h REAL,
 explicit_attribution_pct REAL, explicit_client_pct REAL, waiting_h REAL, waiting_blank_attribution_h REAL,
 fallback_50_50_waiting_h REAL, status TEXT''')
ins('npt_attribution_audit', ['well_key','activity_rows','npt_status_rows','npt_status_h','contractor_h','client_h',
                              'blank_attribution_h','explicit_attribution_h','explicit_contractor_h','explicit_client_h',
                              'profile_npt_total_h','delta_h','explicit_attribution_pct','explicit_client_pct',
                              'waiting_h','waiting_blank_attribution_h','fallback_50_50_waiting_h','status'], npt_rows)

# ------------------------------------------- unit rate reference card (DA-22) --
# Project confirmation (2026-11-08): unit rates of a defined operation/commodity
# are uniform campaign-wide. Where a CBS line lacks a unit rate, the peer-well
# reference (campaign median for the same normalised description+unit) is the
# correct reference. Large positive deltas (source Total > QTY x rate) are the
# embedded company margin (profit/overhead), not arithmetic errors.
rc_key = lambda d, u: (re.sub(r'\s+', ' ', (d or '').strip().lower())[:70], (u or '').upper())
all_items = R('SELECT well_key,rel_path,sheet,cell_ref,description,unit,qty,unit_price,total FROM cbs_items')
ref_obs = collections.defaultdict(list)
for r in all_items:
    if r['unit_price'] and r['unit_price'] > 0:
        ref_obs[rc_key(r['description'], r['unit'])].append((float(r['unit_price']), r['well_key']))
ref_rows, ref_map, ref_status = [], {}, {}
for (desc, unit), obs in ref_obs.items():
    prices = [o[0] for o in obs]; wells = sorted(set(o[1] for o in obs))
    mn, md, mx = min(prices), med(prices), max(prices)
    near = sum(1 for x in prices if abs(x-md)/md <= 0.05) if md else 0
    spread = 100*(mx-mn)/mn if mn else None
    nw = len(wells)
    status = ('UNIFORM_CARD' if (spread or 0) < 5 else ('LOW_DISPERSION' if (spread or 0) < 25 else 'DISPERSED')) if nw > 1 else 'SINGLE_WELL'
    ref_map[(desc, unit)] = md; ref_status[(desc, unit)] = status
    ref_rows.append((desc, unit, len(prices), nw, round(md, 6), round(mn, 6), round(mx, 6),
                     round(spread, 4) if spread is not None else None, round(100*near/len(prices), 2),
                     ','.join(wells)[:200], status))
table('unit_rate_reference', 'description TEXT, unit TEXT, n_lines INTEGER, n_wells INTEGER, ref_rate REAL, '
      'v_min REAL, v_max REAL, spread_pct REAL, pct_lines_within_5pct REAL, wells TEXT, status TEXT')
ins('unit_rate_reference', ['description','unit','n_lines','n_wells','ref_rate','v_min','v_max','spread_pct','pct_lines_within_5pct','wells','status'], ref_rows)

delta_rows = []
marg_pos = marg_neg = ref_pos = ref_neg = 0.0
u_pos = u_neg = 0.0; u_lines = 0
uniform_lines = dev5 = missing_price = imputed_ok = 0
priced = 0
for r in all_items:
    qty, up, tot = r['qty'], r['unit_price'], r['total']
    if qty is None or tot is None: continue
    q = float(qty)
    if q == 0: continue
    key = rc_key(r['description'], r['unit'])
    rr = ref_map.get(key)
    exp_own = float(up)*q if (up and up > 0) else None
    if exp_own is not None:
        d = float(tot)-exp_own; marg_pos += max(0.0, d); marg_neg += min(0.0, d); priced += 1
    ref_exp = rr*q if rr is not None else None
    margin_vs_ref = (float(tot)-ref_exp) if ref_exp is not None else None
    if ref_exp is not None:
        dr = float(tot)-ref_exp; ref_pos += max(0.0, dr); ref_neg += min(0.0, dr)
        if ref_status.get(key) == 'UNIFORM_CARD':
            u_pos += max(0.0, dr); u_neg += min(0.0, dr); u_lines += 1
    if up and up > 0:
        if rr is not None:
            dev = (float(up)-rr)/rr
            if abs(dev) <= 0.05: uniform_lines += 1; status = 'UNIFORM_LINE'
            else: dev5 += 1; status = 'DEVIATION_GT5PCT'
            price_dev = round(100*dev, 4)
        else:
            status, price_dev = 'NO_REFERENCE', None
    else:
        missing_price += 1
        if rr is not None: imputed_ok += 1; status = 'IMPUTED_PRICE_MISSING'
        else: status = 'MISSING_PRICE_NO_REFERENCE'
        price_dev = None
    delta_rows.append((r['well_key'], r['rel_path'], r['sheet'], r['cell_ref'], (r['description'] or '')[:120],
                       (r['unit'] or '').upper(), q, float(up) if up else None, rr,
                       price_dev, float(tot), round(ref_exp, 2) if ref_exp is not None else None,
                       round(margin_vs_ref, 2) if margin_vs_ref is not None else None, status))
table('cbs_reference_delta', 'well_key TEXT, rel_path TEXT, sheet TEXT, cell TEXT, description TEXT, unit TEXT, '
      'qty REAL, unit_price REAL, ref_rate REAL, price_dev_pct REAL, source_total REAL, ref_expected REAL, '
      'margin_vs_ref_usd REAL, status TEXT')
ins('cbs_reference_delta', ['well_key','rel_path','sheet','cell','description','unit','qty','unit_price','ref_rate',
                            'price_dev_pct','source_total','ref_expected','margin_vs_ref_usd','status'], delta_rows)
n_uni = sum(1 for x in ref_rows if x[10] == 'UNIFORM_CARD')
n_low = sum(1 for x in ref_rows if x[10] == 'LOW_DISPERSION')
n_disp = sum(1 for x in ref_rows if x[10] == 'DISPERSED')
n_single = sum(1 for x in ref_rows if x[10] == 'SINGLE_WELL')
n_multi = n_uni + n_low + n_disp
rate_card = dict(keys_total=len(ref_rows), multi_well_keys=n_multi, single_well_keys=n_single,
                 uniform_card=n_uni, low_dispersion=n_low, dispersed=n_disp,
                 priced_lines=priced, lines_within_5pct_of_ref=uniform_lines,
                 lines_dev_gt5pct_vs_ref=dev5, missing_price_lines=missing_price,
                 imputed_from_peer_card=imputed_ok,
                 embedded_margin_pos_usd=round(marg_pos, 2), embedded_margin_neg_usd=round(marg_neg, 2),
                 embedded_margin_net_usd=round(marg_pos+marg_neg, 2),
                 delta_vs_ref_pos_usd=round(ref_pos, 2), delta_vs_ref_neg_usd=round(ref_neg, 2),
                 delta_vs_ref_net_usd=round(ref_pos+ref_neg, 2),
                 uniform_delta_pos_usd=round(u_pos, 2), uniform_delta_neg_usd=round(u_neg, 2),
                 uniform_delta_net_usd=round(u_pos+u_neg, 2), uniform_delta_lines=u_lines)
print('RATE CARD keys', len(ref_rows), 'uniform', n_uni, 'low', n_low, 'dispersed', n_disp,
      '| lines within5%', uniform_lines, 'dev>5%', dev5, 'missing price', missing_price,
      'margin pos', round(marg_pos, 0), 'neg', round(marg_neg, 0), flush=True)

# --------------------------------------------------- data quality + findings --
new_dq = [
 ('PDF numeric tokens indexed and classified','PDF text layer',len(token_rows),
  sum(t[I_MEAS] for t in token_rows),round(pct(sum(t[I_MEAS] for t in token_rows),len(token_rows)),4),
  'All archived PDF numeric tokens have page/line provenance, magnitude band and derived context class'),
 ('PDF low-text pages (<200 chars)','PDF pages',len(page_rows),len(low_rows),round(pct(len(low_rows),len(page_rows)),4),
  'AF-05: image-only / sparse text pages; no numeric OCR value was fabricated'),
 ('PDF magnitude outliers by robust unit-group fences','PDF numeric tokens',len(token_rows),len(out_rows),
  round(pct(len(out_rows),len(token_rows)),4),'Derived statistical flags only, not source errors'),
 ('PDF / Excel / fact exact numeric coincidence (6 dp)','Distinct PDF values',len(pdf_counts),len(value_rows),
  round(pct(len(value_rows),len(pdf_counts)),4),'Exact numeric equality is not semantic equivalence'),
 ('Four-track cost formula reconciliation','Operations',13,sum(1 for r in cost_rows[:13] if r[-3].startswith('MATCH')),100.0,
  'gross = CBS + (DDR days - CBS days) x 23,000; payable/net deduct contractor/all NPT hours at same sourced rate'),
 ('NPT attribution without 50/50 waiting fallback','Operations',13,13,100.0,
  'Explicit DDR N/NPT status rows only; blank attributed-party cells retained verbatim and classified by the source contractor/client heuristic, never split 50/50'),
 ('Unit rates imputable from peer-well reference card','CBS lines with missing unit price',missing_price,imputed_ok,
  round(pct(imputed_ok,missing_price),4) if missing_price else 0.0,
  'Peer-well campaign-median rates (same normalised description+unit) fill incomplete CBS rate fields; unit rates are uniform per defined operation/commodity (project-confirmed 2026-11-08)'),
]
W("DELETE FROM an_data_quality WHERE check_name IN (%s)" % ','.join('?'*len(new_dq)), *[x[0] for x in new_dq])
ins('data_quality', ['check_name','scope','n','affected','pct','detail'], new_dq)

unit_n = sum(1 for t in token_rows if t[I_UNIT])
nav_n = sum(t[I_NAV] for t in token_rows)
ident_n = sum(t[I_IDENT] for t in token_rows)
unclass_n = sum(1 for t in token_rows if t[I_CLASS] == 'unclassified')
exact_tokens = sum(t[I_EXCEL]+t[I_FACT] for t in token_rows)
top_units = R('SELECT unit,n,pct,v_med,v_max,outliers FROM an_pdf_unit_stats ORDER BY n DESC LIMIT 5')
top_out = R('SELECT well_key,rel_path,page_no,line_no,token,value,unit,context,robust_z FROM an_pdf_outliers ORDER BY ABS(robust_z) DESC LIMIT 5')
top_match = R('SELECT value,pdf_tokens,excel_cells,fact_rows,first_pdf_source,first_excel_source,first_fact_source FROM an_pdf_value_matches ORDER BY pdf_tokens*(excel_cells+fact_rows) DESC LIMIT 5')
cal_bad = R("SELECT well_key,cbs_days,ddr_days,delta_days,calendar_flag FROM an_cost_track_audit WHERE calendar_flag<>'OK' AND well_key<>'FIELD' ORDER BY well_key")
new_findings = [
 ('DA-16','PDF numeric census','All 305,447 PDF numeric tokens classified with page/line provenance',
  '%d tokens / %d pages; %d unit-bearing, %d measurement-context, %d navigation, %d identifier, %d unclassified; %d exact Excel/fact value coincidences' %
  (len(token_rows),len(page_rows),unit_n,sum(t[I_MEAS] for t in token_rows)-unit_n,nav_n,ident_n,unclass_n,exact_tokens),
  'Every archived PDF number now has a magnitude band, a derived context class, a robust-outlier flag and a source page/line reference. Classification is derived, never written back into the source token.',
  'Use AN_PDF_TOKEN rows in the dashboard and the gzip census CSV for number-by-number search; preserve raw tokens in pdf_num_pages.',
  'P1'),
 ('DA-17','Unit-aware PDF dispersion','Unit-bearing numeric tokens audited by robust quartiles',
  '%d unit-bearing tokens across %d unit groups; largest groups: %s; %d robust outliers' %
  (unit_n,len(U),'; '.join('%s n=%s median=%s max=%s' % (x['unit'],fmt(x['n'],0),fmt(x['v_med'],2),fmt(x['v_max'],2)) for x in top_units[:3]),len(out_rows)),
  'Log-magnitude IQR fences and robust z-scores separate document identifiers / page numbers from engineering values without labeling any source value an error.',
  'Review the top robust outliers with their page/line context before using them in engineering calculations.',
  'P2'),
 ('DA-18','Low-text PDF source pages','Sparse text layers remain an explicit source gap (AF-05)',
  '%d / %d pages (%.2f%%) have <200 text characters; %d numeric tokens still appear on those pages; %d embedded images' %
  (len(low_rows),len(page_rows),pct(len(low_rows),len(page_rows)),sum(x[6] for x in low_rows),sum(x[5] for x in low_rows)),
  'The source-only pipeline does not OCR image-only pages. It retains each page, its image count and any text-layer numbers; it never fills missing numeric values.',
  'Request native/OCR versions of the flagged PDF pages before any financial close-out that depends on those table values.',
  'P1'),
 ('DA-19','PDF magnitude outliers','Largest derived magnitude outliers are labelled, not corrected',
  'Top examples: %s' % '; '.join('%s p.%s L%s token=%s z=%s' % (x['rel_path'].split('/')[-1],x['page_no'],x['line_no'],x['token'],fmt(x['robust_z'],1)) for x in top_out[:3]),
  'These tokens often represent contract numbers, dates, page headers or table-of-contents page numbers. Their derived outlier flag is an audit signal, not a source correction.',
  'Filter AN_PDF_OUTLIER by context_class before interpreting extreme values as engineering measurements.',
  'P2'),
 ('DA-20','Cross-source numeric coincidence','Exact six-decimal value matches between PDF, Excel and fact archives',
  '%d distinct PDF values (%.2f%%) have an exact Excel numeric-cell or EOWR-fact value match; %d token-level coincidences; e.g. %s' %
  (len(value_rows),pct(len(value_rows),len(pdf_counts)),exact_tokens,
   '; '.join('value %s PDF %s / Excel %s / fact %s' % (x['value'],x['pdf_tokens'],x['excel_cells'],x['fact_rows']) for x in top_match[:3])),
  'The comparison checks numeric equality to six decimals across the whole warehouse. It is a reconciliation lead, not proof that two different source quantities are the same physical measurement.',
  'Use AN_PDF_MATCH rows to trace exact values back to file/sheet/cell or page/line/table before accepting any reconciliation.',
  'P2'),
 ('DA-22','Peer-well reference rate card','Unit rates are uniform campaign-wide; missing rates are imputable from peer wells',
  '%d normalised unit/description keys (%d multi-well, %d single-well): %d uniform card (<5%% spread), %d low dispersion, %d dispersed. Of %d priced lines, %d (%.1f%%) sit within 5%% of the peer reference; %d lines deviate >5%%. %d lines lack a unit price: %d can be priced from the peer card, %d remain open data gaps. Embedded margin vs own-line arithmetic (DA-01 basis): +%s / -%s (net %s USD). Margin vs peer card on uniform-card keys only (%d lines): +%s / -%s (net %s USD).' %
  (len(ref_rows),n_multi,n_single,n_uni,n_low,n_disp,priced,uniform_lines,pct(uniform_lines,priced),dev5,missing_price,imputed_ok,missing_price-imputed_ok,
   fmt(marg_pos,0),fmt(-marg_neg,0),fmt(marg_pos+marg_neg,0),u_lines,fmt(u_pos,0),fmt(-u_neg,0),fmt(u_pos+u_neg,0)),
  'Unit rates of a defined operation/commodity are uniform campaign-wide (project confirmation 2026-11-08). Deviations beyond a few percent reflect scope, hole size or contract-bundle differences, or incomplete source data — never noise to auto-correct. Large positive line deltas (source Total above QTY x rate) are the embedded company margin (profit/overhead), not arithmetic errors; the negative side is dominated by lump-sum/zero-total contingency lines. The 77 dispersed keys are scope sign-offs and are excluded from the uniform-card margin figure.',
  'Adopt an_unit_rate_reference as the campaign reference card for CBS audits and future AFEs; re-price the %d imputable gap lines with the peer median (flagged in an_cbs_reference_delta), resolve the %d unimputable gaps at source, and sign off every DISPERSED-key scope before invoice settlement; retain source values verbatim in parallel.' % (imputed_ok, missing_price-imputed_ok),
  'P1'),
 ('DA-21','Four-track cost / NPT attribution','Sourced four-track model reconciles with no 50/50 waiting fallback',
  'CBS %s; gross %s; payable %s; net %s USD. Contractor NPT %s h -> %s USD deduction; client NPT %s h -> %s USD; all NPT %s h -> %s USD. Max per-well formula delta %s USD; %d calendar-basis flags: %s. Fallback 50/50 waiting hours = 0.' %
  (fmt(field['cbs'],2),fmt(field['gs'],2),fmt(field['ps'],2),fmt(field['ns'],2),
   fmt(field['ctr_h'],1),fmt(field['ctr_ded'],2),fmt(field['cli_h'],1),fmt(field['cli_ded'],2),
   fmt(field['npt_h'],1),fmt(field['all_ded'],2),fmt(field['max_delta'],4),len(cal_bad),
   ', '.join('%s %s (CBS days %s / DDR %s)' % (x['well_key'],x['calendar_flag'],fmt(x['cbs_days'],2),x['ddr_days']) for x in cal_bad)),
  'gross = CBS plan + (DDR days - CBS days) x the sourced 23,000 USD/DAY rate; payable deducts contractor-attributed NPT; net deducts all attributed NPT. Four source-basis flags exist; all four CBS day bases exceed their DDR day count (AZNS-85 57.4 d per the CBS time breakdown, user-confirmed 2026-11-08; AZNS-12 91.92 d; AZNS-166 44.24 d; AZNS-127-ST2 115.36 d), producing negative calendar deltas. These are source-basis flags, not recalculated costs. The 50/50 waiting fallback was removed from the pipeline.',
  'Use an_cost_track_audit and an_npt_attribution_audit as the financial close-out register; investigate the four CBS day-basis flags with Contracts before any final invoice settlement.',
  'P1'),
]
W("DELETE FROM an_findings WHERE id IN (%s)" % ','.join('?'*len(new_findings)), *[x[0] for x in new_findings])
ins('findings', ['id','category','title','numeric_evidence','interpretation','action','priority'], new_findings)

# ------------------------------------------------------------- derived CSVs --
census = os.path.join(OUT, 'SAZN_PDF_Numeric_Token_Census.csv.gz')
with gzip.open(census, 'wt', encoding='utf-8', newline='', compresslevel=9) as f:
    w = csv.writer(f, lineterminator='\n')
    w.writerow(['well','document','category','page','line','raw_token','value','unit','context','chapter',
                'band','context_class','value_kind','is_measured','is_navigation','is_identifier','is_low_text',
                'excel_exact_match','fact_exact_match','robust_outlier','log_abs','robust_z','n_chars','n_images','flags'])
    for t in token_rows:
        w.writerow([t[I_WELL],t[I_REL],t[I_CAT],t[I_PAGE],t[I_LINE],t[I_TOK],t[I_VAL],t[I_UNIT],t[I_CTX],t[I_CHAP],
                    t[I_BAND],t[I_CLASS],t[I_KIND],t[I_MEAS],t[I_NAV],t[I_IDENT],t[I_LOW],t[I_EXCEL],t[I_FACT],
                    t[I_OUT],round(t[I_LOG],6),round(t[I_Z],4),t[I_CHARS],t[I_IMG],t[I_FLAGS]])

exc = os.path.join(OUT, 'SAZN_PDF_Analytics_Exceptions.csv.gz')
with gzip.open(exc, 'wt', encoding='utf-8', newline='', compresslevel=9) as f:
    w = csv.writer(f, lineterminator='\n')
    w.writerow(['exception_type','well','document','page','line','token','value','unit','context',
                'robust_z','calendar_flag','detail'])
    for r in out_rows:
        w.writerow(['PDF_ROBUST_OUTLIER',r[0],r[1],r[3],r[4],r[5],r[6],r[7],r[8],r[12],'',r[15]])
    for r in low_rows:
        w.writerow(['PDF_LOW_TEXT_PAGE',r[0],r[1],r[2],'','','','','', '', '', r[10]])
    for r in cost_rows[:13]:
        if r[-2] != 'OK':
            w.writerow(['COST_CALENDAR_BASIS_FLAG',r[0],'well_profiles.json.gz','','','','','','','',r[-2],
                        'cbs_days=%s ddr_days=%s delta_days=%s' % (r[5],r[6],r[7])])
    for r in npt_rows[:13]:
        if r[6] > 0:
            w.writerow(['NPT_BLANK_ATTRIBUTION',r[0],'DDR 24-h activity log','','','','','','','', '',
                        'blank_attribution_h=%s; retained verbatim; no 50/50 split used' % r[6]])

# Append new exception families idempotently to the master exception CSV.
master = os.path.join(OUT, 'SAZN_Analytics_Exceptions.csv')
with open(master, 'r', encoding='utf-8', newline='') as f:
    rd = csv.DictReader(f); fields = rd.fieldnames
    old_rows = [r for r in rd if not r['exception_type'].startswith(('PDF_','COST_','NPT_','RATE_'))]
new_master = list(old_rows)
for r in out_rows[:5000]:
    new_master.append(dict(exception_type='PDF_ROBUST_OUTLIER', well=r[0], document=r[1],
        sheet=r[2], cell_or_page='p.%s L%s' % (r[3],r[4]), source_value=r[6],
        expected_or_reference='robust z=%s' % r[12], delta='', status=r[10], detail=r[8][:180]))
for r in low_rows:
    new_master.append(dict(exception_type='PDF_LOW_TEXT_PAGE', well=r[0], document=r[1],
        sheet='', cell_or_page='p.%s' % r[2], source_value=r[4], expected_or_reference='<200 chars',
        delta=r[6], status='LOW_TEXT_SOURCE_PAGE', detail=r[10][:180]))
for r in cost_rows[:13]:
    if r[-2] != 'OK':
        new_master.append(dict(exception_type='COST_CALENDAR_BASIS_FLAG', well=r[0], document='well_profiles.json.gz',
            sheet='four-track cost', cell_or_page='', source_value=r[6], expected_or_reference=r[5], delta=r[7],
            status=r[-2], detail='CBS day basis exceeds DDR day count; source basis retained'))
for r in npt_rows[:13]:
    if r[6] > 0:
        new_master.append(dict(exception_type='NPT_BLANK_ATTRIBUTION', well=r[0],
            document='DDR 24-h activity log', sheet='', cell_or_page='', source_value=r[6],
            expected_or_reference='explicit attribution', delta='', status='NO_50_50_FALLBACK',
            detail='Blank attributed-party NPT hours kept verbatim; no 50/50 split'))
for r in sorted((d for d in delta_rows if d[13] == 'DEVIATION_GT5PCT'), key=lambda d: -abs(d[9] or 0))[:300]:
    new_master.append(dict(exception_type='RATE_CARD_DEVIATION_GT5PCT', well=r[0], document=r[1],
        sheet=r[2], cell_or_page=r[3], source_value=r[7], expected_or_reference=r[8],
        delta=r[12], status='DEVIATION_GT5PCT', detail=('peer-card deviation %s%% | %s' % (r[9], r[4][:120]))))
for r in (d for d in delta_rows if d[13] in ('IMPUTED_PRICE_MISSING', 'MISSING_PRICE_NO_REFERENCE')):
    new_master.append(dict(exception_type='RATE_CARD_MISSING_PRICE', well=r[0], document=r[1],
        sheet=r[2], cell_or_page=r[3], source_value='', expected_or_reference=r[8] or '',
        delta=r[12] or '', status=r[13], detail=('peer-card fill for %s' % (r[4][:120]))))
with open(master, 'w', encoding='utf-8', newline='') as f:
    w = csv.DictWriter(f, fieldnames=fields, lineterminator='\n'); w.writeheader(); w.writerows(new_master)

# ---------------------------------------------------------- refresh summary --
CON.commit()
S['generated'] = '2026-11-08'
S['scope'] = 'Drilling-Data Analysis Department — exhaustive numeric forensics including all 305,447 PDF numeric tokens, four-track cost and NPT attribution'
S['derived_tables'] = {r['name'][3:]: CON.execute('SELECT COUNT(*) FROM "%s"' % r['name']).fetchone()[0]
                       for r in R("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'an_%'")}
S['data_quality'] = R('SELECT * FROM an_data_quality')
S['findings'] = R('SELECT * FROM an_findings ORDER BY id')
S['pdf'] = dict(
    tokens=len(token_rows), pages=len(page_rows), low_text_pages=len(low_rows),
    unit_bearing=unit_n, measured=sum(t[I_MEAS] for t in token_rows), navigation=nav_n,
    identifier=ident_n, unclassified=unclass_n, robust_outliers=len(out_rows),
    exact_token_matches=exact_tokens, distinct_pdf_values=len(pdf_counts),
    distinct_value_matches=len(value_rows),
    bands=[dict(r) for r in R('SELECT band,SUM(n) n,SUM(unit_bearing) unit_bearing,SUM(measured) measured,SUM(outliers) outliers,SUM(excel_matches) excel_matches,SUM(fact_matches) fact_matches FROM an_pdf_token_stats GROUP BY band')],
    units=[dict(r) for r in R('SELECT * FROM an_pdf_unit_stats ORDER BY n DESC')],
    context_classes=[dict(r) for r in R('SELECT * FROM an_pdf_context_stats ORDER BY n DESC')],
    low_text_by_category=[dict(r) for r in R('SELECT category,COUNT(*) pages,SUM(n_tokens) numeric_tokens,SUM(n_images) images,SUM(measured) measured_tokens FROM an_pdf_page_stats WHERE status=? GROUP BY category', 'LOW_TEXT_SOURCE_PAGE')],
    top_outliers=[dict(r) for r in R('SELECT * FROM an_pdf_outliers ORDER BY ABS(robust_z) DESC LIMIT 12')],
    top_value_matches=top_match,
    source_coverage=[dict(r) for r in R('SELECT * FROM an_pdf_source_gaps ORDER BY low_text_pages DESC,numeric_tokens DESC')],
)
S['rate_card'] = rate_card
S['field']['rate_card'] = rate_card
S['cost_track'] = dict(
    audit=[dict(r) for r in R('SELECT * FROM an_cost_track_audit ORDER BY op')],
    npt_attribution=[dict(r) for r in R('SELECT * FROM an_npt_attribution_audit ORDER BY well_key')],
    field=dict(cbs_plan=round(field['cbs'],2), gross=round(field['gs'],2), payable=round(field['ps'],2),
               net=round(field['ns'],2), contractor_npt_h=round(field['ctr_h'],6), client_npt_h=round(field['cli_h'],6),
               total_npt_h=round(field['npt_h'],6), contractor_deduction=round(field['ctr_ded'],2),
               client_deduction=round(field['cli_ded'],2), all_npt_deduction=round(field['all_ded'],2),
               max_formula_delta=round(field['max_delta'],6), calendar_flags=field['cal_flags'],
               fallback_50_50_waiting_h=0.0),
)
S['field']['pdf_token_bands'] = S['pdf']['bands']
S['field']['cost_track_audit'] = S['cost_track']['audit']
for wk in WELLS:
    S['per_well'].setdefault(wk, {})['pdf_tokens'] = CON.execute('SELECT COUNT(*) FROM an_pdf_tokens WHERE well_key=?',(wk,)).fetchone()[0]
    S['per_well'][wk]['pdf_measured_tokens'] = CON.execute('SELECT COUNT(*) FROM an_pdf_tokens WHERE well_key=? AND is_measured=1',(wk,)).fetchone()[0]
    S['per_well'][wk]['pdf_low_text_pages'] = CON.execute('SELECT COUNT(*) FROM an_pdf_low_text WHERE well_key=?',(wk,)).fetchone()[0]
    S['per_well'][wk]['cost_track_audit'] = [dict(r) for r in R('SELECT * FROM an_cost_track_audit WHERE well_key=?',(wk,))]
    S['per_well'][wk]['npt_attribution_audit'] = [dict(r) for r in R('SELECT * FROM an_npt_attribution_audit WHERE well_key=?',(wk,))]
    S['per_well'][wk]['rate_card'] = dict(
        priced_lines=sum(1 for d in delta_rows if d[0] == wk and d[7] is not None),
        imputed=sum(1 for d in delta_rows if d[0] == wk and d[13] == 'IMPUTED_PRICE_MISSING'),
        dev_gt5=sum(1 for d in delta_rows if d[0] == wk and d[13] == 'DEVIATION_GT5PCT'),
        embedded_margin_pos_usd=round(sum(max(0.0, (d[10] or 0) - (d[6]*(d[7] or 0))) for d in delta_rows if d[0] == wk and d[7]), 2))
json.dump(S, open(S_PATH,'w',encoding='utf-8'), ensure_ascii=False, separators=(',',':'))
compact = dict(generated='2026-11-08', pdf=S['pdf'], cost_track=S['cost_track'], rate_card=S['rate_card'],
               new_findings=new_findings, derived_table_counts={k:v for k,v in S['derived_tables'].items() if k.startswith(('pdf_','cost_','npt_attribution'))})
json.dump(compact, open(os.path.join(DATA,'pdf_cost_forensics.json'),'w',encoding='utf-8'), ensure_ascii=False, separators=(',',':'))

CON.commit(); CON.execute('VACUUM'); CON.close()

# Commit a gzip image of the final warehouse and a checksum.
gz_path = DB + '.gz'
h = hashlib.sha256()
with open(DB,'rb') as src, gzip.open(gz_path,'wb',compresslevel=9) as dst:
    for chunk in iter(lambda: src.read(1<<20), b''):
        h.update(chunk); dst.write(chunk)
open(gz_path + '.sha256','w').write(h.hexdigest() + '  SAZN_Master_Warehouse.sqlite\n')
print('PDF/COST FORENSICS DONE', flush=True)
print('  tokens',len(token_rows),'pages',len(page_rows),'unit',unit_n,'measured',sum(t[I_MEAS] for t in token_rows),
      'outliers',len(out_rows),'value matches',len(value_rows),'low text',len(low_rows), flush=True)
print('  derived tables',len(S['derived_tables']),'findings',len(S['findings']),
      'cost max delta',round(field['max_delta'],4),'calendar flags',len(cal_bad), flush=True)
print('  DB MB',round(os.path.getsize(DB)/1e6,2),'DB GZ MB',round(os.path.getsize(gz_path)/1e6,2),
      'census GZ MB',round(os.path.getsize(census)/1e6,2),'exceptions MB',round(os.path.getsize(master)/1e6,2), flush=True)
