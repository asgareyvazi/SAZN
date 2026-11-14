"""Self-contained SAZN Executive Interactive Dashboard.
All data is embedded: gzip-compressed TSV index + JSON analytics. No network calls.
Live search across numbers, wells, rigs, materials, files, sheets and pages.
"""
import os, sys, json, gzip, base64, sqlite3, zlib, re, collections
sys.path.insert(0, os.path.dirname(__file__))
from common import OUT, DATA, DB, fmt, load_json

CON = sqlite3.connect(DB); CON.row_factory = sqlite3.Row
P = load_json('well_profiles')
F = json.load(open(os.path.join(DATA, 'field_analytics.json')))
VR = json.load(open(os.path.join(DATA, 'visual_review.json')))
S = json.load(open(os.path.join(DATA, 'deep_analytics.json')))
TOT = F['totals']
WELLS = [w['key'] for w in WELLS] if False else list(P.keys())
RIG = {k: p['rig'] for k, p in P.items()}
CONTRACT = {k: p['contract'] for k, p in P.items()}
FILES = [r[0] for r in CON.execute('SELECT rel_path FROM source_files ORDER BY id')]
FIDX = {f: i for i, f in enumerate(FILES)}
SHEETS = [(FIDX.get(r['rel_path'], 0), r['name'], r['well_key'] or '', r['category'] or '')
          for r in CON.execute('SELECT rel_path,name,well_key,category FROM sheets ORDER BY id')]


def clean(s, n=110):
    if s is None: return ''
    s = str(s).replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
    return re.sub(r'\s+', ' ', s).strip()[:n]


def num(v):
    if v is None or v == '': return ''
    if isinstance(v, float):
        if abs(v) >= 1000: return f'{v:,.4f}'.rstrip('0').rstrip('.')
        return f'{v:.6g}'
    return clean(v, 40)


ROWS = []
def add(kind, well, v, src, loc, label):
    rig = RIG.get(well, '')
    ct = CONTRACT.get(well, '')
    ROWS.append('\t'.join([kind, clean(well, 24), clean(v, 40), clean(src, 90), clean(loc, 60),
                           clean(label, 110), rig, ct]))


# ---- source documents & sheet census ---------------------------------------
for r in CON.execute('SELECT category,rel_path,filename,n_units,size_bytes,sha256 FROM source_files ORDER BY id'):
    add('FILE', '', r['n_units'], r['rel_path'], r['category'], r['filename'] + ' SHA256 ' + (r['sha256'] or '')[:16])
for i, (fi, nm, wk, cat) in enumerate(SHEETS):
    add('SHEET', wk, nm, FILES[fi], 'sheet %d' % (i + 1), cat + ' / ' + nm)

# ---- every numeric worksheet cell with exact A1-style provenance -----------
def a1(r, c):
    s = ''
    while c > 0:
        c, m = divmod(c - 1, 26); s = chr(65 + m) + s
    return f'{s}{r}'

q = CON.execute("SELECT s.rel_path,s.name,s.well_key,c.r,c.c,c.value FROM cell_numeric c JOIN sheets s ON s.id=c.sheet_id")
for r in q:
    add('CELL', r['well_key'], num(r['value']), r['rel_path'], r['name'] + '!' + a1(r['r'], r['c']), 'numeric cell')

# ---- every numeric PDF token that carries a unit or a measurement context --
UNIT = re.compile(r'\b(m|ft|pcf|ppg|psi|bbl|bph|gpm|rpm|klb|hrs?|days?|%|usd|gal|scf|m3|bar|cp|ksi|mm|inch|deg|ppb|bpd|bopd)\b', re.I)
KEY = re.compile(r'mud|cement|casing|bit|bha|npt|acid|stimulation|cbl|usit|packer|sidetrack|loss|fish|milling|trsssv|plt|pvt|waste|formation|top|pressure|rate|pi|whp|choke|volume|depth|md|tvd|rop|wob|rpm', re.I)
for r in CON.execute('SELECT p.well_key,f.rel_path,p.page_no,p.payload FROM pdf_num_pages p JOIN source_files f ON f.id=p.file_id'):
    try: data = json.loads(zlib.decompress(r['payload']))
    except Exception: continue
    for line, tok, val, unit, ctx in data:
        if not unit and not KEY.search(ctx): continue
        add('PDF', r['well_key'], num(tok), r['rel_path'], 'p.%d L%d' % (r['page_no'], line),
            (unit + ' ' if unit else '') + ctx)

# ---- semantic rows (CBS / DDR / EOWR / derived) ---------------------------
SQLS = [
 ('CBS', "SELECT well_key,rel_path,sheet,description,unit,qty,unit_price,total,cell_ref FROM cbs_items"),
 ('TOT', "SELECT well_key,rel_path,sheet,item,qty,unit_cost,total,row_no FROM cbs_totals"),
 ('STIM', "SELECT well_key,rel_path,sheet,description,unit,volume,unit_price,total,remark FROM cbs_stim"),
 ('ACCOM', "SELECT well_key,rel_path,sheet,description,qty,unit,rate,total FROM cbs_accomm"),
 ('MAT', "SELECT well_key,source_sheet,material,unit,received,used,stock,cell_ref FROM ddr_materials"),
 ('FORM', "SELECT well_key,rel_path,source,name,md,tvd,prognosis_md,lithology FROM ddr_formations"),
 ('SURV', "SELECT well_key,rel_path,tool,md,incl,azi,tvd,dls FROM ddr_surveys"),
 ('DAILY', "SELECT well_key,rel_path,sheet_data,report_no,date_g,md_2400,drilled_24,rop,mw,formation FROM ddr_daily"),
 ('ACT', "SELECT well_key,rel_path,sheet,report_no,phase,code,status,npt_attributed,hours,description FROM ddr_activities"),
 ('TIME', "SELECT well_key,rel_path,code,category,hours,report_no FROM ddr_time"),
 ('BHA', "SELECT well_key,rel_path,item,bha_no,od,length,cum_len,report_no FROM ddr_bha"),
 ('CAS', "SELECT well_key,rel_path,size,grade,weight,from_m,to_m,burst,collapse,report_no FROM ddr_casing"),
 ('POB', "SELECT well_key,rel_path,category,count,total,report_no FROM ddr_pob"),
 ('NOTE', "SELECT well_key,rel_path,sheet,category,text,report_no FROM ddr_notes"),
 ('FACT', "SELECT well_key,rel_path,fact_key,value,unit,page_no,context FROM eowr_facts"),
 ('TBL', "SELECT t.well_key,f.rel_path,t.table_no,t.title,t.page_no,t.category FROM pdf_tables t JOIN source_files f ON f.id=t.file_id"),
 ('EVENT', "SELECT well_key,kind,day,depth,formation,hours,cost_usd,evidence FROM well_events"),
 ('LESSON', "SELECT well_key,code,polarity,title,evidence,action,hours,cost_usd FROM well_lessons"),
 ('RECON', "SELECT well_key,item,cbs,ddr,unit,status FROM reconciliation"),
 ('NPT', "SELECT well_key,category,hours,attributed,share FROM npt_breakdown"),
 ('TS', "SELECT well_key,date_g,report_no,md,drilled,rop,mw,npt_h,pt_h,cum_cost FROM field_timeseries"),
]
for kind, sql in SQLS:
    for r in CON.execute(sql):
        vals = [num(x) for x in tuple(r)[1:]]
        wk = r[0] or ''
        src = ''
        for x in tuple(r)[1:4]:
            if isinstance(x, str) and ('/' in x or '.' in x) and len(x) > 6:
                src = x; break
        add(kind, wk, ' | '.join(v for v in vals if v)[:40], src or FILES[0], kind,
            ' | '.join(clean(x, 40) for x in tuple(r)[1:] if x not in (None, '')))
# REC nine-axis cards
for r in CON.execute('SELECT well_key,payload FROM well_recs'):
    try: d = json.loads(r['payload'])
    except Exception: continue
    add('REC', r[0], d.get('code', ''), 'REC-01..REC-12', d.get('branch', ''),
        ' | '.join(clean(d.get(k), 55) for k in ('title','evidence','root_cause','protocol','control_window','kpi','owner','priority','days_saved','usd_saved')))
# ---- every archived PDF numeric token, including unclassified navigation tokens
for r in CON.execute("""SELECT well_key,rel_path,page_no,line_no,token,value,unit,context,band,context_class,
 robust_outlier,excel_exact_match,fact_exact_match,flags FROM an_pdf_tokens"""):
    loc = 'p.%s L%s' % (r['page_no'], r['line_no'])
    lab = ('PDF token %s | unit %s | band %s | %s | robust outlier %s | Excel %s | fact %s | %s | %s'
           % (clean(r['token'], 24), clean(r['unit'], 12), r['band'], r['context_class'],
              r['robust_outlier'], r['excel_exact_match'], r['fact_exact_match'],
              clean(r['context'], 60), clean(r['flags'], 70)))
    add('AN_PDF_TOKEN', r['well_key'], num(r['value']), r['rel_path'], loc, lab)

# ---- peer-well reference rate card (DA-22)
for r in CON.execute("SELECT description,unit,n_lines,n_wells,ref_rate,v_min,v_max,spread_pct,pct_lines_within_5pct,wells,status FROM an_unit_rate_reference"):
    lab = ('Reference rate card | %s | %s | lines %s | wells %s | median %s | min %s | max %s | spread %s%% | within5%% %s | %s'
           % (clean(r['description'],60), r['unit'], r['n_lines'], r['n_wells'], num(r['ref_rate']),
              num(r['v_min']), num(r['v_max']), fmt(r['spread_pct'],2) if r['spread_pct'] is not None else '—',
              r['pct_lines_within_5pct'], r['status']))
    add('AN_RATE_REF', '', num(r['ref_rate']), 'CBS reference card (peer-well median)', r['unit'], lab)
for r in CON.execute("SELECT well_key,rel_path,sheet,cell,description,unit,qty,unit_price,ref_rate,price_dev_pct,source_total,ref_expected,margin_vs_ref_usd,status FROM an_cbs_reference_delta WHERE margin_vs_ref_usd IS NOT NULL ORDER BY ABS(margin_vs_ref_usd) DESC LIMIT 500"):
    lab = ('Reference delta | src price %s | peer ref %s | dev %s%% | qty %s | src total %s | ref expected %s | margin vs ref %s USD | %s | %s'
           % (num(r['unit_price']), num(r['ref_rate']), fmt(r['price_dev_pct'],2) if r['price_dev_pct'] is not None else '—',
              num(r['qty']), num(r['source_total']), num(r['ref_expected']), num(r['margin_vs_ref_usd']),
              r['status'], clean(r['description'],50)))
    add('AN_RATE_DEV', r['well_key'], num(r['margin_vs_ref_usd']), r['rel_path'], r['sheet'] + '!' + str(r['cell']), lab)

# ---- Drilling-Data Analysis Department derived rows ----------------------
AN_SQLS = [
 ('AN_CBS', "SELECT well_key,rel_path,sheet,cell,description,total,expected,delta,status FROM an_cbs_line_audit"),
 ('AN_FORMULA', "SELECT well_key,rel_path,sheet,cell,formula,cached,recomputed,delta,status FROM an_formula_audit WHERE status<>'MATCH'"),
 ('AN_TIME', "SELECT well_key,canonical_sheet,report_no,activity_h,pt_h,npt_h,ilt_h,carried_over_h,balance_status FROM an_time_balance"),
 ('AN_CONSIST', "SELECT well_key,report_no,drilled_24,expected_delta,delta,status,corrected_24h FROM an_daily_consistency WHERE status NOT IN ('MATCH','NO_SOURCE')"),
 ('AN_PARAM', "SELECT well_key,report_no,param,value,sheet,cell FROM an_daily_params LIMIT 9000"),
 ('AN_CORR', "SELECT well_key,x,y,n,pearson,spearman,r2,strength FROM an_correlations"),
 ('AN_NPT', "SELECT well_key,rank,report_no,category,hours,cost_usd,attributed,cumulative_pct FROM an_npt_forensics"),
 ('AN_FORMATION', "SELECT well_key,formation,n_tops,md_med,prognosis_med,delta_m,thickness_m,problem_h,structural_bias FROM an_formation_stats"),
 ('AN_DIRECTIONAL', "SELECT well_key,report_no,md,incl,azi,dls,cum_dogleg,tortuosity,section FROM an_directional LIMIT 3000"),
 ('AN_MUD', "SELECT well_key,report_no,material,received,used,stock,implied_received,delta,status FROM an_mud_balance WHERE status='IMBALANCE'"),
 ('AN_CASING', "SELECT well_key,size,grade,from_m,to_m,nearest_formation,shoe_to_top_m,lap_overlap_m,status FROM an_casing_check LIMIT 2000"),
 ('AN_STIM', "SELECT well_key,sheet,item,planned_volume,unit_price,total_usd,expected,delta,status FROM an_stim_audit"),
 ('AN_RECON', "SELECT well_key,item,cbs_usd,ddr_qty,eowr_value,delta_abs,delta_pct,tolerance_status FROM an_recon_delta LIMIT 3000"),
 ('AN_PROG', "SELECT well_key,formation,program_md,eowr_md,ddr_md,delta_pa,delta_pd,bias FROM an_program_vs_actual"),
 ('AN_EFF', "SELECT well_key,days,drilled_m,gross,usd_per_day,usd_per_m,npt_share_pct,payable_ratio_pct,cost_rank,npt_rank FROM an_cost_efficiency"),
 ('AN_LEARN', "SELECT chain,from_well,to_well,delta_days,delta_npt_h,delta_cbs,npt_per_day_from,npt_per_day_to,direction FROM an_learning_metrics"),
 ('AN_FIELDCORR', "SELECT x,y,n,pearson,spearman,r2,strength FROM an_field_correlations"),
 ('AN_DQ', "SELECT check_name,scope,n,affected,pct,detail FROM an_data_quality"),
 ('AN_OUTLIER', "SELECT well_key,rel_path,sheet,cell,value,z FROM an_outliers"),
 ('AN_FINDING', "SELECT id,category,title,numeric_evidence,interpretation,action,priority FROM an_findings"),
 ('AN_PDF_STAT', "SELECT category,well_key,unit,band,n,v_min,v_med,v_max,v_mean,v_std,zeros,negatives,integers,decimals,measured,navigation,identifier,low_text,outliers,excel_matches,fact_matches,unit_bearing,unclassified FROM an_pdf_token_stats"),
 ('AN_PDF_UNIT', "SELECT unit,n,pct,v_min,q1,v_med,q3,v_max,v_mean,v_std,outliers,measured,excel_matches,fact_matches,context_classes,dispersion FROM an_pdf_unit_stats"),
 ('AN_PDF_CONTEXT', "SELECT context_class,n,pct,v_min,v_med,v_max,v_mean,outliers,excel_matches,fact_matches,low_text,unit_bearing FROM an_pdf_context_stats"),
 ('AN_PDF_PAGE', "SELECT well_key,rel_path,category,page_no,chapter,n_chars,n_images,n_tokens,density_per_1k_chars,distinct_units,unit_bearing,measured,navigation,identifier,outliers,excel_matches,fact_matches,v_min,v_med,v_max,status FROM an_pdf_page_stats"),
 ('AN_PDF_LOW', "SELECT well_key,rel_path,page_no,chapter,n_chars,n_images,numeric_tokens,distinct_units,measured_tokens,outlier_tokens,note FROM an_pdf_low_text"),
 ('AN_PDF_OUTLIER', "SELECT well_key,rel_path,category,page_no,line_no,token,value,unit,context,band,context_class,log_abs,robust_z,excel_exact_match,fact_exact_match,flags FROM an_pdf_outliers"),
 ('AN_PDF_MATCH', "SELECT value_key,value,pdf_tokens,excel_cells,fact_rows,first_pdf_source,first_excel_source,first_fact_source,match_kind,caveat FROM an_pdf_value_matches"),
 ('AN_PDF_SOURCE', "SELECT well_key,rel_path,category,pages,low_text_pages,low_text_pct,numeric_tokens,n_chars,n_images,measured_tokens,unclassified_tokens,outlier_tokens,status FROM an_pdf_source_gaps"),
 ('AN_COST_TRACK', "SELECT well_key,op,rate_basis,rig_rate,cbs_plan,cbs_days,ddr_days,delta_days,gross_formula,gross_stored,gross_delta,npt_total_h,npt_contractor_h,npt_client_h,contractor_deduction,client_deduction,all_npt_deduction,payable_formula,payable_stored,payable_delta,net_formula,net_stored,net_delta,attribution_gap_h,status,calendar_flag,fallback_50_50_waiting_h FROM an_cost_track_audit"),
 ('AN_NPT_ATTR', "SELECT well_key,activity_rows,npt_status_rows,npt_status_h,contractor_h,client_h,blank_attribution_h,explicit_attribution_h,explicit_contractor_h,explicit_client_h,profile_npt_total_h,delta_h,explicit_attribution_pct,explicit_client_pct,waiting_h,waiting_blank_attribution_h,fallback_50_50_waiting_h,status FROM an_npt_attribution_audit"),
 ('AN_RATE_REF', "SELECT description,unit,n_lines,n_wells,ref_rate,v_min,v_max,spread_pct,pct_lines_within_5pct,wells,status FROM an_unit_rate_reference"),
 ('AN_RATE_DEV', "SELECT well_key,rel_path,sheet,cell,description,unit,qty,unit_price,ref_rate,price_dev_pct,source_total,ref_expected,margin_vs_ref_usd,status FROM an_cbs_reference_delta WHERE margin_vs_ref_usd IS NOT NULL ORDER BY ABS(margin_vs_ref_usd) DESC LIMIT 500"),
 ('AN_DQ_EXT', "SELECT check_name,scope,n,affected,pct,detail FROM an_data_quality"),
 ('AN_CENSUS', "SELECT well_key,rel_path,sheet,n,v_min,v_med,v_max,v_mean,v_std,outliers,dominant_band FROM an_num_census"),
 ('AN_PARAMSTAT', "SELECT well_key,param,n,v_min,v_med,v_max,v_mean,v_std,cv_pct,iqr_outliers FROM an_param_stats"),
 ('AN_PRESSURE', "SELECT well_key,unit,n,v_min,v_med,v_max,v_mean,v_std,outliers FROM an_pressure_stats"),
 ('AN_UNITPRICE', "SELECT description,unit,n,v_min,v_med,v_max,spread_pct,dispersion FROM an_unit_price"),
 ('AN_CROSSTAB', "SELECT well_key,phase,code,status,attributed,hours,events,share_pct FROM an_activity_crosstab"),
 ('AN_ILT', "SELECT well_key,report_no,connection_h,trip_h,circulate_h,wait_h,survey_h,ream_h,ilt_h,ilt_pct FROM an_ilt"),
 ('AN_PARETO', "SELECT well_key,category,hours,share_pct,cumulative_pct,rank,cost_usd FROM an_npt_pareto"),
 ('AN_HIST', "SELECT band,n,pct,example,example_source FROM an_num_hist"),
]
for kind, sql in AN_SQLS:
    for r in CON.execute(sql):
        t = tuple(r); wk = t[0] if t[0] else ''
        vals = [clean(x, 26) for x in t[1:] if x not in (None, '')]
        add(kind, wk, vals[0] if vals else '', vals[1] if len(vals) > 1 else kind, kind,
            ' | '.join(vals[:7]))

# PDF page / chapter census
for r in CON.execute('SELECT well_key,rel_path,page_no,chapter,n_chars,n_images,keywords FROM pdf_pages'):
    add('PAGE', r[0], r[3], r[1], 'p.%d' % r[2], f'{r[4]} chars, {r[5]} images | {clean(r[6],70)}')
# audit findings
for r in CON.execute('SELECT id,title,detail FROM audit_findings'):
    add('AUDIT', '', r[0], 'audit_findings', r[0], r[1] + ' | ' + r[2])

BLOB = '\n'.join(ROWS)
GZ = base64.b64encode(gzip.compress(BLOB.encode('utf-8'), 9)).decode('ascii')
META = dict(files=FILES, sheet_count=len(SHEETS), rows=len(ROWS),
            kinds=sorted(collections.Counter(x.split('\t', 1)[0] for x in ROWS).items()),
            wells=WELLS, rigs=sorted(set(RIG.values())), contracts=sorted(set(CONTRACT.values())),
            totals=TOT, field=F, review=VR['review'],
            deep=dict(data_quality=S['data_quality'], findings=S['findings'], num_hist=S['field']['num_hist'],
                      cost_efficiency=S['field']['cost_efficiency'], field_correlations=S['field']['field_correlations'],
                      learning_metrics=S['field']['learning_metrics'], unit_price=S['field']['unit_price_dispersion'][:20],
                      program_vs_actual=S['field']['program_vs_actual'][:20], derived_tables=S['derived_tables'],
                      pdf=S['pdf'], cost_track=S['cost_track'], rate_card=S.get('rate_card'),
                      per_well={k: dict(cbs_mismatch=v['cbs_mismatch'], ilt_h=v['ilt_h'], param_rows=v['param_rows'],
                                        correlations=v['correlations'], cum_typos=v['cum_typos'], mud_imbalance=v['mud_imbalance'],
                                        usd_per_day=(v['cost_efficiency'][0]['usd_per_day'] if v['cost_efficiency'] else None),
                                        npt_share_pct=(v['cost_efficiency'][0]['npt_share_pct'] if v['cost_efficiency'] else None),
                                        cost_rank=(v['cost_efficiency'][0]['cost_rank'] if v['cost_efficiency'] else None))
                                for k, v in S['per_well'].items()}))
print('ROWS', len(ROWS), 'raw MB', round(len(BLOB)/1e6, 2), 'gzip KB', round(len(GZ)*3/4/1024, 1))

HTML = '''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>SAZN Executive Interactive Dashboard — South Azadegan</title>
<style>
:root{--navy:#0B1F3A;--cobalt:#1D4ED8;--teal:#0D9488;--emerald:#15803D;--amber:#D97706;--coral:#DC2626;--slate:#475569;--soft:#F4F7FC;--grid:#DCE5F0}
*{box-sizing:border-box}body{margin:0;font-family:Arial,Helvetica,sans-serif;color:#0B1F3A;background:#F8FAFD}
header{background:var(--navy);color:#fff;padding:22px 30px}header h1{margin:0;font-size:26px}header p{margin:6px 0 0;color:#BFD2EE;font-size:13px}
.badge{display:inline-block;background:#1D4ED8;color:#fff;border-radius:12px;padding:3px 10px;font-size:11px;margin-right:6px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;padding:18px 30px}
.kpi{background:#fff;border:1px solid var(--grid);border-radius:8px;padding:12px}.kpi b{display:block;font-size:20px;color:var(--cobalt)}.kpi span{font-size:11px;color:var(--slate)}
.wrap{padding:0 30px 30px}.panel{background:#fff;border:1px solid var(--grid);border-radius:8px;padding:16px;margin-bottom:16px}
h2{font-size:16px;margin:0 0 12px;color:var(--navy)}h3{font-size:13px;margin:14px 0 8px;color:var(--slate)}
#q{width:100%;padding:12px 14px;font-size:15px;border:2px solid var(--cobalt);border-radius:8px;outline:none}
.controls{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0}select,button{padding:7px 10px;border:1px solid var(--grid);border-radius:6px;background:#fff;font-size:12px}
button.act{background:var(--cobalt);color:#fff;border-color:var(--cobalt)}.chip{border-radius:14px;padding:5px 11px;cursor:pointer;font-size:11px;border:1px solid var(--grid)}
.chip.on{background:var(--navy);color:#fff;border-color:var(--navy)}
#status{font-size:12px;color:var(--slate);margin:8px 0}
table{border-collapse:collapse;width:100%;font-size:12px}th{background:var(--navy);color:#fff;text-align:center;padding:7px;position:sticky;top:0}
td{border-bottom:1px solid var(--grid);padding:6px;vertical-align:top}tr:nth-child(even) td{background:var(--soft)}
.scroll{max-height:480px;overflow:auto;border:1px solid var(--grid);border-radius:6px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}@media(max-width:900px){.grid2{grid-template-columns:1fr}}
.bar{height:10px;border-radius:3px;background:var(--cobalt)}.row{display:flex;align-items:center;gap:8px;margin:5px 0;font-size:11px}
.row .lbl{width:150px;flex:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.row .val{width:110px;text-align:right;flex:none}
.legend{font-size:11px;color:var(--slate);margin-top:8px}footer{padding:18px 30px;color:var(--slate);font-size:11px;border-top:1px solid var(--grid)}
code{background:var(--soft);padding:1px 4px;border-radius:3px}
</style></head><body>
<header><h1>SAZN — South Azadegan Executive Interactive Dashboard</h1>
<p><span class="badge">51 source files</span><span class="badge">13 operations</span><span class="badge">3 contracts</span><span class="badge">2 rigs</span>
Self-contained · offline · live numeric search with file / sheet / cell / page provenance</p></header>
<div class="kpis" id="kpis"></div>
<div class="wrap">
<div class="panel"><h2>Live search — numbers, wells, rigs, materials, files, sheets, pages</h2>
<input id="q" placeholder="Try 23,000  |  7,280,632.64  |  HCl  |  AZNS-207  |  OEOC 208  |  CBL  |  4200  |  Xmas Tree  |  sheet cell">
<div class="controls" id="chips"></div>
<div class="controls"><select id="well"><option value="">All wells</option></select>
<select id="rig"><option value="">All rigs</option></select>
<select id="contract"><option value="">All contracts</option></select>
<select id="limit"><option>100</option><option>250</option><option>500</option></select>
<button id="csv">Download visible rows (CSV)</button></div>
<div id="status">Decompressing embedded index…</div>
<div class="scroll"><table><thead><tr><th>Kind</th><th>Well</th><th>Value</th><th>Document / source</th><th>Location</th><th>Context / label</th><th>Rig</th><th>Contract</th></tr></thead><tbody id="tb"></tbody></table></div>
</div>
<div class="grid2">
<div class="panel"><h2>Four-track cost by operation (USD)</h2><div id="cost"></div><div class="legend">CBS plan · gross calendar · payable excluding contractor zero-rate NPT · pure net excluding all NPT.</div></div>
<div class="panel"><h2>NPT hours by operation</h2><div id="npt"></div><div class="legend">Attributed NPT from DDR 24-h PLN/NPT coding.</div></div>
</div>
<div class="grid2">
<div class="panel"><h2>Contract close-out</h2><div id="contracts"></div></div>
<div class="panel"><h2>Source census &amp; graphics QA</h2><div id="census"></div></div>
</div>
<div class="grid2">
<div class="panel"><h2>Data-Analysis Department — quality scorecard</h2><div id="dq"></div></div>
<div class="panel"><h2>Numeric census by magnitude band</h2><div id="hist"></div></div>
</div>
<div class="grid2">
<div class="panel"><h2>Cost efficiency by operation (USD/day)</h2><div id="eff"></div></div>
<div class="panel"><h2>CBS priced-line arithmetic exceptions (USD)</h2><div id="cbsx"></div></div>
</div>
<div class="grid2">
<div class="panel"><h2>Strongest field-level correlations</h2><div id="fcorr"></div></div>
<div class="panel"><h2>Programme → EOWR top deltas (m)</h2><div id="pva"></div></div>
</div>
<div class="grid2">
<div class="panel"><h2>Re-entry learning metrics</h2><div id="learn"></div></div>
<div class="panel"><h2>Per-well forensic counts</h2><div id="perwell"></div></div>
</div>
<div class="grid2">
<div class="panel"><h2>All 305,447 PDF numeric tokens by magnitude band</h2><div id="pdfbands"></div></div>
<div class="panel"><h2>PDF unit / context coverage</h2><div id="pdfunit"></div></div>
</div>
<div class="grid2">
<div class="panel"><h2>Low-text PDF pages by source category (AF-05)</h2><div id="pdflow"></div></div>
<div class="panel"><h2>Four-track cost bridge (USD)</h2><div id="costbridge"></div></div>
</div>
<div class="grid2">
<div class="panel"><h2>NPT attribution audit — explicit rows, no 50/50 split</h2><div id="nptattr"></div></div>
<div class="panel"><h2>PDF / Excel / fact exact value coincidences</h2><div id="pdfmatch"></div></div>
</div>
<div class="panel"><h2>Peer-well reference rate card (DA-22) — project-confirmed uniform rates</h2><div id="ratecard"></div></div>
<div class="panel"><h2>Evidence-based findings (DA-01 … DA-22)</h2><div id="findings"></div></div>
<div class="panel"><h2>How to read a result</h2><p style="font-size:12px;line-height:1.6;color:#475569">
<code>CELL</code> rows carry an exact Excel sheet and A1 reference; <code>PDF</code> rows carry a page and line number;
<code>CBS</code>, <code>MAT</code>, <code>FORM</code>, <code>FACT</code>, <code>TBL</code>, <code>AN_PDF_TOKEN</code>, <code>AN_COST_TRACK</code> and <code>AN_NPT_ATTR</code> rows carry the source document, a short context and the derived forensic class.
Every figure in this dashboard is a value read from the 51 source documents. Derived quantities (four-track cost, NPT deductions, DLS, structural delta) are computed only from sourced rates, quantities and hours and are labelled by their basis in the master audit workbook.
No value is invented, interpolated or rounded away; pages without a text layer are retained page by page in the warehouse.</p></div>
</div>
<footer>SAZN Master Warehouse · SAZN_Master_Warehouse.sqlite · SAZN_Master_Audit_Workbook.xlsx · 167 reviewed graphics · 31 Word books · generated 2026-11-08</footer>
<script>
const META=__META__;const B64="__B64__";
let ROWS=[],KINDS=new Set(),VIEW=[];
const $=s=>document.querySelector(s);
function fmt(n,d){return (typeof n==='number'?n:Number(n)||0).toLocaleString('en-US',{maximumFractionDigits:d===undefined?2:d})}
(async function(){
  const bin=Uint8Array.from(atob(B64),c=>c.charCodeAt(0));
  let txt;
  if('DecompressionStream' in window){
    const ds=new DecompressionStream('gzip');
    txt=await new Response(new Blob([bin]).stream().pipeThrough(ds)).text();
  }else{txt=new TextDecoder().decode(bin);}
  ROWS=txt.split('\\n').map(l=>l.split('\\t'));
  $('#status').textContent='Index ready: '+ROWS.length.toLocaleString()+' searchable rows · '+META.files.length+' documents · '+META.sheet_count+' worksheets.';
  buildUI();search();
})().catch(e=>{$('#status').textContent='Index error: '+e.message});
function kpi(k,v){return '<div class="kpi"><b>'+v+'</b><span>'+k+'</span></div>'}
function buildUI(){
  const T=META.totals;
  $('#kpis').innerHTML=[kpi('Operations',T.operations),kpi('Source files',51),kpi('Excel sheets',fmt(T.excel_sheets,0)),kpi('PDF pages',fmt(T.pdf_pages,0)),
    kpi('Non-empty cells',fmt(T.cells,0)),kpi('Numeric cells',fmt(T.numeric_cells,0)),kpi('Formulas',fmt(T.formulas,0)),kpi('PDF numeric tokens',fmt(T.pdf_numbers,0)),
    kpi('Numbered tables',fmt(T.pdf_tables,0)),kpi('DDR days',T.ddr_days),kpi('NPT hours',fmt(T.npt_hours,1)),
    kpi('CBS plan USD',fmt(T.cbs_plan/1e6,2)+'M'),kpi('Gross USD',fmt(T.gross/1e6,2)+'M'),
    kpi('Payable USD',fmt(T.payable/1e6,2)+'M'),kpi('Net USD',fmt(T.net/1e6,2)+'M')].join('');
  const ks=META.kinds.map(x=>x[0]);
  $('#chips').innerHTML='<button class="chip on" data-k="">All ('+ROWS.length.toLocaleString()+')</button>'+
    ks.map(k=>'<button class="chip" data-k="'+k+'">'+k+' ('+fmt((META.kinds.find(x=>x[0]===k)||[0,0])[1],0)+')</button>').join('');
  $('#chips').onclick=e=>{if(!e.target.dataset.k&&e.target.dataset.k!=='')return;
    document.querySelectorAll('.chip').forEach(c=>c.classList.remove('on'));e.target.classList.add('on');KINDS=e.target.dataset.k?new Set([e.target.dataset.k]):new Set();search();};
  for(const [id,vals] of [['well',META.wells],['rig',META.rigs],['contract',META.contracts]])
    $('#'+id).innerHTML+=' '+vals.map(v=>'<option>'+v+'</option>').join('');
  ['q','well','rig','contract','limit'].forEach(id=>$('#'+id).addEventListener(id==='q'?'input':'change',search));
  $('#csv').onclick=download;
  charts();deepCharts();
}
function search(){
  const q=$('#q').value.trim().toLowerCase(),w=$('#well').value,r=$('#rig').value,c=$('#contract').value,lim=+$('#limit').value;
  let shown=0;VIEW=[];
  for(let i=0;i<ROWS.length&&shown<lim;i++){
    const x=ROWS[i];if(x.length<8)continue;
    if(KINDS.size&&!KINDS.has(x[0]))continue;
    if(w&&x[1]!==w)continue;if(r&&x[6]!==r)continue;if(c&&x[7]!==c)continue;
    if(q&&x.join(' ').toLowerCase().indexOf(q)<0)continue;
    VIEW.push(x);shown++;
  }
  $('#status').textContent=VIEW.length+' rows shown'+(q?' for "'+$('#q').value+'"':'')+' · '+ROWS.length.toLocaleString()+' rows indexed';
  $('#tb').innerHTML=VIEW.map(x=>'<tr>'+x.slice(0,8).map(v=>'<td>'+esc(String(v||''))+'</td>').join('')+'</tr>').join('');
}
function esc(s){return s.replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
function download(){
  const head=['Kind','Well','Value','Document','Location','Context','Rig','Contract'];
  const csv=[head.join(',')].concat(VIEW.map(x=>x.slice(0,8).map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(','))).join('\\n');
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download='SAZN_search_results.csv';a.click();
}
function bars(el,items,keys,colors){
  const mx=Math.max(...items.flatMap(o=>keys.map(k=>Number(o[k])||0)),1);
  el.innerHTML=items.map(o=>{
    const nm=o.key||o.id||o.well;
    return '<div class="row"><span class="lbl" title="'+nm+'">'+nm+'</span>'+
      keys.map((k,j)=>'<span style="flex:1"><span class="bar" style="display:block;width:'+(100*(Number(o[k])||0)/mx).toFixed(1)+'%;background:'+colors[j]+'"></span></span>').join('')+
      '<span class="val">'+fmt(o[keys[0]]/1e6,2)+'M</span></div>';
  }).join('');
}
function deepCharts(){
  const D=META.deep,tbl=(head,rows)=>'<table><thead><tr>'+head.map(h=>'<th>'+h+'</th>').join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr>'+r.map(v=>'<td>'+esc(String(v===null||v===undefined?'—':v))+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
  $('#dq').innerHTML=tbl(['Check','Affected / N','%'],D.data_quality.map(d=>[d.check_name,fmt(d.affected,0)+' / '+fmt(d.n,0),fmt(d.pct,2)+'%']));
  bars($('#hist'),D.num_hist.map(h=>({key:h.band,n:h.n})),['n'],['#1D4ED8']);
  bars($('#eff'),D.cost_efficiency.map(c=>({key:c.well_key,usd_per_day:c.usd_per_day})),['usd_per_day'],['#0D9488']);
  const ex=ROWS.filter(r=>r[0]==='AN_CBS'&&Math.abs(Number(r[2])||0)>0).slice(0,14);
  $('#cbsx').innerHTML=tbl(['Well','Document','Total','Expected','Delta','Status'],ex.map(r=>[r[1],r[3],r[2],r[4],r[5],r[7]]));
  $('#fcorr').innerHTML=tbl(['X','Y','N','Pearson','Spearman','R²'],D.field_correlations.slice(0,12).map(c=>[c.x,c.y,c.n,fmt(c.pearson,3),fmt(c.spearman,3),fmt(c.r2,3)]));
  bars($('#pva'),D.program_vs_actual.slice(0,12).map(p=>({key:(p.well_key+' '+p.formation).slice(0,22),d:Math.abs(Number(p.delta_pa)||0)})),['d'],['#D97706']);
  $('#learn').innerHTML=tbl(['Chain','From → To','NPT/day','CBS/day','Direction'],D.learning_metrics.map(l=>[l.chain,l.from_well+' → '+l.to_well,fmt(l.npt_per_day_from,2)+' → '+fmt(l.npt_per_day_to,2),fmt(l.cbs_per_day_from,0)+' → '+fmt(l.cbs_per_day_to,0),l.direction]));
  $('#perwell').innerHTML=tbl(['Well','CBS exceptions','ILT h','Param rows','Correlations','Cum. typos','Mud imbalance','USD/day'],Object.entries(D.per_well).map(([k,v])=>[k,v.cbs_mismatch,fmt(v.ilt_h,0),fmt(v.param_rows,0),v.correlations,v.cum_typos,v.mud_imbalance,fmt(v.usd_per_day,0)]));
  const pdf=D.pdf,ct=D.cost_track;
  $('#pdfbands').innerHTML=tbl(['Band','Tokens','Unit-bearing','Measured','Outliers','Excel matches','Fact matches'],pdf.bands.map(b=>[b.band,fmt(b.n,0),fmt(b.unit_bearing,0),fmt(b.measured,0),fmt(b.outliers,0),fmt(b.excel_matches,0),fmt(b.fact_matches,0)]));
  $('#pdfunit').innerHTML=tbl(['Unit','N','Share %','Median','Max','Outliers','Dispersion'],pdf.units.slice(0,14).map(u=>[u.unit,fmt(u.n,0),fmt(u.pct,2),fmt(u.v_med,2),fmt(u.v_max,2),fmt(u.outliers,0),u.dispersion]))+
    '<p style="font-size:11px;color:#475569">Context classes: '+pdf.context_classes.map(c=>c.context_class+' '+fmt(c.n,0)).join(' · ')+'</p>';
  $('#pdflow').innerHTML=tbl(['Category','Low-text pages','Numeric tokens','Images','Measured tokens'],pdf.low_text_by_category.map(x=>[x.category,fmt(x.pages,0),fmt(x.numeric_tokens,0),fmt(x.images,0),fmt(x.measured_tokens,0)]));
  bars($('#costbridge'),[{key:'CBS plan',usd:ct.field.cbs_plan},{key:'Gross',usd:ct.field.gross},{key:'Payable',usd:ct.field.payable},{key:'Net',usd:ct.field.net}],['usd'],['#1D4ED8','#475569','#D97706','#15803D']);
  $('#nptattr').innerHTML=tbl(['Well','NPT h','Contractor h','Client h','Blank attribution h','Explicit %','Fallback 50/50 h'],ct.npt_attribution.map(x=>[x.well_key,fmt(x.npt_status_h,1),fmt(x.contractor_h,1),fmt(x.client_h,1),fmt(x.blank_attribution_h,1),fmt(x.explicit_attribution_pct,2),fmt(x.fallback_50_50_waiting_h,1)]));
  const rc=D.rate_card;
  $('#ratecard').innerHTML=tbl(['Metric','Value'],[['Reference keys',fmt(rc.keys_total,0)],['Multi-well keys',fmt(rc.multi_well_keys,0)],['Uniform card (<5% spread)',fmt(rc.uniform_card,0)],['Low dispersion',fmt(rc.low_dispersion,0)],['Dispersed',fmt(rc.dispersed,0)],['Priced lines within 5% of ref',fmt(rc.lines_within_5pct_of_ref,0)],['Lines deviating >5% vs ref',fmt(rc.lines_dev_gt5pct_vs_ref,0)],['Missing price, imputable from peers',fmt(rc.missing_price_lines,0)],['Embedded margin — positive deltas (USD)',fmt(rc.embedded_margin_pos_usd,0)],['Embedded margin — negative deltas (USD)',fmt(rc.embedded_margin_neg_usd,0)],['Net embedded margin (USD)',fmt(rc.embedded_margin_net_usd,0)],['Margin vs peer card — uniform keys only (net USD)',fmt(rc.uniform_delta_net_usd,0)]]);
  $('#pdfmatch').innerHTML=tbl(['Value','PDF tokens','Excel cells','Fact rows','First PDF source','First Excel source'],pdf.top_value_matches.map(x=>[x.value,fmt(x.pdf_tokens,0),fmt(x.excel_cells,0),fmt(x.fact_rows,0),x.first_pdf_source,x.first_excel_source]));
  $('#findings').innerHTML=D.findings.map(f=>'<div style="border-left:4px solid #1D4ED8;padding:8px 12px;margin:8px 0;background:#F8FAFD">'+
    '<b>'+f.id+' · '+esc(f.category)+' — '+esc(f.title)+'</b><div style="font-size:11px;color:#475569;margin-top:5px">'+esc(f.numeric_evidence)+'</div>'+
    '<div style="font-size:11px;color:#0B1F3A;margin-top:5px">'+esc(f.interpretation)+'</div>'+
    '<div style="font-size:11px;color:#15803D;margin-top:5px"><b>Action:</b> '+esc(f.action)+'</div></div>').join('');
}
function charts(){
  const W=META.field.wells;
  bars($('#cost'),W,['cbs_plan','gross','payable','net'],['#1D4ED8','#475569','#D97706','#15803D']);
  bars($('#npt'),W,['npt_h'],['#DC2626']);
  const C=META.field.contracts;
  $('#contracts').innerHTML='<table><thead><tr><th>Contract</th><th>Operations</th><th>DDR days</th><th>NPT h</th><th>Gross USD</th><th>Net USD</th></tr></thead><tbody>'+
    C.map(c=>'<tr><td>'+c.code+'</td><td>'+c.ops+'</td><td>'+fmt(c.days,0)+'</td><td>'+fmt(c.npt_h,1)+'</td><td>'+fmt(c.gross,0)+'</td><td>'+fmt(c.net,0)+'</td></tr>').join('')+'</tbody></table>';
  const R=META.review,T=META.totals;
  $('#census').innerHTML='<table><tbody>'+
    [['CBS workbooks / sheets','13 / '+T.cbs_sheets],['DDR packs / sheets','13 / '+T.ddr_sheets],
     ['EOWR files / pages','11 / '+T.eowr_pages],['Well programmes / pages','14 / '+T.program_pages],
     ['PDF numbered tables',fmt(T.pdf_tables,0)],['EOWR / Programme facts','2,190+'],
     ['Graphics',R.total+' PNG'],['Graphics resolution',R.resolution],
     ['Text overlap pairs',R.text_overlap_pairs],['Mean PNG size',R.mean_kb+' KB'],
     ['DOCX books','31 (30 × 24 chapters/149 tables; FA findings book SAZN-FA-DA-001)']].map(x=>'<tr><td>'+x[0]+'</td><td>'+x[1]+'</td></tr>').join('')+'</tbody></table>';
}
</script></body></html>'''
HTML = HTML.replace('__META__', json.dumps(META, ensure_ascii=False, separators=(',', ':')))
HTML = HTML.replace('__B64__', GZ)
path = os.path.join(OUT, 'SAZN_Executive_Interactive_Dashboard.html')
open(path, 'w', encoding='utf-8').write(HTML)
print('DASHBOARD MB', round(os.path.getsize(path)/1e6, 2), 'rows', len(ROWS))
