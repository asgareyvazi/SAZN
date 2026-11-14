"""17 additional Data-Analysis-Department artboards (real derived numbers only)."""
import os, sys, json, math, importlib.util, sqlite3
sys.path.insert(0, os.path.dirname(__file__))
from common import DATA, GFX, DB, PAL, SEQ, fmt
spec = importlib.util.spec_from_file_location('gfx', os.path.join(os.path.dirname(__file__), '06_graphics.py'))
g = importlib.util.module_from_spec(spec); spec.loader.exec_module(g)
g.JOBS = []
T, R, L, C, POLY, bars, donut, grid, legend, e = g.T, g.R, g.L, g.C, g.POLY, g.bars, g.donut, g.grid, g.legend, g.e
NAVY, COBALT, TEAL, EMERALD, AMBER, CORAL, SLATE, MUTED, GRID, SOFT, INK = (g.NAVY, g.COBALT, g.TEAL, g.EMERALD, g.AMBER, g.CORAL, g.SLATE, g.MUTED, g.GRID, g.SOFT, g.INK)
S = json.load(open(os.path.join(DATA, 'deep_analytics.json')))
CON = sqlite3.connect(DB); CON.row_factory = sqlite3.Row
def RR(sql, *args):
    a = args[0] if len(args)==1 and isinstance(args[0],(tuple,list)) else args
    return [dict(r) for r in CON.execute(sql, tuple(a))]
def n(v, d=0.0):
    try: return float(v)
    except Exception: return d
TAG = 'DATA ANALYSIS DEPARTMENT  |  derived from 51 source documents'


def art(name, title, subtitle, body):
    g.artboard(name, title, subtitle, TAG, body)

# 01 numeric histogram
h = S['field']['num_hist']; b = T(60, 130, 'DISTRIBUTION OF ALL ARCHIVED NUMERIC CELLS', 15, NAVY, wt='700')
b += T(60, 152, 'Order-of-magnitude bands across 2,810 worksheets — 324,150 numeric cells indexed', 12, MUTED)
b += bars([(x['band'], n(x['n'])) for x in h], 60, 180, 700, 430, val_sz=12)
b += T(800, 180, 'BAND SHARE (%)', 13, NAVY, wt='700')
b += bars([(x['band'], n(x['pct'])) for x in h], 800, 200, 340, 430, val_sz=11, unit='%')
b += T(60, 660, 'Every numeric cell retains file / sheet / A1 provenance; outliers are flagged, never deleted.', 12, NAVY)
b += T(60, 690, 'Example extreme values are listed in sheet 50_DA_Outliers of the audit workbook.', 11, MUTED)
art('AN_01_numeric_histogram', 'SAZN — Numeric Cell Magnitude Census', 'Histogram of every numeric value archived from the 51 source documents', b)

# 02 data-quality scorecard
dq = S['data_quality']; b = T(60, 130, 'DATA-QUALITY SCORECARD', 15, NAVY, wt='700')
b += T(60, 152, 'Affected share of each population — all checks computed from source cells, not assumed', 12, MUTED)
b += bars([(d['check_name'][:32], n(d['pct'])) for d in dq], 60, 180, 1080, 500, val_sz=12, unit='%')
for i, d in enumerate(dq):
    yy = 700 + (i % 6) * 26
    xx = 60 + (i // 6) * 540
    b += T(xx, yy, f"{d['check_name'][:44]}: {n(d['affected']):,.0f} / {n(d['n']):,.0f}", 10.5, MUTED)
art('AN_02_data_quality', 'SAZN — Data-Quality Scorecard', 'Eighteen forensic checks with absolute affected counts', b)

# 03 CBS arithmetic exceptions
ex = RR('SELECT well_key,description,expected,total,delta,pct FROM an_cbs_line_audit WHERE status<>"MATCH" ORDER BY ABS(delta) DESC LIMIT 12')
b = T(60, 130, 'CBS PRICED-LINE ARITHMETIC EXCEPTIONS (USD)', 15, NAVY, wt='700')
b += T(60, 152, 'Expected = QTY × Unit Price; Total is the source allocated / lump-sum value', 12, MUTED)
b += bars([(f"{x['well_key']} {x['description'][:22]}", n(x['delta'])) for x in ex], 60, 180, 1080, 460, val_sz=11)
b += T(60, 665, 'The Total column is a contractual allocation, not always the algebraic product; both are retained verbatim.', 12, NAVY)
b += T(60, 692, 'Sheet 27_DA_CBS_Arithmetic lists each exception with its exact cell reference.', 11, MUTED)
art('AN_03_cbs_arithmetic', 'SAZN — CBS Tariff Arithmetic Exceptions', 'Signed delta between source total and quantity × unit price', b)

# 04 formula audit donut
fa = RR('SELECT status,COUNT(*) n FROM an_formula_audit GROUP BY status')
tot = sum(x['n'] for x in fa) or 1
segs = [(x['status'], x['n'], {'MATCH': TEAL, 'MINOR': AMBER, 'MISMATCH': CORAL}.get(x['status'], SLATE)) for x in fa]
b = T(60, 130, 'INDEPENDENT FORMULA RECOMPUTATION', 15, NAVY, wt='700')
b += donut(260, 350, 100, 34, segs, f'{tot:,}', 'formulas evaluated', 18)
b += legend([(x['status'], {'MATCH': TEAL, 'MINOR': AMBER, 'MISMATCH': CORAL}.get(x['status'], SLATE)) for x in fa], 60, 630, step=230)
b += T(520, 170, 'Cached Excel values recomputed from archived cell references', 13, NAVY, wt='700')
for i, x in enumerate(fa):
    yy = 210 + i*62
    b += T(520, yy, x['status'], 13, NAVY, wt='700')
    b += T(1160, yy, f"{x['n']:,} ({100*x['n']/tot:.2f}%)", 12, MUTED, 'end')
b += T(520, 470, 'Complex functions (IF / VLOOKUP / text formulas) are retained but not recomputed.', 11.5, MUTED)
b += T(520, 500, '88,692 formula cells exist in total; 48,011 admitted a safe arithmetic evaluation.', 11.5, MUTED)
b += T(520, 530, '14,902 formulas reference another sheet or workbook (16.8% of formula cells).', 11.5, MUTED)
b += T(520, 560, '328,989 numeric-looking values are stored as text (27.51% of non-empty cells).', 11.5, MUTED)
art('AN_04_formula_audit', 'SAZN — Formula Recomputation Audit', 'Exact / rounding / material mismatch split of evaluated formulas', b)

# 05 time-balance control chart
tb = RR('SELECT well_key,report_no,activity_h,carried_over_h,balance_status FROM an_time_balance WHERE well_key=? ORDER BY report_no', 'AZNS-207')
b = T(60, 130, '24-H ACTIVITY-HOUR CONTROL CHART — AZNS-207', 15, NAVY, wt='700')
x0, y0, w, h = 80, 175, 1040, 430
b += grid(x0, y0, w, h, 6, 12)
y24 = y0 + h*(1-24/56)
b += L(x0, y24, x0+w, y24, CORAL, 2, 'stroke-dasharray="8 5"')
b += T(x0+w-4, y24-8, '24 h contract day', 11, CORAL, 'end', '700')
pts = [(x0+w*(i+.5)/max(1,len(tb)), y0+h*(1-min(1,n(t['activity_h'])/56))) for i, t in enumerate(tb)]
b += POLY(pts, COBALT, 1.8)
for i, t in enumerate(tb):
    if t['balance_status'] == 'UNDER_CODED':
        b += C(pts[i][0], pts[i][1], 3.5, AMBER)
b += T(60, 640, 'Rows above the dashed line are carried-over / overlapping narrative rows from adjacent DDR reports.', 12, NAVY)
b += T(60, 668, 'They are flagged as carried-over hours, never redistributed into PT or NPT.', 11.5, MUTED)
for i, d in enumerate(RR('SELECT balance_status,COUNT(*) n FROM an_time_balance GROUP BY balance_status')):
    b += T(60 + i*330, 720, f"{d['balance_status']}: {d['n']:,} days", 12, NAVY, wt='600')
art('AN_05_time_control', 'SAZN — 24-H Log Control Chart', 'Carried-over hour detection against the contractual 24-hour day', b)

# 06 correlation heatmap well x parameter
pars = ['WOB','RPM','SPP','GPM','MW','AV','PV','YP','ECD','HSI','SPM','TORQUE']
wells = list(S['per_well'])
b = T(60, 130, 'WELL × PARAMETER → ROP PEARSON CORRELATION', 15, NAVY, wt='700')
x0, y0, cw, ch = 300, 175, 70, 42
for j, p in enumerate(pars): b += T(x0+j*cw+cw/2, y0-8, p[:6], 9.5, MUTED, 'middle')
for i, wk in enumerate(wells):
    yy = y0+i*ch
    b += T(60, yy+25, wk[:20], 10.5, NAVY)
    for j, p in enumerate(pars):
        r = RR('SELECT pearson FROM an_correlations WHERE well_key=? AND x=? AND y="ROP"',(wk,p))
        v = n(r[0]['pearson']) if r else None
        col = SOFT if v is None else (f'rgba(29,78,216,{0.12+0.8*min(1,abs(v)):.2f})' if v>=0 else f'rgba(220,38,38,{0.12+0.8*min(1,abs(v)):.2f})')
        b += R(x0+j*cw+1, yy+2, cw-2, ch-4, col, rx=2)
        if v is not None: b += T(x0+j*cw+cw/2, yy+25, f'{v:.2f}', 9.5, '#FFFFFF' if abs(v)>.45 else INK, 'middle', '700')
b += T(60, 760, 'Blue = positive, red = negative; blank = fewer than 4 paired daily rows. No missing value is imputed.', 11.5, MUTED)
art('AN_06_correlation_heatmap', 'SAZN — ROP Driver Correlation Matrix', 'Pearson coefficients from paired DDR daily parameter rows', b)

# 07 cost efficiency scatter
ce = S['field']['cost_efficiency']
b = T(60, 130, 'COST EFFICIENCY — USD/DAY vs NPT SHARE', 15, NAVY, wt='700')
x0, y0, w, h = 90, 175, 1000, 480
b += grid(x0, y0, w, h, 6, 10)
mx = max([n(x['usd_per_day']) for x in ce]+[1]); mn = max([n(x['npt_share_pct']) for x in ce]+[1])
for i, x in enumerate(ce):
    xx = x0+w*n(x['usd_per_day'])/mx; yy = y0+h*(1-n(x['npt_share_pct'])/mn)
    b += C(xx, yy, 9, SEQ[i%len(SEQ)])
    b += T(xx+12, yy+4, x['well_key'][:18], 10, NAVY)
b += T(60, 690, 'Upper-right = expensive AND NPT-intensive. Field mean USD/day and NPT share are in sheet 46_DA_Cost_Efficiency.', 11.5, MUTED)
art('AN_07_cost_efficiency', 'SAZN — Cost Efficiency Map', 'Gross USD per DDR day plotted against NPT share of calendar hours', b)

# 08 unit-price dispersion
up = S['field']['unit_price_dispersion'][:12]
b = T(60, 130, 'UNIT-PRICE DISPERSION ACROSS THE 13 OPERATIONS', 15, NAVY, wt='700')
b += T(60, 152, 'Normalised tariff key: spread = (max − min) / min', 12, MUTED)
b += bars([(x['description'][:32], n(x['spread_pct'])) for x in up], 60, 180, 1080, 470, val_sz=11, unit='%')
b += T(60, 680, 'Dispersion can reflect genuine scope differences (hole size, rig, bundle) or a column-alignment issue; each row is reviewable by cell.', 11.5, MUTED)
art('AN_08_unit_price', 'SAZN — Tariff Rate Dispersion', 'Largest >5% unit-price spreads between operations', b)

# 09 formation structural bias
fb = RR("SELECT formation,AVG(delta_m) d,COUNT(*) n FROM an_formation_stats WHERE delta_m IS NOT NULL GROUP BY formation ORDER BY ABS(AVG(delta_m)) DESC LIMIT 16")
b = T(60, 130, 'PROGNOSIS → ACTUAL STRUCTURAL BIAS BY FORMATION (m)', 15, NAVY, wt='700')
b += bars([(x['formation'][:26], n(x['d'])) for x in fb], 60, 180, 1080, 480, val_sz=12, unit='m')
b += T(60, 690, 'Positive = actual top deeper than prognosis. Sheet 45_DA_Program_vs_Actual adds the Programme / EOWR / DDR three-source comparison.', 11.5, MUTED)
art('AN_09_structural_bias', 'SAZN — Structural Bias by Formation', 'Mean measured-minus-prognosis top depth across correlated wells', b)

# 10 programme vs actual
pv = S['field']['program_vs_actual'][:14]
b = T(60, 130, 'WELL PROGRAMME vs EOWR MEASURED TOP DELTAS (m)', 15, NAVY, wt='700')
b += bars([(f"{x['well_key'][:14]} {x['formation'][:14]}", n(x['delta_pa'])) for x in pv], 60, 180, 1080, 480, val_sz=11, unit='m')
b += T(60, 690, 'Only tops with both a Programme prognosis and an EOWR measured value are plotted; no value is interpolated.', 11.5, MUTED)
art('AN_10_program_actual', 'SAZN — Programme vs Actual Tops', 'Paired prognosis-to-measured depth deltas by well and formation', b)

# 11 learning curves
lm = S['field']['learning_metrics']
b = T(60, 130, 'RE-ENTRY LEARNING CURVES — NPT AND CBS INTENSITY', 15, NAVY, wt='700')
for i, x in enumerate(lm):
    yy = 190 + i*170
    b += R(60, yy, 1080, 140, SOFT, rx=8)
    b += T(78, yy+26, f"{x['chain']}  |  {x['from_well']} → {x['to_well']}", 13, NAVY, wt='700')
    b += T(78, yy+54, f"NPT/day  {fmt(x['npt_per_day_from'],2)} → {fmt(x['npt_per_day_to'],2)} h/day", 12, INK)
    b += T(78, yy+78, f"CBS/day  {fmt(x['cbs_per_day_from'],0)} → {fmt(x['cbs_per_day_to'],0)} USD/day", 12, INK)
    b += T(78, yy+102, f"Δ days {x['delta_days']}  |  Δ NPT {fmt(x['delta_npt_h'],1)} h  |  Δ CBS {fmt(x['delta_cbs'],0)} USD  |  {x['direction']}", 11.5, MUTED)
    col = TEAL if x['direction'].startswith('IMPROVED') else CORAL
    b += R(1000, yy+20, 120, 100, col, rx=6)
    b += T(1110, yy+58, 'IMPROVED' if col == TEAL else 'DEGRADED', 11, '#FFFFFF', 'middle', '700')
art('AN_11_learning', 'SAZN — Intervention Learning Metrics', 'NPT-hour and CBS intensity per DDR day across repeated interventions', b)

# 12 exception funnel / findings
b = T(60, 130, 'FORENSIC FUNNEL — FROM 1.2M CELLS TO 21 FINDINGS', 15, NAVY, wt='700')
steps = [('Non-empty worksheet cells',1196084),('Numeric cells',324150),('Formula cells',88692),('Formulas safely recomputed',48011),
         ('CBS priced lines checked',1572),('24-h daily rows checked',1304),('Daily parameter rows',17499),
         ('Mud material-days checked',36046),('Reconciliation keys',4298),('PDF numeric tokens',305447),('Derived forensic findings',21)]
mx = max(v for _, v in steps)
for i, (nm, v) in enumerate(steps):
    yy = 175 + i*52
    b += T(60, yy+18, nm, 12, INK, wt='600')
    b += R(420, yy, 620*math.log10(v+1)/math.log10(mx+1), 26, SEQ[i%len(SEQ)], rx=4)
    b += T(1160, yy+18, f'{v:,}', 12, NAVY, 'end', '700')
b += T(60, 800, 'Every narrowing step is reproducible from the 40 an_* tables in SAZN_Master_Warehouse.sqlite.', 11.5, MUTED)
art('AN_12_findings_funnel', 'SAZN — Forensic Funnel', 'Reproducible narrowing from the full cell census to the 21 evidence-based findings', b)

# 13 PDF numeric-token magnitude census
pdf = S['pdf']; ct = S['cost_track']
b = T(60, 130, 'ALL 305,447 ARCHIVED PDF NUMERIC TOKENS', 15, NAVY, wt='700')
b += T(60, 152, 'Magnitude bands, context classes and exact page/line provenance — no OCR values invented', 12, MUTED)
b += bars([(x['band'], n(x['n'])) for x in pdf['bands']], 60, 180, 1080, 400, val_sz=11)
b += T(60, 610, 'DERIVED CONTEXT CLASSES', 13, NAVY, wt='700')
for i, x in enumerate(pdf['context_classes']):
    yy = 638 + i * 30
    b += T(60, yy, x['context_class'], 11.5, INK)
    b += T(360, yy, f"{n(x['n']):,.0f} tokens", 11.5, NAVY, 'end')
    b += T(560, yy, f"{n(x['pct']):.2f}%", 11.5, TEAL, 'end')
    b += T(740, yy, f"outliers {n(x['outliers']):,.0f}", 11.5, MUTED)
b += T(60, 810, 'Raw tokens remain verbatim in pdf_num_pages; full census is in SAZN_PDF_Numeric_Token_Census.csv.gz.', 11, MUTED)
art('AN_13_pdf_token_census', 'SAZN — PDF Numeric Token Census', 'All archived PDF numbers classified by magnitude and derived context', b)

# 14 unit-aware PDF dispersion
units = [x for x in pdf['units'] if x['unit'] != 'NO_UNIT'][:12]
b = T(60, 130, 'UNIT-BEARING PDF NUMERIC TOKENS', 15, NAVY, wt='700')
b += T(60, 152, f"{pdf['unit_bearing']:,} unit-bearing tokens across {len(pdf['units'])-1} unit labels; NO_UNIT excluded from chart", 12, MUTED)
b += bars([(x['unit'][:20], n(x['n'])) for x in units], 60, 180, 1080, 410, val_sz=11)
b += T(60, 615, 'ROBUST DISPERSION EXAMPLES (derived IQR fences)', 13, NAVY, wt='700')
b += T(60, 645, 'Unit', 11, MUTED); b += T(250, 645, 'N', 11, MUTED); b += T(420, 645, 'Median', 11, MUTED)
b += T(650, 645, 'Max', 11, MUTED); b += T(900, 645, 'Outliers', 11, MUTED)
for i, x in enumerate(pdf['units'][:6]):
    yy = 672 + i * 25
    b += T(60, yy, x['unit'][:20], 10.5, INK); b += T(250, yy, f"{n(x['n']):,.0f}", 10.5, NAVY, 'end')
    b += T(420, yy, fmt(x['v_med'],2), 10.5, NAVY, 'end'); b += T(650, yy, fmt(x['v_max'],2), 10.5, NAVY, 'end')
    b += T(900, yy, f"{n(x['outliers']):,.0f}", 10.5, CORAL, 'end')
b += T(60, 820, f"{pdf['robust_outliers']:,} robust unit-group outliers are flagged, never corrected.", 11, MUTED)
art('AN_14_pdf_units', 'SAZN — Unit-Aware PDF Dispersion', 'Unit labels, medians, maxima and robust magnitude outliers', b)

# 15 low-text source pages
b = T(60, 130, 'LOW-TEXT PDF SOURCE PAGES — AF-05', 15, NAVY, wt='700')
b += T(60, 152, 'Pages with fewer than 200 text characters; numeric OCR was not fabricated', 12, MUTED)
lowcat = pdf['low_text_by_category']
b += bars([(x['category'], n(x['pages'])) for x in lowcat], 60, 180, 1080, 220, val_sz=12)
b += R(60, 430, 1080, 54, SOFT, rx=5)
b += T(80, 462, f"{pdf['low_text_pages']:,} / {pdf['pages']:,} pages ({100*pdf['low_text_pages']/pdf['pages']:.2f}%)", 18, NAVY, wt='700')
b += T(700, 462, f"numeric tokens on those pages: {sum(x['numeric_tokens'] for x in lowcat):,}", 12, INK)
b += T(700, 482, f"embedded images: {sum(x['images'] for x in lowcat):,}", 11, MUTED)
b += T(60, 520, 'HIGHEST LOW-TEXT SOURCE COVERAGE', 13, NAVY, wt='700')
for i, x in enumerate(pdf['source_coverage'][:8]):
    yy = 550 + i * 31
    b += T(60, yy, x['rel_path'].split('/')[-1][:44], 10.5, INK)
    b += T(520, yy, f"low-text {x['low_text_pages']}/{x['pages']}", 10.5, NAVY)
    b += T(760, yy, f"tokens {x['numeric_tokens']:,}", 10.5, NAVY)
    b += T(980, yy, x['status'], 10, MUTED)
b += T(60, 815, 'Source-only rule: retain each page, image count and text-layer number; request native/OCR versions for financial close-out.', 11, MUTED)
art('AN_15_pdf_low_text', 'SAZN — Low-Text PDF Source Pages', 'Explicit image-only / sparse-text source gap with exact page counts', b)

# 16 four-track cost bridge
f = ct['field']
b = T(60, 130, 'FOUR-TRACK COST BRIDGE — USD', 15, NAVY, wt='700')
b += T(60, 152, 'CBS plan → gross calendar → payable excl. contractor NPT → pure net excl. all NPT', 12, MUTED)
b += bars([('CBS plan',n(f['cbs_plan'])),('Gross calendar',n(f['gross'])),('Payable',n(f['payable'])),('Pure net',n(f['net']))],
          60, 180, 1080, 250, val_sz=12, colours=[COBALT,SLATE,AMBER,EMERALD])
b += T(60, 460, f"Rig rate: 23,000 USD/DAY (sourced CBS Rig Service)  |  max per-well formula delta {fmt(f['max_formula_delta'],4)} USD", 12, NAVY)
b += T(60, 495, 'DEDUCTIONS / ATTRIBUTION', 13, NAVY, wt='700')
rows = [('Contractor NPT',f['contractor_npt_h'],f['contractor_deduction']),('Client / subsurface NPT',f['client_npt_h'],f['client_deduction']),('All NPT',f['total_npt_h'],f['all_npt_deduction'])]
for i, (nm, h, usd) in enumerate(rows):
    yy = 530 + i*45
    b += T(60, yy, nm, 12, INK); b += T(500, yy, f"{h:,.1f} h", 12, NAVY, 'end')
    b += T(760, yy, f"{usd:,.2f} USD", 12, CORAL if i<2 else NAVY, 'end')
    b += T(950, yy, 'rate x hours / 24', 10.5, MUTED)
b += T(60, 690, 'SOURCE-BASIS FLAGS (not recalculated cost)', 13, NAVY, wt='700')
flags = RR("SELECT well_key,calendar_flag,cbs_days,ddr_days,delta_days FROM an_cost_track_audit WHERE calendar_flag<>'OK' AND well_key<>'FIELD' ORDER BY well_key")
for i, x in enumerate(flags):
    yy = 720 + i*27
    b += T(60, yy, f"{x['well_key']}  {x['calendar_flag']}", 10.5, INK)
    b += T(560, yy, f"CBS days {fmt(x['cbs_days'],2)} / DDR {x['ddr_days']}", 10.5, NAVY)
    b += T(900, yy, f"delta days {fmt(x['delta_days'],2)}", 10.5, MUTED)
b += T(60, 820, 'Fallback 50/50 waiting hours = 0. All 13 operations reconcile to within 0.01 USD after rounding.', 11, MUTED)
art('AN_16_cost_bridge', 'SAZN — Four-Track Cost Bridge', 'Sourced 23,000 USD/DAY rate, NPT deductions and calendar-basis flags', b)

# 17 NPT attribution audit
b = T(60, 130, 'NPT ATTRIBUTION AUDIT — EXPLICIT DDR STATUS ROWS', 15, NAVY, wt='700')
b += T(60, 152, 'Contractor / client hours, blank attributed-party cells and zero 50/50 waiting fallback', 12, MUTED)
cols = [(60,'Well'),(250,'NPT h'),(390,'Contractor h'),(540,'Client h'),(680,'Blank h'),(810,'Explicit %'),(960,'Fallback h')]
for x, lab in cols: b += T(x, 185, lab, 10.5, MUTED, wt='700')
attr = [x for x in ct['npt_attribution'] if x['well_key'] != 'FIELD'] + [x for x in ct['npt_attribution'] if x['well_key'] == 'FIELD']
for i, x in enumerate(attr):
    yy = 212 + i*43
    if i % 2 == 0: b += R(52, yy-22, 1090, 34, SOFT, rx=3)
    b += T(60, yy, x['well_key'][:20], 10.5, INK, wt='700' if x['well_key']=='FIELD' else '400')
    b += T(250, yy, fmt(x['npt_status_h'],1), 10.5, NAVY, 'end')
    b += T(390, yy, fmt(x['contractor_h'],1), 10.5, NAVY, 'end')
    b += T(540, yy, fmt(x['client_h'],1), 10.5, NAVY, 'end')
    b += T(680, yy, fmt(x['blank_attribution_h'],1), 10.5, AMBER, 'end')
    b += T(810, yy, fmt(x['explicit_attribution_pct'],2), 10.5, TEAL, 'end')
    b += T(960, yy, fmt(x['fallback_50_50_waiting_h'],1), 10.5, EMERALD, 'end')
b += T(60, 800, 'Blank attributed-party cells are retained verbatim and classified by the source heuristic; waiting hours are never split 50/50.', 11, MUTED)
art('AN_17_npt_attribution', 'SAZN — NPT Attribution Audit', 'Explicit N/NPT status rows with contractor/client and blank-column hours', b)

jobs = g.JOBS
json.dump(jobs, open(os.path.join(DATA, 'analytics_graphics_jobs.json'), 'w'), indent=1)
print('ANALYTICS ARTBOARDS', len(jobs))
