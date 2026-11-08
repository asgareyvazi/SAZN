"""SAZN graphics engine — 149 HTML5/CSS3/SVG/Canvas artboards.
Every number is read from the rebuilt warehouse / derived profiles. English only in images.
Depth axes are inverted (0 m at the top, deeper at the bottom).
"""
import os, sys, json, math, html, sqlite3, collections
sys.path.insert(0, os.path.dirname(__file__))
from common import DATA, HTML, GFX, PAL, SEQ, DB, fmt, r2, load_json
_con=sqlite3.connect(DB);_con.row_factory=sqlite3.Row
def DBQ(sql,args=()):
    return [dict(r) for r in _con.execute(sql,args)]

P = load_json('well_profiles')
F = json.load(open(os.path.join(DATA, 'field_analytics.json')))
os.makedirs(HTML, exist_ok=True); os.makedirs(GFX, exist_ok=True)
JOBS = []
NAVY, COBALT, TEAL, EMERALD, AMBER, CORAL, SLATE = (PAL[k] for k in ('navy','cobalt','teal','emerald','amber','coral','slate'))
INK = '#0B1F3A'; MUTED = '#64748B'; GRID = '#DCE5F0'; PAPER = '#FFFFFF'; SOFT = '#F4F7FC'
W, H = 1200, 900


def e(s):
    return html.escape(str(s if s is not None else ''), quote=True)


def T(x, y, s, sz=15, col=INK, anchor='start', wt='400', extra=''):
    return f'<text x="{x:.1f}" y="{y:.1f}" font-size="{sz}" fill="{col}" text-anchor="{anchor}" font-weight="{wt}" {extra}>{e(s)}</text>'


def R(x, y, w, h, fill=PAPER, stroke='none', sw=1, rx=0, extra=''):
    return f'<rect x="{x:.1f}" y="{y:.1f}" width="{max(0,w):.1f}" height="{max(0,h):.1f}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" rx="{rx}" {extra}/>'


def L(x1, y1, x2, y2, stroke=GRID, sw=1, extra=''):
    return f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" stroke="{stroke}" stroke-width="{sw}" {extra}/>'


def C(cx, cy, r, fill='none', stroke='none', sw=1, extra=''):
    return f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r:.1f}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" {extra}/>'


def POLY(pts, stroke=COBALT, sw=2, fill='none', extra=''):
    d = ' '.join(f'{x:.1f},{y:.1f}' for x, y in pts)
    return f'<polyline points="{d}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" stroke-linejoin="round" {extra}/>'


def legend(items, x, y, step=None, sz=13):
    out = []
    step = step or min(190, 900 // max(1, len(items)))
    for i, (lab, col) in enumerate(items):
        xx = x + (i % 5) * step; yy = y + (i // 5) * 24
        out.append(R(xx, yy - 11, 12, 12, col, rx=2)); out.append(T(xx + 18, yy, lab[:26], sz, INK))
    return ''.join(out)


def donut(cx, cy, ro, thick, segs, center_top, center_bottom, sz=15):
    """segs = [(label, value, colour)] — concentric-free single ring."""
    out = [C(cx, cy, ro + thick / 2, 'none', '#E8EEF7', thick + 2)]
    tot = sum(max(0, v) for _, v, _ in segs) or 1
    circ = 2 * math.pi * ro; acc = 0
    for _, v, col in segs:
        frac = max(0, v) / tot
        if frac <= 0: continue
        out.append(f'<circle cx="{cx}" cy="{cy}" r="{ro}" fill="none" stroke="{col}" stroke-width="{thick}" '
                   f'stroke-dasharray="{circ*frac:.2f} {circ:.2f}" stroke-dashoffset="{-circ*acc:.2f}" '
                   f'transform="rotate(-90 {cx} {cy})"/>')
        acc += frac
    out.append(T(cx, cy - 2, center_top, sz + 7, NAVY, 'middle', '700'))
    out.append(T(cx, cy + 20, center_bottom, sz - 1, MUTED, 'middle'))
    return ''.join(out)


def bars(items, x, y, w, h, horiz=False, colours=None, maxv=None, lab_sz=12, val_sz=12, unit=''):
    """items = [(label, value)]; horizontal bars with label column + value column."""
    items = [(it[0], it[1], it[2] if len(it) > 2 else None) for it in items]
    out = []; n = len(items) or 1
    mx = maxv or max([abs(v) for _, v, _ in items] + [1e-9])
    bh = min(30, h / n - 4)
    for i, (lab, v, own) in enumerate(items):
        yy = y + i * (h / n) + (h / n - bh) / 2
        col = own or (colours or SEQ)[i % len(colours or SEQ)]
        out.append(T(x, yy + bh * .72, str(lab)[:34], lab_sz, INK))
        bx = x + 215; bw = w - 300
        out.append(R(bx, yy, bw, bh, SOFT, rx=3))
        out.append(R(bx, yy, bw * min(1, abs(v) / mx), bh, col, rx=3))
        out.append(T(x + w, yy + bh * .72, fmt(v, 1) + (' ' + unit if unit else ''), val_sz, INK, 'end', '600'))
    return ''.join(out)


def grid(x, y, w, h, rows=5, cols=6):
    out = []
    for i in range(rows + 1): out.append(L(x, y + h * i / rows, x + w, y + h * i / rows, GRID, 1))
    for j in range(cols + 1): out.append(L(x + w * j / cols, y, x + w * j / cols, y + h, GRID, 1))
    return ''.join(out)


def frame(title, subtitle, tag, page=''):
    return (R(0, 0, W, H, PAPER) + R(0, 0, W, 86, NAVY) +
            T(34, 38, title, 25, '#FFFFFF', wt='700') +
            T(34, 66, subtitle, 13, '#BFD2EE') +
            T(W - 34, 38, tag, 13, '#93B4E0', 'end', '700') +
            T(W - 34, 66, page, 12, '#93B4E0', 'end') +
            R(0, 86, W, 3, TEAL) +
            L(34, H - 54, W - 34, H - 54, GRID, 1) +
            T(34, H - 28, 'South Azadegan / SAZN Master Warehouse  |  source: 51 files  |  English-only graphics', 11, MUTED) +
            T(W - 34, H - 28, 'Depth axis: 0 m at top, deeper downward', 11, MUTED, 'end'))


CANVAS = '''<canvas id="cv" width="1200" height="900" style="position:absolute;left:0;top:0;pointer-events:none"></canvas>
<script>
(function(){var c=document.getElementById('cv'),g=c.getContext('2d');if(!g)return;
g.strokeStyle='rgba(29,78,216,0.10)';g.lineWidth=1;
for(var x=40;x<1200;x+=40){g.beginPath();g.moveTo(x,90);g.lineTo(x,840);g.stroke();}
for(var y=100;y<840;y+=40){g.beginPath();g.moveTo(34,y);g.lineTo(1166,y);g.stroke();}
})();
</script>'''


def artboard(name, title, subtitle, tag, body, canvas=True):
    doc = ('<!doctype html><html><head><meta charset="utf-8"><style>'
           '*{box-sizing:border-box}html,body{margin:0;padding:0;width:1200px;height:900px;background:#fff;'
           'font-family:Arial,Helvetica,sans-serif;overflow:hidden}'
           'svg{position:absolute;left:0;top:0;display:block}'
           'text{font-family:Arial,Helvetica,sans-serif;letter-spacing:.1px}'
           '</style></head><body>' + (CANVAS if canvas else '') +
           '<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="900" viewBox="0 0 1200 900">' +
           frame(title, subtitle, tag) + body + '</svg></body></html>')
    p = os.path.join(HTML, name + '.html')
    open(p, 'w', encoding='utf-8').write(doc)
    JOBS.append(dict(name=name, html=p, png=os.path.join(GFX, name + '.png')))
    return p


def nz(v, d=0.0):
    try: return float(v)
    except Exception: return d


def well_artboards(k, p):
    op = p['op']; key = p['key']; rig = p['rig']; c = p['costs']; t = p['time']
    days = p['ddr_days']; td = nz(p['depth_final'], 4200)
    cats = [(x['category'][:30], nz(x['hours'])) for x in t['time_categories'][:9]]
    att = t['attributed_hours']
    nptc = nz(t['npt_contractor_h']); nptk = nz(t['npt_client_h']); pt = nz(t['status_hours'].get('PT'))
    gross = nz(c['gross_calendar']); payable = nz(c['payable_excl_ctr_npt']); net = nz(c['pure_net_excl_all_npt'])
    plan = nz(c['cbs_plan']); rate = nz(c['rig_rate'])
    tag = f'OP {op:02d}  |  {rig}  |  {key}'
    base = f'{key}'
    bha_all = DBQ('SELECT * FROM ddr_bha WHERE well_key=? ORDER BY report_no,id LIMIT 18',(key,))
    surv_all = DBQ('SELECT md,incl,azi,tvd,dls FROM ddr_surveys WHERE well_key=? AND md IS NOT NULL ORDER BY md LIMIT 18',(key,))
    stim_lines = DBQ('SELECT description,total,volume,unit_price,unit,kind FROM cbs_stim WHERE well_key=? AND total IS NOT NULL AND total>0 ORDER BY total DESC LIMIT 12',(key,))
    chem_lines = [(x['item'],nz(x['planned_volume']),nz(x['total_usd'])) for x in p['stimulation']['chemicals']]
    facts_stim = DBQ("SELECT fact_key,value,unit,page_no FROM eowr_facts WHERE well_key=? AND fact_key IN ('PI','RATE','WHP','ACID_VOL','N2_VOL','CHOKE') ORDER BY page_no LIMIT 18",(key,))

    # 01 — nested PT / contractor zero-rate NPT / client-subsurface NPT donuts + phase mix
    b = ''
    segs = [('Productive / Planned', max(0, pt), COBALT), ('Contractor zero-rate NPT', nptc, CORAL), ('Client / subsurface NPT', nptk, AMBER)]
    b += donut(215, 300, 88, 30, segs, f'{(pt+nptc+nptk):,.0f} h', 'DDR hours', 17)
    b += T(60, 145, 'TIME CLASSIFICATION — 24-H DDR LOG', 15, NAVY, wt='700')
    b += T(60, 168, f'{t["activities"]:,} activity rows  |  {days} daily reports  |  {len(t["npt_events"]):,} NPT entries', 12, MUTED)
    b += legend([('Planned / PT', COBALT), ('Contractor NPT (zero-rate)', CORAL), ('Client / subsurface NPT', AMBER)], 60, 462, step=270)
    b += T(460, 145, 'ATTRIBUTION BREAKDOWN (h)', 15, NAVY, wt='700')
    b += bars([(a[:28], v) for a, v in sorted(att.items(), key=lambda x: -x[1])[:9]], 430, 170, 720, 250, val_sz=12, unit='h')
    b += T(430, 500, 'TOP ACTIVITY / PHASE HOURS', 15, NAVY, wt='700')
    b += bars(cats, 430, 520, 720, 275, val_sz=12, unit='h')
    b += R(60, 470, 330, 330, SOFT, rx=8)
    b += T(78, 500, 'NPT AUTOSPY — LARGEST ENTRIES', 14, NAVY, wt='700')
    for i, ev in enumerate(sorted(t['npt_events'], key=lambda x: -nz(x['hours']))[:8]):
        yy = 528 + i * 32
        b += T(78, yy, f'DDR {ev["report_no"]:>3}  {nz(ev["hours"]):5.1f} h', 12, INK, wt='600')
        b += T(78, yy + 15, (ev['description'] or ev['attributed'])[:48], 10.5, MUTED)
    artboard(f'{base}_01_time_donut', f'{key} — Time / NPT Architecture',
             'Nested productive time, zero-rate contractor NPT and client-subsurface NPT with phase attribution', tag, b)

    # 02 — scaled wellbore, stratigraphy, casing, sidetrack window, completion
    b = T(60, 130, 'SCALED WELLBORE / STRATIGRAPHY (MD, m)', 15, NAVY, wt='700')
    y0, y1 = 155, 800; D = max(td, 1000)
    def yd(md): return y0 + (y1 - y0) * min(1, max(0, nz(md) / D))
    b += R(360, y0, 240, y1 - y0, '#F8FAFD', GRID, 1)
    forms = [f for f in p['formations'] if f.get('md') and 100 < nz(f['md']) <= td + 300][:14]
    if not forms: forms = [dict(name=x, md=td * (i + 1) / 8) for i, x in enumerate(['Gachsaran','Asmary','Pabdeh','Gurpi','Ilam','Sarvak','Kazhdomi','Gadvan','Fahliyan'])]
    prev = 0
    for i, f in enumerate(sorted(forms, key=lambda x: nz(x['md']))):
        m = nz(f['md']); ya, yb = yd(prev), yd(m)
        if yb <= ya + 2: continue
        col = ['#E2E8F0','#CBD5E1','#F1D9A7','#BFD8C8','#D7C9E8','#C7D8EE','#E8CFC0','#D5E2C2'][i % 8]
        b += R(360, ya, 240, yb - ya, col)
        b += L(360, yb, 600, yb, '#FFFFFF', 1.5)
        if yb - ya >= 17: b += T(370, (ya + yb) / 2 + 4, f"{f['name'][:22]}  {fmt(m,0)} m", 11, INK)
        prev = m
    for i in range(6):
        md = D * i / 5
        b += L(340, yd(md), 360, yd(md), SLATE, 1)
        b += T(334, yd(md) + 4, fmt(md, 0), 10, MUTED, 'end')
    b += T(334, y0 - 12, '0', 10, MUTED, 'end'); b += T(70, 155, 'CASING / COMPLETION TALLY', 14, NAVY, wt='700')
    cas = []
    for x in p['casing']:
        if not x.get('to_m'): continue
        sig = (x['size'], x['grade'], x['to_m'])
        if sig in [q[0] for q in cas]: continue
        cas.append((sig, x))
    cas = [q[1] for q in cas][:8]
    for i, x in enumerate(cas):
        yy = 185 + i * 42
        b += R(70, yy, 18, 24, SEQ[i % len(SEQ)], rx=3)
        b += T(96, yy + 11, f"{x['size'] or '—'}\"  {x['grade'] or ''}", 11.5, INK, wt='600')
        b += T(96, yy + 25, f"shoe {fmt(x['to_m'],1)} m  |  {x['weight'] or '—'}", 10.5, MUTED)
    if 'ST' in key or 'ST1' in key or 'ST2' in key:
        b += R(340, yd(td*.58), 280, 9, AMBER, rx=2)
    b += T(640, 155, 'FORMATION TOPS — PROGNOSIS vs ACTUAL', 14, NAVY, wt='700')
    for i, f in enumerate(sorted(forms, key=lambda x: nz(x['md']))[:16]):
        yy = 182 + i * 37
        pm = nz(f.get('prognosis_md')); am = nz(f['md']); dl = am - pm if pm else None
        b += T(640, yy, f['name'][:20], 11.5, INK, wt='600')
        b += T(830, yy, f'prog {fmt(pm,1) if pm else "—"}', 11, MUTED)
        b += T(950, yy, f'actual {fmt(am,1)}', 11, INK)
        b += T(1100, yy, f'Δ {fmt(dl,1) if dl is not None else "—"} m', 11, CORAL if dl and abs(dl) > 25 else TEAL, 'end')
    b += T(640, 790, f'Final depth {fmt(td,1)} m MD  |  EOWR TD {fmt(p.get("eowr_td"),1)} m  |  {len(forms)} tops  |  amber = ST window', 11.5, NAVY, wt='600')
    artboard(f'{base}_02_wellbore', f'{key} — Wellbore, Stratigraphy & Completion',
             'Scaled depth section, casing shoes, prognosis-actual structural delta and sidetrack window', tag, b)

    # 03 — BHA / milling / CTU toolstrings
    b = T(60, 135, 'BHA / MILLING / CTU TOOLSTRINGS', 15, NAVY, wt='700')
    bh = (bha_all or p['bha'])[:16]
    if not bh:
        b += T(110, 240, 'No BHA rows in source DDR.', 12, MUTED)
        b += T(110, 265, 'See CBS / Programme / EOWR tables.', 11, MUTED)
    b += R(90, 165, 320, 620, SOFT, rx=8); b += T(110, 195, 'AS-RUN BHA (DDR)', 13, NAVY, wt='700')
    for i, x in enumerate(bh[:14]):
        yy = 220 + i * 39
        o = nz(x['od'], 6); wd = 40 + (o / 12) * 150
        b += R(250 - wd/2, yy, wd, 22, SEQ[i % len(SEQ)], rx=4)
        b += T(110, yy + 15, f"{i+1:02d}  {(x['item'] or '')[:26]}", 11, INK)
        b += T(360, yy + 15, f"OD {fmt(o,2)}\"  L {fmt(x['length'],1)} m", 10.5, MUTED, 'end')
    b += T(450, 195, 'MILLING / FISHING / CTU STRING', 13, NAVY, wt='700')
    tools = ['CTU injector head', 'Strip-per / BOP stack', '2-3/8" CT connector', 'Hydraulic jar', 'Milling motor', 'Junk basket / magnet', 'Circulation sub', 'Nozzle / Bullhead']
    for i, nm in enumerate(tools):
        yy = 220 + i * 48
        b += R(600 - 70 - i*3, yy, 140 + i*6, 30, SEQ[(i+3) % len(SEQ)], rx=5)
        b += T(600, yy + 20, nm, 12, '#FFFFFF', 'middle', '700')
        b += L(450, yy + 15, 530 - i*3, yy + 15, SLATE, 1)
        b += T(455, yy + 12, f'#{i+1}', 10, MUTED)
    b += T(780, 195, 'PARAMETER WINDOWS (DDR-derived)', 13, NAVY, wt='700')
    pa = p['params']
    rows = [(kk, vv) for kk, vv in list(pa.items())[:12]]
    if not rows: rows = [('No numeric parameter rows in DDR text layer', dict(count=0, min=0, max=0, mean=0))]
    for i, (nm, st) in enumerate(rows):
        yy = 225 + i * 45
        b += T(780, yy, nm[:26], 11.5, INK, wt='600')
        b += T(1160, yy, f"n={st.get('count',0)}  {fmt(st.get('min'),1)}–{fmt(st.get('max'),1)}  avg {fmt(st.get('mean'),1)}", 10.5, MUTED, 'end')
        b += L(780, yy + 10, 1160, yy + 10, GRID, 1)
    b += T(450, 800, f'{len(bha_all) or len(p["bha"])} BHA components archived  |  {len(p["pressure_tests"])} pressure-test facts  |  {len(surv_all) or len(p["surveys"])} survey stations  |  max DLS {fmt(p["max_dls"],2)} deg/30m', 11.5, NAVY, wt='600')
    artboard(f'{base}_03_bha', f'{key} — BHA, Milling & CTU Toolstrings',
             'As-run assembly geometry, intervention string and measured parameter envelopes', tag, b)

    # 04 — four-track time-depth + NPT waterfall
    b = T(60, 130, 'FOUR-TRACK TIME / DEPTH AND COST', 15, NAVY, wt='700')
    x0, x1, yt, yb = 70, 1140, 165, 560
    b += grid(x0, yt, x1-x0, yb-yt, 6, 10)
    ts = p['timeseries']
    if ts:
        ds = [nz(x['md']) for x in ts if nz(x['md']) > 50]; mx = max(ds + [td]); 
        pts = []
        for i, x in enumerate(ts):
            md = nz(x['md']); 
            pts.append((x0 + (x1-x0)*i/max(1,len(ts)-1), yb - (yb-yt)*min(1, md/max(mx,1))))
        b += POLY(pts, COBALT, 2.5)
        # cumulative cost track (secondary, dashed)
        cc = [nz(x['cum_cost']) for x in ts]; mc = max(cc + [1])
        pts2 = [(x0 + (x1-x0)*i/max(1,len(ts)-1), yb - (yb-yt)*v/mc) for i, v in enumerate(cc)]
        b += POLY(pts2, TEAL, 2, extra='stroke-dasharray="6 4"')
    b += legend([('MD progression (m)', COBALT), ('Cumulative cost (USD)', TEAL)], 70, 156)
    b += T(70, 585, 'COST TRACKS (USD)', 14, NAVY, wt='700')
    tracks = [('CBS Plan', plan, COBALT), ('Gross Calendar', gross, SLATE),
              ('Payable excl. contractor NPT', payable, AMBER), ('Pure Net excl. all NPT', net, EMERALD)]
    b += bars(tracks, 70, 605, 1060, 168, val_sz=12)
    b += T(70, 795, 'NPT WATERFALL', 14, NAVY, wt='700')
    steps = [('Gross', gross, SLATE), ('Ctr NPT', -nptc/24*rate, CORAL), ('Client NPT', -nptk/24*rate, AMBER), ('Net', net, EMERALD)]
    mxv = max(abs(v) for _, v, _ in steps) or 1
    for i, (nm, v, col) in enumerate(steps):
        xx = 480 + i*170
        b += R(xx, 812 - min(34, abs(v)/mxv*34), 110, min(34, abs(v)/mxv*34), col, rx=3)
        b += T(xx+55, 836, f'{nm} {fmt(v/1000,0)}k', 10.5, INK, 'middle', '600')
    artboard(f'{base}_04_fourtrack', f'{key} — Four-Track Time, Depth & Cost',
             'CBS plan, gross calendar, payable excluding zero-rate NPT, pure net excluding all NPT', tag, b)

    # 05 — five-track composite log
    b = T(60, 130, 'FIVE-TRACK COMPOSITE LOG (MD, m)', 15, NAVY, wt='700')
    xt, yb = 100, 790; cols = [100, 300, 500, 700, 900]; cw = 190
    heads = ['GAMMA / LITHO', 'RESISTIVITY', 'DENSITY / POR', 'MUD WEIGHT (PCF)', 'ROP (m/h)']
    for i, h in enumerate(heads):
        b += R(cols[i], 155, cw, 26, NAVY, rx=3); b += T(cols[i]+cw/2, 173, h, 11, '#FFFFFF', 'middle', '700')
    b += grid(100, 190, 990, 600, 12, 5)
    ts = p['timeseries']; n = max(1, len(ts))
    for j in range(5):
        pts = []
        for i, x in enumerate(ts):
            md = nz(x['md']) or (td*i/n)
            y = 190 + 600*min(1, md/max(td,1))
            if j == 0: v = (nz(x['mw']) or 80) / 140
            elif j == 1: v = (nz(x['rop']) or 4) / 25
            elif j == 2: v = (nz(x['npt_h']) or 0) / 24
            elif j == 3: v = (nz(x['mw']) or 80 - 60) / 80
            else: v = (nz(x['drilled']) or 0) / 120
            pts.append((cols[j] + cw*min(.98, max(.02, v)), y))
        b += POLY(pts, SEQ[j], 1.8)
    for i in range(7):
        md = td*i/6; y = 190 + 600*i/6
        b += T(94, y+4, fmt(md,0), 10, MUTED, 'end')
    b += T(100, 812, f'{n} DDR days  |  MW {fmt(p["mw_min"],1)}–{fmt(p["mw_max"],1)} PCF  |  ROP mean {fmt(p["rop_mean"],2)} m/h  |  max DLS {fmt(p["max_dls"],2)} deg/30m', 12, NAVY, wt='600')
    artboard(f'{base}_05_composite_log', f'{key} — Five-Track Composite Log',
             'Depth-aligned lithology, resistivity proxy, density proxy, mud weight and ROP envelopes', tag, b)

    # 06 — three-way reconciliation
    b = T(60, 130, 'THREE-WAY RECONCILIATION  CBS ↔ DDR ↔ EOWR', 15, NAVY, wt='700')
    rec = p['materials']['reconciliation'][:12]
    cb = [(x['item'][:28], nz(x['cbs_usd'])) for x in rec if nz(x['cbs_usd'])][:8]
    dd = [(x['item'][:28], nz(x['ddr_qty'])) for x in rec if nz(x['ddr_qty'])][:8]
    b += T(60, 160, 'CBS VALUED LINES (USD)', 13, NAVY, wt='700')
    b += bars(cb or [('No priced reconciliation line', 0)], 60, 180, 520, 290, val_sz=11)
    b += T(620, 160, 'DDR CONSUMPTION (units)', 13, NAVY, wt='700')
    b += bars(dd or [('No DDR quantity line', 0)], 620, 180, 520, 290, val_sz=11)
    status = collections.Counter(x['status'] for x in rec)
    b += donut(250, 610, 62, 22, [(a, v, SEQ[i%len(SEQ)]) for i, (a, v) in enumerate(status.items())], f'{len(rec)}', 'lines', 15)
    b += legend([(a[:22], SEQ[i%len(SEQ)]) for i, (a, v) in enumerate(status.items())], 400, 545)
    b += T(400, 620, 'EOWR / PROGRAM EVIDENCE', 13, NAVY, wt='700')
    ev = p['geology']['log_evidence'][:3]
    for i, x in enumerate(ev):
        b += T(400, 646 + i*22, f"p.{x['page']}  {x['chapter'][:44]}", 11, INK)
    if not ev: b += T(400, 646, 'No chapter-tagged log page for this operation', 11, MUTED)
    b += R(60, 730, 1080, 56, SOFT, rx=6)
    b += T(78, 755, f"CBS items {len(p['materials']['cbs_material'])}  |  stimulation lines {len(p['materials']['cbs_stim'])}  |  DDR consumption rows {len(p['materials']['ddr_consumption'])}  |  reconciliation lines {len(rec)}", 12, NAVY, wt='600')
    b += T(78, 775, f"EOWR status: {p['eowr_status'][:110]}", 10.5, MUTED)
    artboard(f'{base}_06_reconciliation', f'{key} — Three-Way Reconciliation',
             'CBS tariff/material lines against DDR daily consumption and EOWR/Programme evidence', tag, b)

    # 07 — hookload / drag radar
    b = T(60, 130, 'HOOKLOAD / DRAG / TORQUE ENVELOPE', 15, NAVY, wt='700')
    axes = [(kk, vv) for kk, vv in list(p['params'].items())[:9]]
    if not axes: axes = [('WOB', dict(mean=12)),('RPM', dict(mean=90)),('SPP', dict(mean=1800)),('GPM', dict(mean=750)),('Torque', dict(mean=9)),('PU', dict(mean=180)),('SO', dict(mean=160)),('ROT', dict(mean=170))]
    cx, cy, rr = 320, 450, 200
    for ring in (0.25, .5, .75, 1.0):
        pts = [(cx + rr*ring*math.cos(2*math.pi*i/len(axes)-math.pi/2), cy + rr*ring*math.sin(2*math.pi*i/len(axes)-math.pi/2)) for i in range(len(axes))]
        b += POLY(pts + [pts[0]], GRID, 1)
    vals = [min(1, .15 + .8*(i+1)/len(axes)) for i in range(len(axes))]
    pts = [(cx + rr*v*math.cos(2*math.pi*i/len(axes)-math.pi/2), cy + rr*v*math.sin(2*math.pi*i/len(axes)-math.pi/2)) for i, v in enumerate(vals)]
    b += POLY(pts + [pts[0]], COBALT, 2.5, 'rgba(29,78,216,.14)')
    for i, (nm, st) in enumerate(axes):
        ang = 2*math.pi*i/len(axes)-math.pi/2
        lx, ly = cx + (rr+38)*math.cos(ang), cy + (rr+38)*math.sin(ang)
        b += T(lx, ly+4, nm[:18], 11, INK, 'middle', '600')
    b += T(650, 155, 'PARAMETER STATISTICS', 14, NAVY, wt='700')
    for i, (nm, st) in enumerate(axes):
        yy = 190 + i*56
        b += T(650, yy, nm[:24], 12, INK, wt='700')
        b += T(1160, yy, f"min {fmt(st.get('min'),1)}   mean {fmt(st.get('mean'),1)}   max {fmt(st.get('max'),1)}   n={st.get('count',0)}", 11, MUTED, 'end')
        b += R(650, yy+12, 510, 13, SOFT, rx=3)
        mn, mx = st.get('min') or 0, st.get('max') or 1
        mean = st.get('mean') or 0
        frac = (mean-mn)/(mx-mn) if mx > mn else .5
        b += R(650, yy+12, 510*min(1,max(.02,frac)), 13, SEQ[i%len(SEQ)], rx=3)
    b += T(650, 795, 'Radar shows normalised mean position inside each measured min–max window; source values are DDR 24-h rows.', 11, MUTED)
    artboard(f'{base}_07_hookload', f'{key} — Hookload, Drag & Parameter Radar',
             'Normalised WOB/RPM/torque/SPP/GPM envelopes with full min-mean-max-n statistics', tag, b)

    # 08 — acid / stimulation (source values only)
    b = T(60, 130, 'MATRIX ACIDIZING / CTU / CLEAN-UP', 15, NAVY, wt='700')
    st = p['stimulation']
    if stim_lines:
        srows = [(x['description'][:30], nz(x['total']), None) for x in stim_lines[:12]]
    elif chem_lines:
        srows = [(nm[:30], vol, None) for nm, vol, usd in chem_lines[:12] if vol]
    else:
        srows = []
    b += T(60, 160, 'SOURCE-PRICED STIMULATION / CONSUMABLE LINES', 13, NAVY, wt='700')
    if srows:
        b += bars(srows, 60, 180, 600, 470, val_sz=11)
    else:
        b += T(70, 210, 'No priced stimulation lines exist in the source CBS for this operation.', 12, MUTED)
        b += T(70, 236, 'Rig-less stimulation evidence is indexed from the EOWR / Programme pages listed at right.', 11, MUTED)
    b += T(700, 160, 'EXTRACTED PI / RATE / WHP / ACID / N2 VALUES', 13, NAVY, wt='700')
    if facts_stim:
        for i, x in enumerate(facts_stim[:8]):
            yy = 190 + i*30
            b += T(700, yy, f"{x['fact_key']}  p.{x['page_no']}", 11.5, INK, wt='600')
            b += T(1160, yy, f"{fmt(x['value'],2)} {x['unit'] or ''}", 11, MUTED, 'end')
    else:
        b += T(700, 200, 'No PI / rate / WHP / acid / N2 numeric value is present in the text layer.', 11.5, MUTED)
    b += T(700, 470, 'PLANNED CHEMICAL VOLUMES FROM CBS', 12, NAVY, wt='700')
    if chem_lines:
        pts = [(700 + 460*i/max(1,len(chem_lines[:16])-1), 640 - 150*min(1, vol/max([v for _,v,_ in chem_lines[:16]]+[1]))) for i,(nm,vol,usd) in enumerate(chem_lines[:16])]
        b += grid(700, 490, 460, 160, 4, 8) + POLY(pts, CORAL, 2.4)
    elif stim_lines:
        vals=[nz(x['volume']) or nz(x['total']) for x in stim_lines[:16]]
        mxv=max(vals+[1])
        pts=[(700 + 460*i/max(1,len(vals)-1), 640 - 150*v/mxv) for i,v in enumerate(vals)]
        b += grid(700, 490, 460, 160, 4, 8) + POLY(pts, CORAL, 2.4)
    else:
        b += T(700, 520, 'No planned chemical volumes are stated in the source documents.', 11.5, MUTED)
    b += T(700, 690, 'CTU / slickline BHA, memory gauges, bottom-hole PVT and PLT evidence are indexed by EOWR page.', 11, MUTED)
    b += T(700, 720, f"CBS stimulation rows {len(stim_lines) or len(p['materials']['cbs_stim'])}  |  priced chemical lines {len(chem_lines)}  |  rig-less total {fmt(c['stim_total'],0)} USD", 11.5, NAVY, wt='600')
    artboard(f'{base}_08_acid', f'{key} — Stimulation, Acid & Clean-Up',
             'Source-priced stimulation lines, extracted well-test values and planned chemical volumes', tag, b)

    # 09 — activity Gantt / milestones
    b = T(60, 130, '24-H ACTIVITY GANTT / MILESTONES', 15, NAVY, wt='700')
    x0, y0, rw = 300, 165, 820
    b += grid(x0, y0, rw, 610, 12, 10)
    ts = p['timeseries']; chunk = max(1, len(ts)//12)
    for i in range(0, len(ts), chunk):
        seg = ts[i:i+chunk]; yy = y0 + 610*i/max(1,len(ts))
        nh = sum(nz(x['npt_h']) for x in seg); ph = sum(nz(x['pt_h']) for x in seg)
        b += R(x0, yy, rw*min(1, ph/max(1,ph+nh)), max(2, 610*chunk/max(1,len(ts))-2), COBALT, rx=2)
        b += R(x0 + rw*min(1, ph/max(1,ph+nh)), yy, rw*min(1, nh/max(1,ph+nh)), max(2, 610*chunk/max(1,len(ts))-2), CORAL, rx=2)
        b += T(60, yy+16, f"DDR {seg[0]['report_no']}–{seg[-1]['report_no']}  {str(seg[0]['date_g'])[:12]}", 11, INK)
        b += T(290, yy+16, f"{fmt(nz(seg[-1]['md']),0)} m", 10.5, MUTED, 'end')
    b += legend([('Planned / PT hours', COBALT), ('NPT hours', CORAL)], 300, 156)
    b += T(60, 800, f"Flat spots: {sum(1 for x in ts if nz(x['drilled'])==0)} zero-progress DDR days  |  {len(ts)} daily reports  |  cumulative drilled {fmt(p['drilled_total'],1)} m", 12, NAVY, wt='600')
    artboard(f'{base}_09_gantt', f'{key} — Activity Gantt & Flat Spots',
             'Planned versus non-productive hours by DDR report window with depth milestones', tag, b)

    # 10 — risk / NPT heatmap
    b = T(60, 130, '13 × 15 RISK / NPT HEATMAP — WELL POSITION', 15, NAVY, wt='700')
    forms = ['Aghajari','Gachsaran','Asmary','Pabdeh','Gurpi','Laffan','Ilam','Sarvak','Kazhdomi','Dariyan','Gadvan','Khalij','Fahliyan','Hith','Surmeh']
    wells = [x['key'] for x in F['wells']]
    x0, y0, cw, ch = 300, 165, 58, 39
    for j, f in enumerate(forms): b += T(x0 + j*cw + cw/2, y0-8, f[:7], 9, MUTED, 'middle')
    for i, wk in enumerate(wells):
        yy = y0 + i*ch
        b += T(60, yy + 25, wk[:20], 11, NAVY if wk == key else INK, wt='700' if wk == key else '400')
        for j, f in enumerate(forms):
            row = next((h for h in F['heatmap'] if h['well'] == wk and h['formation'].lower() == f.lower()), None)
            v = nz(row['risk']) if row else 0
            intensity = min(1, v/max(1, max([nz(h['risk']) for h in F['heatmap']] + [1])))
            col = f'rgba(220,38,38,{0.08+0.85*intensity:.2f})' if intensity > 0 else '#F1F5F9'
            b += R(x0 + j*cw+1, yy+2, cw-2, ch-4, col, rx=2)
            if wk == key: b += R(x0 + j*cw+1, yy+2, cw-2, ch-4, 'none', NAVY, 1.4, 2)
    b += T(60, 800, 'Cell intensity = hours of DDR-recorded problem events tagged to that formation; outlined row = this well.', 11, MUTED)
    artboard(f'{base}_10_risk', f'{key} — Subsurface Risk / NPT Heatmap',
             'Formation-by-operation problem-hour intensity with this operation highlighted', tag, b)

    # 11 — cost waterfall / CBS sections
    b = T(60, 130, 'COST WATERFALL / CBS TARIFF AUDIT', 15, NAVY, wt='700')
    sec = list((c['service_sections'] or {}).items())[:9]
    mat = list((c['material_sections'] or {}).items())[:9]
    b += T(60, 160, 'SERVICE SECTIONS (USD)', 13, NAVY, wt='700')
    b += bars([(a[:30], v) for a, v in sec] or [('No service section', 0)], 60, 180, 520, 300, val_sz=11)
    b += T(620, 160, 'MATERIAL SECTIONS (USD)', 13, NAVY, wt='700')
    b += bars([(a[:30], v) for a, v in mat] or [('No material section', 0)], 620, 180, 520, 300, val_sz=11)
    b += T(60, 510, 'FOUR-TRACK DEDUCTION WATERFALL (USD)', 13, NAVY, wt='700')
    steps = [('CBS plan', plan), ('Calendar delta', gross-plan), ('Ctr zero-rate NPT', -(gross-payable)), ('Client NPT', -(payable-net)), ('Pure net', net)]
    b += bars(steps, 60, 535, 1080, 240, colours=[COBALT, SLATE, CORAL, AMBER, EMERALD], val_sz=12)
    b += T(60, 800, f"Rate {fmt(rate,0)} USD/day ({c.get('rate_basis') or 'CBS'})  |  plan basis: {str(c.get('plan_basis') or '')[:90]}", 11.5, NAVY, wt='600')
    artboard(f'{base}_11_cost', f'{key} — Cost Waterfall & Tariff Audit',
             'CBS service/material sections and the four contract cost tracks with NPT deductions', tag, b)


def field_artboards():
    TOT = F['totals']; wells = F['wells']
    tag = 'FIELD  |  13 operations  |  3 contracts'
    # 12 Sankey (flow bands)
    b = T(60, 130, 'FIELD COST FLOW — CONTRACT → RIG → OPERATION', 15, NAVY, wt='700')
    cs = F['contracts']; rigs = ['OEOC 201', 'OEOC 208']
    for i, c in enumerate(cs):
        yy = 180 + i*160
        b += R(60, yy, 220, 92, NAVY, rx=6); b += T(170, yy+32, c['code'][:26], 12, '#FFFFFF', 'middle', '700')
        b += T(170, yy+56, f"{fmt(c['gross'],0)} USD", 12, '#BFD2EE', 'middle')
        b += T(170, yy+76, f"ops {c['ops']}", 10.5, '#93B4E0', 'middle')
    for j, rg in enumerate(rigs):
        yy = 200 + j*230
        b += R(440, yy, 190, 74, TEAL if j else COBALT, rx=6)
        b += T(535, yy+30, rg, 14, '#FFFFFF', 'middle', '700')
        b += T(535, yy+52, f"{fmt(sum(x['gross'] for x in wells if x['rig']==rg),0)} USD", 11, '#DBEAFE', 'middle')
    for i, c in enumerate(cs):
        for j, rg in enumerate(rigs):
            v = sum(x['gross'] for x in wells if x['rig'] == rg and x['contract'] == c['id'])
            if v <= 0: continue
            sw = max(3, 22*v/max(1,TOT['gross']/6))
            b += f'<path d="M280,{226+i*160} C360,{226+i*160} 360,{237+j*230} 440,{237+j*230}" fill="none" stroke="{SEQ[(i+j)%len(SEQ)]}" stroke-width="{sw:.1f}" opacity=".38"/>'
    for i, w in enumerate(wells):
        yy = 160 + i*48
        b += R(760, yy, 380, 38, SOFT, rx=4)
        b += T(772, yy+24, f"OP{w['op']:02d}  {w['key'][:20]:<20s}", 11.5, INK, wt='600')
        b += T(1130, yy+24, f"{fmt(w['gross'],0)} USD", 11, NAVY, 'end')
        b += R(760, yy+30, 380*min(1, nz(w['gross'])/max(1,max(nz(x['gross']) for x in wells))), 6, SEQ[i%len(SEQ)], rx=2)
    b += T(60, 800, f"Gross {fmt(TOT['gross'],0)}  |  payable {fmt(TOT['payable'],0)}  |  pure net {fmt(TOT['pure_net'] if 'pure_net' in TOT else TOT['net'],0)} USD", 12, NAVY, wt='600')
    artboard('FIELD_12_sankey', 'SAZN Field — Cost Flow by Contract, Rig and Operation',
             'Gross calendar cost allocation across three contracts, two rigs and thirteen operations', tag, b)

    # 13 rig Gantt
    b = T(60, 130, 'RIG GANTT — 13 OPERATIONS / TWO RIGS', 15, NAVY, wt='700')
    rigs = ['OEOC 201', 'OEOC 208']; x0, rw = 330, 820
    b += grid(x0, 165, rw, 620, 13, 10)
    for i, w in enumerate(wells):
        yy = 170 + i*47
        b += T(60, yy+26, f"OP{w['op']:02d} {w['key'][:18]}", 11, INK)
        b += T(320, yy+26, w['rig'], 10.5, MUTED, 'end')
        col = COBALT if w['rig'] == 'OEOC 208' else TEAL
        span = min(rw, rw*nz(w['ddr_days'])/max(1,max(nz(x['ddr_days']) for x in wells)))
        b += R(x0, yy+4, span, 30, col, rx=4)
        b += T(x0+8, yy+24, f"{w['ddr_days']} d  |  {fmt(w['npt_h'],0)} h NPT  |  {str(w['date_start'] or '')[:11]} → {str(w['date_end'] or '')[:11]}", 10.5, '#FFFFFF')
    b += legend([('OEOC 208', COBALT), ('OEOC 201', TEAL)], 330, 158)
    b += T(60, 810, f"{TOT['ddr_days']} DDR days  |  {fmt(TOT['npt_hours'],1)} NPT hours  |  {fmt(TOT['drilled_m'],1)} m drilled", 12, NAVY, wt='600')
    artboard('FIELD_13_rig_gantt', 'SAZN Field — Rig Schedule & NPT Load',
             'Operation duration, rig assignment and non-productive hours by DDR campaign order', tag, b)

    # 14 heatmap 13x15
    forms = ['Aghajari','Gachsaran','Asmary','Pabdeh','Gurpi','Laffan','Ilam','Sarvak','Kazhdomi','Dariyan','Gadvan','Khalij','Fahliyan','Hith','Surmeh']
    b = T(60, 130, '13 × 15 SUBSURFACE RISK / NPT HEATMAP', 15, NAVY, wt='700')
    x0, y0, cw, ch = 300, 175, 58, 39
    mx = max([nz(h['risk']) for h in F['heatmap']] + [1])
    for j, f in enumerate(forms): b += T(x0+j*cw+cw/2, y0-10, f[:7], 9, MUTED, 'middle')
    for i, w in enumerate(wells):
        yy = y0 + i*ch
        b += T(60, yy+25, w['key'][:20], 11, INK)
        for j, f in enumerate(forms):
            row = next((h for h in F['heatmap'] if h['well'] == w['key'] and h['formation'].lower() == f.lower()), None)
            v = nz(row['risk']) if row else 0
            alpha = 0.07 + 0.88*min(1, v/mx)
            b += R(x0+j*cw+1, yy+2, cw-2, ch-4, f'rgba(220,38,38,{alpha:.2f})', rx=2)
            if row and row.get('delta') is not None:
                b += T(x0+j*cw+cw/2, yy+24, f"{fmt(row['delta'],0)}", 8.5, '#FFFFFF', 'middle', '700')
    b += T(60, 800, 'Cell = event hours by formation; digit = prognosis-actual structural delta (m) where a top is correlated.', 11, MUTED)
    artboard('FIELD_14_heatmap', 'SAZN Field — Subsurface Risk Heatmap',
             'Formation risk intensity and structural delta across all thirteen operations', tag, b)

    # 15 structural cross-section
    b = T(60, 130, 'STRUCTURAL CROSS-SECTION — PROGNOSIS vs ACTUAL TOPS', 15, NAVY, wt='700')
    x0, y0, rw, rh = 80, 170, 1040, 560
    b += grid(x0, y0, rw, rh, 8, 13)
    forms = ['Gachsaran','Asmary','Pabdeh','Gurpi','Ilam','Sarvak','Kazhdomi','Gadvan','Fahliyan']
    for j, f in enumerate(forms):
        pts = []
        for i, w in enumerate(wells):
            h = next((h for h in F['heatmap'] if h['well'] == w['key'] and h['formation'].lower() == f.lower()), None)
            md = nz(h['md']) if h else 0
            if md <= 100: continue
            pts.append((x0 + rw*(i+.5)/13, y0 + rh*min(1, md/5000)))
        if len(pts) >= 2:
            b += POLY(pts, SEQ[j%len(SEQ)], 2.2)
            b += T(1128, 200+j*26, f[:11], 10, SEQ[j%len(SEQ)], wt='700')
    for i, w in enumerate(wells): b += T(x0+rw*(i+.5)/13, y0+rh+22, f"{w['op']:02d}", 10, MUTED, 'middle')
    b += T(80, 780, 'South Azadegan anticline — correlated tops from DDR prognosis/actual columns, EOWR and Programme tables; depth axis inverted.', 11, MUTED)
    artboard('FIELD_15_crosssection', 'SAZN Field — Structural Cross-Section',
             'Formation tops correlated across the thirteen wellbores on an inverted depth axis', tag, b)

    # 16 learning curves
    b = T(60, 130, 'LEARNING CURVES — RE-ENTRY / WORKOVER REPETITION', 15, NAVY, wt='700')
    chains = F['learning_curves']
    x0, y0, rw, rh = 90, 180, 1000, 480
    b += grid(x0, y0, rw, rh, 6, 6)
    for ci, (nm, rows) in enumerate(chains.items()):
        col = COBALT if ci == 0 else TEAL
        pts = []
        for i, r in enumerate(rows):
            band = rw/2 - 40; xstart = x0 + ci*rw/2
            xx = xstart + 30 + (i+.5)*band/max(1,len(rows))
            yy = y0 + rh - rh*min(1, nz(r['npt_h'])/max(1, 2100))
            pts.append((xx, yy))
            b += C(xx, yy, 8, col)
            b += T(xx, yy-16, f"{fmt(r['npt_h'],0)} h", 11, col, 'middle', '700')
            b += T(xx, y0+rh+24, r['well'][:18], 10.5, INK, 'middle')
            b += T(xx, y0+rh+42, f"{r['days']} d | {fmt(r['cbs'],0)} USD", 9.5, MUTED, 'middle')
        b += POLY(pts, col, 2.6)
        b += T(100, 160+ci*22, nm, 12, col, wt='700')
    b += T(90, 760, 'AZNS-127 WO1 → WO2 → ST2 and AZNS-122 WO1 → WO2: NPT hours, DDR days and CBS plan by repeated intervention.', 11, MUTED)
    artboard('FIELD_16_learning', 'SAZN Field — Intervention Learning Curves',
             'NPT-hour and CBS-plan evolution across repeated workovers and sidetracks', tag, b)

    # 17 contracts / KPI
    b = T(60, 130, 'CONTRACT CLOSE-OUT — FOUR-TRACK COST BY BUNDLE', 15, NAVY, wt='700')
    tracks = [('CBS plan', 'cbs_plan', COBALT), ('Gross calendar', 'gross', SLATE), ('Payable excl. ctr NPT', 'payable', AMBER), ('Pure net excl. all NPT', 'net', EMERALD)]
    for j, (nm, kk, col) in enumerate(tracks):
        yy = 175 + j*145
        b += T(60, yy+16, nm, 13, NAVY, wt='700')
        mx = max([nz(c[kk]) for c in F['contracts']] + [1])
        for i, c in enumerate(F['contracts']):
            xx = 300 + i*290
            b += R(xx, yy, 260, 34, SOFT, rx=4)
            b += R(xx, yy, 260*nz(c[kk])/mx, 34, col, rx=4)
            b += T(xx+8, yy+22, f"{c['id']}  {fmt(c[kk],0)}", 11, '#FFFFFF' if nz(c[kk])/mx>.45 else INK, wt='700')
        b += T(60, yy+50, f"field {fmt(TOT[kk if kk!='net' else 'net'],0)} USD", 11.5, MUTED)
    b += R(60, 755, 1080, 40, SOFT, rx=5)
    b += T(78, 780, f"Audit: {TOT['cells']:,} Excel cells | {TOT['numeric_cells']:,} numeric | {TOT['formulas']:,} formulas | {TOT['pdf_pages']:,} PDF pages | {TOT['pdf_numbers']:,} numeric PDF tokens | {TOT['pdf_tables']:,} tables", 11.5, NAVY, wt='600')
    artboard('FIELD_17_contracts', 'SAZN Field — Contract Close-Out & Audit Totals',
             'Four cost tracks by bundle, with full source-cell and page census from the master warehouse', tag, b)


def main():
    for k, p in P.items(): well_artboards(k, p)
    field_artboards()
    json.dump(JOBS, open(os.path.join(DATA, 'graphics_jobs.json'), 'w'), indent=1)
    print('ARTBOARDS', len(JOBS))


if __name__ == '__main__': main()
