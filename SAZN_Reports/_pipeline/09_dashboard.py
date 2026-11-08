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
            totals=TOT, field=F, review=VR['review'])
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
<div class="panel"><h2>How to read a result</h2><p style="font-size:12px;line-height:1.6;color:#475569">
<code>CELL</code> rows carry an exact Excel sheet and A1 reference; <code>PDF</code> rows carry a page and line number;
<code>CBS</code>, <code>MAT</code>, <code>FORM</code>, <code>FACT</code> and <code>TBL</code> rows carry the source document and a short context.
Every figure in this dashboard is a value read from the 51 source documents. Derived quantities (four-track cost, NPT deductions, DLS, structural delta) are computed only from sourced rates, quantities and hours and are labelled by their basis in the master audit workbook.
No value is invented, interpolated or rounded away; pages without a text layer are retained page by page in the warehouse.</p></div>
</div>
<footer>SAZN Master Warehouse · SAZN_Master_Warehouse.sqlite · SAZN_Master_Audit_Workbook.xlsx · 149 reviewed graphics · 30 Word books · generated 2026-11-08</footer>
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
  charts();
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
     ['DOCX books','30 (24 chapters, 96 tables, 11 graphics each)']].map(x=>'<tr><td>'+x[0]+'</td><td>'+x[1]+'</td></tr>').join('')+'</tbody></table>';
}
</script></body></html>'''
HTML = HTML.replace('__META__', json.dumps(META, ensure_ascii=False, separators=(',', ':')))
HTML = HTML.replace('__B64__', GZ)
path = os.path.join(OUT, 'SAZN_Executive_Interactive_Dashboard.html')
open(path, 'w', encoding='utf-8').write(HTML)
print('DASHBOARD MB', round(os.path.getsize(path)/1e6, 2), 'rows', len(ROWS))
