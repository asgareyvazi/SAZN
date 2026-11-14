"""SAZN reference programme — shared structural metadata and helpers.
All factual numbers come from the 51 source documents; only the campaign map lives here."""
import os, re, json, math, hashlib, unicodedata, zlib, sqlite3
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'SAZN_Reports'); DATA = os.path.join(OUT, '_data')
GFX = os.path.join(OUT, 'SAZN_Graphics'); HTML = os.path.join(DATA, 'artboards')
DB = os.path.join(OUT, 'SAZN_Master_Warehouse.sqlite'); XLSX = os.path.join(OUT, 'SAZN_Master_Audit_Workbook.xlsx')
for d in (OUT, DATA, GFX, HTML): os.makedirs(d, exist_ok=True)
# The warehouse is committed as a gzip archive to stay below the workspace patch cap.
# Restore the working (ignored) SQLite file automatically when only the archive exists.
if not os.path.exists(DB) and os.path.exists(DB + '.gz'):
    import gzip as _gz, shutil as _sh
    with _gz.open(DB + '.gz', 'rb') as _f, open(DB, 'wb') as _o:
        _sh.copyfileobj(_f, _o)

CONTRACTS = {'C1':('CCA/C/1401/27','Six-well bundle',list(range(1,7))),
             'C2':('CCA/C/1403/31','4+1 well bundle / extended scope',list(range(7,13))),
             'C3':('CCA/C/1403/31 / ST2 amendment','Single-well second sidetrack — user-map label; source ST#2 EOWR header reads CCA/C/1405/3',[13])}
W = [
 (1,'AZNS-40','AZNS-40 (F-21)','C1','OEOC 208','Drilling / Re-entry','1.AZNS_40_DDR_Pack.xlsx',['1.AZNS_40_CBS.xlsx','10 CBS matrial.xlsx'],['1.AZNS_40_Well_Program.pdf'],['1.AZNS_40_EOWR.pdf']),
 (2,'AZNS-10','AZNS-10','C1','OEOC 201','Workover','2.AZNS_10_DDR_Pack.xlsx',['2.AZNS_10_CBS.xlsx','10 CBS matrial.xlsx'],['2.AZNS_10_Workover_Program.pdf'],['2.AZNS_10_EOWR.pdf']),
 (3,'AZNS-85','AZNS-85','C1','OEOC 201','Workover','3.AZNS_85_DDR_Pack.xlsx',['3.AZNS_85_CBS.xlsx','10 CBS matrial.xlsx'],['3.AZNS_85_Workover_Program.pdf'],['3.AZNS_85_EOWR.pdf']),
 (4,'AZNS-122-WO1','AZNS-122 WO1','C1','OEOC 201','Workover-1','4.AZNS_122_WO1_DDR_Pack.xlsx',['4.AZNS_122_WO1_CBS.xlsx','10 CBS matrial.xlsx'],['4.AZNS_122_Workover-1_Program.pdf'],[]),
 (5,'AZNS-116','AZNS-116','C1','OEOC 208','Workover / Sidetrack','5.AZNS_116_DDR_Pack.xlsx',['5.6AZNS_127WO1_and_116_CBS.xlsx','10 CBS matrial.xlsx'],['5.AZNS_116_Workover_Program.pdf'],['5.AZNS_116_EOWR.pdf']),
 (6,'AZNS-127-WO1-ST1','AZNS-127 WO1 & ST1','C1','OEOC 208','Workover-1 + Sidetrack-1','6.AZNS_127_WO1_and_ST1_DDR_Pack.xlsx',['5.6AZNS_127WO1_and_116_CBS.xlsx','10 CBS matrial.xlsx'],['6.AZNS_127_Workover1_Program.pdf','6.AZNS_127_Side_Track#01_Program.pdf'],['6.AZNS_127_WO1_EOWR.pdf']),
 (7,'AZNS-207','AZNS-207 (F-20)','C2','OEOC 208','Development drilling','7.AZNS_207_DDR_Pack.xlsx',['7.AZNS_207_CBS.xlsx'],['7.AZNS_207_Well_Program.pdf'],['7.AZNS_207_EOWR.pdf']),
 (8,'AZNS-122-WO2','AZNS-122 WO2','C2','OEOC 208','Workover-2','8.AZNS_122_WO2_DDR_Pack.xlsx',['8.AZNS_122_WO2_CBS.xlsx'],['8.AZNS_122_Workover-2_Program.pdf'],['8.AZNS_122_WO2_EOWR.pdf']),
 (9,'AZNS-75','AZNS-75','C2','OEOC 208','Workover-1 / ST1','9.AZNS_75_DDR_Pack.xlsx',['9.AZNS_75_CBS.xlsx','10 CBS matrial.xlsx'],['9.AZNS_75_Workover-1_Program.pdf'],['9.AZNS_75_EOWR.pdf']),
 (10,'AZNS-12','AZNS-12','C2','OEOC 208','Workover / Sidetrack','10.AZNS_12_DDR_Pack.xlsx',['10.AZNS_12_CBS.xlsx','10 CBS matrial.xlsx'],['10.AZNS_12_Workover_Program.pdf'],['10.AZNS_12_EOWR.pdf']),
 (11,'AZNS-166','AZNS-166','C2','OEOC 201','Workover-1','11.AZNS_166_DDR_Pack.xlsx',['11.AZNS_166_CBS.xlsx'],['11.AZNS_166_Workover-1_Program.pdf'],['11.AZNS_166_EOWR.pdf']),
 (12,'AZNS-127-WO2','AZNS-127 WO2','C2','OEOC 208','Workover-2','12.AZNS_127_WO2_DDR_Pack.xlsx',['12.AZNS_127_WO2_CBS.xlsx'],['12.AZNS_127_Workover#2_Program.pdf'],['12,13.AZNS_127_WO2_and ST2_EOWR.pdf']),
 (13,'AZNS-127-ST2','AZNS-127 ST2','C3','OEOC 208','Side Track-2','13.AZNS_127_ST2_DDR_Pack.xlsx',['13.AZNS_127_ST2_CBS.xlsx'],['13.AZNS_127_Side Track-2_Program.pdf'],['12,13.AZNS_127_WO2_and ST2_EOWR.pdf']),
]
WELLS = [dict(op=a,key=b,title=c,contract=d,contract_code=CONTRACTS[d][0],rig=e,kind=f,ddr=g,cbs=h,program=i,eowr=j) for a,b,c,d,e,f,g,h,i,j in W]
BY_KEY={w['key']:w for w in WELLS}; BY_OP={w['op']:w for w in WELLS}
DDR_FILE={w['ddr']:w['key'] for w in WELLS}
FORMATIONS=['Aghajari','Gachsaran','Asmary','Pabdeh','Gurpi','Laffan','Ilam','Sarvak','Kazhdomi','Gadvan','Khalij','Fahliyan','Hith','Surmeh','Najmeh']
FALIAS={'asmari':'Asmary','kazhdumi':'Kazhdomi','fahlian':'Fahliyan','fahliyan':'Fahliyan','lafan':'Laffan'}
PAL=dict(navy='#0B1F3A',cobalt='#1D4ED8',teal='#0D9488',emerald='#15803D',amber='#D97706',coral='#DC2626',slate='#475569')
SEQ=[PAL['cobalt'],PAL['teal'],PAL['emerald'],PAL['amber'],PAL['coral'],PAL['slate'],'#6366F1','#0891B2','#65A30D','#EA580C','#7C3AED','#0F766E']

def norm(x):
    if x is None: return ''
    s=unicodedata.normalize('NFKC',str(x)).replace('\u200c','').replace('\u200f','').replace('\xa0',' ')
    return re.sub(r'\s+',' ',s).strip()
def isnum(v): return isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(float(v))
def tonum(v):
    if isnum(v): return float(v)
    if isinstance(v,str):
        s=v.strip().replace(',','').replace('$','').replace('%','')
        if re.fullmatch(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?',s): return float(s)
    return None
def firstnum(s):
    m=re.search(r'[-+]?\d[\d,]*\.?\d*',str(s or ''))
    return tonum(m.group(0).replace(',','')) if m else None
def nval(v):
    x=tonum(v)
    return x if x is not None else firstnum(v)
def fmt(v,d=2,dash='—'):
    if v is None or v=='':return dash
    if isnum(v): return f'{float(v):,.{d}f}'
    return str(v)
def r2(v): return None if v is None else round(float(v),2)
def med(xs):
    xs=[x for x in xs if x is not None]
    return sorted(xs)[len(xs)//2] if xs else None
def mean(xs):
    xs=[x for x in xs if x is not None]
    return sum(xs)/len(xs) if xs else None
def cl(v,a,b): return max(a,min(b,v))
def sha(p):
    h=hashlib.sha256()
    with open(p,'rb') as f:
        for b in iter(lambda:f.read(1<<20),b''):h.update(b)
    return h.hexdigest()[:16]
def jdump(o,p):
    with open(p,'w',encoding='utf-8') as f: json.dump(o,f,ensure_ascii=False,separators=(',',':'),default=str)
def jload(p):
    with open(p,'r',encoding='utf-8') as f: return json.load(f)
def z(s): return zlib.compress(s if isinstance(s,bytes) else json.dumps(s,ensure_ascii=False,separators=(',',':')).encode('utf-8'),9)
def unz(b): return zlib.decompress(b).decode('utf-8')
def clean(s):
    s=norm(s).rstrip(':').strip(); s=re.sub(r'\s*\(.*?\)\s*',' ',s)
    return re.sub(r'\s+',' ',s).strip().lower()

def load_json(name):
    """Read a derived JSON, preferring the gzip-compressed committed form."""
    import gzip as _gz
    plain=os.path.join(DATA,name+'.json'); packed=plain+'.gz'
    if os.path.exists(packed):
        with _gz.open(packed,'rt',encoding='utf-8') as f: return json.load(f)
    with open(plain,encoding='utf-8') as f: return json.load(f)
def jdump_gz(obj,path):
    import gzip as _gz
    with _gz.open(path,'wt',encoding='utf-8',compresslevel=9) as f: json.dump(obj,f,ensure_ascii=False)
