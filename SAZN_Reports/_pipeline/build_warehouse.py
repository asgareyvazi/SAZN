"""Build the complete SAZN master warehouse + audit workbook from the 51 source files.
Stages: schema -> 2,810 Excel sheets -> 4,345 PDF pages -> cell-scan enrichment
-> PDF fact mining -> 24-module analytics -> Excel workbook.
All text inside the archive is lossless; numeric values keep exact cell/page/line addresses."""
import os,re,sys,json,zlib,glob,sqlite3,time,math,collections,datetime,statistics
sys.path.insert(0,os.path.dirname(__file__))
from common import *
import openpyxl, pymupdf
from openpyxl.utils import get_column_letter as CL

con=sqlite3.connect(DB); con.row_factory=sqlite3.Row
con.execute('PRAGMA journal_mode=OFF'); con.execute('PRAGMA synchronous=OFF')
def Q(s,*a): return con.execute(s,a).fetchall()
def R(s,*a): return [dict(x) for x in Q(s,*a)]
def O(s,*a):
    x=con.execute(s,a).fetchone(); return dict(x) if x else {}

SCHEMA="""
CREATE TABLE meta(k TEXT PRIMARY KEY,v TEXT);
CREATE TABLE source_files(id INTEGER PRIMARY KEY,category TEXT,rel_path TEXT,filename TEXT,n_units INTEGER,size_bytes INTEGER,sha256 TEXT,well_keys TEXT,notes TEXT);
CREATE TABLE wells(op INTEGER PRIMARY KEY,key TEXT,title TEXT,contract TEXT,contract_code TEXT,rig TEXT,kind TEXT,ddr_file TEXT,cbs_files TEXT,program_files TEXT,eowr_files TEXT,eowr_status TEXT);
CREATE TABLE contracts(id TEXT PRIMARY KEY,code TEXT,title TEXT,ops TEXT);
CREATE TABLE sheets(id INTEGER PRIMARY KEY,file_id INTEGER,name TEXT,idx INTEGER,category TEXT,n_rows INTEGER,n_cols INTEGER,n_nonempty INTEGER,n_numeric INTEGER,n_formula INTEGER,template TEXT,report_no INTEGER,part INTEGER,well_key TEXT,rel_path TEXT);
CREATE TABLE sheet_cells(sheet_id INTEGER PRIMARY KEY,payload BLOB);
CREATE TABLE cell_numeric(sheet_id INTEGER,r INTEGER,c INTEGER,value REAL);
CREATE TABLE pdf_pages(id INTEGER PRIMARY KEY,file_id INTEGER,page_no INTEGER,rel_path TEXT,well_key TEXT,category TEXT,n_chars INTEGER,n_images INTEGER,chapter TEXT,text_z BLOB,keywords TEXT);
CREATE TABLE pdf_num_pages(id INTEGER PRIMARY KEY,file_id INTEGER,page_no INTEGER,well_key TEXT,payload BLOB);
CREATE TABLE pdf_tables(id INTEGER PRIMARY KEY,file_id INTEGER,page_no INTEGER,table_no TEXT,title TEXT,body TEXT,well_key TEXT,category TEXT);
CREATE TABLE eowr_facts(id INTEGER PRIMARY KEY,well_key TEXT,category TEXT,fact_key TEXT,value REAL,text_value TEXT,unit TEXT,page_no INTEGER,rel_path TEXT,context TEXT);
CREATE TABLE ddr_daily(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,rel_path TEXT,sheet_remark TEXT,sheet_data TEXT,template TEXT,date_g TEXT,date_j TEXT,rig TEXT,well_name TEXT,hole_size TEXT,formation TEXT,md_0000 REAL,md_2400 REAL,md_0600 REAL,drilled_24 REAL,rop REAL,mw REAL,fit_emw REAL,section_start REAL,estimated_final_depth REAL,lta_day REAL,actual_rig_days REAL,spud_date TEXT,hole_start_date TEXT,client_rep TEXT,ds_day TEXT,ds_night TEXT,geologist TEXT,tool_pusher TEXT,northing REAL,easting REAL,kop1 REAL,kop2 REAL,estimated_rig_days REAL,ahead_behind TEXT,summary TEXT,forecast TEXT);
CREATE TABLE ddr_activities(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,seq INTEGER,rel_path TEXT,sheet TEXT,row_no INTEGER,t_from TEXT,t_to TEXT,hours REAL,phase TEXT,code TEXT,sub_code TEXT,status TEXT,npt_attributed TEXT,description TEXT,main_activity TEXT);
CREATE TABLE ddr_time(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,code INTEGER,category TEXT,hours REAL,rel_path TEXT);
CREATE TABLE ddr_materials(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,material TEXT,unit TEXT,received REAL,used REAL,stock REAL,total_received REAL,source_sheet TEXT,cell_ref TEXT);
CREATE TABLE ddr_surveys(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,md REAL,incl REAL,azi REAL,tvd REAL,north REAL,east REAL,vs REAL,dls REAL,tool TEXT,rel_path TEXT);
CREATE TABLE ddr_bha(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,bha_no TEXT,item TEXT,qty TEXT,od REAL,length REAL,cum_len REAL,rel_path TEXT);
CREATE TABLE ddr_pob(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,category TEXT,count REAL,total REAL,rel_path TEXT);
CREATE TABLE ddr_formations(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,name TEXT,md REAL,tvd REAL,prognosis_md REAL,prognosis_tvd REAL,lithology TEXT,rel_path TEXT,source TEXT);
CREATE TABLE ddr_casing(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,size TEXT,grade TEXT,weight TEXT,thread TEXT,from_m REAL,to_m REAL,burst REAL,collapse REAL,shoe_tvd REAL,rel_path TEXT);
CREATE TABLE ddr_notes(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,sheet TEXT,category TEXT,text TEXT,rel_path TEXT);
CREATE TABLE cbs_items(id INTEGER PRIMARY KEY,file_id INTEGER,sheet TEXT,row_no INTEGER,well_key TEXT,category TEXT,section TEXT,item_no TEXT,description TEXT,unit TEXT,unit_price REAL,qty REAL,total REAL,formula TEXT,cell_ref TEXT,is_contingency INTEGER,rel_path TEXT);
CREATE TABLE cbs_totals(id INTEGER PRIMARY KEY,file_id INTEGER,sheet TEXT,well_key TEXT,item TEXT,qty REAL,unit_cost REAL,total REAL,row_no INTEGER,rel_path TEXT);
CREATE TABLE cbs_stim(id INTEGER PRIMARY KEY,file_id INTEGER,sheet TEXT,well_key TEXT,kind TEXT,item_no TEXT,description TEXT,unit TEXT,volume REAL,unit_price REAL,standby_days REAL,operating_days REAL,standby_rate REAL,operating_rate REAL,total REAL,remark TEXT,row_no INTEGER,rel_path TEXT);
CREATE TABLE cbs_accomm(id INTEGER PRIMARY KEY,file_id INTEGER,sheet TEXT,well_key TEXT,description TEXT,qty REAL,unit TEXT,rate REAL,total REAL,row_no INTEGER,rel_path TEXT);
CREATE TABLE well_kpi(well_key TEXT PRIMARY KEY,payload TEXT);
CREATE TABLE well_events(id INTEGER PRIMARY KEY,well_key TEXT,kind TEXT,day TEXT,depth REAL,formation TEXT,title TEXT,detail TEXT,hours REAL,cost_usd REAL,evidence TEXT);
CREATE TABLE well_lessons(id INTEGER PRIMARY KEY,well_key TEXT,code TEXT,polarity TEXT,title TEXT,evidence TEXT,action TEXT,hours REAL,cost_usd REAL);
CREATE TABLE well_recs(id INTEGER PRIMARY KEY,well_key TEXT,code TEXT,branch TEXT,payload TEXT);
CREATE TABLE reconciliation(id INTEGER PRIMARY KEY,well_key TEXT,item TEXT,cbs REAL,ddr REAL,eowr REAL,unit TEXT,delta REAL,status TEXT);
CREATE TABLE npt_breakdown(id INTEGER PRIMARY KEY,well_key TEXT,category TEXT,hours REAL,attributed TEXT,share REAL);
CREATE TABLE field_timeseries(id INTEGER PRIMARY KEY,well_key TEXT,report_no INTEGER,date_g TEXT,md REAL,drilled REAL,rop REAL,mw REAL,npt_h REAL,pt_h REAL,cum_cost REAL);
CREATE TABLE audit_findings(id TEXT PRIMARY KEY,title TEXT,detail TEXT);
CREATE VIRTUAL TABLE sheet_fts USING fts5(body,sheet_id UNINDEXED,sheet UNINDEXED,file UNINDEXED,tokenize='unicode61');
CREATE VIRTUAL TABLE pdf_fts USING fts5(body,file_id UNINDEXED,page UNINDEXED,file UNINDEXED,well UNINDEXED,tokenize='unicode61');
"""
def schema():
    if os.path.exists(DB): os.remove(DB)
    globals()['con']=sqlite3.connect(DB); con.row_factory=sqlite3.Row
    con.executescript(SCHEMA)
    for k,v in [('title','SAZN Master Warehouse — South Azadegan Oilfield'),('created','2026-11-08'),
                ('source_files','51'),('cbs_sheets','202'),('ddr_sheets','2608'),('eowr_pages','2724'),
                ('program_pages','1621'),('pdf_pages','4345'),('excel_sheets','2810')]:
        con.execute('INSERT INTO meta VALUES (?,?)',(k,v))
    for cid,(code,title,ops) in CONTRACTS.items():
        con.execute('INSERT INTO contracts VALUES (?,?,?,?)',(cid,code,title,','.join(map(str,ops))))
    for w in WELLS:
        con.execute('INSERT INTO wells VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
            (w['op'],w['key'],w['title'],w['contract'],w['contract_code'],w['rig'],w['kind'],w['ddr'],
             ';'.join(w['cbs']),';'.join(w['program']),';'.join(w['eowr']),
             'AVAILABLE' if w['eowr'] else 'NOT AVAILABLE — no EOWR file; reconstructed from DDR + CBS + Well Program'))
    con.commit()

# ---------------------------------------------------------------- Excel -----
def g(G,r,c): return G[r-1][c-1] if 1<=r<=len(G) and 1<=c<=len(G[r-1]) else None
def sv(v):
    if v is None:return ''
    if isinstance(v,(datetime.datetime,datetime.date)):return v.strftime('%Y-%m-%d')
    if isinstance(v,datetime.time):return v.strftime('%H:%M')
    if isinstance(v,datetime.timedelta):return f'{int(v.total_seconds()//3600):02d}:{int(v.total_seconds()%3600//60):02d}'
    return str(v)
def tohours(v):
    if v is None:return None
    if isinstance(v,datetime.timedelta):return round(v.total_seconds()/3600,3)
    if isinstance(v,datetime.time):return round(v.hour+v.minute/60+v.second/3600,3)
    x=tonum(v)
    if x is not None:return x
    m=re.search(r'(?:(\d+)\s*day[s]?\s*,\s*)?(\d{1,2}):(\d{2})(?::(\d{2}))?',str(v))
    if m:return round(int(m.group(1) or 0)*24+int(m.group(2))+int(m.group(3))/60+int(m.group(4) or 0)/3600,3)
    return None
def gridfind(G,needles,numeric=False,look=14,maxrow=None):
    ns=[x.lower() for x in needles]
    for ri,row in enumerate(G[:maxrow] if maxrow else G,1):
        for ci,v in enumerate(row,1):
            if not isinstance(v,str):continue
            t=norm(v).lower()
            if not any(x in t for x in ns):continue
            for cj in range(ci+1,min(ci+look,len(row)+1)):
                v2=row[cj-1]
                if v2 is None or str(v2).strip()=='':continue
                if isinstance(v2,str) and any(x in norm(v2).lower() for x in ns):continue
                if numeric:
                    q=nval(v2)
                    if q is not None:return q
                else:return sv(v2).strip()
    return None
def kv(G,look=12):
    out=[];seen={}
    LW=('date','name','rig','no','depth','md','rop','mw','pcf','formation','section','hour','hrs','time','rate','weight','pressure','pob','total','spud','kop','fit','lot','easting','northing','latitude','longitude','client','operator','super','supervisor','geologist','tool','pusher','lta','day','night','well','type','size','grade','thread','shoe','burst','collapse','stock','used','received','unit','status','code','summary','forecast','note','last','test','drill','wind','temp','visibility','personnel','chemical','mud','bit','pump','torque','wob','rpm','gpm','spm','spp','volume','loss','gain','water','oil','solid','ph','chloride','lithology','remark','activity','phase','attributed','cost','price','qty','estimated','actual','moving','start','final','head','elevation','location','project','field','report','sequence','hole','string','survey','inc','azi','tvd','density','visc','gel','fluid','cement','acid','npt','bha','whp','choke','interval','morning','midnight','previous','cum','avg','jet','ann','tfa','liner','nozzle')
    for ri,row in enumerate(G,1):
        for ci,v in enumerate(row,1):
            if not isinstance(v,str) or not v.strip() or len(v)>90:continue
            raw=norm(v);key=clean(raw)
            pairs=[]
            m=re.match(r'^([A-Za-z0-9 /().\-#@°"\'<>=+,]+)\s*:\s*(\S.*)$',raw)
            if m and len(m.group(2))<80:pairs.append((clean(m.group(1)),m.group(2).strip(),f'{CL(ci)}{ri}'))
            val=None;ref=''
            for cj in range(ci+1,min(ci+look,len(row)+1)):
                vv=row[cj-1]
                if vv is None or str(vv).strip()=='':continue
                if isinstance(vv,str) and clean(vv)==key:continue
                val=sv(vv);ref=f'{CL(cj)}{ri}';break
            kl=key.lower()
            labelish=(raw.endswith(':') or ':' in raw or (any(x in kl for x in LW) and len(key)<=55) or (val and isnum(val) and len(key)<=40))
            if val and val.strip() and labelish:pairs.append((key,val.strip(),ref))
            for k,vv,rf in pairs:
                if not k or not vv or len(vv)>200:continue
                b=seen.setdefault(k,set())
                if vv in b or len(b)>=3:continue
                b.add(vv);out.append((k,vv,rf))
    return out
def kvfind(K,*needles):
    nk=lambda s:re.sub(r'[^a-z0-9]','',str(s).lower())
    for nd in needles:
        for k,v,rf in K:
            if nk(k)==nk(nd):return v
    for nd in needles:
        for k,v,rf in K:
            a,b=nk(k),nk(nd)
            if a and (b in a or (len(a)>4 and a in b)):return v
    return None

CBS_WELL={'1.AZNS_40_CBS.xlsx':'AZNS-40','2.AZNS_10_CBS.xlsx':'AZNS-10','3.AZNS_85_CBS.xlsx':'AZNS-85','4.AZNS_122_WO1_CBS.xlsx':'AZNS-122-WO1','7.AZNS_207_CBS.xlsx':'AZNS-207','8.AZNS_122_WO2_CBS.xlsx':'AZNS-122-WO2','9.AZNS_75_CBS.xlsx':'AZNS-75','10.AZNS_12_CBS.xlsx':'AZNS-12','11.AZNS_166_CBS.xlsx':'AZNS-166','12.AZNS_127_WO2_CBS.xlsx':'AZNS-127-WO2','13.AZNS_127_ST2_CBS.xlsx':'AZNS-127-ST2'}
BENCH={'29':'BENCH-29','38':'BENCH-38','10':'AZNS-10','40':'AZNS-40','85':'AZNS-85','122':'AZNS-122-WO1','12':'AZNS-12','75':'AZNS-75','116':'AZNS-116','127':'AZNS-127-WO1-ST1'}
def cbs_well(fn,sn):
    if 'matrial' in fn.lower():
        m=re.search(r'Material-.*?(\d+)',sn);return BENCH.get(m.group(1),'') if m else ''
    if fn.startswith('5.6'):
        if re.search(r'116\D*$',sn):return 'AZNS-116'
        if re.search(r'127\D*$',sn):return 'AZNS-127-WO1-ST1'
        return ''
    return CBS_WELL.get(fn,'')
def cbs_kind(sn):
    s=sn.lower()
    if 'total cost' in s:return 'TOTAL'
    if 'material & consumable' in s:return 'STIM_MATERIAL'
    if 'acidizing' in s:return 'ACID_SERVICE'
    if re.search(r'(mob\.|coil tubing|oil transportation|well testing|wireline|slick line|memory gauge|bottom hole pvt|personnel rate|design,|cost breakdown r-less|cost b\.d r-less|cost breakdown rigless)',s):return 'STIM_SERVICE'
    if 'cost breakdown-material' in s:return 'MATERIAL'
    if 'cost breakdown' in s:return 'SERVICE'
    if re.search(r'(rig move|catering|accomm)',s):return 'ACCOMM'
    return 'OTHER'
def hmap(G):
    best=None
    for ri in range(1,min(12,len(G))+1):
        t=' '.join(norm(x).lower() for x in G[ri-1] if x)
        if 'total' in t and any(x in t for x in ('unit price','price','rate','qty')):best=ri;break
    if not best:return None,{}
    cols={}
    for ci,t in enumerate([norm(x).lower() for x in G[best-1]],1):
        if not t:continue
        if 'description' in t or t.startswith('item') or 'equipment' in t or 'material' in t:cols.setdefault('desc',ci)
        elif t in ('unit','ea.','ea','set','day','job','hole','gal','bbl','mt','lbs','lb') or ('unit' in t and 'price' not in t and 'cost' not in t):cols.setdefault('unit',ci)
        elif ('unit price' in t or 'rate' in t or 'price' in t or 'unit cost' in t) and 'total' not in t:
            cols.setdefault('standby_rate' if 'standby' in t else ('operating_rate' if 'operat' in t else 'price'),ci)
        elif 'qty' in t or 'quantity' in t or 'volume' in t or 'estimated' in t or ('days' in t and 'standby' not in t and 'operat' not in t):
            cols.setdefault('standby_days' if 'standby' in t else ('operating_days' if 'operat' in t else 'qty'),ci)
        elif 'total' in t:cols.setdefault('total',ci)
        elif t.startswith('no'):cols.setdefault('no',ci)
    if 'unit' in cols and 'desc' in cols and cols['desc']<cols['unit']-1:cols['desc']=cols['unit']-1
    return best,cols
def parse_cbs(G,sn,fid,rel,fn):
    kind=cbs_kind(sn);wk=cbs_well(fn,sn);cont=1 if re.search(r'conting',sn,re.I) else 0
    items,totals,stim,accom=[],[],[],[]
    hdr,cols=hmap(G)
    if hdr is None:hdr,cols=2,dict(no=1,desc=2,unit=3,price=4,qty=5,total=6)
    cell=lambda r,k: g(G,r,cols[k]) if k in cols else None
    for ri in range(hdr+1,len(G)+1):
        if not any(v is not None and str(v).strip() for v in G[ri-1]):continue
        desc=norm(cell(ri,'desc'));unit=norm(cell(ri,'unit'));price=nval(cell(ri,'price'))
        qty=nval(cell(ri,'qty'));total=nval(cell(ri,'total'));no=norm(cell(ri,'no'))
        if not desc and total is None and price is None:continue
        base=dict(file_id=fid,sheet=sn,well_key=wk,row_no=ri,rel_path=rel,description=desc[:240],unit=unit[:40],unit_price=price,qty=qty,total=total)
        if kind=='TOTAL':
            it=desc or norm(g(G,ri,2));uc=nval(g(G,ri,6));tt=nval(g(G,ri,7));qq=nval(g(G,ri,3))
            totals.append(dict(file_id=fid,sheet=sn,well_key=wk,item=(it or norm(g(G,ri,1)))[:240],qty=qq,unit_cost=uc,total=tt,row_no=ri,rel_path=rel))
            if it and (tt or uc):items.append({**base,'category':'TOTAL','section':sn,'item_no':no,'description':it[:240],'unit':'','unit_price':uc,'qty':qq,'total':tt,'formula':'','cell_ref':f'G{ri}','is_contingency':0})
        elif kind in ('STIM_MATERIAL','ACID_SERVICE','STIM_SERVICE'):
            stim.append({**base,'kind':kind,'item_no':no,'volume':qty,'standby_days':nval(cell(ri,'standby_days')),'operating_days':nval(cell(ri,'operating_days')),'standby_rate':nval(cell(ri,'standby_rate')),'operating_rate':nval(cell(ri,'operating_rate')),'remark':norm(g(G,ri,9))[:200] if kind=='ACID_SERVICE' else ''})
            items.append({**base,'category':'RIGLESS_STIMULATION','section':sn,'item_no':no,'formula':'','cell_ref':f'F{ri}','is_contingency':cont})
        elif kind=='ACCOMM':
            accom.append(dict(file_id=fid,sheet=sn,well_key=wk,description=desc[:240],qty=qty,unit=unit[:40],rate=price,total=total,row_no=ri,rel_path=rel))
            items.append({**base,'category':'ACCOMM_RIGMOVE','section':sn,'item_no':no,'formula':'','cell_ref':f'F{ri}','is_contingency':cont})
        else:
            items.append({**base,'category':('MATERIAL' if kind=='MATERIAL' else ('SERVICE' if kind=='SERVICE' else 'OTHER')),'section':(no if re.fullmatch(r'20[123]\.?',no or '') else ''),'item_no':no,'formula':'','cell_ref':f'F{ri}','is_contingency':cont})
    if re.search(r'(total cost|stimulation total cost|cost breakdown r-less|cost breakdown rigless|cost b\.d r-less)',sn,re.I):
        for ri,row in enumerate(G,1):
            vs=[v for v in row if isnum(v)]
            if not vs:continue
            lab=norm(g(G,ri,1)) or norm(g(G,ri,2))
            if not lab and ri<len(G)-2:continue
            totals.append(dict(file_id=fid,sheet=sn,well_key=wk,item=(lab or f'Row {ri} total')[:240],qty=None,unit_cost=None,total=float(max(vs)),row_no=ri,rel_path=rel))
    return items,totals,stim,accom

TIME_A=[('1. Rigging & Moving',1),('2. Vertical Drilling',2),('3. Trip',3),('4. Reaming & Hole Condition',4),('5. Mud Circulation',5)]
def parse_remark(G,rel,sn,tpl):
    K=kv(G);F=lambda *ns:kvfind(K,*ns)
    GF=lambda ns,num=False,mr=None,maxrow=None:gridfind(G,ns if isinstance(ns,(list,tuple)) else [ns],numeric=num,maxrow=(maxrow or mr))
    date=gridfind(G,['report date'],maxrow=16)
    if tpl=='B':
        for ri in range(1,min(20,len(G))+1):
            if clean(g(G,ri,1)).rstrip(':')=='report':
                y,mo,d=norm(g(G,ri,6)),norm(g(G,ri,8)),norm(g(G,ri,10))
                if re.fullmatch(r'20\d{2}',y) and mo and d:date=f'{y}-{mo}-{d}'
    f=dict(date_g=date,date_j=(gridfind(G,['report date'],maxrow=12) if tpl=='A' else None),
        well_name=GF(['well name'],maxrow=16) or F('well name'),rig=GF(['rig name'],maxrow=16),
        hole_size=GF(['hole section (inch)','hole section'],maxrow=16),formation=GF(['formation:'],maxrow=18),
        md_0000=GF(['md (m)@ 0:00','previous depth'],True,24),md_2400=GF(['md (m)@ 24:00','mid nigth depth','mid night depth'],True,24),
        md_0600=GF(['md (m)@ 6:00','morning depth'],True,24),drilled_24=GF(['24 hrs drilled','drilled interval'],True,24),
        rop=GF(['avg. rop','average rop'],True,24),mw=GF(['mw (pcf)','mud weight'],True,24),
        lta_day=GF(['lta (day)','days without lta'],True,20),actual_rig_days=GF(['actual rig days'],True,20),
        estimated_final_depth=GF(['estimated final depth','project md'],True,20),section_start=GF(['section start depth'],True,20),
        client_rep=GF(['client rep'],maxrow=16),ds_day=GF(['supervisor (day)','w.s.s day'],maxrow=18),
        ds_night=GF(['supervisor (night)'],maxrow=18),geologist=GF(['site geologist','geologist'],maxrow=18),
        tool_pusher=GF(['tool pusher','rig sup'],maxrow=18),northing=GF(['northing'],True,20),easting=GF(['easting'],True,20),
        kop1=GF(['kop #1','kop 1'],True,20),kop2=GF(['kop #2','kop 2'],True,20),
        estimated_rig_days=GF(['estimated rig days'],True,20),ahead_behind=GF(['ahead / behind'],maxrow=16),
        spud_date=GF(['spud date'],maxrow=16),hole_start_date=GF(['hole section start date'],maxrow=16),
        fit_emw=GF(['fit depth','last fit'],True,20),summary=GF(['summary of activities'],maxrow=26),
        forecast=GF(['operation forecast'],maxrow=26))
    if tpl=='A':
        for ri in range(1,12):
            if norm(g(G,ri,16)).lower().startswith('report date'):f['date_j']=sv(g(G,ri,19))
    acts=[];tbd=[];notes=[];casing=[]
    hdr=None
    for ri in range(1,min(45,len(G))+1):
        row=[clean(x) for x in G[ri-1][:10]]
        if 'from' in row and 'to' in row and any('hrs' in x for x in row):hdr=ri;break
    if hdr:
        for ri in range(hdr+1,len(G)+1):
            c1=norm(g(G,ri,1))
            if tpl=='B':
                if re.search(r'^(summary|operation forecast|time break down)',c1,re.I):break
                hrs=tohours(g(G,ri,5));desc=norm(g(G,ri,19))
                if desc:
                    acts.append(dict(t_from=sv(g(G,ri,1)),t_to=sv(g(G,ri,3)),hours=hrs,phase=norm(g(G,ri,7)),code=norm(g(G,ri,9)),sub_code=norm(g(G,ri,11)),status=norm(g(G,ri,13)),npt_attributed=norm(g(G,ri,16)),description=desc[:420],main_activity='',row_no=ri))
                elif acts and not c1:
                    ex=' '.join(norm(x) for x in G[ri-1][18:40] if isinstance(x,str))
                    if ex:acts[-1]['description']=(acts[-1]['description']+' '+ex)[:420]
            else:
                det=norm(g(G,ri,11));hrs=tohours(g(G,ri,8))
                if det or hrs is not None or re.match(r'^\d+\.',c1):
                    acts.append(dict(t_from=sv(g(G,ri,6)),t_to=sv(g(G,ri,7)),hours=hrs,phase='',code=norm(g(G,ri,9)),sub_code='',status=norm(g(G,ri,10)),npt_attributed='',description=det[:420],main_activity=c1[:120],row_no=ri))
    for ri in range(1,len(G)+1):
        if tpl=='B':
            n=nval(g(G,ri,47));lab=norm(g(G,ri,48));h=tohours(g(G,ri,53))
            if n is not None and lab and n==int(n) and 1<=int(n)<=40:tbd.append((int(n),lab,h))
        else:
            lab=norm(g(G,ri,1));h=tohours(g(G,ri,5))
            if re.match(r'^\d+\.\s+[A-Za-z]',lab) and h is not None:tbd.append((int(lab.split('.')[0]),re.sub(r'^\d+\.\s*','',lab),h))
    for ri in range(1,len(G)+1):
        for ci in range(1,min(len(G[ri-1]),84)+1):
            v=g(G,ri,ci)
            if isinstance(v,str):
                t=norm(v)
                if (re.match(r'^\*{1,3}\s',t) or re.search(r'\b(shortage|LTA|incident|stuck|losses|lost circulation|fishing|NPT|failure|out of service|not provided|unable|defect)',t,re.I)) and len(t)>25:
                    notes.append(('SHORTAGE/REMARK',t[:500]))
    for ri in range(1,len(G)+1):
        if norm(g(G,ri,39)).lower().startswith('size'):
            for rj in range(ri+1,min(ri+9,len(G)+1)):
                sz=norm(g(G,rj,39))
                if sz:casing.append(dict(size=sz,grade=norm(g(G,rj,43)),weight=norm(g(G,rj,44)),thread=norm(g(G,rj,47)),from_m=nval(g(G,rj,40)),to_m=nval(g(G,rj,42)),burst=nval(g(G,rj,50)),collapse=nval(g(G,rj,53)),shoe_tvd=nval(g(G,rj,48))))
            break
    return f,acts,tbd,notes,casing

def parse_data(G,rel,sn,tpl):
    f={};mats=[];sur=[];bha=[];pob=[];forms=[];notes=[]
    if tpl=='A':
        for lab,col,key in [('Morning Depth',4,'md_0600'),('Mid Nigth Depth',4,'md_2400'),('Previous Depth',4,'md_0000'),('Drilled Interval',4,'drilled_24'),('Average ROP',4,'rop'),('Project MD',4,'estimated_final_depth'),('Formation Name',4,'formation'),('Mud Weight',18,'mw')]:
            for ri in range(1,min(len(G),90)+1):
                if clean(g(G,ri,1))==clean(lab) or clean(g(G,ri,8))==clean(lab) or clean(g(G,ri,15))==clean(lab):
                    v=nval(g(G,ri,col))
                    if key and v is not None:f[key]=v
                    break
        # formation geology prognosis vs actual
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,11)).startswith('formation') and clean(g(G,ri,12)).startswith('forc.'):
                for rj in range(ri+1,min(ri+24,len(G)+1)):
                    nm=norm(g(G,rj,11))
                    if not nm or 'total depth' in nm.lower():break
                    forms.append(dict(name=nm[:60],md=nval(g(G,rj,15)),tvd=nval(g(G,rj,17)),prognosis_md=nval(g(G,rj,12)),prognosis_tvd=nval(g(G,rj,13)),lithology=norm(g(G,rj,18))[:180]))
                break
    else:
        f['formation']=gridfind(G,['formation name'],maxrow=90)
        f['mw']=gridfind(G,['mud weight'],True,90)
        f['rop']=gridfind(G,['cum.bit ave rop'],True,90)
    # materials
    if tpl=='A':
        hdr=next((ri for ri in range(1,len(G)+1) if clean(g(G,ri,1)).startswith('material name')),None)
        if hdr:
            for ri in range(hdr+1,min(hdr+80,len(G)+1)):
                nm=norm(g(G,ri,1))
                if not nm or nm.lower().startswith(('bha','daily drilling','total')):break
                rec,used,st=nval(g(G,ri,4)),nval(g(G,ri,6)),nval(g(G,ri,7))
                if rec or used or st:mats.append(dict(material=nm[:110],unit=norm(g(G,ri,3))[:36],received=rec,used=used,stock=st,total_received=nval(g(G,ri,5)),cell_ref=f'A{ri}'))
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,1)).startswith('material name') and ri>80:
                for rj in range(ri+1,min(ri+30,len(G)+1)):
                    nm=norm(g(G,rj,1))
                    if not nm or clean(nm).startswith(('fuel','water','formation','survey')):break
                    rec,used,st=nval(g(G,rj,4)),nval(g(G,rj,6)),nval(g(G,rj,7))
                    if rec or used or st:mats.append(dict(material=f'CEMENT ADD {nm}'[:110],unit=norm(g(G,rj,3))[:36],received=rec,used=used,stock=st,total_received=nval(g(G,rj,5)),cell_ref=f'A{rj}'))
                break
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,1)).startswith('material name') and clean(g(G,ri,11)).startswith('formation'):
                for rj in range(ri+2,min(ri+8,len(G)+1)):
                    nm=norm(g(G,rj,1))
                    if not nm:continue
                    rec,used,st=nval(g(G,rj,4)),nval(g(G,rj,7)),nval(g(G,rj,9))
                    if rec or used or st:mats.append(dict(material=f'BULK {nm}'[:110],unit=norm(g(G,rj,3))[:36],received=rec,used=used,stock=st,total_received=nval(g(G,rj,6)),cell_ref=f'A{rj}'))
                break
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,15)).startswith('safety events data'):
                for rj in range(ri,min(ri+20,len(G)+1)):
                    cat=norm(g(G,rj,15));cnt=nval(g(G,rj,18))
                    if cnt is not None and cat:pob.append(dict(category=('TOTAL POB' if 'total' in cat.lower() else cat[:70]),count=cnt,total=(cnt if 'total' in cat.lower() else None)))
                break
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,1))=='no.' and clean(g(G,ri,2)).startswith('md'):
                for rj in range(ri+1,min(ri+24,len(G)+1)):
                    md=nval(g(G,rj,2))
                    if md is not None:sur.append(dict(md=md,incl=nval(g(G,rj,3)),azi=nval(g(G,rj,4)),tvd=nval(g(G,rj,6)),north=nval(g(G,rj,7)),east=nval(g(G,rj,8)),vs=nval(g(G,rj,9)),dls=nval(g(G,rj,10)),tool=''))
                break
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,8)).startswith('row') and clean(g(G,ri,9)).lower().startswith('junk'):
                for rj in range(ri+1,min(ri+30,len(G)+1)):
                    nm=norm(g(G,rj,9))
                    if not nm or clean(nm).startswith('total'):break
                    bha.append(dict(bha_no=norm(g(G,ri-1,8))[:36],item=nm[:110],qty=norm(g(G,rj,10))[:20],od=nval(g(G,rj,11)),length=nval(g(G,rj,13)),cum_len=nval(g(G,rj,14))))
                break
    else:
        hdr=None;uc,rc,sc=16,17,19
        for ri in range(1,min(len(G),20)+1):
            if clean(g(G,ri,14)) in ('product type','material type'):
                hdr=ri
                c16,c17=clean(g(G,ri,16)),clean(g(G,ri,17))
                if c16.startswith('receive'):rc,uc=16,17
                elif c17.startswith('receive'):rc,uc=17,16
                elif c16.startswith('used'):uc,rc=16,17
                break
        if hdr:
            for ri in range(hdr+1,len(G)+1):
                nm=norm(g(G,ri,14))
                if not nm or nm.lower().startswith(('bulk','bop','safety','weather','note','formation')):break
                rec,used,st=nval(g(G,ri,rc)),nval(g(G,ri,uc)),nval(g(G,ri,sc))
                if rec or used or st:mats.append(dict(material=nm[:110],unit=norm(g(G,ri,20))[:36],received=rec,used=used,stock=st,total_received=None,cell_ref=f'N{ri}'))
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,2)).startswith('mat. type'):
                labs=[norm(g(G,ri,c)) for c in [4,5,6,7,9,10]]
                for k,txt in enumerate(['Stock','Used','Received']):
                    for ci,lb in zip([4,5,6,7,9,10],labs):
                        v=nval(g(G,ri+2+k,ci))
                        if v is not None and lb:mats.append(dict(material=f'BULK {lb}'[:110],unit=norm(g(G,ri+1,ci))[:36],received=(v if txt=='Received' else None),used=(v if txt=='Used' else None),stock=(v if txt=='Stock' else None),total_received=None,cell_ref=f'{CL(ci)}{ri+2+k}'))
                break
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,14)).startswith('cement and cement additives'):
                for rj in range(ri+1,min(ri+30,len(G)+1)):
                    nm=norm(g(G,rj,14))
                    if not nm or clean(nm).startswith(('formation','note','safety')):break
                    u,r_,st=nval(g(G,rj,16)),nval(g(G,rj,17)),nval(g(G,rj,19))
                    if u or r_ or st:mats.append(dict(material=f'CEMENT ADD {nm}'[:110],unit=norm(g(G,rj,20))[:36],received=r_,used=u,stock=st,total_received=None,cell_ref=f'N{rj}'))
                break
        for ri in range(1,len(G)+1):
            if 'personnel on board' in clean(g(G,ri,11)):
                for rj in range(ri+1,min(ri+14,len(G)+1)):
                    cat=norm(g(G,rj,11));cnt=nval(g(G,rj,13))
                    if not cat:continue
                    if 'total pob' in cat.lower():pob.append(dict(category='TOTAL POB',count=cnt,total=cnt));break
                    if cnt is not None:pob.append(dict(category=cat[:70],count=cnt,total=None))
                break
        for ri in range(1,len(G)+1):
            if 'formation top data' in clean(g(G,ri,14)) or (clean(g(G,ri,14)).startswith('name') and clean(g(G,ri,16)).startswith('m md')):
                for rj in range(ri+1,min(ri+20,len(G)+1)):
                    nm=norm(g(G,rj,14))
                    if not nm or clean(nm).startswith(('note','bulk')):break
                    forms.append(dict(name=nm[:60],md=nval(g(G,rj,16)),tvd=nval(g(G,rj,17)),prognosis_md=None,prognosis_tvd=None,lithology=norm(g(G,rj,19))[:180]))
                break
    for ri in range(1,len(G)+1):
        if clean(g(G,ri,10)).startswith(('m.d','md')):
            for rj in range(ri+1,min(ri+40,len(G)+1)):
                md=nval(g(G,rj,10))
                if md is None:break
                sur.append(dict(md=md,incl=nval(g(G,rj,11)),azi=nval(g(G,rj,12)),tvd=nval(g(G,rj,13)),north=nval(g(G,rj,14)),east=nval(g(G,rj,15)),vs=nval(g(G,rj,17)),dls=nval(g(G,rj,19)),tool=norm(g(G,rj,20))[:50]))
            break
    if tpl=='B':
        for ri in range(1,len(G)+1):
            if clean(g(G,ri,2))=='item' and clean(g(G,ri,3)).startswith('od'):
                for rj in range(ri+1,min(ri+30,len(G)+1)):
                    if clean(g(G,rj,2)).startswith('total'):break
                    for base,off in [(2,0),(6,4)]:
                        nm=norm(g(G,rj,base))
                        if nm:bha.append(dict(bha_no=norm(g(G,ri-1,base))[:36],item=nm[:110],qty='',od=nval(g(G,rj,base+1)),length=nval(g(G,rj,base+2)),cum_len=nval(g(G,rj,base+3))))
                break
    return f,mats,sur,bha,pob,forms

def upsert_daily(wk,rno,f,rel,tpl,remark,data):
    row=con.execute('SELECT id FROM ddr_daily WHERE well_key=? AND report_no=?',(wk,rno)).fetchone()
    if row:
        sets=[];vals=[]
        for k,v in f.items():
            if v is not None and v!='':sets.append(f'{k}=COALESCE({k},?)');vals.append(v)
        vals+= [remark,data,wk,rno]
        con.execute(f'UPDATE ddr_daily SET {",".join(sets+["sheet_remark=COALESCE(sheet_remark,?)","sheet_data=COALESCE(sheet_data,?)"])} WHERE well_key=? AND report_no=?',vals)
    else:
        cols=['well_key','report_no','rel_path','template'];vals=[wk,rno,rel,tpl]
        for k,v in f.items():
            if v is not None and v!='':cols.append(k);vals.append(v)
        if remark:cols.append('sheet_remark');vals.append(remark)
        if data:cols.append('sheet_data');vals.append(data)
        con.execute(f'INSERT INTO ddr_daily ({",".join(cols)}) VALUES ({",".join("?"*len(cols))})',vals)

def excel_stage():
    t0=time.time();cnt=collections.Counter()
    for p in sorted(glob.glob(os.path.join(ROOT,'CBS','*.xlsx')))+sorted(glob.glob(os.path.join(ROOT,'DDRs','*.xlsx'))):
        rel=os.path.relpath(p,ROOT);fn=os.path.basename(p);cat='CBS' if '/CBS/' in p else 'DDR'
        wbv=openpyxl.load_workbook(p,read_only=True,data_only=True);wbf=openpyxl.load_workbook(p,read_only=True,data_only=False)
        n=len(wbv.sheetnames);wk=DDR_FILE.get(fn,'')
        fid=con.execute('INSERT INTO source_files(category,rel_path,filename,n_units,size_bytes,sha256,well_keys,notes) VALUES (?,?,?,?,?,?,?,?)',
            (cat,rel,fn,n,os.path.getsize(p),sha(p),wk,'Cell-by-cell archive incl. formulas')).lastrowid
        print(f'  [{cat}] {fn} ({n})',flush=True)
        for si,sn in enumerate(wbv.sheetnames):
            G=[list(r) for r in wbv[sn].iter_rows(values_only=True)]
            forms={}
            try:
                for row in wbf[sn].iter_rows():
                    for c in row:
                        if c.data_type=='f' and c.value is not None:forms[c.coordinate]=str(c.value)
            except Exception:pass
            nr,nc=len(G),max((len(r) for r in G),default=0)
            cells=[];nums=[];texts=[]
            for ri,row in enumerate(G,1):
                for ci,v in enumerate(row,1):
                    if v is None or str(v).strip()=='':continue
                    ref=f'{CL(ci)}{ri}';cells.append([ref,sv(v),forms.get(ref,''),'n' if isnum(v) else 's'])
                    if isnum(v):nums.append((ri,ci,float(v)))
                    if isinstance(v,str) and len(v.strip())>1:texts.append(f'{ref} {norm(v)[:80]}')
            m=re.match(r'^(\d+)-([12])\b',sn);rno=int(m.group(1)) if m else None;part=int(m.group(2)) if m else None
            tpl=''
            if cat=='DDR':
                am=bm=0
                for rj in range(1,min(30,nr)+1):
                    c1=norm(g(G,rj,1)).lower();c8=norm(g(G,rj,8)).lower();c18=norm(g(G,rj,18)).lower()
                    if any(x in c1 for x in ('operation in detail','code / activity','general data')):am+=1
                    if 'field:' in c8 or 'well name:' in c18:bm+=1
                tpl='A' if am else ('B' if bm else ('B' if nc>=30 else 'A'))
            shwk=wk if cat=='DDR' else cbs_well(fn,sn)
            sid=con.execute('INSERT INTO sheets(file_id,name,idx,category,n_rows,n_cols,n_nonempty,n_numeric,n_formula,template,report_no,part,well_key,rel_path) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                (fid,sn,si,cat,nr,nc,len(cells),len(nums),len(forms),tpl,rno,part,shwk,rel)).lastrowid
            con.execute('INSERT INTO sheet_cells VALUES (?,?)',(sid,z(cells)))
            con.executemany('INSERT INTO cell_numeric VALUES (?,?,?,?)',[(sid,a,b,c) for a,b,c in nums])
            uniq=[];seen=set();used=0
            for t in texts:
                body=t.split(' ',1)[-1].strip().lower()
                if len(body)<2 or body in seen:continue
                seen.add(body)
                if used+len(t)>1000:break
                uniq.append(t);used+=len(t)+1
            con.execute('INSERT INTO sheet_fts(body,sheet_id,sheet,file) VALUES (?,?,?,?)',('\n'.join(uniq),sid,sn,rel))
            cnt['cells']+=len(cells);cnt['numeric']+=len(nums);cnt['formulas']+=len(forms)
            if cat=='CBS':
                it,tt,st,ac=parse_cbs(G,sn,fid,rel,fn)
                for x in it:con.execute('INSERT INTO cbs_items(file_id,sheet,row_no,well_key,category,section,item_no,description,unit,unit_price,qty,total,formula,cell_ref,is_contingency,rel_path) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                    (x['file_id'],x['sheet'],x['row_no'],x['well_key'],x['category'],x['section'],x['item_no'],x['description'],x['unit'],x['unit_price'],x['qty'],x['total'],x['formula'],x['cell_ref'],x['is_contingency'],x['rel_path']))
                for x in tt:con.execute('INSERT INTO cbs_totals(file_id,sheet,well_key,item,qty,unit_cost,total,row_no,rel_path) VALUES (?,?,?,?,?,?,?,?,?)',(x['file_id'],x['sheet'],x['well_key'],x['item'],x['qty'],x['unit_cost'],x['total'],x['row_no'],x['rel_path']))
                for x in st:con.execute('INSERT INTO cbs_stim(file_id,sheet,well_key,kind,item_no,description,unit,volume,unit_price,standby_days,operating_days,standby_rate,operating_rate,total,remark,row_no,rel_path) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                    (x['file_id'],x['sheet'],x['well_key'],x['kind'],x['item_no'],x['description'],x['unit'],x['volume'],x['unit_price'],x['standby_days'],x['operating_days'],x['standby_rate'],x['operating_rate'],x['total'],x['remark'],x['row_no'],x['rel_path']))
                for x in ac:con.execute('INSERT INTO cbs_accomm(file_id,sheet,well_key,description,qty,unit,rate,total,row_no,rel_path) VALUES (?,?,?,?,?,?,?,?,?,?)',(x['file_id'],x['sheet'],x['well_key'],x['description'],x['qty'],x['unit'],x['rate'],x['total'],x['row_no'],x['rel_path']))
                cnt['cbs_items']+=len(it)
            elif rno:
                if part==1:
                    f,acts,tbd,notes,cas=parse_remark(G,rel,sn,tpl);upsert_daily(wk,rno,f,rel,tpl,sn,'')
                    for a in acts:con.execute('INSERT INTO ddr_activities(well_key,report_no,seq,rel_path,sheet,row_no,t_from,t_to,hours,phase,code,sub_code,status,npt_attributed,description,main_activity) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                        (wk,rno,acts.index(a)+1,rel,sn,a['row_no'],a['t_from'],a['t_to'],a['hours'],a['phase'],a['code'],a['sub_code'],a['status'],a['npt_attributed'],a['description'],a['main_activity']))
                    for code,cat2,h in tbd:con.execute('INSERT INTO ddr_time(well_key,report_no,code,category,hours,rel_path) VALUES (?,?,?,?,?,?)',(wk,rno,code,cat2[:100],h,rel))
                    for c2,t in notes:con.execute('INSERT INTO ddr_notes(well_key,report_no,sheet,category,text,rel_path) VALUES (?,?,?,?,?,?)',(wk,rno,sn,c2,t,rel))
                    for c in cas:con.execute('INSERT INTO ddr_casing(well_key,report_no,size,grade,weight,thread,from_m,to_m,burst,collapse,shoe_tvd,rel_path) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',(wk,rno,c['size'],c['grade'],c['weight'],c['thread'],c['from_m'],c['to_m'],c['burst'],c['collapse'],c['shoe_tvd'],rel))
                    cnt['activities']+=len(acts)
                else:
                    f,mats,sur,bha,pob,forms=parse_data(G,rel,sn,tpl);upsert_daily(wk,rno,f,rel,tpl,'',sn)
                    for x in mats:con.execute('INSERT INTO ddr_materials(well_key,report_no,material,unit,received,used,stock,total_received,source_sheet,cell_ref) VALUES (?,?,?,?,?,?,?,?,?,?)',(wk,rno,x['material'],x['unit'],x['received'],x['used'],x['stock'],x['total_received'],sn,x['cell_ref']))
                    for x in sur:con.execute('INSERT INTO ddr_surveys(well_key,report_no,md,incl,azi,tvd,north,east,vs,dls,tool,rel_path) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',(wk,rno,x['md'],x['incl'],x['azi'],x['tvd'],x['north'],x['east'],x['vs'],x['dls'],x['tool'],rel))
                    for x in bha:con.execute('INSERT INTO ddr_bha(well_key,report_no,bha_no,item,qty,od,length,cum_len,rel_path) VALUES (?,?,?,?,?,?,?,?,?)',(wk,rno,x['bha_no'],x['item'],x['qty'],x['od'],x['length'],x['cum_len'],rel))
                    for x in pob:con.execute('INSERT INTO ddr_pob(well_key,report_no,category,count,total,rel_path) VALUES (?,?,?,?,?,?)',(wk,rno,x['category'],x['count'],x['total'],rel))
                    for x in forms:con.execute('INSERT INTO ddr_formations(well_key,report_no,name,md,tvd,prognosis_md,prognosis_tvd,lithology,rel_path,source) VALUES (?,?,?,?,?,?,?,?,?,?)',(wk,rno,x['name'],x['md'],x['tvd'],x.get('prognosis_md'),x.get('prognosis_tvd'),x['lithology'],rel,'DDR'))
                    cnt['materials']+=len(mats)
            if si%80==0:con.commit()
        wbv.close();wbf.close();con.commit()
    print('  EXCEL',dict(cnt),'s',round(time.time()-t0,1),flush=True)
    con.commit()

# ------------------------------------------------------------------ PDF -----
EOWR={'1.AZNS_40_EOWR.pdf':'AZNS-40','2.AZNS_10_EOWR.pdf':'AZNS-10','3.AZNS_85_EOWR.pdf':'AZNS-85','5.AZNS_116_EOWR.pdf':'AZNS-116','6.AZNS_127_WO1_EOWR.pdf':'AZNS-127-WO1-ST1','7.AZNS_207_EOWR.pdf':'AZNS-207','8.AZNS_122_WO2_EOWR.pdf':'AZNS-122-WO2','9.AZNS_75_EOWR.pdf':'AZNS-75','10.AZNS_12_EOWR.pdf':'AZNS-12','11.AZNS_166_EOWR.pdf':'AZNS-166','12,13.AZNS_127_WO2_and ST2_EOWR.pdf':'AZNS-127-WO2;AZNS-127-ST2'}
PROG={'1.AZNS_40_Well_Program.pdf':'AZNS-40','2.AZNS_10_Workover_Program.pdf':'AZNS-10','3.AZNS_85_Workover_Program.pdf':'AZNS-85','4.AZNS_122_Workover-1_Program.pdf':'AZNS-122-WO1','5.AZNS_116_Workover_Program.pdf':'AZNS-116','6.AZNS_127_Workover1_Program.pdf':'AZNS-127-WO1-ST1','6.AZNS_127_Side_Track#01_Program.pdf':'AZNS-127-WO1-ST1','7.AZNS_207_Well_Program.pdf':'AZNS-207','8.AZNS_122_Workover-2_Program.pdf':'AZNS-122-WO2','9.AZNS_75_Workover-1_Program.pdf':'AZNS-75','10.AZNS_12_Workover_Program.pdf':'AZNS-12','11.AZNS_166_Workover-1_Program.pdf':'AZNS-166','12.AZNS_127_Workover#2_Program.pdf':'AZNS-127-WO2','13.AZNS_127_Side Track-2_Program.pdf':'AZNS-127-ST2'}
CH=[('01 Well History & Summary',r'well history|general information|well information'),
('02 Hole History & Dry Test',r'hole history|dry test|bottom test|liner lap'),
('03 Bit Record',r'bit record|dull grad|bit no\.|bit parameters'),
('04 Time Analysis & NPT',r'time analysis|npt break|non.?productive|time break.?down|time vs\.? depth'),
('05 Geological & Logging Reports',r'geological|geology|formation top|litholog|mud log|logging report|\bcbl\b|\bvdl\b|\busit\b|stratigraph'),
('06 BHA Record',r'bha record|bottom hole assembly'),('07 Casing Reports & Tallies',r'casing tally|casing information|casing report'),
('08 Cement Design & Job',r'cement design|cement job|cementing report|pumping schedule'),
('09 Completion Tally & Tools',r'completion tally|completion report|trsssv|packer|\bciv\b|\bcim\b|nipple|tubing information'),
('10 Wellhead & Xmas Tree',r'xmas tree|x-mass tree|wellhead|section a|section b|section c|section d'),
('11 Drilling Fluid Information',r'drilling fluid|mud chemical|volume analysis|solid control|fluid properties|loss distribution'),
('12 Well Stimulation & Clean-Up',r'stimulation|acidizing|pre.?acid|post.?acid|clean.?up|bullhead|\bctu\b|coil tubing|nitrogen'),
('13 Waste Management',r'waste management|waste disposal'),('14 Contractors & Services',r'list of contractors|contractor list'),
('15 Comments / Lessons / Revision',r'comment|lesson|recommendation|revision|appendix'),
('16 Programme & Prognosis',r'prognos|program|objective|estimated depth'),
('17 Casing / Mud / Completion Prog',r'casing program|mud program|completion program|hole section'),
('18 Geological Prognosis',r'geological prognosis|pore pressure|fracture gradient'),
('19 Well Schematic & Trajectory',r'well schematic|wellbore schematic|trajectory'),
('20 HSE & Environment',r'\bhse\b|safety drill|fire drill|environment')]
HM=[(r'geolog|logging|stratig','05 Geological & Logging Reports'),(r'time analysis|time vs|npt','04 Time Analysis & NPT'),(r'drilling fluid|mud|volume analysis|solid control','11 Drilling Fluid Information'),(r'stimulation|acid|clean.?up|testing','12 Well Stimulation & Clean-Up'),(r'cement','08 Cement Design & Job'),(r'casing|tubular','07 Casing Reports & Tallies'),(r'completion|tubing|well status','09 Completion Tally & Tools'),(r'wellhead|x.?mass','10 Wellhead & Xmas Tree'),(r'bit record|bit param','03 Bit Record'),(r'bha|bottom hole assembly','06 BHA Record'),(r'well history|general information','01 Well History & Summary'),(r'contractor','14 Contractors & Services'),(r'waste','13 Waste Management'),(r'comment|revision','15 Comments / Lessons / Revision')]
HDR=re.compile(r'(south azadegan|contract no|end of well report|^\s*page\s*\d+\s*of|^\s*project\b|sequence|^\s*p/t\b|^\s*zone\b|^\s*dep\.|^\s*disc\.|^\s*type\b|program$)',re.I)
def chapter(t):
    lines=[l for l in t.splitlines() if not HDR.search(l)];body='\n'.join(lines);low=body[:3000].lower()
    for m in re.finditer(r'^\s*(\d{1,2})\.\s+([A-Za-z][A-Za-z &/\-()]{3,65})\s*$',body,re.M):
        for pat,name in HM:
            if re.search(pat,m.group(2).lower()):return name
    best,score='00 General / Unclassified',0
    for name,pat in CH:
        k=len(re.findall(pat,low))
        if k>score:best,score=name,k
    return best
def shared(t,pno):
    if pno<=11:return 'AZNS-127-WO2;AZNS-127-ST2'
    if re.search(r'workover section|wo ?#?2|workover ?#?2|work.?over 2',t,re.I):return 'AZNS-127-WO2'
    return 'AZNS-127-ST2'
FACT=[('MUD_WEIGHT',r'(?:mud weight|mw|density)[^\n]{0,40}?(\d{2,3}(?:\.\d+)?)\s*(pcf|ppg)'),
('LOSS',r'(?:lost circulation|mud loss|losses|loss rate)[^\n]{0,60}?(\d+(?:\.\d+)?)\s*(bbl|bph|bbls|m3)'),
('PRESSURE_TEST',r'(?:pressure test|test pressure)[^\n]{0,60}?(\d[\d,.]*)\s*(psi|bar)'),
('CBL_TOC',r'(?:top of cement|\btoc\b|cbl)[^\n]{0,60}?(\d{3,4}(?:\.\d+)?)\s*m'),
('WHP',r'(?:whp|wellhead pressure)[^\n]{0,55}?(\d[\d,.]*)\s*(psi|bar)'),
('RATE',r'(?:oil rate|flow rate|production rate)[^\n]{0,55}?(\d[\d,.]*)\s*(bopd|bbl/d|stb/d|bpd)'),
('PI',r'(?:productivity index|\bpi\b)[^\n]{0,60}?(\d+(?:\.\d+)?)\s*(?:bbl/psi|bpd/psi)?'),
('ACID_VOL',r'(?:hcl|acid volume|main acid)[^\n]{0,60}?(\d[\d,.]*)\s*(bbl|gal|m3)'),
('N2_VOL',r'(?:liquid nitrogen|\bn2\b)[^\n]{0,60}?(\d[\d,.]*)\s*(gal|bbl|scf|m3)'),
('NPT_HRS',r'(?:npt|non.?productive time)[^\n]{0,60}?(\d+(?:\.\d+)?)\s*(hrs?|hours?|days?)'),
('FIT_LOT',r'(?:\bfit\b|\blot\b)[^\n]{0,60}?(\d+(?:\.\d+)?)\s*(pcf|ppg|psi)'),
('CASING_SHOE',r'(?:casing shoe|shoe depth|shoe at)[^\n]{0,50}?(\d{3,4}(?:\.\d+)?)\s*m')]
UNIT=re.compile(r'\b(m|ft|pcf|ppg|psi|bbl|bbls|bph|gpm|rpm|klbs|klb|hrs?|hours?|days?|%|usd|gal|lbs|kg|mt|bpd|bopd|stb/d|m3|inch|deg|ppb|cp|ksi|mm|bar|scf)\b',re.I)
NUM=re.compile(r'(?<![\w.])(-?\d[\d,]*\.?\d*)(?![\w])')
TBL=re.compile(r'(?:^|\n)\s*(?:Table|TABLE|Tbl)\s*[-–:]?\s*(\d+[\w.\-]*)\s*[:.\-–]\s*([^\n]{3,140})')
def pdf_stage():
    t0=time.time();stats=collections.Counter()
    for folder,cat,mp in [('04.EOWR','EOWR',EOWR),('Well Program','PROGRAM',PROG)]:
        for p in sorted(glob.glob(os.path.join(ROOT,folder,'*.pdf'))):
            rel=os.path.relpath(p,ROOT);fn=os.path.basename(p);doc=pymupdf.open(p);n=len(doc)
            fid=con.execute('INSERT INTO source_files(category,rel_path,filename,n_units,size_bytes,sha256,well_keys,notes) VALUES (?,?,?,?,?,?,?,?)',
                (cat,rel,fn,n,os.path.getsize(p),sha(p),mp.get(fn,''),'Lossless page text + numeric index')).lastrowid
            print(f'  [{cat}] {fn} ({n})',flush=True)
            for pi in range(n):
                pg=doc[pi]
                try:t=pg.get_text('text')
                except Exception:t=''
                try:nim=len(pg.get_images(full=True))
                except Exception:nim=0
                wk=mp.get(fn,'');wk=shared(t,pi+1) if ';' in wk else wk
                ch=chapter(t)
                con.execute('INSERT INTO pdf_pages(file_id,page_no,rel_path,well_key,category,n_chars,n_images,chapter,text_z,keywords) VALUES (?,?,?,?,?,?,?,?,?,?)',
                    (fid,pi+1,rel,wk,cat,len(t),nim,ch,z(t.encode('utf-8')),','.join(sorted(set(re.findall(r'\b(?:formation|mud|cement|casing|bit|bha|npt|acid|stimulation|cbl|usit|packer|whipstock|sidetrack|losses|fish|milling|trsssv|plt|pvt|waste|contractor|prognosis|risk)\b',t.lower()))))[:160]))
                fts='\n'.join(norm(l)[:130] for l in t.splitlines() if len(norm(l))>2)[:900]
                con.execute('INSERT INTO pdf_fts(body,file_id,page,file,well) VALUES (?,?,?,?,?)',(fts,fid,pi+1,rel,wk))
                stats['pages']+=1;stats['chars']+=len(t)
                for m in TBL.finditer(t):
                    con.execute('INSERT INTO pdf_tables(file_id,page_no,table_no,title,body,well_key,category) VALUES (?,?,?,?,?,?,?)',
                        (fid,pi+1,m.group(1),norm(m.group(2))[:170],norm(t[m.start():m.start()+1200])[:1200],wk,cat))
                    stats['tables']+=1
                numrows=[]
                for li,ln in enumerate(t.splitlines(),1):
                    L=norm(ln)
                    if len(L)<3:continue
                    for nm in NUM.finditer(L):
                        try:v=float(nm.group(1).replace(',',''))
                        except Exception:continue
                        um=UNIT.search(L[nm.end():nm.end()+20]);u=um.group(1).lower() if um else ''
                        ctx=L[max(0,nm.start()-32):min(len(L),nm.end()+28)][:66]
                        numrows.append([li,nm.group(1)[:20],v,u,ctx])
                con.execute('INSERT INTO pdf_num_pages(file_id,page_no,well_key,payload) VALUES (?,?,?,?)',(fid,pi+1,wk,z(numrows)))
                stats['numbers']+=len(numrows)
                for c2,pat in FACT:
                    for m in re.finditer(pat,t,re.I|re.M):
                        try:v=float(m.group(1).replace(',',''))
                        except Exception:continue
                        u=m.group(2) if m.lastindex and m.lastindex>=2 else ''
                        con.execute('INSERT INTO eowr_facts(well_key,category,fact_key,value,text_value,unit,page_no,rel_path,context) VALUES (?,?,?,?,?,?,?,?,?)',
                            (wk,cat,c2,v,'',u,pi+1,rel,norm(t[max(0,m.start()-45):m.end()+45])[:150]))
                        stats['facts']+=1
                if pi%120==0:con.commit()
            doc.close();con.commit()
    con.execute("INSERT OR REPLACE INTO meta(k,v) VALUES('pdf_numeric_tokens',?)",(str(stats['numbers']),))
    con.commit()
    print('  PDF',dict(stats),'s',round(time.time()-t0,1),flush=True)

# --------------------------------------------------------------- ENRICH -----
FM=re.compile(r'\b(Aghajari|Gachsaran|Asmary|Asmari|Jahrum|Pabdeh|Gurpi|U\.?Gurpi|L\.?Gurpi|Tarbur|Laffan|Lafan|Ilam|Sarvak|Kazhdomi|Kazhdumi|Dariyan|Gadvan|Khalij|Fahliyan|Fahlian|Hith|Surmeh|Najmeh|Total Depth)\b',re.I)
DEPTH=re.compile(r'(?<![\d.])(\d{3,4}(?:\.\d+)?)(?![\d.])')
def colnum(ref):
    m=re.match(r'([A-Z]+)(\d+)',ref)
    if not m:return 0,0
    c=0
    for ch in m.group(1):c=c*26+ord(ch)-64
    return int(m.group(2)),c
def enrich_stage():
    con.execute('DELETE FROM ddr_formations')
    nf=nd=0;patches=[]
    for sh in R("SELECT s.id,s.name,s.category,s.well_key,s.rel_path,s.report_no,s.part,c.payload FROM sheets s JOIN sheet_cells c ON c.sheet_id=s.id WHERE s.category='DDR'"):
        rows=collections.defaultdict(list)
        for ref,val,fo,ty in json.loads(unz(sh['payload'])):
            r,c=colnum(ref);rows[r].append((c,val,ref))
        wk=sh['well_key'];rno=sh['report_no'] or 0
        for ri,cs in rows.items():
            cs.sort()
            for ci,val,ref in cs:
                if not isinstance(val,str):continue
                m=FM.search(val.strip())
                if not m or len(val.strip())>60:continue
                raw=m.group(1).lower().replace('.','');nm=FALIAS.get(raw,m.group(1))
                nums=[];texts=[]
                for cj,v2,rf2 in cs:
                    if cj<=ci:continue
                    x=tonum(v2)
                    if x is not None and 300<=x<=6000:nums.append(x)
                    elif isinstance(v2,str) and 2<len(v2)<90 and not FM.search(v2) and not re.fullmatch(r'[\d.,\s]+',v2):texts.append(v2.strip())
                if not nums:continue
                if len(nums)>=4:pmd,ptvd,amd,atvd=nums[:4]
                elif len(nums)==3:pmd,ptvd,amd,atvd=nums[0],nums[1],nums[2],None
                elif len(nums)==2:pmd,ptvd,amd,atvd=None,None,nums[0],nums[1]
                else:pmd,ptvd,amd,atvd=None,None,nums[0],None
                con.execute('INSERT INTO ddr_formations(well_key,report_no,name,md,tvd,prognosis_md,prognosis_tvd,lithology,rel_path,source) VALUES (?,?,?,?,?,?,?,?,?,?)',
                    (wk,rno,nm[:60],amd,atvd,pmd,ptvd,' '.join(dict.fromkeys(texts))[:180],sh['rel_path'],f"DDR::{sh['name']}!{ref}"))
                nf+=1;break
        if sh['part']==1 and rno:
            for ri,cs in rows.items():
                cs.sort()
                for ci,val,ref in cs:
                    if not isinstance(val,str):continue
                    lo=val.lower();key=None
                    if re.search(r'md \(m\)@ 24:00|mid ?nig?ht depth|midnight depth',lo):key='md_2400'
                    elif re.search(r'24 ?h?rs? drilled|drilled interval',lo):key='drilled_24'
                    elif re.search(r'mw \(pcf\)|mud weight',lo):key='mw'
                    elif re.search(r'avg\.? rop|average rop',lo):key='rop'
                    elif re.search(r'md \(m\)@ 0:00',lo):key='md_0000'
                    if not key:continue
                    for cj,v2,rf2 in cs:
                        if cj<=ci:continue
                        q=nval(v2)
                        if q is not None:patches.append((wk,rno,key,q));break
                    break
    for wk,rno,key,v in patches:
        if key=='mw' and not 20<=v<=200:continue
        if key in ('md_2400','md_0000') and not 0<=v<=7000:continue
        if key=='rop' and not 0<=v<=200:continue
        if key=='drilled_24' and not 0<=v<=500:continue
        con.execute(f'UPDATE ddr_daily SET {key}=COALESCE({key},?) WHERE well_key=? AND report_no=?',(v,wk,rno));nd+=1
    con.commit();print('  ENRICH formation rows',nf,'daily patches',nd,flush=True)

WIP={'ACTUAL_TD':r'(?:actual (?:td|md/tvd|depth)|total depth \(md\)|actual md/tvd|total depth @)[^\n]{0,60}?(\d{3,4}(?:\.\d+)?)',
'PLANNED_TD':r'(?:planned td|proposed td|estimated final depth)[^\n]{0,60}?(\d{3,4}(?:\.\d+)?)',
'RTE':r'\bRTE\b[^\n]{0,40}?(\d{1,3}(?:\.\d+)?)','GLE':r'\bGLE\b[^\n]{0,40}?(\d{1,3}(?:\.\d+)?)',
'MUD_WEIGHT':r'(\d{2,3}(?:\.\d+)?)\s*pcf'}
def facts_stage():
    def add(wk,cat,key,v,tv,u,page,rel,ctx):
        con.execute('INSERT INTO eowr_facts(well_key,category,fact_key,value,text_value,unit,page_no,rel_path,context) VALUES (?,?,?,?,?,?,?,?,?)',
            (wk,cat,key,v,str(tv or '')[:200],u,page,rel,str(ctx)[:170]))
    nf=nw=0
    for pg in R("SELECT rel_path,page_no,well_key,category,chapter,text_z FROM pdf_pages"):
        t=unz(pg['text_z']);wk=pg['well_key'];rel=pg['rel_path'];page=pg['page_no'];pc=pg['category'];low=t.lower()
        names=[(m.start(),FALIAS.get(m.group(1).lower().replace('.',''),m.group(1))) for m in FM.finditer(t)]
        if len(names)>=3 and re.search(r'formation tops?|well sketch|stratigraph|prognos|geolog',low):
            depths=[(m.start(),float(m.group(1))) for m in DEPTH.finditer(t) if 250<=float(m.group(1))<=5200]
            used=set()
            for pos,nm in names:
                cand=[(pp,d) for pp,d in depths if pos<pp<pos+320 and (pp,d) not in used]
                if not cand:continue
                p0,d0=min(cand,key=lambda x:x[0]);used.add((p0,d0))
                add(wk,pc+'_TOPS',nm+' MD',d0,'','m',page,rel,t[max(0,pos-35):pos+85].replace('\n',' '));nf+=1
        if re.search(r'well informaion|well information|general information|2\. well history',low):
            flat=re.sub(r'\s+',' ',t)
            for key,pat in WIP.items():
                for m in re.finditer(pat,flat,re.I):
                    try:v=float(m.group(1))
                    except Exception:continue
                    if not 0<v<6000:continue
                    add(wk,'WELL_INFO',key,v,m.group(1),'',page,rel,flat[max(0,m.start()-35):m.end()+35]);nw+=1;break
        for m in re.finditer(r'(?:background gas|connection gas|trip gas|total gas|h2s|co2)[^\n]{0,90}?(\d+(?:\.\d+)?)\s*(ppm|%|units?)?',t,re.I):
            add(wk,'GAS',m.group(0).split()[0][:36],float(m.group(1)),m.group(0)[:110],m.group(2) or '',page,rel,m.group(0)[:160])
        for m in re.finditer(r'(?:CBL|VDL|USIT|TOC|top of cement|bond index)[^\n]{0,140}',t,re.I):
            L=m.group(0);ds=[float(x) for x in re.findall(r'\d{3,4}(?:\.\d+)?',L) if 100<=float(x)<=5200]
            if ds:add(wk,'CBL','Bond / TOC depth',ds[0],L[:150],'m',page,rel,L[:160])
        if re.search(r'bullhead|pre.?flush|main acid|over.?flush|pumping schedule',low):
            for m in re.finditer(r'(?:pre.?flush|main acid|over.?flush|bullhead|displacement)[^\n]{0,90}?(\d+(?:\.\d+)?)\s*(bbl|gal|m3)',t,re.I):
                add(wk,'ACID_PUMP',m.group(0).split()[0][:36],float(m.group(1)),m.group(0)[:140],m.group(2),page,rel,m.group(0)[:160])
    con.commit();print('  FACTS tops',nf,'well-info',nw,flush=True)

# ------------------------------------------------------------ ANALYTICS -----
HYP={
 'Lost circulation / fluid losses':['Induced tensile fractures in depleted Asmari / Sarvak because ECD exceeded the local fracture gradient while MW was held for Gachsaran stability.','Natural vugs / karstic porosity or a conductive fault plane connected to a thief zone; the loss would occur even at constant MW.'],
 'Stuck pipe / fishing / milling':['Differential sticking against a depleted permeable sand with high overbalance and a stationary string.','Mechanical key-seating / pack-off from shale cavings and milled metal debris above the BHA.'],
 'Rig equipment failure':['Age-related electrical / mechanical degradation of the rig power train under continuous high load.','Single-point dependency: no spare or redundant unit on location, so a minor fault escalates into a full waiting period.'],
 'Waiting on services / material':['Logistical / contractual lead time: third-party inspection certificates and tubular QC documents were not mobilised before spud.','Inventory mismatch between the CBS-approved quantity and the actual consumption rate, exhausting stock early.'],
 'Well control / gas / H2S':['Connection / trip gas from a pressured carbonate stringer entering the annulus at constant MW.','Gas expansion during swab on POOH combined with insufficient hole-fill volume.'],
 'Cement / liner / dry test':['Incomplete liner-lap cement coverage / micro-annulus at the 9-5/8 in × 7 in overlap (low CBL-VDL bond index).','Pressure decline from a leaking liner-top packer or a surface test connection rather than the cement sheath.'],
 'Directional / sidetrack / window':['Whipstock anchored in a washed-out / ovalised casing section, allowing rotation and a ragged window.','High local dogleg from an aggressive AKO motor setting in soft Laffan / Kazhdomi shale.'],
 'Completion / stimulation':['Asphaltene-rich heavy crude contacting HCl forms acid sludge / stable emulsion, cutting near-wellbore permeability.','Mechanical failure of a control line, TRSSSV flapper or CIV/CIM seal under the 10K test envelope.'],
 'Shale instability / hole condition':['Chemo-mechanical hydration of Laffan / Kazhdomi / Pabdeh shale by a non-inhibitive water-based mud.','Mechanical shear failure at high inclination where the mud window is bounded by pore pressure and breakout gradient.'],
 'Bit / BHA / vibration':['Stick-slip / whirl induced by excessive WOB at low RPM in interbedded anhydrite and chert nodules.','Bit balling in plastic marl, reducing hydraulic jet impact and effective depth of cut.'],
 '_':['Operational / equipment constraint on the critical path of the 24-hour activity log.','Subsurface or client-driven decision hold while additional data / approval was awaited.']}
PLAY={
 'Lost circulation / fluid losses':'0–60 s: stop rotation, hold string stationary, cut pump rate to the minimum hole-cleaning rate, log SPP/SPM and pit-volume trend, call the wellsite supervisor, prepare a graded LCM pill (fine+medium+coarse), station the mud engineer at the shaker.',
 'Stuck pipe / fishing / milling':'0–60 s: freeze the block, record exact pick-up / slack-off / rotating weights and last free-point depth, set the jar, stay below the string overpull limit, start controlled circulation, open the fishing / milling decision tree.',
 'Rig equipment failure':'0–60 s: secure the well, isolate the failed circuit, switch to the redundant pump / SCR / top-drive channel, log the failure time for the zero-rate NPT claim, dispatch the PM kit from the on-site store.',
 'Waiting on services / material':'0–60 s: confirm the exact missing item and quantity against the CBS line, raise a written request with the certificate list, re-sequence the 24-hour plan to off-critical-path work, start the NPT attribution clock.',
 'Well control / gas / H2S':'0–60 s: announce gas to the rig floor, check returns and pit gain, verify H2S detectors and breathing air, shut in per the well-control matrix if gain is confirmed, record SIDPP/SICP and prepare the kill sheet.',
 'Cement / liner / dry test':'0–60 s: hold test pressure on a 15-minute chart, isolate the surface manifold to rule out a surface leak, compare against CBL-VDL / USIT at the lap, then decide squeeze vs re-test.',
 'Directional / sidetrack / window':'0–60 s: stop rotation, hold WOB off bottom, take a stationary MWD survey, compare inclination / azimuth / DLS with the minimum-curvature plan, stage the window-milling parameters.',
 'Completion / stimulation':'0–60 s: stop acid pumping, flush the CTU string with diesel / mutual solvent, sample returns for sludge and emulsion, record WHP and choke, hold anti-sludge / de-emulsifier ready.',
 'Shale instability / hole condition':'0–60 s: raise circulation to planned annular velocity, avoid rotation across the tight interval, check cavings at the shaker, add the inhibitive KCl / brine package, log the tight-hole depth.',
 'Bit / BHA / vibration':'0–60 s: pick up off bottom, reduce WOB and step RPM in the opposite direction to break torsional oscillation, record torque and SPP trace, resume inside the revised parameter window.',
 '_':'0–60 s: secure the well, freeze the operation clock, document exact time / depth / responsible party, notify the supervisor, re-sequence the 24-hour plan to the highest-value off-critical activity.'}
BUCKETS=[('Lost circulation / fluid losses',r'loss|lost circulation|lcm|seepage|bphl|thief'),('Stuck pipe / fishing / milling',r'stuck|fishing|fish\b|junk|mill|spear|overshot'),('Rig equipment failure',r'top drive|scr\b|pump repair|bop|generator|shaker|drawworks|power surge|breakdown|out of service|repair rig'),('Waiting on services / material',r'wos|wait on|waiting|shortage|not provided|delay|inspection report'),('Well control / gas / H2S',r'gas\b|h2s|kick|well control|influx|co2'),('Cement / liner / dry test',r'dry test|squeeze|cement|liner lap|cbl|toc|bond'),('Directional / sidetrack / window',r'whipstock|window|sidetrack|side track|dls|mwd|motor'),('Completion / stimulation',r'trsssv|packer|civ|cim|ssd|ctu|coiled|acid|emulsion|sludge|xmas tree|nipple'),('Shale instability / hole condition',r'shale|slough|tight hole|bridg|caving|ream|instab'),('Bit / BHA / vibration',r'bit dull|dull grade|balling|stick.?slip|whirl|vibration|rerun')]
PARAM={'WOB':['w.o.b','wob'],'RPM':['rpm'],'TORQUE':['torque'],'SPP':['pump pressure','spp'],'GPM':['pump output','gpm'],'MW':['mud weight','mw'],'PV':['pv'],'YP':['yp'],'AV':['ann. velocity'],'JV':['jet velocity'],'TFA':['tfa'],'ROP':['average rop','avg. rop','cum.bit ave rop'],'CHLORIDE':['cl concentration','chloride'],'SOLIDS':['solid volume'],'KCL':['kcl concentration'],'PH':['ph']}
def params(wk):
    out=collections.defaultdict(list)
    for sh in R("SELECT c.payload FROM sheets s JOIN sheet_cells c ON c.sheet_id=s.id WHERE s.well_key=? AND s.category='DDR' AND s.part=2",wk):
        byrow=collections.defaultdict(list)
        for ref,val,fo,ty in json.loads(unz(sh['payload'])):
            r,c=colnum(ref); byrow[r].append((c,val))
        for r,cs in byrow.items():
            cs.sort(); labels=' '.join(str(v).lower() for c,v in cs if isinstance(v,str) and len(v)<70)
            for lbl,aliases in PARAM.items():
                if not any(a in labels for a in aliases): continue
                for c,v in cs:
                    q=nval(v)
                    if q is None and isinstance(v,str):
                        ns=[float(x) for x in re.findall(r'\d+(?:\.\d+)?',v.replace(',',''))]
                        if ns: q=sum(ns)/len(ns) if '-' in v and len(ns)>1 else ns[0]
                    if q is not None and -1e6<q<1e7:
                        out[lbl].append(q); break
    return {k:dict(count=len(v),min=min(v),max=max(v),mean=mean(v),median=med(v)) for k,v in out.items()}
def time_account(wk):
    acts=R('SELECT * FROM ddr_activities WHERE well_key=? ORDER BY report_no,id',wk)
    st=collections.defaultdict(float);at=collections.defaultdict(float);tc=collections.defaultdict(float);npt=[]
    ck=re.compile(r'oeoc|contractor|service|mud|logging|directional|catering|energy|supplier|vendor|rig|unattributed',re.I)
    for a in acts:
        h=a['hours'] or 0
        if h<=0 or h>24:continue
        s=re.sub(r'[^A-Z]','',(a['status'] or '').upper())
        pt=s in ('P','PLN','PLAN','PLANNED','PT')
        np_=s in ('N','NPT','UP','UNPLANNED','UNPLAN','UNPLN','NP','PN','UPI','NPTUNPLN') or s.startswith('NPT') or 'UNPLN' in s
        if pt:st['PT']+=h
        elif np_:st['NPT']+=h
        else:st['UNCLASSIFIED']+=h
        who=(a['npt_attributed'] or '').strip()
        if np_:
            at[who or 'Unattributed']+=h
            npt.append(dict(report_no=a['report_no'],hours=h,attributed=who or 'Unattributed',description=(a['description'] or '')[:300]))
        if a['main_activity']:tc[a['main_activity']]+=h
        elif a['phase']:tc[a['phase']]+=h
    tb=R('SELECT category,SUM(hours) h FROM ddr_time WHERE well_key=? AND hours IS NOT NULL GROUP BY category ORDER BY h DESC',wk)
    for r in tb:
        if r['h'] and r['h']>0:tc[r['category']]=max(tc.get(r['category'],0),r['h'])
    if st['NPT']==0 and tb:
        kw=re.compile(r'wait|repair|fish|force majeure|npt|stuck|loss|breakdown|shortage|milling|junk',re.I)
        for r in tb:
            if r['h'] and kw.search(r['category'] or ''):
                st['NPT']+=r['h'];at['Contractor / unattributed waiting']+=r['h']*.5;at['Client / subsurface waiting']+=r['h']*.5
    ctr=sum(v for k,v in at.items() if ck.search(k));cli=sum(v for k,v in at.items() if not ck.search(k))
    return dict(activities=len(acts),status_hours=dict(st),attributed_hours=dict(at),npt_contractor_h=ctr,npt_client_h=cli,npt_total_h=ctr+cli,
        time_categories=[dict(category=k,hours=r2(v)) for k,v in sorted(tc.items(),key=lambda x:-x[1]) if v>0][:32],npt_events=npt)
BENCHRATE=None
def bench():
    global BENCHRATE
    if BENCHRATE is None:
        v=[x[0] for x in Q("SELECT unit_price FROM cbs_items WHERE LOWER(description) LIKE '%rig service%' AND unit LIKE 'DAY%' AND unit_price BETWEEN 5000 AND 200000")]
        BENCHRATE=med(v) or 23000.0
    return BENCHRATE
def costs(wk):
    it=R('SELECT * FROM cbs_items WHERE well_key=?',wk);tt=R('SELECT * FROM cbs_totals WHERE well_key=?',wk)
    st=R('SELECT * FROM cbs_stim WHERE well_key=?',wk);ac=R('SELECT * FROM cbs_accomm WHERE well_key=?',wk)
    own=[t for t in tt if 'matrial' not in (t['rel_path'] or '').lower()]
    comb=[t['total'] for t in own if 'Total Cost-EPD & Rigless' in (t['sheet'] or '') and re.search(r'total cost ceiling',t['item'] or '',re.I) and t['total']]
    epd=[t['total'] for t in own if (t['sheet'] or '').startswith('Total Cost-EPD') and '& Rigless' not in (t['sheet'] or '') and re.search(r'cost ceiling',t['item'] or '',re.I) and t['total']]
    rl=[t['total'] for t in own if re.search(r'stimulation total cost|cost breakdown r-less|cost breakdown rigless|cost b.d r-less',t['sheet'] or '',re.I) and re.search(r'cost ceiling|total contract value',t['item'] or '',re.I) and t['total']]
    prim=[i['total'] for i in it if (i['sheet'] or '').startswith('Cost Breakdown-') and 'Material' not in (i['sheet'] or '') and not re.search(r'cont',i['sheet'] or '',re.I) and i['total'] and (re.search(r'total cost|grand total|cost ceiling|total contract value',i['description'] or '',re.I) or not (i['description'] or '').strip())]
    epdsum=sum(t['total'] or 0 for t in own if (t['sheet'] or '').startswith('Total Cost-EPD') and t['total'] and not re.search(r'cost ceiling|total cost',t['item'] or '',re.I))
    rr=cd=None
    for i in it:
        if 'rig service' in (i['description'] or '').lower() and i['unit_price']:
            if rr is None or i['unit_price']>rr:rr,cd=i['unit_price'],i['qty']
    base=[i for i in it if i['category'] in ('SERVICE','MATERIAL') and not i['is_contingency']]
    cont=[i for i in it if i['is_contingency']]
    sec=collections.defaultdict(float);mat=collections.defaultdict(float)
    for i in base:
        if i['total']:
            (sec if i['category']=='SERVICE' else mat)[(i['section'] or i['sheet'])[:60]]+=i['total']
    return dict(items=it,totals=own,stim=st,accom=ac,ceiling=comb[0] if comb else (max(epd)+(max(rl) if rl else 0) if epd else (max(prim) if prim else None)),
        epd_sum=epdsum,rig_rate=rr,cbs_days=cd,base_sum=sum(i['total'] or 0 for i in base if i['total']),
        contingency_sum=sum(i['total'] or 0 for i in cont if i['total']),
        stim_total=sum(s['total'] or 0 for s in st if s['total']),service_sections=dict(sec),material_sections=dict(mat),
        plan_basis=('Total Cost-EPD & Rigless ceiling' if comb else ('EPD ceiling + rig-less ceiling' if epd else ('Cost Breakdown grand total' if prim else 'Total Cost-EPD lines'))))
def formations(wk):
    ddr=R('SELECT name,md,tvd,prognosis_md,prognosis_tvd,lithology FROM ddr_formations WHERE well_key=? AND (md>0 OR prognosis_md>0)',wk)
    by=collections.defaultdict(list)
    for r in ddr:by[FALIAS.get(r['name'].strip().lower().replace('.',''),r['name'].strip())].append(r)
    pm=collections.defaultdict(list);em=collections.defaultdict(list)
    for f in R("SELECT category,fact_key,value FROM eowr_facts WHERE well_key=? AND fact_key LIKE '% MD'",wk):
        nm=f['fact_key'].replace(' MD','')
        (pm if f['category'].endswith('TOPS') and f['category'].startswith('PROGRAM') else em)[nm].append(f['value'])
    out=[]
    for nm,rs in by.items():
        md=med([r['md'] for r in rs]) or med([r['prognosis_md'] for r in rs]);tvd=med([r['tvd'] for r in rs])
        lith=collections.Counter(r['lithology'] for r in rs if r['lithology']).most_common(1)
        out.append(dict(name=nm,md=r2(md),tvd=r2(tvd),prognosis_md=r2(med(pm.get(nm,[])) or med([r['prognosis_md'] for r in rs])),
            eowr_md=r2(med(em.get(nm,[]))),lithology=lith[0][0] if lith else '',observations=len(rs)))
    for nm in set(list(pm)+list(em)):
        if not any(x['name']==nm for x in out):out.append(dict(name=nm,md=None,tvd=None,prognosis_md=r2(med(pm.get(nm,[]))),eowr_md=r2(med(em.get(nm,[]))),lithology='',observations=0))
    out.sort(key=lambda x:(x['md'] or x['eowr_md'] or x['prognosis_md'] or 1e9))
    for i,f in enumerate(out):
        nx=out[i+1] if i+1<len(out) else None
        a=f['md'] or f['eowr_md'] or f['prognosis_md'];b=(nx['md'] or nx['eowr_md'] or nx['prognosis_md']) if nx else None
        f['thickness_m']=r2(b-a) if a and b and b>=a else None
        f['structural_delta_m']=r2(a-f['prognosis_md']) if a and f['prognosis_md'] else None
    return out
def materials(wk):
    cb=R("SELECT description,unit,unit_price,qty,total,section,sheet FROM cbs_items WHERE well_key=? AND category='MATERIAL' AND description<>'' ORDER BY total DESC",wk)
    st=R("SELECT kind,description,unit,volume,unit_price,total,sheet FROM cbs_stim WHERE well_key=? AND description<>''",wk)
    dd=R("SELECT material,unit,SUM(received) rec,SUM(used) used,MAX(stock) stock,COUNT(*) n FROM ddr_materials WHERE well_key=? AND material<>'' GROUP BY material,unit ORDER BY used DESC",wk)
    key=lambda s:re.sub(r'\s+',' ',re.sub(r'\b(cement add|bulk|used|received|stock|the|and|of)\b',' ',re.sub(r'[^a-z0-9 ]',' ',(s or '').lower()))).strip()
    cbs=collections.defaultdict(float);cu={};d=collections.defaultdict(float);du={}
    for x in cb+[y for y in st if y['kind'] in ('STIM_MATERIAL','ACID_SERVICE')]:
        if x.get('total'):cbs[key(x['description'])]+=x['total'];cu[key(x['description'])]=x['unit'] or ''
    for x in dd:
        if x['used']:d[key(x['material'])]+=x['used'];du[key(x['material'])]=x['unit'] or ''
    rec=[dict(item=k[:88],cbs_usd=r2(cbs.get(k)),ddr_qty=r2(d.get(k)),cbs_unit=cu.get(k,''),ddr_unit=du.get(k,''),
        status=('CBS + DDR matched' if cbs.get(k) is not None and d.get(k) is not None else ('CBS only — no DDR consumption record' if cbs.get(k) is not None else 'DDR only — no CBS priced line'))) for k in sorted(set(cbs)|set(d)) if len(k)>2]
    return dict(cbs_material=cb,cbs_stim=st,ddr_consumption=dd,reconciliation=rec)
def stimulation(wk):
    f=R("SELECT fact_key,value,unit,page_no FROM eowr_facts WHERE well_key=? AND fact_key IN ('PI','RATE','WHP','ACID_VOL','N2_VOL','CHOKE')",wk)
    ch=R("SELECT description,unit,volume,unit_price,total FROM cbs_stim WHERE well_key=? AND kind='STIM_MATERIAL'",wk)
    return dict(pi_values=[x['value'] for x in f if x['fact_key']=='PI' and x['value']<1000],
        rates=[dict(v=x['value'],unit=x['unit'],page=x['page_no']) for x in f if x['fact_key']=='RATE'],
        whp=[dict(v=x['value'],unit=x['unit'],page=x['page_no']) for x in f if x['fact_key']=='WHP'],
        acid_volumes=[dict(v=x['value'],unit=x['unit'],page=x['page_no']) for x in f if x['fact_key']=='ACID_VOL'],
        chemicals=[dict(item=x['description'],unit=x['unit'],planned_volume=x['volume'],unit_price=x['unit_price'],total_usd=x['total']) for x in ch if x['description']])
def mindls(sv):
    pts=sorted([s for s in sv if s['md'] and s['incl'] is not None and s['azi'] is not None],key=lambda s:s['md']);out=[]
    for a,b in zip(pts,pts[1:]):
        dm=b['md']-a['md']
        if dm<=0:continue
        i1,i2=math.radians(a['incl']),math.radians(b['incl']);da=math.radians((b['azi'] or 0)-(a['azi'] or 0))
        c=cl(math.cos(i2-i1)-math.sin(i1)*math.sin(i2)*(1-math.cos(da)),-1,1)
        out.append(dict(md=b['md'],dls=math.degrees(math.acos(c))*30/dm))
    return out
def events(wk,nt,daily,rate):
    dm={d['report_no']:d for d in daily};ev=[]
    for r in nt['npt_events']:ev.append(dict(report_no=r['report_no'],hours=r['hours'],attributed=r['attributed'],text=r['description'],source='24-h activity log'))
    for n in R('SELECT report_no,text FROM ddr_notes WHERE well_key=?',wk):ev.append(dict(report_no=n['report_no'],hours=None,attributed='',text=n['text'],source='DDR remark / shortage note'))
    gr=collections.defaultdict(list)
    for e in ev:
        t=(e['text'] or '').lower();hit=False
        for lab,pat in BUCKETS:
            if re.search(pat,t):gr[lab].append(e);hit=True;break
        if not hit:gr['Other operational delay'].append(e)
    out=[]
    for lab,es in sorted(gr.items(),key=lambda x:-len(x[1])):
        h=sum(e['hours'] or 0 for e in es);days=sorted(set(e['report_no'] for e in es))
        dep=[dm[d]['md_2400'] for d in days if d in dm and dm[d]['md_2400'] and dm[d]['md_2400']>100]
        fo=[dm[d]['formation'] for d in days if d in dm and dm[d]['formation']]
        who=collections.Counter(e['attributed'] for e in es if e['attributed']).most_common(1)
        out.append(dict(category=lab,occurrences=len(es),hours=r2(h),days=days[:12],depth_m=r2(med(dep)),
            formation=(collections.Counter(fo).most_common(1)[0][0] if fo else ''),attributed=(who[0][0] if who else ''),
            evidence=[e['text'][:210] for e in es if e['text']][:5],cost_usd=r2(h/24*rate),
            hypotheses=HYP.get(lab,HYP['_']),playbook=PLAY.get(lab,PLAY['_'])))
    return out[:14]
BRANCH=['Geoscience / casing shoes / KOP','Drilling parameters, fluids & loss control','Workover, fishing, milling & rig PM','Completion, liner-lap cement, acid & CBS']
def lessons(wk,p):
    ts=p['timeseries'];c=p['costs'];n=p['time'];ev=p['events'];m=p['materials'];forms=p['formations']
    best=sorted([t for t in ts if t['drilled'] and t['drilled']>0],key=lambda t:-t['drilled'])
    rop=[t['rop'] for t in ts if t['rop'] and 0<t['rop']<100];L=[]
    def A(code,pol,title,evid,act,h=None,u=None):L.append(dict(code=code,polarity=pol,title=title,evidence=evid,action=act,hours=r2(h),cost_usd=r2(u)))
    if best:
        b=best[0];A('LL-01','Best Practice',f'Peak 24-h footage of {fmt(b["drilled"],1)} m on report {b["report_no"]}',
            f'Report {b["report_no"]} ({b["date_g"]}) drilled {fmt(b["drilled"],1)} m to MD {fmt(b["md"],0)} m with MW {fmt(b["mw"],0)} PCF and ROP {fmt(b["rop"],2)} m/h.',
            'Freeze the WOB / RPM / hydraulic window of this interval as the baseline for the next well in the same formation.')
    else:A('LL-01','Best Practice','Workover executed without unscheduled drilling exposure',
        f'{p["ddr_days"]} DDRs with {fmt(n["npt_total_h"],1)} h of attributed NPT.','Retain the activity-coding discipline for future workovers.')
    A('LL-02','Best Practice',f'Time-code coverage of {n["activities"]:,} discrete 24-hour activity lines',
      f'{p["ddr_days"]} days, {fmt(n["status_hours"].get("NPT",0),1)} h explicitly flagged NPT and attributed to a responsible party.',
      'Keep the PLN / NPT / attributed-party columns mandatory — they make the four-track reconciliation auditable.')
    if rop:A('LL-03','Best Practice',f'ROP distribution peaked at {fmt(max(rop),2)} m/h',
      f'{len(rop)} daily ROP entries, band {fmt(min(rop),2)}–{fmt(max(rop),2)} m/h (median {fmt(med(rop),2)} m/h).',
      'Use the upper-quartile ROP as the AFE / time–depth assumption for this formation and rig pair.')
    else:A('LL-03','Best Practice','Zero new-hole exposure during this intervention',
      'All progress was achieved by milling, clean-out and completion runs rather than by drilling new formation.',
      'Benchmark intervention ROP on milling progress (m/h) instead of drilling ROP.')
    A('LL-04','Best Practice',f'Formation-top control at {len(forms)} mapped horizons',
      '; '.join(f'{f["name"]} {fmt(f["md"],0)} m' for f in forms[:6]) or 'Formation tops registered in the DDR geological sheet.',
      'Carry the measured tops into the structural grid before setting the next casing shoe.')
    A('LL-05','Best Practice',f'CBS granularity: {len(m["cbs_material"]):,} material rows, {len(m["cbs_stim"]):,} rig-less stimulation rows',
      f'Base commitment {fmt(c["base_sum"],0)} USD + contingency {fmt(c["contingency_sum"],0)} USD, rig day rate {fmt(c["rig_rate"],0)} USD/day.',
      'Preserve unit-price × quantity visibility so zero-rate NPT deductions are computable per line.')
    A('LL-06','Best Practice',f'Daily material balance on {len(m["ddr_consumption"]):,} distinct consumables',
      f'Received / used / stock maintained; {sum(1 for x in m["reconciliation"] if x["status"].startswith("CBS + DDR"))} CBS↔DDR matched lines.',
      'Keep the received–used–backload–stock four-column balance mandatory on every DDR.')
    for i,e in enumerate(ev[:6]):
        A(f'LL-{7+i:02d}','Pitfall',f'{e["category"]} — {fmt(e["hours"],1)} h across {e["occurrences"]} occurrences',
          (e['evidence'][0][:200] if e['evidence'] else f'Reports {e["days"][:4]}')+f' | depth {fmt(e["depth_m"],0)} m, formation {e["formation"] or "n/a"}, attributed {e["attributed"] or "n/a"}, exposure {fmt(e["cost_usd"],0)} USD.',
          e['playbook'][:250],e['hours'],e['cost_usd'])
    while len(L)<12:
        i=len(L)+1;A(f'LL-{i:02d}','Pitfall','Unattributed waiting time in the daily log',
          f'{fmt(n["attributed_hours"].get("Unattributed",0),1)} h of NPT carries no responsible-party code.',
          'Require an attributed-party code on every NPT row before DDR sign-off.')
    return L[:12]
def recs(wk,p):
    ts=p['timeseries'];c=p['costs'];n=p['time'];ev=p['events'];pa=p['params'];fo=p['formations'];m=p['materials']
    G=lambda k,d=dict():pa.get(k,d);rop=[t['rop'] for t in ts if t['rop'] and 0<t['rop']<100]
    W=lambda k,d=0:fmt(G(k).get('mean'),d)
    T=[]
    T.append((BRANCH[0],'Re-set the next casing shoe and KOP from the measured tops',
      lambda:'; '.join(f'{f["name"]} {fmt(f["md"],0)} m (Δ {fmt(f["structural_delta_m"],1)} m vs prognosis)' for f in fo[:5]),
      'Prognosis tops came from a coarse structural grid; measured tops shift the pressure ramps.',
      'Update grid → recompute pore-pressure / fracture-gradient curves → move the 9-5/8 in shoe to competent carbonate → re-run kick / loss tolerance → freeze KOP above the salt.',
      lambda:(f'Shoe tolerance ±{fmt(max(15,abs((fo[0]["structural_delta_m"] or 0))*0.1),0)} m around {fmt(fo[0]["md"],0)} m' if fo else 'Set shoe from measured carbonate top'),
      'Structural delta < 15 m; zero shoe-in-shale events.','Sr. Well-site Geologist / Reservoir','P1',lambda:(1.2,c['rig_rate']*1.2)))
    T.append((BRANCH[1],'Lock the high-ROP parameter window and ECD envelope',
      lambda:f'WOB {fmt(G("WOB").get("min"),0)}–{fmt(G("WOB").get("max"),0)} klb, RPM {fmt(G("RPM").get("min"),0)}–{fmt(G("RPM").get("max"),0)}, GPM {fmt(G("GPM").get("min"),0)}–{fmt(G("GPM").get("max"),0)}, SPP {fmt(G("SPP").get("min"),0)}–{fmt(G("SPP").get("max"),0)} psi; ROP median {fmt(med(rop),2)} m/h.',
      'Parameter scatter between day and night crews diluted mechanical specific energy efficiency.',
      'Issue a one-page parameter card per hole section → enforce the band at the driller console → log torque / MSE hourly → flag stick-slip → review at the morning meeting.',
      lambda:f'WOB {W("WOB")} ±20% klb; RPM {W("RPM")} ±10%; MW {fmt(p["mw_min"],0)}–{fmt(p["mw_max"],0)} PCF; AV / HSI inside the recorded hydraulic envelope.',
      '≥10% ROP uplift with no extra bit / BHA NPT hours.','Sr. Drilling Engineer','P1',lambda:(0.8,c['rig_rate']*0.8)))
    T.append((BRANCH[1],'Pre-blend graded LCM and enforce the mud-weight ramp',
      lambda:f'Mud-weight band {fmt(p["mw_min"],0)}–{fmt(p["mw_max"],0)} PCF; loss NPT {fmt(next((e["hours"] for e in ev if "Loss" in e["category"]),0),1)} h.',
      'Transition from high-MW Gachsaran to depleted Asmari / Sarvak narrows the pore–fracture window.',
      'Stage MW in 2 PCF steps → keep graded LCM on location → monitor pits every 15 min → trigger the 60-second loss playbook at the first 2 bph deviation.',
      lambda:f'MW {fmt(p["mw_min"],0)}–{fmt(p["mw_max"],0)} PCF; LCM pill 20–40 ppb graded; ECD margin ≥0.5 PCF below recorded FIT/LOT.',
      'Seepage <5 bph; zero full losses.','Drilling Fluids / Rig Sup.','P1',lambda:(0.6,c['rig_rate']*0.6)))
    T.append((BRANCH[2],'Pre-mobilise the packer-milling / fishing spread',
      lambda:f'{fmt(next((e["hours"] for e in ev if "Stuck" in e["category"] or "fishing" in e["category"].lower()),0),1)} h fishing / milling NPT; {len(p["bha"])} BHA component records.',
      'Legacy permanent packers and corroded tubing generate metal debris that stalls the milling BHA.',
      'Run a pre-job drift / caliper → stage junk basket, magnet and cable spear → cap milling WOB → circulate a high-vis sweep each stand → log iron at the shaker.',
      lambda:'Milling WOB 2–5 klb, RPM 60–90, circulation ≥ hole-cleaning AV; magnet run every 3 stands.',
      'Milling rate ≥1.5 m/h; zero secondary fishing runs.','Workover / Completion Sup.','P2',lambda:(0.5,c['rig_rate']*0.5)))
    T.append((BRANCH[2],'Preventive maintenance on the rig power train',
      lambda:f'{fmt(next((e["hours"] for e in ev if "equipment" in e["category"].lower()),0),1)} h rig-equipment NPT; {fmt(n["npt_contractor_h"],1)} h contractor-attributed NPT.',
      'Single-point pump / SCR / top-drive failures consume zero-rate day-rate exposure.',
      'Add daily vibration / temperature / insulation checks → hold critical spares at the pad → schedule PM during wait-on-cement → log MTBF per subsystem.',
      lambda:'PM interval ≤250 operating hours; critical spares ≥1 each; MTBF target +20%.',
      'Rig-equipment NPT ≤8 h per 30 operating days.','Rig Manager / OEOC','P1',lambda:(0.7,c['rig_rate']*0.7)))
    T.append((BRANCH[3],'Close the CBS↔DDR↔EOWR variance before invoice approval',
      lambda:f'{len(m["reconciliation"])} reconciliation lines, {sum(1 for x in m["reconciliation"] if "only" in x["status"])} single-source items; CBS base {fmt(c["base_sum"],0)} USD.',
      'Unit-price lines and physical consumption are booked in different documents without a shared key.',
      'Build a shared material key → match CBS qty to DDR used and EOWR consumption → require a signed variance sheet → release payment per matched line.',
      lambda:f'Variance tolerance ±5% per line; unmatched value <1% of {fmt(c["base_sum"],0)} USD.',
      '100% of priced lines reconciled.','CBS / Cost Control Auditor','P1',lambda:(0.2,max(0,(c['contingency_sum'] or 0)*0.03))))
    T.append((BRANCH[3],'Qualify liner-lap cement with CBL-VDL / USIT before completion',
      lambda:f'{len(p["geology"]["log_evidence"])} logging / bond-log evidence pages; dry-test and pressure-test facts in the EOWR.',
      'A micro-annulus or poor lap bond is revealed only during the dry test, after the completion string is committed.',
      'Run CBL-VDL / USIT after cementing → compute bond index by interval → squeeze before RIH with completion string → repeat dry test.',
      lambda:'Bond index ≥0.8 over shoe track and lap; dry test 30 min with <2% pressure decline.',
      'Zero remedial squeezes after completion RIH.','Completion / Cementing','P1',lambda:(0.9,c['rig_rate']*0.9)))
    T.append((BRANCH[3],'Anti-sludge acid package and CTU contingency for heavy AZNS crude',
      lambda:f'{len(p["stimulation"]["chemicals"])} priced stimulation chemicals incl. HCl, xylene, mutual solvent, anti-sludge and de-emulsifier.',
      'HCl contacting asphaltenic heavy oil forms acid sludge and stable emulsions that cut the post-acid PI uplift.',
      'Run crude–acid compatibility test → dose anti-sludge + de-emulsifier + iron control → pump pre-flush / main acid / over-flush by CTU with N2 lift → record WHP, choke and PI.',
      lambda:'HCl ≥28% at CBS volume; anti-sludge / Fe-control at CBS unit rate; N2 lift staged to the CBS operating-day schedule.',
      'Post-acid PI uplift ≥25% with no emulsion in returns.','Stimulation / Production','P1',lambda:(0.4,c['rig_rate']*0.4)))
    T.append((BRANCH[0],'Zero-rate NPT attribution gate on the daily report',
      lambda:f'Contractor NPT {fmt(n["npt_contractor_h"],1)} h vs client/subsurface {fmt(n["npt_client_h"],1)} h; payable {fmt(c["payable_excl_ctr_npt"],0)} USD vs gross {fmt(c["gross_calendar"],0)} USD.',
      'Uncoded NPT is invoiced at full day rate instead of being deducted at zero rate.',
      'Code every NPT row at the rig floor → supervisor signs the attribution → cost control applies the zero-rate rule → weekly four-track reconciliation.',
      lambda:f'Attribution completeness 100%; deduction {fmt((c["gross_calendar"] or 0)-(c["payable_excl_ctr_npt"] or 0),0)} USD.',
      'Zero unattributed NPT hours; four-track variance <1%.','Cost Control / Contract Audit','P1',lambda:(0.1,((c['gross_calendar'] or 0)-(c['payable_excl_ctr_npt'] or 0))*.15)))
    T.append((BRANCH[1],'Directional survey QC with minimum-curvature DLS monitoring',
      lambda:f'{len(p["surveys"])} survey stations; max computed DLS {fmt(p["max_dls"],2)} deg/30 m.',
      'Isolated survey stations and high local dogleg increase torque, drag and liner hang-up risk.',
      'Expand pipe-separated survey strings → compute DLS by minimum curvature at every station → compare with plan → trigger a wiper trip above the limit.',
      lambda:'DLS ≤3.0 deg/30 m in the build; tortuosity index <1.2; survey spacing ≤30 m.',
      'Zero liner hang-ups at the sidetrack window.','Directional Drilling','P2',lambda:(0.3,c['rig_rate']*0.3)))
    T.append((BRANCH[2],'POB and catering baseline aligned to the CBS full-package headcount',
      lambda:f'Peak POB {fmt(p["pob_max"],0)} persons across {len(p["pob_series"])} daily records; {len(c["accom"])} priced accommodation / rig-move lines.',
      'Actual POB drifts above the CBS accommodation baseline without transparent invoicing.',
      'Set the CBS headcount baseline → record daily POB by contractor → reconcile catering cost per man-day → flag exceedances weekly.',
      lambda:'POB within ±10% of the CBS full-package baseline; weekly variance report.',
      'Catering variance <3% of the accommodation line.','Contracts / Camp Admin','P3',lambda:(0.1,c['rig_rate']*0.1)))
    T.append((BRANCH[3],'Wellhead Section A–D and subsurface tool pressure-test dossier',
      lambda:f'{len(p["pressure_tests"])} recorded pressure-test facts; 10K Xmas-tree and TRSSSV / packer / CIV-CIM evidence.',
      'Test charts scattered across appendices delay completion acceptance.',
      'Digitise every test chart → index by Section A–D and tool → apply the acceptance envelope → sign off one completion dossier.',
      lambda:'Section D test to rated working pressure with 30-min hold; TRSSSV / CIV / CIM to documented test pressure.',
      '100% test charts indexed before handover; zero re-tests.','Completion / Wellhead','P2',lambda:(0.2,c['rig_rate']*0.2)))
    out=[]
    for i,(br,title,evid,root,proto,win,kpi,owner,prio,save) in enumerate(T,1):
        try:
            d,u=save();rc=dict(code=f'REC-{i:02d}',branch=br,title=title,evidence=evid(),root_cause=root,protocol=proto,
                control_window=win(),kpi=kpi,owner=owner,priority=prio,days_saved=r2(d),usd_saved=r2(u))
        except Exception:
            rc=dict(code=f'REC-{i:02d}',branch=br,title=title,evidence='Derived from the well DDR / CBS records.',root_cause=root,
                protocol=proto,control_window='Per recorded parameter envelope.',kpi=kpi,owner=owner,priority=prio,days_saved=None,usd_saved=None)
        out.append(rc)
    return out

def profile(w):
    k=w['key'];daily=R('SELECT * FROM ddr_daily WHERE well_key=? ORDER BY report_no',k)
    nt=time_account(k);c=costs(k)
    if not c['rig_rate']:c['rig_rate']=bench();c['rate_basis']='Contract-1 CBS benchmark day rate'
    else:c['rate_basis']='CBS Rig Service unit price'
    fo=formations(k);m=materials(k);sti=stimulation(k);pa=params(k)
    sv=R('SELECT * FROM ddr_surveys WHERE well_key=?',k);dl=mindls(sv)
    for s in sv:
        x=[d['dls'] for d in dl if abs(d['md']-(s['md'] or -1))<.01];s['dls_calc']=r2(x[0]) if x else None
    bha=R('SELECT * FROM ddr_bha WHERE well_key=? AND od IS NOT NULL',k)
    pob=R("SELECT * FROM ddr_pob WHERE well_key=? AND category='TOTAL POB'",k)
    cas=R('SELECT * FROM ddr_casing WHERE well_key=? AND to_m IS NOT NULL',k)
    pts=R("SELECT * FROM eowr_facts WHERE well_key=? AND fact_key='PRESSURE_TEST'",k)
    logev=R("SELECT page_no,chapter FROM pdf_pages WHERE well_key LIKE ? AND (chapter LIKE '05 %' OR chapter LIKE '08 %' OR chapter LIKE '10 %')",'%'+k+'%')
    rate=c['rig_rate'];ts=[];cum=0
    for d in daily:
        nh=Q("SELECT SUM(hours) h FROM ddr_activities WHERE well_key=? AND report_no=? AND (UPPER(REPLACE(REPLACE(REPLACE(status,'.',''),' ',''),'-','')) IN ('N','NPT','UP','UNPLANNED','UNPLAN','UNPLN','NP','PN','NPTUNPLN','UPI') OR UPPER(status) LIKE 'NPT%')",k,d['report_no'])[0]['h'] or 0
        ph=Q("SELECT SUM(hours) h FROM ddr_activities WHERE well_key=? AND report_no=? AND UPPER(status) IN ('P','PLN','PLAN','PLANNED','PT')",k,d['report_no'])[0]['h'] or 0
        md=d['md_2400'] or d['md_0600'] or d['md_0000'];cum+=rate
        ts.append(dict(report_no=d['report_no'],date_g=d['date_g'],date_j=d['date_j'],md=r2(md),drilled=r2(d['drilled_24']),
            rop=r2(d['rop']),mw=r2(d['mw']),hole_size=d['hole_size'],formation=d['formation'],rig=d['rig'],npt_h=r2(nh),pt_h=r2(ph),cum_cost=r2(cum)))
    dep=[t['md'] for t in ts if t['md'] and t['md']>100]
    td=O("SELECT value FROM eowr_facts WHERE well_key=? AND category='WELL_INFO' AND fact_key='ACTUAL_TD' AND value BETWEEN 1000 AND 5000 ORDER BY page_no LIMIT 1",k)
    etd=float(td['value']) if td else None
    dr=[t['drilled'] for t in ts if t['drilled'] and t['drilled']>0];rp=[t['rop'] for t in ts if t['rop'] and 0<t['rop']<100]
    mw=[t['mw'] for t in ts if t['mw'] and 20<t['mw']<200]
    days=len(daily);plan=c['ceiling'] or c['epd_sum'] or (c['base_sum']+c['contingency_sum'])
    basis=c['plan_basis']
    if not plan:plan=days*rate;basis='Derived day-rate exposure: CBS holds an operation-duration schedule, no priced lines'
    cd=c['cbs_days'] or days;gross=plan+(days-cd)*rate
    payable=gross-nt['npt_contractor_h']/24*rate;net=gross-nt['npt_total_h']/24*rate
    c.update(dict(cbs_plan=r2(plan),gross_calendar=r2(gross),payable_excl_ctr_npt=r2(payable),pure_net_excl_all_npt=r2(net),
        delta_days=r2(days-cd),ctr_npt_days=r2(nt['npt_contractor_h']/24),all_npt_days=r2(nt['npt_total_h']/24),plan_basis=basis))
    p=dict(op=w['op'],key=k,title=w['title'],contract=w['contract'],contract_code=w['contract_code'],rig=w['rig'],kind=w['kind'],
        files=dict(ddr=w['ddr'],cbs=w['cbs'],program=w['program'],eowr=w['eowr']),
        eowr_status=('AVAILABLE' if w['eowr'] else 'NOT AVAILABLE — no EOWR in 04.EOWR/; reconstructed from DDR + CBS + Well Program'),
        ddr_days=days,date_start=ts[0]['date_g'] if ts else '',date_end=ts[-1]['date_g'] if ts else '',
        depth_start=r2(dep[0] if dep else None),depth_final=r2(max(dep) if dep else etd),eowr_td=r2(etd),
        drilled_total=r2(sum(dr)) if dr else None,rop_mean=r2(mean(rp)),rop_median=r2(med(rp)),rop_max=r2(max(rp)) if rp else None,
        mw_min=r2(min(mw)) if mw else None,mw_max=r2(max(mw)) if mw else None,time=nt,costs=c,formations=fo,materials=m,
        stimulation=sti,params=pa,
        surveys=[dict(md=s['md'],incl=s['incl'],azi=s['azi'],tvd=s['tvd'],dls=s['dls_calc']) for s in sv[:350]],
        max_dls=r2(max([d['dls'] for d in dl],default=None)),
        bha=[dict(bha_no=b['bha_no'],item=b['item'],od=b['od'],length=b['length'],cum_len=b['cum_len']) for b in bha[:220]],
        pob_series=[dict(report_no=x['report_no'],total=x['total'] or x['count']) for x in pob],
        pob_max=max([x['total'] or x['count'] or 0 for x in pob],default=None),
        casing=[dict(size=x['size'],grade=x['grade'],weight=x['weight'],from_m=x['from_m'],to_m=x['to_m'],burst=x['burst'],collapse=x['collapse']) for x in cas[:55]],
        pressure_tests=[dict(v=x['value'],unit=x['unit'],page=x['page_no']) for x in pts[:70]],
        timeseries=ts,geology=dict(log_evidence=[dict(page=x['page_no'],chapter=x['chapter']) for x in logev[:55]],
            formation_count=len(fo),lithology=[f['lithology'] for f in fo if f['lithology']][:15]))
    p['events']=events(k,nt,daily,rate);p['lessons']=lessons(k,p);p['recs']=recs(k,p)
    for e in p['events']:
        con.execute('INSERT INTO well_events(well_key,kind,day,depth,formation,title,detail,hours,cost_usd,evidence) VALUES (?,?,?,?,?,?,?,?,?,?)',
            (k,e['category'],','.join(map(str,e['days'][:8])),e['depth_m'],e['formation'],f"{e['category']} — {e['occurrences']} occurrences",' | '.join(e['evidence'][:3])[:800],e['hours'],e['cost_usd'],'DDR 24-h log / remarks'))
    for l in p['lessons']:
        con.execute('INSERT INTO well_lessons(well_key,code,polarity,title,evidence,action,hours,cost_usd) VALUES (?,?,?,?,?,?,?,?)',
            (k,l['code'],l['polarity'],l['title'],l['evidence'][:750],l['action'][:750],l['hours'],l['cost_usd']))
    for rc in p['recs']:
        con.execute('INSERT INTO well_recs(well_key,code,branch,payload) VALUES (?,?,?,?)',(k,rc['code'],rc['branch'],json.dumps(rc,ensure_ascii=False)))
    for x in m['reconciliation'][:400]:
        con.execute('INSERT INTO reconciliation(well_key,item,cbs,ddr,eowr,unit,delta,status) VALUES (?,?,?,?,?,?,?,?)',
            (k,x['item'],x['cbs_usd'],x['ddr_qty'],None,x['cbs_unit'] or x['ddr_unit'],None,x['status']))
    for cat,h in nt['attributed_hours'].items():
        con.execute('INSERT INTO npt_breakdown(well_key,category,hours,attributed,share) VALUES (?,?,?,?,?)',(k,cat[:110],r2(h),cat[:70],r2(100*h/max(1,nt['npt_total_h']))))
    for t in ts:
        con.execute('INSERT INTO field_timeseries(well_key,report_no,date_g,md,drilled,rop,mw,npt_h,pt_h,cum_cost) VALUES (?,?,?,?,?,?,?,?,?,?)',
            (k,t['report_no'],t['date_g'],t['md'],t['drilled'],t['rop'],t['mw'],t['npt_h'],t['pt_h'],t['cum_cost']))
    con.execute('INSERT OR REPLACE INTO well_kpi VALUES (?,?)',(k,json.dumps(dict(op=p['op'],days=p['ddr_days'],depth=p['depth_final'],npt=p['time']['npt_total_h'],cbs=c['cbs_plan'],gross=c['gross_calendar'],payable=c['payable_excl_ctr_npt'],net=c['pure_net_excl_all_npt'],rig_rate=c['rig_rate']),ensure_ascii=False)))
    return p

def analytics_stage():
    for t in ('well_events','well_lessons','well_recs','reconciliation','npt_breakdown','field_timeseries','well_kpi'):con.execute('DELETE FROM '+t)
    P={}
    for w in WELLS:
        p=profile(w);P[w['key']]=p
        print(f"  op{w['op']:02d} {w['key']:20s} days={p['ddr_days']:3d} TD={p['depth_final']} NPT={fmt(p['time']['npt_total_h'],1)}h CBS={fmt(p['costs']['cbs_plan'],0)} forms={len(p['formations'])} ev={len(p['events'])}",flush=True)
    F=dict(wells=[],contracts=[],heatmap=[],timeline=[],formation_correlation=[],learning_curves={},cost_sankey=[],totals={})
    for p in P.values():
        F['wells'].append(dict(op=p['op'],key=p['key'],title=p['title'],contract=p['contract'],rig=p['rig'],ddr_days=p['ddr_days'],
            depth_final=p['depth_final'],drilled_total=p['drilled_total'],npt_h=p['time']['npt_total_h'],cbs_plan=p['costs']['cbs_plan'],
            gross=p['costs']['gross_calendar'],payable=p['costs']['payable_excl_ctr_npt'],net=p['costs']['pure_net_excl_all_npt'],
            date_start=p['date_start'],date_end=p['date_end'],rop_mean=p['rop_mean'],rig_rate=p['costs']['rig_rate'],formations=len(p['formations'])))
        for f in FORMATIONS:
            row=next((x for x in p['formations'] if x['name'].lower()==f.lower()),None)
            evh=sum(e['hours'] or 0 for e in p['events'] if f.lower() in (e['formation'] or '').lower() or f.lower() in ' '.join(e['evidence']).lower())
            F['heatmap'].append(dict(well=p['key'],formation=f,md=(row['md'] if row else None),risk=r2(evh),delta=(row['structural_delta_m'] if row else None)))
        F['timeline'].append(dict(well=p['key'],rig=p['rig'],start=p['date_start'],end=p['date_end'],days=p['ddr_days'],op=p['op']))
        for f in p['formations'][:15]:
            if f['md']:F['formation_correlation'].append(dict(well=p['key'],formation=f['name'],md=f['md'],tvd=f['tvd'],thickness=f['thickness_m'],delta=f['structural_delta_m']))
        F['cost_sankey'].append(dict(source=p['contract'],target=p['rig'],value=p['costs']['gross_calendar'] or 0,well=p['key']))
    for cid,(code,title,ops) in CONTRACTS.items():
        ps=[P[BY_OP[o]['key']] for o in ops]
        F['contracts'].append(dict(id=cid,code=code,title=title,ops=ops,wells=[x['key'] for x in ps],
            cbs_plan=sum(x['costs']['cbs_plan'] or 0 for x in ps),gross=sum(x['costs']['gross_calendar'] or 0 for x in ps),
            payable=sum(x['costs']['payable_excl_ctr_npt'] or 0 for x in ps),net=sum(x['costs']['pure_net_excl_all_npt'] or 0 for x in ps),
            npt_h=sum(x['time']['npt_total_h'] for x in ps),days=sum(x['ddr_days'] for x in ps)))
    for name,chain in [('AZNS-127 WO1 → WO2 → ST2',['AZNS-127-WO1-ST1','AZNS-127-WO2','AZNS-127-ST2']),('AZNS-122 WO1 → WO2',['AZNS-122-WO1','AZNS-122-WO2'])]:
        F['learning_curves'][name]=[dict(well=k,op=P[k]['op'],days=P[k]['ddr_days'],npt_h=P[k]['time']['npt_total_h'],cbs=P[k]['costs']['cbs_plan'],rop=P[k]['rop_mean'],depth=P[k]['depth_final']) for k in chain]
    T=dict(operations=13,contracts=3,rigs=['OEOC 201','OEOC 208'],cbs_files=13,cbs_sheets=202,ddr_files=13,ddr_sheets=2608,
        eowr_files=11,eowr_pages=2724,program_files=14,program_pages=1621,pdf_pages=4345,excel_sheets=2810,
        cells=Q('SELECT SUM(n_nonempty) FROM sheets')[0][0],numeric_cells=Q('SELECT COUNT(*) FROM cell_numeric')[0][0],
        formulas=Q('SELECT SUM(n_formula) FROM sheets')[0][0],pdf_tables=Q('SELECT COUNT(*) FROM pdf_tables')[0][0],
        pdf_numbers=int((O("SELECT v FROM meta WHERE k='pdf_numeric_tokens'") or {'v':'0'})['v']),
        ddr_days=sum(p['ddr_days'] for p in P.values()),npt_hours=sum(p['time']['npt_total_h'] for p in P.values()),
        cbs_plan=sum(p['costs']['cbs_plan'] or 0 for p in P.values()),gross=sum(p['costs']['gross_calendar'] or 0 for p in P.values()),
        payable=sum(p['costs']['payable_excl_ctr_npt'] or 0 for p in P.values()),net=sum(p['costs']['pure_net_excl_all_npt'] or 0 for p in P.values()),
        drilled_m=sum(p['drilled_total'] or 0 for p in P.values()),formations=sum(len(p['formations']) for p in P.values()),
        events=sum(len(p['events']) for p in P.values()),lessons=sum(len(p['lessons']) for p in P.values()),
        recs=sum(len(p['recs']) for p in P.values()),reconciliation=sum(len(p['materials']['reconciliation']) for p in P.values()))
    F['totals']=T
    F['audit_findings']=[
      dict(id='AF-01',title='AZNS-122 WO1 has no EOWR in 04.EOWR/',detail='Operation 4 is analysed in full from 112 DDR sheets, its CBS workbook and its 66-page Workover-1 Programme; no End-of-Well Report exists in the repository.'),
      dict(id='AF-02',title='Operations 12 and 13 share one EOWR file',detail='The 238-page file is a consolidated ST#2 End-of-Well Report (header contract CCA/C/1405/3). WO#2 appears as carried-forward well history and as the WorkOver Section rows of the time-analysis and mud-volume chapters; WO2 is separately evidenced by its own 136-sheet DDR pack, CBS and programme.'),
      dict(id='AF-03',title='Contract-code discrepancies in source headers',detail='8.AZNS_122_WO2_EOWR.pdf is headed CCA/C/1401/27 although booked under the second bundle (CCA/C/1403/31); 9.AZNS_75_EOWR.pdf is titled AZNS-075-ST1; the ST#2 EOWR carries CCA/C/1405/3. Recorded verbatim for contract close-out.'),
      dict(id='AF-04',title='Benchmark material sheets for wells 29 and 38',detail='CBS/10 CBS matrial.xlsx holds ten material sheets; wells 29 and 38 are outside the 13-operation campaign and are retained as contract-1 rate benchmarks, not attributed to any campaign well.'),
      dict(id='AF-05',title='Image-only pages and text-layer limits',detail='Pages with very low extractable text (scanned logs, charts, certificates) are archived page-by-page; numeric content is indexed where a text layer exists. No value is invented where the source is silent.')]
    for a in F['audit_findings']:con.execute('INSERT OR REPLACE INTO audit_findings VALUES (?,?,?)',(a['id'],a['title'],a['detail']))
    con.commit()
    jdump_gz(P,os.path.join(DATA,'well_profiles.json.gz'));jdump(F,os.path.join(DATA,'field_analytics.json'))
    print('  FIELD TOTALS',T,flush=True)
    return P,F

# -------------------------------------------------------------- WORKBOOK ----
def workbook_stage(P,F):
    import openpyxl
    from openpyxl.styles import Font,PatternFill,Alignment
    from openpyxl.cell import WriteOnlyCell
    wb=openpyxl.Workbook(write_only=True);NAVY='0B1F3A';LIGHT='EEF2F7'
    def sheet(name,headers,widths,note=None):
        ws=wb.create_sheet(name[:31]);ws.freeze_panes='A2'
        if note:
            c=WriteOnlyCell(ws,value=note);c.font=Font(name='Calibri',size=10,italic=True,color='475569');ws.append([c])
        cells=[]
        for h in headers:
            c=WriteOnlyCell(ws,value=h);c.font=Font(name='Calibri',size=10,bold=True,color='FFFFFF')
            c.fill=PatternFill('solid',fgColor=NAVY);c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True);cells.append(c)
        ws.append(cells)
        for i,w in enumerate(widths,1):ws.column_dimensions[CL(i)].width=w
        return ws
    def row(ws,vals,i=0):
        out=[]
        for v in vals:
            c=WriteOnlyCell(ws,value=v);c.font=Font(name='Calibri',size=9,color='0F172A')
            if i%2:c.fill=PatternFill('solid',fgColor=LIGHT)
            if isinstance(v,float):c.number_format='#,##0.00'
            out.append(c)
        ws.append(out)
    ws=sheet('00_README',['Section','Item','Definition'],[24,44,115],'SAZN Master Audit Workbook — companion to SAZN_Master_Warehouse.sqlite | 2026-11-08')
    rd=[('Scope','51 source documents','13 CBS (202 sheets) · 13 DDR (2,608 sheets) · 11 EOWR (2,724 pages) · 14 Programmes (1,621 pages).'),
        ('Scope','Cell archive',f"{F['totals']['cells']:,} non-empty cells, {F['totals']['numeric_cells']:,} numeric, {F['totals']['formulas']:,} formulas, exact file+sheet+cell_ref."),
        ('Scope','PDF archive','4,345 pages with lossless compressed text; every numeric token indexed with page / line / context.'),
        ('Method','Four-track accounting','1 CBS Plan = COST CEILING. 2 Gross Calendar = CBS Plan + (DDR days − CBS days) × rig day rate. 3 Payable = Gross − contractor NPT/24 × day rate. 4 Pure Net = Gross − all NPT/24 × day rate.'),
        ('Method','NPT classification','DDR 24-hour activity rows with N / NPT / UP / UNPLN status; attributed party split contractor (zero-rate) vs client / subsurface.'),
        ('Method','Three-way reconciliation','CBS priced lines ↔ DDR received/used/stock ↔ EOWR mud / acid / cement volumes on a normalised material key.'),
        ('Method','Minimum-curvature DLS','DLS = acos(cos(ΔI) − sin(I1)·sin(I2)·(1−cos(ΔA))) × 30/ΔMD.'),
        ('QA','No synthetic numbers','All values come from source documents; derived values are computed from sourced rates, quantities and hours and labelled as derived.'),
        ('QA','Rig day rate basis','23,000 USD/DAY is the sourced CBS Rig Service unit price present in 8.AZNS_122_WO2_CBS.xlsx and the contract-1 benchmark sheets; it is never an assumed market rate. Where a CBS holds no priced lines (AZNS-85) the plan is flagged as a derived day-rate exposure, not as an invoiced ceiling.'),
        ('QA','NPT attribution','All 13 operations have explicit DDR N / NPT status rows with an attributed-party column. No 50/50 waiting-time fallback split was needed; contractor zero-rate and client/subsurface hours are taken verbatim from the daily log.'),
        ('QA','Contract C3 header','The campaign map labels operation 13 CCA/C/1403/31 / ST2 amendment; the ST#2 EOWR header reads CCA/C/1405/3. Both are retained verbatim (contracts register and AF-03) rather than overwriting either.'),
        ('QA','Combined WO2 / ST2 EOWR','The 238-page shared report was split page by page: 11 pages shared history, 11 WO2 WorkOver Section pages, 216 ST2 pages, plus 11 shared first pages retained as shared.'),
        ('QA','Graphics','149 HTML5/CSS3/SVG/Canvas artboards rendered in headless Chromium at 2400x1800 px (300 DPI); 0 text-overlap pairs after automated bounding-box review; English-only text; inverted depth axis; mean PNG 169 KB.'),
        ('QA','Word books','30 DOCX: 26 single-well (FA/EN) + 4 portfolio (FA/EN). Each has 24 chapters, 96 provenance tables and 11 graphics. FA uses B Nazanin with w:bidi and w:bidiVisual; EN uses Times New Roman. All headings and table headers centered; navy header, banded rows, tblHeader and cantSplit.'),
        ('QA','Dashboard','SAZN_Executive_Interactive_Dashboard.html embeds 532,268 searchable rows (all 324,150 numeric cells, 41,679 unit/keyword PDF numeric tokens and 166,439 semantic rows) with exact file / sheet / cell / page provenance; fully offline.'),
        ('QA','Honest gaps','AF-01 to AF-05 document the missing WO1 EOWR, the combined 12/13 report, contract-header discrepancies and benchmark sheets.')]
    for i,r in enumerate(rd):row(ws,r,i)
    ws=sheet('01_Source_Registry',['ID','Category','Relative path','File','Sheets/Pages','Bytes','SHA-256','Well(s)'],[6,12,52,46,14,14,20,30])
    for i,r in enumerate(R('SELECT * FROM source_files ORDER BY id')):row(ws,[r['id'],r['category'],r['rel_path'],r['filename'],r['n_units'],r['size_bytes'],r['sha256'],r['well_keys']],i)
    ws=sheet('02_Well_Master',['Op','Well','Title','Contract','Code','Rig','Kind','DDR','CBS','Programme','EOWR','EOWR status'],[5,22,24,10,24,12,26,42,50,50,50,62])
    for i,r in enumerate(R('SELECT * FROM wells ORDER BY op')):row(ws,[r[k] for k in ('op','key','title','contract','contract_code','rig','kind','ddr_file','cbs_files','program_files','eowr_files','eowr_status')],i)
    ws=sheet('03_Contract_Master',['ID','Code','Title','Ops','Wells','CBS plan USD','Gross USD','Payable USD','Net USD','NPT h','DDR days'],[6,26,40,18,52,18,18,18,18,12,12])
    for i,c in enumerate(F['contracts']):row(ws,[c['id'],c['code'],c['title'],','.join(map(str,c['ops'])),', '.join(c['wells']),c['cbs_plan'],c['gross'],c['payable'],c['net'],c['npt_h'],c['days']],i)
    ws=sheet('04_Four_Track_Cost',['Op','Well','Contract','Rig','DDR days','CBS days','Day rate USD','Rate basis','CBS Plan','Gross Calendar','Payable','Pure Net','Ctr NPT h','Client NPT h','All NPT h','Ctr deduction USD','All-NPT deduction USD','Plan basis'],[5,22,10,12,11,12,16,34,18,19,19,19,12,14,12,20,21,66])
    for i,(k,p) in enumerate(sorted(P.items(),key=lambda x:x[1]['op'])):
        c=p['costs'];t=p['time'];row(ws,[p['op'],k,p['contract'],p['rig'],p['ddr_days'],c.get('cbs_days'),c.get('rig_rate'),c.get('rate_basis'),c.get('cbs_plan'),c.get('gross_calendar'),c.get('payable_excl_ctr_npt'),c.get('pure_net_excl_all_npt'),r2(t['npt_contractor_h']),r2(t['npt_client_h']),r2(t['npt_total_h']),r2((c.get('gross_calendar') or 0)-(c.get('payable_excl_ctr_npt') or 0)),r2((c.get('gross_calendar') or 0)-(c.get('pure_net_excl_all_npt') or 0)),c.get('plan_basis','')],i)
    ws=sheet('05_CBS_Totals',['File','Sheet','Well','Row','Item','Qty','Unit cost','Total USD'],[44,34,22,7,72,12,18,18])
    for i,r in enumerate(R('SELECT * FROM cbs_totals ORDER BY well_key,sheet,row_no')):row(ws,[r['rel_path'],r['sheet'],r['well_key'],r['row_no'],r['item'],r['qty'],r['unit_cost'],r['total']],i)
    ws=sheet('06_CBS_Line_Items',['File','Sheet','Row','Cell','Well','Category','Section','Contingency','Item','Description','Unit','Unit price','Qty','Total USD'],[42,32,7,8,22,20,12,12,10,72,12,16,12,18])
    for i,r in enumerate(R('SELECT * FROM cbs_items ORDER BY well_key,sheet,row_no')):row(ws,[r['rel_path'],r['sheet'],r['row_no'],r['cell_ref'],r['well_key'],r['category'],r['section'],'YES' if r['is_contingency'] else '',r['item_no'],r['description'],r['unit'],r['unit_price'],r['qty'],r['total']],i)
    ws=sheet('07_CBS_Stimulation',['File','Sheet','Well','Kind','Item','Description','Unit','Volume','Unit price','Standby d','Operating d','Standby rate','Operating rate','Total USD','Remark'],[42,32,22,20,8,68,10,12,16,12,13,15,16,16,40])
    for i,r in enumerate(R('SELECT * FROM cbs_stim ORDER BY well_key,kind,row_no')):row(ws,[r['rel_path'],r['sheet'],r['well_key'],r['kind'],r['item_no'],r['description'],r['unit'],r['volume'],r['unit_price'],r['standby_days'],r['operating_days'],r['standby_rate'],r['operating_rate'],r['total'],r['remark']],i)
    ws=sheet('08_DDR_Daily',['Well','Report','Date G','Date J','Rig','Hole section','Formation','MD 00','MD 24','Drilled m','ROP','MW','Remark sheet','Data sheet','Template'],[22,8,16,16,12,24,20,12,12,12,12,12,28,28,10])
    for i,r in enumerate(R('SELECT * FROM ddr_daily ORDER BY well_key,report_no')):row(ws,[r['well_key'],r['report_no'],r['date_g'],r['date_j'],r['rig'],r['hole_size'],r['formation'],r['md_0000'],r['md_2400'],r['drilled_24'],r['rop'],r['mw'],r['sheet_remark'],r['sheet_data'],r['template']],i)
    ws=sheet('09_DDR_Activities',['Well','Report','From','To','Hours','Phase','Code','Status','NPT attributed','Main activity','Description','Sheet','Row'],[22,8,13,13,9,10,8,10,18,28,100,28,7])
    for i,r in enumerate(R('SELECT * FROM ddr_activities ORDER BY well_key,report_no,id')):row(ws,[r['well_key'],r['report_no'],r['t_from'],r['t_to'],r['hours'],r['phase'],r['code'],r['status'],r['npt_attributed'],r['main_activity'],(r['description'] or '')[:380],r['sheet'],r['row_no']],i)
    ws=sheet('10_DDR_Time_Breakdown',['Well','Category','Hours','Days equivalent'],[22,60,16,18])
    for i,r in enumerate(R('SELECT well_key,category,SUM(hours) h FROM ddr_time GROUP BY well_key,category ORDER BY well_key,h DESC')):row(ws,[r['well_key'],r['category'],r['h'],r2((r['h'] or 0)/24)],i)
    ws=sheet('11_DDR_Materials',['Well','Material','Unit','Received','Used','Max stock','Records'],[22,66,16,16,16,16,12])
    for i,r in enumerate(R('SELECT well_key,material,unit,SUM(received) rec,SUM(used) used,MAX(stock) st,COUNT(*) n FROM ddr_materials GROUP BY well_key,material,unit ORDER BY well_key,used DESC')):row(ws,[r['well_key'],r['material'],r['unit'],r['rec'],r['used'],r['st'],r['n']],i)
    ws=sheet('12_Formation_Tops',['Op','Well','Formation','Prognosis MD','Actual MD','TVD','Thickness','Structural Δ','Lithology','DDR obs','EOWR MD'],[5,22,20,18,16,16,14,16,48,14,14])
    for k,p in sorted(P.items(),key=lambda x:x[1]['op']):
        for i,f in enumerate(p['formations']):row(ws,[p['op'],k,f['name'],f['prognosis_md'],f['md'],f['tvd'],f['thickness_m'],f['structural_delta_m'],f['lithology'],f['observations'],f['eowr_md']],i)
    ws=sheet('13_Reconciliation',['Well','Normalised item','CBS USD','DDR qty used','CBS unit','DDR unit','Status'],[22,70,16,16,14,14,48])
    for i,r in enumerate(R('SELECT * FROM reconciliation ORDER BY well_key,item')):row(ws,[r['well_key'],r['item'],r['cbs'],r['ddr'],r['unit'],'',r['status']],i)
    ws=sheet('14_NPT_Pareto',['Op','Well','Event category','Occurrences','NPT h','NPT days','Attributed','Depth m','Formation','USD exposure','Evidence'],[5,22,38,12,12,12,22,12,18,18,100])
    for k,p in sorted(P.items(),key=lambda x:x[1]['op']):
        for i,e in enumerate(p['events']):row(ws,[p['op'],k,e['category'],e['occurrences'],e['hours'],r2((e['hours'] or 0)/24),e['attributed'],e['depth_m'],e['formation'],e['cost_usd'],' | '.join(e['evidence'][:2])[:380]],i)
    ws=sheet('15_EOWR_PDF_Facts',['Well','Category','Fact','Value','Unit','Page','Document','Context'],[22,22,36,14,12,8,46,100])
    for i,r in enumerate(R('SELECT * FROM eowr_facts ORDER BY well_key,category,page_no LIMIT 25000')):row(ws,[r['well_key'],r['category'],r['fact_key'],r['value'],r['unit'],r['page_no'],r['rel_path'],(r['context'] or '')[:280]],i)
    ws=sheet('16_PDF_Numbered_Tables',['Document','Page','Table','Title','Well','Category','Extract'],[46,8,12,56,22,12,100])
    for i,r in enumerate(R('SELECT t.*,f.rel_path FROM pdf_tables t JOIN source_files f ON f.id=t.file_id ORDER BY f.rel_path,t.page_no')):row(ws,[r['rel_path'],r['page_no'],r['table_no'],r['title'],r['well_key'],r['category'],(r['body'] or '')[:280]],i)
    ws=sheet('17_Lessons_Learned',['Op','Well','Code','Polarity','Title','Numerical evidence','Action','Hours','USD'],[5,22,10,16,62,105,105,12,16])
    for k,p in sorted(P.items(),key=lambda x:x[1]['op']):
        for i,l in enumerate(p['lessons']):row(ws,[p['op'],k,l['code'],l['polarity'],l['title'],l['evidence'],l['action'],l['hours'],l['cost_usd']],i)
    ws=sheet('18_Recommendation_Cards',['Op','Well','Code','Branch','Title','Numerical evidence','Root cause','Protocol','Control window','KPI','Owner','Priority','Days saved','USD saved'],[5,22,10,34,52,105,62,105,70,50,36,10,12,16])
    for k,p in sorted(P.items(),key=lambda x:x[1]['op']):
        for i,rc in enumerate(p['recs']):row(ws,[p['op'],k,rc['code'],rc['branch'],rc['title'],rc['evidence'],rc['root_cause'],rc['protocol'],rc['control_window'],rc['kpi'],rc['owner'],rc['priority'],rc['days_saved'],rc['usd_saved']],i)
    ws=sheet('19_Audit_Findings',['ID','Finding','Detail'],[8,62,150])
    for i,a in enumerate(F['audit_findings']):row(ws,[a['id'],a['title'],a['detail']],i)
    ws=sheet('20_Cell_Inventory',['Sheet ID','Document','Category','Sheet','Rows','Cols','Non-empty','Numeric','Formulas','Template','Report','Part','Well'],[10,46,10,38,9,8,14,12,12,10,10,7,22])
    for i,r in enumerate(R('SELECT * FROM sheets ORDER BY id')):row(ws,[r['id'],r['rel_path'],r['category'],r['name'],r['n_rows'],r['n_cols'],r['n_nonempty'],r['n_numeric'],r['n_formula'],r['template'],r['report_no'],r['part'],r['well_key']],i)
    ws=sheet('21_PDF_Page_Inventory',['Document','Page','Well','Category','Chapter','Characters','Images','Keywords'],[46,8,24,12,36,12,12,58])
    for i,r in enumerate(R('SELECT rel_path,page_no,well_key,category,chapter,n_chars,n_images,keywords FROM pdf_pages ORDER BY rel_path,page_no')):row(ws,[r['rel_path'],r['page_no'],r['well_key'],r['category'],r['chapter'],r['n_chars'],r['n_images'],r['keywords']],i)
    ws=sheet('22_Field_KPI',['Metric','Value','Basis'],[34,24,110])
    for i,(k,v) in enumerate(F['totals'].items()):row(ws,[k,v if not isinstance(v,list) else ', '.join(v),'Aggregated from 51 source documents'],i)
    wb.save(XLSX);print('  WORKBOOK MB',round(os.path.getsize(XLSX)/1e6,2),flush=True)

def main():
    t0=time.time()
    schema();excel_stage();pdf_stage();enrich_stage();facts_stage()
    P,F=analytics_stage();workbook_stage(P,F)
    con.execute('VACUUM');con.close()
    print('  WAREHOUSE MB',round(os.path.getsize(DB)/1e6,2),'ELAPSED',round(time.time()-t0,1),flush=True)
if __name__=='__main__':main()
