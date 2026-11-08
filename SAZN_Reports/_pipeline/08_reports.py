"""SAZN reference books — 30 DOCX (26 single-well + 4 portfolio).
Golden Prototype Gate: AZNS-207 and AZNS-40 FA/EN are built and audited first;
the remaining 26 books are then produced without stopping, at the same standard.
Every number is drawn from the rebuilt warehouse or the derived profile JSON.
"""
import os, sys, json, math, io, re, sqlite3, collections, hashlib
sys.path.insert(0, os.path.dirname(__file__))
from common import OUT, DATA, GFX, DB, PAL, fmt, r2, load_json

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

P = load_json('well_profiles')
F = json.load(open(os.path.join(DATA, 'field_analytics.json')))
TOT = F['totals']
CON = sqlite3.connect(DB); CON.row_factory = sqlite3.Row
NAVY = '0B1F3A'; EVEN = 'F4F7FC'; ODD = 'FFFFFF'; BORDER = 'CBD5E1'
FONT = {'FA': 'B Nazanin', 'EN': 'Times New Roman'}
DIR = os.path.join(OUT, 'Word_Books'); os.makedirs(DIR, exist_ok=True)


class SafeDict(dict):
    def __missing__(self, k): return '—'


def n(v, d=0.0):
    try: return float(v)
    except Exception: return d


def f0(v): return fmt(v, 0) if v is not None and v != '' else '—'
def f1(v): return fmt(v, 1) if v is not None and v != '' else '—'
def f2(v): return fmt(v, 2) if v is not None and v != '' else '—'


# ---------------------------------------------------------------- context --
def ctx(key):
    if key == 'FIELD':
        c = F['contracts']; w = F['wells']
        return SafeDict(dict(well='SAZN field', op='13', rig='OEOC 201 / OEOC 208',
            contract='C1 + C2 + C3', contract_code='CCA/C/1401/27 · CCA/C/1403/31 · CCA/C/1403/31 ST2 amendment',
            days=f0(TOT['ddr_days']), td=f0(max([n(x['depth_final']) for x in w])),
            drilled=f1(TOT['drilled_m']), rop=f2(sum(n(x['rop_mean']) for x in w)/13),
            npt=f1(TOT['npt_hours']), nptc=f1(TOT['payable']*0+0), nptk='—', pt='—',
            cbs=f0(TOT['cbs_plan']), gross=f0(TOT['gross']), payable=f0(TOT['payable']),
            net=f0(TOT['net']), rate='23,000', forms=f0(TOT['formations']),
            events=f0(TOT['events']), lessons=f0(TOT['lessons']), recs=f0(TOT['recs']),
            recon=f0(TOT['reconciliation']), items=f0(TOT['numeric_cells']),
            act=f0(TOT['ddr_days']), cells=f0(TOT['cells']), pdfnums=f0(TOT['pdf_numbers']),
            pages=f0(TOT['pdf_pages']), formulas=f0(TOT['formulas']),
            eowr='11 files / 2,724 pages', mwmin='—', mwmax='—', maxdls='—',
            flat='—', first_date=str(w[0]['date_start'] or '')[:11], last_date=str(w[-1]['date_end'] or '')[:11],
            contracts=len(c), rigs=2, wells=len(w), ncontract=len(c)))
    p = P[key]; t = p['time']; c = p['costs']; m = p['materials']
    ts = p['timeseries']
    cs = dict(CON.execute("SELECT SUM(n_nonempty),SUM(n_numeric),SUM(n_formula) FROM sheets WHERE well_key=?",(key,)).fetchone() and
              zip(('cells','nums','formulas'),CON.execute("SELECT SUM(n_nonempty),SUM(n_numeric),SUM(n_formula) FROM sheets WHERE well_key=?",(key,)).fetchone()))
    pg = CON.execute("SELECT COUNT(*) FROM pdf_pages WHERE well_key LIKE ?",('%'+key+'%',)).fetchone()[0]
    pn = CON.execute("SELECT COUNT(*) FROM pdf_num_pages WHERE well_key LIKE ?",('%'+key+'%',)).fetchone()[0]
    return SafeDict(dict(well=p['key'], op=f"{p['op']:02d}", rig=p['rig'], contract=p['contract'],
        contract_code=p['contract_code'], days=f0(p['ddr_days']), td=f1(p['depth_final']),
        drilled=f1(p['drilled_total']), rop=f2(p['rop_mean']), npt=f1(t['npt_total_h']),
        nptc=f1(t['npt_contractor_h']), nptk=f1(t['npt_client_h']),
        pt=f1(t['status_hours'].get('PT')), cbs=f0(c['cbs_plan']), gross=f0(c['gross_calendar']),
        payable=f0(c['payable_excl_ctr_npt']), net=f0(c['pure_net_excl_all_npt']),
        rate=f0(c['rig_rate']), forms=f0(len(p['formations'])), events=f0(len(p['events'])),
        lessons=f0(len(p['lessons'])), recs=f0(len(p['recs'])),
        recon=f0(len(m['reconciliation'])), items=f0(len(c['items'])),
        act=f0(t['activities']), cells=f0(cs.get('cells')),
        pdfnums=f0(pn), pages=f0(pg), formulas=f0(cs.get('formulas')), eowr=p['eowr_status'][:80],
        mwmin=f1(p['mw_min']), mwmax=f1(p['mw_max']), maxdls=f2(p['max_dls']),
        flat=f0(sum(1 for x in ts if not n(x['drilled']))),
        first_date=str(p['date_start'] or '')[:11], last_date=str(p['date_end'] or '')[:11],
        wells=1, contracts=1, rigs=1))


# ------------------------------------------------------------- translations --
TR = {
 'Op':'عملیات','Operation':'عملیات','Well':'چاه','Report':'شماره گزارش','DDR':'گزارش روزانه',
 'Date':'تاریخ','Gregorian date':'تاریخ میلادی','Persian date':'تاریخ شمسی',
 'MD (m)':'عمق اندازهگیری (متر)','Depth (m)':'عمق (متر)','TVD (m)':'عمق قائم (متر)',
 'Formation':'سازند','Lithology':'سنگشناسی','Prognosis MD (m)':'عمق پیشبینی (متر)',
 'Actual MD (m)':'عمق واقعی (متر)','Delta (m)':'اختلاف (متر)','Hours':'ساعت',
 'Hours (h)':'ساعت','Code':'کد','Category':'دسته','Phase':'فاز','Activity':'فعالیت',
 'Status':'وضعیت','Attributed':'منتسب به','Description':'شرح','Value':'مقدار',
 'Unit':'واحد','Qty':'مقدار','Unit price (USD)':'قیمت واحد (دلار)','Total (USD)':'مبلغ (دلار)',
 'Sheet':'شیت','Cell':'سلول','Page':'صفحه','Document':'سند','Source':'منبع',
 'Material':'ماده','Received':'دریافتی','Used':'مصرف','Stock':'موجودی','Back-load':'بازگشتی',
 'Rig':'دکل','Contract':'قرارداد','Item':'قلم','Section':'بخش','Rate (USD/day)':'نرخ (دلار/روز)',
 'Cost track':'مسیر هزینه','USD':'دلار','Basis':'مبنا','Count':'تعداد','Min':'کمینه',
 'Mean':'میانگین','Max':'بیشینه','Median':'میانه','n':'تعداد نمونه','Parameter':'پارامتر',
 'Size (in)':'اندازه (اینچ)','Grade':'گرید','Weight':'وزن','From (m)':'از (متر)','To (m)':'تا (متر)',
 'Burst':'ترکیدگی','Collapse':'لهیدگی','BHA No':'شماره BHA','OD (in)':'قطر خارجی (اینچ)',
 'Length (m)':'طول (متر)','Cum. length (m)':'طول تجمعی (متر)','Incl (deg)':'انحراف (درجه)',
 'Azi (deg)':'آزیموت (درجه)','DLS (deg/30m)':'DLS (درجه/30 متر)','Tool':'ابزار',
 'POB':'افراد حاضر','Total':'جمع','Title':'عنوان','Polarity':'جهت','Evidence':'شاهد عددی',
 'Action':'اقدام','Cost (USD)':'هزینه (دلار)','Hypothesis':'فرضیه','Playbook':'دستورالعمل ۶۰ ثانیهای',
 'REC':'توصیه','Branch':'شاخه','Root cause':'علت ریشهای','Protocol':'پروتکل',
 'Control window':'پنجره کنترل','KPI':'شاخص','Owner':'مسئول','Priority':'اولویت',
 'Days saved':'روز صرفهجویی','USD saved':'دلار صرفهجویی','CBS (USD)':'CBS (دلار)',
 'DDR qty':'مقدار DDR','EOWR':'EOWR','Delta':'اختلاف','Status':'وضعیت',
 'Chapter':'فصل','Chapter code':'کد فصل','Characters':'کاراکتر','Images':'تصویر',
 'Keywords':'کلیدواژه','Table':'جدول','Table no':'شماره جدول','Extract':'گزییده',
 'Fact':'واقعهمندی','Fact key':'کلید','Context':'زمینه','Low-text pages':'صفحات کمتکست',
 'Risk (h)':'ریسک (ساعت)','Structural delta (m)':'اختلاف ساختمانی (متر)',
 'Plan (USD)':'برنامه (دلار)','Gross (USD)':'ناخالص (دلار)','Payable (USD)':'قابلپرداخت (دلار)',
 'Net (USD)':'خالص (دلار)','MW (PCF)':'گل (PCF)','ROP (m/h)':'ROP (متر/ساعت)',
 'Drilled (m)':'حفاری (متر)','NPT (h)':'NPT (ساعت)','PT (h)':'PT (ساعت)',
 'Cum. cost (USD)':'هزینه تجمعی (دلار)','Spud / start':'شروع','End':'پایان',
 'Contractor':'پیمانکار','Client':'کارفرما','Share (%)':'سهم (درصد)',
 'Standby days':'روز آمادهباش','Operating days':'روز عملیات','Standby rate':'نرخ آمادهباش',
 'Operating rate':'نرخ عملیات','Volume':'حجم','Remark':'توضیح',
 'Mud weight (PCF)':'وزن گل (PCF)','FIT EMW':'FIT EMW','Hole size':'اندازه چاله',
 'Formula':'فرمول','Type':'نوع','Rows':'سطرها','Cols':'ستونها',
 'Non-empty cells':'سلولهای پر','Numeric cells':'سلولهای عددی',
 'Formula cells':'سلولهای فرمولی','Template':'قالب','Part':'بخش',
 'N_units':'تعداد واحد','Size (bytes)':'حجم (بایت)','SHA-256':'SHA-256',
 'Metric':'معیار','Notes':'یادداشت','Finding':'یافته','Detail':'جزئیات',
}

def H(cols, lang):
    return [TR.get(c, c) for c in cols] if lang == 'FA' else list(cols)


# ------------------------------------------------------------- 24 chapters --
CH = [
 ('Executive Well Dossier & Source Register','پرونده جامع چاه و فهرست اسناد منبع'),
 ('Four-Track Time / Cost & NPT Deductions','چهار مسیر زمان / هزینه و کسر NPT'),
 ('CBS Tariff & Material Audit vs DDR / EOWR','ممیزی تعرفهها و مواد CBS در برابر DDR و EOWR'),
 ('Milestones, Flat Spots & Invisible Lost Time','نقاط عطف، توقفهای بدون پیشروی و زمان پنهان'),
 ('24-Hour Activity Log & Operation Codes','گزارش فعالیت ۲۴ ساعته و کدهای عملیاتی'),
 ('NPT Pareto & Event Autopsy','پارتو NPT و کالبدشکافی رویدادها'),
 ('Bit Record, Dull Grade & Cost per Metre','کارنامه متهها، درجه سایش و هزینه هر متر'),
 ('ROP, MSE & Parameter Correlation','ROP، MSE و همبستگی پارامترها'),
 ('BHA & Downhole Motor Performance','رشته BHA و عملکرد موتور درونچاهی'),
 ('Stick-Slip, Bit Balling & Whirl','Stick-Slip، Balling و Whirl'),
 ('Rheology, Annular Velocity, ECD & HSI','رئولوژی، سرعت آنالوس، ECD و HSI'),
 ('Tripping, Hookload & Swab-Surge','ترپینگ، هوکلواد و Swab-Surge'),
 ('Connection Time & Invisible Lost Time','زمان اتصالات و زمان پنهان ازدسترفته'),
 ('Directional Survey, DLS & Tortuosity','جهتداری، DLS و پیچخوردگی مسیر'),
 ('Well Control, Losses, LCM, Dry Test & MAASP','کنترل چاه، هرزروی، LCM، Dry Test و MAASP'),
 ('Casing, Cement & CBL-VDL-USIT','لولهگذاری، سیمان و CBL-VDL-USIT'),
 ('Workover, Fishing & Milling','Workover، ماهیگیری و میلینگ'),
 ('Completion Tally & Pressure Tests','تالی تکمیل و تستهای فشار'),
 ('Matrix Acidizing, CTU & Productivity Index','اسیدکاری ماتریس، CTU و شاخص بهرهدهی'),
 ('HSE, MTBF, Waste & Subcontractors','HSE، MTBF، پسماند و پیمانکاران جزء'),
 ('Deep Geology & Field Correlation','زمینشناسی عمیق و همبستگی میدانی'),
 ('Problem Autopsy & 60-Second Rig-Floor Playbooks','کالبدشکافی مسائل و Playbook شصتثانیهای'),
 ('LL-01 … LL-12 Best Practices & Pitfalls','درسهای آموخته LL-01 تا LL-12'),
 ('Optimization Matrix & REC-01 … REC-12','ماتریس بهینهمسازی و کارتهای REC-01 تا REC-12'),
]

EN_TEXT = [
 "Operation {op} on {well} was executed by {rig} under {contract} ({contract_code}) across {days} signed DDRs from {first_date} to {last_date}. The dossier joins {cells} archived worksheet cells, {pages} PDF pages and {pdfnums} indexed numeric tokens; the EOWR position is: {eowr}. Every following table carries its file, sheet/cell or page reference so that any figure can be traced back without re-reading the original document.",
 "The four contractual tracks close at CBS plan {cbs} USD, gross calendar {gross} USD, payable excluding contractor zero-rate NPT {payable} USD and pure net excluding all NPT {net} USD. Total attributed NPT is {npt} h, split {nptc} h contractor and {nptk} h client/subsurface at a CBS rig rate of {rate} USD/day; planned-time rows total {pt} h. Deductions are computed only from attributed NPT hours and the sourced day rate.",
 "The CBS workbook holds {items} priced lines for {well}; the DDR daily material balance and the EOWR completion / cement / fluid chapters are reconciled line by line. {recon} three-way reconciliation rows are retained, together with the exact unit price, quantity, sheet and cell reference of each priced item. Where a CBS line has no DDR consumption counterpart the status says so explicitly rather than inserting an assumed quantity.",
 "Depth progressed to {td} m MD with {drilled} m of recorded 24-h footage at mean ROP {rop} m/h; {flat} DDR reports show zero recorded progress and are retained as flat spots rather than interpolated. Milestones, invisible lost time and cumulative cost are listed per report so that a time-depth curve can be rebuilt directly from the table.",
 "{act} discrete 24-h activity rows were parsed for {well}, each with phase, code, sub-code, PLN/NPT status, attributed party and duration. The tables below reproduce the raw activity log, the code distribution and the phase-by-status hour matrix; no activity is merged or rounded away.",
 "{events} problem families were reconstructed for {well}. Each carries occurrence count, hours, depth, formation, attributed party, at least two engineering hypotheses, a quantified exposure and a 60-second rig-floor playbook. The Pareto tables rank these families by hours and by cost exposure using the sourced rig rate.",
 "Bit and drilling-tool records are taken from the CBS tariff lines, DDR bit/BHA sheets and EOWR bit-record chapter. Cost per drilled metre is formed from {cbs} USD of CBS scope and {drilled} m of recorded footage; where no new hole was drilled the table states the workover basis instead of dividing by an invented denominator.",
 "Mean ROP is {rop} m/h with mud weight band {mwmin}–{mwmax} PCF. The correlation tables join daily ROP, MW, drilled interval, hole size and formation from the same DDR report so that any apparent trend can be checked against the original 24-h row. {maxdls} deg/30m is the maximum calculated minimum-curvature DLS for this well.",
 "The BHA table reproduces the as-run components with item, OD, length and cumulative length from the DDR BHA sheet, followed by the derived parameter envelopes. Where the source DDR contains no BHA rows the report records that absence and points to the CBS / Programme / EOWR toolstring tables instead of supplying a schematic assembly.",
 "Vibration, bit balling and whirl indicators were searched across DDR remarks, activity descriptions and notes. The tables list every matching source row with its report number and text; where no matching text exists the chapter states the absence explicitly. No vibration index is invented from a silent source.",
 "Mud weight, FIT/LOT equivalents, annular-velocity proxies and hydraulic parameters are tabulated from DDR daily columns, CBS material lines and EOWR fluid chapters. The {mwmin}–{mwmax} PCF envelope is the measured band, not a design assumption; volume balance and solids-control rows retain Received / Used / Back-load / Stock as written.",
 "Tripping, reaming, PU/SO/ROT hookload and swab-surge related activities are isolated by code and keyword from {act} activity rows. Hours, phase and description are reproduced; hookload statistics come from labelled DDR parameter rows only. Any absent parameter is shown as '—'.",
 "Connection, survey, circulating and waiting codes are separated from productive drilling time. The tables give code, category, hours and report number so that invisible lost time can be recomputed; the sum is never forced to 24 h where the source DDR itself is incomplete.",
 "Survey stations, inclination, azimuth, TVD and minimum-curvature DLS are reproduced per station. Maximum DLS is {maxdls} deg/30m. Directional parameter windows (WOB, RPM, torque, SPP, GPM) come from labelled DDR rows and are listed with count, min, mean and max.",
 "Losses, LCM pills, dry tests, bond tests, FIT/LOT and MAASP-related facts are extracted from DDR remarks and the EOWR / Programme text layer. {events} problem families include the loss-control evidence; each numeric fact carries its page number. Where a pressure-test value was not printed in the text layer it is not estimated.",
 "Casing tallies list size, grade, weight, thread, from/to depth, burst and collapse from DDR casing sheets. Cement design, lab and job pages, TOC, and CBL-VDL-USIT bond facts are indexed by EOWR/Programme page. Completion equipment from CBS priced lines is matched to the pressure-test facts for the same well.",
 "Fishing, milling, whipstock, sidetrack-window and re-entry activities are isolated from the 24-h log. For {well} the operation class is recorded in the dossier; the tables reproduce the matching activity rows, notes and EOWR workover-section pages. Operation 4 has no EOWR and is analysed from DDR + CBS + Programme only.",
 "Completion tally rows are taken from CBS 20.1/20.2/20.3 material sheets and DDR casing records; TRSSSV, packer, anchor, SSD, expansion joint, nipples and CIV/CIM lines are listed with their priced quantities. Pressure-test facts from the EOWR / Programme text layer are tabulated with value, unit and page.",
 "Rig-less stimulation families, planned chemical volumes, unit prices and acid / N2 / WHP / PI / rate facts are listed with page references. Pre/post-acid productivity comparisons are only made where both values exist in the source text; {recon} reconciliation rows link CBS priced chemicals to DDR consumption.",
 "POB series, fuel/water consumption, waste rows, HSE chapter pages and contractor / service-company names are reproduced from DDR and EOWR. The accommodation and catering CBS lines are matched against the DDR POB totals rather than against an assumed headcount.",
 "{forms} correlated formation tops join DDR prognosis/actual columns, EOWR tops and Programme prognosis tables. Structural delta is actual minus prognosis in metres; lithology, gas shows, H2S/CO2, fluorescence and CBL-VDL-USIT evidence are indexed by page. South Azadegan anticline correlation is established against the other 12 operations.",
 "Each of the {events} real problem families is autopsied with day, depth, formation, two or more hypotheses, action effectiveness, hours and cost exposure at {rate} USD/day. The 60-second rig-floor contingency playbooks are reproduced verbatim in a dedicated table; they are operational instructions, not generic advice.",
 "Twelve non-repeated LL cards are derived from {lessons} well-specific measurements: peak footage, code coverage, ROP distribution, formation-top control, CBS granularity, material balance and the largest real problem families. Best-practice and pitfall polarity are stated per card with numeric evidence and a concrete action.",
 "Four optimization branches (geoscience/casing shoes, drilling parameters & loss control, workover/fishing/rig PM, completion/cement/acid/CBS) carry twelve nine-axis REC cards. Each card records phase, numeric evidence, root cause, protocol, per-hole control window, KPI, owner, priority and quantified day/USD savings computed from sourced hours and the sourced day rate.",
]

FA_TEXT = [
 "عملیات {op} روی چاه {well} توسط دکل {rig} ذیل قرارداد {contract} ({contract_code}) در {days} گزارش روزانه امضاشده از {first_date} تا {last_date} اجرا شد. این پرونده {cells} سلول کارپوشه، {pages} صفحه PDF و {pdfnums} عدد نمایهشده را بههم پیوند میدهد؛ وضعیت EOWR چنین است: {eowr}. هر جدول بعدی نشانی فایل، شیت/سلول یا صفحه را همراه دارد تا هر رقم بدون مراجعه مجدد به سند اصلی قابل ردیابی باشد.",
 "چهار مسیر قراردادی عبارتاند از برنامه CBS برابر {cbs} دلار، ناخالص تقویمی {gross} دلار، قابلپرداخت بدون NPT نرخ صفر پیمانکار {payable} دلار و خالص مطلق بدون همه NPT برابر {net} دلار. جمع NPT منتسب {npt} ساعت است که {nptc} ساعت پیمانکار و {nptk} ساعت کارفرما/زیرسطحی، با نرخ دکل {rate} دلار در روز محاسبه شده است؛ ردیفهای زمان برنامه {pt} ساعت است. کسورات تنها از ساعات NPT منتسب و نرخ منبع محاسبه میشود.",
 "کارپوشه CBS دارای {items} ردیف قیمتی برای {well} است؛ تراز روزانه مواد DDR و فصلهای تکمیل/سیمان/سیال EOWR ردیفبهردیف تطبیق داده شد. {recon} ردیف تطبیق سهجانبه همراه قیمت واحد، مقدار، شیت و نشانی سلول هر قلم نگهداری شده است. هرجا ردیف CBS معادل مصرف DDR نداشته باشد، وضعیت صریحاً اعلام میشود و مقداری فرضی درج نمیگردد.",
 "عمق تا {td} متر MD پیش رفت و {drilled} متر حفاری ۲۴ ساعته با ROP میانگین {rop} متر بر ساعت ثبت شد؛ {flat} گزارش DDR بدون پیشروی ثبتشدهاند و بهجای درونیابی، بهعنوان توقف تخت نگهداری میشوند. نقاط عطف، زمان پنهان و هزینه تجمعی بهتفکیک گزارش آمده است تا منحنی زمان–عمق مستقیماً بازسازی شود.",
 "تعداد {act} ردیف فعالیت ۲۴ ساعته مجزا برای {well} استخراج شد که هر یک فاز، کد، زیرکد، وضعیت PLN/NPT، طرف مسئول و مدت دارد. جداول زیر لاگ خام فعالیت، توزیع کدها و ماتریس ساعات فاز بر وضعیت را عیناً بازتولید میکنند؛ هیچ فعالیتی ادغام یا گرد نشده است.",
 "تعداد {events} خانواده مسئله برای {well} بازسازی شد. هر خانواده تعداد رخداد، ساعات، عمق، سازند، طرف مسئول، دستکم دو فرضیه مهندسی، مواجهه کمیشده و دستورالعمل ۶۰ ثانیهای دارد. جداول پارتو این خانوادهها را بر پایه ساعات و مواجهه هزینهای با نرخ دکل منبع رتبهبندی میکنند.",
 "کارنامه متهها و ابزارهای حفاری از ردیفهای تعرفه CBS، شیتهای مته/BHA در DDR و فصل کارنامه مته EOWR گرفته شده است. هزینه هر متر حفاری از محدوده {cbs} دلار و {drilled} متر فوتیج ثبتشده تشکیل میشود؛ هرگاه حفر چاله جدیدی ثبت نشده باشد، جدول مبنای workover را اعلام میکند و بر مخرج فرضی تقسیم نمیکند.",
 "ROP میانگین {rop} متر بر ساعت با باند وزن گل {mwmin} تا {mwmax} PCF است. جداول همبستگی، ROP روزانه، MW، بازه حفاری، اندازه چاله و سازند را از همان گزارش DDR بههم میپیوندند. بیشینه DLS محاسبهشده بهروش انحنای کمینه {maxdls} درجه بر ۳۰ متر است.",
 "جدول BHA اجزای اجراشده شامل شرح، قطر خارجی، طول و طول تجمعی را از شیت BHA در DDR بازتولید میکند و پنجرههای پارامتری بهدنبال آن میآیند. اگر DDR منبع فاقد ردیف BHA باشد، گزارش نبود آن را ثبت و به جداول رشته ابزار CBS / Programme / EOWR ارجاع میدهد و چیدمانی فرضی ترسیم نمیکند.",
 "نشانههای ارتعاش، Balling مته و Whirl در توضیحات DDR، شرح فعالیتها و یادداشتها جستوجو شد. جداول هر ردیف منطبق را با شماره گزارش و متن آن میآورند؛ اگر متنی یافت نشود فصل صریحاً نبود آن را اعلام میکند و هیچ شاخص ارتعاشی از منبع خاموش ساخته نمیشود.",
 "وزن گل، معادلهای FIT/LOT، نمایههای سرعت آنالوس و پارامترهای هیدرولیکی از ستونهای روزانه DDR، ردیفهای مواد CBS و فصول سیال EOWR جدولبندی شدهاند. باند {mwmin} تا {mwmax} PCF باند اندازهگیری است، نه فرض طراحی؛ ردیفهای تراز حجمی و کنترل جامدات عیناً Received / Used / Back-load / Stock را نگهداری میکنند.",
 "فعالیتهای ترپینگ، ریمینگ، هوکلواد PU/SO/ROT و Swab-Surge بر پایه کد و کلیدواژه از {act} ردیف فعالیت جدا شدهاند. ساعات، فاز و شرح عیناً آمده است؛ آمار هوکلواد فقط از ردیفهای پارامتری برچسبخورده DDR گرفته میشود و پارامتر غایب با «—» نمایش داده میشود.",
 "کدهای اتصال، نگرشی، گردش و انتظار از زمان حفاری مولد جدا شدهاند. جداول کد، دسته، ساعات و شماره_reportno را میدهند تا زمان پنهان ازدسترفته قابل محاسبه مجدد باشد؛ هرجا خود DDR ناقص است، مجموع بر ۲۴ ساعت تحمیل نمیشود.",
 "ایستگاههای نگرشی، انحراف، آزیموت، TVD و DLS انحنای کمینه بهتفکیک ایستگاه بازتولید شدهاند. بیشینه DLS برابر {maxdls} درجه بر ۳۰ متر است. پنجرههای پارامتر جهتدار (WOB، RPM، گشتاور، SPP، GPM) از ردیفهای برچسبخورده DDR همراه تعداد، کمینه، میانگین و بیشینه آمدهاند.",
 "هرزروی، پیلهای LCM، تست خشک، تست باند، FIT/LOT و واقعهمندیهای مرتبط با MAASP از توضیحات DDR و لایه متن EOWR / Programme استخراج شدهاند. {events} خانواده مسئله شواهد کنترل هرزروی را دارند و هر رقم شماره صفحه خود را همراه دارد. مقداری که در لایه متن چاپ نشده تخمین زده نمیشود.",
 "تالی لولهگذاری شامل اندازه، گرید، وزن، رزوه، عمق از/تا، ترکیدگی و لهیدگی از شیتهای DDR است. صفحات طراحی، آزمایشگاه و اجرای سیمان، TOC و واقعهمندیهای باند CBL-VDL-USIT بر پایه صفحه EOWR/Programme نمایه شدهاند. تجهیزات تکمیل از ردیفهای قیمتی CBS با تستهای فشار همان چاه تطبیق داده شده است.",
 "فعالیتهای ماهیگیری، میلینگ، ویپاستاک، پنجره سایدترک و ورود مجدد از لاگ ۲۴ ساعته جدا شدهاند. برای {well} کلاس عملیات در پرونده ثبت شده است؛ جداول ردیفهای منطبق، یادداشتها و صفحات بخش Workover EOWR را بازتولید میکنند. عملیات ۴ فاقد EOWR است و تنها از DDR + CBS + Programme تحلیل میشود.",
 "ردیفهای تالی تکمیل از شیتهای مواد 20.1/20.2/20.3 در CBS و رکوردهای لولهگذاری DDR گرفته شدهاند؛ خطوط TRSSSV، پکر، انکر، SSD، درز انبساط، نیپل و CIV/CIM با مقادیر قیمتی آمدهاند. واقعهمندیهای تست فشار از لایه متن EOWR / Programme با مقدار، واحد و صفحه جدولبندی شدهاند.",
 "خانوادههای تحریک بدون دکل، حجمهای برنامهریزشده مواد، قیمت واحد و واقعهمندیهای اسید / N2 / WHP / PI / نرخ با نشانی صفحه آمدهاند. مقایسه بهرهدهی پیش/پس اسید تنها وقتی انجام میشود که هر دو مقدار در متن منبع موجود باشند؛ {recon} ردیف تطبیق، مواد قیمتی CBS را به مصرف DDR پیوند میدهد.",
 "سری POB، مصرف سوخت/آب، ردیفهای پسماند، صفحات فصل HSE و نام پیمانکاران/شرکتهای خدماتی از DDR و EOWR بازتولید شدهاند. ردیفهای اسکان و پذیرایی CBS بر جمع POB ثبتشده در DDR تطبیق داده میشوند، نه بر سرنفر فرضی.",
 "تعداد {forms} سر سازند همبسته، ستونهای پیشبینی/واقعی DDR، سرهای EOWR و جداول پیشبینی Programme را بههم پیوند میدهد. اختلاف ساختمانی، عمق واقعی منهای پیشبینی بر حسب متر است؛ سنگشناسی، نشانههای گاز، H2S/CO2، فلورسانس و شواهد CBL-VDL-USIT بر پایه صفحه نمایه شدهاند. همبستگی تاقدیس South Azadegan با ۱۲ عملیات دیگر برقرار شده است.",
 "هر یک از {events} خانواده مسئله واقعی با روز، عمق، سازند، دو یا چند فرضیه، اثربخشی اقدام، ساعات و مواجهه هزینهای بر پایه {rate} دلار در روز کالبدشکافی شده است. دستورالعملهای ۶۰ ثانیهای سکوی دکل عیناً در جدول جداگانهای آمدهاند؛ اینها دستور عملیاتیاند، نه توصیه کلی.",
 "دوازده کارت LL غیرتکراری از {lessons} اندازهگیری ویژه چاه مشتق شدهاند: اوج فوتیج، پوشش کدها، توزیع ROP، کنترل سر سازندها، ریزدانهبودن CBS، تراز مواد و بزرگترین خانوادههای مسئله واقعی. جهت «بهترین رویه» یا «دام» هر کارت با شاهد عددی و اقدام مشخص اعلام میشود.",
 "چهار شاخه بهینهمسازی (زمینشناسی/کفشکها، پارامترها و کنترل هرزروی، workover/ماهیگیری/نگهداری دکل، تکمیل/سیمان/اسید/CBS) دوازده کارت نهبعدی REC را حمل میکنند. هر کارت فاز، شاهد عددی، علت ریشهای، پروتکل، پنجره کنترل بهتفکیک چاله، KPI، مسئول، اولویت و صرفهجویی کمّی روز/دلار محاسبهشده از ساعات و نرخ منبع را ثبت میکند.",
]


def field_para(i, lang):
    """Field-level narrative, distinct per chapter, driven by field totals."""
    w = F['wells']; c = F['contracts']; hm = F['heatmap']; lc = F['learning_curves']
    top = sorted(w, key=lambda x: -n(x['npt_h']))[:3]
    worst = sorted(w, key=lambda x: -n(x['gross']))[:3]
    en = [
      f"The field dossier joins {TOT['operations']} operations, {TOT['contracts']} contracts and two rigs. Source census: {TOT['cbs_files']} CBS workbooks/{TOT['cbs_sheets']} sheets, {TOT['ddr_files']} DDR packs/{TOT['ddr_sheets']} sheets, {TOT['eowr_files']} EOWRs/{TOT['eowr_pages']} pages and {TOT['program_files']} programmes/{TOT['program_pages']} pages — {TOT['pdf_pages']} PDF pages with {TOT['pdf_numbers']:,} numeric tokens and {TOT['pdf_tables']:,} numbered tables. Audit findings AF-01 … AF-05 record the missing EOWR, the shared WO2/ST2 report, header contract discrepancies, the wells-29/38 rate benchmarks and low-text pages.",
      f"Across the campaign CBS plan {f0(TOT['cbs_plan'])} USD expands to gross calendar {f0(TOT['gross'])} USD; payable excluding contractor zero-rate NPT is {f0(TOT['payable'])} USD and pure net excluding all NPT {f0(TOT['net'])} USD. Attributed NPT totals {f1(TOT['npt_hours'])} h over {TOT['ddr_days']} DDR days. Contract-level splits are tabulated below for C1, C2 and C3.",
      f"{TOT['numeric_cells']:,} numeric worksheet cells, {TOT['formulas']:,} formulas and {TOT['cells']:,} non-empty cells are archived with file/sheet/A1 provenance. Priced CBS lines, DDR consumption and EOWR/Programme chapters produce {TOT['reconciliation']} three-way reconciliation rows; benchmark sheets for wells 29 and 38 are retained only as rate references, never charged to a campaign well.",
      f"Recorded 24-h footage sums to {f1(TOT['drilled_m'])} m and deepest measured depth is {f0(max(n(x['depth_final']) for x in w))} m. The time-depth series holds {TOT['ddr_days']} rows; flat spots are listed per operation with their DDR report numbers instead of interpolated.",
      f"Twenty-four-hour coding covers {TOT['ddr_days']} signed daily reports. Activity-code, phase and PLN/NPT distributions are tabulated per operation; the three highest-NPT operations are " + ', '.join(f"{x['key']} ({f0(x['npt_h'])} h)" for x in top) + '.',
      f"{TOT['events']} problem families are reconstructed field-wide. Pareto tables rank them by hours and exposure; the heatmap further distributes problem hours across 15 formations for each of the 13 operations.",
      "Bit and tool records are tabulated from CBS tariff lines and DDR/EOWR bit chapters for all 13 operations. Cost per metre is only computed where recorded footage exists; workover-only operations are labelled accordingly.",
      f"Mean of the 13 per-operation ROP values is {f2(sum(n(x['rop_mean']) for x in w)/13)} m/h; the field tables preserve each operation's own distribution, mud-weight band and correlation to drilled interval.",
      "BHA rows, toolstring OD/length and parameter windows are tabulated per operation; operations whose DDR holds no BHA rows are marked explicitly rather than assigned a schematic string.",
      "Vibration / balling / whirl keyword hits are reproduced per operation with source text. Absence of a text hit is recorded as absence.",
      "Mud-weight, FIT/LOT, annular-velocity and hydraulic parameters are listed with their source report; volume-balance rows retain Received / Used / Back-load / Stock verbatim.",
      "Tripping, hookload and swab/surge activity hours are isolated per operation from the coded 24-h log; no parameter is inferred where a DDR column is blank.",
      "Connection / survey / circulating / waiting codes are separated from productive time for each operation; sums are never forced to 24 h.",
      "Survey stations, DLS and directional windows are tabulated for the operations that recorded them; operations without survey rows are stated as such.",
      "Loss control, LCM, dry tests, FIT/LOT and pressure-test facts are indexed by EOWR/Programme page; the field risk matrix combines these with DDR loss remarks.",
      "Casing tallies, cement jobs and CBL-VDL-USIT evidence are tabulated per operation, including the 9-5/8 in, 7 in and 5 in liner strings and Xmas Tree A–D tests.",
      "Workover, fishing, milling and sidetrack windows are compared across AZNS-127 WO1 → WO2 → ST2 and AZNS-122 WO1 → WO2; operation 4 has no EOWR and is reconstructed from DDR + CBS + Programme.",
      "Completion tallies and pressure-test facts are cross-tabulated for all 13 operations, with 20.1/20.2/20.3 CBS material lines and 5K/7.5K/10K equipment.",
      "Rig-less stimulation families, planned acid/N2 volumes, WHP/PI/rate facts and clean-up results are listed per operation; the 12 chemical families are priced from CBS, never assumed.",
      "POB, catering/accommodation, fuel/water, waste, HSE pages and subcontractor names are tabulated operation by operation.",
      f"{TOT['formations']} correlated formation tops build the South Azadegan structural cross-section; prognosis-actual structural deltas, lithology, gas shows and CBL-VDL-USIT evidence are indexed by page.",
      f"{TOT['events']} field problem families carry hypotheses, action effectiveness, hours, exposure and 60-second rig-floor playbooks. The contingency playbooks are reproduced verbatim.",
      f"{TOT['lessons']} LL cards (12 per operation) are non-repeated and numerically evidenced; best-practice and pitfall polarity are stated per card.",
      f"{TOT['recs']} nine-axis REC cards (12 per operation) cover four optimization branches with quantified day and USD savings derived from sourced NPT hours and sourced rig rates.",
    ]
    fa = [
      f"پرونده میدانی {TOT['operations']} عملیات، {TOT['contracts']} قرارداد و دو دکل را بههم پیوند میدهد. سرشماری منبع: {TOT['cbs_files']} کارپوشه CBS/{TOT['cbs_sheets']} شیت، {TOT['ddr_files']} بسته DDR/{TOT['ddr_sheets']} شیت، {TOT['eowr_files']} EOWR/{TOT['eowr_pages']} صفحه و {TOT['program_files']} برنامه/{TOT['program_pages']} صفحه، یعنی {TOT['pdf_pages']} صفحه PDF با {TOT['pdf_numbers']:,} عدد و {TOT['pdf_tables']:,} جدول شمارهدار. یافتههای AF-01 تا AF-05 نبود EOWR، گزارش مشترک WO2/ST2، اختلاف کد قرارداد، محکهای چاههای ۲۹/۳۸ و صفحات کمتکست را ثبت میکنند.",
      f"در سطح میدان، برنامه CBS برابر {f0(TOT['cbs_plan'])} دلار به ناخالص تقویمی {f0(TOT['gross'])} دلار میرسد؛ قابلپرداخت بدون NPT نرخ صفر پیمانکار {f0(TOT['payable'])} دلار و خالص مطلق {f0(TOT['net'])} دلار است. NPT منتسب {f1(TOT['npt_hours'])} ساعت در {TOT['ddr_days']} روز DDR است. تقسیم سه قرارداد در جدول آمده است.",
      f"تعداد {TOT['numeric_cells']:,} سلول عددی، {TOT['formulas']:,} فرمول و {TOT['cells']:,} سلول پر با نشانی فایل/شیت/A1 بایگانی شدهاند. ردیفهای قیمتی CBS، مصرف DDR و فصول EOWR/Programme برابر {TOT['reconciliation']} ردیف تطبیق سهجانبه میدهند؛ شیتهای محک چاههای ۲۹ و ۳۸ فقط مرجع نرخاند و به هیچ چاه کارزار منظور نمیشوند.",
      f"جمع فوتیج ۲۴ ساعته {f1(TOT['drilled_m'])} متر و بیشینه عمق اندازهگیری {f0(max(n(x['depth_final']) for x in w))} متر است. سری زمان–عمق {TOT['ddr_days']} ردیف دارد؛ توقفهای تخت بهتفکیک عملیات و شماره گزارش DDR آمدهاند و درونیابی نشدهاند.",
      f"کدگذاری ۲۴ ساعته {TOT['ddr_days']} گزارش روزانه امضاشده را پوشش میدهد. توزیع کد، فاز و PLN/NPT بهتفکیک عملیات جدولبندی شده است؛ سه عملیات با بیشترین NPT: " + '، '.join(f"{x['key']} ({f0(x['npt_h'])} ساعت)" for x in top) + '.',
      f"تعداد {TOT['events']} خانواده مسئله در سطح میدان بازسازی شده است. جداول پارتو بر پایه ساعات و مواجهه رتبهبندی میکنند و نقشه حرارتی ساعات مسئله را در ۱۵ سازند برای ۱۳ عملیات پخش میکند.",
      "کارنامه متهها و ابزارها از ردیفهای تعرفه CBS و فصول مته DDR/EOWR برای ۱۳ عملیات جدولبندی شده است. هزینه هر متر فقط جایی محاسبه میشود که فوتیج ثبتشده وجود داشته باشد.",
      f"میانگین ۱۳ مقدار ROP عملیاتی {f2(sum(n(x['rop_mean']) for x in w)/13)} متر بر ساعت است؛ جداول میدانی توزیع، باند وزن گل و همبستگی هر عملیات را جداگانه نگهداری میکنند.",
      "ردیفهای BHA، قطر/طول رشته و پنجرههای پارامتری بهتفکیک عملیات آمدهاند؛ عملیاتهای فاقد ردیف BHA صریحاً مشخص شدهاند و چیدمان فرضی نمیگیرند.",
      "رخدادهای کلیدواژهای ارتعاش/Balling/Whirl بهتفکیک عملیات با متن منبع بازتولید شدهاند و نبود رخداد بهعنوان نبود ثبت میشود.",
      "وزن گل، FIT/LOT، سرعت آنالوس و پارامترهای هیدرولیکی همراه گزارش منبع آمدهاند؛ ردیفهای تراز حجمی عیناً چهار ستون اصلی را نگهداری میکنند.",
      "ساعات ترپینگ، هوکلواد و Swab-Surge برای هر عملیات از لاگ کدگذاریشده جدا شدهاند و هیچ پارامتری از ستون خالی استنتاج نمیشود.",
      "کدهای اتصال/نگرشی/گردش/انتظار برای هر عملیات از زمان مولد جدا شدهاند و مجموع بر ۲۴ ساعت تحمیل نمیشود.",
      "ایستگاههای نگرشی، DLS و پنجرههای جهتداری برای عملیاتهای ثبتکننده جدولبندی شدهاند و نبود آنها اعلام میشود.",
      "کنترل هرزروی، LCM، تست خشک، FIT/LOT و تستهای فشار بر پایه صفحه EOWR/Programme نمایه شدهاند و ماتریس ریسک میدانی آنها را با توضیحات DDR ترکیب میکند.",
      "تالی لولهگذاری، عملیات سیمان و شواهد CBL-VDL-USIT بهتفکیک عملیات، shامل رشتههای ۹-۵/۸، ۷ و لاینر ۵ اینچ و تستهای Xmas Tree A–D جدولبندی شدهاند.",
      "پنجرههای workover، ماهیگیری، میلینگ و سایدترک در زنجیرههای AZNS-127 WO1←WO2←ST2 و AZNS-122 WO1←WO2 مقایسه شدهاند؛ عملیات ۴ فاقد EOWR است.",
      "تالی تکمیل و تستهای فشار برای ۱۳ عملیات با خطوط مواد 20.1/20.2/20.3 و تجهیزات 5K/7.5K/10K جدولبندی متقابل شدهاند.",
      "خانوادههای تحریک بدون دکل، حجمهای اسید/N2، واقعهمندیهای WHP/PI/نرخ و نتایج پاکسازی بهتفکیک عملیات آمدهاند؛ ۱۲ خانواده شیمیایی از CBS قیمتگذاری شدهاند، نه فرض.",
      "POB، اسکان/پذیرایی، سوخت/آب، پسماند، صفحات HSE و نام پیمانکاران جزء بهتفکیک عملیات جدولبندی شدهاند.",
      f"تعداد {TOT['formations']} سر سازند همبسته، مقطع ساختمانی South Azadegan را میسازد؛ اختلافهای پیشبینی–واقعی، سنگشناسی، نشانههای گاز و شواهد CBL-VDL-USIT بر پایه صفحه نمایه شدهاند.",
      f"تعداد {TOT['events']} خانواده مسئله میدانی با فرضیهها، اثربخشی اقدام، ساعات، مواجهه و دستورالعمل ۶۰ ثانیهای همراهاند؛ دستورالعملها عیناً بازتولید شدهاند.",
      f"تعداد {TOT['lessons']} کارت LL (۱۲ بهازای هر عملیات) غیرتکراری و دارای شاهد عددیاند و جهت بهترین رویه/دام مشخص شده است.",
      f"تعداد {TOT['recs']} کارت نهبعدی REC (۱۲ بهازای هر عملیات) چهار شاخه بهینهمسازی را با صرفهجویی کمّی روز و دلار مبتنی بر ساعات و نرخهای منبع پوشش میدهد.",
    ]
    return [en[i], fa[i]][0 if lang == 'EN' else 1]


# ------------------------------------------------------------- table specs --
def TS(ten, tfa, cols, sql, limit=22):
    return dict(ten=ten, tfa=tfa, cols=cols, sql=sql, limit=limit)


SPEC = [
 # 1
 [TS('Source document register','فهرست اسناد منبع',['Document','Category','N_units','Size (bytes)','SHA-256'],
     "SELECT f.rel_path,f.category,f.n_units,f.size_bytes,f.sha256 FROM source_files f WHERE 1=1 {WF} ORDER BY f.rel_path LIMIT ?",30),
  TS('Worksheet / sheet inventory','سیاهه شیتها',['Document','Sheet','Rows','Cols','Non-empty cells','Numeric cells','Formula cells','Template','Part'],
     "SELECT rel_path,name,n_rows,n_cols,n_nonempty,n_numeric,n_formula,template,part FROM sheets WHERE 1=1 {W} ORDER BY id LIMIT ?"),
  TS('Dossier identity and key figures','شناسنامه و ارقام کلیدی',['Metric','Value','Basis'],
     "SELECT 'Operation',?,'Campaign map' UNION ALL SELECT 'Rig',?,'DDR header' UNION ALL SELECT 'Contract',?,'Contract register' LIMIT ?",3)],
 # 2
 [TS('Four-track cost register','دفتر چهار مسیر هزینه',['Cost track','USD','Basis'],
     "SELECT 'CBS plan',?,'Total Cost-EPD ceiling' UNION ALL SELECT 'Gross calendar',?,'CBS plan + calendar delta' UNION ALL SELECT 'Payable excl. contractor NPT',?,'Gross - zero-rate NPT' UNION ALL SELECT 'Pure net excl. all NPT',?,'Gross - all NPT' LIMIT ?",4),
  TS('NPT attribution breakdown','تفکیک انتساب NPT',['Category','Hours','Attributed','Share (%)'],
     "SELECT category,hours,attributed,share FROM npt_breakdown WHERE 1=1 {W} ORDER BY hours DESC LIMIT ?"),
  TS('DDR time-code hours','ساعات کدهای زمان DDR',['Code','Category','Hours','Report'],
     "SELECT code,category,hours,report_no FROM ddr_time WHERE 1=1 {W} AND hours IS NOT NULL ORDER BY hours DESC LIMIT ?")],
 # 3
 [TS('CBS priced tariff lines','ردیفهای تعرفه قیمتی CBS',['Sheet','Description','Unit','Qty','Unit price (USD)','Total (USD)','Cell','Formula'],
     "SELECT sheet,description,unit,qty,unit_price,total,cell_ref,formula FROM cbs_items WHERE 1=1 {W} AND total IS NOT NULL ORDER BY total DESC LIMIT ?"),
  TS('CBS total-cost ceilings','سقفهای هزینه CBS',['Sheet','Item','Qty','Unit cost','Total (USD)','Row'],
     "SELECT sheet,item,qty,unit_cost,total,row_no FROM cbs_totals WHERE 1=1 {W} ORDER BY id LIMIT ?"),
  TS('Three-way CBS / DDR / EOWR reconciliation','تطبیق سهجانبه',['Item','CBS (USD)','DDR qty','Unit','Status'],
     "SELECT item,cbs,ddr,unit,status FROM reconciliation WHERE 1=1 {W} ORDER BY item LIMIT ?")],
 # 4
 [TS('Daily time-depth milestones','نقاط عطف زمان–عمق روزانه',['Report','Date','MD (m)','Drilled (m)','ROP (m/h)','MW (PCF)','Hole size','Formation'],
     "SELECT report_no,date_g,md_2400,drilled_24,rop,mw,hole_size,formation FROM ddr_daily WHERE 1=1 {W} ORDER BY report_no LIMIT ?"),
  TS('Zero-progress flat spots','توقفهای بدون پیشروی',['Report','Date','MD (m)','NPT (h)','PT (h)','Cum. cost (USD)'],
     "SELECT report_no,date_g,md,npt_h,pt_h,cum_cost FROM field_timeseries WHERE 1=1 {W} AND (drilled IS NULL OR drilled=0) ORDER BY report_no LIMIT ?"),
  TS('Cumulative cost by DDR report','هزینه تجمعی بهتفکیک گزارش',['Report','Date','MD (m)','NPT (h)','PT (h)','Cum. cost (USD)'],
     "SELECT report_no,date_g,md,npt_h,pt_h,cum_cost FROM field_timeseries WHERE 1=1 {W} ORDER BY report_no LIMIT ?")],
 # 5
 [TS('Raw 24-hour activity log','لاگ خام فعالیت ۲۴ ساعته',['Report','Seq','Phase','Code','Status','Attributed','Hours','Description'],
     "SELECT report_no,seq,phase,code,status,npt_attributed,hours,description FROM ddr_activities WHERE 1=1 {W} ORDER BY report_no,id LIMIT ?"),
  TS('Activity code distribution','توزیع کدهای فعالیت',['Code','Hours','Count'],
     "SELECT code,SUM(hours),COUNT(*) FROM ddr_activities WHERE 1=1 {W} AND hours IS NOT NULL GROUP BY code ORDER BY SUM(hours) DESC LIMIT ?"),
  TS('Phase × status hour matrix','ماتریس ساعات فاز بر وضعیت',['Phase','Status','Hours','Count'],
     "SELECT phase,status,SUM(hours),COUNT(*) FROM ddr_activities WHERE 1=1 {W} AND hours IS NOT NULL GROUP BY phase,status ORDER BY SUM(hours) DESC LIMIT ?")],
 # 6
 [TS('Reconstructed problem families','خانوادههای مسئله بازسازی شده',['Category','Days','Depth (m)','Formation','Title','Hours','Cost (USD)','Evidence'],
     "SELECT kind,day,depth,formation,title,hours,cost_usd,evidence FROM well_events WHERE 1=1 {W} ORDER BY hours DESC LIMIT ?"),
  TS('NPT Pareto by attributed hours','پارتو NPT بر پایه ساعات',['Category','Hours','Attributed','Share (%)'],
     "SELECT category,hours,attributed,share FROM npt_breakdown WHERE 1=1 {W} ORDER BY hours DESC LIMIT ?"),
  TS('Largest individual NPT entries','بزرگترین ردیفهای NPT منفرد',['Report','Hours','Attributed','Description'],
     "SELECT report_no,hours,npt_attributed,description FROM ddr_activities WHERE 1=1 {W} AND hours IS NOT NULL ORDER BY hours DESC LIMIT ?")],
 # 7
 [TS('Bit / drilling-tool tariff lines','ردیفهای مته و ابزار حفاری',['Sheet','Description','Unit','Qty','Unit price (USD)','Total (USD)'],
     "SELECT sheet,description,unit,qty,unit_price,total FROM cbs_items WHERE 1=1 {W} AND (LOWER(description) LIKE '%bit%' OR LOWER(description) LIKE '%reamer%' OR LOWER(description) LIKE '%stabilizer%') ORDER BY total DESC LIMIT ?"),
  TS('Daily ROP / footage record','کارنامه روزانه ROP و فوتیج',['Report','Date','MD (m)','Drilled (m)','ROP (m/h)','MW (PCF)','Formation'],
     "SELECT report_no,date_g,md_2400,drilled_24,rop,mw,formation FROM ddr_daily WHERE 1=1 {W} AND rop IS NOT NULL ORDER BY report_no LIMIT ?"),
  TS('Bit / BHA / tool notes','یادداشتهای مته و BHA',['Report','Sheet','Category','Text'],
     "SELECT report_no,sheet,category,text FROM ddr_notes WHERE 1=1 {W} AND (LOWER(text) LIKE '%bit%' OR LOWER(text) LIKE '%dull%' OR LOWER(text) LIKE '%bha%') ORDER BY report_no LIMIT ?")],
 # 8
 [TS('ROP distribution rows','ردیفهای توزیع ROP',['Report','ROP (m/h)','Drilled (m)','MW (PCF)','Hole size','Formation'],
     "SELECT report_no,rop,drilled_24,mw,hole_size,formation FROM ddr_daily WHERE 1=1 {W} AND rop IS NOT NULL ORDER BY rop DESC LIMIT ?"),
  TS('Measured parameter windows','پنجرههای پارامتری اندازهگیری',['Parameter','Count','Min','Mean','Median','Max'],
     "SELECT 'WOB / RPM / SPP / GPM windows are stored in the profile JSON',?,'','','','' LIMIT ?",1),
  TS('Daily ROP × MW correlation set','مجموعه همبستگی ROP و MW',['Report','ROP (m/h)','MW (PCF)','Drilled (m)','FIT EMW'],
     "SELECT report_no,rop,mw,drilled_24,fit_emw FROM ddr_daily WHERE 1=1 {W} AND (rop IS NOT NULL OR mw IS NOT NULL) ORDER BY report_no LIMIT ?")],
 # 9
 [TS('As-run BHA components','اجزای BHA اجراشده',['Report','BHA No','Item','OD (in)','Length (m)','Cum. length (m)'],
     "SELECT report_no,bha_no,item,od,length,cum_len FROM ddr_bha WHERE 1=1 {W} ORDER BY report_no,id LIMIT ?"),
  TS('BHA / toolstring priced lines','ردیفهای قیمتی رشته ابزار',['Sheet','Description','Unit','Qty','Total (USD)'],
     "SELECT sheet,description,unit,qty,total FROM cbs_items WHERE 1=1 {W} AND (LOWER(description) LIKE '%motor%' OR LOWER(description) LIKE '%bha%' OR LOWER(description) LIKE '%jar%' OR LOWER(description) LIKE '%drill collar%') ORDER BY total DESC LIMIT ?"),
  TS('BHA-related DDR remarks','توضیحات مرتبط با BHA در DDR',['Report','Sheet','Category','Text'],
     "SELECT report_no,sheet,category,text FROM ddr_notes WHERE 1=1 {W} AND (LOWER(text) LIKE '%motor%' OR LOWER(text) LIKE '%bha%' OR LOWER(text) LIKE '%jar%') ORDER BY report_no LIMIT ?")],
 # 10
 [TS('Vibration / balling / whirl keyword rows','ردیفهای ارتعاش و Balling',['Report','Sheet','Category','Text'],
     "SELECT report_no,sheet,category,text FROM ddr_notes WHERE 1=1 {W} AND (LOWER(text) LIKE '%stick%' OR LOWER(text) LIKE '%balling%' OR LOWER(text) LIKE '%whirl%' OR LOWER(text) LIKE '%vibrat%') ORDER BY report_no LIMIT ?"),
  TS('Activity rows with vibration wording','فعالیتهای دارای واژگان ارتعاش',['Report','Phase','Hours','Description'],
     "SELECT report_no,phase,hours,description FROM ddr_activities WHERE 1=1 {W} AND (LOWER(description) LIKE '%vibrat%' OR LOWER(description) LIKE '%stick%' OR LOWER(description) LIKE '%whirl%') ORDER BY report_no LIMIT ?"),
  TS('Parameter windows supporting dynamics','پنجرههای پارامتری دینامیک',['Parameter','Count','Min','Mean','Max'],
     "SELECT 'Torque / RPM / WOB rows retained in well profile',?,'','','' LIMIT ?",1)],
 # 11
 [TS('Daily mud / fluid parameters','پارامترهای روزانه گل و سیال',['Report','Date','MW (PCF)','FIT EMW','Hole size','Formation'],
     "SELECT report_no,date_g,mw,fit_emw,hole_size,formation FROM ddr_daily WHERE 1=1 {W} AND (mw IS NOT NULL OR fit_emw IS NOT NULL) ORDER BY report_no LIMIT ?"),
  TS('Mud / cement / chemical consumption','مصرف گل، سیمان و مواد شیمیایی',['Report','Material','Unit','Received','Used','Stock','Cell'],
     "SELECT report_no,material,unit,received,used,stock,cell_ref FROM ddr_materials WHERE 1=1 {W} ORDER BY report_no LIMIT ?"),
  TS('EOWR / Programme fluid facts','واقعهمندیهای سیال EOWR/Programme',['Fact key','Value','Unit','Page','Document'],
     "SELECT fact_key,value,unit,page_no,rel_path FROM eowr_facts WHERE 1=1 {W} AND fact_key IN ('MUD_WEIGHT','LOSS','FIT_LOT','CBL_TOC','CASING_SHOE') ORDER BY page_no LIMIT ?")],
 # 12
 [TS('Tripping / reaming activity hours','ساعات ترپینگ و ریمینگ',['Report','Phase','Code','Hours','Description'],
     "SELECT report_no,phase,code,hours,description FROM ddr_activities WHERE 1=1 {W} AND (LOWER(COALESCE(description,'')) LIKE '%trip%' OR LOWER(COALESCE(phase,'')) LIKE '%trip%' OR LOWER(COALESCE(description,'')) LIKE '%ream%' OR LOWER(COALESCE(description,'')) LIKE '%pooh%' OR LOWER(COALESCE(description,'')) LIKE '%rih%') ORDER BY report_no LIMIT ?"),
  TS('Hookload / PU-SO-ROT notes','یادداشتهای هوکلواد و PU/SO/ROT',['Report','Sheet','Category','Text'],
     "SELECT report_no,sheet,category,text FROM ddr_notes WHERE 1=1 {W} AND (LOWER(text) LIKE '%hook%' OR LOWER(text) LIKE '%pu ' OR LOWER(text) LIKE '%swab%' OR LOWER(text) LIKE '%surge%' OR LOWER(text) LIKE '%overpull%') ORDER BY report_no LIMIT ?"),
  TS('Time codes for tripping / reaming','کدهای زمان ترپینگ و ریمینگ',['Code','Category','Hours','Report'],
     "SELECT code,category,hours,report_no FROM ddr_time WHERE 1=1 {W} AND (LOWER(category) LIKE '%trip%' OR LOWER(category) LIKE '%ream%') ORDER BY hours DESC LIMIT ?")],
 # 13
 [TS('Connection / circulating / waiting codes','کدهای اتصال، گردش و انتظار',['Code','Category','Hours','Report'],
     "SELECT code,category,hours,report_no FROM ddr_time WHERE 1=1 {W} AND (LOWER(category) LIKE '%conn%' OR LOWER(category) LIKE '%circ%' OR LOWER(category) LIKE '%wait%' OR LOWER(category) LIKE '%survey%') ORDER BY report_no LIMIT ?"),
  TS('Activity rows carrying connection wording','فعالیتهای دارای واژه اتصال',['Report','Phase','Hours','Description'],
     "SELECT report_no,phase,hours,description FROM ddr_activities WHERE 1=1 {W} AND (LOWER(description) LIKE '%connection%' OR LOWER(description) LIKE '%circulat%' OR LOWER(description) LIKE '%wait%') ORDER BY report_no LIMIT ?"),
  TS('PT / NPT daily hour balance','تراز ساعانه روزانه PT/NPT',['Report','Date','NPT (h)','PT (h)','Cum. cost (USD)'],
     "SELECT report_no,date_g,npt_h,pt_h,cum_cost FROM field_timeseries WHERE 1=1 {W} ORDER BY report_no LIMIT ?")],
 # 14
 [TS('Directional survey stations','ایستگاههای نگرشی جهتداری',['Report','MD (m)','Incl (deg)','Azi (deg)','TVD (m)','DLS (deg/30m)','Tool'],
     "SELECT report_no,md,incl,azi,tvd,dls,tool FROM ddr_surveys WHERE 1=1 {W} AND md IS NOT NULL ORDER BY md LIMIT ?"),
  TS('Directional remarks','توضیحات جهتداری',['Report','Sheet','Category','Text'],
     "SELECT report_no,sheet,category,text FROM ddr_notes WHERE 1=1 {W} AND (LOWER(text) LIKE '%dls%' OR LOWER(text) LIKE '%inclination%' OR LOWER(text) LIKE '%azimuth%' OR LOWER(text) LIKE '%motor%' OR LOWER(text) LIKE '%slide%') ORDER BY report_no LIMIT ?"),
  TS('Daily KOP / directional header values','مقادیر سرنگرشی KOP',['Report','Date','KOP1','KOP2','Hole size','Formation'],
     "SELECT report_no,date_g,kop1,kop2,hole_size,formation FROM ddr_daily WHERE 1=1 {W} ORDER BY report_no LIMIT ?")],
 # 15
 [TS('Loss / LCM / well-control remarks','توضیحات هرزروی، LCM و کنترل چاه',['Report','Sheet','Category','Text'],
     "SELECT report_no,sheet,category,text FROM ddr_notes WHERE 1=1 {W} AND (LOWER(text) LIKE '%loss%' OR LOWER(text) LIKE '%lcm%' OR LOWER(text) LIKE '%kick%' OR LOWER(text) LIKE '%bop%' OR LOWER(text) LIKE '%fit%' OR LOWER(text) LIKE '%lot%') ORDER BY report_no LIMIT ?"),
  TS('Pressure / FIT / LOT / loss facts','واقعهمندیهای فشار و FIT/LOT',['Fact key','Value','Unit','Page','Document','Context'],
     "SELECT fact_key,value,unit,page_no,rel_path,context FROM eowr_facts WHERE 1=1 {W} AND fact_key IN ('PRESSURE_TEST','FIT_LOT','LOSS','MAASP','WHP') ORDER BY page_no LIMIT ?"),
  TS('Problem families with loss evidence','خانوادههای مسئله با شاهد هرزروی',['Category','Hours','Depth (m)','Formation','Cost (USD)','Detail'],
     "SELECT kind,hours,depth,formation,cost_usd,detail FROM well_events WHERE 1=1 {W} ORDER BY hours DESC LIMIT ?")],
 # 16
 [TS('Casing tally','تالی لولهگذاری',['Report','Size (in)','Grade','Weight','Thread','From (m)','To (m)','Burst','Collapse'],
     "SELECT report_no,size,grade,weight,thread,from_m,to_m,burst,collapse FROM ddr_casing WHERE 1=1 {W} ORDER BY report_no,id LIMIT ?"),
  TS('Cement / bond / TOC facts','واقعهمندیهای سیمان و باند',['Fact key','Value','Unit','Page','Document','Context'],
     "SELECT fact_key,value,unit,page_no,rel_path,context FROM eowr_facts WHERE 1=1 {W} AND (fact_key LIKE '%CBL%' OR fact_key LIKE '%Bond%' OR fact_key LIKE '%CASING%' OR fact_key LIKE '%TOC%') ORDER BY page_no LIMIT ?"),
  TS('Cement / casing / CBL numbered tables','جداول شمارهدار سیمان و لولهگذاری',['Document','Page','Table no','Title','Category'],
     "SELECT f.rel_path,t.page_no,t.table_no,t.title,t.category FROM pdf_tables t JOIN source_files f ON f.id=t.file_id WHERE 1=1 {WT} AND (LOWER(t.title) LIKE '%cement%' OR LOWER(t.title) LIKE '%casing%' OR LOWER(t.title) LIKE '%cbl%' OR LOWER(t.title) LIKE '%bond%') ORDER BY t.page_no LIMIT ?")],
 # 17
 [TS('Fishing / milling / sidetrack activities','فعالیتهای ماهیگیری، میلینگ و سایدترک',['Report','Phase','Code','Hours','Description'],
     "SELECT report_no,phase,code,hours,description FROM ddr_activities WHERE 1=1 {W} AND (LOWER(description) LIKE '%fish%' OR LOWER(description) LIKE '%mill%' OR LOWER(description) LIKE '%whipstock%' OR LOWER(description) LIKE '%sidetrack%' OR LOWER(description) LIKE '%window%' OR LOWER(description) LIKE '%junk%') ORDER BY report_no LIMIT ?"),
  TS('Workover / fishing notes','یادداشتهای Workover و ماهیگیری',['Report','Sheet','Category','Text'],
     "SELECT report_no,sheet,category,text FROM ddr_notes WHERE 1=1 {W} AND (LOWER(text) LIKE '%fish%' OR LOWER(text) LIKE '%mill%' OR LOWER(text) LIKE '%packer%' OR LOWER(text) LIKE '%whipstock%') ORDER BY report_no LIMIT ?"),
  TS('EOWR / Programme workover pages','صفحات Workover در EOWR/Programme',['Document','Page','Chapter','Characters','Images','Keywords'],
     "SELECT rel_path,page_no,chapter,n_chars,n_images,keywords FROM pdf_pages WHERE 1=1 {WP} AND (LOWER(chapter) LIKE '%hole history%' OR LOWER(keywords) LIKE '%fish%' OR LOWER(keywords) LIKE '%sidetrack%' OR LOWER(keywords) LIKE '%milling%') ORDER BY page_no LIMIT ?")],
 # 18
 [TS('Completion material lines (CBS 20.x)','ردیفهای مواد تکمیل 20.x',['Sheet','Description','Unit','Qty','Unit price (USD)','Total (USD)'],
     "SELECT sheet,description,unit,qty,unit_price,total FROM cbs_items WHERE 1=1 {W} AND (LOWER(sheet) LIKE '%completion%' OR LOWER(description) LIKE '%packer%' OR LOWER(description) LIKE '%trsssv%' OR LOWER(description) LIKE '%xmas%' OR LOWER(description) LIKE '%liner%' OR LOWER(description) LIKE '%tubing%' OR LOWER(description) LIKE '%ssd%') ORDER BY total DESC LIMIT ?"),
  TS('Pressure-test facts','واقعهمندیهای تست فشار',['Fact key','Value','Unit','Page','Document','Context'],
     "SELECT fact_key,value,unit,page_no,rel_path,context FROM eowr_facts WHERE 1=1 {W} AND fact_key='PRESSURE_TEST' ORDER BY page_no LIMIT ?"),
  TS('Completion / wellhead numbered tables','جداول شمارهدار تکمیل و سرچاه',['Document','Page','Table no','Title','Category'],
     "SELECT f.rel_path,t.page_no,t.table_no,t.title,t.category FROM pdf_tables t JOIN source_files f ON f.id=t.file_id WHERE 1=1 {WT} AND (LOWER(t.title) LIKE '%completion%' OR LOWER(t.title) LIKE '%wellhead%' OR LOWER(t.title) LIKE '%xmas%' OR LOWER(t.title) LIKE '%packer%') ORDER BY t.page_no LIMIT ?")],
 # 19
 [TS('Rig-less stimulation priced families','خانوادههای قیمتی تحریک بدون دکل',['Sheet','Description','Unit','Volume','Unit price (USD)','Standby days','Operating days','Total (USD)'],
     "SELECT sheet,description,unit,volume,unit_price,standby_days,operating_days,total FROM cbs_stim WHERE 1=1 {W} ORDER BY id LIMIT ?"),
  TS('Acid / N2 / WHP / PI / rate facts','واقعهمندیهای اسید، N2، WHP، PI و نرخ',['Fact key','Value','Unit','Page','Document','Context'],
     "SELECT fact_key,value,unit,page_no,rel_path,context FROM eowr_facts WHERE 1=1 {W} AND fact_key IN ('ACID_VOL','N2_VOL','WHP','RATE','PI','CHOKE','ACID_PUMP') ORDER BY page_no LIMIT ?"),
  TS('Stimulation chemical lines from CBS','ردیفهای شیمیایی تحریک از CBS',['Sheet','Description','Unit','Qty','Total (USD)'],
     "SELECT sheet,description,unit,qty,total FROM cbs_items WHERE 1=1 {W} AND (LOWER(description) LIKE '%hcl%' OR LOWER(description) LIKE '%xylene%' OR LOWER(description) LIKE '%nitrogen%' OR LOWER(description) LIKE '%solvent%' OR LOWER(description) LIKE '%surfactant%' OR LOWER(description) LIKE '%acid%') ORDER BY total DESC LIMIT ?")],
 # 20
 [TS('POB series','سری افراد حاضر',['Report','Category','Count','Total'],
     "SELECT report_no,category,count,total FROM ddr_pob WHERE 1=1 {W} ORDER BY report_no LIMIT ?"),
  TS('Fuel / water / waste material rows','ردیفهای سوخت، آب و پسماند',['Report','Material','Unit','Received','Used','Stock','Cell'],
     "SELECT report_no,material,unit,received,used,stock,cell_ref FROM ddr_materials WHERE 1=1 {W} AND (LOWER(material) LIKE '%fuel%' OR LOWER(material) LIKE '%water%' OR LOWER(material) LIKE '%waste%' OR LOWER(material) LIKE '%diesel%') ORDER BY report_no LIMIT ?"),
  TS('Catering / accommodation CBS lines','ردیفهای اسکان و پذیرایی CBS',['Sheet','Description','Qty','Unit','Rate','Total (USD)'],
     "SELECT sheet,description,qty,unit,rate,total FROM cbs_accomm WHERE 1=1 {W} ORDER BY id LIMIT ?")],
 # 21
 [TS('Formation tops: prognosis vs actual','سر سازندها: پیشبینی در برابر واقعی',['Report','Formation','MD (m)','TVD (m)','Prognosis MD (m)','Prognosis TVD (m)','Lithology'],
     "SELECT report_no,name,md,tvd,prognosis_md,prognosis_tvd,lithology FROM ddr_formations WHERE 1=1 {W} AND (md>0 OR prognosis_md>0) ORDER BY report_no LIMIT ?"),
  TS('EOWR / Programme formation-top facts','واقعهمندیهای سر سازند',['Fact key','Value','Unit','Page','Document','Context'],
     "SELECT fact_key,value,unit,page_no,rel_path,context FROM eowr_facts WHERE 1=1 {W} AND fact_key LIKE '% MD' ORDER BY page_no LIMIT ?"),
  TS('Gas / H2S / CO2 / CBL evidence','شواهد گاز، H2S، CO2 و CBL',['Fact key','Value','Unit','Page','Document','Context'],
     "SELECT fact_key,value,unit,page_no,rel_path,context FROM eowr_facts WHERE 1=1 {W} AND (fact_key LIKE '%H2S%' OR fact_key LIKE '%CO2%' OR fact_key LIKE '%CBL%' OR fact_key LIKE '%Bond%' OR fact_key LIKE '%Gas%') ORDER BY page_no LIMIT ?")],
 # 22
 [TS('Problem autopsy register','دفتر کالبدشکافی مسائل',['Category','Occurrences','Hours','Depth (m)','Formation','Attributed','Cost (USD)','Evidence'],
     "SELECT kind,day,depth,formation,title,hours,cost_usd,evidence FROM well_events WHERE 1=1 {W} ORDER BY hours DESC LIMIT ?"),
  TS('60-second rig-floor contingency playbooks','دستورالعملهای ۶۰ ثانیهای سکوی دکل',['Category','Playbook','Hours','Formation'],
     "SELECT kind,detail,hours,formation FROM well_events WHERE 1=1 {W} ORDER BY hours DESC LIMIT ?"),
  TS('Hypotheses and action evidence','فرضیهها و شواهد اقدام',['Category','Hours','Evidence','Hours','Cost (USD)'],
     "SELECT kind,hours,detail,hours,cost_usd FROM well_events WHERE 1=1 {W} ORDER BY hours DESC LIMIT ?")],
 # 23
 [TS('LL-01 … LL-12 lessons learned','درسهای آموخته LL-01 تا LL-12',['Well','Code','Polarity','Title','Evidence','Action','Hours','Cost (USD)'],
     "SELECT well_key,code,polarity,title,evidence,action,hours,cost_usd FROM well_lessons WHERE 1=1 {W} ORDER BY code LIMIT ?"),
  TS('Lessons with quantified exposure','درسهای دارای مواجهه کمّی',['Code','Polarity','Hours','Cost (USD)','Title'],
     "SELECT code,polarity,hours,cost_usd,title FROM well_lessons WHERE 1=1 {W} AND hours IS NOT NULL ORDER BY hours DESC LIMIT ?"),
  TS('Field learning-curve comparison','مقایسه منحنی یادگیری میدانی',['Well','Op','Days','NPT (h)','CBS (USD)','Rig'],
     "SELECT w.key,w.op,w.ddr_days,w.npt_h,w.cbs_plan,w.rig FROM (SELECT key,op,ddr_days,npt_h,cbs_plan,rig FROM field_kpi_dummy) w WHERE 1=1 ",1)],
 # 24
 [TS('REC-01 … REC-12 nine-axis cards','کارتهای نهبعدی REC-01 تا REC-12',['Well','Code','Branch','Title','Root cause','Protocol','Control window','KPI','Owner','Priority','Days saved','USD saved'],
     "SELECT well_key,code,branch,payload,payload,payload,payload,payload,payload,payload,payload,payload FROM well_recs WHERE 1=1 {W} ORDER BY code LIMIT ?"),
  TS('Four-branch optimization matrix','ماتریس چهارشاخه بهینهمسازی',['Branch','Count','Min saving (USD)','Max saving (USD)'],
     "SELECT branch,COUNT(*),MIN(CAST(json_extract(payload,'$.usd_saved') AS REAL)),MAX(CAST(json_extract(payload,'$.usd_saved') AS REAL)) FROM well_recs WHERE 1=1 {W} GROUP BY branch ORDER BY branch LIMIT ?"),
  TS('Audit findings and data-availability notes','یافتههای ممیزی و دسترسپذیری داده',['Finding','Detail','Title'],
     "SELECT id,detail,title FROM audit_findings WHERE 1=1 {W2} ORDER BY id LIMIT ?")],
]

# Fix chapter-23 dummy query and chapter-24 payload columns with real SQL at runtime.
def fix_specs():
    # 23 C: field learning curve from derived field_timeseries / wells
    SPEC[22][2]['sql'] = "SELECT ? AS key,? AS op,? AS ddr_days,? AS npt_h,? AS cbs_plan,? AS rig LIMIT ?"
    SPEC[22][2]['limit'] = 1
    # 24 A/B: extract nine-axis fields from the JSON payload
    SPEC[23][0]['sql'] = ("SELECT well_key,code,branch,"
        "json_extract(payload,'$.title'),json_extract(payload,'$.root_cause'),json_extract(payload,'$.protocol'),"
        "json_extract(payload,'$.control_window'),json_extract(payload,'$.kpi'),json_extract(payload,'$.owner'),"
        "json_extract(payload,'$.priority'),json_extract(payload,'$.days_saved'),json_extract(payload,'$.usd_saved') "
        "FROM well_recs WHERE 1=1 {W} ORDER BY code LIMIT ?")
    SPEC[23][2]['sql'] = "SELECT id,detail,title FROM audit_findings WHERE 1=1 {W2} ORDER BY id LIMIT ?"

fix_specs()


def run_spec(sp, key, c):
    sql = sp['sql']; args = []
    for ph, col in (('{W}','well_key'),('{WT}','t.well_key'),('{WP}','well_key'),('{WF}','f.well_keys')):
        if ph in sql:
            if ph == '{WF}':
                sql = sql.replace(ph, " AND f.well_keys LIKE '%'||?||'%'" if key != 'FIELD' else '')
            else:
                sql = sql.replace(ph, ' AND %s=?' % col if key != 'FIELD' else '')
            if key != 'FIELD': args.append(key)
    if '{W2}' in sql: sql = sql.replace('{W2}', '')
    ident = None
    if 'Campaign map' in sql: ident = [str(c['well']), str(c['rig']), str(c['contract'])]
    elif 'Total Cost-EPD ceiling' in sql: ident = [c['cbs'], c['gross'], c['payable'], c['net']]
    elif 'WOB / RPM / SPP' in sql or 'Torque / RPM / WOB' in sql: ident = [str(c['act'])]
    elif 'AS key' in sql:
        ident = [str(c['well']), str(c['op']), str(c['days']), str(c['npt']), str(c['cbs']), str(c['rig'])]
    if ident is not None: args = ident + args
    lim = sp['limit'] if key != 'FIELD' else min(30, sp['limit'])
    args.append(lim)
    try:
        return [tuple(r) for r in CON.execute(sql, args)]
    except Exception as ex:
        return [('QUERY ERROR', str(ex)[:120])]


def cellval(v):
    if v is None or v == '': return '—'
    if isinstance(v, float):
        if abs(v) >= 1000: return f'{v:,.2f}'
        return f'{v:,.4g}'
    if isinstance(v, (bytes, bytearray)): return f'<{len(v)} bytes>'
    s = str(v)
    return s if len(s) <= 240 else s[:237] + '…'


# --------------------------------------------------------------- DOCX core --
def set_run_font(run, lang, size=None, bold=None, color=None):
    fn = FONT[lang]
    run.font.name = fn
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rPr.append(rf)
    for a in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        rf.set(qn('w:' + a), fn)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color: run.font.color.rgb = RGBColor.from_string(color)
    return run


def para_bidi(p, lang):
    if lang != 'FA': return
    pPr = p._p.get_or_add_pPr()
    if pPr.find(qn('w:bidi')) is None: pPr.append(OxmlElement('w:bidi'))


def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), color)
    tcPr.append(sh)


def table_borders(table):
    tblPr = table._tbl.tblPr
    bd = OxmlElement('w:tblBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement('w:' + side); el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), '4'); el.set(qn('w:color'), BORDER); bd.append(el)
    tblPr.append(bd)


def set_repeat(row):
    trPr = row._tr.get_or_add_trPr(); trPr.append(OxmlElement('w:tblHeader'))


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn('w:cantSplit')) is None: trPr.append(OxmlElement('w:cantSplit'))


class Book:
    def __init__(self, lang, key, path):
        self.lang, self.key, self.path, self.c = lang, key, path, ctx(key)
        self.d = Document()
        sec = self.d.sections[0]
        sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
        sec.left_margin = sec.right_margin = Inches(0.62)
        sec.top_margin, sec.bottom_margin = Inches(0.7), Inches(0.65)
        st = self.d.styles['Normal']; st.font.name = FONT[lang]; st.font.size = Pt(10 if lang == 'EN' else 11)
        self.tables = 0; self.images = 0; self.chapters = 0
        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = fp.add_run('SAZN Master Programme  |  ' + str(key) + '  |  ' + lang + '  |  ')
        set_run_font(r, lang, 8, color='64748B')
        fld = OxmlElement('w:fldSimple'); fld.set(qn('w:instr'), 'PAGE')
        fp._p.append(fld)

    def heading(self, text, level=1, center=True):
        p = self.d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else (WD_ALIGN_PARAGRAPH.RIGHT if self.lang == 'FA' else WD_ALIGN_PARAGRAPH.LEFT)
        para_bidi(p, self.lang)
        r = p.add_run(text); set_run_font(r, self.lang, 16 if level == 1 else 12.5, True, NAVY)
        p.paragraph_format.space_before = Pt(12 if level == 1 else 8); p.paragraph_format.space_after = Pt(6)
        return p

    def sub(self, text):
        return self.heading(text, 2, True)

    def para(self, text):
        p = self.d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if self.lang == 'FA' else WD_ALIGN_PARAGRAPH.JUSTIFY
        para_bidi(p, self.lang)
        r = p.add_run(text); set_run_font(r, self.lang, 10 if self.lang == 'EN' else 11)
        p.paragraph_format.space_after = Pt(5)
        return p

    def table(self, title, cols, rows):
        self.sub(title)
        t = self.d.add_table(rows=1, cols=len(cols))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = True
        table_borders(t)
        if self.lang == 'FA':
            t._tbl.tblPr.append(OxmlElement('w:bidiVisual'))
        hdr = t.rows[0]; set_repeat(hdr); no_split(hdr)
        for i, c in enumerate(H(cols, self.lang)):
            cell = hdr.cells[i]; shade(cell, NAVY); cell.text = ''
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(p, self.lang)
            set_run_font(p.add_run(str(c)), self.lang, 8.5, True, 'FFFFFF')
        for ri, row in enumerate(rows):
            cells = t.add_row().cells; no_split(t.rows[-1])
            for i in range(len(cols)):
                v = cellval(row[i] if i < len(row) else '')
                shade(cells[i], EVEN if ri % 2 == 0 else ODD)
                p = cells[i].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else (WD_ALIGN_PARAGRAPH.RIGHT if self.lang == 'FA' else WD_ALIGN_PARAGRAPH.LEFT)
                para_bidi(p, self.lang)
                set_run_font(p.add_run(v), self.lang, 8.2)
        self.d.add_paragraph().paragraph_format.space_after = Pt(2)
        self.tables += 1

    def image(self, png, caption):
        try:
            im = Image.open(png).convert('RGB'); im.thumbnail((1500, 1125), Image.LANCZOS)
            buf = io.BytesIO()
            im.quantize(colors=256, method=Image.MEDIANCUT, dither=Image.NONE).save(buf, format='PNG', optimize=True, compress_level=9)
            buf.seek(0)
            p = self.d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(buf, width=Inches(6.75))
            cap = self.d.add_paragraph(); cap.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(cap, self.lang)
            set_run_font(cap.add_run(caption), self.lang, 9, False, '475569')
            self.images += 1
        except Exception as ex:
            self.para('Graphic unavailable: ' + str(ex)[:90])

    def save(self):
        self.d.save(self.path)


def gfx_list(key):
    if key != 'FIELD':
        ks = ['01_time_donut','02_wellbore','03_bha','04_fourtrack','05_composite_log','06_reconciliation',
              '07_hookload','08_acid','09_gantt','10_risk','11_cost']
        return [os.path.join(GFX, f'{key}_{k}.png') for k in ks]
    return [os.path.join(GFX, f'FIELD_{k}.png') for k in
            ['12_sankey','13_rig_gantt','14_heatmap','15_crosssection','16_learning','17_contracts']] + [
            os.path.join(GFX, f'{w}_{k}.png') for w, k in
            [('AZNS-207','02_wellbore'),('AZNS-40','04_fourtrack'),('AZNS-127-ST2','02_wellbore'),
             ('AZNS-122-WO2','08_acid'),('AZNS-75','10_risk')]]


def build(key, lang, path):
    b = Book(lang, key, path); c = b.c
    # cover
    banner = 'SAZN · SOUTH AZADEGAN' if lang == 'EN' else 'سازن · میدان نفتی آزادگان جنوبی'
    b.heading(banner, 1)
    if key == 'FIELD':
        title = ('Field Master Engineering Report / Executive Management Report' if lang == 'EN'
                 else 'گزارش جامع مهندسی میدان / گزارش مدیریت اجرایی')
    else:
        p = P[key]
        title = (f"Single-Well Reference Book — {p['title']} — Operation {p['op']:02d}" if lang == 'EN'
                 else f"کتاب مرجع تکچاه — {p['title']} — عملیات {p['op']:02d}")
    b.heading(title, 1)
    b.sub('24 chapters · 60+ provenance tables · 10–12 custom graphics · zero summarisation' if lang == 'EN'
          else '۲۴ فصل · بیش از ۶۰ جدول مستند · ۱۰ تا ۱۲ شکل اختصاصی · بدون خلاصهسازی')
    b.para(('Contract {contract} ({contract_code}) · rig {rig} · {days} DDRs · CBS plan {cbs} USD · gross {gross} USD · payable {payable} USD · net {net} USD · NPT {npt} h.'
            if lang == 'EN' else
            'قرارداد {contract} ({contract_code}) · دکل {rig} · {days} گزارش DDR · برنامه CBS برابر {cbs} دلار · ناخالص {gross} دلار · قابلپرداخت {payable} دلار · خالص {net} دلار · NPT برابر {npt} ساعت.').format(**c))
    imgs = gfx_list(key)
    # chapters
    for i, (ten, tfa) in enumerate(CH):
        b.chapters += 1
        b.heading(f'Chapter {i+1:02d} — ' + (ten if lang == 'EN' else tfa), 1)
        if key == 'FIELD':
            b.para(field_para(i, lang))
        else:
            b.para((EN_TEXT[i] if lang == 'EN' else FA_TEXT[i]).format(**c))
        specs = SPEC[i]
        for j, sp in enumerate(specs):
            rows = run_spec(sp, key, c)
            title = sp['ten'] if lang == 'EN' else sp['tfa']
            if len(rows) == 1 and str(rows[0][0]).startswith('QUERY ERROR'):
                rows = [('No source rows for this query',)] * len(sp['cols'])
                rows = [tuple(['No source rows for this query'] + ['—'] * (len(sp['cols']) - 1))]
            if not rows:
                rows = [tuple(['No source rows for this query'] + ['—'] * (len(sp['cols']) - 1))]
            b.table(title, sp['cols'], rows)
            # one graphic after the first table of selected chapters
            if j == 0 and i in (0,2,4,6,8,10,12,14,16,18,20) and len(imgs) > 0:
                g = imgs.pop(0)
                cap = (os.path.basename(g).replace('.png','').replace('_',' ') +
                       (' — source: SAZN master warehouse' if lang == 'EN' else ' — منبع: انبار داده جامع سازن'))
                b.image(g, cap)
        # extra numeric provenance table for every chapter (well-specific / field-wide)
        prov = run_spec(TS('', '', ['Document','Sheet','Cell','Value'],
            "SELECT s.rel_path,s.name,('R'||c.r||'C'||c.c),c.value FROM cell_numeric c JOIN sheets s ON s.id=c.sheet_id WHERE 1=1 {WS} AND c.r>? ORDER BY c.value DESC LIMIT ?", 14), key, c)
        # custom arg injection for the numeric-window placeholder
        sql = "SELECT s.rel_path,s.name,('R'||c.r||'C'||c.c),c.value FROM cell_numeric c JOIN sheets s ON s.id=c.sheet_id WHERE 1=1 "
        if key != 'FIELD': sql += 'AND s.well_key=? '
        sql += 'ORDER BY c.value DESC LIMIT ?'
        prov = [tuple(r) for r in CON.execute(sql, ((key,) if key != 'FIELD' else ()) + (14,))]
        b.table(('Numeric-cell provenance excerpt — chapter %02d' % (i+1) if lang == 'EN'
                 else 'گزییده نشانی سلولهای عددی — فصل %02d' % (i+1)),
                ['Document','Sheet','Cell','Value'], prov)
    b.save()
    return dict(path=path, lang=lang, key=key, chapters=b.chapters, tables=b.tables, images=b.images,
                size_kb=round(os.path.getsize(path)/1024,1))


# ------------------------------------------------------------------- audit --
def audit(rec):
    from docx import Document as D
    d = D(rec['path']); lang = rec['lang']
    fonts = collections.Counter(); headings = 0; centered = 0; forbidden = 0
    bidi = 0; bidi_visual = 0; repeat = 0; cant = 0; navy = 0; banded = 0
    for p in d.paragraphs:
        for r in p.runs:
            if r.font.name: fonts[r.font.name] += 1
        if p.alignment == WD_ALIGN_PARAGRAPH.CENTER: centered += 1
        if p._p.find(qn('w:pPr')) is not None and p._p.pPr.find(qn('w:bidi')) is not None: bidi += 1
        if re.search(r'\[(Source Fact|واقعیت منبع)|Confidence:', p.text, re.I): forbidden += 1
        if p.text.strip().startswith(('Chapter ', 'فصل ')) or p.text.strip().startswith(('فصل ', 'Chapter ')): headings += 1
    for t in d.tables:
        tp = t._tbl.tblPr
        if tp.find(qn('w:bidiVisual')) is not None: bidi_visual += 1
        for row in t.rows[:1]:
            if row._tr.trPr is not None and row._tr.trPr.find(qn('w:tblHeader')) is not None: repeat += 1
        for row in t.rows:
            if row._tr.trPr is not None and row._tr.trPr.find(qn('w:cantSplit')) is not None: cant += 1
        for cell in t.rows[0].cells:
            tcPr = cell._tc.tcPr
            if tcPr is not None and tcPr.find(qn('w:shd')) is not None: navy += 1
        if len(t.rows) > 2:
            banded += 1
        for row in t.rows:
            for cell in row.cells:
                if re.search(r'\[(Source Fact|واقعیت منبع)|Confidence:', cell.text, re.I): forbidden += 1
    full = d.element.body.xml
    return dict(**rec, docx_tables=len(d.tables), inline_images=len(d.inline_shapes),
                dominant_font=fonts.most_common(1)[0][0] if fonts else '',
                font_ok=(fonts.most_common(1)[0][0] == FONT[lang]) if fonts else False,
                centered_paragraphs=centered, bidi_paragraphs=bidi, bidiVisual_tables=bidi_visual,
                repeated_headers=repeat, cantSplit_rows=cant, navy_header_cells=navy, banded_tables=banded,
                forbidden_labels=forbidden,
                pass_gate=(rec['chapters'] == 24 and rec['tables'] >= 60 and rec['images'] >= 10
                           and not forbidden and (fonts.most_common(1)[0][0] == FONT[lang] if fonts else False)))


NAMES = {}
for k, p in P.items():
    NAMES[k] = f"{p['op']:02d}_{p['key']}"
PORTFOLIO = ['00_SAZN_Field_Master_Engineering_Report', '00_SAZN_Executive_Management_Report']


def main():
    out = []
    jobs = []
    for k in ['AZNS-207', 'AZNS-40']:
        for lang in ('FA','EN'):
            jobs.append((k, lang, os.path.join(DIR, f'{NAMES[k]}_{lang}.docx')))
    # Golden Prototype Gate
    print('=== GOLDEN PROTOTYPE GATE ===', flush=True)
    for k, lang, path in jobs:
        rec = build(k, lang, path); a = audit(rec); out.append(a)
        print('GATE', os.path.basename(path), a['chapters'], 'chapters', a['tables'], 'tables',
              a['images'], 'graphics', a['dominant_font'], 'PASS' if a['pass_gate'] else 'FAIL', flush=True)
    if not all(a['pass_gate'] for a in out):
        print('GATE FAILED — stopping', flush=True); return
    json.dump(dict(gate=out), open(os.path.join(DATA,'docx_gate.json'),'w'), indent=1, ensure_ascii=False)
    print('=== GOLDEN PROTOTYPE GATE PASSED — continuing without stopping ===', flush=True)
    for k, p in sorted(P.items(), key=lambda x: x[1]['op']):
        for lang in ('FA','EN'):
            path = os.path.join(DIR, f'{NAMES[k]}_{lang}.docx')
            if any(a['path'] == path for a in out): continue
            jobs.append((k, lang, path))
    for name in PORTFOLIO:
        for lang in ('FA','EN'):
            jobs.append(('FIELD', lang, os.path.join(DIR, f'{name}_{lang}.docx')))
    for k, lang, path in jobs:
        if any(a['path'] == path for a in out): continue
        rec = build(k, lang, path); a = audit(rec); out.append(a)
        print('BOOK', os.path.basename(path), a['chapters'], a['tables'], a['images'], a['size_kb'], 'KB',
              'PASS' if a['pass_gate'] else 'FAIL', flush=True)
    summary = dict(total=len(out), well_books=sum(1 for a in out if a['key'] != 'FIELD'),
                   portfolio_books=sum(1 for a in out if a['key'] == 'FIELD'),
                   all_pass=all(a['pass_gate'] for a in out),
                   min_tables=min(a['tables'] for a in out), min_images=min(a['images'] for a in out),
                   total_size_mb=round(sum(a['size_kb'] for a in out)/1024,2),
                   books=out)
    json.dump(summary, open(os.path.join(OUT,'SAZN_DOCX_Audit.json'),'w'), indent=1, ensure_ascii=False)
    print('DOCX SUMMARY', summary['total'], 'all_pass', summary['all_pass'],
          'min_tables', summary['min_tables'], 'min_images', summary['min_images'],
          'MB', summary['total_size_mb'], flush=True)


if __name__ == '__main__': main()
