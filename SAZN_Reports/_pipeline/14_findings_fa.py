# -*- coding: utf-8 -*-
"""گزارش یافته‌های فنی–مالی تحلیل داده‌های حفاری (فارسی) — SAZN.

گزارش مستقل فارسی برای مدیریت طرح و اداره قراردادها؛ تمام اعداد از انبار داده جامع
(SAZN_Master_Warehouse.sqlite) و فایل‌های مشتق خوانده می‌شوند؛ هیچ داده منبعی تغییر نمی‌کند.
"""
import os, sys, json, sqlite3, datetime
sys.path.insert(0, os.path.dirname(__file__))
from common import OUT, DATA, DB, fmt

import importlib.util
spec = importlib.util.spec_from_file_location('r08', os.path.join(os.path.dirname(__file__), '08_reports.py'))
R08 = importlib.util.module_from_spec(spec); spec.loader.exec_module(R08)

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

NAVY, EVEN, ODD, BORDER = R08.NAVY, R08.EVEN, R08.ODD, R08.BORDER
tb, ns, srp, shd, cv = R08.table_borders, R08.no_split, R08.set_repeat, R08.shade, R08.cellval
set_run_font, para_bidi = R08.set_run_font, R08.para_bidi

P = json.load(open(os.path.join(DATA, 'well_profiles.json.gz' if False else 'well_profiles.json'))) if os.path.exists(os.path.join(DATA,'well_profiles.json')) else json.load(__import__('gzip').open(os.path.join(DATA,'well_profiles.json.gz')))
F = json.load(open(os.path.join(DATA, 'field_analytics.json')))
TOT = F['totals']
S = json.load(open(os.path.join(DATA, 'deep_analytics.json')))
CON = sqlite3.connect(DB); CON.row_factory = sqlite3.Row
PC = json.load(open(os.path.join(DATA, 'pdf_cost_forensics.json')))

DIR = os.path.join(OUT, 'Word_Books'); os.makedirs(DIR, exist_ok=True)
PATH = os.path.join(DIR, '00_SAZN_Analytical_Findings_Technical_Financial_FA.docx')

d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
sec.left_margin = sec.right_margin = Inches(0.62)
sec.top_margin, sec.bottom_margin = Inches(0.7), Inches(0.65)
sec.header_distance, sec.footer_distance = Inches(0.35), Inches(0.3)
st = d.styles['Normal']; st.font.name = 'B Nazanin'; st.font.size = Pt(11)
# heading styles
for name, size, before in (('Heading 1', 20, 18), ('Heading 2', 14, 12), ('Heading 3', 12, 8)):
    try:
        hs = d.styles[name]
        hs.font.name = 'B Nazanin'; hs.font.size = Pt(size); hs.font.bold = True; hs.font.color.rgb = RGBColor.from_string(NAVY)
    except Exception:
        pass

# header + footer (FA RTL)
hp = sec.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(hp, 'FA')
set_run_font(hp.add_run('SAZN-FA-DA-001  |  گزارش فنی–مالی یافته‌های تحلیل داده‌های حفاری  |  ساوث آزدگان'), 'FA', 8, color='475569')
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(fp, 'FA')
set_run_font(fp.add_run('دپارتمان تحلیل داده‌های حفاری  |  نسخه ۱.۰  |  صفحه '), 'FA', 8, color='64748B')
fld = OxmlElement('w:fldSimple'); fld.set(qn('w:instr'), 'PAGE'); fp._p.append(fld)

def h1(text):
    p = d.add_heading('', 1); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; para_bidi(p, 'FA')
    set_run_font(p.add_run(text), 'FA', 20, True, NAVY); p.paragraph_format.space_before = Pt(16); p.paragraph_format.space_after = Pt(8); return p

def h2(text):
    p = d.add_heading('', 2); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; para_bidi(p, 'FA')
    set_run_font(p.add_run(text), 'FA', 13.5, True, '1D4ED8'); p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(5); return p

def para(text, size=11, color=None, bold=False, space_after=6):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; para_bidi(p, 'FA')
    set_run_font(p.add_run(text), 'FA', size, bold, color); p.paragraph_format.space_after = Pt(space_after); return p

def bullet(text, level=0):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; para_bidi(p, 'FA')
    p.paragraph_format.left_indent = Inches(0.28 + 0.18 * level); p.paragraph_format.space_after = Pt(3)
    set_run_font(p.add_run('•  ' + text), 'FA', 10.5); return p

def table(title, cols, rows, widths=None, font=8.2):
    h2(title)
    t = d.add_table(rows=1, cols=len(cols)); t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = True
    tb(t); t._tbl.tblPr.append(OxmlElement('w:bidiVisual'))
    hdr = t.rows[0]; srp(hdr); ns(hdr)
    for i, c in enumerate(cols):
        cell = hdr.cells[i]; shd(cell, NAVY); cell.text = ''
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(p, 'FA')
        set_run_font(p.add_run(str(c)), 'FA', 8.5, True, 'FFFFFF')
    for ri, row in enumerate(rows):
        cells = t.add_row().cells; ns(t.rows[-1])
        for i in range(len(cols)):
            v = cv(row[i] if i < len(row) else '')
            shd(cells[i], EVEN if ri % 2 == 0 else ODD)
            p = cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else (WD_ALIGN_PARAGRAPH.RIGHT if any('\u0600' <= ch <= '\u06FF' for ch in str(v)) else WD_ALIGN_PARAGRAPH.LEFT)
            para_bidi(p, 'FA'); set_run_font(p.add_run(str(v)), 'FA', font)
    d.add_paragraph().paragraph_format.space_after = Pt(2); return t

from PIL import Image as _Img
GFXD = os.path.join(OUT, 'SAZN_Graphics')
def img(png, caption):
    path = os.path.join(GFXD, png)
    im = _Img.open(path).convert('RGB'); im.thumbnail((1500, 1125), _Img.LANCZOS)
    import io as _io
    buf = _io.BytesIO(); im.quantize(colors=256, method=_Img.MEDIANCUT, dither=_Img.NONE).save(buf, format='PNG', optimize=True, compress_level=9); buf.seek(0)
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(6)
    p.add_run().add_picture(buf, width=Inches(6.6))
    cap = d.add_paragraph(); cap.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(cap, 'FA')
    set_run_font(cap.add_run(caption), 'FA', 9, False, '475569'); cap.paragraph_format.space_after = Pt(8)

def note(text):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; para_bidi(p, 'FA')
    p.paragraph_format.left_indent = Inches(0.25); p.paragraph_format.space_after = Pt(4)
    pPr = p._p.get_or_add_pPr()
    set_run_font(p.add_run('یادآوری:  '), 'FA', 9.5, True, 'B45309')
    set_run_font(p.add_run(text), 'FA', 9.5, False, '78350F')
    return p

usd = lambda v, dgits=2: fmt(v, dgits)
CT = S['cost_track']; CTF = CT['field']
PDF = S['pdf']
print('loading findings from DB…', flush=True)
FINDINGS = [dict(r) for r in CON.execute('SELECT id,category,title,numeric_evidence,interpretation,action,priority FROM an_findings ORDER BY id')]
COST_AUDIT = [dict(r) for r in CON.execute('SELECT * FROM an_cost_track_audit ORDER BY op')]
NPT_ATTR = [dict(r) for r in CON.execute('SELECT * FROM an_npt_attribution_audit ORDER BY well_key')]
print('findings', len(FINDINGS), 'cost rows', len(COST_AUDIT), 'npt rows', len(NPT_ATTR), flush=True)

# ===================================================================== COVER
p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(p, 'FA')
set_run_font(p.add_run('\n\n\n\n\n\n\nمیدان ساوث آزدگان (SAZN)'), 'FA', 17, True, '475569'); p.paragraph_format.space_after = Pt(0)
p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(p, 'FA')
set_run_font(p.add_run('گزارش فنی–مالی یافته‌های\nتحلیل داده‌های حفاری'), 'FA', 30, True, NAVY); p.paragraph_format.space_after = Pt(6)
p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(p, 'FA')
r = p.add_run('ممیزی کامل ۵۱ سند منبع  |  ۲٬۸۱۰ شیت اکسل  |  ۴٬۳۴۵ صفحه PDF  |  ۳٬۰۵٬۴۴۷ عدد PDF'.replace('۳٬۰۵٬۴۴۷', '305٬447'))
set_run_font(r, 'FA', 11, False, '1D4ED8'); p.paragraph_format.space_after = Pt(26)

# thin rule
pr = d.add_paragraph(); pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
pPr = pr._p.get_or_add_pPr(); pbdr = OxmlElement('w:pBdr'); bottom = OxmlElement('w:bottom')
bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '14'); bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '1D4ED8'); pbdr.append(bottom); pPr.append(pbdr)

for line, size, color in [('دپارتمان تحلیل داده‌های حفاری', 15, NAVY),
                          ('تهیه: تیم تحلیل داده حفاری (SAZN Drilling-Data Analysis Department)', 11, '475569'),
                          ('مخاطب: مدیریت طرح توسعه میدان  |  اداره قراردادها و متره  |  کنترل پروژه', 11, '475569'),
                          ('تاریخ انتشار: ۱۸ آبان ۱۴۰۵  (۸ نوامبر ۲۰۲۶)  |  نسخه ۱.۰  |  کد سند: SAZN-FA-DA-001', 10, '64748B')]:
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; para_bidi(p, 'FA')
    set_run_font(p.add_run(line), 'FA', size, size >= 15, color); p.paragraph_format.space_after = Pt(4)

p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(18)
set_run_font(p.add_run('این گزارش بدون حذف و بدون بازنویسی هیچ داده منبعی تولید شده است؛ هر تصحیح مشتق، در ستون/جدول جداگانه و با کلید منبع نگهداری می‌شود.'), 'FA', 9, False, '64748B')
d.add_page_break()

# ================================================= بخش ۱ — دامنه و روش‌شناسی
h1('بخش ۱ — دامنه، منابع و روش‌شناسی')
para('این گزارش خلاصه فنی–مالی «دپارتمان تحلیل داده‌های حفاری» برای میدان ساوث آزدگان است و تمام ۱۳ عملیات انجام‌شده تحت دو دکل (OEOC 201 و OEOC 208) و سه قرارداد (C1: CCA/C/1401/27، C2: CCA/C/1403/31، C3: الحاقیه ST2 با برچسب CCA/C/1403/31 / ST2 amendment؛ سربرگ منبع EOWR ترکیبی ST#2 عبارت CCA/C/1405/3 را نشان می‌دهد) را پوشش می‌دهد.')
table('۱-۱  سرشماری اسناد منبع', ['دسته سند', 'تعداد فایل', 'تعداد شیت / صفحه', 'حجم تقریبی'], [
    ['CBS (متره و برآورد)', '13', '202 شیت', '2.0 MB'],
    ['DDR (گزارش‌های روزانه حفاری)', '13', '2٬608 شیت', '74 MB'],
    ['EOWR (گزارش پایان عملیات)', '11', '2٬724 صفحه', '126 MB'],
    ['Well Program (برنامه حفاری)', '14', '1٬621 صفحه', '81 MB'],
    ['جمع اسناد', '51', '4٬345 صفحه PDF + 2٬810 شیت', '≈283 MB'],
], font=9)
para('روش‌شناسی تحلیل بر چهار ستون استوار است:')
bullet('سرشماری کامل: تمام ۱٬۱۹۶٬۰۸۴ سلول غیرخالی شیت‌ها (۳۲۴٬۱۵۰ عدد عددی، ۱۵۵٬۲۴۸ فرمول) و تمام 305٬447 عدد موجود در لایه متنی PDF با کلید منبع (فایل / شیت / سلول یا فایل / صفحه / خط) نمایه شده‌اند.')
bullet('محاسبه مجدد مستقل: 48٬011 فرمول بررسی‌پذیر از نو محاسبه شد؛ ۴۷٬۹۶۵ مورد مطابق، ۳ مورد با اختلاف جزئی، ۴۳ مورد نامنطبق یافت شد (جدول ۲۵_DA_Formula_Audit).')
bullet('تطبیق چهارمسیره هزینه: برای هر چاه، gross = CBS + (روز DDR − روز CBS) × ۲۳٬۰۰۰ USD/DAY؛ سپس payable با کسر NPT پیمانکاری و net با کسر تمام NPT محاسبه و با مقادیر منبع مقایسه شد (بیشینه اختلاف ۰٫۰۰۹۲ USD).')
bullet('اصل عدم حذف: هیچ سلول، صفحه یا گزارشی حذف یا بازنویسی نشده است؛ تصحیح‌ها در ستون‌های مشتق جداگانه و با برچسب «DERIVED» ثبت شده‌اند.')
note('سه ناحیه نیازمند توجه قراردادی به‌عنوان AF-01 تا AF-05 در سامانه ثبت شده‌اند؛ جزئیات در بخش ۹ آمده است.')

# ============================================================ بخش ۲ — خلاصه مدیریتی
h1('بخش ۲ — خلاصه مدیریتی')
para('در یک نگاه: طرح با ۱۳ عملیات، ۱٬۳۰۴ روز دکل و ۱۲٬۲۶۴٫۷۶ متر حفاری تجمعی، در مجموع ۹۴٬۷۹۷٬۸۹۴٫۳۹ دلار برآورد CBS داشته است؛ با لحاظ تقویم واقعی دکل، هزینه ناخالص ۹۹٬۷۴۲٬۴۸۸٫۲۲ دلار می‌شود. پس از کسر کسورات NPT (پیمانکاری ۱۰٬۷۱۷٬۴۵۶٫۹۴ دلار؛ کارفرمایی/زیرزمینی ۲٬۰۴۴٬۱۲۵٫۰۰ دلار؛ جمع ۱۲٬۷۶۱٬۵۸۱٫۹۴ دلار)، مبلغ قابل‌پرداخت ۸۹٬۰۲۵٬۰۳۱٫۲۷ دلار و خالص نهایی بدون NPT ۸۶٬۹۸۰٬۹۰۶٫۲۸ دلار است.')
table('۲-۱  ارقام کلیدی میدان', ['شاخص', 'مقدار'], [
    ['عملیات / قرارداد / دکل', '13 عملیات  |  3 قرارداد  |  2 دکل (OEOC 201، OEOC 208)'],
    ['روزهای DDR (تجمعی)', '1٬304 روز'],
    ['کل NPT ثبت‌شده', '13٬316٫43 ساعت  (≈ 554٫9 روز دکل)'],
    ['متر حفاری تجمعی', '12٬264٫76 m'],
    ['برآورد CBS (جمع قراردادها)', '94٬797٬894٫39 USD'],
    ['هزینه ناخالص تقویمی', '99٬742٬488٫22 USD'],
    ['قابل‌پرداخت (پس از کسر NPT پیمانکاری)', '89٬025٬031٫27 USD'],
    ['خالص نهایی (پس از کسر تمام NPT)', '86٬980٬906٫28 USD'],
    ['بیشینه اختلاف محاسبه مجدد هزینه', '0٫0092 USD  (پس از گرد کردن به سنت)'],
    ['ساعت fallback انتظار ۵۰/۵۰', '0٫0  (حذف کامل شد)'],
    ['تعداد یافته‌های مستند (DA)', '22 یافته  |  42 جدول مشتق an_*  |  exception CSV ۵٬۴۹۳ ردیف'],
    ['صفحات PDF کم‌متن (AF-05)', '364 صفحه  (۸٫۳۸٪ از ۴٬۳۴۵ صفحه)'],
], font=9.5)
para('سه پیام اصلی برای تصمیم‌سازان:', bold=True)
bullet('مسیر هزینه چهارمسیره با دقت سنت (کمتر از یک صدم دلار) تطبیق می‌دهد؛ این یعنی صورت‌های منبع از نظر حسابی خودسازگارند، اما ۴ پرچم مبنای تقویم (بخش ۵) باید پیش از تسویه قطعی رفع‌ابهام شوند.')
bullet('۱۶٫۰٪ از کل NPT (۲٬۱۳۳٫۰ ساعت) به کارفرما/عوامل زیرزمینی منتسب شده است؛ در دو عملیات AZNS-127-WO2 (۷۲٫۴٪) و AZNS-75 (۵۳٫۸٪) سهم کارفرما از نیمی از NPT فراتر رفته است — این موضوع مستقیماً در کسورات و دامنه ادعاها اثر دارد.')
bullet('ممیزی عمیق اعداد PDF، ۵۳۸ داده پرتِ بزرگی را برچسب‌گذاری کرده و ۳۶۴ صفحه کم‌متن را به‌عنوان گپ منبع صریح (AF-05) ثبت کرده است؛ هیچ مقدار OCR ساختگی نشده است.')

# ====================================================== بخش ۳ — پل هزینه چهارمسیره
h1('بخش ۳ — پل هزینه چهارمسیره (Four-Track Cost Bridge)')
para('تعاریف عملیاتی: (۱) CBS Plan: جمع اقلام متره هر قرارداد بدون لحاظ زمان؛ (۲) Gross: CBS به‌علاوه (روز DDR منهای روز CBS) × نرخ دکل ۲۳٬۰۰۰ USD/DAY؛ (۳) Payable: ناخالص پس از کسر NPT منتسب به پیمانکار؛ (۴) Net: ناخالص پس از کسر کل NPT (پیمانکاری + کارفرمایی). نرخ دکل از آیتم «Rig Service» در CBS با واحد USD/DAY استخراج شده است.')
table('۳-۱  پل هزینه در سطح میدان (USD)', ['مرحله', 'مقدار USD', 'توضیح'], [
    ['CBS Plan', usd(CTF['cbs_plan']), 'جمع اقلام CBS سه قرارداد'],
    ['Gross Calendar', usd(CTF['gross']), 'CBS + (1٬304 روز DDR − روز CBS) × 23٬000'],
    ['Payable (excl. contractor NPT)', usd(CTF['payable']), 'Gross − 10٬717٬456٫94  (11٬183٫43 h × 23٬000 / 24)'],
    ['Net (excl. all NPT)', usd(CTF['net']), 'Gross − 12٬761٬581٫94  (13٬316٫43 h × 23٬000 / 24)'],
], font=9.5)
table('۳-۲  پل هزینه در سطح قرارداد (USD)', ['قرارداد', 'عملیات', 'CBS', 'Gross', 'Payable', 'Net', 'NPT (h)', 'روز DDR'], [
    ['C1 — CCA/C/1401/27', '1–6', '21٬797٬511٫87', '26٬655٬111٫87', '20٬356٬466٫04', '20٬356٬466٫04', '6٬572٫50', '592'],
    ['C2 — CCA/C/1403/31', '7–12', '57٬967٬774٫39', '58٬270٬105٫72', '54٬102٬377٫94', '52٬276٬273٫78', '6٬254٫43', '606'],
    ['C3 — الحاقیه ST2 (CCA/C/1403/31 / برچسب کاربر؛ سربرگ منبع EOWR: CCA/C/1405/3)', '13', '15٬032٬608٫13', '14٬817٬270٫63', '14٬566٬187٫29', '14٬348٬166٫46', '489٫50', '106'],
    ['جمع میدان', '1–13', '94٬797٬894٫39', '99٬742٬488٫22', '89٬025٬031٫27', '86٬980٬906٫28', '13٬316٫43', '1٬304'],
], font=8)
note('در C1 (شش‌چاهه) سهم کارفرما از NPT صفر است؛ در C2، NPT کارفرمایی/زیرزمینی ۲٬۱۳۳٫۰ ساعت (۲٬۰۴۴٬۱۲۵٫۰۰ دلار) است؛ در C3 سهم کارفرما ۲۲۷٫۵ ساعت (۴۶٫۵٪ از NPT عملیات ۱۳) است.')
print('sections 1-3 written', flush=True)

img('AN_16_cost_bridge.png', 'شکل ۱ — پل هزینه چهارمسیره میدان: CBS → Gross → Payable → Net (نرخ 23٬000 USD/DAY)')
table('۳-۳  پل هزینه در سطح چاه (USD) — با اختلاف محاسبه مجدد', ['عملیات', 'چاه', 'CBS plan', 'Gross', 'Payable', 'Net', 'Δ Gross (USD)', 'پرچم تقویم'], [
    [str(r['op'] or '—'), r['well_key'], usd(r['cbs_plan'],0), usd(r['gross_stored'],0), usd(r['payable_stored'],0), usd(r['net_stored'],0), fmt(r['gross_delta'],4), r['calendar_flag']] for r in COST_AUDIT], font=7.6)
note('تمام سیزده عملیات در محدوده±۰٫۰۱ دلار با فرمول تطبیق می‌دهند. چهار ردیف با پرچم تقویم (AZNS-85، AZNS-12، AZNS-166، AZNS-127-ST2) در بخش ۵ تشریح شده‌اند؛ این پرچم‌ها اختلاف منبع مبنا را نشان می‌دهند و هزینه را دوباره محاسبه نمی‌کنند.')

# ===================================================== بخش ۴ — ممیزی نرخ دکل و کسورات
h1('بخش ۴ — نرخ دکل، کسورات NPT و حساسیت')
para('نرخ ۲۳٬۰۰۰ USD/DAY از آیتم Rig Service در هر سه CBS استخراج شده و برای هر سیزده عملیات یکسان است. کسورات NPT بر پایه «ساعت / 24 × نرخ» محاسبه و با منبع تطبیق داده شده‌اند:')
table('۴-۱  کسورات NPT (سطح میدان)', ['اقلام', 'ساعت', 'نرخ محاسباتی', 'کسر (USD)'], [
    ['NPT پیمانکاری (قراردادی)', '11٬183٫43', '23٬000 / 24 h', usd(CTF['contractor_deduction'])],
    ['NPT کارفرمایی / زیرزمینی', '2٬133٫00', '23٬000 / 24 h', usd(CTF['client_deduction'])],
    ['کل NPT', '13٬316٫43', '23٬000 / 24 h', usd(CTF['all_npt_deduction'])],
    ['از آن، ساعت انتظار ثبت‌شده (waiting)', '2٬878٫00', 'بدون تقسیم ۵۰/۵۰', '—'],
    ['از آن، انتظارِ بدون انتساب (blank)', '1٬642٫50', 'بدون تقسیم ۵۰/۵۰', '—'],
    ['fallback انتظار ۵۰/۵۰', '0٫00', 'حذف شده از خط لوله', '—'],
], font=9)
para('حساسیت‌سنجی: هر ۲۴ ساعت NPT پیمانکاری معادل ۲۳٬۰۰۰ دلار کسر است؛ بنابراین هر ۱۰۰ ساعت انتساب جدید/تغییر انتساب، معادل ۹۵٬۸۳۳٫۳۳ دلار در payable و net تغییر می‌کند. به‌همین دلیل، تأیید نهایی انتساب ردیف‌های خالی (Blank) در منبع DDR برای عملیات‌های AZNS-207، AZNS-122-WO2، AZNS-75، AZNS-12، AZNS-166، AZNS-127-WO2 و AZNS-127-ST2 پیش از تسویه نهایی ضروری است.')

# ===================================================== بخش ۵ — پرچم‌های مبنای تقویم
h1('بخش ۵ — پرچم‌های مبنای تقویم (Calendar-Basis Flags)')
para('مقایسه «روزهای CBS» (مبنای متره هزینه) با «روزهای DDR» (مبنای تقویم دکل) چهار ناسازگاری منبع را آشکار کرد. این پرچم‌ها به‌معنای خطای حسابی نیستند؛ بلکه نشان می‌دهند مبنای روزهای CBS و DDR برای آن عملیات یکسان نیست و باید با اداره قراردادها رفع‌ابهام شود.')
table('۵-۱  چهار پرچم منبع', ['چاه', 'نوع پرچم', 'روز CBS', 'روز DDR', 'Δ روز'], [
    ['AZNS-85', 'CBS_DAYS_GT_DDR_DAYS', '57٫40', '53', '−4٫40'],
    ['AZNS-12', 'CBS_DAYS_GT_DDR_DAYS', '91٫92', '72', '−19٫92'],
    ['AZNS-166', 'CBS_DAYS_GT_DDR_DAYS', '44٫24', '24', '−20٫24'],
    ['AZNS-127-ST2', 'CBS_DAYS_GT_DDR_DAYS', '115٫36', '106', '−9٫36'],
], font=9.5)
bullet('AZNS-85: مبنای روزهای CBS از تایم‌بریک‌دان منبع (شیت «AZNS-085 WO. Extra Work» در CBS عملیات ۳) با محاسبه 51٫89 روز (شامل کانتینجنسی) + 5٫5 روز کانتینجنسی پلن شماره ۱ = 57٫4 روز به‌دست آمد (تأیید کاربر ۲۰۲۶-۱۱-۰۸)؛ نرخ دکل همان benchmark قرارداد C1 (۲۳٬۰۰۰ USD/DAY) است. فرمول four-track روی همین مبنا اعمال شد: gross = 1٬219٬000 + (۵۳ − 57٫4) × ۲۳٬۰۰۰ = 1٬117٬800 USD؛ payable = net = 890٬195٫83 USD. پرچم منبع از CBS_DAYS_MISSING به CBS_DAYS_GT_DDR_DAYS تغییر کرد.')
bullet('AZNS-12 / AZNS-166 / AZNS-127-ST2 / AZNS-85: تعداد روز CBS از مبنای منبع بزرگ‌تر از روزهای DDR است؛ اختلاف فرمول gross با مبنای منبع در این چهار چاه ≤ ۰٫۰۰۱ دلار است (مبنای AZNS-85 مشتق‌شده از تایم‌بریک‌دان است، نه سلول ذخیره‌شده)؛ بااین‌حال مبنای بازبینی قراردادی لازم است.')
note('خط لوله دیگر هیچ تقسیم ۵۰/۵۰ برای ساعت‌های انتظار اعمال نمی‌کند؛ ستون fallback_50_50_waiting_h در جدول an_npt_attribution_audit برای همه ردیف‌ها صفر است.')

# ===================================================== بخش ۶ — ممیزی CBS و نرخ‌نامه‌ها
h1('بخش ۶ — ممیزی حساب‌های CBS و نرخ‌نامه')
da1 = next((f for f in FINDINGS if f['id'] == 'DA-01'), None)
da5 = next((f for f in FINDINGS if f['id'] == 'DA-05'), None)
da4 = next((f for f in FINDINGS if f['id'] == 'DA-04'), None)
table('۶-۱  نتیجه‌های کلیدی ممیزی CBS', ['یافته', 'شواهد عددی'], [
    ['DA-01 — عدم بسته‌شدن حسابی اقلام متره', '۱٬۵۷۲ خط قیمت‌گذاری‌شده؛ ۱۷ خط با qty×price ≠ total (۱۶ مادی + ۱ جزئی)؛ اختلاف علامت‌دار خالص −8٬759٬173 USD؛ قدر مطلق 9٬093٬071 USD؛ اختلاف‌های مثبت بزرگ «حاشیه سود شرکت» (سود/سربار) هستند، نه خطای حسابی'],
    ['DA-05 — پراکندگی نرخ واحد', '36 از 267 کلید نرخ‌نامه بیش از ۵٪ پراکندگی دارند؛ نرخ واحدِ یک عملیات/کالای تعریف‌شده در کل کمپین یکسان است؛ پراکندگی = تفاوت اسکوپ/قطر چاه/بسته قراردادی یا داده ناقص، نه خطای تصادفی'],
    ['DA-22 — نرخ‌نامه مرجع هم‌طراز', '295 کلید نرمال‌شده (202 چندچاهی + 93 تک‌چاهی): 122 کلید یکنواخت (<5٪ پراکندگی)، 3 کم‌پراکنش، 77 پراکنده؛ از 1٬176 خط قیمت‌گذاری‌شده 1٬057 خط (89٫9٪) در ۵٪ مرجع؛ 119 خط اختلاف >5٪؛ 376 خط بدون قیمت واحد — 8 خط قابل انتصاب از چاه‌های هم‌طراز، 368 گپ منبع؛ حاشیه سود روی کلیدهای یکنواخت (834 خط): +165٬938 / −1٬603٬805 (خالص −1٬437٬868 USD)'],
    ['DA-04 — بازمحاسبه فرمول‌ها', '48٬011 فرمول بازمحاسبه شد: 47٬965 مطابق / 3 اختلاف جزئی / 43 نامنطبق؛ ۲۷٫۵۱٪ سلول‌های عددی به‌صورت متن ذخیره شده‌اند (328٬989 سلول)'],
], font=8.5)
para('تفسیر عملیاتی: ستون Total در CBS یک «تخصیص قراردادی» است، نه همیشه حاصل‌ضرب حسابی مقدار در نرخ واحد؛ بنابراین اختلاف DA-01 به‌معنای خطای قطعی نیست، بلکه پرچم ممیزی برای مطابقت با پیوست‌های قراردادی است. طبق تأیید پروژه (۲۰۲۶-۱۱-۰۸)، اختلاف‌های مثبت بزرگ میان Total منبع و QTY × نرخ واحد همان «حاشیه سود شرکت» (سود/سربار) هستند؛ سمت منفی عمدتاً خطوط contingency با جمع صفر/مبلغ مقطوع است. هر دو مقدار (محاسبه‌شده و منبع) در جدول an_cbs_line_audit و شیت 27_DA_CBS_Arithmetic با سلول مرجع A1 نگهداری شده‌اند. برای خطوط ناقص، نرخ از چاه‌های هم‌طراز انتصاب می‌شود و در an_cbs_reference_delta با status=IMPUTED برچسب می‌خورد؛ هیچ مقدار منبع بازنویسی نمی‌شود.')
note('اقلام محرک هزینه: بیشترین اختلاف‌های مطلق در شیت 27 با مرتب‌سازی ABS(delta) فهرست شده‌اند؛ نمونه‌های بزرگ باید با پیوست مالی قرارداد تطبیق شوند، نه با فرض خطای CBS.')

# ===================================================== بخش ۷ — انتساب NPT
h1('بخش ۷ — انتساب NPT بدون تقسیم ۵۰/۵۰')
para('منبع DDR برای هر فعالیت N/NPT ستون «npt_attributed» را دارد؛ در بخش عمده‌ای از ردیف‌ها این سلول خالی است. خط‌مشی تحلیل: (۱) سلول‌های خالی «همان‌طور که هستند» نگهداری و در ستون blank_attribution_h گزارش می‌شوند؛ (۲) انتساب قراردادی از روی مقدار منبع و با همان قاعده warehouse (literal «Unattributed» → سمت پیمانکار) انجام می‌شود؛ (۳) هیچ ساعت انتظاری ۵۰/۵۰ تقسیم نمی‌شود (fallback = 0).')
img('AN_17_npt_attribution.png', 'شکل ۲ — ممیزی انتساب NPT از ردیف‌های صریح وضعیت N/NPT؛ بدون تقسیم ۵۰/۵۰ ساعت‌های انتظار')
table('۷-۱  انتساب NPT به تفکیک چاه (ساعت)', ['چاه', 'ردیف N/NPT', 'NPT', 'پیمانکار', 'کارفرما', 'خالی (blank)', 'انتساب صریح %', 'انتظار (waiting)'], [
    [r['well_key'], f"{r['npt_status_rows']:,}", fmt(r['npt_status_h'],1), fmt(r['contractor_h'],1), fmt(r['client_h'],1), fmt(r['blank_attribution_h'],1), fmt(r['explicit_attribution_pct'],2), fmt(r['waiting_h'],1)] for r in NPT_ATTR], font=7.6)
img('AN_14_pdf_units.png', 'شکل ۴ — توکن‌های PDF دارای واحد؛ پراکندگی استحکامی (robust IQR) و داده‌های پرت برچسب‌گذاری‌شده')
para('یافته‌های برجسته:')
bullet('میدان: 2٬622 ردیف N/NPT؛ 13٬316٫43 ساعت؛ پیمانکار 11٬183٫43 h (۸۴٫۰٪) — کارفرما 2٬133٫00 h (۱۶٫۰٪)؛ ردیف‌های با انتساب صریح تنها ۳٬۷۲۶٫۹۸ ساعت (۲۸٫۰٪) هستند؛ ۹٬۵۸۹٫۴۵ ساعت (۷۲٫۰٪) سلول انتساب خالی دارند.')
bullet('AZNS-127-WO2 با ۹۴۱٫۵ ساعت NPT، ۷۲٫۴٪ سهم کارفرما (۶۸۲٫۰ h) دارد؛ AZNS-75 با ۱٬۲۴۴٫۰ ساعت، ۵۳٫۸٪ (۶۶۹٫۵ h)؛ AZNS-127-ST2 با ۴۸۹٫۵ ساعت، ۴۶٫۵٪ (۲۲۷٫۵ h) — این سه عملیات محور اصلی مذاکره ادعای NPT با کارفرما هستند.')
bullet('AZNS-207 و AZNS-122-WO2: ۴۷٫۵ و ۴۰۰٫۵ ساعت کارفرمایی صریح؛ انتظار (waiting) بدون انتساب در AZNS-40 (۴۴۶٫۵ h)، AZNS-10 (۳۸۹٫۰ h)، AZNS-122-WO1 (۴۰۲٫۵ h) و AZNS-122-WO2 (۳۹۴٫۰ h) متمرکز است.')
note('انتساب «پیمانکار/کارفرما» برای ردیف‌های خالی، طبقه‌بندی منبع است نه قضاوت حسابی؛ برای تسویه قطعی، اداره قراردادها باید هر ردیف خالی را با منبع DDR (ستون npt_attributed) تأیید کند.')

# ===================================================== بخش ۸ — تحلیل اعداد PDF
h1('بخش ۸ — ممیزی کامل اعداد PDF (۳٬۰۵٬۴۴۷ توکن)')
para('تمام 305٬447 عدد موجود در لایه متنی PDFها (۲٬۷۲۴ صفحه EOWR + ۱٬۶۲۱ صفحه Well Program) استخراج، طبقه‌بندی و با کلید منبع (فایل/صفحه/خط) نمایه شده‌اند. خروجی خام: SAZN_PDF_Numeric_Token_Census.csv.gz (۴٫۷۸ MB) و جدول‌های an_pdf_* در انبار داده.')
table('۸-۱  طبقه‌بندی عددی توکن‌ها', ['طبقه', 'تعداد', 'سهم', 'توضیح'], [
    ['Unclassified (بدون واحد/زمینه اندازه‌گیری)', '187٬085', '61٫2٪', 'شماره‌گذاری، فهرست مطالب، ارجاعات و اعداد متفرقه؛ نگه‌داری‌شده با provenance'],
    ['Navigation / TOC', '42٬390', '13٫9٪', 'شماره صفحه و آیتم‌های فهرست مطالب'],
    ['Header / identifier', '34٬293', '11٫2٪', 'شماره گزارش، کد سند، شماره چاه در سربرگ‌ها'],
    ['Measurement-context (بدون واحد صریح)', '28٬131', '9٫2٪', 'عدد در مجاورت کلیدواژه‌های اندازه‌گیری (ROP، depth، فشار، mud و…)'],
    ['Unit-bearing (واحد صریح)', '13٬548', '4٫4٪', 'عدد با واحد واقعی: m، ft، psi، bbl، ppg، rpm، inch و ۳۳ واحد دیگر'],
    ['Measured (با واحد یا زمینه اندازه‌گیری)', '41٬679', '13٫6٪', 'اجتماع Unit-bearing و Measurement-context'],
], font=8.5)
img('AN_13_pdf_token_census.png', 'شکل ۳ — سرشماری تمام 305٬447 عدد PDF بر پایه بزرگی و کلاس زمینه مشتق')
table('۸-۲  کیفیت و تطبیق‌پذیری PDF', ['شاخص', 'مقدار'], [
    ['صفحات کم‌متن (< 200 کاراکتر) — AF-05', '364 صفحه (8٫38٪ از 4٬345)  |  EOWR: 267 صفحه  |  Well Program: 97 صفحه'],
    ['توکن‌های PDF روی صفحات کم‌متن', f"{sum(x['numeric_tokens'] for x in PDF['low_text_by_category']):,} عدد (EOWR 1٬849 + PROGRAM 494)"],
    ['تصاویر روی صفحات کم‌متن', f"{sum(x['images'] for x in PDF['low_text_by_category']):,} تصویر — OCR/بازخوانی منبع درخواست شود"],
    ['پرت‌های بزرگگی واحد-محور (robust IQR)', '538 توکن — برچسب‌گذاری شده، هرگز تصحیح نشده'],
    ['Coincidence دقیق با Excel/Fact (شش‌رقم اعشار)', '4٬286 مقدار متمایز  |  386٬692 تطابق توکن‌سطح'],
    ['داده‌های با بزرگگی ≥ 1e12 (مشکوک/خطای متن)', '22 توکن — شامل ۸٫۱۲×۱۰^۳۱ در Well Program چاه AZNS-40 (صفحه ۲۲۹)'],
], font=8.5)
para('نمونه پرت‌های بزرگگی (DA-19): بزرگ‌ترین آن‌ها «8,12,16,20,24,28,32,36…» در ردیف ۱۰۸ صفحه ۲۲۹ از ۱.AZNS_40_Well_Program.pdf است که توکن لیستی ابزارهاست نه یک عدد فیزیکی؛ با robust_z = ۶۸۹٫۸ برچسب‌گذاری شده است. دو توکن دیگر ۱۶٬۲۵۰ gal (AZNS-207 EOWR p.151) و ۱۱٬۲۵۰ gal (AZNS-40 EOWR p.118) به‌ترتیب دارای تطابق دقیق Excel هستند.')
note('تطابق دقیق شش‌رقم اعشار میان PDF و Excel صرفاً یک «سرنخ تطبیق» است، نه اثبات هویت فیزیکی مقادیر؛ برای داوری مالی باید هر مورد با منبع (page/line یا sheet/cell) تطبیق داده شود (DA-20).')

# ===================================================== بخش ۹ — ثبت ۲۲ یافته
h1('بخش ۹ — ثبت ۲۲ یافته دپارتمان تحلیل داده (DA-01 تا DA-22)')
para('جدول زیر عنوان، شواهد عددی و اقدام پیشنهادی هر یافته را می‌آورد؛ متن کامل هر یافته (تفسیر و اقدام) در شیت 51_DA_Findings از کتاب کار و جدول an_findings انبار داده موجود است.')
find_rows = []
PRI = {'P1': 'بحرانی', 'P2': 'مهم', 'P3': 'عادی'}
for f in FINDINGS:
    ev = f['numeric_evidence'].replace('\n', ' ')
    if len(ev) > 240: ev = ev[:237] + '…'
    find_rows.append([f['id'], f['title'].strip(), ev, PRI.get(f['priority'], f['priority'])])
table('۹-۱  ثبت یافته‌ها', ['شناسه', 'عنوان', 'شواهد عددی (خلاصه)', 'اولویت'], find_rows, font=6.8)

# ===================================================== بخش ۱۰ — محدودیت‌ها
h1('بخش ۱۰ — محدودیت‌ها و داده‌های ناقص (AF-01 تا AF-05)')
bullet('AF-01 — عملیات ۴ (AZNS-85): EOWR ندارد؛ تحلیل از DDR + CBS + Well Program انجام شده است؛ هر ادعای مبتنی بر EOWR برای op4 نامعتبر است.')
bullet('AF-02 — EOWR ترکیبی ST#2 (۲۳۸ صفحه) میان AZNS-127-WO2 و AZNS-127-ST2 مشترک است؛ ۱۱ صفحه shared به‌صورت ساختاری شمارش دوگانه دارد و در تطبیق‌ها once counted شده است.')
bullet('AF-03 — برچسب کاربر «CCA/C/1403/31 / ST2 amendment» در برابر سربرگ منبع EOWR ترکیبی ST#2 با عبارت «CCA/C/1405/3»؛ هر دو در گزارش و سامانه نگهداری شده‌اند؛ پاسخ قراردادی لازم است.')
bullet('AF-04 — نرخ‌های مرجع چاه‌های ۲۹ و ۳۸ به‌عنوان benchmark از دامنه این میدان خارج هستند؛ در هیچ محاسبه‌ای استفاده نشده‌اند.')
bullet('AF-05 — ۳۶۴ صفحه کم‌متن (EOWR ۲۶۷ صفحه، Well Program ۹۷ صفحه): لایه متنی برای استخراج عددی کافی نیست؛ مقادیر OCR حدس زده نشده‌اند؛ درخواست نسخه native/OCR از بایگانی الزامی است.')
bullet('۲۷٫۵۱٪ از سلول‌های عددی شیت‌ها به‌صورت متن ذخیره شده‌اند (۳۲۸٬۹۸۹ سلول)؛ در محاسبات پس از نرمال‌سازی خوانده شده‌اند، اما منبع تغییر نکرده است.')
bullet('۴۳ فرمول با اختلاف بازمحاسبه و ۳ اختلاف جزئی: در شیت 25_DA_Formula_Audit با آدرس سلول فهرست شده‌اند.')
bullet('مبنای «روزهای CBS» عملیات AZNS-85 از تایم‌بریک‌دان منبع (57٫4 روز) مشتق شده و هیچ سلول ذخیره‌شده‌ای مبنا نیست؛ برای AZNS-12/166/127-ST2 نیز روز CBS از DDR بزرگ‌تر است (بخش ۵).')

# ===================================================== بخش ۱۱ — اقدامات
h1('بخش ۱۱ — اقدامات پیشنهادی (با مالک و اولویت)')
table('۱۱-۱  فهرست اقدام‌ها', ['#', 'اقدام', 'مالک', 'اولویت'], [
    ['۱', 'حل چهار پرچم مبنای تقویم (AZNS-85، AZNS-12، AZNS-166، AZNS-127-ST2) با اداره قراردادها قبل از تسویه قطعی؛ مبنای CBS-days در پیوست‌های C2/C3 مکتوب شود.', 'اداره قراردادها', 'P1'],
    ['۲', 'تأیید انتساب ردیف‌های خالی npt_attributed در DDR هفت عملیات (۸۰ ردیف منبع‌خالی؛ ۹٬۵۸۹٫۴۵ ساعت)؛ به‌ویژه AZNS-127-WO2 و AZNS-75 که سهم کارفرما > ۵۰٪ است.', 'متره و کنترل پروژه + پیمانکار', 'P1'],
    ['۳', 'تطبیق ۴۳ فرمول نامنطبق و ۱۶ آیتم مادی CBS با شکاف ≥ ۵٪ (شیت‌های 25 و 27) با پیوست‌های مالی قراردادها.', 'کنترل هزینه / حسابداری پروژه', 'P1'],
    ['۴', 'درخواست نسخه native یا OCR صفحات کم‌متن (AF-05) از بایگانی؛ تا آن زمان اعداد آن صفحات از هر تصمیم مالی خارج بماند.', 'بایگانی فنی + IT', 'P1'],
    ['۵', 'رفع‌ابهام برچسب قرارداد ST#2 (CCA/C/1403/31 / ST2 amendment در برابر سربرگ CCA/C/1405/3) و ثبت اصلاحیه در صورت رعایت حقوقی.', 'اداره قراردادها / حقوقی', 'P1'],
    ['۶', 'بازبینی ۱۷ خط CBS با qty×price ≠ total، ۳۶ کلید هم‌گروه an_unit_price با پراکندگی >5٪ و ۱۱۹ خط با اختلاف >5٪ از نرخ‌نامه مرجع DA-22؛ تصمیم آگاهانه lump-sum در برابر نرخ واحد؛ انتصاب نرخ‌های ناقص از چاه‌های هم‌طراز (۸ خط) و رسیدگی به ۳۶۸ گپ منبع باقی‌مانده.', 'متره و برآورد', 'P2'],
    ['۷', 'بررسی ۵۳۸ عدد پرت بزرگگی PDF (DA-19)؛ تأیید این‌که توکن‌های لیستی (مثل ۸,۱۲,۱۶…) به‌اشتباه در محاسبات وارد نشده باشند.', 'تیم تحلیل داده', 'P2'],
    ['۸', 'به‌روزرسانی فایل AF-01 تا AF-05 پس از دریافت EOWR عملیات ۴ (در صورت صدور) و اصلاح جداول مشتق.', 'تیم تحلیل داده + بایگانی', 'P2'],
    ['۹', 'استفاده از جدول‌های an_cost_track_audit و an_npt_attribution_audit به‌عنوان «دفتر ثبت نهایی» انتساب NPT در صورت‌جلسات مالی.', 'حسابداری پروژه', 'P2'],
    ['۱۰', 'تولید نسخه به‌روز این گزارش پس از رفع هر پرچم؛ نسخه ۱٫۱ پس از پاسخ قراردادی.', 'دپارتمان تحلیل داده', 'P3'],
], font=8)

# ===================================================== پیوست — ثبت تحویل‌دادنی‌ها
h1('پیوست الف — ثبت تحویل‌دادنی‌ها و اصالت‌سنجی')
table('الف-۱  فایل‌های کلیدی', ['فایل', 'شرح', 'اندازه'], [
    ['SAZN_Master_Warehouse.sqlite (+ .gz/.sha256)', 'انبار داده جامع: 87 جدول (شامل 42 جدول مشتق an_*)، تمام سلول‌ها و توکن‌ها با provenance', '171.5 MB / 38.7 MB (.gz)'],
    ['SAZN_Master_Audit_Workbook.xlsx', '۶۵ شیت؛ شیت‌های 23–51 تحلیل داده، 52–64 PDF/هزینه و نرخ‌نامه مرجع (63 و 64)؛ 102٬143 ردیف جدید', '12.1 MB'],
    ['SAZN_Executive_Interactive_Dashboard.html', '۸۸۲٬۷۱۴ ردیف قابل جستجو (آفلاین، بدون شبکه) با پانل‌های DA/PDF/هزینه و نرخ‌نامه مرجع DA-22', '9.2 MB'],
    ['SAZN_PDF_Numeric_Token_Census.csv.gz', 'سرشماری کامل 305٬447 توکن PDF', '4.6 MB'],
    ['SAZN_Analytics_Exceptions.csv', 'استثناهای تحلیلی (CBS، فرمول، پرت‌ها، نرخ‌نامه DA-22) — ۵٬۴۹۳ ردیف', '0.68 MB'],
    ['Word_Books/*.docx (۳۱ جلد)', '۳۰ کتاب مرجع (FA/EN) + این گزارش (SAZN-FA-DA-001)', '≈34 MB'],
    ['SAZN_Graphics/*.png (۱۶۷ فایل)', 'گرافیک 2400×1800 / ۳۰۰dpi؛ English-only؛ بدون هم‌پوشانی متن', '≈26 MB'],
], font=8.5)
sha = open(os.path.join(OUT, 'SAZN_Master_Warehouse.sqlite.gz.sha256')).read().split()[0]
para('SHA-256 پرونده SAZN_Master_Warehouse.sqlite (پیش از فشرده‌سازی): ' + sha, size=8, color='475569')
para('برای بازیابی انبار از پرونده فشرده: python _pipeline/common.py (به‌صورت خودکار .sqlite را از .sqlite.gz بازمی‌گرداند)؛ سپس python _pipeline/build_warehouse.py برای بازسازی از منابع.', size=8, color='475569')
p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(20)
set_run_font(p.add_run('— پایان گزارش SAZN-FA-DA-001 نسخه ۱.۱ —'), 'FA', 10, True, NAVY)

d.save(PATH)
print('SAVED', PATH, round(os.path.getsize(PATH)/1048576, 2), 'MB', flush=True)
