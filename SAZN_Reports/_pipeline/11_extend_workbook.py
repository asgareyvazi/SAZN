"""Append the Drilling-Data Analysis Department sheets to the audit workbook."""
import os, sys, sqlite3, json
sys.path.insert(0, os.path.dirname(__file__))
from common import OUT, DB
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

XLSX = os.path.join(OUT, 'SAZN_Master_Audit_Workbook.xlsx')
CON = sqlite3.connect(DB); CON.row_factory = sqlite3.Row
wb = openpyxl.load_workbook(XLSX)
NAVY, LIGHT, WHITE = '0B1F3A', 'EEF2F7', 'FFFFFF'

SHEETS = [
 ('23_DA_Numeric_Census', 'an_num_census', ['Well','Document','Sheet','Category','N','Min','Q1','Median','Q3','Max','Mean','Std','Sum','Zeros','Negatives','Integers','IQR outliers','Band'],
  'SELECT well_key,rel_path,sheet,category,n,v_min,q1,v_med,q3,v_max,v_mean,v_std,v_sum,zeros,negatives,integers,outliers,dominant_band FROM an_num_census ORDER BY n DESC'),
 ('24_DA_Numeric_Histogram','an_num_hist',['Magnitude band','Count','Percent','Example value','Example source'],
  'SELECT band,n,pct,example,example_source FROM an_num_hist'),
 ('25_DA_Formula_Audit','an_formula_audit',['Well','Document','Sheet','Cell','Formula','Refs','Cached','Recomputed','Delta','Kind','Status'],
  'SELECT * FROM an_formula_audit WHERE status<>"MATCH" UNION ALL SELECT * FROM (SELECT * FROM an_formula_audit WHERE status="MATCH" LIMIT 400)'),
 ('26_DA_Cell_Quality','an_cell_quality',['Document','Sheet','Numeric as text','Formula cells','Formulas with refs','Hardcoded formulas','Cross-sheet','Blanks','Total cells'],
  'SELECT * FROM an_cell_quality'),
 ('27_DA_CBS_Arithmetic','an_cbs_line_audit',['Well','Document','Sheet','Cell','Description','Unit','Qty','Unit price','Total','Expected','Delta','Delta %','Status'],
  'SELECT well_key,rel_path,sheet,cell,description,unit,qty,unit_price,total,expected,delta,pct,status FROM an_cbs_line_audit ORDER BY ABS(delta) DESC'),
 ('28_DA_Unit_Price','an_unit_price',['Normalised description','Unit','N','Min','Median','Max','Spread %','Wells','Dispersion'],
  'SELECT * FROM an_unit_price ORDER BY spread_pct DESC'),
 ('29_DA_Time_Balance','an_time_balance',['Well','Report','Date','Canonical sheet','Activity h','PT h','NPT h','Unclassified h','Gap h','Carried-over h','Status','Time-code h','Max row h','Rows','Status rows'],
  'SELECT well_key,report_no,date_g,canonical_sheet,activity_h,pt_h,npt_h,unclassified_h,ilt_h,carried_over_h,balance_status,time_code_h,max_row_h,n_rows,status_rows FROM an_time_balance ORDER BY well_key,report_no'),
 ('30_DA_Daily_Consistency','an_daily_consistency',['Well','Report','MD 00:00','MD 06:00','MD 24:00','Drilled 24h','Expected delta','Delta','Status','Corrected 24h','ROP','ROP recomputed','ROP delta'],
  'SELECT * FROM an_daily_consistency ORDER BY well_key,report_no'),
 ('31_DA_Activity_Crosstab','an_activity_crosstab',['Well','Phase','Code','Status','Attributed','Hours','Events','Share %'],
  'SELECT well_key,phase,code,status,attributed,hours,events,share_pct FROM an_activity_crosstab ORDER BY well_key,hours DESC'),
 ('32_DA_ILT','an_ilt',['Well','Report','Connection h','Trip h','Circulate h','Wait h','Survey h','Ream h','Other h','Gap h','Gap %'],
  'SELECT well_key,report_no,connection_h,trip_h,circulate_h,wait_h,survey_h,ream_h,other_h,ilt_h,ilt_pct FROM an_ilt ORDER BY well_key,report_no'),
 ('33_DA_Daily_Parameters','an_daily_params',['Well','Report','Parameter','Value','Unit','Sheet','Cell','Label'],
  'SELECT well_key,report_no,param,value,unit,sheet,cell,label FROM an_daily_params ORDER BY well_key,report_no,param'),
 ('34_DA_Param_Stats','an_param_stats',['Well','Parameter','N','Min','Q1','Median','Q3','Max','Mean','Std','CV %','IQR outliers'],
  'SELECT well_key,param,n,v_min,q1,v_med,q3,v_max,v_mean,v_std,cv_pct,iqr_outliers FROM an_param_stats ORDER BY well_key,param'),
 ('35_DA_Correlations','an_correlations',['Well','X','Y','N','Pearson','Spearman','Slope','Intercept','R²','Strength'],
  'SELECT well_key,x,y,n,pearson,spearman,slope,intercept,r2,strength FROM an_correlations ORDER BY ABS(pearson) DESC'),
 ('36_DA_NPT_Forensics','an_npt_forensics',['Well','Rank','Report','Category','Hours','Cost USD','Attributed','Depth','Formation','Days since prev','Cumulative h','Cumulative %','MTBF h','Payer'],
  'SELECT well_key,rank,report_no,category,hours,cost_usd,attributed,depth_m,formation,days_since_prev,cumulative_h,cumulative_pct,mtbf_h,payer FROM an_npt_forensics ORDER BY well_key,rank'),
 ('37_DA_NPT_Pareto','an_npt_pareto',['Well','Category','Hours','Share %','Cumulative %','Rank','Cost USD'],
  'SELECT well_key,category,hours,share_pct,cumulative_pct,rank,cost_usd FROM an_npt_pareto ORDER BY well_key,rank'),
 ('38_DA_Formation_Stats','an_formation_stats',['Well','Formation','N tops','MD med','TVD med','Prognosis med','Delta m','Thickness m','Lithology','Problem h','MW min','MW max','Loss events','Structural bias'],
  'SELECT well_key,formation,n_tops,md_med,tvd_med,prognosis_med,delta_m,thickness_m,lithology,problem_h,mw_min,mw_max,loss_events,structural_bias FROM an_formation_stats ORDER BY well_key,md_med'),
 ('39_DA_Directional','an_directional',['Well','Report','MD','Incl','Azi','TVD','DLS','Cum dogleg','Tortuosity','Section'],
  'SELECT well_key,report_no,md,incl,azi,tvd,dls,cum_dogleg,tortuosity,section FROM an_directional ORDER BY well_key,md'),
 ('40_DA_Mud_Balance','an_mud_balance',['Well','Report','Material','Received','Used','Stock prev','Stock','Implied received','Delta','Status'],
  'SELECT well_key,report_no,material,received,used,stock_prev,stock,implied_received,delta,status FROM an_mud_balance WHERE status="IMBALANCE"'),
 ('41_DA_Casing_Checks','an_casing_check',['Well','Size','Grade','From m','To m','Burst','Collapse','Nearest formation','Nearest top m','Shoe-to-top m','Lap overlap m','Status'],
  'SELECT well_key,size,grade,from_m,to_m,burst,collapse,nearest_formation,nearest_top_m,shoe_to_top_m,lap_overlap_m,status FROM an_casing_check ORDER BY well_key,to_m'),
 ('42_DA_Pressure_Stats','an_pressure_stats',['Well','Unit','N','Min','Median','Max','Mean','Std','Outliers'],
  'SELECT well_key,unit,n,v_min,v_med,v_max,v_mean,v_std,outliers FROM an_pressure_stats ORDER BY well_key'),
 ('43_DA_Stimulation_Audit','an_stim_audit',['Well','Document','Sheet','Item','Unit','Planned volume','Unit price','Total USD','Expected','Delta','Family','Status'],
  'SELECT well_key,rel_path,sheet,item,unit,planned_volume,unit_price,total_usd,expected,delta,family,status FROM an_stim_audit ORDER BY ABS(delta) DESC'),
 ('44_DA_Reconciliation_Delta','an_recon_delta',['Well','Item','CBS USD','DDR qty','EOWR value','Delta abs','Delta %','Status','Basis'],
  'SELECT well_key,item,cbs_usd,ddr_qty,eowr_value,delta_abs,delta_pct,tolerance_status,basis FROM an_recon_delta ORDER BY well_key,ABS(delta_abs) DESC'),
 ('45_DA_Program_vs_Actual','an_program_vs_actual',['Well','Formation','Program MD','EOWR MD','DDR MD','Delta P→A','Delta P→D','Bias','Program source','EOWR source'],
  'SELECT well_key,formation,program_md,eowr_md,ddr_md,delta_pa,delta_pd,bias,source_program,source_eowr FROM an_program_vs_actual ORDER BY well_key,program_md'),
 ('46_DA_Cost_Efficiency','an_cost_efficiency',['Well','Days','Drilled m','CBS','Gross','Payable','Net','NPT h','USD/day','USD/m','NPT share %','Payable ratio %','Cost rank','NPT rank','Efficiency index'],
  'SELECT well_key,days,drilled_m,cbs,gross,payable,net,npt_h,usd_per_day,usd_per_m,npt_share_pct,payable_ratio_pct,cost_rank,npt_rank,efficiency_index FROM an_cost_efficiency ORDER BY cost_rank'),
 ('47_DA_Learning_Metrics','an_learning_metrics',['Chain','From','To','Δ days','Δ NPT h','Δ CBS','Δ ROP','NPT/day from','NPT/day to','CBS/day from','CBS/day to','Direction'],
  'SELECT chain,from_well,to_well,delta_days,delta_npt_h,delta_cbs,delta_rop,npt_per_day_from,npt_per_day_to,cbs_per_day_from,cbs_per_day_to,direction FROM an_learning_metrics'),
 ('48_DA_Field_Correlations','an_field_correlations',['X','Y','N','Pearson','Spearman','Slope','R²','Strength'],
  'SELECT x,y,n,pearson,spearman,slope,r2,strength FROM an_field_correlations ORDER BY ABS(pearson) DESC'),
 ('49_DA_Data_Quality','an_data_quality',['Check','Scope','N','Affected','Percent','Detail'],
  'SELECT check_name,scope,n,affected,pct,detail FROM an_data_quality'),
 ('50_DA_Outliers','an_outliers',['Well','Document','Sheet','Cell','Value','Z-score','Sheet mean','Sheet std'],
  'SELECT well_key,rel_path,sheet,cell,value,z,sheet_mean,sheet_std FROM an_outliers ORDER BY ABS(z) DESC'),
 ('51_DA_Findings','an_findings',['ID','Category','Title','Numeric evidence','Interpretation','Action','Priority'],
  'SELECT id,category,title,numeric_evidence,interpretation,action,priority FROM an_findings'),
 ('52_PDF_Token_Sample','an_pdf_tokens',['Well','Document','Category','Page','Line','Token','Value','Unit','Context','Chapter','Band','Context class','Value kind','Measured','Navigation','Identifier','Low text','Excel match','Fact match','Robust outlier','Robust z','Flags'],
  'SELECT well_key,rel_path,category,page_no,line_no,token,value,unit,substr(context,1,120),chapter,band,context_class,value_kind,is_measured,is_navigation,is_identifier,is_low_text,excel_exact_match,fact_exact_match,robust_outlier,robust_z,flags FROM an_pdf_tokens WHERE robust_outlier=1 OR is_measured=1 ORDER BY robust_outlier DESC,page_no LIMIT 40000'),
 ('53_PDF_Token_Stats','an_pdf_token_stats',['Category','Well','Unit','Band','N','Sum','Min','Median','Max','Mean','Std','Zeros','Negatives','Integers','Decimals','Measured','Navigation','Identifier','Low text','Outliers','Excel matches','Fact matches','Unit-bearing','Unclassified'],
  'SELECT * FROM an_pdf_token_stats ORDER BY n DESC'),
 ('54_PDF_Unit_Stats','an_pdf_unit_stats',['Unit','N','Share %','Min','Q1','Median','Q3','Max','Mean','Std','Outliers','Measured','Excel matches','Fact matches','Context classes','Dispersion'],
  'SELECT * FROM an_pdf_unit_stats ORDER BY n DESC'),
 ('55_PDF_Context_Stats','an_pdf_context_stats',['Context class','N','Share %','Min','Median','Max','Mean','Outliers','Excel matches','Fact matches','Low text','Unit-bearing'],
  'SELECT * FROM an_pdf_context_stats ORDER BY n DESC'),
 ('56_PDF_Page_Stats','an_pdf_page_stats',['Well','Document','Category','Page','Chapter','Chars','Images','Tokens','Density/1k chars','Distinct units','Unit-bearing','Measured','Navigation','Identifier','Outliers','Excel matches','Fact matches','Min','Median','Max','Status'],
  'SELECT * FROM an_pdf_page_stats ORDER BY rel_path,page_no'),
 ('57_PDF_Low_Text','an_pdf_low_text',['Well','Document','Page','Chapter','Chars','Images','Tokens','Distinct units','Measured tokens','Outlier tokens','Note'],
  'SELECT * FROM an_pdf_low_text ORDER BY n_chars,rel_path,page_no'),
 ('58_PDF_Outliers','an_pdf_outliers',['Well','Document','Category','Page','Line','Token','Value','Unit','Context','Band','Context class','log10|v|','Robust z','Excel match','Fact match','Flags'],
  'SELECT well_key,rel_path,category,page_no,line_no,token,value,unit,substr(context,1,120),band,context_class,log_abs,robust_z,excel_exact_match,fact_exact_match,flags FROM an_pdf_outliers ORDER BY ABS(robust_z) DESC'),
 ('59_PDF_Value_Matches','an_pdf_value_matches',['Value key','Value','PDF tokens','Excel cells','Fact rows','First PDF source','First Excel source','First fact source','Match kind','Caveat'],
  'SELECT * FROM an_pdf_value_matches ORDER BY pdf_tokens DESC'),
 ('60_PDF_Source_Gaps','an_pdf_source_gaps',['Well','Document','Category','Pages','Low-text pages','Low-text %','Tokens','Chars','Images','Measured','Unclassified','Outliers','Status'],
  'SELECT * FROM an_pdf_source_gaps ORDER BY low_text_pages DESC,numeric_tokens DESC'),
 ('61_Cost_Track_Audit','an_cost_track_audit',['Well','Op','Rate basis','Rig rate','CBS plan','CBS days','DDR days','Δ days','Gross (formula)','Gross (stored)','Gross Δ','NPT total h','Contractor h','Client h','Contractor deduction','Client deduction','All-NPT deduction','Payable (formula)','Payable (stored)','Payable Δ','Net (formula)','Net (stored)','Net Δ','Attribution gap h','Status','Calendar flag','Fallback 50/50 h'],
  'SELECT * FROM an_cost_track_audit ORDER BY op'),
 ('62_NPT_Attribution_Audit','an_npt_attribution_audit',['Well','Activity rows','NPT status rows','NPT status h','Contractor h','Client h','Blank attribution h','Explicit attribution h','Explicit contractor h','Explicit client h','Profile NPT total h','Delta h','Explicit attribution %','Explicit client %','Waiting h','Waiting blank h','Fallback 50/50 h','Status'],
  'SELECT * FROM an_npt_attribution_audit ORDER BY well_key'),
]

def add(name, headers, sql, cap=40000):
    if name in wb.sheetnames: del wb[name]
    ws = wb.create_sheet(name[:31]); ws.freeze_panes = 'A2'
    note = ('SAZN Drilling-Data Analysis Department — derived forensic table. '
            'Every row is computed from source values; corrections are stored separately and labelled.')
    c = ws.cell(1, 1, note); c.font = Font(name='Calibri', size=10, italic=True, color='475569')
    for i, h in enumerate(headers, 1):
        c = ws.cell(2, i, h); c.font = Font(name='Calibri', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY); c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = min(46, max(13, len(h) + 3))
    rr = 3
    for row in CON.execute(sql):
        for i, v in enumerate(row, 1):
            if isinstance(v, str) and len(v) > 220: v = v[:217] + '…'
            c = ws.cell(rr, i, v); c.font = Font(name='Calibri', size=9, color='0F172A')
            if (rr-3) % 2: c.fill = PatternFill('solid', fgColor=LIGHT)
            if isinstance(v, float): c.number_format = '#,##0.0000'
        rr += 1
        if rr - 3 >= cap: break
    return rr - 3

total = 0
for name, _, headers, sql in SHEETS:
    n = add(name, headers, sql); total += n; print('  SHEET', name, n, flush=True)
wb.save(XLSX)
print('EXTENDED WORKBOOK sheets', len(wb.sheetnames), 'new rows', total, 'MB', round(os.path.getsize(XLSX)/1e6, 2))
