# SAZN — South Azadegan Reference Programme

Complete, zero-summarisation deliverable built from the 51 source documents in `CBS/`, `DDRs/`, `04.EOWR/` and `Well Program/`. The four original source folders are never modified.

## Deliverables

| File / folder | Content |
|---|---|
| `SAZN_Master_Warehouse.sqlite.gz` + `.sha256` | Gzip of the full warehouse (36.9 MB, expands to ~163 MB). `common.load_json()`/pipeline restore the `.sqlite` automatically from the `.gz` when absent. Tracked in Git; the raw `.sqlite` is gitignored (over GitHub's file-size limit) and is restored locally or rebuilt from sources. |
| `SAZN_Master_Warehouse.sqlite` | Full archive (local only, not tracked): 2,810 exact worksheet cell blobs, 324,150 numeric cells, 155,248 formulas, 4,345 lossless PDF page texts, **305,447 numeric PDF tokens**, 1,349 numbered tables, FTS indexes, semantic DDR/CBS tables, derived KPI/events/lessons/recommendations/reconciliation/NPT/time-series, **40 `an_*` deep-analytics tables**. Every value carries `file → sheet/page → cell/row/table` provenance. |
| `SAZN_Master_Audit_Workbook.xlsx` | **63 sheets**: source registry, four-track cost, CBS tariff & materials, DDR 24-h log, NPT, formations, BHA, casing, surveys, EOWR facts, PDF tables, lessons, REC cards, audit findings, full cell/page inventory, field KPI, **29 DA sheets (23–51)**, **11 PDF/cost forensic sheets (52–62)**. ~100,000 DA/PDF/cost audit rows. |
| `SAZN_Executive_Interactive_Dashboard.html` | Self-contained, offline, **881,121 searchable rows**. Live search across numbers, wells, rigs, materials, files, sheets, pages, **all 305,447 PDF numeric tokens** and every DA/PDF/cost forensic row, with exact provenance; four-track charts, contract close-out, deep-analytics panels (DQ scorecard, DA findings, PDF census, NPT attribution, four-track cost bridge). |
| `Word_Books/` | **31 DOCX**: 13 single-well FA + 13 single-well EN, plus 2 field-master FA/EN, 2 executive-management FA/EN, and **`00_SAZN_Analytical_Findings_Technical_Financial_FA.docx`** (گزارش فنی–مالی یافته‌های تحلیل داده‌های حفاری، کد SAZN-FA-DA-001). Each reference book has **24 chapters, 149 provenance tables and 11 custom graphics**. FA = B Nazanin + full RTL (`w:bidi`, `w:bidiVisual`); EN = Times New Roman. |
| `SAZN_Graphics/` | **166 PNG** (2400×1800 px / 300 DPI) rendered from HTML5 + CSS3 + SVG + Canvas in headless Chromium. 0 text-overlap pairs, English-only text, inverted depth axis, mean ≈160 KB. Includes 17 `AN_*` deep-analytics artboards. |
| `SAZN_PDF_Numeric_Token_Census.csv.gz` | Census of **all 305,447 PDF numeric tokens** with page/line provenance, magnitude band, derived context class, robust unit-group outlier flag, Excel/fact coincidence flags. 4.6 MB (gzip). |
| `SAZN_PDF_Analytics_Exceptions.csv.gz` | Exception register (PDF + cost forensics). |
| `SAZN_Analytics_Exceptions.csv` | Deep-analytics exceptions (formula mismatches, CBS arithmetic, time balance, mud imbalance, outliers). 0.6 MB. |
| `SAZN_DOCX_Audit.json` | Gate + full audit of all 30 Word books. `SAZN_Final_Audit.json` | Final gate for the whole patchset (patchset files, MB, size cap, DB/workbook/dashboard/graphics/DOCX checks). |
| `_data/` | `well_profiles.json.gz`, `field_analytics.json`, `deep_analytics.json`, `pdf_cost_forensics.json`, `visual_review.json`, `analytics_overlap.json` (0 overlaps). |
| `_pipeline/` | Reproducible generators: `common.py`, `build_warehouse.py`, `06_graphics.py`…`09_dashboard.py`, `10_deep_analytics.py`, `11_extend_workbook.py`, `12_analytics_graphics.py`, `13_pdf_and_cost_forensics.py`, `14_findings_fa.py`. |

## Source census

13 CBS workbooks / 202 sheets · 13 DDR packs / 2,608 sheets · 11 EOWR / 2,724 pages · 14 Well Programmes / 1,621 pages = **51 files**. 1,196,084 non-empty Excel cells, 324,150 numeric cells, 155,248 formulas, 4,345 PDF pages, **305,447 numeric PDF tokens** (13,548 unit-bearing; 41,679 measured / context-bearing), 1,349 numbered PDF tables.

## Four-track time / cost accounting

1. **CBS Plan** = sourced cost ceiling (`Total Cost-EPD & Rigless`, EPD + rig-less, or `Cost Breakdown` grand total).
2. **Gross Calendar** = CBS Plan + (DDR days − CBS planned days) × sourced rig day rate.
3. **Payable excl. contractor zero-rate NPT** = Gross − contractor NPT hours / 24 × rate.
4. **Pure Net excl. all NPT** = Gross − all attributed NPT hours / 24 × rate.

Field totals: CBS plan **94,797,894.39 USD** · gross **99,843,688.22 USD** · payable **89,126,231.27 USD** · pure net **87,082,106.28 USD** · **13,316.43 h** attributed NPT (contractor **11,183.43 h** → **10,717,456.94 USD**; client/subsurface **2,133.00 h** → **2,044,125.00 USD**; all-NPT deduction **12,761,581.94 USD**) · 1,304 DDR days · 12,264.76 m recorded 24-h footage.

The 23,000 USD/DAY rate is the **sourced** CBS `Rig Service` unit price, never an assumed market number. Every derived basis is recorded in `04_Four_Track_Cost` and in `an_cost_track_audit` (per well + FIELD). **The 50/50 waiting-time fallback was removed**; waiting hours are never split. Blank attributed-party cells in DDR are retained verbatim (`blank_attribution_h`) and classified only by the source convention (literal `Unattributed` → contractor-side); the per-well register is `an_npt_attribution_audit`. All 13 operations reconcile to **≤ 0.01 USD** after rounding (max delta **0.0092 USD**).

**Calendar-basis flags (4)** — source-basis inconsistencies, not recalculated cost: `AZNS-85` CBS days missing; `AZNS-12` / `AZNS-166` / `AZNS-127-ST2` CBS days > DDR days. Raised as DA-21 / P1 action for Contracts before final invoice settlement.

## Drilling-Data Analysis Department (deep forensics)

- **40 derived `an_*` tables** (29 from `10_deep_analytics.py` + 11 from `13_pdf_and_cost_forensics.py`), all additive, none overwriting source.
- **21 findings (DA-01…DA-21)** with numeric evidence, interpretation, action and priority (`an_findings`, workbook `51_DA_Findings`, dashboard DA panels, FA report sections 6–9).
- Data-quality highlights: numeric-as-text **328,989 / 1,196,084 cells (27.51%)**; formula audit **47,965 MATCH / 3 MINOR / 43 MISMATCH** of 48,011 recomputable formulas; CBS priced-line arithmetic **17 lines** qty×price≠total + **16** material ≥5% gaps (net signed delta **−8,759,173 USD**, abs **9,093,071 USD**); 24-h time balance **16 under / 1,224 over-coded / 7,393 h carried-over**; mud balance **1,846 / 36,046 (5.12%)** imbalance; **364 low-text PDF pages (8.38%)**.
- **PDF numeric forensics**: 305,447 tokens classified (band × context class × unit group); **538 robust unit-group outliers** flagged, never corrected; **4,286 exact six-decimal PDF↔Excel/fact value coincidences** (a reconciliation lead, not proof of identity); unit stats across **34 unit labels**; page-level table `an_pdf_page_stats` (4,345 rows).
- **Cross-source**: CBS ↔ DDR ↔ EOWR ↔ Programme reconciliations (`an_recon_delta`, 4,298 keys), prognosis vs actual tops (`an_program_vs_actual`, 188), casing checks (3,237), directional (4,388), correlations (125 well + 28 field), learning curves (3 chains).

## Honest source gaps (AF-01 … AF-05)

- **AF-01** — Operation 4 (`AZNS-85`) has no EOWR in `04.EOWR/`; it is analysed in full from DDR + CBS + Programme.
- **AF-02** — Operations 12 and 13 share the 238-page `12,13.AZNS_127_WO2_and ST2_EOWR.pdf` (consolidated ST#2 report). WO2 is separately evidenced by its own DDR pack, CBS and Programme. 11 shared pages counted once.
- **AF-03** — Header discrepancies: `8.AZNS_122_WO2_EOWR.pdf` headed `CCA/C/1401/27`; `9.AZNS_75_EOWR.pdf` titled `AZNS-075-ST1`; ST#2 EOWR header `CCA/C/1405/3`. The campaign map labels C3 `CCA/C/1403/31 / ST2 amendment`. Both readings are retained verbatim.
- **AF-04** — `10 CBS matrial.xlsx` wells 29 and 38 are contract-1 rate benchmarks, not campaign wells.
- **AF-05** — Image-only / low-text pages (364) are archived page by page; no value is invented where the source is silent; native/OCR copies requested.

## Querying the warehouse

```sql
-- every numeric cell with exact provenance
SELECT s.rel_path, s.name, c.r, c.c, c.value
FROM cell_numeric c JOIN sheets s ON s.id = c.sheet_id
WHERE c.value = 23000;

-- every PDF numeric token with provenance
SELECT rel_path, page_no, line_no, token, value, unit, context_class, robust_outlier
FROM an_pdf_tokens WHERE value = 23000;

-- four tracks per well, refreshed with explicit attribution
SELECT * FROM an_cost_track_audit ORDER BY op;
SELECT * FROM an_npt_attribution_audit ORDER BY well_key;
```

Full-text indexes: `sheet_fts` (worksheet text) and `pdf_fts` (page text). Semantic tables: `cbs_items`, `cbs_totals`, `ddr_daily`, `ddr_activities`, `ddr_materials`, `ddr_formations`, `ddr_surveys`, `ddr_bha`, `ddr_casing`, `eowr_facts`, `pdf_tables`, `pdf_num_pages`, `well_events`, `well_lessons`, `well_recs`, `reconciliation`, `npt_breakdown`, `field_timeseries` + **40 `an_*` forensic tables**.

## Reproduce

```bash
.venv/bin/python SAZN_Reports/_pipeline/build_warehouse.py        # schema → Excel → PDF → enrich → facts → analytics → workbook (DB restore: needs only .gz if .sqlite absent)
.venv/bin/python SAZN_Reports/_pipeline/10_deep_analytics.py      # 29 an_* tables, 15 findings, exceptions CSV
.venv/bin/python SAZN_Reports/_pipeline/13_pdf_and_cost_forensics.py  # PDF token census, cost/NPT audit, DA-16…DA-21
.venv/bin/python SAZN_Reports/_pipeline/11_extend_workbook.py     # workbook → 63 sheets
.venv/bin/python SAZN_Reports/_pipeline/06_graphics.py            # 149 base artboards
.venv/bin/python SAZN_Reports/_pipeline/12_analytics_graphics.py  # + 17 AN_* deep-analytics artboards
LD_LIBRARY_PATH=/tmp/alibs/lib/lib:/tmp/alibs/lib node .tools/render_graphics.mjs SAZN_Reports/_data/graphics_jobs.json SAZN_Reports/_data/graphics_overlap.json
LD_LIBRARY_PATH=/tmp/alibs/lib/lib:/tmp/alibs/lib node .tools/render_graphics.mjs SAZN_Reports/_data/analytics_graphics_jobs.json SAZN_Reports/_data/analytics_overlap.json
.venv/bin/python SAZN_Reports/_pipeline/07_optimize_graphics.py
.venv/bin/python SAZN_Reports/_pipeline/08_reports.py             # Golden Prototype Gate → 30 DOCX (+ DOCX audit JSON)
.venv/bin/python SAZN_Reports/_pipeline/14_findings_fa.py         # گزارش فنی–مالی (FA), SAZN-FA-DA-001
.venv/bin/python SAZN_Reports/_pipeline/09_dashboard.py           # self-contained HTML dashboard
LD_LIBRARY_PATH=/tmp/alibs/lib/lib:/tmp/alibs/lib node .tools/check_dashboard.mjs  # headless Chromium verifier
```

Generated 2026-11-08 on branch `arena/96a68a69-sazn`.
