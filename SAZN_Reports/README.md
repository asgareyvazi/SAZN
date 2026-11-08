# SAZN — South Azadegan Reference Programme

Complete, zero-summarisation deliverable built from the 51 source documents in `CBS/`, `DDRs/`, `04.EOWR/` and `Well Program/`. The four original source folders are never modified.

## Deliverables

| File / folder | Content |
|---|---|
| `SAZN_Master_Warehouse.sqlite` | Full archive: 2,810 exact worksheet cell blobs, 324,150 numeric cells, 155,248 formulas, 4,345 lossless PDF page texts, 305,447 numeric PDF tokens, 1,349 numbered tables, FTS indexes, semantic DDR/CBS tables, derived KPI / events / lessons / recommendations / reconciliation / NPT / time-series. Every value carries `file → sheet/page → cell/row/table` provenance. |
| `SAZN_Master_Audit_Workbook.xlsx` | 23 sheets: source registry, four-track cost, CBS tariff & materials, DDR 24-h log, NPT, formations, BHA, casing, surveys, EOWR facts, PDF tables, lessons, REC cards, audit findings, full cell/page inventory and field KPI. |
| `SAZN_Executive_Interactive_Dashboard.html` | Self-contained, offline, 532,268 searchable rows. Live search across numbers, wells, rigs, materials, files, sheets and pages with exact provenance; four-track charts and contract close-out. |
| `Word_Books/` | 30 DOCX: 13 single-well FA + 13 single-well EN, plus 2 field-master FA/EN and 2 executive-management FA/EN. Each book has **24 chapters, 96 provenance tables and 11 custom graphics**. FA = B Nazanin + full RTL (`w:bidi`, `w:bidiVisual`); EN = Times New Roman. All headings/table headers centred; navy header, banded rows, repeated header, `cantSplit`. |
| `SAZN_Graphics/` | 149 PNG (2400×1800 px / 300 DPI) rendered from HTML5 + CSS3 + SVG + Canvas in headless Chromium. 0 text-overlap pairs, English-only text, inverted depth axis, mean ≈169 KB. |
| `SAZN_DOCX_Audit.json` | Gate + full audit of all 30 Word books. |
| `_data/visual_review.json` | Graphics QA: overlap census, size distribution, sample visual review. |
| `_data/well_profiles.json.gz` | Gzip-compressed derived per-well profiles (0.65 MB); `common.load_json()` reads it transparently.
| `_pipeline/` | Reproducible generators: `common.py`, `build_warehouse.py`, `06_graphics.py`, `07_optimize_graphics.py`, `08_reports.py`, `09_dashboard.py`, plus `.tools/render_graphics.mjs` (workspace tool). |

## Source census

13 CBS workbooks / 202 sheets · 13 DDR packs / 2,608 sheets · 11 EOWR / 2,724 pages · 14 Well Programmes / 1,621 pages = **51 files**. 1,196,084 non-empty Excel cells, 324,150 numeric cells, 155,248 formulas, 4,345 PDF pages, 305,447 numeric PDF tokens, 1,349 numbered PDF tables.

## Four-track time / cost accounting

1. **CBS Plan** = sourced cost ceiling (`Total Cost-EPD & Rigless`, EPD + rig-less, or `Cost Breakdown` grand total).
2. **Gross Calendar** = CBS Plan + (DDR days − CBS planned days) × sourced rig day rate.
3. **Payable excl. contractor zero-rate NPT** = Gross − contractor NPT hours / 24 × rate.
4. **Pure Net excl. all NPT** = Gross − all attributed NPT hours / 24 × rate.

Field totals: CBS plan **94,797,894.39 USD** · gross **99,843,688.22 USD** · payable **89,126,231.27 USD** · pure net **87,082,106.28 USD** · **13,316.43 h** attributed NPT · 1,304 DDR days · 12,264.76 m recorded 24-h footage.

The 23,000 USD/DAY rate is the **sourced** CBS `Rig Service` unit price (visible in `8.AZNS_122_WO2_CBS.xlsx` and the contract-1 benchmark sheets), never an assumed market number. Every derived basis is recorded in `04_Four_Track_Cost` (`plan_basis`, `rate_basis`) and in `00_README`. No 50/50 waiting-time fallback was needed: all 13 operations have explicit DDR `N/NPT` rows with an attributed party.

## Honest source gaps (AF-01 … AF-05)

- **AF-01** — Operation 4 (`AZNS-122 WO1`) has no EOWR in `04.EOWR/`; it is analysed in full from 112 DDR sheets, its CBS workbook and the 66-page Workover-1 Programme.
- **AF-02** — Operations 12 and 13 share the 238-page `12,13.AZNS_127_WO2_and ST2_EOWR.pdf`. It is a consolidated ST#2 report; WO#2 appears as carried-forward history and WorkOver Section rows. Page split: 11 shared, 11 WO2, 216 ST2 (plus 11 shared first pages retained as shared). WO2 is separately evidenced by its own 136-sheet DDR pack, CBS and Programme.
- **AF-03** — Header discrepancies: `8.AZNS_122_WO2_EOWR.pdf` headed `CCA/C/1401/27`; `9.AZNS_75_EOWR.pdf` titled `AZNS-075-ST1`; ST#2 EOWR header `CCA/C/1405/3`. The campaign map labels C3 `CCA/C/1403/31 / ST2 amendment`. Both readings are retained verbatim.
- **AF-04** — `10 CBS matrial.xlsx` wells 29 and 38 are contract-1 rate benchmarks, not campaign wells.
- **AF-05** — Image-only / low-text pages are archived page by page; no value is invented where the source is silent.

## Querying the warehouse

```sql
-- every numeric cell with exact provenance
SELECT s.rel_path, s.name, c.r, c.c, c.value
FROM cell_numeric c JOIN sheets s ON s.id = c.sheet_id
WHERE c.value = 23000;

-- any number inside a PDF page
SELECT f.rel_path, p.page_no, p.chapter, p.n_chars
FROM pdf_pages p JOIN source_files f ON f.id = p.file_id
WHERE p.well_key = 'AZNS-207' AND p.chapter LIKE '08 %';

-- four tracks per well
SELECT well_key, payload FROM well_kpi;
```

Full-text indexes: `sheet_fts` (worksheet text) and `pdf_fts` (page text). Semantic tables: `cbs_items`, `cbs_totals`, `cbs_stim`, `cbs_accomm`, `ddr_daily`, `ddr_activities`, `ddr_time`, `ddr_materials`, `ddr_formations`, `ddr_surveys`, `ddr_bha`, `ddr_casing`, `ddr_pob`, `ddr_notes`, `eowr_facts`, `pdf_tables`, `pdf_num_pages`, `well_events`, `well_lessons`, `well_recs`, `reconciliation`, `npt_breakdown`, `field_timeseries`.

## Reproduce

```bash
.venv/bin/python SAZN_Reports/_pipeline/build_warehouse.py     # schema → Excel → PDF → enrich → facts → analytics → workbook
.venv/bin/python SAZN_Reports/_pipeline/06_graphics.py         # 149 artboards
LD_LIBRARY_PATH=/tmp/alibs/lib node .tools/render_graphics.mjs # headless Chromium @ 2400 px
.venv/bin/python SAZN_Reports/_pipeline/07_optimize_graphics.py
.venv/bin/python SAZN_Reports/_pipeline/08_reports.py          # Golden Prototype Gate → 30 DOCX
.venv/bin/python SAZN_Reports/_pipeline/09_dashboard.py
```

Generated 2026-11-08 on branch `arena/96a68a69-sazn`.
