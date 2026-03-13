# Cost Allocations — Iteration 2: Consolidate & Polish the Reconciliation Workbook

## What you've done so far:

You've already completed the initial reconciliation and produced several output files in `C:\Users\iams395\Cost Allocations`:

- **reconciliation_output.xlsx** — v1, early attempt. Known issue: the "Persp Paste" sheet contained ALL RBC data, not just US Wealth Management, which caused a significant variance.
- **reconciliation_output_v2.xlsx** — v2, refined. Properly filtered for US Wealth Management and used the EPM Detail sheet for correct function breakdowns.
- **reconciliation_final.xlsx** — final version, comprehensive reconciliation with 4 tabs. Total-level variance of $0.26MM (0.2%).
- **Cost Allocations explore.ipynb** — interactive notebook with 15 cells of analysis.
- **Python scripts:** `explore_files.py`, `reconciliation_analysis.py`, `reconciliation_v2.py`, `reconciliation_final.py`

## The two source files:

- `Copy of Continuity Report 2025 Act - 2026 Plan v1 USWM SP.xlsb` (detail file, CAD-denominated)
- `Copy of WMUS Functions Allocations 2025-2026.xlsb` (EPM file, USD-denominated)

---

## What I need you to do now:

Create **ONE** consolidated Excel workbook called `Cost Allocations Reconciliation - Master.xlsx` that merges all output iterations into a single, well-organized deliverable. Structure the sheets left to right from most valuable/current to oldest/archival.

### Sheet 1: Executive Summary

This is the presentation-ready sheet I'd show Eric. Include:

- **Reconciliation status:** RECONCILED, 0.2% total variance, $0.26MM
- **Grand total reconciliation table:** Detail File 233,679K CAD → $169.78MM USD vs EPM $169.52MM USD, variance $0.26MM (0.2%)
- **Function-level breakdown table** with all functions and their variances:
  - GRM Group: $32.99MM vs $32.94MM (+0.05, 0.2%)
  - CFO Group: $25.08MM vs $24.15MM (+0.93, 3.9%) — **FLAG THIS**
  - Corp Expenditures: $10.82MM vs $10.80MM (+0.02, 0.2%)
  - CAE Group: $15.51MM vs $15.49MM (+0.02, 0.2%)
  - CLAO Group: $44.18MM vs $44.11MM (+0.07, 0.2%)
  - US Regional Support: $11.72MM vs $10.88MM (+0.84, 7.7%) — **FLAG THIS**
  - HR & BMCC: Check if this function appears in the detail file. It was listed in the EPM summary but did not appear in the final function-level table. If it exists, include it. If it was mapped into another category, add a note explaining where it went.
- **YoY Bridge** (Prior Year → 2026 Plan): Prior Year Base 216,319K CAD → FX Impact -2,401K → Methodology Changes +13,960K (largest driver) → Challenge Savings -7,490K → 2026 Plan Total 233,679K CAD (+8% YoY)
- **FX rate:** 1.3764 CAD per USD (2026 Plan Rate from PPTX sheet)
- **Open items / next steps:** (1) Investigate CFO Group variance $0.93MM / 3.9%, (2) Validate US Regional Support mapping in EPM 7.7% variance, (3) Document methodology codes for reference, (4) Share findings with Eric

**Format this professionally:** bold headers, borders, number formatting with commas and 2 decimals for currency, conditional formatting on variance percentages (green <1%, yellow 1-5%, red >5%), freeze top row.

### Sheet 2: Function Detail

Full function-level reconciliation with sub-function breakdowns (the 3-4 layers of hierarchy from the continuity file). Include methodology codes for each allocation line. Include transit codes and geography tagging (Canadian vs US transit) where available. Show amounts in both CAD and USD columns. Enable filters on all columns. This is the working analytical sheet.

### Sheet 3: GRM Deep Dive

Group Risk Management broken out specifically since that was Eric's test case ($32.99MM detail vs $32.94MM EPM). Show all sub-functions: Operational Risk, Enterprise Risk, Operational Risk Executive, etc. Break down by methodology (QESO Pool, Pool, Direct, Metric/Usage-Based). Compare each sub-function line to EPM where possible. This sheet serves as the template for how we'd eventually analyze every function.

### Sheet 4: Methodology Reference

List every unique methodology code found in the detail file. For each code show: count of allocation lines using it, total CAD allocated, total USD allocated, percentage of total allocations. Add a description column — populate what you can (QESO Pool = pool-based using revenue + non-variable expense ratios, Pool = similar pool methodology, Direct = direct charge, Metric/Usage-Based = allocated by specific usage driver). Leave blanks for any codes you're unsure about so I can fill in later with Craig's help.

### Sheet 5: Data Profile

Structural summary of both source `.xlsb` files: sheet names, row counts, column names, data types. Unique values for key categorical fields (Line of Business labels, Function names, Sub-function names, Methodology codes). Any data quality notes — missing values, unexpected labels, format issues encountered during the analysis.

### Sheet 6: Recon v2 (Archive)

Pull in the content from `reconciliation_output_v2.xlsx`. Add a text note in cell A1: *"V2 — Refined analysis. Properly filtered for Wealth Management USA using EPM Detail sheet. Resolved the v1 scoping issue."* Then the data starts below.

### Sheet 7: Recon v1 (Archive)

Pull in the content from `reconciliation_output.xlsx`. Add a text note in cell A1: *"V1 — Initial attempt. Known issue: Persp Paste sheet contained ALL RBC data, not scoped to US Wealth Management. Replaced by v2."* Then the data starts below.

---

## Formatting across all sheets:

- Currency columns: comma-separated with 2 decimal places
- Percentage columns: formatted as percentages with 1 decimal
- Header rows: bold, dark blue background (`#003366`), white text
- Freeze panes on the header row for every sheet
- Auto-fit column widths
- Filters enabled on sheets 2 through 5
- Conditional formatting on any variance % column: green (<1%), yellow (1-5%), red (>5%)

---

## After creating the workbook:

1. **Delete** the old intermediate output files: `reconciliation_output.xlsx`, `reconciliation_output_v2.xlsx`, `reconciliation_final.xlsx`
2. **Keep** all `.py` scripts and the `.ipynb` notebook as-is — those are code artifacts, not deliverables
3. **Print** a summary confirming what was consolidated, the final file path, and sheet count

---

## Important notes:

- Pull data from the existing output files and your scripts wherever possible. Only re-read the source `.xlsb` files if you need to fill in gaps (like the HR & BMCC question or missing sub-function detail).
- The FX rate is **1.3764 CAD per USD**. Do not recalculate it.
- Output must be `.xlsx` format, not `.xlsb`. Use `openpyxl` for all formatting.
- If any of the archived v1/v2 workbooks had multiple tabs, flatten each into a single sheet with clear labels.
