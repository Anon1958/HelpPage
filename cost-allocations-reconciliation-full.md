# Cost Allocations Reconciliation — Claude Code Prompt

## Context

You are helping David, a CFA working in FP&A at RBC Wealth Management (US), build a Python-based reconciliation and analysis tool for global functions cost allocations. This is a high-priority project assigned by Eric (senior stakeholder). The goal is to take a detailed Canadian-dollar-denominated allocations file from Functions Finance in Toronto and reconcile it back to USD figures in RBC's EPM (Enterprise Performance Management) system.

This project lives under: **FPNA > Projects > Allocations Review**

---

## Input Files

There are two primary Excel workbooks that drive this analysis. David will specify exact file paths once uploaded.

### File 1: Global Functions Allocations Continuity File (the "detail file")

- **Source:** Functions Finance team in Toronto (pulled from SharePoint)
- **Description:** A large, granular file containing all global functions cost allocations for the 2026 plan
- **Currency:** Canadian dollars (CAD) — conversion to USD will be required
- **Key columns/fields to expect** (verify by inspecting the actual file):
  - **Line of Business** — filter for Wealth Management USA (may also appear as US Wealth Management or similar; confirm the exact label)
  - **Function / Sub-function hierarchy** — e.g., Group Risk Management > Operational Risk > Operational Risk Executive (typically 3-4 layers deep)
  - **Methodology codes** — e.g., `QESO Pool`, `Pool`, `Metric/Usage-Based`, `Direct`, etc. These describe HOW costs are allocated. There are only ~3-4 distinct methods in this file.
  - **Transit codes** — internal codes that may indicate geography (Canadian vs. US transit). This is a secondary enrichment task.
  - **Destination transit** — may help determine if the allocating department is Canada-based
  - **Allocation amounts** — the dollar figures being allocated (Column V or similar; verify)
- The file is structured as a pivot table or pivot-table-like layout with many rows

### File 2: EPM Report / Query (the "summary file")

- **Source:** RBC's EPM system
- **Description:** A summary-level report showing planned cost allocation line items in USD for 2026
- **Key fields:** Function-level totals (e.g., Group Risk Management total = ~$8M USD for 2026 plan)
- This is the "target" that the detail file must reconcile to

---

## Task Breakdown

### Phase 1: Data Discovery & Profiling

1. Load and inspect both Excel files — read all sheets, identify headers, data types, row counts, and key columns
2. Print a structural summary of each file: sheet names, column names, sample rows, unique values in key categorical columns (Line of Business, Function, Sub-function, Methodology, etc.)
3. Identify the correct Line of Business filter — find the exact label for US Wealth Management (could be `Wealth Management USA`, `US Wealth Management`, `WM USA`, etc.)
4. Identify the currency — confirm the detail file is in CAD; confirm the EPM file is in USD
5. Identify the allocation amount column in the detail file

### Phase 2: Total-Level Reconciliation

1. Filter the detail file to only Wealth Management USA (or equivalent label)
2. Sum all allocation amounts from the filtered detail file (this gives the CAD total)
3. Convert CAD to USD — use a simple conversion factor (start with 1 CAD ≈ 0.72 USD, or prompt David for the plan rate). The plan FX rate should ideally come from the EPM file or a known planning assumption.
4. Compare the converted USD total to the EPM total for global functions allocations
5. Report the variance: absolute dollar difference and percentage difference
6. If the totals are "close" (Eric's word), proceed to Phase 3. If there's a material gap, flag it and investigate potential causes (missing rows, FX rate mismatch, scope differences, etc.)

### Phase 3: Function-Level Reconciliation (Start with Group Risk Management)

1. Build a mapping from the detail file's function hierarchy to EPM line items
2. Start with Group Risk Management (GRM) as the test case:
   - Sum all GRM sub-allocations from the detail file (CAD), convert to USD
   - Compare to the GRM line item in EPM (~$8M USD for 2026 plan)
   - Report the variance
3. Expand to all functions — produce a reconciliation table:

| Function | Detail File (CAD) | Detail File (USD) | EPM (USD) | Variance ($) | Variance (%) |
|----------|-------------------|-------------------|-----------|-------------|-------------|

4. Flag any functions with material variances for investigation

### Phase 4: Sub-Function Detail & Enrichment

1. For each function, break down by sub-function (the 3-4 layers of hierarchy)
2. Add methodology descriptions — tag each allocation line with its methodology code and (when available) a plain-English description
3. Add geography enrichment — using transit codes, flag whether each allocating department appears to be Canada-based or US-based
4. Produce a detailed output table that could serve as the working analytical file:

| Function | Sub-Function 1 | Sub-Function 2 | Sub-Function 3 | Methodology | Transit | Geography (CA/US) | Amount (CAD) | Amount (USD) | EPM Line Item | Variance |
|----------|---------------|---------------|---------------|-------------|---------|-------------------|-------------|-------------|--------------|----------|

---

## Technical Requirements

- **Language:** Python (pandas, openpyxl)
- **Output:**
  - Console summary of reconciliation results at each phase
  - A clean Excel workbook (`reconciliation_output.xlsx`) with tabs for:
    - **Summary** — total-level recon
    - **By Function** — function-level recon table
    - **Detail** — full filtered dataset with enrichments
    - **Data Profile** — structural summary of source files
- **Code style:** Well-commented, modular functions, clear variable names
- **Error handling:** Graceful handling of missing columns, unexpected labels, or format issues — print warnings rather than crashing
- **FX Rate:** Parameterize the CAD/USD conversion rate at the top of the script so it can be easily updated

---

## File Paths (to be filled in by David)

```python
# Update these paths to point to your actual files
DETAIL_FILE_PATH = "path/to/allocations_continuity_2026.xlsx"
EPM_FILE_PATH = "path/to/epm_report_2026.xlsx"

# Specify sheet names if needed (or set to None to auto-detect)
DETAIL_SHEET_NAME = None    # e.g., "Sheet1" or "Allocations"
EPM_SHEET_NAME = None       # e.g., "Sheet1" or "Query Results"

# FX conversion rate (CAD to USD) — update with the 2026 plan rate
CAD_TO_USD_RATE = 0.72

# Output file
OUTPUT_FILE_PATH = "reconciliation_output.xlsx"
```

---

## Important Notes

- This data is for the **2026 plan**. There may also be 2025 actuals or forecast versions of the detail file — ignore those for now.
- The detail file is the **ONLY** granular data source. There are no other feeds. The EPM is the control total.
- Methodology codes (QESO Pool, Pool, Metric/Usage-Based, Direct, etc.) are critical context. Eventually David will need a reference document defining each code, but for now, just capture and categorize them.
- **"Pool" methodology** generally means costs are allocated based on a rolling multi-quarter ratio of (revenue + non-variable expense) as a percentage of RBC total. The more revenue/expense a business unit generates, the higher its allocation share.
- Prior attempts at this reconciliation have not been successful. Approach this methodically and document assumptions clearly.
- **Eric wants David to understand this data better than anyone.** The code should be exploratory and educational, not just a black box.
- Start with the total-level recon first. Only proceed to detailed analysis once we confirm the data is reasonably complete and ties to EPM.
- **Key contacts:** Craig has familiarity with this project. Corey was the previous owner. Nick is David's new direct manager.
