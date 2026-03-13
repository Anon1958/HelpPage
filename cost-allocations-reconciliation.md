# Cost Allocations Reconciliation — 2026 Plan

## Phase 3: Function-Level Reconciliation

- Compare to the GRM line item in EPM (~$8M USD for 2026 plan)
- Report the variance

1. Expand to all functions — produce a reconciliation table:

| Function | Detail File (CAD) | Detail File (USD) | EPM (USD) | Variance ($) | Variance (%) |
|----------|-------------------|-------------------|-----------|-------------|-------------|

1. Flag any functions with material variances for investigation

---

## Phase 4: Sub-Function Detail & Enrichment

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
DETAIL_SHEET_NAME = None  # e.g., "Sheet1" or "Allocations"
EPM_SHEET_NAME = None     # e.g., "Sheet1" or "Query Results"

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
