# Cost Allocations — Iteration 3: Formula Rebuild (Replace Hardcoded Values)

This is a critical rebuild of the Executive Summary and supporting analysis sheets. The problem right now is that every number on the Executive Summary is a hardcoded value — there are no formulas linking back to the source data sheets in this same workbook. That means I can't trace any number, I can't cross-check anything, and I can't answer questions about where a figure comes from. Eric will ask, and I need to be able to click into any cell and see the formula trail.

---

## What I need you to do:

### 1. Executive Summary — replace all hardcoded values with formulas

Every number on this sheet should be a formula that references the underlying data in the **Source Data (Full)**, **EPM Detail (Full)**, or **EPM Persp Paste (Full)** sheets within this same workbook. Specifically:

- **Grand Total Reconciliation:** The Detail File CAD total (233,679.06K) should be a SUMIF or SUMIFS pulling from Source Data (Full), filtered to the correct Line of Business. The USD conversion should be a formula dividing by the FX rate cell. The EPM total ($169.52MM) should be a SUM or SUMIFS pulling from EPM Detail (Full).

- **Function-Level Breakdown:** Each function row (GRM Group, CFO Group, Corp Expenditures, CAE Group, HR & BMCC, CLAO Group, US Regional Support) — the Detail CAD column should be SUMIFS on Source Data (Full) grouped by the function/group column. The Detail USD column should be the CAD value divided by FX rate. The EPM USD column should be SUMIFS on EPM Detail (Full) matching to the corresponding function.

- **Variance ($)** should be `=Detail USD - EPM USD`. **Variance (%)** should be `=Variance/EPM USD`.

- **The FX rate** should live in ONE cell on the Executive Summary (or Sheet Guide) and every USD conversion should reference that single cell using an absolute reference. Do not repeat the FX rate number anywhere — always point to the one cell.

- **The YoY Bridge values** — if these come from specific columns in Source Data (Full) or from the PPTX sheet, link them with formulas. If they were calculated in the Python script and aren't directly in the source data, then leave them hardcoded BUT add a comment or note cell explaining the calculation logic and source.

### 2. GRM Deep Dive — same treatment

The GRM Deep Dive sheet should also use formulas referencing Source Data (Full), not hardcoded values. The sub-function breakdown counts and amounts should be COUNTIFS and SUMIFS against the source data. The EPM comparison figures should reference EPM Detail (Full).

### 3. Methodology Reference — formula-based where possible

The count of allocations by methodology, total CAD, total USD, and percentage breakdowns should all be COUNTIFS and SUMIFS referencing Source Data (Full), not pasted values.

### 4. Create a named cell or small reference area for key parameters

Put these somewhere clean — either on the Sheet Guide or a dedicated area on Executive Summary:

- **FX Rate** (1.3764) — single cell, named reference, everything points here
- **Line of Business filter value** ("Wealth Management USA" or whatever the exact label is)
- **Plan year** (2026)
- **Source file names**

### 5. Keep the raw data sheets untouched

Keep the **Sheet Guide**, **Source Data (Full)**, **EPM Detail (Full)**, **EPM Persp Paste (Full)** sheets exactly as they are. Do not modify the raw data sheets. Only rebuild the analytical sheets (Executive Summary, GRM Deep Dive, Methodology Reference) with formulas.

### 6. Data Profile stays separate

Data Profile stays as its own standalone file — you already separated that out, keep it separate.

### 7. No emojis

No emojis, no checkmarks, no alert symbols anywhere in the workbook. Plain text only.

### 8. Test the formulas

After rebuilding, verify that the formula-driven totals on Executive Summary still produce the same numbers: Detail File 233,679.06K CAD, $169.78MM USD, EPM $169.52MM USD, variance $0.26MM (0.2%). If they don't match, tell me what's different and why before saving.

---

## The Standard

The end result should be a workbook where I can click into any cell on the Executive Summary and trace it all the way back to a row in the source data. That's the standard for anything I'd show Eric.
