# GEL HR Metrics – FY2027 workbook set

All workbooks are built from **0 GCG (33 Companies) HR Metrics FY2027**, the new model:
- 70 metric rows
- Quarter tab: ∑Q1–∑Q4 plus an FY2027 column
- 12 monthly tabs from OCT to SEP

The folder tree matches the SharePoint library, so the relative links between files resolve once the files are uploaded to the same places.

Sheet password: **GCG** on every sheet.

## Data flow

```
Company tabs ─► monthly tabs (OCT…SEP) ─► Quarter ─┬─► GEL HR Metrics FY2027           (Auto, Building, Head Office,
                                                   │                                       Manufacturing, Services, Shipping)
GCG cluster workbooks (monthly tabs) ─► 0 GCG (33 Companies) monthly tabs ─► Quarter ─┘  (GCG)
```

| File | Feeds |
|---|---|
| `GCG HR Metrics/FY 2027/GCG HDQ · Central America · North Caribbean · South America · South Caribbean HR Metrics FY2027.xlsx` | 0 GCG (33 Companies) |
| `GCG HR Metrics/FY 2027/0 GCG (33 Companies) HR Metrics FY2027.xlsx` | GEL |
| `Auto / Building / Head Office / Manufacturing / Services / Shipping HR Metrics FY2027.xlsx` | GEL |
| `All GEL Divisions FY2025/GEL HR Metrics FY2027.xlsx` (same folder as GEL FY2026) | – |

## What goes where

**Monthly tabs (cluster and division workbooks):** type the monthly data into the company columns. Only the input rows are unlocked:
- Headcount split
- Standard hours
- Subcontracted headcount and hours
- HR FTE
- Hires, exits and terminations
- Vacancies
- Check-ins % and objectives
- Days lost
- Training hours, trained staff, programmes and costs
- Headcount last Month (OCT tab only; later months carry it over automatically)

**Company tabs** (one per company):
- **Absence:** enter the *Working Days* for each month. Total Employees and Days Lost come from the monthly tab. The monthly Non-certified and Certified absence rates feed rows 46 and 47 of the monthly tab.
- **Time to Hire:** enter the Job Title, No. of Posts, Job Opening Date and Contract Signed Date for permanent and temporary hires. The *Time to Hire by Month* block works out the YTD Time to Hire for each month. It counts contracts signed from 1-Oct-2026 up to that month and weights them by posts. It feeds rows 39 and 40 of the monthly tab.
- The Check-ins YTD block from FY2026 has been removed. Check-ins % is now typed straight into the monthly tabs (row 41).

**0 GCG (33 Companies) and GEL:** these contain no inputs. Every value is linked or calculated.

## Formula fixes to the model

| Where | Problem in the model | Fix |
|---|---|---|
| Quarter, NOV–SEP columns | Pulled `NOV!AI`, `DEC!AH` … `SEP!V`, drifting one column per month away from TOTAL (`AJ`), so they showed individual companies instead of the total | Every month now reads the TOTAL column |
| Monthly and Quarter row 11, TOTAL column | Employee FTE gave `#VALUE!` when a headcount bucket was blank | Blank buckets now count as 0 (`N()`) |
| Monthly and Quarter row 37 | Authorized Positions gave `#VALUE!` when vacancies were blank | Blanks now count as 0 (`N()`) |
| Turnover rows 25/26/29/30/33, Training Invest. rows 68/69 | Errored when a linked input was blank (breaks once values are linked) | Blank inputs now count as 0 (`N()`) |
| Monthly TOTAL rows 46/47 (weighted absence) | Errored when a headcount cell held a linked blank | Each term is wrapped in `IFERROR` (the pattern the model already uses in its Quarter tab) |
| Quarter rows 39/40 (Time to Hire YTD) | Stayed blank until the last month of the quarter or year was entered | Show the latest month reported |
| OCT tab | Held test values and was unprotected | Cleared and protected |

## Changes from FY2026

- **Companies follow the FY2027 model:** 33 GCG companies. GCG Ground/Security Costa Rica, AGO Security El Salvador, GCG Bermuda and Island Grill Barbados are no longer included.
- **The non-GCG divisions list only their own companies.** FY2026 had about 23 template columns; each division now has only the companies with a company tab.
- **Tab names made consistent:** `(JBH_Auto)`, `(Culinary_CRI)`, `(GEL_Holding)`. The months are in order: HDQ had APR before JAN, and Central America used `ABR`.
- **FY2026 cross-link bug fixed:** Sint Maarten Airport Dining linked to the Sky Dining tab. Every company now links to its own tab.
- **Time to Hire YTD** is now weighted by posts filled. FY2026 divided total days by *all* posts, including unfilled ones.
- **GEL consolidation for each quarter and FY2027:**
  - Counts are summed.
  - Ratios are recalculated from the consolidated figures.
  - Absence rates and Staff Trained % are weighted by headcount.
  - Time to Hire, Check-ins and standard hours are averaged.

## Not changed

- `GCG HR Metrics/Overtime/OT_Model.xlsx` still links to the FY2026 GCG workbook.
- `Data Collection Compliance Dashboard.xlsx` links to FY2025 files, so it was left out.

## Rebuilding

`tools/build_fy2027.py` regenerates the whole set from the model and the FY2026 files. The FY2026 files are used only to size each company's Time to Hire table.

`tools/test_fy2027.py` checks the set:
1. Fills every workbook with random data.
2. Recalculates it in LibreOffice.
3. Checks the results against an independent Python calculation.
4. Replaces the cross-workbook links with the linked values and checks that the 0 GCG and GEL roll-ups match the source workbooks.
