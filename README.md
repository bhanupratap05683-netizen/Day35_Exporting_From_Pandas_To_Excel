# Day 35 — Exporting pandas to Excel: Multi-Sheet & Polished Reports

**Date:** Day 35 of 84 | 84-Day Python & Excel Roadmap
**Author:** Bhanu Pratap Singh
**Phase:** Phase 2 — Advanced Excel + pandas Basics

---

## What This Does

Reads raw portfolio transaction data from Excel, builds four analytical views using pandas, exports all views as separate sheets into a single workbook, then applies professional styling (title rows, colour-coded headers, conditional P&L formatting) using openpyxl.

---

## Key Concepts Practiced

| Concept | Description |
|---|---|
| `pd.ExcelWriter` | Context manager that writes multiple DataFrames to one file |
| `to_excel(startrow=1)` | Leaves row 0 for a merged title above the header |
| Multi-sheet export | Four sheets from four DataFrames in one `with` block |
| Post-export openpyxl pass | Load saved file, insert titles, style headers, colour P&L |
| `get_column_letter()` | Converts column number to Excel letter for `merge_cells` |

---

## Output

`day35_output.xlsx` — 4 sheets:
- **Transactions** — Full trade log with computed P&L and Return %
- **Sector_Summary** — Aggregated performance by sector
- **Stock_Summary** — Per-stock average return
- **Monthly_PnL** — Month-wise profit & loss trend

---

## Files

```
day35_input.xlsx      ← Raw transaction data (16 rows, pre-styled)
day35_practice.py     ← Main script (Steps 1–5)
day35_output.xlsx     ← Final polished multi-sheet report
README_Day35.md       ← This file
```

---

## Portfolio Connection

This pattern — **pandas for transformation → ExcelWriter for export → openpyxl for styling** — is the backbone of the Financial Dashboard (Day 78) and Expense Tracker (Day 80). Every polished report in your portfolio will reuse this exact three-step workflow.
