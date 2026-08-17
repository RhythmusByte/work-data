# Salary Report

This tool builds a salary-only workbook from multiple store expense files.

## Expected input

- One sheet per workbook
- Header row on row 1
- Columns:
  - `Date`
  - `Expense`
  - `Amount`
  - `Notes`

The script filters rows where `salary` appears in the expense category or notes.

## What it does

- Prompts you to confirm the store name for each file
- Creates one main summary sheet with store totals
- Creates one sheet per store with the salary rows and total salary given
- Saves a combined workbook in the same directory

## Run

From the repo root:

```bash
python salary_report/salary_summary.py
```

You can also pass files explicitly:

```bash
python salary_report/salary_summary.py Ayoor.xlsx Ettumanoor.xlsx Ulloor.xlsx "World-Market.xlsx"
```
