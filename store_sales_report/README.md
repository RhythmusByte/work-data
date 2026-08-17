# Store Sales Report

This small program reads multiple Excel files, treats each file name as a store name, sums the sales column in each workbook, and writes a summary workbook.

## Usage

From this directory or the repo root:

```bash
python store_sales_report/sales_summary.py store1.xlsx store2.xlsx store3.xlsx store4.xlsx
```

If you omit file arguments, it will look for all `.xlsx` files in the current directory.

## Output

- Console output with each store total and the grand total
- `store_sales_summary.xlsx` with one row per store and a final total row
