# Store Sales Report

This program is tailored to the four store sales workbooks.

## Expected layout

Each workbook should have:

- One sheet
- Header row on row 1
- Sales values in column `H`
- The exact headers:
  - `Date`
  - `Bill No`
  - `Customer Name`
  - `Customer Mobile No`
  - `Payment Method`
  - `Sales Via`
  - `Coupon Discount`
  - `Total Rate(INR)`

## What it does

- Uses each Excel file name as the store name
- Sums the sales in column `H`
- Adds `Date From` and `Date To` details for each store
- Formats totals with the rupee symbol
- Prints each store total and the grand total
- Writes `store_sales_summary.xlsx` with a row for each store and a per-store detail sheet

## Run

```bash
python store_sales_report/sales_summary.py
```

You can also pass the files explicitly:

```bash
python store_sales_report/sales_summary.py Store1.xlsx Store2.xlsx Store3.xlsx "Store4.xlsx"
```
