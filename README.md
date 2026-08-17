# Shop Data

Small Excel reporting tools for shop and supermarket data.

## What’s in this repo

- [`expense_statement.py`](./expense_statement.py): builds an expense summary workbook from one or more `expenses_report*.xlsx` files.
- [`salary_report.py`](./salary_report.py): filters expense rows containing `salary` and generates a salary-only statement.
- [`pnl_statement.py`](./pnl_statement.py): builds a profit and loss workbook from sales, purchase, expense, and purchase-return reports.
- [`store_sales_report/sales_summary.py`](./store_sales_report/sales_summary.py): sums sales for multiple store workbooks and writes a combined summary.
- [`supermarket/`](./supermarket): notebook, sample workbook, charts, and JSON sales data for supermarket analysis.

## Requirements

- Python 3.14+
- `openpyxl`
- `pandas`
- `python-dateutil`

Install dependencies with:

```bash
uv sync
```

## Usage

Run the scripts from the repo root.

### Expense summary

```bash
python expense_statement.py
```

### Salary-only statement

```bash
python salary_report.py
```

### Profit and loss statement

```bash
python pnl_statement.py
```

### Store sales summary

```bash
python store_sales_report/sales_summary.py
```

You can also pass the workbook paths explicitly:

```bash
python store_sales_report/sales_summary.py Ayoor.xlsx Ettumanoor.xlsx Ulloor.xlsx "World-Market.xlsx"
```

## Expected workbook layout

The store sales tool expects the first sheet of each workbook to use this layout:

- Row 1 headers
- Column `H` as `Total Rate(INR)`
- Store name taken from the Excel filename

## Generated files

The repo ignores uploaded Excel inputs, generated summary workbooks, and Python cache files. Those files are local-only and should not be committed.
