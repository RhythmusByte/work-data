from __future__ import annotations

import argparse
import glob
from pathlib import Path
from typing import Iterable

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


HEADER_CANDIDATES = [
    "total rate",
    "total rate (inr)",
    "sale amount",
    "sales amount",
    "amount",
    "total",
    "net total",
    "grand total",
    "sales",
]

HEADER_SCAN_LIMIT = 20
COL_SCAN_LIMIT = 20


def normalize_text(value) -> str:
    return str(value or "").strip().lower()


def find_header_row(ws) -> int:
    for row_idx in range(1, min(ws.max_row, HEADER_SCAN_LIMIT) + 1):
        row_values = [normalize_text(ws.cell(row_idx, col_idx).value) for col_idx in range(1, min(ws.max_column, COL_SCAN_LIMIT) + 1)]
        if any(value in HEADER_CANDIDATES for value in row_values):
            return row_idx
    raise ValueError("Could not find a header row with a sales/amount column.")


def find_amount_column(ws, header_row: int) -> int:
    for col_idx in range(1, min(ws.max_column, COL_SCAN_LIMIT) + 1):
        header = normalize_text(ws.cell(header_row, col_idx).value)
        if not header:
            continue
        if header in HEADER_CANDIDATES:
            return col_idx
        for candidate in HEADER_CANDIDATES:
            if candidate in header:
                return col_idx
    raise ValueError("Could not identify the sales amount column.")


def to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    try:
        return float(text)
    except ValueError:
        return None


def sum_sales_file(path: Path) -> float:
    workbook = openpyxl.load_workbook(path, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]

    header_row = find_header_row(worksheet)
    amount_col = find_amount_column(worksheet, header_row)

    total = 0.0
    for row_idx in range(header_row + 1, worksheet.max_row + 1):
        value = to_number(worksheet.cell(row_idx, amount_col).value)
        if value is not None:
            total += value

    return total


def collect_files(file_args: list[str]) -> list[Path]:
    if file_args:
        paths = [Path(item).expanduser().resolve() for item in file_args]
    else:
        paths = [Path(p).resolve() for p in glob.glob("*.xlsx")]

    paths = [p for p in paths if p.suffix.lower() == ".xlsx" and p.is_file() and not p.name.startswith("~$")]
    return sorted(paths)


def write_summary_workbook(rows: Iterable[tuple[str, str, float]], output_path: Path) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Store Sales"

    headers = ["Store Name", "Source File", "Total Sales"]
    worksheet.append(headers)

    header_fill = PatternFill(fill_type="solid", fgColor="FF1B365D")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFFFF")
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    number_format = "#,##0.00"
    grand_total = 0.0
    for store_name, source_file, total_sales in rows:
        worksheet.append([store_name, source_file, total_sales])
        grand_total += total_sales

    total_row = worksheet.max_row + 1
    worksheet.cell(total_row, 1, "TOTAL")
    worksheet.cell(total_row, 3, grand_total)

    for row_idx in range(2, total_row + 1):
        worksheet.cell(row_idx, 3).number_format = number_format

    for cell in worksheet[total_row]:
        cell.font = Font(name="Arial", size=11, bold=True)
        cell.fill = PatternFill(fill_type="solid", fgColor="FFE6ECF4")

    worksheet.cell(total_row, 3).number_format = number_format

    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 40
    worksheet.column_dimensions["C"].width = 16
    worksheet.freeze_panes = "A2"

    workbook.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Calculate total sales for multiple store Excel files."
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="Excel files to process. If omitted, all .xlsx files in the current directory are used.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="store_sales_summary.xlsx",
        help="Output workbook name.",
    )
    args = parser.parse_args()

    paths = collect_files(args.files)
    if not paths:
        print("No Excel files found.")
        return 1

    results: list[tuple[str, str, float]] = []
    skipped: list[tuple[str, str]] = []

    for path in paths:
        store_name = path.stem
        try:
            total_sales = sum_sales_file(path)
        except Exception as exc:  # noqa: BLE001 - report and continue for batch inputs
            skipped.append((path.name, str(exc)))
            continue

        results.append((store_name, path.name, total_sales))
        print(f"{store_name}: {total_sales:,.2f}")

    if not results:
        print("No files could be processed.")
        for filename, reason in skipped:
            print(f"Skipped {filename}: {reason}")
        return 1

    grand_total = sum(total for _, _, total in results)
    print(f"Grand Total: {grand_total:,.2f}")

    output_path = Path(args.output).expanduser().resolve()
    write_summary_workbook(results, output_path)
    print(f"Saved summary workbook: {output_path}")

    if skipped:
        print("\nSkipped files:")
        for filename, reason in skipped:
            print(f"- {filename}: {reason}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
