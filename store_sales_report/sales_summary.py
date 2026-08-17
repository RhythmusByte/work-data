from __future__ import annotations

import argparse
import glob
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill


HEADER_ROW = 1
DATA_START_ROW = 2
SALES_COLUMN = 8  # Column H
EXPECTED_HEADERS = [
    "Date",
    "Bill No",
    "Customer Name",
    "Customer Mobile No",
    "Payment Method",
    "Sales Via",
    "Coupon Discount",
    "Total Rate(INR)",
]


def normalize_text(value) -> str:
    return str(value or "").strip().replace(" ", "").lower()


def collect_files(file_args: list[str], output_name: str) -> list[Path]:
    if file_args:
        paths = [Path(arg).expanduser().resolve() for arg in file_args]
    else:
        paths = [Path(path).resolve() for path in glob.glob("*.xlsx")]

    result = []
    for path in paths:
        if not path.is_file():
            continue
        if path.suffix.lower() != ".xlsx":
            continue
        if path.name.startswith("~$"):
            continue
        if path.name == output_name:
            continue
        result.append(path)

    return sorted(result)


def validate_layout(ws, source_name: str) -> None:
    headers = [ws.cell(HEADER_ROW, col).value for col in range(1, len(EXPECTED_HEADERS) + 1)]
    normalized_headers = [normalize_text(value) for value in headers]
    normalized_expected = [normalize_text(value) for value in EXPECTED_HEADERS]

    if normalized_headers != normalized_expected:
        raise ValueError(
            f"{source_name}: unexpected header row. "
            f"Expected {EXPECTED_HEADERS}, found {headers}."
        )


def to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def sum_store_sales(path: Path) -> float:
    workbook = openpyxl.load_workbook(path, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    validate_layout(worksheet, path.name)

    total = 0.0
    for row_idx in range(DATA_START_ROW, worksheet.max_row + 1):
        value = to_number(worksheet.cell(row_idx, SALES_COLUMN).value)
        if value is not None:
            total += value
    return total


def write_summary_workbook(rows: list[tuple[str, str, float]], output_path: Path) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Store Sales Summary"

    headers = ["Store Name", "Source File", "Total Sales"]
    worksheet.append(headers)

    header_fill = PatternFill(fill_type="solid", fgColor="FF1B365D")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFFFF")
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    total_sales_all_stores = 0.0
    for store_name, source_file, total_sales in rows:
        worksheet.append([store_name, source_file, total_sales])
        total_sales_all_stores += total_sales

    total_row = worksheet.max_row + 1
    worksheet.cell(total_row, 1, "TOTAL")
    worksheet.cell(total_row, 3, total_sales_all_stores)

    number_format = "#,##0.00"
    for row_idx in range(2, total_row + 1):
        worksheet.cell(row_idx, 3).number_format = number_format

    for cell in worksheet[total_row]:
        cell.font = Font(name="Arial", size=11, bold=True)
        cell.fill = PatternFill(fill_type="solid", fgColor="FFE6ECF4")

    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 34
    worksheet.column_dimensions["C"].width = 16
    worksheet.freeze_panes = "A2"

    workbook.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Calculate total sales for each store workbook and the overall total."
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
        help="Name of the summary workbook to create.",
    )
    args = parser.parse_args()

    output_name = Path(args.output).name
    paths = collect_files(args.files, output_name)
    if not paths:
        print("No Excel files found.")
        return 1

    results: list[tuple[str, str, float]] = []
    failed: list[tuple[str, str]] = []

    for path in paths:
        try:
            total_sales = sum_store_sales(path)
        except Exception as exc:
            failed.append((path.name, str(exc)))
            continue

        store_name = path.stem
        results.append((store_name, path.name, total_sales))
        print(f"{store_name}: {total_sales:,.2f}")

    if not results:
        print("No files could be processed.")
        for filename, reason in failed:
            print(f"Skipped {filename}: {reason}")
        return 1

    grand_total = sum(total for _, _, total in results)
    print(f"Grand Total: {grand_total:,.2f}")

    output_path = Path(args.output).expanduser().resolve()
    write_summary_workbook(results, output_path)
    print(f"Saved summary workbook: {output_path}")

    if failed:
        print("\nSkipped files:")
        for filename, reason in failed:
            print(f"- {filename}: {reason}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
