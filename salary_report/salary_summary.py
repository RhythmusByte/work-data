from __future__ import annotations

import argparse
import glob
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
from dateutil import parser as dateparser
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SCRIPT_DIR = Path(__file__).resolve().parent
HEADER_SCAN_LIMIT = 20
MAX_PREVIEW_COLUMNS = 6
NUMFMT = "#,##0.00"
TITLE_FILL = "FF1B365D"
LIGHT_FILL = "FFE6ECF4"
OUTPUT_PREFIX = "Salary_Summary_"
EXPECTED_HEADERS = ("Date", "Expense", "Amount", "Notes")

ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
INVALID_SHEET_CHARS = re.compile(r"[\[\]\:\*\?\/\\]")


@dataclass
class StoreReport:
    store_name: str
    source_file: str
    total_salary: float
    salary_rows: list[list[object]]
    period_start: datetime | None
    period_end: datetime | None
    sheet_name: str
    category_totals: dict[str, list] = field(default_factory=dict)
    month_totals: dict[str, list] = field(default_factory=dict)
    month_labels: dict[str, str] = field(default_factory=dict)


def normalize_text(value) -> str:
    return str(value or "").strip().lower()


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


def parse_any_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if not text:
        return None
    if ISO_DATE_RE.match(text):
        try:
            return datetime.strptime(text, "%Y-%m-%d")
        except ValueError:
            return None
    try:
        return dateparser.parse(text, dayfirst=True)
    except (ValueError, TypeError):
        return None


def find_header_row(ws) -> int:
    expected = [normalize_text(h) for h in EXPECTED_HEADERS]
    for row_idx in range(1, min(ws.max_row, HEADER_SCAN_LIMIT) + 1):
        values = [
            normalize_text(ws.cell(row_idx, col_idx).value)
            for col_idx in range(1, min(ws.max_column, MAX_PREVIEW_COLUMNS) + 1)
        ]
        if values[: len(expected)] == expected:
            return row_idx
    raise ValueError(f"Could not find a header row matching {EXPECTED_HEADERS}.")


def is_salary_row(row) -> bool:
    category = normalize_text(row[1])
    notes = normalize_text(row[3])
    return "salary" in category or "salary" in notes


def collect_files(file_args: list[str], output_name: str) -> list[Path]:
    if file_args:
        paths = [Path(arg).expanduser().resolve() for arg in file_args]
    else:
        try:
            from google.colab import files

            uploaded = files.upload()
            paths = [Path(name).resolve() for name in uploaded.keys()]
        except ImportError:
            paths = [Path(path).resolve() for path in glob.glob(str(SCRIPT_DIR / "*.xlsx"))]
            if not paths:
                paths = [Path(path).resolve() for path in glob.glob("*.xlsx")]

    result = []
    for path in paths:
        if not path.is_file():
            continue
        if path.suffix.lower() != ".xlsx":
            continue
        if path.name.startswith("~$"):
            continue
        if path.name.startswith(OUTPUT_PREFIX):
            continue
        if path.name == output_name:
            continue
        result.append(path)

    return sorted(result)


def prompt_store_names(paths: list[Path]) -> dict[Path, str]:
    print("Files received:")
    for path in paths:
        print(f"  - {path.name}")

    interactive = sys.stdin.isatty()
    used_names: set[str] = set()
    store_names: dict[Path, str] = {}

    for path in paths:
        default_name = path.stem.replace("_", " ").replace("-", " ").strip() or "Store"
        if interactive:
            while True:
                response = input(f"Store name for file {path.name} [{default_name}]: ").strip()
                candidate = " ".join((response or default_name).split())
                if candidate not in used_names:
                    break
                print(f"Store name '{candidate}' is already used. Choose a unique name.")
        else:
            candidate = default_name if default_name not in used_names else f"{default_name} {len(used_names) + 1}"
        used_names.add(candidate)
        store_names[path] = candidate

    return store_names


def make_unique_sheet_title(base_name: str, existing: set[str]) -> str:
    cleaned = INVALID_SHEET_CHARS.sub(" ", base_name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned) or "Store"
    suffix = " Summary"
    max_base_len = 31 - len(suffix)
    base = cleaned[:max_base_len].rstrip()
    candidate = f"{base}{suffix}"
    counter = 2
    while candidate in existing:
        extra = f" ({counter})"
        max_base_len = 31 - len(suffix) - len(extra)
        base = cleaned[:max_base_len].rstrip()
        candidate = f"{base}{suffix}{extra}"
        counter += 1
    existing.add(candidate)
    return candidate


def extract_store_report(path: Path, store_name: str, sheet_name: str) -> StoreReport:
    workbook = openpyxl.load_workbook(path, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    header_row = find_header_row(worksheet)

    salary_rows: list[list[object]] = []
    for row_idx in range(header_row + 1, worksheet.max_row + 1):
        row = [worksheet.cell(row_idx, col_idx).value for col_idx in range(1, 5)]
        if row[0] is None and row[1] is None and row[2] is None and row[3] is None:
            continue
        if is_salary_row(row):
            salary_rows.append([row[0], row[1], to_number(row[2]), row[3]])

    total_salary = sum((row[2] or 0) for row in salary_rows)
    dates = [parse_any_date(row[0]) for row in salary_rows]
    dates = [dt for dt in dates if dt is not None]
    period_start = min(dates) if dates else None
    period_end = max(dates) if dates else None

    category_totals: dict[str, list] = {}
    month_totals: dict[str, list] = {}
    month_labels: dict[str, str] = {}
    for row in salary_rows:
        category_name = str(row[1] or "").strip() or "Uncategorised"
        cat_entry = category_totals.setdefault(category_name, [0, 0.0])
        cat_entry[0] += 1
        cat_entry[1] += row[2] or 0

        row_date = parse_any_date(row[0])
        month_key = row_date.strftime("%Y-%m") if row_date else "Unknown"
        month_labels[month_key] = row_date.strftime("%B %Y") if row_date else "Unknown"
        month_entry = month_totals.setdefault(month_key, [0, 0.0])
        month_entry[0] += 1
        month_entry[1] += row[2] or 0

    return StoreReport(
        store_name=store_name,
        source_file=path.name,
        total_salary=total_salary,
        salary_rows=salary_rows,
        period_start=period_start,
        period_end=period_end,
        sheet_name=sheet_name,
        category_totals=category_totals,
        month_totals=month_totals,
        month_labels=month_labels,
    )


def _write_section_header(ws, row_idx: int, text: str) -> None:
    cell = ws.cell(row_idx, 2, text)
    cell.font = Font(name="Arial", size=11, bold=True, color=TITLE_FILL)


def _write_table_header(ws, row_idx: int, headers: list[str], start_col: int = 2) -> None:
    for offset, header in enumerate(headers):
        cell = ws.cell(row_idx, start_col + offset, header)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor=TITLE_FILL)
        cell.alignment = Alignment(horizontal="center")


def _write_total_row(ws, row_idx: int, start_col: int, end_col: int) -> None:
    for col_idx in range(start_col, end_col + 1):
        ws.cell(row_idx, col_idx).font = Font(name="Arial", size=10, bold=True)
        ws.cell(row_idx, col_idx).fill = PatternFill(fill_type="solid", fgColor=LIGHT_FILL)


def write_store_sheet(ws, report: StoreReport) -> None:
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 18

    ws["B2"] = f"{report.store_name} Salary Summary"
    ws["B2"].font = Font(name="Arial", size=14, bold=True, color=TITLE_FILL)
    ws["B3"] = f"Source file: {report.source_file}"
    ws["B3"].font = Font(name="Arial", size=10, color="FF555555")
    ws["B4"] = f"Salary rows matched: {len(report.salary_rows)}"
    ws["B4"].font = Font(name="Arial", size=10, color="FF555555")
    period_text = "Period: not available"
    if report.period_start and report.period_end:
        period_text = f"Period: {report.period_start.strftime('%d %B %Y')} to {report.period_end.strftime('%d %B %Y')}"
    ws["B5"] = period_text
    ws["B5"].font = Font(name="Arial", size=10, color="FF555555")
    ws["B6"] = f"Total salary given (INR): {report.total_salary:,.2f}"
    ws["B6"].font = Font(name="Arial", size=11, bold=True)

    row = 8

    # Monthly breakdown
    _write_section_header(ws, row, "Monthly Breakdown")
    row += 1
    _write_table_header(ws, row, ["Month", "Rows", "Total (INR)"])
    row += 1
    sorted_months = sorted(report.month_totals.items(), key=lambda item: item[0])
    for month_key, (rows_count, month_total) in sorted_months:
        ws.cell(row, 2, report.month_labels.get(month_key, month_key))
        ws.cell(row, 3, rows_count)
        ws.cell(row, 4, month_total)
        ws.cell(row, 4).number_format = NUMFMT
        row += 1
    if sorted_months:
        ws.cell(row, 2, "TOTAL")
        ws.cell(row, 3, len(report.salary_rows))
        ws.cell(row, 4, report.total_salary)
        ws.cell(row, 4).number_format = NUMFMT
        _write_total_row(ws, row, 2, 4)
        row += 1
    else:
        ws.cell(row, 2, "No salary rows were found in this file.")
        ws.cell(row, 2).font = Font(name="Arial", size=10, italic=True, color="FF888888")
        row += 1
    row += 1

    # Category breakdown
    _write_section_header(ws, row, "Category Breakdown")
    row += 1
    _write_table_header(ws, row, ["Category", "Rows", "Total (INR)"])
    row += 1
    sorted_categories = sorted(
        report.category_totals.items(), key=lambda item: item[1][1], reverse=True
    )
    for category_name, (rows_count, category_total) in sorted_categories:
        ws.cell(row, 2, category_name)
        ws.cell(row, 3, rows_count)
        ws.cell(row, 4, category_total)
        ws.cell(row, 4).number_format = NUMFMT
        row += 1
    if sorted_categories:
        ws.cell(row, 2, "TOTAL")
        ws.cell(row, 3, len(report.salary_rows))
        ws.cell(row, 4, report.total_salary)
        ws.cell(row, 4).number_format = NUMFMT
        _write_total_row(ws, row, 2, 4)
        row += 1
    else:
        ws.cell(row, 2, "No salary categories found.")
        ws.cell(row, 2).font = Font(name="Arial", size=10, italic=True, color="FF888888")
        row += 1
    row += 1

    # Detail table
    header_row = row
    headers = ["Date", "Expense Category", "Amount", "Notes"]
    _write_table_header(ws, header_row, headers)

    for row_idx, salary_row in enumerate(report.salary_rows, start=header_row + 1):
        ws.cell(row_idx, 2, parse_any_date(salary_row[0]) or salary_row[0])
        ws.cell(row_idx, 3, salary_row[1] or "Uncategorised")
        ws.cell(row_idx, 4, salary_row[2])
        ws.cell(row_idx, 5, salary_row[3])

    for r in ws.iter_rows(min_row=header_row + 1, min_col=2, max_col=2):
        r[0].number_format = "dd-mmm-yyyy"
    for r in ws.iter_rows(min_row=header_row + 1, min_col=4, max_col=4):
        r[0].number_format = NUMFMT

    if report.salary_rows:
        total_row = header_row + len(report.salary_rows) + 1
        ws.cell(total_row, 3, "TOTAL")
        ws.cell(total_row, 4, report.total_salary)
        ws.cell(total_row, 4).number_format = NUMFMT
        _write_total_row(ws, total_row, 3, 5)
    else:
        ws.cell(header_row + 1, 2, "No salary rows were found in this file.")
        ws.cell(header_row + 1, 2).font = Font(name="Arial", size=10, italic=True, color="FF888888")

    ws.freeze_panes = f"B{header_row + 1}"


def write_combined_summary_sheet(ws, reports: list[StoreReport]) -> None:
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 30
    for idx in range(len(reports)):
        ws.column_dimensions[get_column_letter(4 + idx)].width = 18
    ws.column_dimensions[get_column_letter(4 + len(reports))].width = 18

    ws["B2"] = "Salary Summary"
    ws["B2"].font = Font(name="Arial", size=14, bold=True, color=TITLE_FILL)

    all_dates = [dt for report in reports for dt in (report.period_start, report.period_end) if dt is not None]
    if all_dates:
        period_label = f"{min(all_dates).strftime('%d %B %Y')} to {max(all_dates).strftime('%d %B %Y')}"
    else:
        period_label = "not available"
    ws["B3"] = f"Period: {period_label}"
    ws["B3"].font = Font(name="Arial", size=10, color="FF555555")

    row = 5

    # Store totals table
    _write_section_header(ws, row, "Store Totals")
    row += 1
    _write_table_header(ws, row, ["Store Name", "Source File", "Salary Rows", "Total Salary (INR)"])
    row += 1
    for report in reports:
        ws.cell(row, 2, report.store_name)
        ws.cell(row, 3, report.source_file)
        ws.cell(row, 4, len(report.salary_rows))
        ws.cell(row, 5, report.total_salary)
        ws.cell(row, 5).number_format = NUMFMT
        row += 1
    ws.cell(row, 2, "TOTAL")
    ws.cell(row, 4, sum(len(report.salary_rows) for report in reports))
    ws.cell(row, 5, sum(report.total_salary for report in reports))
    ws.cell(row, 5).number_format = NUMFMT
    _write_total_row(ws, row, 2, 5)
    row += 2

    # Monthly totals matrix: Month | Store1 | Store2 | ... | Grand Total
    month_keys: list[str] = []
    seen_months: set[str] = set()
    month_grand_totals: dict[str, float] = {}
    month_display_labels: dict[str, str] = {}
    for report in reports:
        for month_key, (_, month_total) in report.month_totals.items():
            month_grand_totals[month_key] = month_grand_totals.get(month_key, 0.0) + month_total
            month_display_labels[month_key] = report.month_labels.get(month_key, month_key)
            if month_key not in seen_months:
                seen_months.add(month_key)
                month_keys.append(month_key)
    month_keys.sort()

    _write_section_header(ws, row, "Monthly Totals")
    row += 1
    month_header_row = row
    ws.cell(month_header_row, 2, "Month")
    for idx, report in enumerate(reports):
        ws.cell(month_header_row, 4 + idx, report.store_name)
    month_grand_col = 4 + len(reports)
    ws.cell(month_header_row, month_grand_col, "Grand Total (INR)")
    _write_table_header(ws, month_header_row, [ws.cell(month_header_row, c).value for c in range(2, month_grand_col + 1)])
    row += 1
    for month_key in month_keys:
        ws.cell(row, 2, month_display_labels[month_key])
        for idx, report in enumerate(reports):
            _, month_total = report.month_totals.get(month_key, (0, 0.0))
            ws.cell(row, 4 + idx, month_total)
            ws.cell(row, 4 + idx).number_format = NUMFMT
        ws.cell(row, month_grand_col, month_grand_totals[month_key])
        ws.cell(row, month_grand_col).number_format = NUMFMT
        row += 1
    ws.cell(row, 2, "TOTAL")
    for idx, report in enumerate(reports):
        ws.cell(row, 4 + idx, report.total_salary)
        ws.cell(row, 4 + idx).number_format = NUMFMT
    ws.cell(row, month_grand_col, sum(month_grand_totals.values()))
    ws.cell(row, month_grand_col).number_format = NUMFMT
    _write_total_row(ws, row, 2, month_grand_col)
    row += 2

    # Category totals matrix: Category | Store1 | Store2 | ... | Grand Total
    category_names: list[str] = []
    seen_categories: set[str] = set()
    category_grand_totals: dict[str, float] = {}
    for report in reports:
        for category_name, (_, category_total) in report.category_totals.items():
            category_grand_totals[category_name] = category_grand_totals.get(category_name, 0.0) + category_total
            if category_name not in seen_categories:
                seen_categories.add(category_name)
                category_names.append(category_name)
    category_names.sort(key=lambda name: category_grand_totals[name], reverse=True)

    _write_section_header(ws, row, "Category Totals")
    row += 1
    cat_header_row = row
    ws.cell(cat_header_row, 2, "Category")
    for idx, report in enumerate(reports):
        ws.cell(cat_header_row, 4 + idx, report.store_name)
    cat_grand_col = 4 + len(reports)
    ws.cell(cat_header_row, cat_grand_col, "Grand Total (INR)")
    _write_table_header(ws, cat_header_row, [ws.cell(cat_header_row, c).value for c in range(2, cat_grand_col + 1)])
    row += 1
    for category_name in category_names:
        ws.cell(row, 2, category_name)
        for idx, report in enumerate(reports):
            _, category_total = report.category_totals.get(category_name, (0, 0.0))
            ws.cell(row, 4 + idx, category_total)
            ws.cell(row, 4 + idx).number_format = NUMFMT
        ws.cell(row, cat_grand_col, category_grand_totals[category_name])
        ws.cell(row, cat_grand_col).number_format = NUMFMT
        row += 1
    ws.cell(row, 2, "TOTAL")
    for idx, report in enumerate(reports):
        ws.cell(row, 4 + idx, report.total_salary)
        ws.cell(row, 4 + idx).number_format = NUMFMT
    ws.cell(row, cat_grand_col, sum(category_grand_totals.values()))
    ws.cell(row, cat_grand_col).number_format = NUMFMT
    _write_total_row(ws, row, 2, cat_grand_col)

    ws.freeze_panes = "D6"


def build_workbook(reports: list[StoreReport], output_path: Path) -> None:
    workbook = openpyxl.Workbook()
    summary_ws = workbook.active
    summary_ws.title = "Salary Summary"
    write_combined_summary_sheet(summary_ws, reports)

    for report in reports:
        ws = workbook.create_sheet(report.sheet_name)
        write_store_sheet(ws, report)

    workbook.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a salary report with one summary sheet and one sheet per store."
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="Excel files to process. If omitted, all .xlsx files in the salary_report directory are used.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Optional output workbook name. Defaults to Salary_Summary_<period>.xlsx.",
    )
    args = parser.parse_args()

    output_name = Path(args.output).name if args.output else ""
    paths = collect_files(args.files, output_name)
    if not paths:
        print("No Excel files found.")
        return 1

    store_names = prompt_store_names(paths)
    sheet_names: set[str] = set()
    reports: list[StoreReport] = []

    for path in paths:
        store_name = store_names[path]
        sheet_name = make_unique_sheet_title(store_name, sheet_names)
        report = extract_store_report(path, store_name, sheet_name)
        reports.append(report)
        print(f"{store_name} ({path.name}): {report.total_salary:,.2f}")

    all_dates = [dt for report in reports for dt in (report.period_start, report.period_end) if dt is not None]
    if all_dates:
        period_start = min(all_dates)
        period_end = max(all_dates)
        default_output = f"{OUTPUT_PREFIX}{period_start.strftime('%Y-%m-%d')}_to_{period_end.strftime('%Y-%m-%d')}.xlsx"
    else:
        default_output = f"{OUTPUT_PREFIX}report.xlsx"

    output_path = Path(args.output or (SCRIPT_DIR / default_output)).expanduser().resolve()
    build_workbook(reports, output_path)

    grand_total = sum(report.total_salary for report in reports)
    print(f"Grand Total Salary: {grand_total:,.2f}")
    print(f"Saved summary workbook: {output_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())