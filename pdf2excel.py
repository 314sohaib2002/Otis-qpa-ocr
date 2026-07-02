"""
pdf2excel.py — Accurate PDF -> Excel table extraction for born-digital PDFs
(e.g. PDFs exported from Excel with a fixed template).

Approach (based on research):
  * Native text extraction with pdfplumber (NO OCR — the PDF text layer is
    lossless; OCR would only introduce errors on born-digital PDFs).
  * Two modes:
      - AUTO mode: detect tables from ruled lines ("lines" strategy).
      - TEMPLATE mode: a JSON template supplies fixed crop boxes and explicit
        cell-grid coordinates, which is deterministic and preserves empty and
        merged cells exactly. Use template_tool.py to build a template.
  * Validates that each page has a real text layer; scanned pages are
    reported instead of silently producing empty tables.
  * Writes one Excel sheet per extracted table with openpyxl.

Usage:
    python pdf2excel.py input.pdf                     # auto mode -> input.xlsx
    python pdf2excel.py input.pdf -o out.xlsx
    python pdf2excel.py input.pdf -t template.json    # template mode
"""

import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

MIN_CHARS_FOR_TEXT_LAYER = 20  # below this a page is treated as scanned/empty

EU_NUMBER = re.compile(r"^-?\d{1,3}(\.\d{3})*(,\d+)?$|^-?\d+(,\d+)?$")


def clean_cell(value, parse_numbers=True):
    """Normalize one extracted cell. Empty cells stay None (never dropped)."""
    if value is None:
        return None
    text = unicodedata.normalize("NFKC", str(value)).strip()
    if text == "":
        return None
    if parse_numbers and EU_NUMBER.match(text):
        # European format from Italian/German Excel exports: 1.016,72 -> 1016.72
        try:
            return float(text.replace(".", "").replace(",", "."))
        except ValueError:
            return text
    return text


def has_text_layer(page):
    text = page.extract_text() or ""
    return len(text.strip()) >= MIN_CHARS_FOR_TEXT_LAYER


def extract_auto(page):
    """Detect ruled-line tables on a page and return (bbox, grid) pairs."""
    results = []
    for table in page.find_tables(
        {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
    ):
        results.append((table.bbox, table.extract()))
    return results


def extract_with_template(pdf, template):
    """
    Extract each table region defined in the template.

    Template JSON format (all coordinates in PDF points, origin top-left):
    {
      "name": "otis-qpa-v1",
      "tables": [
        {
          "name": "line_items",
          "page": 1,
          "bbox": [x0, top, x1, bottom],
          "explicit_vertical_lines":   [x0, ..., x1],   # optional
          "explicit_horizontal_lines": [top, ..., bottom],  # optional
          "expected_headers": ["Item code", "Quantity"]     # optional check
        }
      ]
    }
    When explicit lines are given the cell grid is fixed — the most accurate
    option for a layout that never changes. Without them, ruled-line detection
    runs inside the cropped region only.
    """
    results = []
    for spec in template["tables"]:
        page = pdf.pages[spec["page"] - 1]
        region = page.crop(spec["bbox"])

        settings = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
        if spec.get("explicit_vertical_lines"):
            settings["vertical_strategy"] = "explicit"
            settings["explicit_vertical_lines"] = spec["explicit_vertical_lines"]
        if spec.get("explicit_horizontal_lines"):
            settings["horizontal_strategy"] = "explicit"
            settings["explicit_horizontal_lines"] = spec["explicit_horizontal_lines"]

        grid = region.extract_table(settings)
        if grid is None:
            raise ValueError(
                f"Table '{spec['name']}' not found on page {spec['page']} "
                f"in bbox {spec['bbox']} — has the template drifted?"
            )

        expected = spec.get("expected_headers")
        if expected:
            headers = [(c or "").strip() for c in grid[0]]
            missing = [h for h in expected
                       if not any(h in cell for cell in headers)]
            if missing:
                raise ValueError(
                    f"Validation failed for table '{spec['name']}': expected "
                    f"headers {missing} not found in {headers}. "
                    "The PDF layout may have changed — refusing to extract."
                )

        results.append((spec["name"], spec["bbox"], grid))
    return results


def write_excel(tables, out_path, parse_numbers=True):
    """tables: list of (sheet_name, bbox, grid). One sheet per table."""
    wb = Workbook()
    wb.remove(wb.active)
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")

    for name, bbox, grid in tables:
        ws = wb.create_sheet(title=name[:31])
        for r, row in enumerate(grid, start=1):
            for c, cell in enumerate(row, start=1):
                ws.cell(row=r, column=c, value=clean_cell(cell, parse_numbers))
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
        for c in range(1, ws.max_column + 1):
            width = max(
                (len(str(ws.cell(row=r, column=c).value or "").split("\n")[0])
                 for r in range(1, ws.max_row + 1)),
                default=8,
            )
            ws.column_dimensions[get_column_letter(c)].width = min(width + 2, 60)

    wb.save(out_path)


def convert(pdf_path, out_path=None, template_path=None, parse_numbers=True):
    pdf_path = Path(pdf_path)
    out_path = Path(out_path) if out_path else pdf_path.with_suffix(".xlsx")

    with pdfplumber.open(pdf_path) as pdf:
        scanned = [i + 1 for i, p in enumerate(pdf.pages) if not has_text_layer(p)]
        if scanned:
            print(
                f"WARNING: pages {scanned} have no usable text layer (scanned?). "
                "This tool extracts the native text layer and does not run OCR; "
                "those pages will be skipped.",
                file=sys.stderr,
            )

        if template_path:
            template = json.loads(Path(template_path).read_text(encoding="utf-8"))
            tables = extract_with_template(pdf, template)
        else:
            tables = []
            for i, page in enumerate(pdf.pages, start=1):
                if i in scanned:
                    continue
                for j, (bbox, grid) in enumerate(extract_auto(page), start=1):
                    tables.append((f"p{i}_table{j}", bbox, grid))

    if not tables:
        raise SystemExit("No tables found. Try building a template with template_tool.py.")

    write_excel(tables, out_path, parse_numbers)
    print(f"Extracted {len(tables)} table(s) -> {out_path}")
    for name, bbox, grid in tables:
        rows, cols = len(grid), max(len(r) for r in grid)
        print(f"  {name}: {rows} rows x {cols} cols  (bbox {[round(v,1) for v in bbox]})")
    return out_path


def main():
    ap = argparse.ArgumentParser(description="Extract tables from a born-digital PDF into Excel.")
    ap.add_argument("pdf", help="input PDF file")
    ap.add_argument("-o", "--output", help="output .xlsx path (default: alongside PDF)")
    ap.add_argument("-t", "--template", help="template JSON for fixed-layout extraction")
    ap.add_argument("--raw-text", action="store_true",
                    help="keep all cells as text (skip European-number parsing)")
    args = ap.parse_args()
    convert(args.pdf, args.output, args.template, parse_numbers=not args.raw_text)


if __name__ == "__main__":
    main()
