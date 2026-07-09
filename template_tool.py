"""
template_tool.py — helper for building extraction templates for pdf2excel.py.

Because the PDF layout is fixed, you measure the table geometry ONCE and
freeze it into a template JSON. This tool does the measuring for you:

  1. Renders each page with detected table grids drawn on top
     (page_N_debug.png) so you can visually confirm the cell boundaries.
  2. Prints/saves a starter template JSON with each detected table's bbox and
     the exact ruled-line coordinates, ready to paste into a template file
     and hand-tune.

Usage:
    python template_tool.py input.pdf
    python template_tool.py input.pdf -o my_template.json
"""

import argparse
import json
from pathlib import Path

import pdfplumber


def build_template(pdf_path, out_json=None, save_images=True):
    pdf_path = Path(pdf_path)
    template = {"name": pdf_path.stem, "tables": []}

    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            tables = page.find_tables(
                {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
            )
            for idx, table in enumerate(tables, start=1):
                x0, top, x1, bottom = table.bbox
                # Reconstruct the cell grid lines from the table's cells so the
                # template pins the exact geometry, not just the outer box.
                xs = sorted({round(c[0], 2) for c in table.cells}
                            | {round(c[2], 2) for c in table.cells})
                ys = sorted({round(c[1], 2) for c in table.cells}
                            | {round(c[3], 2) for c in table.cells})
                # Derive headers from the SAME explicit grid the template will
                # use, so pdf2excel.py's validation is self-consistent.
                grid = page.crop(table.bbox).extract_table({
                    "vertical_strategy": "explicit",
                    "horizontal_strategy": "explicit",
                    "explicit_vertical_lines": xs,
                    "explicit_horizontal_lines": ys,
                })
                headers = [(c or "").strip().split("\n")[0] for c in grid[0]] if grid else []
                template["tables"].append({
                    "name": f"p{page_no}_table{idx}",
                    "page": page_no,
                    "bbox": [round(x0, 2), round(top, 2), round(x1, 2), round(bottom, 2)],
                    "explicit_vertical_lines": xs,
                    "explicit_horizontal_lines": ys,
                    "expected_headers": [h for h in headers if h],
                })

            if save_images:
                img = page.to_image(resolution=120)
                img.debug_tablefinder(
                    {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
                )
                img_path = pdf_path.parent / f"{pdf_path.stem}_page{page_no}_debug.png"
                img.save(img_path)
                print(f"debug image: {img_path}")

    out_json = Path(out_json) if out_json else pdf_path.with_suffix(".template.json")
    out_json.write_text(json.dumps(template, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"template written: {out_json}  ({len(template['tables'])} table(s))")
    print("Review the debug images, rename tables, delete regions you don't need,")
    print("and trim expected_headers to the labels that never change.")
    return out_json


def main():
    ap = argparse.ArgumentParser(description="Generate a starter extraction template from a sample PDF.")
    ap.add_argument("pdf", help="sample PDF with the fixed layout")
    ap.add_argument("-o", "--output", help="output template JSON path")
    ap.add_argument("--no-images", action="store_true", help="skip debug PNG rendering")
    args = ap.parse_args()
    build_template(args.pdf, args.output, save_images=not args.no_images)


if __name__ == "__main__":
    main()
