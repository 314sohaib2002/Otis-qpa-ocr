# Otis-qpa-ocr

Accurate PDF → Excel table extraction for born-digital PDFs (PDFs exported
from Excel with a fixed template). Despite the repo name, **no OCR is used for them**:
these PDFs carry a lossless embedded text layer, and research shows native
extraction reaches ~100% accuracy where OCR tops out around 97–99%. The tool
reads the text layer directly with pdfplumber and writes Excel with openpyxl.

## Setup

```sh
pip install -r requirements.txt
```

## Desktop app (timesheet mapper)

```sh
python gui.py
```

Add the monthly report PDFs, pick the month from the dropdown and click
**Process**. Each file is mapped into the built-in Admin / Absence / Overtime
templates; files for a different month, or that fail validation, are rejected
and listed with the reason. **Export to Excel...** saves the compiled report.
The same pipeline is available from the command line with
`python main.py <folder>`.

### Build the .exe

```sh
pip install -r requirements.txt -r requirements-build.txt
pyinstaller --noconfirm --onefile --windowed --name OtisTimesheetMapper ^
  --add-data "templates/Admin.xlsx;templates" ^
  --add-data "templates/Absense.xlsx;templates" ^
  --add-data "templates/overtime.xlsx;templates" ^
  --collect-all pdfminer --collect-all pdfplumber gui.py
```

The result is `dist/OtisTimesheetMapper.exe` (the templates are bundled inside).
`OtisTimesheetMapper.exe --selftest <month-number> <out.txt> <pdf>...` maps the
PDFs without opening a window and writes a summary, to check a build.

## Scanned timesheets (OCR)

Timesheets that were printed, signed and scanned have no text layer. They are read
with OCR (`ocr_timesheet.py`): the table grid is located from its ruled lines, every
cell is read on its own, and the result is cross-checked against the totals printed on
the sheet (hours, sick days, vacation days, remaining holiday). A sheet that fails a
check is rejected instead of being mapped with wrong numbers. Names and comments are
free text and may contain OCR typos - use **Export extracted data** to review them.

This needs the Tesseract program installed
(<https://github.com/UB-Mannheim/tesseract/wiki>, default path
`C:\Program Files\Tesseract-OCR`). It is not bundled in the .exe. A scan takes about a
minute.

## Quick start (auto mode)

Detects tables from ruled lines and writes one Excel sheet per table:

```sh
python pdf2excel.py "invoice.pdf"                # -> invoice.xlsx
python pdf2excel.py "invoice.pdf" -o out.xlsx
```

## Accurate mode (fixed-layout template) — recommended

Because the PDF layout never changes, measure the table geometry once and
freeze it into a template. Extraction then becomes deterministic: every cell
(including empty and merged ones) lands in the same position every time, and
the run fails loudly if the layout ever drifts.

1. Generate a starter template + debug images from one sample PDF:

   ```sh
   python template_tool.py "sample.pdf" -o qpa_template.json
   ```

2. Open the `*_debug.png` images to confirm the detected cell grid. Edit the
   JSON: rename tables, delete header/footer regions you don't need, and trim
   `expected_headers` down to labels that never change (they are used as a
   layout-drift check on every run).

3. Extract any future PDF with the frozen template:

   ```sh
   python pdf2excel.py "new_document.pdf" -t qpa_template.json -o out.xlsx
   ```

## Notes

- European number formats (`1.016,72`) are converted to real numbers; use
  `--raw-text` to keep everything as literal text.
- Pages with no text layer (true scans) are skipped with a warning — this
  tool deliberately does not OCR them, since OCR would introduce errors on
  the born-digital documents this project targets.
- Multi-line cells from the PDF are kept as multi-line text in Excel.
