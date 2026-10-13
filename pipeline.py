"""Shared pipeline used by the UI: PDFs in, compiled Admin/Absence/Overtime report out."""

import contextlib
import io
import sys
import tempfile
from pathlib import Path

import pandas as pd

from mapper import load_templates, process_timesheet
from pdf2excel import convert as pdf_to_excel

# When packaged with PyInstaller the templates are unpacked next to the code in _MEIPASS
_BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
TEMPLATES_DIR = _BASE_DIR / "templates"


def build_report(files, expected_month=None):
    """files: list of (file_name, pdf_bytes). Returns (report, results).

    report is {"Admin": df, "Absence": df, "Overtime": df}; results is one dict per
    file with its name, status ("OK" / "Rejected") and a message. A file that fails
    is skipped and never contributes rows to the report.
    """
    df_admin, df_absence, df_overtime = load_templates(TEMPLATES_DIR)
    results = []

    with tempfile.TemporaryDirectory() as tmp:
        for i, (name, data) in enumerate(files):
            try:
                pdf_path = Path(tmp) / f"{i}.pdf"
                pdf_path.write_bytes(data)
                # A windowed .exe has no console, so keep library output off stdout/stderr
                with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                    xlsx_path = pdf_to_excel(pdf_path)
                # Map into copies so a failure half-way never leaves partial rows behind
                df_admin, df_absence, df_overtime = process_timesheet(
                    xlsx_path, df_admin, df_absence, df_overtime, expected_month
                )
                results.append({"File": name, "Status": "OK", "Message": ""})
            except (Exception, SystemExit) as e:
                results.append({"File": name, "Status": "Rejected", "Message": str(e)})

    report = {"Admin": df_admin, "Absence": df_absence, "Overtime": df_overtime}
    return report, results


def report_to_xlsx_bytes(report):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in report.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()
