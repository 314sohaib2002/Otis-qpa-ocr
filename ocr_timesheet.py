"""OCR for scanned monthly timesheets (printed, signed and scanned: no text layer).

The report has a fixed table layout, so instead of OCR-ing the whole page the grid is
located from its ruled lines and every cell is read on its own (digits only for numeric
cells). The result is written in the same layout pdf2excel.py produces for the
born-digital PDFs, so mapper.py handles both the same way.

Because OCR can misread, the values are cross-checked against the totals printed on the
sheet (sum of the days = SUM row, remaining holiday, ...). A sheet that fails a check is
rejected instead of being mapped with wrong numbers.

Needs the Tesseract program (https://github.com/UB-Mannheim/tesseract/wiki).
"""

import calendar
import difflib
import re
import shutil
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import cv2
import numpy as np
import openpyxl
import pypdfium2 as pdfium
import pytesseract

RENDER_SCALE = 4  # 4x = ~288 dpi
GRID_COLUMNS = 12  # 11 table columns A-K -> 12 vertical lines
FIRST_ROWS = 5  # lines above the first day row: title, name, header(2 lines merged), carry-over
BLOCK_ROWS = 8  # lines from the SUM row down to the bottom of the sheet
MONTHS = ["Jan", "Feb", "Mar", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Des"]
OCR_WORKERS = 6  # parallel Tesseract processes
TOLERANCE = 0.05  # hours; the sheets print totals with 1-2 decimals

_WINDOWS_TESSERACT = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")


class OcrError(ValueError):
    """The scan could not be read reliably."""


def _find_tesseract():
    found = shutil.which("tesseract")
    if found:
        return found
    if _WINDOWS_TESSERACT.exists():
        return str(_WINDOWS_TESSERACT)
    raise OcrError(
        "This PDF is a scan and needs OCR, but Tesseract is not installed "
        "(https://github.com/UB-Mannheim/tesseract/wiki)."
    )


# --------------------------------------------------------------------------
# Page -> grid
# --------------------------------------------------------------------------
def _render(pdf_path):
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        return np.array(pdf[0].render(scale=RENDER_SCALE).to_pil().convert("L"))
    finally:
        pdf.close()


def _deskew(gray):
    dark = (gray < 110).astype(np.uint8) * 255
    horizontal = cv2.morphologyEx(dark, cv2.MORPH_OPEN,
                                  cv2.getStructuringElement(cv2.MORPH_RECT, (150, 1)))
    lines = cv2.HoughLinesP(horizontal, 1, np.pi / 3600, 150, minLineLength=500, maxLineGap=30)
    if lines is None:
        return gray
    angle = float(np.median([np.degrees(np.arctan2(y2 - y1, x2 - x1))
                             for x1, y1, x2, y2 in lines[:, 0]]))
    if abs(angle) < 0.05 or abs(angle) > 5:
        return gray
    h, w = gray.shape
    matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1)
    return cv2.warpAffine(gray, matrix, (w, h), flags=cv2.INTER_CUBIC, borderValue=255)


def _group(indices, gap=8):
    if len(indices) == 0:
        return []
    out, start, prev = [], indices[0], indices[0]
    for x in indices[1:]:
        if x - prev > gap:
            out.append(int((start + prev) // 2))
            start = x
        prev = x
    out.append(int((start + prev) // 2))
    return out


def _find_grid(gray):
    dark = (gray < 100).astype(np.uint8) * 255
    vertical = cv2.morphologyEx(dark, cv2.MORPH_OPEN,
                                cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40)))
    horizontal = cv2.morphologyEx(dark, cv2.MORPH_OPEN,
                                  cv2.getStructuringElement(cv2.MORPH_RECT, (120, 1)))
    xs = _group(np.where(vertical.sum(0) > 255 * 120)[0])
    ys = _group(np.where(horizontal.sum(1) > 255 * 500)[0])
    if len(xs) != GRID_COLUMNS:
        raise OcrError(f"Could not recognise the table: found {len(xs)} column lines "
                       f"(expected {GRID_COLUMNS}). Is the scan sharp and upright?")
    days = len(ys) - FIRST_ROWS - BLOCK_ROWS - 1
    if not 28 <= days <= 31:
        raise OcrError(f"Could not recognise the table rows ({days} day rows found, "
                       "expected 28-31). Is the whole page in the scan?")
    return xs, ys, days


# --------------------------------------------------------------------------
# Cell reading
# --------------------------------------------------------------------------
def _drop_line_fragments(ink):
    """Remove pieces of the neighbouring ruled lines that reach into the crop: thin
    components that run (almost) the full height of the crop."""
    count, labels, stats, _ = cv2.connectedComponentsWithStats(ink.astype(np.uint8), connectivity=8)
    height = ink.shape[0]
    keep = ink.copy()
    for i in range(1, count):
        width_i, height_i = stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT]
        if height_i >= 0.9 * height and width_i <= 12:
            keep[labels == i] = False
    return keep


def _cell_image(gray, x0, x1, y0, y1, scale=1.5, thin=0):
    margin = 6  # stay clear of the ruled lines
    crop = gray[y0 + margin:y1 - margin, x0 + margin:x1 - margin]
    if crop.size == 0:
        return None
    # Text is black or dark red; cell fills (blue/pink) are lighter
    ink = crop < 105
    if ink.sum() < 25:
        return None  # empty cell
    ink = _drop_line_fragments(ink)
    if ink.sum() < 40:
        return None
    if thin:  # heavily inked glyphs (a bold '1') become blobs; thinning restores the shape
        ink = cv2.erode(ink.astype(np.uint8), np.ones((3, 3), np.uint8), iterations=thin) > 0
    out = np.where(ink, 0, 255).astype(np.uint8)
    out = cv2.resize(out, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    return cv2.copyMakeBorder(out, 25, 25, 25, 25, cv2.BORDER_CONSTANT, value=255)


def _read(image, whitelist=None, psm=7):
    if image is None:
        return ""
    config = f"--psm {psm}"
    if whitelist:
        config += f" -c tessedit_char_whitelist={whitelist}"
    return pytesseract.image_to_string(image, config=config).strip()


def _text(image, psm=7):
    """Free text (names, comments) without the stray '|' / '[' the cell borders cause."""
    return " ".join(_read(image, psm=psm).split()).strip("|[]_!. ")


_TWO_DECIMALS = re.compile(r"^-?\d+,\d\d$|^-+$")
_VARIANTS = ((1.5, 0), (3.0, 0), (3.0, 1), (3.0, 2), (4.5, 1))  # (scale, thinning)


def _clean_number(text):
    """Keep the leading number-shaped part of an OCR result ('.0.00' -> '0.00',
    '7,501' -> '7,50'); printed numbers have at most 2 decimals."""
    text = text.replace(" ", "").lstrip(".,")
    match = re.match(r"-?\d+(?:[.,]\d{1,2})?|-+", text)
    text = match.group() if match else text
    return "" if not text.strip(".,") else text


def _number_text(gray, x0, x1, y0, y1, two_decimals=False):
    """Text of a numeric cell: '7,50', '-', '' ...

    Retries with other scales/thinning when Tesseract returns nothing. With
    two_decimals (the day columns always print 0,00 / 7,50 ...) every read is voted on,
    because a single read can be wrong yet well-formed (0,00 read as 0,90)."""
    votes, order = {}, []
    reads = 0
    for scale, thin in _VARIANTS:
        image = _cell_image(gray, x0, x1, y0, y1, scale, thin)
        if image is None:
            return ""
        for psm in (7, 8, 10):
            text = _clean_number(_read(image, "0123456789,.-", psm))
            if not text:
                continue
            if not two_decimals:
                return text  # first usable read (callers cross-check the totals)
            reads += 1
            normal = text.replace(".", ",")
            if _TWO_DECIMALS.match(normal):
                if normal not in votes:
                    order.append(normal)
                votes[normal] = votes.get(normal, 0) + 1
        # Stop early once one reading clearly wins
        if votes and max(votes.values()) >= 3 and max(votes.values()) > sum(votes.values()) / 2:
            break
    if not votes:
        return ""
    return max(order, key=lambda t: votes[t])  # ties -> the earliest reading


def _looks_like_one(gray, x0, x1, y0, y1):
    """True when the cell holds a single tall, narrow mark: a bold printed '1' that
    Tesseract cannot recognise on its own (the sick/vacation columns only hold 1)."""
    margin = 6
    crop = gray[y0 + margin:y1 - margin, x0 + margin:x1 - margin]
    if crop.size == 0:
        return False
    ink = _drop_line_fragments(crop < 105)
    count, _, stats, _ = cv2.connectedComponentsWithStats(ink.astype(np.uint8), connectivity=8)
    blobs = [stats[i] for i in range(1, count) if stats[i, cv2.CC_STAT_AREA] >= 60 and stats[i, cv2.CC_STAT_WIDTH] >= 4]
    if len(blobs) != 1:
        return False
    width, height = blobs[0][cv2.CC_STAT_WIDTH], blobs[0][cv2.CC_STAT_HEIGHT]
    return height >= 14 and width <= 0.55 * height


def _to_number(text):
    """7,50 -> 7.5; '-' / '' -> None; unreadable -> raises ValueError."""
    if text in ("", "-", "--"):
        return None
    cleaned = text.replace(",", ".").strip(".-")
    if cleaned.count(".") > 1:
        raise ValueError(text)
    return float(cleaned)


def _month(gray, x0, x1, y0, y1):
    image = _cell_image(gray, x0, x1, y0, y1)
    if image is None:
        return ""
    for psm in (7, 8, 13):
        text = _read(image, "ADEFGJLMNOPSTUVYabceghiklmnoprstuvy", psm).capitalize()[:3]
        if text in MONTHS:
            return text
    # A single clearly-best fuzzy match is accepted (e.g. 'Mal' for 'Mai' is NOT: ties)
    close = difflib.get_close_matches(text, MONTHS, n=2, cutoff=0.6)
    if len(close) == 1:
        return close[0]
    raise OcrError(f"Could not read the month on the sheet (read '{text}').")


# --------------------------------------------------------------------------
# Whole sheet
# --------------------------------------------------------------------------
def read_scanned_timesheet(pdf_path):
    """Returns (rows, problems): the sheet as a list of rows (A1 = rows[0][0]) in the
    pdf2excel layout, and a list of failed cross-checks (empty when all is consistent)."""
    pytesseract.pytesseract.tesseract_cmd = _find_tesseract()
    gray = _deskew(_render(pdf_path))
    xs, ys, days = _find_grid(gray)

    def cell(col, row, to_col=None):
        """Pixel box of the cell in table column `col` (0=A) of grid row `row`."""
        return xs[col], xs[(to_col if to_col is not None else col) + 1], ys[row], ys[row + 1]

    rows = {}  # (row, col) -> value, 1-based Excel coordinates

    def put(row, col, value):
        if value is not None and value != "":
            rows[(row, col + 1)] = value

    # --- header ---
    name_img = _cell_image(gray, *cell(1, 1, 3))
    put(2, 0, "NAVN:")
    put(2, 1, _text(name_img))
    put(2, 5, int(_number_text(gray, *cell(5, 1, 6)) or 0) or None)
    put(2, 8, _month(gray, *cell(8, 1)))
    year = _number_text(gray, *cell(10, 1)).strip(".,-")
    put(2, 10, int(year) if year.isdigit() else None)
    problems = []

    def number(text, where):
        try:
            return _to_number(text)
        except ValueError:
            problems.append(f"unreadable number '{text}' at {where}")
            return None

    carry = _number_text(gray, *cell(4, 4))
    put(5, 4, number(carry, "E5") if carry not in ("", "-") else carry or None)
    put(5, 5, "Gjenværende feriedager")
    holiday_left = number(_number_text(gray, *cell(10, 4)), "K5")
    put(5, 10, holiday_left)

    # --- one row per day (rows are independent, so read them in parallel) ---
    day_columns = ((1, "B"), (2, "C"), (4, "E"), (5, "F"), (6, "G"), (7, "H"), (8, "I"),
                   (9, "J"), (10, "K"))

    def read_day(d):
        grid_row = 4 + d  # ys[4 + d] .. ys[5 + d]
        values = {}
        for col, letter in day_columns:
            box = cell(col, grid_row)
            text = _number_text(gray, *box, two_decimals=letter not in ("J", "K"))
            if not text and letter in ("J", "K") and _looks_like_one(gray, *box):
                text = "1"
            values[col] = text
        comment = _text(_cell_image(gray, *cell(3, grid_row)), psm=6) or None
        return values, comment

    sums = {"E": 0.0, "J": 0.0, "K": 0.0}
    with ThreadPoolExecutor(max_workers=OCR_WORKERS) as pool:
        day_results = list(pool.map(read_day, range(1, days + 1)))
    for d, (values, comment) in enumerate(day_results, start=1):
        excel_row = FIRST_ROWS + d
        put(excel_row, 0, d)
        for col, letter in day_columns:
            text = values[col]
            if letter not in ("J", "K") and text and not _TWO_DECIMALS.match(text.replace(".", ",")):
                problems.append(f"{letter}{excel_row} read as '{text}' (expected e.g. 7,50)")
            value = number(text, f"{letter}{excel_row}")
            put(excel_row, col, value if value is not None else (text or None))
            if letter in sums and value:
                sums[letter] += value
        put(excel_row, 3, comment)

    # --- totals block (SUM row and the five rows below it) ---
    sum_grid = 4 + days + 1
    sum_row = FIRST_ROWS + days + 1
    put(sum_row, 0, "SUM denne måned")
    totals = {}
    for col, letter in ((4, "E"), (7, "H"), (8, "I"), (9, "J"), (10, "K")):
        text = _number_text(gray, *cell(col, sum_grid))
        totals[letter] = number(text, f"{letter}{sum_row}")
        put(sum_row, col, totals[letter] if totals[letter] is not None else (text or None))
    for offset, label in ((1, "Sum overtid denne måneden"), (2, "Overtidstimer til avspasering"),
                          (3, "Overtidstimer til utbetaling"),
                          (4, "Time saldo overføres til neste måned Månedslønnede")):
        put(sum_row + offset, 0, label)
        text = _number_text(gray, *cell(4, sum_grid + offset))
        put(sum_row + offset, 4, number(text, f"E{sum_row + offset}")
            if text not in ("", "-") else (text or None))
    put(sum_row + 5, 5, "Restferie:")
    remaining = number(_number_text(gray, *cell(10, sum_grid + 5)), f"K{sum_row + 5}")
    put(sum_row + 5, 10, remaining)

    # --- cross-checks against the totals printed on the sheet ---
    if totals["E"] is None or abs(sums["E"] - totals["E"]) > TOLERANCE:
        problems.append(f"worked hours: days add up to {sums['E']:g}, sheet says {totals['E']}")
    for letter, label in (("J", "sick days"), ("K", "vacation days")):
        if abs(sums[letter] - (totals[letter] or 0)) > TOLERANCE:
            problems.append(f"{label}: days add up to {sums[letter]:g}, "
                            f"sheet says {totals[letter]}")
    if holiday_left is not None and remaining is not None and \
            abs(holiday_left - (totals["K"] or 0) - remaining) > TOLERANCE:
        problems.append(f"remaining holiday: {holiday_left:g} - {totals['K'] or 0:g} "
                        f"taken should be {holiday_left - (totals['K'] or 0):g}, "
                        f"sheet says {remaining:g}")
    year_cell, month_cell = rows.get((2, 11)), rows.get((2, 9))
    if year_cell and month_cell in MONTHS:
        expected = calendar.monthrange(year_cell, MONTHS.index(month_cell) + 1)[1]
        if expected != days:
            problems.append(f"{month_cell} {year_cell} has {expected} days but the sheet "
                            f"has {days} day rows")

    last_row = sum_row + 6
    matrix = [[rows.get((r, c)) for c in range(1, 12)] for r in range(1, last_row + 1)]
    return matrix, problems


def scanned_pdf_to_excel(pdf_path, out_path):
    """OCR a scanned timesheet into an .xlsx laid out like pdf2excel's output.
    Raises OcrError (with the failed checks) when the values are not consistent."""
    matrix, problems = read_scanned_timesheet(pdf_path)
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in matrix:
        ws.append(row)
    wb.save(out_path)
    if problems:
        raise OcrError("OCR result failed its checks - " + "; ".join(problems))
    return Path(out_path)
