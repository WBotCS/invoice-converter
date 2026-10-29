#!/usr/bin/env python3
"""
pdf_to_invoice.py - Convert a scanned MLL Services timesheet PDF into the
Excel invoice format (one invoice workbook per PDF page).

The Excel template (e.g. MLL_UNWRAPPED_INVOICE_UN-B-79.xlsx) is used as-is:
logo, "INVOICE" banner, address boxes, styles, Employee List sheet and the
dropdown are all kept. Only the header fields (date, invoice no, start/end
date) and the employee table + totals are rewritten, with live formulas.

Because the PDF is a scan, the numbers are read with OCR. Every row is then
cross-checked against the totals printed on the PDF itself (Reg-Hrly pay,
Straight Time, Total Pay, column TOTALS, Total due) and anything that does
not add up is reported so you can fix it before sending.

Usage:
    Keep template.xlsx in the same folder as this script, then:

    python pdf_to_invoice.py SCAN.pdf --invoice-no UN-B-80
    python pdf_to_invoice.py SCAN.pdf --invoice-no UN-B-80 \
        --invoice-date 2026-10-08 --pages 2 --out-dir invoices

Requirements:
    pip install pymupdf opencv-python-headless pytesseract numpy
    + the Tesseract OCR program itself
      (Windows: https://github.com/UB-Mannheim/tesseract/wiki ;
       Mac: brew install tesseract ; Linux: apt install tesseract-ocr)
"""
import argparse
import datetime as dt
import difflib
import os
import re
import sys
import zipfile
from concurrent.futures import ThreadPoolExecutor
from xml.sax.saxutils import escape

import cv2
import numpy as np
import pymupdf
import pytesseract

DPI = 300
TOL = 0.015  # money tolerance when cross-checking

# Columns on the PDF, in order, starting at "Reg HRS".
PDF_COLS = ["reg", "sick", "holiday", "adj_hrs", "ot_half_hrs", "pcs", "reg_pay",
            "bonus", "misc", "straight", "adj_amt", "adj_ot", "ot_pay", "total",
            "bill", "ot_rate"]
PCS_OFFSET = PDF_COLS.index("pcs")
TOTAL_OFFSET = PDF_COLS.index("total")

# Template layout (rows of the original 38-employee template)
FIRST_DATA_ROW = 30
TEMPLATE_LAST_DATA_ROW = 67
TEMPLATE_TOTAL_ROW = 68


# --------------------------------------------------------------------------
# Image / OCR
# --------------------------------------------------------------------------
def render_pages(pdf_path):
    """Render each page as grayscale at 300 DPI. Red marks (e.g. a red pen
    stroke or a red scanner streak across the sheet) are removed first: the
    timesheet itself is printed in black, so anything clearly red is
    not part of it and would only confuse the table reader."""
    doc = pymupdf.open(pdf_path)
    for page in doc:
        pix = page.get_pixmap(dpi=DPI, colorspace=pymupdf.csRGB)
        rgb = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.stride)[:, :pix.w * 3]
        rgb = rgb.reshape(pix.h, pix.w, 3).astype(np.int32)
        r, g, b = rgb[..., 0], rgb[..., 1], rgb[..., 2]
        gray = ((r * 299 + g * 587 + b * 114) // 1000).astype(np.uint8)
        red = (r - np.maximum(g, b) > 50) & (r > 110)
        gray[red] = 255
        yield gray

def line_masks(img):
    b = cv2.adaptiveThreshold(~img, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2)
    h, w = b.shape
    hl = cv2.morphologyEx(b, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (w // 30, 1)))
    vl = cv2.morphologyEx(b, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, h // 60)))
    return hl, vl


def skew_angle(hl):
    n, lab, st, _ = cv2.connectedComponentsWithStats(hl)
    angles = []
    for k in range(1, n):
        if st[k][2] < hl.shape[1] * 0.4:
            continue
        ys, xs = np.where(lab == k)
        angles.append(np.degrees(np.arctan(np.polyfit(xs, ys, 1)[0])))
    return float(np.median(angles)) if angles else 0.0


def group_positions(idx, gap=6):
    groups = []
    for x in idx:
        if groups and x - groups[-1][-1] <= gap:
            groups[-1].append(x)
        else:
            groups.append([x])
    return [int(np.mean(g)) for g in groups]


def find_lines(profile, thresh=0.25, win=13):
    """Positions of grid lines from a projection profile. Smoothing over a small
    window keeps slightly slanted / broken scan lines from being missed."""
    smooth = np.convolve(profile, np.ones(win), mode="same")
    idx = np.where(smooth > thresh)[0]
    groups = []
    for x in idx:
        if groups and x - groups[-1][-1] <= 2:
            groups[-1].append(x)
        else:
            groups.append([x])
    return [int(g[int(np.argmax(smooth[g]))]) for g in groups]


def drop_slanted(hl, max_angle=0.3):
    """Remove pieces of 'horizontal' line mask that are actually slanted
    (stray pen/scanner streaks). After deskewing, real grid lines are flat."""
    n, lab, st, _ = cv2.connectedComponentsWithStats(hl)
    out = hl.copy()
    for k in range(1, n):
        x, y, w, h, area = st[k]
        if w < 120:
            continue
        ys, xs = np.where(lab == k)
        ang = abs(np.degrees(np.arctan(np.polyfit(xs, ys, 1)[0])))
        if ang > max_angle:
            out[lab == k] = 0
    return out


def stray_line_mask(img):
    """Find long, thin, slightly slanted lines that are not part of the grid
    (pen strokes, scanner streaks) so they can be erased before OCR."""
    h, w = img.shape
    b = cv2.adaptiveThreshold(~img, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2)
    lines = cv2.HoughLinesP(b, 1, np.pi / 1440, threshold=w // 8,
                            minLineLength=w // 6, maxLineGap=12)
    mask = np.zeros_like(img)
    if lines is None:
        return mask
    for x1, y1, x2, y2 in lines.reshape(-1, 4):
        ang = abs(np.degrees(np.arctan2(y2 - y1, x2 - x1)))
        ang = min(ang, 180 - ang)
        if 0.25 < ang < 20:  # not a grid line (those are ~0 or ~90 after deskew)
            cv2.line(mask, (int(x1), int(y1)), (int(x2), int(y2)), 255, 4)
    return mask


def split_tall_rows(rows):
    """If a grid line is missing (faint, or hidden by a stray mark on the scan),
    two table rows merge into one tall band. Split bands that are a whole
    multiple of the normal row height back into equal separate rows."""
    if len(rows) < 6:
        return rows
    gaps = np.diff(rows)
    typical = float(np.median(gaps))
    # Drop a "line" that sits off the regular row spacing on both sides
    # (e.g. a stray streak through text) when the two gaps together make a
    # whole number of normal rows; the split below then restores the real line.
    rows = list(rows)
    i = 1
    while i < len(rows) - 1:
        g1, g2 = rows[i] - rows[i - 1], rows[i + 1] - rows[i]
        off1, off2 = abs(g1 - typical) > typical * 0.15, abs(g2 - typical) > typical * 0.15
        k = (g1 + g2) / typical
        if off1 and off2 and abs(k - round(k)) * typical < typical * 0.15 and round(k) >= 1:
            del rows[i]
            continue
        i += 1
    out = [rows[0]]
    for a, b in zip(rows, rows[1:]):
        k = int(round((b - a) / typical))
        if k >= 2 and abs((b - a) / k - typical) < typical * 0.15:
            step = (b - a) / k
            for i in range(1, k):
                out.append(int(round(a + step * i)))
        out.append(b)
    return out


class Page:
    """A deskewed page with its table grid detected and lines erased."""

    def __init__(self, img):
        hl, _ = line_masks(img)
        angle = skew_angle(hl)
        if abs(angle) > 0.01:
            h, w = img.shape
            m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1)
            img = cv2.warpAffine(img, m, (w, h), flags=cv2.INTER_CUBIC, borderValue=255)
        hl, vl = line_masks(img)
        hl = drop_slanted(hl)
        h, w = img.shape
        rows = find_lines(hl.sum(1) / 255 / w)
        self.cols = find_lines(vl.sum(0) / 255 / h)
        # If the scan cut off the table's outer border (page shifted on the
        # scanner), use the page edge as the missing border so the first /
        # last column (NAME, OT Rate) isn't lost.
        if self.cols and self.cols[0] > 40:
            self.cols.insert(0, 0)
        if self.cols and self.cols[-1] < w - 40:
            self.cols.append(w - 1)
        self.rows = split_tall_rows(rows)
        # Only erase horizontal line pixels that lie on a real row boundary;
        # anything else (fragments of streaks) stays so it can't eat digits.
        near = np.zeros(h, bool)
        for y in rows:
            near[max(y - 6, 0):y + 7] = True
        hl = hl.copy()
        hl[~near] = 0
        self.img, self.hl = img, hl
        mask = cv2.dilate(hl | vl, np.ones((7, 7), np.uint8))
        self.clean = img.copy()
        self.clean[mask > 0] = 255
        # Alternative image with long slanted streaks (pen strokes, scanner
        # lines) also erased - only used as a second opinion for rows whose
        # numbers don't add up, because erasing can also nick real digits.
        self.clean_alt = self.clean.copy()
        self.clean_alt[cv2.dilate(stray_line_mask(img), np.ones((3, 3), np.uint8)) > 0] = 255

    def crop(self, r, c0, c1=None, alt=False):
        c1 = c0 + 1 if c1 is None else c1
        y0, y1 = self.rows[r], self.rows[r + 1]
        x0, x1 = self.cols[c0], self.cols[c1]
        return (self.clean_alt if alt else self.clean)[y0:y1, x0:x1]

    def name_text(self, r):
        """Gentler OCR for the NAME column: only the horizontal rule is removed
        (thinly), so the tails of g / y / p that touch it survive."""
        y0, y1 = max(self.rows[r] - 2, 0), self.rows[r + 1] + 8
        x0, x1 = self.cols[0] + 6, self.cols[1] - 4
        crop = self.img[y0:y1, x0:x1].copy()
        m = cv2.dilate(self.hl[y0:y1, x0:x1], np.ones((3, 1), np.uint8))
        crop[m > 0] = 255
        crop = cv2.morphologyEx(crop, cv2.MORPH_OPEN, np.ones((3, 1), np.uint8))
        crop = cv2.resize(crop, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        crop = cv2.copyMakeBorder(crop, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=255)
        return pytesseract.image_to_string(crop, config="--psm 7").strip()

    def row_cells(self, r, pad=3, scale=1.0, alt=False):
        """OCR a whole table row as one line and assign each word to its column
        by position. Returns {column_index: text}."""
        y0, y1 = max(self.rows[r] - pad, 0), self.rows[r + 1] + pad
        crop = (self.clean_alt if alt else self.clean)[y0:y1, :]
        if scale != 1.0:
            crop = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
        crop = cv2.copyMakeBorder(crop, 20, 20, 0, 0, cv2.BORDER_CONSTANT, value=255)
        d = pytesseract.image_to_data(crop, config="--psm 6", output_type=pytesseract.Output.DICT)
        cells = {}
        for t, x, w in zip(d["text"], d["left"], d["width"]):
            t = t.strip()
            if not t:
                continue
            c = int(np.searchsorted(self.cols, (x + w / 2) / scale)) - 1
            cells.setdefault(c, []).append(t)
        return {c: " ".join(v) for c, v in cells.items()}

    @staticmethod
    def ocr(crop, numeric=False):
        if crop.size == 0 or (crop < 128).sum() < 20:
            return ""
        crop = cv2.copyMakeBorder(crop, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=255)
        cfg = "--psm 7"
        if numeric:
            cfg += " -c tessedit_char_whitelist=0123456789.,$"
        return pytesseract.image_to_string(crop, config=cfg).strip()

    def text(self, r, c0, c1=None, numeric=False, alt=False):
        return self.ocr(self.crop(r, c0, c1, alt=alt), numeric)

    def band_text(self, r):
        """OCR a whole row band across the full width (used for headers/footers)."""
        y0, y1 = self.rows[r], self.rows[r + 1]
        crop = self.clean[y0:y1, self.cols[0]:self.cols[-1]]
        if (crop < 128).sum() < 20:
            return ""
        crop = cv2.copyMakeBorder(crop, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=255)
        return pytesseract.image_to_string(crop, config="--psm 6").strip()


def parse_number(s):
    """Every amount/hour on the timesheet is printed with exactly 2 decimals, so
    read just the digits and divide by 100. This makes a dropped or misread
    decimal point / thousands comma harmless ('$1.01583', '$1,015.83' -> 1015.83)."""
    if s is None:
        return None
    digits = re.sub(r"[^0-9]", "", s)
    if not digits:
        return None
    return int(digits) / 100.0


# --------------------------------------------------------------------------
# Extraction
# --------------------------------------------------------------------------
def row_problems(v):
    """Arithmetic checks that must hold on every timesheet row.
    Returns a list of human-readable problems (empty = row is consistent)."""
    out = []
    bill = v.get("bill") or 15.0
    hrs = z(v.get("reg")) + z(v.get("sick")) + z(v.get("holiday"))
    if v.get("reg_pay") is not None and abs(hrs * bill - v["reg_pay"]) > TOL:
        out.append(f"Reg hrs x rate = {hrs * bill:.2f} but REG-Hrly pay on PDF = {v['reg_pay']:.2f}")
    straight = z(v.get("pcs")) + z(v.get("reg_pay")) + z(v.get("bonus")) + z(v.get("misc"))
    if v.get("straight") is not None and abs(straight - v["straight"]) > TOL:
        out.append(f"Pcs+Reg pay+Bonus+Misc = {straight:.2f} but Straight Time on PDF = {v['straight']:.2f}")
    total = z(v.get("straight")) + z(v.get("adj_amt")) + z(v.get("adj_ot")) + z(v.get("ot_pay"))
    if v.get("total") is not None and abs(total - v["total"]) > TOL:
        out.append(f"Straight+Adj+OT pay = {total:.2f} but Total Pay on PDF = {v['total']:.2f}")
    # OT half pay = OT half hrs x 1/2 x (straight time / total hrs)
    if v.get("total_hrs"):
        exp = z(v.get("ot_half_hrs")) * 0.5 * z(v.get("straight")) / v["total_hrs"]
        if abs(exp - z(v.get("ot_pay"))) > 0.02:
            out.append(f"OT half hrs {z(v.get('ot_half_hrs')):.2f} gives OT pay {exp:.2f} but PDF shows {z(v.get('ot_pay')):.2f}")
    if v.get("bill") is not None and v.get("ot_rate") is not None and abs(v["bill"] * 1.5 - v["ot_rate"]) > TOL:
        out.append(f"Bill rate {v['bill']:.2f} / OT rate {v['ot_rate']:.2f} don't match")
    return out


def reconcile(readings):
    """Given several OCR readings of the same row, pick per-field values so the
    row's arithmetic checks pass. Returns (values, ok)."""
    primary = readings[0]
    if not row_problems(primary):
        return primary, True
    keys = list(primary)
    options = {}
    for k in keys:
        opts = []
        for rd in readings:
            if rd.get(k) not in opts:
                opts.append(rd.get(k))
        options[k] = opts
    varying = [k for k in keys if len(options[k]) > 1]
    import itertools
    combos = 1
    for k in varying:
        combos *= len(options[k])
    if not varying or combos > 20000:
        return primary, False
    for choice in itertools.product(*(options[k] for k in varying)):
        cand = dict(primary)
        cand.update(zip(varying, choice))
        if not row_problems(cand):
            return cand, True
    return primary, False


# Header words -> position relative to the "Reg HRS" column. Each header cell
# that is recognised casts a vote for where Reg HRS is; the majority wins, so a
# few misread header words (different scans / Tesseract versions) don't matter.
HEADER_ANCHORS = [
    (lambda t: "TOTAL" in t and "HR" in t, -1),
    (lambda t: "REG" in t and "HR" in t and "HRLY" not in t and "PAY" not in t, 0),
    (lambda t: "SICK" in t, 1),
    (lambda t: "HOLID" in t, 2),
    (lambda t: "OT" in t and "HALF" in t and "PAY" not in t, 4),
    (lambda t: "PCS" in t, 5),
    (lambda t: "HRLY" in t or "S+HLD" in t, 6),
    (lambda t: "BONUS" in t, 7),
    (lambda t: "MISC" in t, 8),
    (lambda t: "STRAIGHT" in t or "EARNING" in t, 9),
    (lambda t: "AMOUNT" in t, 10),
    (lambda t: "HALF" in t and "PAY" in t, 12),
    (lambda t: "TOTAL" in t and "PAY" in t, 13),
    (lambda t: "BILL" in t, 14),
    (lambda t: "OT" in t and "RATE" in t, 15),
]


def locate_reg_column(cells, header_r, ncols):
    votes = {}
    for r in (header_r - 1, header_r, header_r + 1):
        if r < 0 or r >= len(cells):
            continue
        for c, text in cells[r].items():
            t = text.upper().replace(" ", "")
            for test, offset in HEADER_ANCHORS:
                if test(t):
                    votes[c - offset] = votes.get(c - offset, 0) + 1
                    break
    if not votes:
        raise RuntimeError("Could not recognise any column headers.")
    reg_c, n = max(votes.items(), key=lambda kv: kv[1])
    if n < 3 or reg_c < 1 or reg_c + len(PDF_COLS) > ncols:
        raise RuntimeError("Column layout not recognised - the header row could not be read "
                           "clearly enough (votes: %s)." % votes)
    return reg_c


def extract_page(img, employee_names, workers=8):
    page = Page(img)
    nrows = len(page.rows) - 1
    if nrows < 5 or len(page.cols) < 10:
        raise RuntimeError("Could not find the table grid on this page.")
    ncols = len(page.cols) - 1

    with ThreadPoolExecutor(workers) as ex:
        cells = list(ex.map(page.row_cells, range(nrows)))

    def line(r):
        return " ".join(cells[r][c] for c in sorted(cells[r]))

    header_r = next((r for r in range(min(nrows, 15))
                     if sum(w in line(r).upper() for w in ("PCS", "BONUS", "MISC", "SICK", "BILL")) >= 2), None)
    if header_r is None:
        raise RuntimeError("Could not find the column header row (PCS-WORK / BONUS / MISC).")
    totals_r = next((r for r in range(header_r + 1, nrows)
                     if "TOTAL" in cells[r].get(0, "").upper()), None)
    has_totals = totals_r is not None
    if not has_totals:          # the list continues on the next page
        totals_r = nrows
    reg_c = locate_reg_column(cells, header_r, ncols)
    col_of = {k: reg_c + i for i, k in enumerate(PDF_COLS) if reg_c + i < ncols}
    col_of["total_hrs"] = reg_c - 1

    def values_from(cellmap):
        return {k: parse_number(cellmap.get(c)) for k, c in col_of.items()}

    def cell_reading(r, alt=False):
        return {k: parse_number(page.text(r, c, numeric=True, alt=alt)) for k, c in col_of.items()}

    candidates = []
    for r in range(header_r + 1, totals_r):
        name_raw = cells[r].get(0, "")
        vals = values_from(cells[r])
        if not re.search(r"[A-Za-z]{2}", name_raw):
            continue
        if all(vals.get(k) is None for k in ("reg", "pcs", "reg_pay", "straight", "total")):
            continue  # stray text, not an employee line
        candidates.append((r, name_raw, vals))

    # Second opinion for rows whose arithmetic doesn't add up
    def fix(item):
        r, name_raw, vals = item
        if not row_problems(vals):
            return vals, True
        readings = [vals, values_from(page.row_cells(r, scale=1.5)), cell_reading(r),
                    values_from(page.row_cells(r, pad=6)),
                    values_from(page.row_cells(r, alt=True)),
                    values_from(page.row_cells(r, alt=True, scale=1.5)),
                    cell_reading(r, alt=True)]
        return reconcile(readings)

    with ThreadPoolExecutor(workers) as ex:
        fixed = list(ex.map(fix, candidates))

    def names_for(item):
        r, name_raw, _ = item
        name, status, sc = match_name(name_raw, employee_names)
        if status == "exact":
            return name, status, sc
        more = [name_raw, page.name_text(r), page.text(r, 0),
                page.row_cells(r, scale=1.5).get(0, ""), page.row_cells(r, pad=0).get(0, "")]
        return match_name([m for m in more if m], employee_names)

    with ThreadPoolExecutor(workers) as ex:
        matched_names = list(ex.map(names_for, candidates))

    # Bill rate: if a row's Bill/OT rate pair is unreadable (doesn't satisfy
    # OT = 1.5 x Bill), use the rate most rows on this page have.
    good_rates = [v["bill"] for v, _ in fixed
                  if v.get("bill") and v.get("ot_rate") and abs(v["bill"] * 1.5 - v["ot_rate"]) <= TOL]
    page_rate = max(set(good_rates), key=good_rates.count) if good_rates else 15.0
    for v, _ in fixed:
        if not (v.get("bill") and v.get("ot_rate") and abs(v["bill"] * 1.5 - v["ot_rate"]) <= TOL):
            if v.get("bill") != page_rate:
                v["bill_note"] = f"Bill rate unclear on scan (read {v.get('bill')}), used {page_rate:.2f} like the rest of the page"
            v["bill"], v["ot_rate"] = page_rate, page_rate * 1.5

    employees = []
    for (r, name_raw, _), (vals, ok), (name, matched, score) in zip(candidates, fixed, matched_names):
        employees.append({"raw_name": name_raw, "name": name, "match": matched, "score": score,
                          "consistent": ok, **vals})
    totals = values_from(cells[totals_r]) if has_totals else {}

    top_text = " ".join(line(r) for r in range(header_r + 1))
    week_end = None
    m = re.search(r"Week\s*End\w*\s*[:\-]?\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", top_text, re.I) \
        or re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", top_text)
    if m:
        try:
            week_end = dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
        except ValueError:
            week_end = None

    label = ""
    m = re.search(r"\(([A-Za-z ]{2,15})\)", top_text) or \
        re.search(r"(\d+\s*(?:st|nd|rd|th)\s*shift)", top_text, re.I)
    if m:
        label = m.group(1).strip()

    total_due = None
    for r in range(totals_r + 1, nrows):
        t = line(r)
        if "due" in t.lower():
            nums = re.findall(r"\$?\s*\d[\d,.' ]*\d", t)
            if nums:
                total_due = parse_number(nums[-1])
    return {"employees": employees, "totals": totals, "week_end": week_end,
            "label": label, "total_due": total_due, "has_totals": has_totals}


def _fold(s):
    return s.replace("y", "v").replace("g", "a")


def _norm(s):
    return re.sub(r"[^a-z]", "", s.lower())


def _clean_name(raw):
    cleaned = re.sub(r"\(.*$", "", raw)           # drop "(sophany)", "(POI..." etc
    cleaned = re.sub(r"[^A-Za-z' ]", " ", cleaned)
    return re.sub(r"\s+", " ", cleaned).strip()


def match_name(readings, names):
    """Match one or more OCR readings of a name against the Employee List.
    Leading junk words (scan specks read as text) are ignored by also trying
    word-suffixes of each reading (at least 2 words, so 'Pech Soeun' can never
    collapse to just 'Soeun').
    Returns (name, status, score) with status 'exact' | 'fuzzy' | 'new'."""
    if isinstance(readings, str):
        readings = [readings]
    keyed = {_norm(n): n for n in names}
    scores = {}
    fallback = None
    for raw in readings:
        words = _clean_name(raw).split()
        for i in range(max(len(words) - 1, 1)):
            cand = " ".join(words[i:])
            if fallback is None and re.match(r"[A-Z][A-Za-z']", words[i]):  # skip stray letters/specks
                fallback = cand
            nc = _norm(cand)
            if len(nc) < 4:
                continue
            fc = _fold(nc)
            for key, real in keyed.items():
                # also compare with the usual scan confusions folded together
                # (tail of y/g cut off by the table line: y->v, g->a)
                sc = max(difflib.SequenceMatcher(None, nc, key).ratio(),
                         difflib.SequenceMatcher(None, fc, _fold(key)).ratio() - 0.001)
                if sc > scores.get(real, 0):
                    scores[real] = sc
    ranked = sorted(scores.items(), key=lambda kv: -kv[1])
    if fallback is None:
        fallback = _clean_name(readings[0])
    # names not on the list: a word ending in "v" is almost always a "y" whose tail was cut off
    fallback = re.sub(r"\B([a-z])v\b", r"\1y", fallback)
    if not ranked:
        return fallback, "new", 0.0
    best, sc = ranked[0]
    runner = ranked[1][1] if len(ranked) > 1 else 0.0
    if sc >= 0.999:
        return best, "exact", sc
    if sc >= 0.84 and sc - runner >= 0.05:
        return best, "fuzzy", sc
    return fallback, "new", sc


# --------------------------------------------------------------------------
# Checks
# --------------------------------------------------------------------------
def z(v):
    return v or 0.0


def check(data, name_notes=False):
    issues = []
    emps = data["employees"]
    for i, e in enumerate(emps, 1):
        who = f"Row {i} ({e['name']})"
        if not name_notes:
            pass
        elif e["match"] == "new":
            issues.append(f"{who}: not in Employee List - typed as OCR read it, check spelling")
        elif e["match"] == "fuzzy" and e["score"] < 0.92:
            issues.append(f"{who}: scan was unclear, matched to Employee List name '{e['name']}' - confirm")
        for p in row_problems(e):
            issues.append(f"{who}: {p}")
        if e.get("bill_note"):
            issues.append(f"{who}: {e['bill_note']}")

    t = data["totals"]
    for k, label in [("reg", "Reg HRS"), ("ot_half_hrs", "OT Half Hrs"), ("pcs", "PCS-Work"),
                     ("reg_pay", "REG-Hrly Pay"), ("bonus", "Bonus"), ("misc", "Misc"),
                     ("straight", "Straight Time"), ("ot_pay", "OT Half Hrs-Pay"), ("total", "Total Pay")]:
        s = sum(z(e[k]) for e in emps)
        if t.get(k) is not None and abs(s - t[k]) > 0.05:
            issues.append(f"Column {label}: rows add up to {s:,.2f} but PDF TOTALS row says {t[k]:,.2f}")
    return issues


# --------------------------------------------------------------------------
# Excel writing (direct XML edit so the template's shapes/logo survive)
# --------------------------------------------------------------------------
def excel_serial(d):
    return (d - dt.date(1899, 12, 30)).days


def read_employee_list(zf):
    try:
        ss = zf.read("xl/sharedStrings.xml").decode("utf8")
    except KeyError:
        ss = ""
    strings = []
    for si in re.findall(r"<si>(.*?)</si>", ss, re.S):
        strings.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
    sheet_path = sheet_path_by_name(zf, "Employee List")
    if not sheet_path:
        return []
    xml = zf.read(sheet_path).decode("utf8")
    names = []
    for m in re.finditer(r'<c r="A(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S):
        row, attrs, inner = int(m.group(1)), m.group(2), m.group(3) or ""
        if row == 1:
            continue
        v = re.search(r"<v>(.*?)</v>", inner)
        if 't="s"' in attrs and v:
            names.append(unescape(strings[int(v.group(1))]))
        elif "inlineStr" in attrs:
            names.append(unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner))))
    return [n.strip() for n in names if n.strip()]


def unescape(s):
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def sheet_path_by_name(zf, name):
    wb = zf.read("xl/workbook.xml").decode("utf8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf8")
    m = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="(rId\d+)"' % re.escape(name), wb)
    if not m:
        return None
    t = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels) or \
        re.search(r'Target="([^"]+)"[^>]*Id="%s"' % m.group(1), rels)
    return "xl/" + t.group(1).lstrip("/").replace("xl/", "") if t else None


def num(v):
    return repr(round(float(v), 6)).rstrip("0").rstrip(".") if v is not None else ""


def cell(ref, style, value=None, formula=None, text=None):
    s = f' s="{style}"' if style else ""
    if text is not None:
        return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>'
    if formula is not None:
        v = f"<v>{num(value)}</v>" if value is not None else ""
        return f'<c r="{ref}"{s}><f>{escape(formula)}</f>{v}</c>'
    if value is None:
        return f'<c r="{ref}"{s}/>'
    return f'<c r="{ref}"{s}><v>{num(value)}</v></c>'


def get_style(row_xml, ref):
    m = re.search(r'<c r="%s"[^>]*?s="(\d+)"' % ref, row_xml)
    return m.group(1) if m else None


def set_cell(sheet, ref, **kw):
    """Replace one cell (keeping its style) inside the sheet xml."""
    row = int(re.sub(r"[A-Z]", "", ref))
    pat = re.compile(r'<c r="%s"([^>]*?)(?:/>|>.*?</c>)' % ref, re.S)
    m = pat.search(sheet)
    if m:
        st = re.search(r's="(\d+)"', m.group(1))
        new = cell(ref, st.group(1) if st else None, **kw)
        return sheet[:m.start()] + new + sheet[m.end():]
    # cell not present: append to its row
    rm = re.search(r'(<row r="%d"[^>]*>)(.*?)(</row>)' % row, sheet, re.S)
    if not rm:
        raise RuntimeError(f"Template row {row} not found")
    return sheet[:rm.end(2)] + cell(ref, None, **kw) + sheet[rm.end(2):]


def shift_refs(text, delta, from_row):
    def f(m):
        r = int(m.group(2))
        return f"{m.group(1)}{r + delta if r >= from_row else r}"
    return re.sub(r"\b([A-Z]{1,2})(\d+)\b", f, text)


def build_invoice(template_path, out_path, data, invoice_no, invoice_date, week_end):
    with zipfile.ZipFile(template_path) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
        inv_path = sheet_path_by_name(zin, "Invoice")
    sheet = files[inv_path].decode("utf8")

    emps = data["employees"]
    n = len(emps)
    first = FIRST_DATA_ROW
    last = first + n - 1
    delta = (last + 1) - TEMPLATE_TOTAL_ROW  # how far everything below the table moves
    T = TEMPLATE_TOTAL_ROW + delta           # new totals row

    head, rest = sheet.split("<sheetData>", 1)
    body, tail = rest.split("</sheetData>", 1)
    rows = {int(m.group(1)): m.group(0) for m in re.finditer(r'<row r="(\d+)".*?</row>', body, re.S)}

    tmpl_row = rows[FIRST_DATA_ROW + 1]
    st = {c: get_style(tmpl_row, f"{c}{FIRST_DATA_ROW + 1}") for c in "ABCDEFGHIJKLMNOPQ"}
    row_attr = re.match(r'<row r="\d+"([^>]*)>', tmpl_row).group(1)
    row_attr = re.sub(r'\s*spans="[^"]*"', "", row_attr)

    out_rows = []
    for r in sorted(rows):
        if r < FIRST_DATA_ROW:
            out_rows.append(rows[r])

    # ---- employee rows
    tot = {k: 0.0 for k in "CDEFGHIJKLMNO"}
    for i, e in enumerate(emps):
        r = first + i
        bill = e["bill"] if e["bill"] else 15.0
        C = e["reg"] or None
        D = (z(e["sick"]) + z(e["holiday"])) or None
        E = e["adj_hrs"] or None
        F = e["ot_half_hrs"] or None
        G = z(e["pcs"])
        H = (z(C) + z(D)) * bill
        I = e["bonus"] or None
        J = e["misc"] or None
        K = G + H + z(I) + z(J)
        L = e["adj_amt"] or None
        M = e["adj_ot"] or None
        N = z(e["ot_pay"])
        O = K + z(L) + z(M) + N
        for k, v in zip("CDEFGHIJKLMNO", [C, D, E, F, G, H, I, J, K, L, M, N, O]):
            tot[k] += z(v)
        cells = [
            cell(f"A{r}", st["A"], i + 1),
            cell(f"B{r}", st["B"], text=e["name"]),
            cell(f"C{r}", st["C"], C), cell(f"D{r}", st["D"], D), cell(f"E{r}", st["E"], E),
            cell(f"F{r}", st["F"], F), cell(f"G{r}", st["G"], G),
            cell(f"H{r}", st["H"], H, formula=f"(C{r}+D{r})*P{r}"),
            cell(f"I{r}", st["I"], I), cell(f"J{r}", st["J"], J),
            cell(f"K{r}", st["K"], K, formula=f"SUM(G{r}:J{r})"),
            cell(f"L{r}", st["L"], L), cell(f"M{r}", st["M"], M), cell(f"N{r}", st["N"], N),
            cell(f"O{r}", st["O"], O, formula=f"SUM(K{r}:N{r})"),
            cell(f"P{r}", st["P"], bill),
            cell(f"Q{r}", st["Q"], bill * 1.5, formula=f"P{r}*1.5"),
        ]
        out_rows.append(f'<row r="{r}"{row_attr}>' + "".join(cells) + "</row>")

    # ---- everything from the template totals row down, shifted
    r70, r71, r72, r73, r74, r76 = (x + delta for x in (70, 71, 72, 73, 74, 76))
    c70 = tot["H"]
    c71 = tot["O"] - tot["H"]
    c72 = (tot["C"] + tot["D"]) * 1.7
    c73 = c71 * 0.17
    formulas = {f"{c}{T}": (f"SUM({c}{first}:{c}{last})", tot[c]) for c in "CDEFGHIJKLMNO"}
    formulas.update({
        f"C{r70}": (f"SUM(H{first}:H{last})", c70),
        f"C{r71}": (f"O{T}-H{T}", c71),
        f"C{r72}": (f"SUM((C{T}+D{T})*1.7)", c72),
        f"C{r73}": (f"SUM(C{r71}*0.17)", c73),
        f"C{r74}": (f"SUM(C{r70}:F{r71})", c70 + c71),
        f"C{r76}": (f"SUM(C{r70}:E{r73})", c70 + c71 + c72 + c73),
    })
    for r in sorted(rows):
        if r < TEMPLATE_TOTAL_ROW:
            continue
        rx = rows[r]
        new_r = r + delta
        rx = re.sub(r'<row r="\d+"', f'<row r="{new_r}"', rx, count=1)
        rx = re.sub(r'<c r="([A-Z]+)\d+"', lambda m: f'<c r="{m.group(1)}{new_r}"', rx)
        for ref in re.findall(r'<c r="([A-Z]+\d+)"[^>]*>\s*<f', rx):
            if ref in formulas:
                f, v = formulas[ref]
                rx = re.sub(r'(<c r="%s"[^>]*>).*?</c>' % ref,
                            lambda m: f"{m.group(1)}<f>{escape(f)}</f><v>{num(v)}</v></c>", rx, flags=re.S)
            else:  # unknown formula: shift its references generically, drop cached value
                rx = re.sub(r'(<c r="%s"[^>]*>)<f>(.*?)</f>(<v>.*?</v>)?' % ref,
                            lambda m: f"{m.group(1)}<f>{shift_refs(m.group(2), delta, TEMPLATE_TOTAL_ROW)}</f>", rx)
        out_rows.append(rx)

    sheet = head + "<sheetData>" + "".join(out_rows) + "</sheetData>" + tail

    # ---- header fields
    grand = c70 + c71 + c72 + c73
    sheet = set_cell(sheet, "J4", value=excel_serial(invoice_date))
    sheet = set_cell(sheet, "O4", text=invoice_no)
    sheet = set_cell(sheet, "J8", value=excel_serial(week_end - dt.timedelta(days=6)))
    sheet = set_cell(sheet, "O8", value=excel_serial(week_end))
    sheet = set_cell(sheet, "J22", formula=f"(SUM(K{T}:N{T})+C{r72}+C{r73})",
                     value=tot["K"] + tot["L"] + tot["M"] + tot["N"] + c72 + c73)

    # ---- merged cells, validation, dimension, view
    def fix_merge(m):
        return f'<mergeCell ref="{shift_refs(m.group(1), delta, TEMPLATE_TOTAL_ROW)}"/>'
    sheet = re.sub(r'<mergeCell ref="([^"]+)"/>', fix_merge, sheet)
    sheet = re.sub(r'(<dataValidation [^>]*sqref=")[^"]*(")', rf"\g<1>B{first}:B{last}\g<2>", sheet)
    sheet = re.sub(r"<ignoredErrors>.*?</ignoredErrors>", "", sheet, flags=re.S)
    sheet = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+)(\d+)"/>',
                   lambda m: f'<dimension ref="{m.group(1)}:{m.group(2)}{int(m.group(3)) + delta}"/>', sheet)
    sheet = re.sub(r'\s*topLeftCell="[^"]*"', "", sheet)
    sheet = re.sub(r'<selection [^/]*/>', '<selection activeCell="A1" sqref="A1"/>', sheet)
    files[inv_path] = sheet.encode("utf8")

    # ---- force Excel to recalculate everything on open; drop stale calc chain
    wb = files["xl/workbook.xml"].decode("utf8")
    wb = re.sub(r"<calcPr([^>]*?)\s*/>",
                lambda m: "<calcPr" + re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wb)
    files["xl/workbook.xml"] = wb.encode("utf8")
    if "xl/calcChain.xml" in files:
        del files["xl/calcChain.xml"]
        rels = files["xl/_rels/workbook.xml.rels"].decode("utf8")
        files["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", rels).encode("utf8")
        ct = files["[Content_Types].xml"].decode("utf8")
        files["[Content_Types].xml"] = re.sub(r'<Override [^>]*calcChain[^>]*/>', "", ct).encode("utf8")

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in files.items():
            zout.writestr(name, blob)
    return grand


# --------------------------------------------------------------------------
def next_invoice_no(no, k):
    m = re.match(r"^(.*?)(\d+)$", no)
    if not m or k == 0:
        return no if k == 0 else f"{no}-{k + 1}"
    return f"{m.group(1)}{int(m.group(2)) + k:0{len(m.group(2))}d}"


def extract_page_any_orientation(img, names):
    """Read a page; if the table can't be found, the scan may be sideways or
    upside down, so try the other three orientations before giving up."""
    try:
        return extract_page(img, names)
    except RuntimeError as first_error:
        for rot in (cv2.ROTATE_90_CLOCKWISE, cv2.ROTATE_90_COUNTERCLOCKWISE, cv2.ROTATE_180):
            try:
                return extract_page(cv2.rotate(img, rot), names)
            except RuntimeError:
                continue
        raise first_error


def find_tesseract():
    """Find the Tesseract program. Order: a copy bundled inside the app (Windows
    build), then PATH, then the usual install locations (apps launched from
    Finder don't get Homebrew's PATH)."""
    import shutil
    exe = "tesseract.exe" if os.name == "nt" else "tesseract"
    bundle = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    cands = [os.path.join(bundle, "tesseract", exe), shutil.which("tesseract")]
    if os.name == "nt":
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                     os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs")):
            cands.append(os.path.join(base, "Tesseract-OCR", exe))
    else:
        cands += ["/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract", "/usr/bin/tesseract"]
    for cand in cands:
        if cand and os.path.exists(cand):
            return cand
    return None


def unique_path(path):
    """Don't overwrite an existing invoice: add (2), (3)... to the name."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"


def group_pages(extracted):
    """Join pages that belong to one timesheet. A page without a TOTALS row
    continues onto the next page, unless the next page starts a new timesheet
    (it has its own 'Week Ending' heading). Returns a list of groups, each a
    list of (page_number, data_or_error)."""
    groups, current = [], None
    for pno, data in extracted:
        if isinstance(data, Exception):
            groups.append([(pno, data)])
            current = None
            continue
        continues = (current is not None and not current[-1][1]["has_totals"]
                     and data["week_end"] is None)
        if continues:
            current.append((pno, data))
        else:
            current = [(pno, data)]
            groups.append(current)
        if data["has_totals"]:
            current = None
    return groups


def confirm_check_columns(employees, totals):
    """REG-Hrly pay, Straight Time and Total Pay are only used for checking (the
    invoice recalculates them). If some rows disagree with their own arithmetic
    and using the calculated values makes the column match the PDF's TOTALS row,
    the PDF itself confirms the correction (typically a column cut off or smudged
    on the scan)."""
    calc = {
        "reg_pay": lambda v: (z(v.get("reg")) + z(v.get("sick")) + z(v.get("holiday"))) * (v.get("bill") or 15.0),
        "straight": lambda v: z(v.get("pcs")) + z(v.get("reg_pay")) + z(v.get("bonus")) + z(v.get("misc")),
        "total": lambda v: z(v.get("straight")) + z(v.get("adj_amt")) + z(v.get("adj_ot")) + z(v.get("ot_pay")),
    }
    for k in ("reg_pay", "straight", "total"):
        if totals.get(k) is None:
            continue
        wrong = [e for e in employees if e.get(k) is None or abs(calc[k](e) - e[k]) > TOL]
        if not wrong or len(wrong) > max(3, len(employees) // 10):
            continue
        s = sum(calc[k](e) if e in wrong else z(e.get(k)) for e in employees)
        if abs(s - totals[k]) <= 0.05:
            for e in wrong:
                e[k] = round(calc[k](e), 2)


def combine(pages_data):
    """Merge the pages of one timesheet into a single data dict."""
    first = pages_data[0]
    closing = next((d for d in reversed(pages_data) if d["has_totals"]), None)
    employees = [e for d in pages_data for e in d["employees"]]
    if closing:
        confirm_check_columns(employees, closing["totals"])
    return {
        "employees": employees,
        "totals": closing["totals"] if closing else {},
        "week_end": next((d["week_end"] for d in pages_data if d["week_end"]), None),
        "label": first["label"],
        "total_due": closing["total_due"] if closing else None,
        "has_totals": closing is not None,
    }


def convert_pdf(pdf, template, invoice_no, out_dir, invoice_date=None, week_ending=None,
                pages=None, prefix="MLL_UNWRAPPED_INVOICE", check_names=False,
                progress=None, overwrite=True):
    """Convert a timesheet PDF into invoice workbooks: one per timesheet. Usually
    that's one per page; a timesheet that runs over several pages (no TOTALS row
    until its last page) becomes a single invoice. Returns a list of result dicts.
    A page that can't be read gets an 'error' entry instead of stopping the rest."""
    invoice_date = invoice_date or dt.date.today()
    os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(template) as zf:
        names = read_employee_list(zf)
    total_pages = len(pymupdf.open(pdf))

    extracted = []
    for pno, img in enumerate(render_pages(pdf), 1):
        if pages and pno not in pages:
            continue
        if progress:
            progress(pno, total_pages)
        try:
            extracted.append((pno, extract_page_any_orientation(img, names)))
        except Exception as e:
            extracted.append((pno, e))

    results = []
    k = 0
    for group in group_pages(extracted):
        nums = [pno for pno, _ in group]
        page_ref = str(nums[0]) if len(nums) == 1 else f"{nums[0]}-{nums[-1]}"
        res = {"page": page_ref}
        try:
            if isinstance(group[0][1], Exception):
                raise group[0][1]
            data = combine([d for _, d in group])
            week_end = week_ending or data["week_end"]
            if week_end is None:
                raise RuntimeError("Could not read the 'Week Ending' date on this page.")
            inv_no = next_invoice_no(invoice_no, k)
            out = os.path.join(out_dir, f"{prefix}_{inv_no}.xlsx")
            if not overwrite:
                out = unique_path(out)
            grand = build_invoice(template, out, data, inv_no, invoice_date, week_end)
            k += 1
            issues = check(data, name_notes=check_names)
            due = data["total_due"]
            # rounding: the PDF adds hidden decimals, so allow a few cents when every row checks out
            rows_ok = all(not row_problems(e) for e in data["employees"])
            total_ok = due is not None and abs(grand - due) <= (0.05 if rows_ok else 0.02)
            if not data["has_totals"]:
                issues.append("No TOTALS row found at the end of this timesheet - couldn't cross-check "
                              "against the PDF; check the total by hand")
            elif due is not None and not total_ok:
                issues.append(f"Invoice total ${grand:,.2f} differs from the PDF's Total due ${due:,.2f}")
            elif due is None:
                issues.append("Couldn't read 'Total due' on the PDF - check the total by hand")
            label = data["label"] or f"Page {page_ref}"
            if len(nums) > 1:
                label += f" (pages {nums[0]}-{nums[-1]})"
            res.update(label=label, employees=len(data["employees"]),
                       week_start=week_end - dt.timedelta(days=6), week_end=week_end,
                       invoice_no=inv_no, total=grand, pdf_total=due, total_ok=total_ok,
                       issues=issues, path=out)
        except Exception as e:  # keep going with the other pages
            res.update(label=f"Page {page_ref}", error=str(e))
        results.append(res)
    return results


def default_template():
    here = os.path.dirname(os.path.abspath(__file__))
    for cand in (os.path.join(here, "template.xlsx"), "template.xlsx"):
        if os.path.exists(cand):
            return cand
    return None


def main():
    ap = argparse.ArgumentParser(description="Convert a scanned timesheet PDF into Excel invoice(s).")
    ap.add_argument("pdf")
    ap.add_argument("--template", default=None,
                    help="Invoice .xlsx to copy the layout from (default: template.xlsx next to this script)")
    ap.add_argument("--invoice-no", required=True,
                    help="Invoice number for the first page; later pages count up (UN-B-80, UN-B-81, ...)")
    ap.add_argument("--invoice-date", help="Invoice date YYYY-MM-DD (default: today)")
    ap.add_argument("--week-ending", help="Override week-ending date YYYY-MM-DD (default: read from the PDF)")
    ap.add_argument("--pages", help="Which pages to convert, e.g. '1' or '1,2' (default: all)")
    ap.add_argument("--out-dir", default=".", help="Folder for the output workbooks")
    ap.add_argument("--prefix", default="MLL_UNWRAPPED_INVOICE", help="Output file name prefix")
    ap.add_argument("--check-names", action="store_true",
                    help="Also list names the scan made unclear / that aren't in the template's Employee List")
    ap.add_argument("--tesseract", help="Path to tesseract executable if it is not on PATH")
    args = ap.parse_args()

    tess = args.tesseract or find_tesseract()
    if tess:
        pytesseract.pytesseract.tesseract_cmd = tess
    template = args.template or default_template()
    if not template:
        sys.exit("Could not find template.xlsx. Put your invoice template in the same folder as "
                 "this script and name it template.xlsx (or pass --template FILE.xlsx).")
    print(f"Using template: {template}")

    results = convert_pdf(
        args.pdf, template, args.invoice_no, args.out_dir,
        invoice_date=dt.date.fromisoformat(args.invoice_date) if args.invoice_date else None,
        week_ending=dt.date.fromisoformat(args.week_ending) if args.week_ending else None,
        pages={int(p) for p in args.pages.split(",")} if args.pages else None,
        prefix=args.prefix, check_names=args.check_names,
        progress=lambda p, n: print(f"\n=== Page {p}: reading (OCR)..."))

    any_issues = False
    for r in results:
        if "error" in r:
            any_issues = True
            print(f"    !! Page {r['page']} could not be converted: {r['error']}")
            continue
        print(f"    {r['label']}: {r['employees']} employees, week {r['week_start']} to {r['week_end']}")
        print(f"    Total due: ${r['total']:,.2f}" +
              (f"  (PDF says ${r['pdf_total']:,.2f})" if r["pdf_total"] is not None else ""))
        print(f"    Saved -> {r['path']}")
        if r["total_ok"]:
            print("    Invoice total matches the PDF's 'Total due'.")
        if r["issues"]:
            any_issues = True
            print(f"    !! {len(r['issues'])} item(s) to review:")
            for s_ in r["issues"]:
                print("       - " + s_)
        else:
            print("    All rows and totals match the PDF.")
    if any_issues:
        print("\nPlease check the flagged items in Excel before sending.")


if __name__ == "__main__":
    main()
