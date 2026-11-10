"""
Label sheet (employee stickers) and check amounts.

Writes a Word file laid out exactly like labels_template.docx: Letter paper,
3 labels across, 1" rows. Only the text inside each label changes, so it
lines up with the pre-cut sticker paper.

Some employees get part of their pay as a check. Their label shows
"Total Pay = <cash> + <check>" with the cash part in bold, where
cash = (Total Pay from the timesheet, rounded to the dollar) - check.
"""
import difflib
import re
import shutil
import tempfile
import unicodedata
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from xml.sax.saxutils import escape

# Starting check list (client UW), from the ChecksTable in the payroll workbook.
# "aliases" are other ways the name is written. The app keeps its own copy of
# this list and updates it whenever the amounts are confirmed.
DEFAULT_CHECKS = [
    {"name": "Ana Caraballo", "aliases": [], "amount": 300},
    {"name": "Sovannary Chay", "aliases": [], "amount": 300},
    {"name": "Lam KimAnh", "aliases": ["Kim Duong"], "amount": 240},
    {"name": "Sophea Hor", "aliases": [], "amount": 300},
    {"name": "Sera Krim", "aliases": [], "amount": 300},
    {"name": "Sreylen Leng", "aliases": [], "amount": 450},
    {"name": "Leilanys Mercado", "aliases": [], "amount": 300},
    {"name": "Khan Nhep", "aliases": [], "amount": 300},
    {"name": "Tola Sam", "aliases": [], "amount": 300},
    {"name": "Kimsrey Taing", "aliases": [], "amount": 300},
    {"name": "Savun An Buth", "aliases": [], "amount": 300},
    {"name": "Visal Kim", "aliases": [], "amount": 300},
    {"name": "Yon Chheng", "aliases": [], "amount": 240},
    {"name": "Ravy Suon", "aliases": [], "amount": 0},
]


# ---------------------------------------------------------------- numbers ---
def z(v):
    return v or 0.0


def dollars(v):
    """Regular rounding to the dollar: 50 cents or more rounds up (750.5 -> 751, 166.31 -> 166)."""
    return int(Decimal(str(round(z(v), 4))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def fmt(v, blank_if_zero=False):
    """38.50 -> '38.5', 577.00 -> '577', 0 -> '0' (or '' when blank_if_zero)."""
    v = round(z(v), 2)
    if v == 0:
        return "" if blank_if_zero else "0"
    return f"{v:.2f}".rstrip("0").rstrip(".")


def usd(v, blank_if_zero=False):
    """Dollar amounts on the label: whole dollars, 50 cents and up rounds up
    (610.52 -> '611', 166.31 -> '166')."""
    d = dollars(v)
    return "" if (d == 0 and blank_if_zero) else str(d)


# ------------------------------------------------------------------ people ---
def label_people(groups):
    """One entry per employee across every timesheet in the PDF (someone on two
    shifts gets one label with both added together), sorted by name."""
    people = {}
    for g in groups:
        for e in g.get("data", {}).get("employees", []):
            name = (e.get("name") or e.get("raw_name") or "").strip()
            if not name:
                continue
            reg_hrs = z(e.get("reg")) + z(e.get("sick")) + z(e.get("holiday"))
            reg_pay = e.get("reg_pay")
            if reg_pay is None:
                reg_pay = reg_hrs * (e.get("bill") or 15.0)
            item = {
                "reg_hrs": reg_hrs, "reg_pay": z(reg_pay), "pcs": z(e.get("pcs")),
                "ot_hrs": z(e.get("ot_half_hrs")), "ot_pay": z(e.get("ot_pay")) + z(e.get("adj_ot")),
                "bonus": z(e.get("bonus")), "misc": z(e.get("misc")), "adj_min": z(e.get("adj_amt")),
                "total": z(e.get("total")),
            }
            p = people.setdefault(_key(name), {"name": name, **{k: 0.0 for k in item}})
            for k, v in item.items():
                p[k] += v
    return sorted(people.values(), key=lambda p: _fold(p["name"]).lower())


# ---------------------------------------------------------- name matching ---
def _fold(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def _tokens(s):
    return sorted(re.findall(r"[a-z]+", _fold(s).lower()))


def _key(s):
    return " ".join(_tokens(s))


def name_score(a, b):
    """1.0 for the same name in any order ('Savun An Buth' = 'Buth Savun An'),
    high for small spelling differences ('Chay' vs 'Chhay')."""
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return 0.0
    if ta == tb:
        return 1.0
    whole = difflib.SequenceMatcher(None, " ".join(ta), " ".join(tb)).ratio()
    # also try with spaces removed ("KimAnh" vs "Kim Anh")
    joined = difflib.SequenceMatcher(None, "".join(ta), "".join(tb)).ratio()
    return max(whole, joined)


MATCH_AT = 0.86


def match_checks(people, checklist):
    """Pair check-list entries with people on this timesheet.
    Returns (matches, missing): matches is a list of
    {"name", "total", "amount", "entry": index into checklist}; missing lists
    check-list names not found on this timesheet."""
    pairs = []
    for i, entry in enumerate(checklist):
        names = [entry["name"]] + list(entry.get("aliases", []))
        for j, p in enumerate(people):
            s = max(name_score(n, p["name"]) for n in names)
            if s >= MATCH_AT:
                pairs.append((s, i, j))
    pairs.sort(reverse=True)
    used_i, used_j, matches = set(), set(), []
    for s, i, j in pairs:
        if i in used_i or j in used_j:
            continue
        used_i.add(i)
        used_j.add(j)
        p = people[j]
        matches.append({"name": p["name"], "total": dollars(p["total"]),
                        "amount": checklist[i].get("amount") or 0, "entry": i})
    matches.sort(key=lambda m: _fold(m["name"]).lower())
    missing = [e["name"] for i, e in enumerate(checklist) if i not in used_i]
    return matches, missing


def update_checklist(checklist, people, confirmed):
    """Fold the amounts confirmed in the pop-up back into the saved list.
    confirmed: list of {"name": <name on timesheet>, "amount": n,
    "entry": index or None (someone newly added)}; removed: entries with
    "remove": True are dropped from the list."""
    out = [dict(e, aliases=list(e.get("aliases", []))) for e in checklist]
    drop = set()
    for c in confirmed:
        i = c.get("entry")
        if c.get("remove"):
            if i is not None:
                drop.add(i)
            continue
        if i is None:
            out.append({"name": c["name"], "aliases": [], "amount": c["amount"]})
            continue
        e = out[i]
        e["amount"] = c["amount"]
        known = [e["name"]] + e["aliases"]
        if not any(_key(n) == _key(c["name"]) for n in known):
            e["aliases"].append(c["name"])  # remember how the timesheet spells it
    return [e for i, e in enumerate(out) if i not in drop]


# --------------------------------------------------------------- the docx ---
SZ = '<w:sz w:val="18"/><w:szCs w:val="18"/>'
B = '<w:b/><w:bCs/>'


def _run(text, bold=False, noproof=False):
    if text == "":
        return ""
    rpr = (B if bold else "") + ("<w:noProof/>" if noproof else "") + SZ
    sp = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t{sp}>{escape(text)}</w:t></w:r>'


def _para(runs, bold_mark=False, ind=False):
    ppr = ('<w:ind w:right="115"/>' if ind else "") + f'<w:rPr>{B if bold_mark else ""}{SZ}</w:rPr>'
    return f'<w:p><w:pPr>{ppr}</w:pPr>{"".join(runs)}</w:p>'


def _cell(width, *paras):
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{"".join(paras)}</w:tc>'


def _total_runs(p, check):
    total = dollars(p["total"])
    if check:
        return [_run("Total Pay = ", bold=True), _run(usd(total - dollars(check)), bold=True, noproof=True),
                _run(f" + {usd(check)}", noproof=True)]
    return [_run("Total Pay = ", bold=True), _run(str(total), bold=True, noproof=True)]


def _last_line(p):
    if p["adj_min"] and not p["misc"]:
        return [_run("Adj Min = "), _run(usd(p["adj_min"]), noproof=True)]
    return [_run("Misc = "), _run(usd(p["misc"] + p["adj_min"], blank_if_zero=True), noproof=True)]


def label_cell(p, check=0):
    rows = [
        (_cell(1844, _para([_run(p["name"], bold=True, noproof=True)], bold_mark=True, ind=True),
               _para([], bold_mark=True, ind=True)),
         _cell(1710, _para(_total_runs(p, check)))),
        (_cell(1844, _para([_run("Reg Hrs: "), _run(fmt(p["reg_hrs"], True), noproof=True),
                            _run(" = "), _run(usd(p["reg_pay"]), noproof=True)])),
         _cell(1710, _para([_run("PCS Work = "), _run(usd(p["pcs"]), noproof=True)]))),
        (_cell(1844, _para([_run("OT Hrs: "), _run(fmt(p["ot_hrs"], True), noproof=True),
                            _run(" = "), _run(usd(p["ot_pay"]), noproof=True)])),
         _cell(1710, _para([_run("Bonus = "), _run(usd(p["bonus"], True), noproof=True)]))),
        (_cell(1844, _para([], bold_mark=True)), _cell(1710, _para(_last_line(p)))),
    ]
    none = '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    look = '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
    inner = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="3554" w:type="dxa"/>'
             f'<w:tblBorders>{none}</w:tblBorders><w:tblLayout w:type="fixed"/>{look}</w:tblPr>'
             '<w:tblGrid><w:gridCol w:w="1844"/><w:gridCol w:w="1710"/></w:tblGrid>'
             + "".join(f"<w:tr>{a}{b}</w:tr>" for a, b in rows) + "</w:tbl>")
    small_b = '<w:p><w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:pPr></w:p>'
    middle = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="3636" w:type="dxa"/>'
              f'<w:tblInd w:w="108" w:type="dxa"/><w:tblLayout w:type="fixed"/>{look}</w:tblPr>'
              '<w:tblGrid><w:gridCol w:w="3636"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="311"/></w:trPr>'
              '<w:tc><w:tcPr><w:tcW w:w="3636" w:type="dxa"/><w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/>'
              f'<w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders></w:tcPr>{inner}{small_b}</w:tc></w:tr></w:tbl>')
    after = ('<w:p><w:pPr><w:spacing w:before="111"/><w:ind w:left="108" w:right="108"/>'
             '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:pPr></w:p>')
    return f'<w:tc><w:tcPr><w:tcW w:w="3780" w:type="dxa"/></w:tcPr>{middle}{after}</w:tc>'


EMPTY_LABEL = ('<w:tc><w:tcPr><w:tcW w:w="3780" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:before="111"/>'
               '<w:ind w:left="108" w:right="108"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:pPr></w:p></w:tc>')
GAP = '<w:tc><w:tcPr><w:tcW w:w="180" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:ind w:left="108" w:right="108"/></w:pPr></w:p></w:tc>'
ROW_PR = '<w:trPr><w:cantSplit/><w:trHeight w:hRule="exact" w:val="1440"/></w:trPr>'


def build_labels(template_docx, out_path, people, checks=None):
    """Write the label sheet. checks: {employee name: check amount}."""
    checks = checks or {}
    with zipfile.ZipFile(template_docx) as zf:
        doc = zf.read("word/document.xml").decode("utf-8")
    # keep everything around the label table (page size, margins, column widths)
    first_row = re.search(r'<w:tr[ >]', doc).start()
    sect = doc.rindex("<w:sectPr")
    table_end = doc.rindex("</w:tbl>", 0, sect)
    head, tail = doc[:first_row], doc[table_end:]

    cells = [label_cell(p, checks.get(p["name"], 0)) for p in people]
    while len(cells) % 3 or not cells:
        cells.append(EMPTY_LABEL)
    rows = ["<w:tr>" + ROW_PR + cells[i] + GAP + cells[i + 1] + GAP + cells[i + 2] + "</w:tr>"
            for i in range(0, len(cells), 3)]
    new_doc = head + "".join(rows) + tail

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    with zipfile.ZipFile(template_docx) as src, zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = new_doc.encode("utf-8") if item.filename == "word/document.xml" else src.read(item.filename)
            dst.writestr(item, data)
    shutil.move(tmp.name, out_path)
    return out_path
