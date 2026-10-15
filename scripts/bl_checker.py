"""
B/L CHECKER  -  compare a CLIENT B/L PDF against the HAL DRAFT B/L PDF.

Reads the 27 B/L fields from each PDF (by position on the form), compares them
field by field and writes a colour-coded Excel report:  BL CHECK <B/L NO>.xlsx

Usage
  BL CHECKER.exe                       -> opens the window
  BL CHECKER.exe CLIENT.pdf DRAFT.pdf  -> checks the pair, writes the report
  BL CHECKER.exe <folder>              -> checks every CLIENT*/DRAFT* pair in the folder
  bl_checker.py --cli <folder> --outdir <out>   -> same, reports written to <out>

Form layouts (field boxes) live in layouts.json next to the program; a default
copy is written there on first run so new B/L forms can be added without code.
"""
import difflib
import json
import os
import re
import sys
from datetime import datetime

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

APP_NAME = "B/L CHECKER"
APP_VERSION = "1.1"

# --------------------------------------------------------------------------
# The 27 fields of the check list (BL DETAILS.xlsx), in order.
# --------------------------------------------------------------------------
FIELDS = [
    ("bl_no", "B/L NO."),
    ("shipper", "SHIPPER"),
    ("consignee", "CONSIGNEE"),
    ("notify", "NOTIFY PARTY"),
    ("pre_carriage", "PRE CARRIAGE BY"),
    ("place_receipt", "PLACE OF RECEIPT"),
    ("vessel_voy", "OCEAN VESSEL VOY NO."),
    ("pol", "PORT OF LOADING"),
    ("pod", "PORT OF DISCHARGE"),
    ("place_delivery", "PLACE OF DELIVERY"),
    ("final_dest", "FINAL DESTINATION (FOR THE MERCHANTS REFERENCE ONLY)"),
    ("container_seal", "CONTAINER NO. / SEAL NO."),
    ("no_pkgs", "NO. OF CONTAINERS OR PKGS"),
    ("marks", "MARKS AND NUMBERS"),
    ("description", "KIND OF PACKAGES / DESCRIPTION OF GOODS"),
    ("gross_weight", "GROSS WEIGHT"),
    ("measurement", "MEASUREMENT"),
    ("total_words", "TOTAL NO. OF CONTAINERS OR PACKAGES (IN WORDS)"),
    ("freight", "FREIGHT AND CHARGES / REVENUE TONS / RATE / PER / PREPAID / COLLECT"),
    ("ex_rate", "EX RATE"),
    ("prepaid_at", "PREPAID AT"),
    ("payable_at", "PAYABLE AT"),
    ("place_date_issue", "PLACE AND DATE OF ISSUE"),
    ("total_prepaid_local", "TOTAL PREPAID IN LOCAL CURRENCY"),
    ("no_originals", "NO. OF ORIGINAL B(S)/L"),
    ("laden_on_board", "DATE / LADEN ON BOARD THE VESSEL"),
    ("service_term", "CY/CY (SERVICE TERM)"),
]
FIELD_LABEL = dict(FIELDS)
RIDER_ROW = ("rider", "RIDER / ATTACHED SHEET (AS PER ATTACHED RIDER / DESCRIPTIONS)")

# Boxes are [x0, top, x1, bottom] in PDF points (A4 = 595 x 842, origin top-left).
# "cargo" is the container/goods table, split into columns by x position.
DEFAULT_LAYOUTS = {
    "HAL DRAFT (system print)": {
        "detect_label": "Shipper",
        "ignore_font_below": 6.0,
        "fields": {
            "bl_no": [355, 28, 595, 53],
            "shipper": [12, 15, 355, 91],
            "consignee": [12, 91, 355, 166],
            "notify": [12, 166, 355, 240],
            "pre_carriage": [12, 240, 195, 259],
            "place_receipt": [195, 240, 355, 259],
            "vessel_voy": [12, 259, 195, 278],
            "pol": [195, 259, 355, 278],
            "pod": [12, 278, 195, 296],
            "place_delivery": [195, 278, 372, 296],
            "final_dest": [372, 278, 595, 296],
            "total_words": [12, 544, 470, 566],
            "service_term": [470, 544, 595, 566],
            "freight": [12, 566, 595, 692],
            "ex_rate": [12, 692, 82, 717],
            "prepaid_at": [82, 692, 243, 717],
            "payable_at": [243, 692, 383, 717],
            "place_date_issue": [383, 692, 595, 717],
            "total_prepaid_local": [82, 717, 243, 747],
            "no_originals": [243, 717, 383, 747],
            "laden_on_board": [12, 747, 383, 842],
        },
        "cargo": {"top": 296, "bottom": 544,
                  "columns": {"left": [12, 164], "no_pkgs": [164, 228], "description": [228, 393],
                              "gross_weight": [393, 491], "measurement": [491, 595]}},
    },
    "CLIENT (shipper print)": {
        "detect_label": "",
        "ignore_font_below": 6.0,
        "fields": {
            "bl_no": [370, 20, 595, 50],
            "shipper": [12, 15, 360, 87],
            "consignee": [12, 87, 360, 164],
            "notify": [12, 164, 360, 240],
            "pre_carriage": [12, 240, 214, 263],
            "place_receipt": [214, 240, 360, 263],
            "vessel_voy": [12, 263, 214, 286],
            "pol": [214, 263, 360, 286],
            "pod": [12, 286, 214, 315],
            "place_delivery": [214, 286, 390, 315],
            "final_dest": [390, 286, 595, 315],
            "total_words": [12, 551, 470, 573],
            "service_term": [470, 551, 595, 573],
            "freight": [12, 573, 595, 722],
            "ex_rate": [12, 722, 103, 748],
            "prepaid_at": [103, 722, 258, 748],
            "payable_at": [258, 722, 385, 748],
            "place_date_issue": [385, 722, 595, 748],
            "total_prepaid_local": [103, 748, 258, 773],
            "no_originals": [258, 748, 385, 773],
            "laden_on_board": [12, 773, 385, 842],
        },
        "cargo": {"top": 315, "bottom": 551,
                  "columns": {"left": [12, 183.5], "no_pkgs": [183.5, 226], "description": [226, 414],
                              "gross_weight": [414, 502], "measurement": [502, 595]}},
    },
}

# Words that belong to the printed form itself, never to the B/L data.
FORM_NOISE = {"BY", "AS CARRIER", "HEUNG A LINE CO., LTD."}

CONTAINER_RE = re.compile(r"\b([A-Z]{4})\s?(\d{6})(\d)\b")
BL_RE = re.compile(r"\b[A-Z]{3,6}\d{6,}\b")
NUM_RE = re.compile(r"\d[\d,]*\.?\d*")
DATE_RE = re.compile(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*[.\s/-]*(\d{1,2})[.,\s/-]*(\d{4})\b")
# "AS PER ATTACHED RIDER", "SEE ATTACHED SHEET", "AS PER ATTACHED DESCRIPTIONS", ...
ATTACH_RE = re.compile(
    r"\b(AS\s+PER|SEE|PL(?:EA)?SE?\s+SEE|PLS\s+SEE|DETAILS?\s+AS\s+PER|CONT(?:INUED|'?D)?\s+(?:ON|IN)|REFER\s+TO)"
    r"\s+(?:THE\s+)?(?:ATTACHED|ATTACHMENT|RIDER|ATTACH)"
    r"|\bATTACHED\s+(?:RIDER|SHEETS?|LISTS?|DESCRIPTIONS?|PAGES?|DETAILS?)\b")
# Lines on rider pages that are only page furniture, not B/L data.
RIDER_NOISE_RE = re.compile(
    r"^(PAGE\s*\d+(\s*(OF|/)\s*\d+)?|-?\s*\d+\s*-?|RIDER( PAGE)?|ATTACHED SHEET|ATTACHMENT|CONTINUED|CONT'D|"
    r"\(?CONTINUED\)?|HEUNG A LINE CO\., LTD\.|"
    # page titles such as "RIDER - B/L NO. HASLS22260901995" or "B/L NO.: HASLS22260901995"
    r"((RIDER|ATTACHED SHEET|ATTACHMENT|CONTINUATION SHEET)( PAGE)?\W*(TO\W*)?)?"
    r"(B/L|BILL OF LADING)\W*NO\.?\W*[A-Z]{3,6}\d{6,}|(RIDER|ATTACHED SHEET|ATTACHMENT)\W+[A-Z]{3,6}\d{6,})$")
WEIGHT_RE = re.compile(r"([\d,]+\.?\d*)\s*KGS?\b")
NUMBER_WORDS = {w: i for i, w in enumerate(
    "ZERO ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN ELEVEN TWELVE THIRTEEN FOURTEEN "
    "FIFTEEN SIXTEEN SEVENTEEN EIGHTEEN NINETEEN TWENTY".split())}


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_layouts():
    path = os.path.join(app_dir(), "layouts.json")
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_LAYOUTS, f, indent=2)
    except Exception:
        pass
    return DEFAULT_LAYOUTS


# --------------------------------------------------------------------------
# PDF reading
# --------------------------------------------------------------------------
def group_lines(words, tol=3.0):
    """Group words into text lines (top to bottom, left to right)."""
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(lines[-1][0]["top"] - w["top"]) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in lines]


def line_text(line):
    return " ".join(w["text"] for w in line)


def in_box(w, box):
    cx = (w["x0"] + w["x1"]) / 2
    cy = (w["top"] + w["bottom"]) / 2
    return box[0] <= cx < box[2] and box[1] <= cy < box[3]


def page_lines(page, min_size=6.0):
    """All text lines of a rider / attached page (tiny printed form labels dropped)."""
    words = [w for w in page.extract_words(extra_attrs=["size"]) if w.get("size", 99) >= min_size]
    lines = [line_text(l) for l in group_lines(words)]
    return [l for l in lines if l.strip() and not RIDER_NOISE_RE.match(norm(l))]


def read_words(pdf_path):
    """Words of page 1 (the B/L form) and the text lines of pages 2+ (rider)."""
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            raise ValueError("PDF has no pages")
        words = pdf.pages[0].extract_words(extra_attrs=["size"], keep_blank_chars=False)
        rider = []
        for pg in pdf.pages[1:]:
            rider += page_lines(pg)
        n_pages = len(pdf.pages)
    if not words:
        raise ValueError("No text found in the PDF (it may be a scanned image). "
                         "Please use a PDF with selectable text.")
    return words, rider, n_pages


def read_rider_file(pdf_path):
    """A rider sent as its own PDF: every page is rider text."""
    with pdfplumber.open(pdf_path) as pdf:
        lines = []
        for pg in pdf.pages:
            lines += page_lines(pg)
    return lines


def extract_with_layout(words, layout):
    min_size = layout.get("ignore_font_below", 0)
    words = [w for w in words if w.get("size", 99) >= min_size]
    out = {}
    for key, box in layout["fields"].items():
        lines = [line_text(l) for l in group_lines([w for w in words if in_box(w, box)])]
        out[key] = [l for l in lines if l.strip().upper() not in FORM_NOISE]

    # Cargo table: split each line into segments at wide gaps, then put every
    # segment in the column it belongs to.
    cargo = layout["cargo"]
    cols = cargo["columns"]
    order = list(cols.keys())
    table = {c: [] for c in order}
    cwords = [w for w in words if cargo["top"] <= (w["top"] + w["bottom"]) / 2 < cargo["bottom"]]
    for line in group_lines(cwords):
        segs, cur = [], [line[0]]
        for w in line[1:]:
            if w["x0"] - cur[-1]["x1"] > 6.5:
                segs.append(cur)
                cur = [w]
            else:
                cur.append(w)
        segs.append(cur)
        row = {}
        for seg in segs:
            x0, x1 = seg[0]["x0"], seg[-1]["x1"]
            col = order[-1]
            for i, c in enumerate(order):
                if cols[c][0] <= x0 + 2 < cols[c][1]:
                    col = c
                    # right-aligned figures that start just left of a column line
                    if c != "description" and i + 1 < len(order):
                        edge = cols[c][1]
                        if edge - x0 < 10 and (x1 - edge) > (edge - x0):
                            col = order[i + 1]
                    break
            row.setdefault(col, []).append(" ".join(w["text"] for w in seg))
        for c, parts in row.items():
            table[c].append("  ".join(parts))
    out["container_seal"] = [l for l in table["left"] if CONTAINER_RE.search(l.upper())]
    out["marks"] = [l for l in table["left"] if not CONTAINER_RE.search(l.upper())]
    for c in ("no_pkgs", "description", "gross_weight", "measurement"):
        out[c] = table.get(c, [])
    return out


def layout_score(fields, words, layout):
    score = 0
    label = layout.get("detect_label")
    if label and any(w["text"] == label and w.get("size", 99) < 6 for w in words):
        score += 5
    j = lambda k: " ".join(fields.get(k, [])).upper()
    score += 3 if BL_RE.fullmatch(j("bl_no").replace(" ", "")) else 0
    score += 2 if CONTAINER_RE.search(j("container_seal")) else 0
    score += 1 if "SAY" in j("total_words") or "ONLY" in j("total_words") else 0
    score += 1 if DATE_RE.search(j("place_date_issue")) else 0
    score += 1 if re.search(r"\d", j("gross_weight")) else 0
    score += 1 if fields.get("vessel_voy") else 0
    score += 1 if fields.get("shipper") else 0
    return score


def extract_bl(pdf_path, layouts=None, rider_pdf=None):
    layouts = layouts or load_layouts()
    words, rider, n_pages = read_words(pdf_path)
    if rider_pdf:
        rider += read_rider_file(rider_pdf)
    best = None
    for name, lay in layouts.items():
        f = extract_with_layout(words, lay)
        s = layout_score(f, words, lay)
        if best is None or s > best[0]:
            best = (s, name, f)
    refs = [FIELD_LABEL[k] for k, v in best[2].items() if attach_ref(v)]
    return {"file": pdf_path, "rider_file": rider_pdf, "layout": best[1], "score": best[0], "pages": n_pages,
            "fields": best[2], "rider": rider, "rider_refs": refs}


# --------------------------------------------------------------------------
# Comparison
# --------------------------------------------------------------------------
def norm(s):
    return re.sub(r"\s+", " ", s.upper()).strip()


def compact(s):
    return re.sub(r"[\s]", "", s.upper())


def numbers(s):
    vals = []
    for m in NUM_RE.findall(s):
        try:
            vals.append(float(m.replace(",", "")))
        except ValueError:
            pass
    return vals


def count_value(s):
    """'3 / THREE', 'THREE(3)', '3' -> 3"""
    s = s.upper()
    n = re.findall(r"\d+", s)
    if n:
        return int(n[0])
    for w, v in NUMBER_WORDS.items():
        if re.search(r"\b%s\b" % w, s):
            return v
    return None


def dates(s):
    return ["%s %02d %s" % (m[0], int(m[1]), m[2]) for m in DATE_RE.findall(s.upper())]


def token_diff(a_lines, b_lines):
    """Describe which words are only in CLIENT / only in DRAFT."""
    a = norm(" ".join(a_lines)).split()
    b = norm(" ".join(b_lines)).split()
    only_a, only_b = [], []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if op in ("replace", "delete"):
            only_a.append(" ".join(a[i1:i2]))
        if op in ("replace", "insert"):
            only_b.append(" ".join(b[j1:j2]))
    parts = []
    if only_a:
        parts.append("CLIENT has: " + " | ".join('"%s"' % x for x in only_a))
    if only_b:
        parts.append("DRAFT has: " + " | ".join('"%s"' % x for x in only_b))
    return "; ".join(parts)


MATCH, MISMATCH, ONLY_CLIENT, ONLY_DRAFT, EMPTY, FORMAT, VIA_RIDER = (
    "MATCH", "NOT MATCH", "MISSING IN DRAFT", "MISSING IN CLIENT", "BOTH EMPTY", "MATCH (format differs)",
    "MATCH (via rider)")


def attach_ref(lines):
    """The 'AS PER ATTACHED RIDER'-type phrase in these lines, or ''."""
    m = ATTACH_RE.search(norm(" ".join(lines)))
    return m.group(0) if m else ""


def compare_field(key, a_lines, b_lines):
    a, b = norm(" ".join(a_lines)), norm(" ".join(b_lines))
    if not a and not b:
        return EMPTY, ""
    if a and not b:
        return ONLY_CLIENT, 'CLIENT shows "%s" but DRAFT is blank' % " / ".join(a_lines)
    if b and not a:
        return ONLY_DRAFT, 'DRAFT shows "%s" but CLIENT is blank' % " / ".join(b_lines)
    if a == b:
        return MATCH, ""

    if key in ("gross_weight", "measurement"):
        na, nb = numbers(a), numbers(b)
        if na and na == nb:
            return FORMAT, "Same value, written differently"
        if na and nb:
            unit = "KGS" if key == "gross_weight" else "CBM"
            d = (na[0] - nb[0])
            return MISMATCH, "CLIENT %s vs DRAFT %s  (difference %s%s %s)" % (
                a, b, "+" if d > 0 else "", f"{d:,.3f}", unit)
    if key == "no_originals":
        ca, cb = count_value(a), count_value(b)
        if ca is not None and ca == cb:
            return FORMAT, 'Both = %d  ("%s" vs "%s")' % (ca, a, b)
        return MISMATCH, 'CLIENT "%s" (=%s) vs DRAFT "%s" (=%s)' % (a, ca, b, cb)
    if compact(a) == compact(b):
        return FORMAT, "Only spacing differs"
    return MISMATCH, token_diff(a_lines, b_lines)


def check_against_rider(key, ref_side, ref_lines, ref_rider, other_side, other_lines):
    """One B/L says 'AS PER ATTACHED RIDER' in this field: every detail the other
    B/L prints in the field must be found in this B/L's field text or its rider."""
    phrase = attach_ref(ref_lines)
    if not ref_rider:
        return MISMATCH, '%s says "%s" but NO rider / attached page was found in the %s PDF' % (
            ref_side, phrase, ref_side)
    pool = compact(" ".join(ref_lines + ref_rider))
    missing = [l for l in other_lines if compact(l) not in pool and not attach_ref([l])]
    if key in ("gross_weight", "measurement"):
        # compare the figures, not the spelling
        have = set(numbers(" ".join(ref_lines + ref_rider)))
        missing = [l for l in other_lines if not attach_ref([l]) and any(n not in have for n in numbers(l))]
    if not missing:
        return VIA_RIDER, '%s says "%s"; everything %s shows here is found in the %s rider' % (
            ref_side, phrase, other_side, ref_side)
    return MISMATCH, '%s says "%s"; NOT found in %s rider: %s' % (
        ref_side, phrase, ref_side, " | ".join('"%s"' % l for l in missing))


def rider_containers(lines):
    """{container no.: text of the rider line(s) that mention it}"""
    out = {}
    for l in lines:
        for m in CONTAINER_RE.finditer(norm(l)):
            out.setdefault("".join(m.groups()), []).append(norm(l))
    return out


def compare_rider(a, b, page1_a="", page1_b=""):
    """Compare the rider / attached-sheet text of CLIENT (a) and DRAFT (b).
    page1_a / page1_b: the page-1 text of each B/L - a rider detail that the other
    B/L prints on its page 1 instead of on a rider still counts as matched.
    Returns (status, summary, [ {client, draft, status, detail} ... ])."""
    if not a and not b:
        return EMPTY, "No rider / attached page on either B/L", []
    for side, other, lines, pool in (("CLIENT", "DRAFT", a, page1_b), ("DRAFT", "CLIENT", b, page1_a)):
        if lines and not (b if side == "CLIENT" else a):
            pool = compact(pool)
            missing = [norm(l) for l in lines if compact(l) not in pool]
            if not missing:
                return VIA_RIDER, "Only %s has a rider; all %d rider line(s) are printed on %s page 1" % (
                    side, len(lines), other), []
            st = ONLY_CLIENT if side == "CLIENT" else ONLY_DRAFT
            diffs = [{"client": l if side == "CLIENT" else "", "draft": l if side == "DRAFT" else "",
                      "status": st, "detail": "In %s rider, not found anywhere on %s" % (side, other)}
                     for l in missing]
            return st, "%s has NO rider; %d of %d %s rider line(s) not found on %s page 1: %s%s" % (
                other, len(missing), len(lines), side, other, " | ".join(missing[:3]),
                " ..." if len(missing) > 3 else ""), diffs
    ja, jb = compact(" ".join(a) + " " + page1_a), compact(" ".join(b) + " " + page1_b)
    if norm(" ".join(a)) == norm(" ".join(b)):
        return MATCH, "Rider text is the same (%d lines)" % len(a), []
    diffs = []
    if ja != jb:
        # lines of one side that cannot be found anywhere in the other side's rider
        # (searching the joined text, so a line wrapped differently still matches)
        only_a = [norm(l) for l in a if compact(l) not in jb]
        only_b = [norm(l) for l in b if compact(l) not in ja]
        # pair up lines that were changed rather than added / removed
        for la in only_a:
            best, score = None, 0.6
            for lb in only_b:
                r = difflib.SequenceMatcher(None, la, lb).ratio()
                if r > score:
                    best, score = lb, r
            if best:
                only_b.remove(best)
                diffs.append({"client": la, "draft": best, "status": MISMATCH, "detail": token_diff([la], [best])})
            else:
                diffs.append({"client": la, "draft": "", "status": ONLY_CLIENT, "detail": "Line not in DRAFT rider"})
        for lb in only_b:
            diffs.append({"client": "", "draft": lb, "status": ONLY_DRAFT, "detail": "Line not in CLIENT rider"})
        # container by container
        ca, cb = rider_containers(a), rider_containers(b)
        for c in sorted(set(ca) | set(cb)):
            if c not in cb:
                diffs.append({"client": c, "draft": "", "status": ONLY_CLIENT,
                              "detail": "Container %s is in CLIENT rider only" % c})
            elif c not in ca:
                diffs.append({"client": "", "draft": c, "status": ONLY_DRAFT,
                              "detail": "Container %s is in DRAFT rider only" % c})
        # total weight written on the rider
        wa = sum(float(x.replace(",", "")) for x in WEIGHT_RE.findall(norm(" ".join(a))) if x.replace(",", ""))
        wb = sum(float(x.replace(",", "")) for x in WEIGHT_RE.findall(norm(" ".join(b))) if x.replace(",", ""))
        if wa and wb and abs(wa - wb) > 0.0005:
            diffs.append({"client": f"{wa:,.3f} KGS", "draft": f"{wb:,.3f} KGS", "status": MISMATCH,
                          "detail": "Sum of all KGS figures on the rider differs by %s KGS" % f"{wa - wb:+,.3f}"})
    if not diffs:
        return FORMAT, "Same rider content, only line breaks / spacing differ", []
    first = "; ".join(d["detail"] or ("CLIENT: " + d["client"] if d["client"] else "DRAFT: " + d["draft"])
                      for d in diffs[:3])
    return MISMATCH, "%d difference(s) in rider - see RIDER CHECK. %s%s" % (
        len(diffs), first, " ..." if len(diffs) > 3 else ""), diffs


def iso6346_ok(code):
    """Check-digit test for a container number like SKLU1597799."""
    m = CONTAINER_RE.fullmatch(code)
    if not m:
        return None
    vals = {}
    v = 10
    for ch in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        if v % 11 == 0:
            v += 1
        vals[ch] = v
        v += 1
    body = m.group(1) + m.group(2)
    total = sum((vals[c] if c.isalpha() else int(c)) * (2 ** i) for i, c in enumerate(body))
    return (total % 11) % 10 == int(m.group(3))


def extra_checks(client, draft):
    """Sanity checks beyond field-by-field matching."""
    notes = []
    for side, doc in (("CLIENT", client), ("DRAFT", draft)):
        f = doc["fields"]
        txt = lambda k: norm(" ".join(f.get(k, [])))
        # container check digit
        for m in CONTAINER_RE.finditer(txt("container_seal")):
            code = "".join(m.groups())
            if not iso6346_ok(code):
                notes.append((side, "Container %s fails the ISO 6346 check digit - please verify the number" % code))
        # gross weight vs net weight in description
        gw = numbers(txt("gross_weight"))
        nw = re.search(r"N\.?\s?W\.?\s*[:]?\s*([\d,]+\.?\d*)", txt("description"))
        if gw and nw:
            nwv = float(nw.group(1).replace(",", ""))
            if gw[0] < nwv:
                notes.append((side, "GROSS WEIGHT %s KGS is LESS than N.W. %s KGS in the description" % (
                    f"{gw[0]:,.3f}", f"{nwv:,.3f}")))
        # "AS PER ATTACHED RIDER" but no rider pages
        if doc["rider_refs"] and not doc["rider"]:
            notes.append((side, '"AS PER ATTACHED" found in %s but the PDF has NO rider / attached page' %
                          ", ".join(doc["rider_refs"])))
        for m in CONTAINER_RE.finditer(norm(" ".join(doc["rider"]))):
            code = "".join(m.groups())
            if not iso6346_ok(code):
                notes.append((side, "Rider container %s fails the ISO 6346 check digit - please verify" % code))
        # containers listed (page 1 + rider) vs container count in words
        n_cntr = len(set("".join(m) for m in CONTAINER_RE.findall(txt("container_seal") + " " + norm(" ".join(doc["rider"])))))
        words = txt("total_words")
        cw = None
        for w, v in NUMBER_WORDS.items():
            if re.search(r"\b%s\b" % w, words):
                cw = v
                break
        if cw is not None and "CONTAINER" in words and n_cntr and cw != n_cntr:
            notes.append((side, "Total in words says %d container(s) but %d container number(s) are listed" % (cw, n_cntr)))
        # file name B/L no vs B/L no on the document
        bl = compact(txt("bl_no"))
        fname = os.path.basename(doc["file"]).upper()
        fbl = BL_RE.search(fname)
        if fbl and bl and fbl.group(0) != bl:
            notes.append((side, "File name has B/L %s but the document shows %s" % (fbl.group(0), bl)))
        # on-board date after issue date
        ob, iss = dates(txt("laden_on_board")), dates(txt("place_date_issue"))
        if ob and iss:
            try:
                d1 = datetime.strptime(ob[0], "%b %d %Y")
                d2 = datetime.strptime(iss[0], "%b %d %Y")
                if d1 > d2:
                    notes.append((side, "On-board date %s is AFTER the issue date %s" % (ob[0], iss[0])))
            except ValueError:
                pass
        if doc.get("score", 0) < 5:
            notes.append((side, "Form layout not recognised well (layout '%s') - results may be incomplete" % doc["layout"]))
    return notes


def check_pair(client_pdf, draft_pdf, client_rider=None, draft_rider=None):
    layouts = load_layouts()
    client = extract_bl(client_pdf, layouts, client_rider)
    draft = extract_bl(draft_pdf, layouts, draft_rider)
    rows = []
    for i, (key, label) in enumerate(FIELDS, 1):
        a, b = client["fields"].get(key, []), draft["fields"].get(key, [])
        status, detail = compare_field(key, a, b)
        ra, rb = attach_ref(a), attach_ref(b)
        if is_problem(status) and (ra or rb) and not (ra and rb):
            # one side moved this field to the rider: match it against that rider
            if ra:
                status, detail = check_against_rider(key, "CLIENT", a, client["rider"], "DRAFT", b)
            else:
                status, detail = check_against_rider(key, "DRAFT", b, draft["rider"], "CLIENT", a)
        rows.append({"no": i, "key": key, "field": label, "client": a, "draft": b,
                     "status": status, "detail": detail})
    page1 = lambda doc: " ".join(" ".join(v) for v in doc["fields"].values())
    status, detail, rider_diffs = compare_rider(client["rider"], draft["rider"], page1(client), page1(draft))
    refs = sorted(set(client["rider_refs"] + draft["rider_refs"]))
    if refs and status == EMPTY:
        status, detail = MISMATCH, '"AS PER ATTACHED" is written in %s but neither PDF has a rider page' % ", ".join(refs)
    elif refs:
        detail = "Referred to in: %s. %s" % (", ".join(refs), detail)
    rows.append({"no": len(FIELDS) + 1, "key": RIDER_ROW[0], "field": RIDER_ROW[1],
                 "client": client["rider"], "draft": draft["rider"], "status": status, "detail": detail})
    return {"client": client, "draft": draft, "rows": rows, "rider_diffs": rider_diffs,
            "notes": extra_checks(client, draft), "checked_at": datetime.now()}


def is_problem(status):
    return status in (MISMATCH, ONLY_CLIENT, ONLY_DRAFT)


# --------------------------------------------------------------------------
# Excel report
# --------------------------------------------------------------------------
FILL = {
    MATCH: "C6EFCE", FORMAT: "E2F0D9", EMPTY: "F2F2F2",
    MISMATCH: "FFC7CE", ONLY_CLIENT: "FFEB9C", ONLY_DRAFT: "FFEB9C", VIA_RIDER: "DDEBF7",
}


def cell_lines(r):
    """What the field cells show; the rider row is summarised (full text on RIDER CHECK)."""
    if r["key"] != RIDER_ROW[0]:
        return "\n".join(r["client"]), "\n".join(r["draft"])
    show = lambda ls: ("%d line(s) - full text on RIDER CHECK sheet\n%s%s" % (
        len(ls), "\n".join(ls[:4]), "\n..." if len(ls) > 4 else "")) if ls else ""
    return show(r["client"]), show(r["draft"])


def bl_number(result):
    for side in ("draft", "client"):
        bl = compact(" ".join(result[side]["fields"].get("bl_no", [])))
        if bl:
            return bl
    return "UNKNOWN"


def write_report(result, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "B-L CHECK"
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    bold = Font(bold=True)

    rows = result["rows"]
    n_bad = sum(is_problem(r["status"]) for r in rows)
    ws["A1"] = "B/L CHECK REPORT  -  B/L NO. %s" % bl_number(result)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "CLIENT B/L : %s   (layout: %s)" % (os.path.basename(result["client"]["file"]), result["client"]["layout"])
    ws["A3"] = "DRAFT HAL  : %s   (layout: %s)" % (os.path.basename(result["draft"]["file"]), result["draft"]["layout"])
    ws["A4"] = "Checked    : %s" % result["checked_at"].strftime("%d-%b-%Y %H:%M")
    ws["A5"] = ("RESULT: ALL %d FIELDS MATCH" % len(rows)) if n_bad == 0 else \
        "RESULT: %d FIELD(S) NOT MATCHED  -  see red / yellow rows" % n_bad
    ws["A5"].font = Font(bold=True, size=12, color="006100" if n_bad == 0 else "9C0006")

    head = ["NO.", "FIELD", "CLIENT B/L", "DRAFT HAL B/L", "RESULT", "DETAILS OF DIFFERENCE"]
    hr = 7
    for c, h in enumerate(head, 1):
        cell = ws.cell(hr, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for i, r in enumerate(rows):
        ctext, dtext = cell_lines(r)
        vals = [r["no"], r["field"], ctext, dtext, r["status"], r["detail"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(hr + 1 + i, c, v)
            cell.alignment = wrap
            cell.border = border
            cell.fill = PatternFill("solid", fgColor=FILL[r["status"]])
            if c in (2, 5):
                cell.font = bold
        lines = max(ctext.count("\n") + 1, dtext.count("\n") + 1, len(r["detail"]) // 45 + 1)
        ws.row_dimensions[hr + 1 + i].height = max(15, 13 * lines)
    for col, w in zip("ABCDEF", (5, 26, 45, 45, 17, 48)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = ws.cell(hr + 1, 1)

    r0 = hr + len(rows) + 2
    ws.cell(r0, 1, "OTHER WARNINGS").font = Font(bold=True, size=12)
    if result["notes"]:
        for i, (side, msg) in enumerate(result["notes"], 1):
            ws.cell(r0 + i, 1, i)
            ws.cell(r0 + i, 2, side).font = bold
            c = ws.cell(r0 + i, 3, msg)
            ws.merge_cells(start_row=r0 + i, start_column=3, end_row=r0 + i, end_column=6)
            for cc in range(1, 7):
                ws.cell(r0 + i, cc).fill = PatternFill("solid", fgColor="FCE4D6")
            c.alignment = wrap
    else:
        ws.cell(r0 + 1, 2, "None")

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "%d:%d" % (hr, hr)
    write_rider_sheet(wb, result, border, wrap)
    wb.save(out_path)
    return out_path


def write_rider_sheet(wb, result, border, wrap):
    a, b = result["client"]["rider"], result["draft"]["rider"]
    refs = sorted(set(result["client"]["rider_refs"] + result["draft"]["rider_refs"]))
    if not (a or b or refs):
        return
    ws = wb.create_sheet("RIDER CHECK")
    rider_row = result["rows"][-1]
    src = lambda doc, ls: ("%d line(s) from %s" % (
        len(ls), os.path.basename(doc.get("rider_file") or doc["file"]))) if ls else "NONE"
    ws["A1"] = "RIDER / ATTACHED SHEET CHECK  -  B/L NO. %s" % bl_number(result)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "CLIENT rider: " + src(result["client"], a)
    ws["A3"] = "DRAFT rider : " + src(result["draft"], b)
    ws["A4"] = "'AS PER ATTACHED' written in: %s" % (", ".join(refs) if refs else "-")
    ws["A5"] = "RESULT: %s  -  %s" % (rider_row["status"], rider_row["detail"])
    ws["A5"].font = Font(bold=True, size=12, color="9C0006" if is_problem(rider_row["status"]) else "006100")

    def header(row, titles):
        for c, h in enumerate(titles, 1):
            cell = ws.cell(row, c, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

    row = 7
    ws.cell(row, 1, "DIFFERENCES").font = Font(bold=True, size=12)
    row += 1
    header(row, ["NO.", "CLIENT RIDER", "DRAFT RIDER", "RESULT", "DETAILS"])
    diffs = result.get("rider_diffs", [])
    if not diffs:
        row += 1
        ws.cell(row, 2, "No differences")
    for i, d in enumerate(diffs, 1):
        row += 1
        for c, v in enumerate([i, d["client"], d["draft"], d["status"], d["detail"]], 1):
            cell = ws.cell(row, c, v)
            cell.alignment = wrap
            cell.border = border
            cell.fill = PatternFill("solid", fgColor=FILL[d["status"]])
    row += 2
    ws.cell(row, 1, "FULL RIDER TEXT (side by side)").font = Font(bold=True, size=12)
    row += 1
    header(row, ["LINE", "CLIENT RIDER", "DRAFT RIDER"])
    for i in range(max(len(a), len(b))):
        row += 1
        ws.cell(row, 1, i + 1).border = border
        for c, ls in ((2, a), (3, b)):
            cell = ws.cell(row, c, ls[i] if i < len(ls) else "")
            cell.alignment = wrap
            cell.border = border
    for col, w in zip("ABCDE", (6, 60, 60, 20, 50)):
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def default_report_path(result):
    folder = os.path.dirname(os.path.abspath(result["client"]["file"]))
    return os.path.join(folder, "BL CHECK %s.xlsx" % bl_number(result))


def summary_text(result):
    out = ["B/L NO. %s" % bl_number(result)]
    bad = [r for r in result["rows"] if is_problem(r["status"])]
    if not bad:
        out.append("ALL %d FIELDS MATCH" % len(result["rows"]))
    else:
        out.append("%d FIELD(S) NOT MATCHED:" % len(bad))
        for r in bad:
            out.append("  %2d. %s - %s: %s" % (r["no"], r["field"], r["status"], r["detail"]))
    for d in result.get("rider_diffs", []):
        out.append("      rider: %s | CLIENT: %s | DRAFT: %s" % (d["detail"] or d["status"], d["client"], d["draft"]))
    if result["notes"]:
        out.append("OTHER WARNINGS:")
        out += ["  [%s] %s" % n for n in result["notes"]]
    return "\n".join(out)


# --------------------------------------------------------------------------
# Folder (batch) mode
# --------------------------------------------------------------------------
def is_draft_name(name):
    u = os.path.basename(name).upper()
    return "DRAFT" in u or "HAL" in u.replace("HASL", "")


def is_rider_name(name):
    u = os.path.basename(name).upper()
    return "RIDER" in u or "ATTACH" in u


def find_pairs(folder):
    """[(client, draft, client_rider or None, draft_rider or None), ...] matched by B/L no.
    A separate rider PDF has RIDER / ATTACH in its name (plus DRAFT for the HAL one)."""
    pdfs = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
    found = {}
    for f in pdfs:
        m = BL_RE.search(f.upper())
        key = m.group(0) if m else f.upper()
        slot = ("draft" if is_draft_name(f) else "client") + ("_rider" if is_rider_name(f) else "")
        found.setdefault(key, {})[slot] = os.path.join(folder, f)
    return [(v["client"], v["draft"], v.get("client_rider"), v.get("draft_rider"))
            for k, v in sorted(found.items()) if "client" in v and "draft" in v]


def run_cli(args):
    outdir = None
    if "--outdir" in args:
        i = args.index("--outdir")
        outdir = args[i + 1]
        args = args[:i] + args[i + 2:]
        os.makedirs(outdir, exist_ok=True)
    pairs = []
    if len(args) == 1 and os.path.isdir(args[0]):
        pairs = find_pairs(args[0])
        if not pairs:
            print("No CLIENT/DRAFT PDF pairs found in", args[0])
            return 1
    elif len(args) == 2:
        pairs = [(args[0], args[1], None, None)]
    elif len(args) == 4:
        pairs = [(args[0], args[1], args[2] or None, args[3] or None)]
    else:
        print(__doc__)
        return 1
    for c, d, cr, dr in pairs:
        res = check_pair(c, d, cr, dr)
        path = default_report_path(res)
        if outdir:
            path = os.path.join(outdir, os.path.basename(path))
        try:
            write_report(res, path)
        except PermissionError:
            print("!! Cannot save %s - close it in Excel and run again" % path)
            path = "(not saved)"
        print(summary_text(res))
        print("Report:", path)
        print("-" * 70)
    return 0


# --------------------------------------------------------------------------
# Window
# --------------------------------------------------------------------------
def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("%s  v%s  -  Client B/L vs HAL Draft" % (APP_NAME, APP_VERSION))
    root.geometry("1250x720")
    root.minsize(900, 500)
    state = {"result": None}

    top = ttk.Frame(root, padding=8)
    top.pack(fill="x")
    client_var, draft_var = tk.StringVar(), tk.StringVar()
    crider_var, drider_var = tk.StringVar(), tk.StringVar()
    all_vars = (client_var, draft_var, crider_var, drider_var)

    def pick(var, title):
        p = filedialog.askopenfilename(title=title, filetypes=[("PDF files", "*.pdf")])
        if not p:
            return
        var.set(p)
        if var not in (client_var, draft_var):
            return
        # auto-fill the partner file (and separate rider PDFs) from the same folder
        idx = all_vars.index(var)
        for group in find_pairs(os.path.dirname(p)):
            if group[idx] and os.path.samefile(group[idx], p):
                for v, f in zip(all_vars, group):
                    if f and not v.get():
                        v.set(f)

    entries = (("CLIENT B/L (PDF):", client_var), ("DRAFT HAL B/L (PDF):", draft_var),
               ("CLIENT RIDER (optional):", crider_var), ("DRAFT RIDER (optional):", drider_var))
    for r, (lbl, var) in enumerate(entries):
        ttk.Label(top, text=lbl, width=24).grid(row=r, column=0, sticky="w", pady=2)
        ttk.Entry(top, textvariable=var).grid(row=r, column=1, sticky="ew", padx=4)
        ttk.Button(top, text="Browse...", command=lambda v=var, l=lbl: pick(v, l)).grid(row=r, column=2)
        if r >= 2:
            ttk.Button(top, text="Clear", command=lambda v=var: v.set("")).grid(row=r, column=3, padx=2)
    ttk.Label(top, foreground="#555555", text="Rider / attached pages inside the B/L PDF (page 2 onwards) are read "
              "automatically. Use the RIDER boxes only when the rider is a separate PDF file.").grid(
        row=4, column=0, columnspan=4, sticky="w")
    top.columnconfigure(1, weight=1)

    btns = ttk.Frame(root, padding=(8, 0))
    btns.pack(fill="x")
    status = tk.StringVar(value="Choose the two PDFs, then press CHECK.")
    status_lbl = tk.Label(root, textvariable=status, font=("Segoe UI", 12, "bold"), anchor="w", padx=10)
    status_lbl.pack(fill="x")

    body = ttk.PanedWindow(root, orient="vertical")
    body.pack(fill="both", expand=True, padx=8, pady=4)
    cols = ("no", "field", "client", "draft", "result")
    tree = ttk.Treeview(body, columns=cols, show="headings")
    for c, t, w in zip(cols, ("NO.", "FIELD", "CLIENT B/L", "DRAFT HAL B/L", "RESULT"), (40, 260, 360, 360, 170)):
        tree.heading(c, text=t)
        tree.column(c, width=w, anchor="w", stretch=c in ("client", "draft"))
    tree.tag_configure("bad", background="#FFC7CE")
    tree.tag_configure("miss", background="#FFEB9C")
    tree.tag_configure("ok", background="#C6EFCE")
    tree.tag_configure("fmt", background="#E2F0D9")
    tree.tag_configure("empty", background="#F2F2F2")
    tree.tag_configure("rider", background="#DDEBF7")
    sb = ttk.Scrollbar(tree, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    body.add(tree, weight=3)
    detail = tk.Text(body, height=10, wrap="word", font=("Consolas", 10))
    body.add(detail, weight=1)

    def show_detail(_=None):
        sel = tree.selection()
        res = state["result"]
        if not sel or not res:
            return
        r = res["rows"][int(sel[0])]
        detail.delete("1.0", "end")
        detail.insert("end", "%d. %s   ->   %s\n" % (r["no"], r["field"], r["status"]))
        if r["detail"]:
            detail.insert("end", "DIFFERENCE: %s\n" % r["detail"])
        if r["key"] == RIDER_ROW[0] and res.get("rider_diffs"):
            detail.insert("end", "\nRIDER DIFFERENCES:\n")
            for i, d in enumerate(res["rider_diffs"], 1):
                detail.insert("end", "%3d. %s\n     CLIENT: %s\n     DRAFT : %s\n" % (
                    i, d["detail"] or d["status"], d["client"] or "-", d["draft"] or "-"))
        detail.insert("end", "\n--- CLIENT B/L ---\n%s\n\n--- DRAFT HAL B/L ---\n%s\n" % (
            "\n".join(r["client"]) or "(blank)", "\n".join(r["draft"]) or "(blank)"))

    tree.bind("<<TreeviewSelect>>", show_detail)

    def do_check():
        c, d = client_var.get().strip(), draft_var.get().strip()
        if not (os.path.isfile(c) and os.path.isfile(d)):
            messagebox.showwarning(APP_NAME, "Please choose both PDF files.")
            return
        try:
            root.config(cursor="watch")
            root.update()
            res = check_pair(c, d, crider_var.get().strip() or None, drider_var.get().strip() or None)
        except Exception as e:
            messagebox.showerror(APP_NAME, "Could not read the PDFs:\n%s" % e)
            return
        finally:
            root.config(cursor="")
        state["result"] = res
        tree.delete(*tree.get_children())
        tagmap = {MISMATCH: "bad", ONLY_CLIENT: "miss", ONLY_DRAFT: "miss", MATCH: "ok", FORMAT: "fmt",
                  EMPTY: "empty", VIA_RIDER: "rider"}
        for i, r in enumerate(res["rows"]):
            tree.insert("", "end", iid=str(i), tags=(tagmap[r["status"]],), values=(
                r["no"], r["field"], " / ".join(r["client"])[:300], " / ".join(r["draft"])[:300], r["status"]))
        n_bad = sum(is_problem(r["status"]) for r in res["rows"])
        if n_bad:
            status.set("B/L %s :  %d FIELD(S) NOT MATCHED  (red / yellow rows)%s" % (
                bl_number(res), n_bad, "   +  %d warning(s)" % len(res["notes"]) if res["notes"] else ""))
            status_lbl.config(fg="#9C0006")
        else:
            status.set("B/L %s :  ALL FIELDS MATCH%s" % (
                bl_number(res), "   ( %d warning(s) )" % len(res["notes"]) if res["notes"] else ""))
            status_lbl.config(fg="#006100")
        detail.delete("1.0", "end")
        detail.insert("end", summary_text(res))

    def do_save():
        res = state["result"]
        if not res:
            messagebox.showinfo(APP_NAME, "Press CHECK first.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                         initialfile=os.path.basename(default_report_path(res)),
                                         initialdir=os.path.dirname(default_report_path(res)))
        if not p:
            return
        try:
            write_report(res, p)
        except PermissionError:
            messagebox.showerror(APP_NAME, "Cannot save - the file is open in Excel. Close it and try again.")
            return
        if messagebox.askyesno(APP_NAME, "Report saved:\n%s\n\nOpen it now?" % p):
            os.startfile(p)

    def do_folder():
        folder = filedialog.askdirectory(title="Folder with CLIENT and DRAFT PDFs")
        if not folder:
            return
        pairs = find_pairs(folder)
        if not pairs:
            messagebox.showinfo(APP_NAME, "No CLIENT / DRAFT pairs found.\nFile names must contain the B/L no.; "
                                          "the HAL file name must contain DRAFT.")
            return
        lines = []
        for c, d, cr, dr in pairs:
            try:
                res = check_pair(c, d, cr, dr)
                write_report(res, default_report_path(res))
                n = sum(is_problem(r["status"]) for r in res["rows"])
                lines.append("%s : %s" % (bl_number(res), "ALL MATCH" if n == 0 else "%d NOT MATCHED" % n))
            except Exception as e:
                lines.append("%s : ERROR %s" % (os.path.basename(c), e))
        messagebox.showinfo(APP_NAME, "Checked %d B/L(s). Reports saved in the folder.\n\n%s" % (
            len(pairs), "\n".join(lines)))

    ttk.Button(btns, text="CHECK", command=do_check, width=16).pack(side="left", pady=6)
    ttk.Button(btns, text="Save Excel report", command=do_save, width=20).pack(side="left", padx=6)
    ttk.Button(btns, text="Check whole folder...", command=do_folder, width=22).pack(side="left")
    ttk.Label(btns, text="Green = match   Blue = match via rider   Red = not match   Yellow = missing on one side"
              ).pack(side="right")

    if len(sys.argv) == 3:
        client_var.set(sys.argv[1])
        draft_var.set(sys.argv[2])
        root.after(200, do_check)
    root.mainloop()


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--cli"]
    if "--cli" in sys.argv or (len(args) == 1 and os.path.isdir(args[0])):
        sys.exit(run_cli(args))
    run_gui()
