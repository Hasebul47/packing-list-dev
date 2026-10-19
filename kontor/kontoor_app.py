"""
Kontoor Brands (Lee / Wrangler) PO  ->  Packing List
=====================================================
Module for the unified app. The router calls  run(pdf_file).
Also runs alone:  streamlit run kontoor_app.py

Layout follows the sample packing lists:
  A/B = case from/to | C = COLOR | D.. = sizes | then PCS IN CTNS, CTN QTY,
  TTL PCS, G.W, N.W, TTL/GW, TTL/NW.
Full cartons are numbered first, partial cartons after them.
"""
import datetime
import io
import re
from collections import OrderedDict

import sys
from pathlib import Path

import pdfplumber
import streamlit as st

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:                      # pl_kit.py lives in the project root
    sys.path.insert(0, _ROOT)
import pl_kit as kit                           # noqa: E402
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter as L

# ============================================================
# DEFAULTS  (NOT in the PO - from the sample sheets, editable in the UI)
# ============================================================
SIZE_RANK = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "2XL", "3XL", "4XL", "5XL"]
PACK_DEFAULT = {"XS": 34, "S": 34, "M": 34, "L": 30, "XL": 30, "XXL": 30}
DEFAULT_PACK = 20                                   # fallback for numeric / compound sizes
NETWT_DEFAULT = {"XS": 0.253, "S": 0.264, "M": 0.273, "L": 0.290, "XL": 0.308, "XXL": 0.333}
DEFAULT_NETWT = 0.30
CARTONS = OrderedDict([
    ("R13", (0.80, (51, 33, 32), "20X13X12.5 INCH")),
    ("R07", (0.60, (51, 33, 19), "20X13X7.5 INCH")),
    ("R04", (0.45, (51, 33, 9),  "20X13X3.5 INCH")),
])
R04_MAX_PCS, R07_MAX_PCS = 10, 20
MAX_SIZES = 40                                      # Lee waist/inseam sheets reach 39 sizes (style 112382674)
CHECKLIST = ["Test Report", "Po Sheet", "TOP", "CTL Comments", "Irregular", "25 CTN Check", "HL",
             "1 HL", "MMT Summary", "Barcode", "Moisture", "V-code & FWL", "Pre-Final(By Mail)", "AQL"]

# ---- derived Excel columns (kept in sync with MAX_SIZES) ----
C_SZ0 = 4                                           # D - first size column
C_SZL = C_SZ0 + MAX_SIZES - 1                       # last size column
C_PCS = C_SZL + 1                                   # PCS IN CTNS
C_CTN = C_PCS + 1                                   # CTN QTY
C_TPC = C_CTN + 1                                   # TTL PCS
C_GW  = C_TPC + 1                                   # G.W
C_NW  = C_GW  + 1                                   # N.W
C_TGW = C_NW  + 1                                   # TTL/GW
C_TNW = C_TGW + 1                                   # TTL/NW

# ============================================================
# PART 1 - PDF PARSING
# ============================================================
# Sizes: compound ("30 30"/"30/30"), plain numeric ("10"), or letter ("XS").
LINE_RE = re.compile(
    r"^(?P<color>[A-Z]\d+)\s+(?P<cdesc>.+?)\s+"
    r"(?P<size>\d{2,3}[/ ]\d{2,3}|\d{1,3}|[A-Z]{1,4})\s+"
    r"(?P<upc>\d{12,13})\s+(?P<oq>[\d,]+)\s+(?P<cq>[\d,]+)\s+(?P<sq>[\d,]+)\s+"
    r"\$(?P<cost>[\d.,]+)\s+\$(?P<total>[\d.,]+)\s*$", re.M)
STYLE_RE      = re.compile(r"Style:\s*(\S+)")
STYLE_DESC_RE = re.compile(r"Style Description:\s*(.+?)(?=\s*(?:COMMERCIAL|Style:|$))", re.S)
GOODS_RE      = re.compile(r"COMMERCIAL GOODS\s*DESCRIPTION:\s*(.+?)(?=\n|$)", re.S)


def _num(s):
    return int(str(s).replace(",", ""))


def _norm_size(s):
    """'30 30' -> '30/30' so it renders like the sample sheets."""
    return re.sub(r"\s+", "/", s.strip())


def _size_sort(sizes):
    def key(s):
        return (SIZE_RANK.index(s), 0) if s in SIZE_RANK else (len(SIZE_RANK), sizes.index(s))
    return sorted(sizes, key=key)


def _style_blocks(text):
    """Yield (style_no, block_text) for each 'Style: <no>' occurrence."""
    marks = list(STYLE_RE.finditer(text))
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        yield m.group(1), text[m.end():end]


def parse_po(pdf_bytes: bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        p1 = pdf.pages[0]
        dest_lines = []
        hit  = [w for w in p1.extract_words() if w["text"] == "Shipping"]
        stop = [w for w in p1.extract_words() if w["text"] == "Master"]
        if hit:
            x0, top = hit[0]["x0"] - 4, hit[0]["top"] + 8
            bottom = stop[0]["top"] - 2 if stop else top + 110
            crop = p1.crop((max(x0, 0), top, p1.width, bottom)).extract_text() or ""
            dest_lines = [ln.strip() for ln in crop.splitlines() if ln.strip()]

    def grab(pat, default=""):
        m = re.search(pat, text)
        return m.group(1).strip() if m else default

    crd = grab(r"Current CRD Date:\s*(\d\d/\d\d/\d{4})")
    try:
        crd_dt = datetime.datetime.strptime(crd, "%m/%d/%Y")
    except ValueError:
        crd_dt = None

    head = {
        "po": grab(r"\bPO:\s*(\d+)"),
        "brand": grab(r"BRAND DESC:\s*(.+?)\s+PRODUCT CATEGORY"),
        "over_tol": grab(r"Over Ship Tolerance:\s*(\d+)", "0"),
        "under_tol": grab(r"Under Ship Tolerance:\s*(\d+)", "0"),
        "crd": crd_dt, "crd_text": crd,
        "dest": dest_lines,
        "total_qty": grab(r"PO TOTAL QUANTITY:\s*([\d,]+)"),
    }

    groups = []
    for style, block in _style_blocks(text):
        sd = STYLE_DESC_RE.search(block)
        style_desc = sd.group(1).strip() if sd else ""
        gd = GOODS_RE.search(block)
        goods = gd.group(1).strip() if gd else ""

        by_color = OrderedDict()
        for ln in LINE_RE.finditer(block):
            c = by_color.setdefault(ln["color"], {"color": ln["color"], "cdesc": ln["cdesc"], "lines": []})
            c["lines"].append({"size": _norm_size(ln["size"]), "upc": ln["upc"],
                               "qty": _num(ln["cq"]), "orig": _num(ln["oq"]),
                               "cost": ln["cost"]})
        for c in by_color.values():
            sizes = _size_sort([x["size"] for x in c["lines"]])
            order = {x["size"]: x["qty"] for x in c["lines"]}
            groups.append({
                "style": style, "style_desc": style_desc, "goods": goods,
                "color": c["color"], "cdesc": c["cdesc"], "sizes": sizes, "order": order,
                "upc": {x["size"]: x["upc"] for x in c["lines"]},
            })
    return head, groups


# ============================================================
# PART 2 - CARTON LOGIC
# ============================================================
def carton_type(pcs):
    return "R04" if pcs <= R04_MAX_PCS else ("R07" if pcs <= R07_MAX_PCS else "R13")


def default_rows(sizes, order, pack):
    rows = []
    for sz in sizes:
        qty, p = int(order.get(sz, 0)), max(int(pack.get(sz, DEFAULT_PACK)), 1)
        full, rem = divmod(qty, p)
        if full:
            rows.append({"Kind": "Full", "Carton": "Auto", "sizes": {sz: p}, "CTN QTY": full})
        if rem:
            rows.append({"Kind": "Partial", "Carton": "Auto", "sizes": {sz: rem}, "CTN QTY": 1})
    return rows


def calc_rows(rows, sizes, netwt):
    out = []
    for r in rows:
        sz = {s: int(r["sizes"].get(s, 0) or 0) for s in sizes}
        pcs = sum(sz.values())
        ctn = int(r.get("CTN QTY") or 0)
        ctype = r.get("Carton") if r.get("Carton") in CARTONS else carton_type(pcs)
        empty = CARTONS[ctype][0]
        net_raw = sum(sz[s] * netwt.get(s, DEFAULT_NETWT) for s in sizes)
        out.append({"kind": r.get("Kind") or "Partial", "ctype": ctype, "sizes": sz, "pcs": pcs,
                    "ctn": ctn, "total_pcs": pcs * ctn, "gw": round(net_raw + empty, 2), "empty": empty})
    return out


def ship_by_size(calc, sizes):
    return {s: sum(r["sizes"][s] * r["ctn"] for r in calc) for s in sizes}


# ============================================================
# PART 3 - EXCEL
# ============================================================
def build_workbook(head, groups_data):
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for gd in groups_data:
        name = re.sub(r"[\[\]:*?/\\]", "-", str(gd["group"]["style"]))[:31]
        if len(groups_data) > 1 and sum(1 for g in groups_data
                                        if g["group"]["style"] == gd["group"]["style"]) > 1:
            name = f"{gd['group']['style']}_{gd['group']['color']}"[:31]
        base, i = name, 1
        while name in used:
            i += 1
            name = f"{base[:28]}_{i}"
        used.add(name)
        _sheet(wb.create_sheet(title=name), head, gd["group"], gd["calc"], gd["cut"])
    return wb


def _sheet(ws, head, g, calc, cut):
    sizes = g["sizes"][:MAX_SIZES]
    TNR, ARI, VER = "Times New Roman", "Arial", "Verdana"
    thin, med = Side(style="thin"), Side(style="medium")
    CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center")

    def put(r, c, v=None, name=ARI, sz=48, al=CEN, fmt=None, bold=True, u=None):
        cell = ws.cell(row=r, column=c)
        if v is not None:
            cell.value = v
        cell.font = Font(name=name, size=sz, bold=bold, underline=u)
        cell.alignment = al
        if fmt:
            cell.number_format = fmt
        return cell

    def merge(r1, c1, r2, c2):
        if (r1, c1) != (r2, c2):
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    def grid(r1, c1, r2, c2, side=thin):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).border = Border(left=side, right=side, top=side, bottom=side)

    def outline(r1, c1, r2, c2, side=med):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                b = ws.cell(row=r, column=c).border
                ws.cell(row=r, column=c).border = Border(
                    left=side if c == c1 else b.left, right=side if c == c2 else b.right,
                    top=side if r == r1 else b.top, bottom=side if r == r2 else b.bottom)

    # column widths (derived so they follow MAX_SIZES)
    widths = {1: 79.7, 2: 85, 3: 84.3}
    for i in range(MAX_SIZES):
        widths[C_SZ0 + i] = 38.1
    widths.update({C_PCS: 31.3, C_CTN: 43.1, C_TPC: 53.1, C_GW: 38.1,
                   C_NW: 40.4, C_TGW: 68.7, C_TNW: 48.8})
    for c, w in widths.items():
        ws.column_dimensions[L(c)].width = w

    R_DATA = 16
    R_END = R_DATA + len(calc) - 1
    R_TOT = R_END + 2
    R_SUMT = R_TOT + 2
    R_SIZE = R_SUMT + 1
    R_ORD, R_CUT, R_SHIP, R_EXS, R_PCT = R_SIZE + 1, R_SIZE + 2, R_SIZE + 3, R_SIZE + 4, R_SIZE + 5
    R_MARK = R_PCT + 7
    R_LAST = max(R_MARK + 7, R_PCT + 3 + len(CHECKLIST))

    heights = {1: 75.75, 2: 10.5, 3: 58.5, 4: 58.5, 5: 78.75, 6: 81, 7: 81, 8: 81, 9: 81, 10: 81,
               11: 126, 12: 133.5, 13: 18, 14: 75, 15: 144, R_TOT - 1: 74.25, R_TOT: 89.25,
               R_TOT + 1: 89.25, R_SUMT: 80.25, R_SIZE: 153, R_ORD: 122.25, R_CUT: 94.5,
               R_SHIP: 94.5, R_EXS: 94.5, R_PCT: 94.5}
    for r in range(R_DATA, R_END + 1):
        heights[r] = 93.75
    for r in range(R_PCT + 1, R_LAST + 1):
        heights[r] = 78
    for r, h in heights.items():
        ws.row_dimensions[r].height = h

    # ---- title + header block
    merge(1, 1, 1, 16)
    put(1, 1, "PACKING LIST DETAILS", TNR, al=Alignment(horizontal="center", vertical="bottom"))
    L_LBL = lambda r, t: put(r, 1, t, TNR, al=LEFT)
    L_LBL(6, "BUYER                :"); put(6, 2, "KONTOOR", TNR, al=LEFT)
    L_LBL(7, "STYLE                :"); put(7, 2, g["style"], TNR, al=LEFT)
    L_LBL(8, "P.O                      :")
    put(8, 2, int(head["po"]) if head["po"].isdigit() else head["po"], TNR, al=LEFT)
    L_LBL(9, "DESCRIPTION :"); put(9, 2, g["style_desc"], TNR, al=LEFT)
    put(10, 1, "Goods Description", TNR, al=Alignment(horizontal="left", vertical="bottom"))
    put(10, 2, g["goods"], TNR, al=Alignment(horizontal="left", vertical="bottom"))
    put(11, 1, "Overship Tolerance:", ARI, al=Alignment(horizontal="left", vertical="center", wrap_text=True))
    put(11, 2, int(head["over_tol"]), ARI, al=LEFT)
    put(12, 1, "Undership Tolerance:", ARI, al=Alignment(horizontal="left", vertical="center", wrap_text=True))
    put(12, 2, int(head["under_tol"]), ARI, al=LEFT)
    put(11, C_PCS, "CARTOON TYPE :", ARI, 72, al=Alignment(vertical="center"))
    put(11, C_PCS + 3, "/".join(CARTONS.keys()), ARI, 72, al=Alignment(vertical="center"))

    put(5, C_PCS, "CRD DATE", TNR, al=Alignment(vertical="center"))
    put(5, C_PCS + 2, head["crd"] or head["crd_text"], TNR, al=LEFT, fmt="m/d/yyyy")
    for r, t in ((6, "ORDER QTY"), (7, "PACK QTY"), (8, "EXCESS QTY"),
                 (9, "PERCENTAGE"), (10, "CARTON QTY")):
        put(r, C_PCS, t, ARI, al=Alignment(vertical="center"))
    put(6, C_PCS + 2, f"=+{L(C_TPC)}{R_ORD}", TNR, al=LEFT)
    put(7, C_PCS + 2, f"=+{L(C_TPC)}{R_TOT}", TNR, al=LEFT)
    put(8, C_PCS + 2, f"={L(C_PCS + 2)}7-{L(C_PCS + 2)}6", TNR, al=LEFT)
    merge(9, C_PCS + 2, 9, C_PCS + 3)
    put(9, C_PCS + 2, f"=IFERROR({L(C_PCS + 2)}8/{L(C_PCS + 2)}6,0)", TNR, al=LEFT, fmt="0.00%")
    put(10, C_PCS + 2, f"=+{L(C_CTN)}{R_TOT}", TNR, al=LEFT)
    outline(5, C_PCS, 10, C_TNW)

    # ---- table header
    merge(14, 1, 15, 2); put(14, 1, "Case SL#")
    put(14, 3, "COLOR"); put(15, 3, g["cdesc"], ARI, 36)
    merge(14, C_SZ0, 14, C_SZL); put(14, C_SZ0, "SIZE ")
    for i, s in enumerate(sizes):
        put(15, C_SZ0 + i, s, VER, fmt="@")
    for c, t in ((C_PCS, "PCS IN CTNS"), (C_CTN, "CTN \nQTY"), (C_TPC, "TTL PCS"),
                 (C_GW, "G.W"), (C_NW, "N.W")):
        merge(14, c, 15, c); put(14, c, t)
    put(15, C_TGW, "TTL/GW"); put(15, C_TNW, "TTL/NW")
    grid(14, 1, 15, C_TNW)

    # ---- data rows
    full_rows = [i for i, r in enumerate(calc) if r["kind"] == "Full"]
    part_rows = [i for i, r in enumerate(calc) if r["kind"] != "Full"]
    last_full_excel = R_DATA + full_rows[-1] if full_rows else None
    for idx, r in enumerate(calc):
        xr = R_DATA + idx
        if r["kind"] == "Full":
            pos = full_rows.index(idx)
            a = 1 if pos == 0 else f"=+B{R_DATA + full_rows[pos - 1]}+1"
        else:
            pos = part_rows.index(idx)
            a = (f"=+B{last_full_excel}+1" if last_full_excel else 1) if pos == 0 \
                else f"=+B{R_DATA + part_rows[pos - 1]}+1"
        put(xr, 1, a, TNR, fmt="0")
        put(xr, 2, f"=A{xr}+{L(C_CTN)}{xr}-1", TNR, fmt="0")
        for i, s in enumerate(sizes):
            v = r["sizes"].get(s)
            put(xr, C_SZ0 + i, v if v else None, ARI)
        put(xr, C_PCS, f"=SUM({L(C_SZ0)}{xr}:{L(C_SZL)}{xr})")
        put(xr, C_CTN, r["ctn"])
        put(xr, C_TPC, f"=+{L(C_CTN)}{xr}*{L(C_PCS)}{xr}")
        put(xr, C_GW,  r["gw"], fmt="0.00")
        put(xr, C_NW,  f"={L(C_GW)}{xr}-{r['empty']}", fmt="0.00")
        put(xr, C_TGW, f"={L(C_GW)}{xr}*{L(C_CTN)}{xr}", fmt="0.00")
        put(xr, C_TNW, f"=+{L(C_CTN)}{xr}*{L(C_NW)}{xr}", fmt="0.00")
        put(xr, 3, None)
    grid(R_DATA, 1, R_END, C_TNW)
    outline(14, 1, R_END, C_TNW)

    # ---- G.TTL
    merge(R_TOT, 1, R_TOT, C_PCS); put(R_TOT, 1, "G.TTL", TNR)
    for c in (C_CTN, C_TPC, C_GW, C_NW, C_TGW, C_TNW):
        put(R_TOT, c, f"=SUM({L(c)}{R_DATA}:{L(c)}{R_END})",
            fmt="0.00" if c >= C_GW else None)
    grid(R_TOT, 1, R_TOT, C_TNW)

    # ---- summary
    merge(R_SUMT, 1, R_SUMT, 14); put(R_SUMT, 1, "SUMMARY: COLOUR & SIZE WISE BREAKDOWN", TNR)
    put(R_SIZE, 1, "SIZE", TNR); merge(R_SIZE, 2, R_SIZE, 3)
    for i, s in enumerate(sizes):
        put(R_SIZE, C_SZ0 + i, f"=+{L(C_SZ0 + i)}15", VER, fmt="@")
    put(R_SIZE, C_TPC, "Total Pcs", TNR)
    for r, t, sz in ((R_ORD, "ORDER QTY", 40), (R_CUT, "CUT QTY", 40), (R_SHIP, "SHIP QTY", 40),
                     (R_EXS, "EXCESS/ SHORT", 40), (R_PCT, "%", 48)):
        put(r, 1, t, TNR, sz, al=LEFT if r != R_PCT else CEN)
    merge(R_EXS, 1, R_EXS, 3)
    first, last = L(C_SZ0), L(C_SZL)
    for i, s in enumerate(sizes):
        c = C_SZ0 + i; cl = L(c)
        put(R_ORD,  c, g["order"].get(s, 0), VER)
        put(R_CUT,  c, cut.get(s) or None, VER)
        put(R_SHIP, c, f"=SUMPRODUCT({cl}${R_DATA}:{cl}${R_END},"
                       f"${L(C_CTN)}${R_DATA}:${L(C_CTN)}${R_END})", VER)
        put(R_EXS,  c, f"=+{cl}{R_SHIP}-{cl}{R_ORD}", VER)
        put(R_PCT,  c, f"=IFERROR({cl}{R_EXS}/{cl}{R_ORD},0)", TNR, fmt="0.0%")
    tp = L(C_TPC)
    for r in (R_ORD, R_CUT, R_SHIP, R_EXS):
        put(r, C_TPC, f"=SUM({first}{r}:{L(C_TPC - 1)}{r})", TNR, 72)
    put(R_PCT, C_TPC, f"=IFERROR({tp}{R_EXS}/{tp}{R_ORD},0)", TNR, 72, fmt="0.00%")
    grid(R_SIZE, 1, R_PCT, C_TPC); outline(R_SUMT, 1, R_PCT, C_TPC)

    # ---- shipping marks / carton measure / CBM
    put(R_MARK, 2, "SHIPPING MARKS", ARI, al=LEFT, u="single")
    for i, ln in enumerate((head["dest"] or [])[:6]):
        put(R_MARK + 1 + i, 2, ln, VER, al=LEFT, fmt="@")
    for i, (k, (_, _, label)) in enumerate(CARTONS.items()):
        put(R_MARK + 1 + i, C_PCS, f"CTN MEASUR :{label}", ARI,
            al=Alignment(horizontal="general", vertical="bottom"))
    put(R_MARK + 4, C_PCS, "CBM:", TNR, al=LEFT, fmt="0.00")
    terms = []
    for k, (_, (a, b, c), _) in CARTONS.items():
        cells = [f"{L(C_CTN)}{R_DATA + i}" for i, r in enumerate(calc) if r["ctype"] == k]
        if cells:
            terms.append(f"{a}*{b}*{c}*SUM({','.join(cells)})/1000000")
    put(R_MARK + 4, C_PCS + 1, "=" + "+".join(terms) if terms else 0, ARI, al=LEFT, fmt="0.00")

    # ---- checklist
    r0 = R_PCT + 3
    merge(r0, 13, r0, 14); put(r0, 13, "Particulars", ARI, 70)
    put(r0, 15, "OK", TNR, 70); put(r0, 16, "NOT OK", TNR, 55)
    for i, t in enumerate(CHECKLIST):
        merge(r0 + 1 + i, 13, r0 + 1 + i, 14)
        put(r0 + 1 + i, 13, t, ARI, 55, al=Alignment(horizontal="center", vertical="bottom"))
    grid(r0, 13, r0 + len(CHECKLIST), 16)

    ws.print_area = f"A1:{L(C_TNW)}{R_LAST}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.25


# ============================================================
# PART 4 - STREAMLIT UI
# ============================================================
@st.cache_data(show_spinner=False, max_entries=10)
def _parse_cached(b: bytes):
    return parse_po(b)


def _group_ui(head, g, key):
    import pandas as pd
    sizes = g["sizes"]
    if len(sizes) > MAX_SIZES:
        st.error(f"{len(sizes)} sizes found; the sheet supports {MAX_SIZES}.")
        sizes = sizes[:MAX_SIZES]
    order = g["order"]

    # ---- 1 summary
    kit.step(1, "PO summary")
    kit.summary_cards([("Style", g["style"]), ("Color", f"{g['color']}"),
                       ("Order pcs", f"{sum(order.get(s, 0) for s in sizes):,}"),
                       ("Sizes", f"{len(sizes)}"),
                       ("Tolerance", f"+{head['over_tol']}% / -{head['under_tol']}%")])
    st.caption(f"{g.get('style_desc', '')} · {g.get('cdesc', '')} · {g.get('goods', '')}")

    # ---- 2 variables
    kit.step(2, "Packing settings", "Not in the PO. Pieces per full carton, net weight per piece (kg) and cut qty, per size.")
    cols = st.columns(min(len(sizes), 10))
    pack, netwt, cut = {}, {}, {}
    for i, s in enumerate(sizes):
        with cols[i % len(cols)]:
            st.markdown(f"**{s}**")
            pack[s] = st.number_input("Pcs / carton", 1, 500, PACK_DEFAULT.get(s, DEFAULT_PACK), 1, key=f"{key}_pk_{s}")
            netwt[s] = st.number_input("Net kg / pc", 0.0, 5.0, NETWT_DEFAULT.get(s, DEFAULT_NETWT), 0.001,
                                       format="%.3f", key=f"{key}_nw_{s}")
            cut[s] = st.number_input("Cut qty", 0, 1000000, 0, 1, key=f"{key}_cut_{s}")

    # ---- 3 carton table
    kit.step(3, "Carton table",
             "Edit any cell, add a row with +, or tick 🗑 Delete and press “Remove ticked rows”. "
             "Changing “Pcs / carton” above rebuilds this table.")
    sig = "_".join(str(pack[s]) for s in sizes)
    base = default_rows(sizes, order, pack)
    df0 = pd.DataFrame([{"Kind": r["Kind"], "Carton": r["Carton"],
                         **{s: r["sizes"].get(s, 0) for s in sizes}, "CTN QTY": r["CTN QTY"]} for r in base])
    cfg = {"Kind": st.column_config.SelectboxColumn("Kind", options=["Full", "Partial"], required=True),
           "Carton": st.column_config.SelectboxColumn("Carton", options=["Auto"] + list(CARTONS.keys()), required=True),
           "CTN QTY": st.column_config.NumberColumn("CTN QTY", min_value=0, step=1)}
    for s in sizes:
        cfg[s] = st.column_config.NumberColumn(s, min_value=0, step=1)
    edited = kit.editable_grid(df0, f"{key}_ed_{sig}", column_config=cfg)

    rows = []
    for _, r in edited.iterrows():
        if r.get("CTN QTY") is None or str(r.get("CTN QTY")) == "nan":
            continue
        rows.append({"Kind": r.get("Kind") or "Partial", "Carton": r.get("Carton") or "Auto",
                     "sizes": {s: (0 if str(r.get(s)) == "nan" else int(r.get(s) or 0)) for s in sizes},
                     "CTN QTY": int(r["CTN QTY"])})
    calc = calc_rows(rows, sizes, netwt)
    ship = ship_by_size(calc, sizes)

    # ---- 4 check
    kit.step(4, "Check against the PO")
    tot_ctn = sum(r["ctn"] for r in calc)
    tot_pcs = sum(ship.values())
    gw = sum(r["gw"] * r["ctn"] for r in calc)
    nw = sum((r["gw"] - r["empty"]) * r["ctn"] for r in calc)
    kit.summary_cards([("Cartons", f"{tot_ctn:,}"), ("Ship pcs", f"{tot_pcs:,}"),
                       ("Order pcs", f"{sum(order.get(s, 0) for s in sizes):,}"),
                       ("Gross kg", f"{gw:,.1f}"), ("Net kg", f"{nw:,.1f}")])
    kit.check_by_size(sizes, {s: order.get(s, 0) for s in sizes}, ship,
                      over_pct=float(head["over_tol"]), under_pct=float(head["under_tol"]))

    g2 = dict(g, sizes=sizes)
    return {"group": g2, "calc": calc, "cut": cut}


def run(pdf_file=None):
    st.title("Kontoor / Lee - PO to Packing List")
    if pdf_file is None:
        st.info("Upload a Kontoor Brands purchase order PDF.")
        return
    try:
        head, groups = _parse_cached(pdf_file.getvalue())
    except Exception as e:
        st.error(f"Could not read the PDF: {e}")
        return
    if not groups:
        st.warning("No style / size lines were found - the PDF layout may differ.")
        return

    st.success(f"PO {head['po']} - brand {head['brand']} - {len(groups)} style/colour group(s) - "
               f"CRD {head['crd_text']}")
    if head["dest"]:
        st.caption("Ship to: " + " | ".join(head["dest"]))

    results = []
    if len(groups) == 1:
        results.append(_group_ui(head, groups[0],
                                 f"kt_{head['po']}_{groups[0]['style']}_{groups[0]['color']}"))
    else:
        tabs = st.tabs([f"{g['style']} {g['color']}" for g in groups])
        for tab, g in zip(tabs, groups):
            with tab:
                results.append(_group_ui(head, g, f"kt_{head['po']}_{g['style']}_{g['color']}"))

    kit.step(5, "Download", "One Excel file with a sheet for every style / colour.")
    wb = build_workbook(head, results)
    buf = io.BytesIO()
    wb.save(buf)
    kit.download_block("⬇ Download packing list (Excel)", buf.getvalue(), f"packing_list_{head['po']}.xlsx",
                       "Cut qty and carton weights are not in the PO - they use the values in 'Packing settings'.")


if __name__ == "__main__":
    st.set_page_config(page_title="Kontoor PO -> Packing List", layout="wide")
    run(st.file_uploader("Upload Kontoor PO PDF", type=["pdf"]))