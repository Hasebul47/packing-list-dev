
"""
Kontoor Brands (Lee / Wrangler) PO  ->  Packing List
=====================================================
Module for the unified app. The router calls  run(pdf_file).
Also runs alone:  streamlit run kontoor_app.py
"""
import datetime
import io
import re
from collections import OrderedDict

import sys
from pathlib import Path

import pdfplumber
import streamlit as st
import pandas as pd

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
import pl_kit as kit
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

# ============================================================
# DEFAULTS
# ============================================================
SIZE_RANK = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "2XL", "3XL", "4XL", "5XL"]
PACK_DEFAULT = {"XS": 34, "S": 34, "M": 34, "L": 30, "XL": 30, "XXL": 30}
DEFAULT_PACK = 20
NETWT_DEFAULT = {"XS": 0.253, "S": 0.2643, "M": 0.2725, "L": 0.290, "XL": 0.3075, "XXL": 0.333}
DEFAULT_NETWT = 0.30
# N.N.W. (Net-Net Weight) defaults = N.W. minus polybag / packaging allowance (~0.03 kg)
NNWT_DEFAULT = {s: round(v - 0.03, 4) for s, v in NETWT_DEFAULT.items()}
DEFAULT_NNWT = 0.27
CARTONS = OrderedDict([
    ("R13", (0.80, (51, 33, 32), "20X13X12.5 INCH")),
    ("R07", (0.60, (51, 33, 19), "20X13X7.5 INCH")),
    ("R04", (0.45, (51, 33, 9),  "20X13X3.5 INCH")),
])
R04_MAX_PCS, R07_MAX_PCS = 10, 20
MAX_SIZES = 40
SHEET_ZOOM = 15         # opening zoom % (~whole sheet on one screen; LibreOffice clamps to 20)
CHECKLIST = ["Test Report", "Po Sheet", "TOP", "CTL Comments", "Irregular", "25 CTN Check", "HL",
             "1 HL", "MMT Summary", "Barcode", "Moisture", "V-code & FWL", "Pre-Final(By Mail)", "AQL"]

C_SZ0 = 4
C_SZL = C_SZ0 + MAX_SIZES - 1
C_PCS = C_SZL + 1
C_CTN = C_PCS + 1
C_TPC = C_CTN + 1
C_GW  = C_TPC + 1
C_NW  = C_GW  + 1
C_TGW = C_NW  + 1
C_TNW = C_TGW + 1

# ============================================================
# PART 1 - PDF PARSING
# ============================================================
LINE_RE = re.compile(
    r"^(?P<color>[A-Z]\d+)\s+(?P<cdesc>.+?)\s+"
    r"(?P<size>\d{2,3}[/ ]\d{2,3}|\d{1,3}|[A-Z]{1,4})\s+"
    r"(?P<upc>\d{12,13})\s+(?P<oq>[\d,]+)\s+(?P<cq>[\d,]+)\s+(?P<sq>[\d,]+)\s+"
    r"\$(?P<cost>[\d.,]+)\s+\$(?P<total>[\d.,]+)\s*$", re.M)
STYLE_RE      = re.compile(r"Style:\s*(\S+)")
STYLE_DESC_RE = re.compile(r"Style Description:\s*(.+?)(?=\s*(?:COMMERCIAL|Style:|$))", re.S)
GOODS_RE      = re.compile(r"COMMERCIAL GOODS[ \t]+([^\n]+?)[ \t]*\n\s*DESCRIPTION:"      # value between the 2 label halves
                           r"|COMMERCIAL GOODS\s*DESCRIPTION:[ \t]*\n?[ \t]*([^\n]+)")           # label then value

def _num(s):
    return int(str(s).replace(",", ""))

def _norm_size(s):
    return re.sub(r"\s+", "/", s.strip())

def _size_sort(sizes):
    def key(s):
        return (SIZE_RANK.index(s), 0) if s in SIZE_RANK else (len(SIZE_RANK), sizes.index(s))
    return sorted(sizes, key=key)

def _style_blocks(text):
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
        goods = (gd.group(1) or gd.group(2) or "").strip() if gd else ""

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

def auto_pack(sizes, order, pack, mode):
    rows = []
    if mode == "Solid Size":
        for sz in sizes:
            qty, p = int(order.get(sz, 0)), max(int(pack.get(sz, DEFAULT_PACK)), 1)
            full, rem = divmod(qty, p)
            if full:
                rows.append({"Kind": "Full", "Carton": "Auto", "sizes": {sz: p}, "CTN QTY": full})
            if rem:
                rows.append({"Kind": "Partial", "Carton": "Auto", "sizes": {sz: rem}, "CTN QTY": 1})
    else:
        rem_order = {s: int(order.get(s, 0)) for s in sizes}
        while sum(rem_order.values()) > 0:
            p_limit = DEFAULT_PACK
            for s in sizes:
                if rem_order[s] > 0:
                    p_limit = max(int(pack.get(s, DEFAULT_PACK)), 1)
                    break
            c_sz = {}
            cur_pcs = 0
            for s in sizes:
                if rem_order[s] <= 0: continue
                take = min(rem_order[s], p_limit - cur_pcs)
                if take > 0:
                    c_sz[s] = take
                    rem_order[s] -= take
                    cur_pcs += take
                if cur_pcs >= p_limit:
                    break
            if cur_pcs > 0:
                rows.append({"Kind": "Full" if cur_pcs == p_limit else "Partial",
                             "Carton": "Auto", "sizes": c_sz, "CTN QTY": 1})
            else:
                break
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

        gw = round(net_raw + empty, 2)
        nw = round(gw - empty, 2)
        tgw = round(gw * ctn, 2)
        tnw = round(nw * ctn, 2)

        out.append({"kind": r.get("Kind") or "Partial", "ctype": ctype, "sizes": sz, "pcs": pcs,
                    "ctn": ctn, "total_pcs": pcs * ctn,
                    "gw": gw, "nw": nw, "tgw": tgw, "tnw": tnw,
                    "empty": empty, "netwt": {s: float(netwt.get(s, DEFAULT_NETWT)) for s in sizes}})
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
        _sheet(wb.create_sheet(title=name), head, gd["group"], gd["calc"], gd["cut"],
               gd.get("netwt", {}), gd.get("nnwt", {}))
    return wb

def _sheet(ws, head, g, calc, cut, netwt=None, nnwt=None):
    netwt = netwt or {}
    nnwt  = nnwt  or {}
    sizes = g["sizes"][:MAX_SIZES]
    n_sz = max(len(sizes), 1)      # no empty size columns
    C_SZ0 = 4
    C_SZL = C_SZ0 + n_sz - 1
    C_PCS = C_SZL + 1
    C_CTN = C_PCS + 1
    C_TPC = C_CTN + 1
    C_GW  = C_TPC + 1
    C_NW  = C_GW  + 1
    C_TGW = C_NW  + 1
    C_TNW = C_TGW + 1
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

    # --- fixed rows (two blank rows between every pair of tables) ---
    GAP_H = 40
    R_NW_HDR, R_NW, R_NNW = 15, 16, 17            # SIZE / N.W. / N.N.W. table   (rows 13-14 blank)
    R_CT_HDR, R_CT_W, R_CT_D = 20, 21, 22         # carton table: R13/R07/R04, Empty carton weight, Dimension (rows 18-19 blank)
    R_H1, R_H2 = 25, 26                           # main carton table header     (rows 23-24 blank)
    R_DATA = 27

    widths = {1: 79.7, 2: 85, 3: 84.3}
    for i in range(n_sz):
        widths[C_SZ0 + i] = 38.1
    widths.update({C_PCS: 31.3, C_CTN: 43.1, C_TPC: 53.1, C_GW: 38.1,
                   C_NW: 40.4, C_TGW: 68.7, C_TNW: 48.8})
    for c, w in widths.items():
        ws.column_dimensions[L(c)].width = w

    # ---- centring helper: pick the column span whose middle is closest to the middle of the sheet ----
    cum = [0.0]
    for c in range(1, C_TNW + 1):
        cum.append(cum[-1] + widths.get(c, 8.43))
    sheet_mid = cum[-1] / 2

    def place(n_val, label_min, val_min, tot_min=None):
        """Layout [label cols][n_val value cols][optional total col], centred on the sheet.
        Returns (label_first, label_last, [value cols], total_col or None)."""
        extra = 1 if tot_min else 0
        best = None
        for c1 in range(1, C_TNW + 1):
            for k in range(1, 7):                              # label spans k columns
                v0 = c1 + k
                c2 = v0 + n_val - 1 + extra
                if c2 > C_TNW:
                    continue
                if cum[v0 - 1] - cum[c1 - 1] < label_min:
                    continue
                vcols = list(range(v0, v0 + n_val))
                if any(widths.get(c, 8.43) < val_min for c in vcols):
                    continue
                if tot_min and widths.get(c2, 8.43) < tot_min:
                    continue
                dist = abs((cum[c1 - 1] + cum[c2]) / 2 - sheet_mid)
                key = (dist, c2 - c1)
                if best is None or key < best[0]:
                    best = (key, c1, v0 - 1, vcols, c2 if tot_min else None)
        if best is None:                                       # fallback: left-aligned like the main table
            vcols = list(range(C_SZ0, C_SZ0 + n_val))
            return 1, 3, vcols, (vcols[-1] + 1 if tot_min else None)
        return best[1], best[2], best[3], best[4]

    NW_L1, NW_L2, NW_COLS, _ = place(n_sz, 100, 31)
    CT_L1, CT_L2, CT_COLS, _ = place(len(CARTONS), 110, 31)
    CT_COL = {k: CT_COLS[i] for i, k in enumerate(CARTONS.keys())}
    SM_L1, SM_L2, SM_COLS, SM_TOT = place(n_sz, 110, 31, tot_min=40)

    R_END = R_DATA + len(calc) - 1
    R_TOT = R_END + 1                  # G.TTL is the last row of the same table
    R_SUMT = R_TOT + 3                 # two blank rows after the carton table
    R_SIZE = R_SUMT + 1
    R_ORD, R_CUT, R_SHIP, R_EXS, R_PCT = R_SIZE + 1, R_SIZE + 2, R_SIZE + 3, R_SIZE + 4, R_SIZE + 5
    R_MARK = R_PCT + 2                 # two blank rows, then the CTN MEASUR / CBM block (left)
    R_LAST = R_MARK + 5

    heights = {1: 90, 2: 10.5, 3: 58.5, 4: 58.5, 5: 78.75, 6: 81, 7: 81, 8: 81, 9: 81, 10: 81,
               11: 126, 12: 40, 13: GAP_H, 14: GAP_H,
               R_NW_HDR: 70, R_NW: 70, R_NNW: 70, 18: GAP_H, 19: GAP_H,
               R_CT_HDR: 70, R_CT_W: 70, R_CT_D: 100, 23: GAP_H, 24: GAP_H,
               R_H1: 75, R_H2: 144,
               R_TOT: 110, R_TOT + 1: GAP_H, R_TOT + 2: GAP_H,
               R_SUMT: 80.25, R_SIZE: 153, R_ORD: 122.25, R_CUT: 94.5,
               R_SHIP: 94.5, R_EXS: 94.5, R_PCT: 94.5}
    for r in range(R_DATA, R_END + 1):
        heights[r] = 93.75
    for r in range(R_PCT + 1, R_LAST + 1):
        heights[r] = 78
    heights[R_PCT + 1] = heights[R_PCT + 2] = GAP_H
    for r, h in heights.items():
        ws.row_dimensions[r].height = h

    merge(1, 1, 1, 16)
    put(1, 1, "PACKING LIST DETAILS", TNR, 72, al=Alignment(horizontal="center", vertical="center"), bold=True)
    L_LBL = lambda r, t: put(r, 1, t, TNR, al=LEFT)
    L_LBL(6, "BUYER                :"); put(6, 2, "KONTOOR", TNR, al=LEFT)
    L_LBL(7, "STYLE                :"); put(7, 2, int(g["style"]) if str(g["style"]).isdigit() else g["style"], TNR, al=LEFT)
    L_LBL(8, "P.O                      :")
    put(8, 2, int(head["po"]) if head["po"].isdigit() else head["po"], TNR, al=LEFT)
    L_LBL(9, "DESCRIPTION :"); put(9, 2, g["style_desc"], TNR, al=LEFT)
    put(10, 1, "Goods Description", TNR, al=Alignment(horizontal="left", vertical="bottom"))
    put(10, 2, g["goods"], TNR, al=Alignment(horizontal="left", vertical="bottom"))
    merge(10, 2, 10, 3)                       # long goods description spans B:C
    put(11, 1, "Overship Tolerance:", ARI, al=Alignment(horizontal="left", vertical="center", wrap_text=True))
    put(11, 2, int(head["over_tol"]), ARI, al=LEFT)
    put(12, 1, "Undership Tolerance:", ARI, al=Alignment(horizontal="left", vertical="center", wrap_text=True))
    put(12, 2, int(head["under_tol"]), ARI, al=LEFT)
    put(11, 7, "CARTOON TYPE :", ARI, 72, al=Alignment(vertical="center"))
    put(11, 10, "/".join(CARTONS.keys()), ARI, 72, al=Alignment(vertical="center"))

    put(5, C_PCS, "CRD DATE", TNR, al=Alignment(vertical="center"))
    put(5, C_PCS + 2, head["crd"] or head["crd_text"], TNR, al=LEFT, fmt="m/d/yyyy")
    for r, t in ((6, "ORDER QTY"), (7, "PACK QTY"), (8, "EXCESS QTY"),
                 (9, "PERCENTAGE"), (10, "CARTON QTY")):
        put(r, C_PCS, t, ARI, al=Alignment(vertical="center"))
    put(6, C_PCS + 2, f"=+{L(SM_TOT)}{R_ORD}", TNR, al=LEFT)
    put(7, C_PCS + 2, f"=+{L(C_TPC)}{R_TOT}", TNR, al=LEFT)
    put(8, C_PCS + 2, f"={L(C_PCS + 2)}7-{L(C_PCS + 2)}6", TNR, al=LEFT)
    merge(9, C_PCS + 2, 9, C_PCS + 3)
    put(9, C_PCS + 2, f"=IFERROR({L(C_PCS + 2)}8/{L(C_PCS + 2)}6,0)", TNR, al=LEFT, fmt="0.00%")
    put(10, C_PCS + 2, f"=+{L(C_CTN)}{R_TOT}", TNR, al=LEFT)
    outline(5, C_PCS, 10, C_TNW)

    # ---- N.W. / N.N.W. reference table (yellow, centred on the sheet) ----
    WT_SZ = 60
    for r_, lbl in ((R_NW_HDR, "SIZE"), (R_NW, "N.W."), (R_NNW, "N.N.W.")):
        merge(r_, NW_L1, r_, NW_L2)
        put(r_, NW_L1, lbl, ARI, WT_SZ, bold=True, al=CEN)
    for i, s in enumerate(sizes):
        put(R_NW_HDR, NW_COLS[i], s, VER, WT_SZ, fmt="@", al=CEN)
        put(R_NW, NW_COLS[i], netwt.get(s, DEFAULT_NETWT), ARI, WT_SZ, fmt="0.000", al=CEN)
        put(R_NNW, NW_COLS[i], nnwt.get(s, DEFAULT_NNWT), ARI, WT_SZ, fmt="0.000", al=CEN)
    grid(R_NW_HDR, NW_L1, R_NNW, NW_COLS[-1])
    yw = PatternFill("solid", fgColor="FFF2CC")
    for r_ in (R_NW_HDR, R_NW, R_NNW):
        for c_ in range(NW_L1, NW_COLS[-1] + 1):
            ws.cell(row=r_, column=c_).fill = yw

    # ---- Carton table (green, centred): R13/R07/R04, Empty carton weight, Dimension ----
    #      every G.W / N.W formula below reads its carton weight from row R_CT_W
    gw_fill = PatternFill("solid", fgColor="E2EFDA")
    for r_, lbl in ((R_CT_HDR, ""), (R_CT_W, "Empty carton weight"), (R_CT_D, "Dimension")):
        merge(r_, CT_L1, r_, CT_L2)
        put(r_, CT_L1, lbl, ARI, 48, bold=True, al=CEN)
    for k, (empty, _, label) in CARTONS.items():
        c_ = CT_COL[k]
        put(R_CT_HDR, c_, k, ARI, 48, bold=True, al=CEN)
        put(R_CT_W, c_, empty, ARI, 48, fmt="0.00", al=CEN)
        put(R_CT_D, c_, label, ARI, 26, bold=True, al=CEN)
    CT_LAST = CT_COLS[-1]
    grid(R_CT_HDR, CT_L1, R_CT_D, CT_LAST)
    for r_ in (R_CT_HDR, R_CT_W, R_CT_D):
        for c_ in range(CT_L1, CT_LAST + 1):
            ws.cell(row=r_, column=c_).fill = gw_fill

    # ---- main carton table header ----
    merge(R_H1, 1, R_H2, 2); put(R_H1, 1, "Case SL#", al=CEN)
    put(R_H1, 3, "COLOR", al=CEN); put(R_H2, 3, g["cdesc"], ARI, 36, al=CEN)
    merge(R_H1, C_SZ0, R_H1, C_SZL); put(R_H1, C_SZ0, "SIZE ", al=CEN)
    for i, s in enumerate(sizes):
        put(R_H2, C_SZ0 + i, s, VER, fmt="@", al=CEN)
    for c, t in ((C_PCS, "PCS IN CTNS"), (C_CTN, "CTN \nQTY"), (C_TPC, "TTL PCS"),
                 (C_GW, "G.W"), (C_NW, "N.W")):
        merge(R_H1, c, R_H2, c); put(R_H1, c, t, al=CEN)
    put(R_H2, C_TGW, "TTL/GW", al=CEN); put(R_H2, C_TNW, "TTL/NW", al=CEN)
    grid(R_H1, 1, R_H2, C_TNW)

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

        # G.W = sum(pcs x N.W. per size) + empty carton weight; both read from the reference tables at the top
        _terms = "+".join(f"{L(C_SZ0 + i)}{xr}*{L(NW_COLS[i])}${R_NW}" for i in range(len(sizes)))
        _empty = f"${L(CT_COL[r['ctype']])}${R_CT_W}"
        put(xr, C_GW, f"=ROUND({_terms}+{_empty},2)", fmt="0.00")

        put(xr, C_NW,  f"={L(C_GW)}{xr}-{_empty}", fmt="0.00")
        put(xr, C_TGW, f"={L(C_GW)}{xr}*{L(C_CTN)}{xr}", fmt="0.00")
        put(xr, C_TNW, f"=+{L(C_CTN)}{xr}*{L(C_NW)}{xr}", fmt="0.00")
        put(xr, 3, None)
    grid(R_DATA, 1, R_END, C_TNW)

    # ---- G.TTL: last row of the carton table - size totals, bigger bold text, light background
    TOT_SZ, TOT_FILL = 60, PatternFill("solid", fgColor="DDEBF7")
    merge(R_TOT, 1, R_TOT, 3); put(R_TOT, 1, "G.TTL", TNR, TOT_SZ)
    for i in range(len(sizes)):
        cl = L(C_SZ0 + i)
        put(R_TOT, C_SZ0 + i, f"=SUMPRODUCT({cl}{R_DATA}:{cl}{R_END},${L(C_CTN)}${R_DATA}:${L(C_CTN)}${R_END})",
            ARI, TOT_SZ)
    for c in (C_CTN, C_TPC, C_GW, C_NW, C_TGW, C_TNW):
        put(R_TOT, c, f"=SUM({L(c)}{R_DATA}:{L(c)}{R_END})", ARI, TOT_SZ,
            fmt="0.00" if c >= C_GW else None)
    grid(R_TOT, 1, R_TOT, C_TNW)
    for c in range(1, C_TNW + 1):
        ws.cell(row=R_TOT, column=c).fill = TOT_FILL
    outline(R_H1, 1, R_TOT, C_TNW)

    # ---- SUMMARY: colour & size wise breakdown (centred on the sheet) ----
    SM_LAST = SM_TOT
    merge(R_SUMT, SM_L1, R_SUMT, SM_LAST); put(R_SUMT, SM_L1, "SUMMARY: COLOUR & SIZE WISE BREAKDOWN", TNR)
    put(R_SIZE, SM_L1, "SIZE", TNR); merge(R_SIZE, SM_L1, R_SIZE, SM_L2)
    for i, s in enumerate(sizes):
        put(R_SIZE, SM_COLS[i], f"=+{L(C_SZ0 + i)}{R_H2}", VER, fmt="@")
    put(R_SIZE, SM_TOT, "Total Pcs", TNR)
    for r, t, sz in ((R_ORD, "ORDER QTY", 40), (R_CUT, "CUT QTY", 40), (R_SHIP, "SHIP QTY", 40),
                     (R_EXS, "EXCESS/ SHORT", 40), (R_PCT, "%", 48)):
        put(r, SM_L1, t, TNR, sz, al=CEN)            # summary labels centred
        merge(r, SM_L1, r, SM_L2)
    first, last = L(SM_COLS[0]), L(SM_COLS[-1])
    for i, s in enumerate(sizes):
        sc = L(SM_COLS[i]); cl = L(C_SZ0 + i)           # sc = summary column, cl = same size in the main table
        put(R_ORD,  SM_COLS[i], g["order"].get(s, 0), VER)
        put(R_CUT,  SM_COLS[i], cut.get(s) or None, VER)
        put(R_SHIP, SM_COLS[i], f"=SUMPRODUCT({cl}${R_DATA}:{cl}${R_END},"
                                f"${L(C_CTN)}${R_DATA}:${L(C_CTN)}${R_END})", VER)
        put(R_EXS,  SM_COLS[i], f"=+{sc}{R_SHIP}-{sc}{R_ORD}", VER)
        put(R_PCT,  SM_COLS[i], f"=IFERROR({sc}{R_EXS}/{sc}{R_ORD},0)", TNR, fmt="0.0%")
    tp = L(SM_TOT)
    for r in (R_ORD, R_CUT, R_SHIP, R_EXS):
        put(r, SM_TOT, f"=SUM({first}{r}:{last}{r})", TNR, 72)
    put(R_PCT, SM_TOT, f"=IFERROR({tp}{R_EXS}/{tp}{R_ORD},0)", TNR, 72, fmt="0.00%")
    grid(R_SIZE, SM_L1, R_PCT, SM_LAST); outline(R_SUMT, SM_L1, R_PCT, SM_LAST)

    # ---- carton measure / CBM block: LEFT edge of the sheet (column A), text left-aligned ----
    MB1, MB2 = 1, 3
    LEFT_C = Alignment(horizontal="left", vertical="center")
    for i, (k, (_, _, label)) in enumerate(CARTONS.items()):
        merge(R_MARK + 1 + i, MB1, R_MARK + 1 + i, MB2)
        put(R_MARK + 1 + i, MB1, f"CTN MEASUR :{label}", ARI, al=LEFT_C)
    put(R_MARK + 4, MB1, "CBM:", TNR, al=LEFT_C)
    terms = []
    for k, (_, (a, b, c), _) in CARTONS.items():
        cells = [f"{L(C_CTN)}{R_DATA + i}" for i, r in enumerate(calc) if r["ctype"] == k]
        if cells:
            terms.append(f"{a}*{b}*{c}*SUM({','.join(cells)})/1000000")
    merge(R_MARK + 4, MB1 + 1, R_MARK + 4, MB2)
    put(R_MARK + 4, MB1 + 1, "=" + "+".join(terms) if terms else 0, ARI, al=LEFT_C, fmt="0.00")

    ws.sheet_view.zoomScale = SHEET_ZOOM
    ws.sheet_view.zoomScaleNormal = SHEET_ZOOM
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

    # ---- 2 packing settings
    kit.step(2, "Packing settings",
             "Not in the PO. Pieces per full carton, N.W. / N.N.W. per piece (kg) and cut qty, per size.")

    df_settings = pd.DataFrame(
        {
            "": ["Pcs / carton", "N.W. kg / pc", "N.N.W. kg / pc", "Cut qty"],
            **{
                s: [
                    float(PACK_DEFAULT.get(s, DEFAULT_PACK)),
                    float(NETWT_DEFAULT.get(s, DEFAULT_NETWT)),
                    float(NNWT_DEFAULT.get(s, DEFAULT_NNWT)),
                    0.0,
                ]
                for s in sizes
            },
        }
    )

    cfg = {
        "": st.column_config.TextColumn("", disabled=True),
        **{s: st.column_config.NumberColumn(s, format="%g") for s in sizes},
    }

    edited_settings = st.data_editor(
        df_settings,
        column_config=cfg,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key=f"{key}_settings_h"
    )

    pack  = {s: max(int(edited_settings.at[0, s] or 0), 1) for s in sizes}
    netwt = {s: float(edited_settings.at[1, s] or 0.0) for s in sizes}
    nnwt  = {s: float(edited_settings.at[2, s] or 0.0) for s in sizes}
    cut   = {s: int(edited_settings.at[3, s] or 0) for s in sizes}

    # ============================================================
    # STEP 3 - Carton table (improved live calculation)
    # ============================================================
    kit.step(3, "Carton table", "Edit any cell and press Enter. Calculated columns update immediately.")

    mode = st.radio(
        "Packing Mode",
        [
            "Solid Size", "Solid Color", "Assort Size", "Assort Color",
            "Assort Size + Solid Color", "Assort Color + Solid Size", "All Mixed"
        ],
        horizontal=True,
        key=f"{key}_mode"
    )

    sig = "_".join(str(pack[s]) for s in sizes) + "_" + mode
    state_key = f"{key}_carton_{sig}"

    # Initialize session state with editable data only
    if state_key not in st.session_state:
        base = auto_pack(sizes, order, pack, mode)
        st.session_state[state_key] = [
            {
                "Kind": r["Kind"],
                "Carton": r["Carton"],
                **{s: r["sizes"].get(s, 0) for s in sizes},
                "CTN QTY": r["CTN QTY"],
            }
            for r in base
        ]

    def build_display(rows):
        display = []
        cur = 1
        for r in rows:
            sz = {s: int(r.get(s, 0) or 0) for s in sizes}
            pcs = sum(sz.values())
            ctn = int(r.get("CTN QTY") or 0)
            ctype = r.get("Carton") if r.get("Carton") in CARTONS else carton_type(pcs)
            empty = CARTONS[ctype][0]
            net_raw = sum(sz[s] * netwt.get(s, DEFAULT_NETWT) for s in sizes)
            gw = round(net_raw + empty, 2)
            nw = round(gw - empty, 2)

            display.append({
                "CTN Range": f"{cur}" if ctn == 1 else f"{cur} - {cur + ctn - 1}",
                "Kind": r.get("Kind") or "Partial",
                "Carton": r.get("Carton") or "Auto",
                **sz,
                "PCS IN CTNS": pcs,
                "CTN QTY": ctn,
                "Total CTN": pcs * ctn,
                "Gross weight": gw,
                "Net weight": nw,
                "Total gross weight": round(gw * ctn, 2),
                "Total net weight": round(nw * ctn, 2),
            })
            cur += ctn
        return display

    disp_rows = build_display(st.session_state[state_key])
    ttl = {k: None for k in (disp_rows[0].keys() if disp_rows else [])}
    if disp_rows:
        ttl["CTN Range"] = "G.TTL"
        for s_ in sizes:
            ttl[s_] = sum(int(r[s_]) * int(r["CTN QTY"]) for r in disp_rows)
        ttl["CTN QTY"] = sum(int(r["CTN QTY"]) for r in disp_rows)
        ttl["Total CTN"] = sum(int(r["Total CTN"]) for r in disp_rows)
        ttl["Total gross weight"] = round(sum(r["Total gross weight"] for r in disp_rows), 2)
        ttl["Total net weight"] = round(sum(r["Total net weight"] for r in disp_rows), 2)
        disp_rows = disp_rows + [ttl]
    df_display = pd.DataFrame(disp_rows)

    cfg = {
        "CTN Range": st.column_config.TextColumn("CTN Range", disabled=True, width="small"),
        "Kind": st.column_config.SelectboxColumn("Kind", options=["Full", "Partial"], required=True),
        "Carton": st.column_config.SelectboxColumn("Carton", options=["Auto"] + list(CARTONS.keys()), required=True),
        "PCS IN CTNS": st.column_config.NumberColumn("PCS IN CTNS", disabled=True),
        "CTN QTY": st.column_config.NumberColumn("CTN QTY", min_value=0, step=1),
        "Total CTN": st.column_config.NumberColumn("Total CTN", disabled=True),
        "Gross weight": st.column_config.NumberColumn("Gross weight", format="%.2f", disabled=True),
        "Net weight": st.column_config.NumberColumn("Net weight", format="%.2f", disabled=True),
        "Total gross weight": st.column_config.NumberColumn("Total gross weight", format="%.2f", disabled=True),
        "Total net weight": st.column_config.NumberColumn("Total net weight", format="%.2f", disabled=True),
    }
    for s in sizes:
        cfg[s] = st.column_config.NumberColumn(s, min_value=0, step=1)

    edited = st.data_editor(
        df_display,
        column_config=cfg,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key=f"{state_key}_editor",
    )

    # Save editable columns back
    new_rows = []
    for _, row in edited.iterrows():
        if row.get("CTN Range") == "G.TTL" or pd.isna(row.get("CTN QTY")):
            continue
        new_rows.append({
            "Kind": row.get("Kind") or "Partial",
            "Carton": row.get("Carton") or "Auto",
            **{s: int(row.get(s) or 0) for s in sizes},
            "CTN QTY": int(row["CTN QTY"]),
        })

    # Force immediate refresh when data actually changed
    if new_rows != st.session_state[state_key]:
        st.session_state[state_key] = new_rows
        st.rerun()

    # Convert to format expected by calc_rows
    rows_for_calc = []
    for r in new_rows:
        rows_for_calc.append({
            "Kind": r["Kind"],
            "Carton": r["Carton"],
            "sizes": {s: r.get(s, 0) for s in sizes},
            "CTN QTY": r["CTN QTY"],
        })

    calc = calc_rows(rows_for_calc, sizes, netwt)
    ship = ship_by_size(calc, sizes)

    # ---- Totals + Summary (directly below the carton table) ----
    tot_ctn = sum(r["ctn"] for r in calc)
    tot_pcs = sum(ship.values())
    tgw = sum(r["tgw"] for r in calc)
    tnw = sum(r["tnw"] for r in calc)

    # SUMMARY: COLOUR & SIZE WISE BREAKDOWN (same as the Excel sheet)
    st.markdown("**SUMMARY: COLOUR & SIZE WISE BREAKDOWN**")
    o_ = {s: int(order.get(s, 0)) for s in sizes}
    c_ = {s: int(cut.get(s, 0) or 0) for s in sizes}
    sh_ = {s: int(ship.get(s, 0)) for s in sizes}
    ex_ = {s: sh_[s] - o_[s] for s in sizes}
    pc_ = {s: (f"{ex_[s] / o_[s] * 100:.1f}%" if o_[s] else "0.0%") for s in sizes}
    tO, tC, tS = sum(o_.values()), sum(c_.values()), sum(sh_.values())
    tE = tS - tO
    summary_df = pd.DataFrame({
        "SIZE": ["ORDER QTY", "CUT QTY", "SHIP QTY", "EXCESS/ SHORT", "%"],
        **{s: [o_[s], c_[s], sh_[s], ex_[s], pc_[s]] for s in sizes},
        "Total Pcs": [tO, tC, tS, tE, f"{tE / tO * 100:.2f}%" if tO else "0.00%"],
    })
    st.dataframe(summary_df, hide_index=True, use_container_width=True)

    st.markdown("---")
    st.markdown("### Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Cartons", f"{tot_ctn:,}")
    c2.metric("Total PCS", f"{tot_pcs:,}")
    c3.metric("Total Gross Weight", f"{tgw:,.2f} kg")
    c4.metric("Total Net Weight", f"{tnw:,.2f} kg")
    st.markdown("---")

    # ---- 4 check
    kit.step(4, "Check against the PO")
    gw = sum(r["gw"] * r["ctn"] for r in calc)
    nw = sum((r["gw"] - r["empty"]) * r["ctn"] for r in calc)
    kit.summary_cards([("Cartons", f"{tot_ctn:,}"), ("Ship pcs", f"{tot_pcs:,}"),
                       ("Order pcs", f"{sum(order.get(s, 0) for s in sizes):,}"),
                       ("Gross kg", f"{gw:,.1f}"), ("Net kg", f"{nw:,.1f}")])
    kit.check_by_size(sizes, {s: order.get(s, 0) for s in sizes}, ship,
                      over_pct=float(head["over_tol"]), under_pct=float(head["under_tol"]))

    g2 = dict(g, sizes=sizes)
    return {"group": g2, "calc": calc, "cut": cut, "netwt": netwt, "nnwt": nnwt}

def run(pdf_file=None):
    st.title("Kontoor / Lee - PO to Packing List")

    # If called from parent app → use the file that was already uploaded
    # If running standalone → show the uploader here
    if pdf_file is None:
        pdf_file = st.file_uploader(
            "Upload Kontoor Brands Purchase Order (PDF)",
            type=["pdf"],
            key="kontoor_uploader"
        )

    if pdf_file is None:
        st.info("Please upload a Kontoor Brands purchase order PDF to continue.")
        st.stop()

    try:
        head, groups = _parse_cached(pdf_file.getvalue())
    except Exception as e:
        st.error(f"Could not read the PDF: {e}")
        st.stop()

    if not groups:
        st.warning("No style / size lines were found - the PDF layout may differ.")
        st.stop()

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
    run()