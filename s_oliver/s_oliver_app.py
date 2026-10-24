


"""
s.Oliver Overseas Ltd. (Bangladesh) PO -> Packing List
=======================================================
Module for the unified app.  The router calls  run(pdf_file).
Also runs alone:  streamlit run s_oliver_app.py

Reads the s.Oliver "Purchase Order" PDF (article 2184662, PO 4810116911 ...) and
produces the "Final Packing List" workbook the buyer expects.
"""
import io
import re
from collections import OrderedDict

import sys
from pathlib import Path

import pdfplumber
import pandas as pd
import streamlit as st

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:                      # pl_kit.py lives in the project root
    sys.path.insert(0, _ROOT)
import pl_kit as kit                           # noqa: E402
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter as L

# ============================================================
# DEFAULTS
# ============================================================
WIDTHS = ["32", "34", "36", "38", "40", "42", "44", "46", "48"]
LENGTHS = ["28", "30", "32", "34"]
PACK_DEFAULT = 20
EMPTY_CARTON = 1.20
NETWT_DEFAULT = {
    "32": 0.36, "34": 0.36, "36": 0.40, "38": 0.40, "40": 0.42,
    "42": 0.44, "44": 0.44, "46": 0.46, "48": 0.46,
}
CARTON_DIM = "60 X 40 X 20"
CARTON_CBM = 0.6 * 0.4 * 0.2
CONSIGNEE = [
    "S.OLIVER Bernd Freier GmbH & Co.KG",
    "S.Oliver - Straße 3",
    "D - 97228 Rottendorf",
    "GERMANY",
]

# ============================================================
# PDF PARSING
# ============================================================
ITEM_RE = re.compile(
    r"(\d{10})"
    r"\s*"
    r"([0-9]{2}[A-Z][0-9])"
    r"\s*"
    r"(\d{2})"
    r"\s*"
    r"(\d{2})"
    r"\s*"
    r"(\d{1,4}?)"
    r"\s*"
    r"(\d{2}\.\d{2})"
)


def _grab(text, pat, default=""):
    m = re.search(pat, text)
    return m.group(1).strip() if m else default


def parse_po(pdf_bytes: bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    head = {
        "po":          _grab(text, r"PO-\s*No\.?:\s*(\d+)") or _grab(text, r"\bPO:\s*(\d+)"),
        "style":       _grab(text, r"Style-\s*ID:\s*(\d+)"),
        "article":     _grab(text, r"Article Description:\s*(.+?)\s+Gender:"),
        "division":    _grab(text, r"Division:\s*(.+?)\s+Line:"),
        "line":        _grab(text, r"Line:\s*(\d+)"),
        "brand":       _grab(text, r"Brand:\s*(\d+)"),
        "pg":          _grab(text, r"Purchasing Group:\s*(\w+)"),
        "handover":    _grab(text, r"Handover:\s*(\S+)"),
        "ship_mode":   _grab(text, r"Shipment Mode:\s*(\w+)"),
        "terms":       _grab(text, r"Terms of delivery:\s*(.+?)(?:\s+Handover|\n)"),
        "sap_id":      _grab(text, r"(10\.\d+\.\d+\.\d+\.\d+\.\d+)"),
    }

    items, seen = [], set()
    for m in ITEM_RE.finditer(text):
        article, color, w, l, q, p = m.groups()
        key = (article, color, w, l)
        if key in seen:
            continue
        seen.add(key)
        items.append({"article": article, "color": color, "width": w,
                      "length": l, "qty": int(q), "price": float(p)})
    return head, items


def group_items(items):
    g = OrderedDict()
    for it in items:
        g.setdefault(it["color"], OrderedDict()).setdefault(it["length"], {})[it["width"]] = it["qty"]
    return g


# ============================================================
# CARTON LOGIC
# ============================================================
def build_carton_rows(groups, pack, netwt):
    rows, cno = [], 1
    for color, lengths in groups.items():
        for length, widths in lengths.items():
            for w in WIDTHS:
                qty = widths.get(w, 0)
                if qty <= 0:
                    continue
                full, rem = divmod(qty, pack)
                if full:
                    rows.append({"from": cno, "to": cno + full - 1, "num": full,
                                 "color": color, "length": length, "width": w, "pcs": pack,
                                 "per_gw": round(pack * netwt[w] + EMPTY_CARTON, 2)})
                    cno += full
                if rem:
                    rows.append({"from": cno, "to": cno, "num": 1,
                                 "color": color, "length": length, "width": w, "pcs": rem,
                                 "per_gw": round(rem * netwt[w] + EMPTY_CARTON, 2)})
                    cno += 1
    return rows


def rows_from_edit(recs, groups, netwt):
    """Edited carton table -> carton rows (numbered, weights). recs keys: color, length, width, pcs, num."""
    rows, problems, cno = [], [], 1
    for n, r in enumerate(recs, 1):
        color, length, w = r.get("color"), r.get("length"), r.get("width")
        try:
            pcs, num = int(float(r.get("pcs"))), int(float(r.get("num")))
        except (TypeError, ValueError):
            if color or length or w:
                problems.append(f"Row {n}: pieces and cartons must be numbers - row ignored.")
            continue
        if color not in groups or length not in LENGTHS or w not in WIDTHS:
            problems.append(f"Row {n}: pick colour, inseam and waist from the lists - row ignored.")
            continue
        if pcs <= 0 or num <= 0:
            problems.append(f"Row {n}: pieces and cartons must be above 0 - row ignored.")
            continue
        rows.append({"from": cno, "to": cno + num - 1, "num": num, "color": color, "length": length,
                     "width": w, "pcs": pcs, "per_gw": round(pcs * netwt[w] + EMPTY_CARTON, 2)})
        cno += num
    return rows, problems


# ============================================================
# EXCEL
# ============================================================
def build_workbook(head, groups, netwt, pack, rows=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Final Packing List"

    TNR, ARI, VER = "Times New Roman", "Arial", "Verdana"
    thin = Side(style="thin")
    CEN  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left",   vertical="center")

    def put(r, c, v=None, name=ARI, sz=10, al=CEN, fmt=None, bold=True, u=None):
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

    def box(r1, c1, r2, c2, side=thin):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).border = Border(left=side, right=side, top=side, bottom=side)

    # ---------------- header block ----------------
    put(1, 1, "Supplier Name and Address", TNR, 11, al=LEFT)
    put(2, 1, "CREATIVE COLLECTIONS LTD", TNR, 11, al=LEFT)
    put(3, 1, "Nishat Nagar , Tongi , Gazipur .", TNR, 11, al=LEFT)
    put(4, 1, "Bangladesh", TNR, 11, al=LEFT)

    put(6, 1, "Consignee:", TNR, 11, al=LEFT)
    for i, ln in enumerate(CONSIGNEE):
        put(7 + i, 1, ln, TNR, 11, al=LEFT)

    put(12, 1, "PACKING LIST", TNR, 14, al=LEFT, u="single")

    put(14, 1, "ORDER NO:",              TNR, 11, al=LEFT); put(14, 4, head["po"],    TNR, 11, al=LEFT)
    put(15, 1, "STYLE NO:",              TNR, 11, al=LEFT); put(15, 4, head["style"], TNR, 11, al=LEFT)
    put(16, 1, "SAP ID NO:",             TNR, 11, al=LEFT)
    put(16, 4, head["sap_id"] or f"10.2.11.26.185.{head['style']}", TNR, 11, al=LEFT)
    total_order = sum(sum(w.values()) for L2 in groups.values() for w in L2.values())
    put(17, 1, "TOTAL ORDER QTY",        TNR, 11, al=LEFT); put(17, 4, total_order, TNR, 11, al=LEFT)
    put(18, 1, "PACKING TYPE (FLAT or GOH)", TNR, 11, al=LEFT); put(18, 4, "Flat pack", TNR, 11, al=LEFT)
    put(19, 1, "FROM BANGLADESH TO:",    TNR, 11, al=LEFT); put(19, 4, "GERMANY",   TNR, 11, al=LEFT)
    put(20, 1, "DESCRIPTION",            TNR, 11, al=LEFT); put(20, 4, head["article"], TNR, 11, al=LEFT)

    put(14, 8, "SHIPPING MARK:",          TNR, 11, al=LEFT); put(14, 11, "S.OLIVER", TNR, 11, al=LEFT)
    put(15, 8, "Bernd Freier GmbH & Co.KG", TNR, 11, al=LEFT)
    put(16, 8, "Haus 7",                  TNR, 11, al=LEFT)
    put(17, 8, "S.Oliver - Straße 3",     TNR, 11, al=LEFT)
    put(18, 8, "D - 97228 Rottendorf",    TNR, 11, al=LEFT)
    put(19, 8, "Germany",                 TNR, 11, al=LEFT)

    # ---------------- SUMMARY ----------------
    R_TITLE = 22
    R_H1, R_H2 = 24, 25
    R_D0 = 26
    put(R_TITLE, 1, "SUMMARY OF TOTAL BREAKDOWN", TNR, 12, al=LEFT)

    C0 = 6
    put(R_H1, 1, "Color", TNR); put(R_H1, 4, "Size", TNR)
    put(R_H2, 1, "Inseam", TNR)
    for i, w in enumerate(WIDTHS):
        c = C0 + i * 4
        merge(R_H1, c, R_H1, c + 3)
        put(R_H1, c, w, TNR)
        put(R_H2, c,     "ship",   TNR, 9)
        put(R_H2, c + 1, "order",  TNR, 9)
        put(R_H2, c + 2, "Dif.",   TNR, 9)
        put(R_H2, c + 3, "Diff.%", TNR, 9)
    C_TOT = C0 + len(WIDTHS) * 4
    merge(R_H1, C_TOT, R_H1, C_TOT + 3)
    put(R_H1, C_TOT,     "TOTAL",  TNR)
    put(R_H2, C_TOT,     "ship",   TNR, 9)
    put(R_H2, C_TOT + 1, "order",  TNR, 9)
    put(R_H2, C_TOT + 2, "Dif.",   TNR, 9)
    put(R_H2, C_TOT + 3, "Diff.%", TNR, 9)
    box(R_H1, 1, R_H2, C_TOT + 3)

    r = R_D0
    color_blocks = []
    ship_cells = {}
    for color, lengths in groups.items():
        start = r
        for length in LENGTHS:
            widths = lengths.get(length, {})
            if not any(widths.values()):
                continue
            put(r, 1, color,  TNR, 10)
            put(r, 4, length, TNR, 10)
            for i, w in enumerate(WIDTHS):
                c = C0 + i * 4
                q = widths.get(w, 0)
                put(r, c,     q, VER, 10)            # ship: filled with a live formula once the carton table exists
                ship_cells[(color, length, w)] = (r, c)
                put(r, c + 1, q, VER, 10)
                put(r, c + 2, f"={L(c)}{r}-{L(c+1)}{r}",       VER, 10)
                put(r, c + 3, f"=IFERROR({L(c+2)}{r}/{L(c+1)}{r},0)", TNR, 10, fmt="0.0%")
            put(r, C_TOT,     f"=SUM({L(C0)}{r}:{L(C_TOT-1)}{r})", TNR, 10)
            put(r, C_TOT + 1, f"=SUM({L(C0+1)}{r}:{L(C_TOT)}{r})", TNR, 10)
            put(r, C_TOT + 2, f"={L(C_TOT)}{r}-{L(C_TOT+1)}{r}",   TNR, 10)
            put(r, C_TOT + 3, f"=IFERROR({L(C_TOT+2)}{r}/{L(C_TOT+1)}{r},0)", TNR, 10, fmt="0.0%")
            r += 1
        # subtotal row
        put(r, 1, "TOTAL", TNR, 10)
        for i in range(len(WIDTHS) * 4 + 4):
            c = C0 + i
            put(r, c, f"=SUM({L(c)}{start}:{L(c)}{r-1})", TNR, 10)
        put(r, C_TOT + 2, f"={L(C_TOT)}{r}-{L(C_TOT+1)}{r}", TNR, 10)
        put(r, C_TOT + 3, f"=IFERROR({L(C_TOT+2)}{r}/{L(C_TOT+1)}{r},0)", TNR, 10, fmt="0.0%")
        box(r, 1, r, C_TOT + 3)
        color_blocks.append((start, r))
        r += 1

    # G.TOTAL row — sum of the per-colour subtotal rows
    R_GT = r
    put(R_GT, 1, "G. TOTAL", TNR, 10)
    for i in range(len(WIDTHS) * 4 + 4):
        c = C0 + i
        formula = "=" + "+".join(f"{L(c)}{row}" for _, row in color_blocks)
        put(R_GT, c, formula, TNR, 10)
    box(R_GT, 1, R_GT, C_TOT + 3)

    # ---------------- carton detail ----------------
    rows = rows if rows is not None else build_carton_rows(groups, pack, netwt)
    R_DET0 = R_GT + 2

    put(R_DET0, 1, "Carton No.", TNR, 10)
    put(R_DET0, 4, "PREP",       TNR, 10)
    put(R_DET0, 5, "COLOR",      TNR, 10)
    put(R_DET0, 7, "SIZE",       TNR, 10)
    for i, w in enumerate(WIDTHS):
        put(R_DET0, 8 + i, w, TNR, 9)
    C_PCS = 8 + len(WIDTHS)
    for j, t in enumerate(("Total", "Par/ctn Gr.Wt", "GR.WT", "NT.WT", "MEASUREMENT")):
        put(R_DET0, C_PCS + j, t, TNR, 10)
    box(R_DET0, 1, R_DET0, C_PCS + 4)

    cur = R_DET0 + 1
    prev_color_len = None
    det_rng = {}
    for row in rows:
        cl = f"{row['color']}-{row['length']}"
        put(cur, 1, row["from"] if cur == R_DET0 + 1 else f"=C{cur - 1}+1", TNR, 10, fmt="0")
        put(cur, 2, "-",          TNR, 10)
        put(cur, 3, f"=A{cur}+D{cur}-1", TNR, 10, fmt="0")
        put(cur, 4, row["num"],   TNR, 10)
        put(cur, 5, cl if cl != prev_color_len else None, TNR, 10)
        put(cur, 8 + WIDTHS.index(row["width"]), row["pcs"], VER, 10)
        put(cur, C_PCS,     f"=SUM(H{cur}:{L(8+len(WIDTHS)-1)}{cur})", TNR, 10)
        # LIVE: gross per carton = SUM(pcs in each waist-width column x net kg/pc of that width) + empty carton
        _terms = "+".join(f"{L(8 + i)}{cur}*{float(netwt[w]):.5f}" for i, w in enumerate(WIDTHS) if w in netwt)
        put(cur, C_PCS + 1, f"=ROUND({_terms}+{EMPTY_CARTON},2)" if _terms else row["per_gw"], TNR, 10, fmt="0.00")
        put(cur, C_PCS + 2, f"={L(C_PCS+1)}{cur}*D{cur}",                 TNR, 10, fmt="0.00")
        put(cur, C_PCS + 3, f"={L(C_PCS+2)}{cur}-{EMPTY_CARTON}*D{cur}",  TNR, 10, fmt="0.00")
        put(cur, C_PCS + 4, CARTON_DIM, TNR, 10)
        prev_color_len = cl
        _rng = det_rng.setdefault((row["color"], row["length"]), [cur, cur])
        _rng[1] = cur
        cur += 1

    R_DET_END = cur - 1

    # LIVE summary "ship" = pcs per carton x cartons, summed over that colour / inseam block of the carton table
    for (color_, length_, w_), (rr_, cc_) in ship_cells.items():
        rg = det_rng.get((color_, length_))
        if rg:
            wc = L(8 + WIDTHS.index(w_))
            put(rr_, cc_, f"=SUMPRODUCT({wc}{rg[0]}:{wc}{rg[1]},$D${rg[0]}:$D${rg[1]})", VER, 10)
    put(cur, 1, "G.TOTAL", TNR, 10)
    for i in range(len(WIDTHS)):
        c = 8 + i
        put(cur, c, f"=SUM({L(c)}{R_DET0+1}:{L(c)}{R_DET_END})", TNR, 10)
    put(cur, C_PCS,     f"=SUM({L(C_PCS)}{R_DET0+1}:{L(C_PCS)}{R_DET_END})",         TNR, 10)
    put(cur, C_PCS + 1, f"=SUM({L(C_PCS+1)}{R_DET0+1}:{L(C_PCS+1)}{R_DET_END})",     TNR, 10, fmt="0.00")
    put(cur, C_PCS + 2, f"=SUM({L(C_PCS+2)}{R_DET0+1}:{L(C_PCS+2)}{R_DET_END})",     TNR, 10, fmt="0.00")
    put(cur, C_PCS + 3, f"=SUM({L(C_PCS+3)}{R_DET0+1}:{L(C_PCS+3)}{R_DET_END})",     TNR, 10, fmt="0.00")
    box(cur, 1, cur, C_PCS + 4)

    # ---------------- footer ----------------
    R_FOOT = cur + 2
    total_ctn = sum(row["num"] for row in rows)
    put(R_FOOT,     1, "TOTAL GROSS WEIGHT", TNR, 11, al=LEFT)
    put(R_FOOT,     7, f"={L(C_PCS+2)}{cur}", TNR, 11); put(R_FOOT,     8, "Kgs",  TNR, 11)
    put(R_FOOT + 1, 1, "TOTAL NET WEIGHT",   TNR, 11, al=LEFT)
    put(R_FOOT + 1, 7, f"={L(C_PCS+3)}{cur}", TNR, 11); put(R_FOOT + 1, 8, "Kgs",  TNR, 11)
    put(R_FOOT + 2, 1, "TOTAL CTN",          TNR, 11, al=LEFT)
    put(R_FOOT + 2, 7, f"=SUM(D{R_DET0 + 1}:D{R_DET_END})", TNR, 11); put(R_FOOT + 2, 8, "Ctns", TNR, 11)
    put(R_FOOT + 3, 1, "TOTAL CBM",          TNR, 11, al=LEFT)
    put(R_FOOT + 3, 7, f"=G{R_FOOT + 2}*{CARTON_CBM}", TNR, 11, fmt="0.000"); put(R_FOOT + 3, 8, "CBM", TNR, 11)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 7
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 3
    ws.column_dimensions["G"].width = 8
    for i in range(len(WIDTHS)):
        ws.column_dimensions[L(8 + i)].width = 6
    for j in range(5):
        ws.column_dimensions[L(C_PCS + j)].width = 12

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.25

    return wb


# ============================================================
# STREAMLIT RENDER HELPERS
# ============================================================
def _render_order_table(groups):
    """Order qty table — one row per (colour, inseam), sizes as columns."""
    st.markdown("**Order quantity**  (summary of total breakdown)")

    rows = []
    for color, lengths in groups.items():
        for length in LENGTHS:
            widths = lengths.get(length, {})
            if not any(widths.values()):
                continue
            r = {"Color": color, "Inseam": length}
            for w in WIDTHS:
                r[w] = widths.get(w, 0)
            r["Total"] = sum(r[w] for w in WIDTHS)
            rows.append(r)
        # subtotal per colour
        subtotal = {"Color": f"TOTAL {color}", "Inseam": ""}
        for w in WIDTHS:
            subtotal[w] = sum(r[w] for r in rows
                              if r["Color"] == color)
        subtotal["Total"] = sum(subtotal[w] for w in WIDTHS)
        rows.append(subtotal)

    gt = {"Color": "G. TOTAL", "Inseam": ""}
    for w in WIDTHS:
        gt[w] = sum(r[w] for r in rows
                    if not str(r["Color"]).startswith("TOTAL")
                    and r["Color"] != "G. TOTAL")
    gt["Total"] = sum(gt[w] for w in WIDTHS)
    rows.append(gt)

    df = pd.DataFrame(rows)
    left_cols = ["Color", "Inseam"]
    styler = (
        df.style
        .set_properties(subset=left_cols, **{"text-align": "left"})
        .set_properties(
            subset=[c for c in df.columns if c not in left_cols],
            **{"text-align": "center"},
        )
        .apply(lambda r: [
            "font-weight: bold; background-color: #F0F0F0"
            if str(r["Color"]).startswith("TOTAL") or r["Color"] == "G. TOTAL"
            else "" for _ in r
        ], axis=1)
    )
    st.dataframe(styler, use_container_width=True, hide_index=True)


def _render_carton_table(rows, head):
    """Carton detail table — matches the Excel layout."""
    st.markdown("**Carton details**  (calculated from your table)")
    if not rows:
        st.info("No carton rows to display.")
        return

    table_rows = []
    prev_cl = None
    for row in rows:
        cl = f"{row['color']}-{row['length']}"
        r = {
            "Carton No. From": row["from"],
            "To":              row["to"],
            "PREP":            row["num"],
            "COLOR-LEN":       cl if cl != prev_cl else "",
        }
        for w in WIDTHS:
            r[w] = row["pcs"] if w == row["width"] else ""
        r["Total"]         = row["pcs"]
        r["Par/ctn Gr.Wt"] = row["per_gw"]
        r["GR.WT"]         = round(row["per_gw"] * row["num"], 2)
        r["NT.WT"]         = round(row["per_gw"] * row["num"] - EMPTY_CARTON * row["num"], 2)
        r["MEASUREMENT"]   = CARTON_DIM
        table_rows.append(r)
        prev_cl = cl

    gt = {
        "Carton No. From": "G.TOTAL",
        "To":              "",
        "PREP":            sum(r["PREP"] for r in table_rows),
        "COLOR-LEN":       "",
    }
    for w in WIDTHS:
        gt[w] = sum(int(r[w]) for r in table_rows if r[w] != "")
    gt["Total"]         = sum(r["Total"] for r in table_rows)
    gt["Par/ctn Gr.Wt"] = round(sum(r["Par/ctn Gr.Wt"] for r in table_rows), 2)
    gt["GR.WT"]         = round(sum(r["GR.WT"] for r in table_rows), 2)
    gt["NT.WT"]         = round(sum(r["NT.WT"] for r in table_rows), 2)
    gt["MEASUREMENT"]   = ""
    table_rows.append(gt)

    df = pd.DataFrame(table_rows)
    left_cols = ["Carton No. From", "To", "COLOR-LEN", "MEASUREMENT"]
    left_cols = [c for c in left_cols if c in df.columns]
    styler = (
        df.style
        .set_properties(subset=left_cols, **{"text-align": "left"})
        .set_properties(
            subset=[c for c in df.columns if c not in left_cols],
            **{"text-align": "center"},
        )
        .apply(lambda r: [
            "font-weight: bold; background-color: #F0F0F0"
            if r["Carton No. From"] == "G.TOTAL" else ""
            for _ in r
        ], axis=1)
    )
    st.dataframe(styler, use_container_width=True, hide_index=True)


def _render_summary_table(rows):
    """Summary table below the carton detail."""
    st.markdown("**Summary**")
    total_ctn = sum(r["num"] for r in rows)
    total_pcs = sum(r["pcs"] * r["num"] for r in rows)
    total_gw  = sum(r["per_gw"] * r["num"] for r in rows)
    total_nw  = sum((r["per_gw"] * r["num"]) - EMPTY_CARTON * r["num"] for r in rows)
    total_cbm = CARTON_CBM * total_ctn

    summary = pd.DataFrame([
        {"Description": "TOTAL CARTONS",      "Value": f"{total_ctn:,}"},
        {"Description": "TOTAL PIECES",       "Value": f"{total_pcs:,}"},
        {"Description": "TOTAL GROSS WEIGHT", "Value": f"{total_gw:,.2f} kg"},
        {"Description": "TOTAL NET WEIGHT",   "Value": f"{total_nw:,.2f} kg"},
        {"Description": "TOTAL CBM",          "Value": f"{total_cbm:,.3f} m³"},
        {"Description": "CARTON DIMENSION",   "Value": f"{CARTON_DIM} cm"},
    ])

    styler = (
        summary.style
        .set_properties(subset=["Description"],
                        **{"text-align": "left", "font-weight": "bold"})
        .set_properties(subset=["Value"], **{"text-align": "center"})
    )
    st.dataframe(styler, use_container_width=True, hide_index=True)


# ============================================================
# STREAMLIT UI
# ============================================================
@st.cache_data(show_spinner=False, max_entries=10)
def _parse_cached(b: bytes):
    return parse_po(b)


def run(pdf_file=None):
    st.title("s.Oliver - PO to Packing List")
    if pdf_file is None:
        st.info("Upload an s.Oliver purchase order PDF.")
        return
    try:
        head, items = _parse_cached(pdf_file.getvalue())
    except Exception as e:
        st.error(f"Could not read the PDF: {e}")
        return
    if not items:
        st.warning("No article / size lines were found - the PDF layout may differ.")
        return

    groups = group_items(items)
    total_order = sum(it["qty"] for it in items)
    key = f"so_{head['po']}"

    # ---- 1 summary
    kit.step(1, "PO summary")
    kit.summary_cards([("PO", head["po"]), ("Style", head["style"]), ("Colours", f"{len(groups)}"),
                       ("Handover", head.get("handover") or "-"), ("Pieces", f"{total_order:,}")])
    st.caption(f"Article: {head['article']}")
    _render_order_table(groups)

    # ---- 2 variables
    kit.step(2, "Packing settings", "Not in the PO - edit if different.")
    c1, c2 = st.columns([1, 3])
    pack = int(c1.number_input("Pieces per full carton", 1, 200, PACK_DEFAULT, 1, key=f"{key}_pack"))
    c2.caption(f"Empty carton: {EMPTY_CARTON} kg · Carton {CARTON_DIM} cm")
    st.markdown("**Net weight per piece (kg) by waist size**")
    cols = st.columns(len(WIDTHS))
    netwt = {}
    for col, w in zip(cols, WIDTHS):
        with col:
            netwt[w] = st.number_input(w, 0.10, 2.00, NETWT_DEFAULT.get(w, 0.40), 0.01,
                                       format="%.2f", key=f"so_nw_{w}")

    # ---- 3 carton table
    kit.step(3, "Carton table",
             "Edit any cell, add a row with +, or tick 🗑 Delete and press “Remove ticked rows”. "
             "Changing “Pieces per full carton” above rebuilds this table.")
    auto = build_carton_rows(groups, pack, netwt)
    base = pd.DataFrame([{"Color": r["color"], "Inseam": r["length"], "Waist": r["width"],
                          "Pcs/Ctn": r["pcs"], "Cartons": r["num"]} for r in auto])
    cfg = {"Color": st.column_config.SelectboxColumn("Color", options=list(groups.keys()), required=True),
           "Inseam": st.column_config.SelectboxColumn("Inseam", options=LENGTHS, required=True),
           "Waist": st.column_config.SelectboxColumn("Waist", options=WIDTHS, required=True),
           "Pcs/Ctn": st.column_config.NumberColumn("Pcs/Ctn", min_value=1, step=1, format="%d"),
           "Cartons": st.column_config.NumberColumn("Cartons", min_value=1, step=1, format="%d")}
    ed = kit.editable_grid(base, f"{key}_ct_{pack}", column_config=cfg)
    recs = [{"color": r["Color"], "length": r["Inseam"], "width": r["Waist"], "pcs": r["Pcs/Ctn"], "num": r["Cartons"]}
            for r in ed.to_dict("records")]
    rows, probs = rows_from_edit(recs, groups, netwt)
    kit.problems(probs)
    if not rows:
        st.error("The carton table is empty. Press “Reset table to PO” or add a row.")
        return
    _render_carton_table(rows, head)

    # ---- 4 check
    kit.step(4, "Check against the PO")
    order, packed = {}, {}
    for color, lengths in groups.items():
        for length, widths in lengths.items():
            for w, q in widths.items():
                lab = f"{color}-{length} / {w}"
                order[lab] = q
                packed[lab] = sum(r["pcs"] * r["num"] for r in rows
                                  if (r["color"], r["length"], r["width"]) == (color, length, w))
    kit.check_by_size(list(order), order, packed)
    _render_summary_table(rows)

    # ---- 5 download
    kit.step(5, "Download")
    wb = build_workbook(head, groups, netwt, int(pack), rows=rows)
    buf = io.BytesIO()
    wb.save(buf)
    kit.download_block("⬇ Download packing list (Excel)", buf.getvalue(), f"packing_list_{head['po']}.xlsx",
                       "Empty-carton weight, carton size and net weight per piece use the settings above.")


if __name__ == "__main__":
    st.set_page_config(page_title="s.Oliver PO -> Packing List", layout="wide")
    run(st.file_uploader("Upload s.Oliver PO PDF", type=["pdf"]))