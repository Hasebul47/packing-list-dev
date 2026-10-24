

"""
Banana Republic / GAP Canada — PO to Packing List
=================================================
Handles the GAP Canada PO PDF format (e.g. 62049433.pdf) where each item
row is glued, and the color name sits on the line below.

Run:  streamlit run br_packing.py
"""
import io
import re
import sys
from collections import OrderedDict
from datetime import date
from pathlib import Path

import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

import grid_editor
import sys as _sys
from pathlib import Path as _Path
_ROOT = str(_Path(__file__).resolve().parent.parent)
if _ROOT not in _sys.path:                     # pl_kit.py lives in the project root
    _sys.path.insert(0, _ROOT)
import pl_kit as kit                           # noqa: E402


# ============================================================
# 1. PDF parsing — line by line
# ============================================================
ITEM_LINE_RE = re.compile(
    r'^(?P<dept>\d{4})\s+'
    r'(?P<style>\d{7})\s+'
    r'(?P<prepack>Bulk|Single|Multi)\s+'
    r'(?P<prepack_type>Bulk|Single|Multi)\s+'
    r'(?P<sku>\d{9,})\s+'
    r'(?P<upp>Bulk|\d+)\s+'
    r'(?P<rest>.+)$'
)

TOTAL_LINE_RE = re.compile(r'^Total\s+(?P<full_style>\S+)\s+(?P<qty>[\d,]+)\s+'
                            r'(?P<color_glued>\S+)\s+(?P<cost>[\d,.]+)$')

PO_RE      = re.compile(r'Destination Purchase Order\s+\S+\s+(\d+)')
VENDOR_RE  = re.compile(r'Vendor\s+([A-Z][A-Za-z0-9 &().,\-/\']+?)\s*\(\d+\)')
FACTORY_RE = re.compile(r'Factory\s+([A-Z][A-Za-z0-9 &().,\-/\']+?)\s*\(\d+\)')
BUYER_RE   = re.compile(r'Brand:\s+([A-Z][A-Z ]+)')


def _parse_item_line(line: str):
    m = ITEM_LINE_RE.match(line)
    if not m:
        return None
    g = m.groupdict()
    rest = g["rest"].split()

    if len(rest) < 7:
        return None

    try:
        total_cost = float(rest[-1].replace(",", ""))
        unit_cost  = float(rest[-2].replace(",", ""))
        suffix     = rest[-3]
        size_num   = int(rest[-4])
        glued      = rest[-5]
        qty        = int(rest[-6].replace(",", ""))
        icpf       = int(rest[-7])
    except (ValueError, IndexError):
        return None

    if "-" in glued:
        dash = glued.index("-")
        color_start = dash - 7
        prepacks = glued[:color_start] or "00"
        color_code = glued[color_start:]
    else:
        prepacks = "00"
        color_code = glued

    desc = " ".join(rest[:-7]) if len(rest) > 7 else ""

    return {
        "dept":      g["dept"],
        "style":     g["style"],
        "sku":       g["sku"],
        "upp":       g["upp"],
        "desc":      desc,
        "icpf":      icpf,
        "qty":       qty,
        "prepacks":  prepacks,
        "color_code": color_code,
        "size":      size_num,
        "suffix":    suffix,
        "unit_cost": unit_cost,
        "total_cost": total_cost,
        "full_style": None,
    }


@st.cache_data(show_spinner=False, max_entries=5)
def parse_pdf(pdf_bytes: bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    lines = [ln.rstrip() for ln in text.splitlines()]

    items = []
    current_block_start = 0
    for i, line in enumerate(lines):
        line_s = line.strip()

        item = _parse_item_line(line_s)
        if item:
            color_name = ""
            for j in range(i + 1, min(i + 4, len(lines))):
                nxt = lines[j].strip()
                if not nxt:
                    continue
                if re.fullmatch(r'[A-Z][A-Z ]+', nxt):
                    color_name = nxt
                break
            item["color_name"] = color_name
            items.append(item)
            continue

        tm = TOTAL_LINE_RE.match(line_s)
        if tm:
            full_style = tm.group("full_style")
            for it in items[current_block_start:]:
                if it["full_style"] is None:
                    it["full_style"] = full_style
            current_block_start = len(items)

    po_m      = PO_RE.search(text)
    vendor_m  = VENDOR_RE.search(text)
    factory_m = FACTORY_RE.search(text)
    buyer_m   = BUYER_RE.search(text)

    colors: "OrderedDict[tuple, list]" = OrderedDict()
    for it in items:
        key = (it["color_code"], it["color_name"] or "UNKNOWN")
        colors.setdefault(key, []).append(it)

    return {
        "po_no":     po_m.group(1) if po_m else "",
        "vendor":    vendor_m.group(1).strip() if vendor_m else "",
        "factory":   factory_m.group(1).strip() if factory_m else "",
        "buyer":     buyer_m.group(1).strip() if buyer_m else "BANANA REPUBLIC",
        "style":     items[0]["style"] if items else "",
        "dept":      items[0]["dept"]  if items else "",
        "colors":    colors,
        "row_count": len(items),
    }


# ============================================================
# 2. Carton plan
# ============================================================
DEFAULT_CAP = 24
DEFAULT_EMPTY = {"G81": 0.91, "G82": 0.72, "G84": 0.41, "G85": 0.41}
DEFAULT_NET = {
    "25 R": 0.75, "26 R": 0.77, "27 R": 0.79, "28 R": 0.83,
    "29 R": 0.86, "30 R": 0.89, "31 R": 0.92, "32 R": 0.94,
}
DEFAULT_NN_OFFSET = 0.03


def carton_type_for(pcs: int, full_pcs: int) -> str:
    if pcs >= full_pcs:
        return "G81"
    if pcs >= 15:
        return "G82"
    if pcs >= 10:
        return "G85"
    return "G84"


def build_plan(sizes_qty: "OrderedDict[str, int]", cap: int):
    rows, cno = [], 1
    for size_label, qty in sizes_qty.items():
        full, rem = divmod(int(qty), cap)
        if full:
            rows.append({
                "ctn_from": cno, "ctn_to": cno + full - 1,
                "ctn_qty": full, "size": size_label, "pcs": cap,
                "ctn_type": "G81",
            })
            cno += full
        if rem:
            rows.append({
                "ctn_from": cno, "ctn_to": cno,
                "ctn_qty": 1, "size": size_label, "pcs": rem,
                "ctn_type": carton_type_for(rem, cap),
            })
            cno += 1
    return rows


def _size_sort_key(label: str):
    m = re.match(r"(\d+)", label or "")
    return (int(m.group(1)) if m else 0, label or "")


# ============================================================
# 3. Excel export
# ============================================================
FONT = "Courier New"
_thin = Side(style="thin", color="000000")
_box  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_CEN  = Alignment(horizontal="center", vertical="center")
_LEFT = Alignment(horizontal="left",   vertical="center")
_YELLOW = PatternFill("solid", fgColor="FFFF00")


def _put(ws, r, c, v=None, *, size=10, bold=False, align=_CEN, fmt=None, fill=None):
    cell = ws.cell(row=r, column=c)
    if v is not None:
        cell.value = v
    cell.font = Font(name=FONT, size=size, bold=bold)
    cell.alignment = align
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


def _merge(ws, r1, c1, r2, c2):
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def export_xlsx(po_data, weight_by_size, empty_by_type, cap, nn_offset=DEFAULT_NN_OFFSET, edited_plans=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "PKL"
    ws.sheet_view.showGridLines = False

    # size columns start at G (7); the trailing columns follow the widest colour block, so
    # a colour with 9+ sizes no longer collides with QTY PER CTN / TOTAL QTY / weights.
    max_n = max((len({f"{it['size']} {it['suffix'][:1].upper()}" for it in items})
                 for items in po_data["colors"].values()), default=8)
    C_QPC = max(15, 7 + max_n)
    C_TQ, C_GW, C_NW = C_QPC + 1, C_QPC + 2, C_QPC + 3
    LC = C_NW                                   # last used column
    QPC, TQ, GWL, NWL = L(C_QPC), L(C_TQ), L(C_GW), L(C_NW)

    widths = {"A": 12, "B": 8, "C": 10, "D": 9, "E": 16, "F": 20}
    for c in range(7, C_QPC):
        widths[L(c)] = 8
    widths[QPC] = 14
    widths[TQ] = 12
    widths[GWL] = 15
    widths[NWL] = 14
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    _merge(ws, 1, 1, 1, LC)
    _put(ws, 1, 1, "CREATIVE COLLECTIONS LTD", size=14, bold=True, align=_LEFT)
    _merge(ws, 2, 1, 2, LC)
    _put(ws, 2, 1, "Nishatnagar,Tongi,Gazipur,Dhaka.", size=11, align=_LEFT)
    _merge(ws, 3, 1, 3, LC)
    _put(ws, 3, 1, "PACKING LIST", size=12, bold=True)

    r = 5
    for section_idx, ((cc, cname), items) in enumerate(po_data["colors"].items(), 1):
        sizes_qty: "OrderedDict[str, int]" = OrderedDict()
        for it in sorted(items, key=lambda x: x["size"]):
            label = f"{it['size']} {it['suffix'][:1].upper()}"
            sizes_qty[label] = sizes_qty.get(label, 0) + it["qty"]
        size_list = list(sizes_qty.keys())
        if section_idx > 1:                     # one colour per printed page
            from openpyxl.worksheet.pagebreak import Break
            ws.row_breaks.append(Break(id=r - 1))

        plan_rows = (edited_plans.get((cc, cname)) if edited_plans else None) or build_plan(sizes_qty, cap)
        total_cartons = sum(int(p.get("total_ctn", p.get("ctn_qty", 0))) for p in plan_rows)
        order_qty = sum(sizes_qty.values())
        ship_qty = sum(int(p.get("total_pcs", 0) or 0) for p in plan_rows) if (edited_plans and (cc, cname) in edited_plans) else order_qty

        full_style = next((it["full_style"] for it in items if it["full_style"]),
                          po_data["style"])

        # ---------- dynamic header: SIZE / N.WT / N.N.WT ----------
        _put(ws, r, 5, "SIZE", bold=True, align=_LEFT)
        for i, s in enumerate(size_list):
            _put(ws, r, 7 + i, s, bold=True, align=_CEN)
        r += 1

        nw_row_idx = r  # SAVE THIS ROW INDEX
        _put(ws, r, 5, "N.WT", bold=True, align=_LEFT)
        for i, s in enumerate(size_list):
            _put(ws, r, 7 + i, round(weight_by_size.get(s, 0.0), 3), fmt="0.000")
        r += 1

        _put(ws, r, 5, "N.N. WT", bold=True, align=_LEFT)
        for i, s in enumerate(size_list):
            _put(ws, r, 7 + i, f"={L(7+i)}{r-1}-{nn_offset}", fmt="0.000")
        r += 2

        # yellow strip
        _merge(ws, r, 1, r, LC)
        _put(ws, r, 1, "BULK PACK", size=12, bold=True, fill=_YELLOW)
        r += 1

        # left PO block
        left_start = r
        _put(ws, r, 1, "BUYER             :", bold=True, align=_LEFT)
        _put(ws, r, 4, po_data["buyer"], align=_LEFT); r += 1
        _put(ws, r, 1, "STYLE             :", bold=True, align=_LEFT)
        _put(ws, r, 4, full_style, align=_LEFT); r += 1
        _put(ws, r, 1, "P. O. #           :", bold=True, align=_LEFT)
        _put(ws, r, 4, po_data["po_no"], align=_LEFT); r += 1
        _put(ws, r, 1, "O/QTY             :", bold=True, align=_LEFT)
        oq_r = r; _put(ws, r, 4, order_qty, align=_LEFT); _put(ws, r, 5, "PCS"); r += 1
        _put(ws, r, 1, "SHIP QTY         :", bold=True, align=_LEFT)
        sq_r = r; _put(ws, r, 4, ship_qty, align=_LEFT); _put(ws, r, 5, "PCS"); r += 1
        _put(ws, r, 1, "EX/SHORT        :", bold=True, align=_LEFT)
        xq_r = r; _put(ws, r, 4, ship_qty - order_qty, align=_LEFT); _put(ws, r, 5, "PCS"); r += 1
        _put(ws, r, 1, "CTN QTY          :", bold=True, align=_LEFT)
        cq_r = r; _put(ws, r, 4, total_cartons, align=_LEFT); _put(ws, r, 5, "CTNS"); r += 1
        for lbl in ["58.67 X 38.48 X 29.72 CM-G81",
                    "58.67 X 38.48 X 14.86 CM-G82",
                    "38.48 X 29.34 X 14.86 CM-G84",
                    "29.34 X 19.25 X 14.86 CM-G85"]:
            _put(ws, r, 1, "CTN MEAS.      :", bold=True, align=_LEFT)
            _put(ws, r, 4, lbl, align=_LEFT); r += 1

        rr = left_start + 3
        for lbl, val in [
            ("PO #",            po_data["po_no"]),
            ("SKU / Item",      full_style),
            ("Unit / Prepack",  f"{cap}/{cap // 2}"),
            ("CARTON",          f"1 OF {total_cartons}"),
        ]:
            _put(ws, rr, 12, lbl, bold=True, align=_LEFT)
            _put(ws, rr, 14, val, align=_LEFT)
            rr += 1

        r += 1

        for i, h in enumerate(["CTN", "", "CTN", "CTN QTY", "COLOR", "PREPACK STECKER"]):
            _put(ws, r, 1 + i, h, bold=True)
        _merge(ws, r, 7, r, 7 + len(size_list) - 1)
        _put(ws, r, 7, "SIZE", bold=True)
        _put(ws, r, C_QPC, "QTY PER CTN", bold=True)
        _put(ws, r, C_TQ,  "TOTAL QTY",   bold=True)
        _put(ws, r, C_GW,  "GROSS WEIGHT",bold=True)
        _put(ws, r, C_NW,  "NET WEIGHT",  bold=True)
        for c in range(1, LC + 1):
            ws.cell(row=r, column=c).border = _box
        r += 1

        _put(ws, r, 1, "NO", bold=True)
        _put(ws, r, 3, "MEA", bold=True)
        for i, s in enumerate(size_list):
            _put(ws, r, 7 + i, s, bold=True)
        for c in range(1, LC + 1):
            ws.cell(row=r, column=c).border = _box
        r += 1

        data_start = r
        cum_ctn = 0
        for idx, p in enumerate(plan_rows):
            tctn = int(p.get("total_ctn", p.get("ctn_qty", 0)))
            c_start = cum_ctn + 1
            cum_ctn += tctn
            c_end = cum_ctn
            c_type = p.get("ctn_type") or "G81"

            _put(ws, r, 1, 1 if idx == 0 else f"=B{r-1}+1", fmt="0")      # CTN from
            _put(ws, r, 2, f"=A{r}+D{r}-1", fmt="0")                       # CTN to
            _put(ws, r, 3, c_type)
            _put(ws, r, 4, tctn,  fmt="0")
            _put(ws, r, 5, cname if idx == 0 else "")
            _put(ws, r, 6, "")

            row_sizes = p.get("sizes") if "sizes" in p else {s: (p["pcs"] if s == p.get("size") else 0) for s in size_list}
            for i, s in enumerate(size_list):
                sz_val = row_sizes.get(s, 0)
                _put(ws, r, 7 + i, sz_val if sz_val else None, fmt="0")

            _put(ws, r, C_QPC, f"=SUM(G{r}:{L(6+len(size_list))}{r})", fmt="0")
            _put(ws, r, C_TQ, f"={QPC}{r}*D{r}", fmt="0")

            tare = float(empty_by_type.get(c_type, 0.0) or 0.0)
            # LIVE for ALL size columns (a size cell that is empty now can be filled in Excel later):
            #   NET = SUMPRODUCT(size qty row, N.WT row) ; GROSS = NET + empty carton
            _a, _z = L(7), L(6 + len(size_list))
            _net = f"SUMPRODUCT({_a}{r}:{_z}{r},{_a}${nw_row_idx}:{_z}${nw_row_idx})"
            _put(ws, r, C_GW, f"={_net}+{tare}", fmt="0.00")
            _put(ws, r, C_NW, f"={_net}", fmt="0.00")

            for c in range(1, LC + 1):
                ws.cell(row=r, column=c).border = _box
            r += 1
        data_end = r - 1

        # TOTAL row for the main table
        _merge(ws, r, 1, r, 3)
        _put(ws, r, 1, "TOTAL", bold=True)
        _put(ws, r, 4, f"=SUM(D{data_start}:D{data_end})", bold=True, fmt="0")
        _put(ws, r, 5, "CTNS", bold=True)
        for i, s in enumerate(size_list):
            _put(ws, r, 7 + i, f"=SUMPRODUCT({L(7+i)}{data_start}:{L(7+i)}{data_end},$D{data_start}:$D{data_end})", bold=True, fmt="0")   # pcs per carton x cartons
        _put(ws, r, C_TQ, f"=SUM({TQ}{data_start}:{TQ}{data_end})", bold=True, fmt="0")
        _put(ws, r, C_GW, f"=SUM({GWL}{data_start}:{GWL}{data_end})", bold=True, fmt="0.00")
        _put(ws, r, C_NW, f"=SUM({NWL}{data_start}:{NWL}{data_end})", bold=True, fmt="0.00")
        for c in range(1, LC + 1):
            ws.cell(row=r, column=c).border = _box

        total_row_idx = r
        r += 2  # blank row

        # ---------- SUMMARY TABLE (After Main Table) ----------
        n_sizes = len(size_list)
        total_col_idx = n_sizes + 3
        summary_start_row = r

        _merge(ws, r, 1, r, total_col_idx)
        _put(ws, r, 1, "SUMMARY", bold=True, fill=_YELLOW, align=_CEN)
        r += 1

        _put(ws, r, 1, "SIZE", bold=True, align=_LEFT)
        for i, s in enumerate(size_list):
            _put(ws, r, 2 + i, s, bold=True, align=_CEN)
        _put(ws, r, total_col_idx, "TOTAL", bold=True, align=_CEN)
        r += 1

        row_cut = r
        _put(ws, row_cut, 1, "CUT QTY", bold=True, align=_LEFT)
        for i in range(n_sizes):
            col = L(7 + i)
            _put(ws, row_cut, 2 + i, f"={col}{total_row_idx}", fmt="0")
        _put(ws, row_cut, total_col_idx, f"=SUM(B{row_cut}:{L(total_col_idx-1)}{row_cut})", fmt="0", bold=True)
        r += 1

        row_order = r
        _put(ws, row_order, 1, "ORDER QTY", bold=True, align=_LEFT)
        for i, s_lbl in enumerate(size_list):
            _put(ws, row_order, 2 + i, int(sizes_qty.get(s_lbl, 0)), fmt="0")      # PO quantity (input)
        _put(ws, row_order, total_col_idx, f"=SUM(B{row_order}:{L(total_col_idx-1)}{row_order})", fmt="0", bold=True)
        r += 1

        row_ship = r
        _put(ws, row_ship, 1, "SHIP QTY", bold=True, align=_LEFT)
        for i in range(n_sizes):
            col = L(7 + i)
            _put(ws, row_ship, 2 + i, f"={col}{total_row_idx}", fmt="0")
        _put(ws, row_ship, total_col_idx, f"=SUM(B{row_ship}:{L(total_col_idx-1)}{row_ship})", fmt="0", bold=True)
        r += 1

        row_ex = r
        _put(ws, row_ex, 1, "EX/SHORT", bold=True, align=_LEFT)
        for i in range(n_sizes):
            col = L(2 + i)
            _put(ws, row_ex, 2 + i, f"={col}{row_ship}-{col}{row_order}", fmt="0")
        _put(ws, row_ex, total_col_idx, f"=SUM(B{row_ex}:{L(total_col_idx-1)}{row_ex})", fmt="0", bold=True)
        r += 1

        row_pct = r
        _put(ws, row_pct, 1, "%", bold=True, align=_LEFT)
        for i in range(n_sizes):
            col = L(2 + i)
            _put(ws, row_pct, 2 + i, f"=IF({col}{row_order}=0, 0, {col}{row_ex}/{col}{row_order})", fmt="0.00")
        _put(ws, row_pct, total_col_idx, f"=IF({L(total_col_idx)}{row_order}=0, 0, {L(total_col_idx)}{row_ex}/{L(total_col_idx)}{row_order})", fmt="0.00", bold=True)
        r += 1

        for rr in range(summary_start_row, r):
            for cc in range(1, total_col_idx + 1):
                ws.cell(row=rr, column=cc).border = _box

        tc = L(total_col_idx)
        ws.cell(row=oq_r, column=4).value = f"={tc}{row_order}"
        ws.cell(row=sq_r, column=4).value = f"={tc}{row_ship}"
        ws.cell(row=xq_r, column=4).value = f"=D{sq_r}-D{oq_r}"
        ws.cell(row=cq_r, column=4).value = f"=D{total_row_idx}"

        r += 3

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0           # as many pages as colours (was 1 page = tiny print)
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def run(pdf_file=None):
    st.title("📦 Banana Republic / GAP Canada — Packing List")
    st.caption("Upload the GAP Canada PO PDF. Produces a BULK PACK packing list with live editable carton grid.")

    if pdf_file is None:
        pdf_file = st.file_uploader("Upload PO PDF", type=["pdf"], key="br_pdf_upload")

    if pdf_file is None:
        st.info("Upload a PDF to get started.")
        return

    try:
        data = parse_pdf(pdf_file.getvalue())
    except Exception as e:
        st.error(f"Could not parse the PDF: {e}")
        return

    if not data or data.get("row_count", 0) == 0:
        st.warning("No item rows were found.")
        return

    sizes_in_po = sorted(
        {f"{it['size']} {it['suffix'][:1].upper()}"
         for items in data["colors"].values() for it in items},
        key=_size_sort_key,
    )

    st.success(
        f"PO **{data['po_no']}** · Style **{data['style']}** · "
        f"**{len(data['colors'])}** color(s) · {data['row_count']} item row(s) parsed"
    )
    kit.step(1, "PO summary")
    kit.summary_cards([("PO", str(data["po_no"])), ("Style", str(data["style"])),
                       ("Colours", f"{len(data['colors'])}"), ("Sizes", f"{len(sizes_in_po)}"),
                       ("Item rows", f"{data['row_count']}")])

    kit.step(2, "Variables", "Not in the PO - edit if different. The carton tables below update straight away.")
    with st.container(border=True):
        st.markdown("**Carton settings**")
        cap = st.number_input("Pieces per full carton", 1, 200, DEFAULT_CAP, 1, key=f"br_cap_{data['po_no']}")

        st.markdown("**Net weight (kg / piece)**")
        st.caption(f"{len(sizes_in_po)} size(s) detected in this PO — edit values:")

        wdf = pd.DataFrame({
            "SIZE": sizes_in_po,
            "NET kg / pc": [float(DEFAULT_NET.get(s, 0.0)) for s in sizes_in_po],
        })
        edited_w = st.data_editor(
            wdf,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            key=f"br_w_ed_{data['po_no']}",
            column_config={
                "SIZE": st.column_config.TextColumn(required=True, width="small"),
                "NET kg / pc": st.column_config.NumberColumn(format="%.3f", step=0.001, min_value=0.0),
            },
        )
        weight_by_size = {
            str(row["SIZE"]).strip(): float(row["NET kg / pc"])
            for _, row in edited_w.iterrows()
            if str(row.get("SIZE", "")).strip() and pd.notna(row["NET kg / pc"])
        }

        st.markdown("**Net-net offset (kg)**")
        nn_offset = st.number_input(
            "N.WT − N.N.WT", 0.0, 1.0, float(DEFAULT_NN_OFFSET), 0.01,
            key=f"br_nno_{data['po_no']}",
            help="Amount subtracted from N.WT to produce N.N.WT in the header.",
        )

        st.markdown("**Empty carton weight (kg)**")
        edf = pd.DataFrame({
            "TYPE": list(DEFAULT_EMPTY.keys()),
            "EMPTY kg": [DEFAULT_EMPTY[t] for t in DEFAULT_EMPTY],
        }).set_index("TYPE")
        edited_e = st.data_editor(
            edf, use_container_width=True,
            key=f"br_e_ed_{data['po_no']}",
            column_config={"EMPTY kg": st.column_config.NumberColumn(format="%.2f", step=0.01)}
        )
        empty_by_type = {t: float(edited_e.loc[t, "EMPTY kg"]) for t in edited_e.index}

    kit.step(3, "Carton tables",
             "One editable grid per colour. Click a cell to change it, add or delete carton rows in the grid; "
             "the ORDER QTY / TOTAL / REMAIN rows update live and edited cells get a green border.")
    edited_plans = {}
    order_by = {}

    for (cc, cname), items in data["colors"].items():
        full_style = next((it["full_style"] for it in items if it["full_style"]), "—")
        st.markdown(f"#### Color: **{cname}** ({cc}) · Style {full_style}")

        sizes_qty: "OrderedDict[str, int]" = OrderedDict()
        for it in sorted(items, key=lambda x: x["size"]):
            label = f"{it['size']} {it['suffix'][:1].upper()}"
            sizes_qty[label] = sizes_qty.get(label, 0) + it["qty"]
        size_list = list(sizes_qty.keys())
        order_by[cname] = dict(sizes_qty)

        # Generate initial default plan
        plan = build_plan(sizes_qty, int(cap))
        initial_rows = []
        for p in plan:
            initial_rows.append({
                "color": cname,
                "ctn_type": p["ctn_type"],
                "sizes": {s: (p["pcs"] if s == p["size"] else 0) for s in size_list},
                "total_ctn": p["ctn_qty"],
            })

        # Render VF-style editable carton grid
        grid_res = grid_editor.render_carton_grid(
            key_prefix=f"br_{data['po_no']}_{cc}",
            sizes=size_list,
            order_qty=sizes_qty,
            initial_rows=initial_rows,
            frozen_columns=[
                {"header": "Color", "key": "color", "width": 120},
                {"header": "Type", "key": "ctn_type", "width": 80},
            ],
            net_weight_by_size=weight_by_size,
            empty_carton_weight=float(empty_by_type.get("G81", 0.0) or 0.0),
            show_weights=True,
            show_metrics=True,
            title=f"Main Carton Table (Editable) — {cname}",
        )
        edited_plans[(cc, cname)] = grid_res["rows"]

    # ---- 4 check (all colours together)
    kit.step(4, "Check against the PO")
    labels, order_d, packed_d = [], {}, {}
    for (cc, cname), rows_ in edited_plans.items():
        for sz, q in order_by.get(cname, {}).items():
            lab = f"{cname} / {sz}" if len(edited_plans) > 1 else sz
            labels.append(lab)
            order_d[lab] = q
            packed_d[lab] = sum(int((r.get("sizes") or {}).get(sz, 0) or 0) * int(r.get("total_ctn", 0) or 0)
                                for r in rows_)
    if labels:
        kit.check_by_size(labels, order_d, packed_d)
        kit.summary_cards([
            ("Cartons", f"{sum(int(r.get('total_ctn', 0) or 0) for rs in edited_plans.values() for r in rs):,}"),
            ("Pieces", f"{sum(packed_d.values()):,}"),
            ("Order pcs", f"{sum(order_d.values()):,}"),
        ])

    kit.step(5, "Download")
    prep_col1, prep_col2 = st.columns([1, 2])
    with prep_col1:
        if st.button("Prepare packing list (.xlsx)", type="primary", key=f"br_prep_{data['po_no']}"):
            st.session_state[f"br_xlsx_{data['po_no']}"] = export_xlsx(
                data, weight_by_size, empty_by_type, int(cap), float(nn_offset),
                edited_plans=edited_plans,
            )

    if st.session_state.get(f"br_xlsx_{data['po_no']}"):
        st.download_button(
            "⬇ Download packing list (.xlsx)",
            data=st.session_state[f"br_xlsx_{data['po_no']}"],
            file_name=f"packing_list_{data['po_no']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"br_dl_{data['po_no']}",
        )


if __name__ == "__main__":
    st.set_page_config(page_title="BR / GAP Canada — Packing List", layout="wide", page_icon="📦")
    run()