"""
PO -> Packing List Streamlit App
=================================
Upload a VF/Vans-style PO PDF, review the extracted data, and download
a packing-list-style Excel workbook (one sheet per PO line).

HOW TO RUN:
    pip install streamlit pdfplumber openpyxl reportlab pandas
    streamlit run VF_Packing_app.py
"""
import io
import json
import re
import uuid
import datetime

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# ============================================================
# PART 1 - PDF EXTRACTION
# ============================================================

LINE_PATTERN = re.compile(
    r"""
    (?P<po_line_no>\d{15})
    \s+
    (?P<style>[A-Z0-9]+)
    \s+
    (?P<description>[A-Z0-9 \-/]+?)
    \s+
    (?P<order_qty>[\d,]+)
    \s+
    (?:PIECE|EACH)
    \s+
    (?P<unit_price>[\d.]+)
    \s+
    (?P<line_cost>[\d,]+\.\d{2})
    """,
    re.IGNORECASE | re.VERBOSE,
)

SUBLINE_START_PATTERN = re.compile(
    r"(?P<subline_no>\d{12,13})\s+(?P<style>VN[A-Z0-9]+)",
    re.IGNORECASE,
)


def extract_sublines(full_text):
    text = full_text.replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
    matches = list(SUBLINE_START_PATTERN.finditer(text))
    sublines = []

    for i, start_match in enumerate(matches):
        block_start = start_match.start()
        block_end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        block = text[block_start:block_end]

        if not re.search(r"UnitOfMeasureCode\s+(?:PC|EA)\b", block, re.I):
            continue

        qty_cost = re.search(
            r"(?P<style>VN[A-Z0-9]+)\s+.*?"
            r"(?P<qty>[\d,]+)\s+(?P<unit_cost>\d+(?:\.\d+)?)\s+.*?"
            r"UnitOfMeasureCode",
            block, re.IGNORECASE | re.DOTALL,
        )
        color = re.search(
            r"UnitOfMeasureCode\s+(?:PC|EA)\s+Color\s+(?P<color>.*?)\s+Size",
            block, re.IGNORECASE | re.DOTALL,
        )
        size = re.search(
            r"Size\s+(?P<size>.*?)\s+Dimension\s+1",
            block, re.IGNORECASE | re.DOTALL,
        )
        sku = re.search(r"SKU\s+Number\s+(?P<sku>\S+)", block, re.IGNORECASE)
        upc = re.search(r"UPC\s+Number\s+(?P<upc>\d{12,13})", block, re.IGNORECASE)
        pack = re.search(
            r"Items\s+Per\s+Outer\s+(?:Pack\s+)?(?P<pack>[\d,]+)",
            block, re.IGNORECASE | re.DOTALL,
        )

        if not all((qty_cost, color, size, sku, upc, pack)):
            continue

        sublines.append({
            "subline_no": start_match.group("subline_no"),
            "style": start_match.group("style"),
            "qty": int(qty_cost.group("qty").replace(",", "")),
            "unit_cost": float(qty_cost.group("unit_cost").replace(",", "")),
            "color": re.sub(r"\s+", " ", color.group("color")).strip(),
            "size": re.sub(r"\s+", " ", size.group("size")).strip(),
            "sku": sku.group("sku"),
            "upc": upc.group("upc"),
            "items_per_outer_pack": int(pack.group("pack").replace(",", "")),
        })

    return sublines


CRD_PATTERN = re.compile(r"Brand Requested CRD\s+(\d{4}-\d{2}-\d{2})")
COUNTRY_PATTERN = re.compile(r"([A-Z]{3,})\s*\na\s*\na\s*\nBrand Buyer Code")


def extract_data(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join((page.extract_text() or "") for page in pdf.pages)

    lines = []
    for m in LINE_PATTERN.finditer(full_text):
        lines.append({
            "po_line_no": m.group("po_line_no"),
            "style": m.group("style"),
            "description": re.sub(r"\s+", " ", m.group("description")).strip(),
            "order_qty": int(m.group("order_qty").replace(",", "")),
            "unit_price": float(m.group("unit_price").replace(",", "")),
            "line_cost": float(m.group("line_cost").replace(",", "")),
        })

    sublines = extract_sublines(full_text)

    crd_matches = CRD_PATTERN.findall(full_text)
    for i, l in enumerate(lines):
        l["crd"] = crd_matches[i] if i < len(crd_matches) else None

    country_match = COUNTRY_PATTERN.search(full_text)
    destination_country = country_match.group(1) if country_match else None
    for l in lines:
        l["destination_country"] = destination_country

    return {"lines": lines, "sublines": sublines}


@st.cache_data(show_spinner=False, max_entries=10)
def extract_data_cached(pdf_bytes: bytes):
    return extract_data(io.BytesIO(pdf_bytes))


_HAS_FRAGMENT = hasattr(st, "fragment")


def _fragment(fn):
    return st.fragment(fn) if _HAS_FRAGMENT else fn


# ============================================================
# PART 2 - CARTON ROW COMPUTATION
# ============================================================

def compute_carton_rows_v3(sublines, max_by_size, mode="solid_mixed"):
    if not sublines:
        return []

    size_order = [s["size"] for s in sublines]
    full_rows = []
    partial_rows = []
    remainder_pool = []

    for s in sublines:
        size = s["size"]
        qty = int(s["qty"])
        pack = int(max_by_size.get(size) or s.get("items_per_outer_pack") or 1)
        if pack < 1:
            pack = 1

        full_ctn = qty // pack
        rem = qty % pack

        if full_ctn > 0:
            full_rows.append({
                "sizes": {size: pack},
                "total_ctn": full_ctn,
                "ctn_pcs": pack,
                "color": s["color"],
                "upc": s["upc"],
                "mixed": False,
            })

        if rem > 0:
            if mode == "solid_no_leftover":
                pass
            elif mode == "solid_all":
                partial_rows.append({
                    "sizes": {size: rem},
                    "total_ctn": 1,
                    "ctn_pcs": rem,
                    "color": s["color"],
                    "upc": s["upc"],
                    "mixed": False,
                })
            else:
                remainder_pool.append({
                    "size": size,
                    "rem": rem,
                    "color": s["color"],
                    "upc": s["upc"],
                    "max": pack,
                })

    size_to_full = {}
    for r in full_rows:
        for sz in r["sizes"]:
            size_to_full.setdefault(sz, []).append(r)

    size_to_partial = {}
    for r in partial_rows:
        for sz in r["sizes"]:
            size_to_partial.setdefault(sz, []).append(r)

    solid_block = []
    for sz in size_order:
        for r in size_to_full.get(sz, []):
            solid_block.append(r)
        for r in size_to_partial.get(sz, []):
            solid_block.append(r)

    if mode in ("solid_no_leftover", "solid_all"):
        return solid_block

    mixed_rows = []
    idx = 0
    while idx < len(remainder_pool):
        carton = {}
        carton_total = 0
        carton_cap = None
        progressed = False

        while idx < len(remainder_pool):
            entry = remainder_pool[idx]
            size = entry["size"]
            size_max = entry["max"]
            proposed_cap = size_max if carton_cap is None else min(carton_cap, size_max)

            if carton_total > proposed_cap:
                break
            space = proposed_cap - carton_total
            if space <= 0:
                break

            take = min(entry["rem"], space)
            if take > 0:
                carton[size] = carton.get(size, 0) + take
                carton_total += take
                entry["rem"] -= take
                carton_cap = proposed_cap
                progressed = True

            if entry["rem"] == 0:
                idx += 1
            else:
                break

            if carton_total >= carton_cap:
                break

        if carton:
            mixed_rows.append({
                "sizes": dict(carton),
                "total_ctn": 1,
                "ctn_pcs": carton_total,
                "color": None,
                "upc": None,
                "mixed": True,
            })
        if not progressed:
            break

    return solid_block + mixed_rows


# ============================================================
# PART 3 - EXCEL BUILDING
# ============================================================

FONT = "Arial"
bold = Font(name=FONT, bold=True, size=10)
normal = Font(name=FONT, size=10)
title_font = Font(name=FONT, bold=True, size=13)
thin = Side(style="thin")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
yellow = PatternFill("solid", fgColor="FFFF00")
grey = PatternFill("solid", fgColor="D9D9D9")


def sheet_name_for(line):
    raw = f"{line['style']}"
    return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


def build_sheet(ws, line, sublines):
    sizes = [s["size"] for s in sublines]
    n_sizes = len(sizes)

    ws.merge_cells("A1:H1")
    ws["A1"] = "CREATIVE COLLECTIONS LTD"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    ws.merge_cells("A2:H2")
    ws["A2"] = "PACKING LIST DETAILS"
    ws["A2"].font = Font(name=FONT, bold=True, size=10)
    ws["A2"].alignment = center

    spec_header_row = 4
    ws.cell(row=spec_header_row, column=1, value="SIZE").font = bold
    ws.cell(row=spec_header_row, column=1).border = box
    for sc_idx, size_name in enumerate(sizes):
        c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
        c.font = bold
        c.alignment = center
        c.border = box
        c.fill = grey

    spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
    for i, label in enumerate(spec_rows):
        rr = spec_header_row + 1 + i
        lc = ws.cell(row=rr, column=1, value=label)
        lc.font = bold
        lc.border = box
        for sc_idx, size_name in enumerate(sizes):
            c = ws.cell(row=rr, column=2 + sc_idx, value=1)
            c.alignment = center
            c.border = box

    r = spec_header_row + len(spec_rows) + 2
    ws.cell(row=r, column=1, value="BUYER").font = bold
    ws.cell(row=r, column=2, value="VANS").font = normal
    r += 1
    ws.cell(row=r, column=1, value="LOT NO").font = bold
    ws.cell(row=r, column=2, value=line["style"]).font = normal
    r += 1
    ws.cell(row=r, column=1, value="P.O NO").font = bold
    ws.cell(row=r, column=2, value=line["po_line_no"]).font = normal
    r += 1
    ws.cell(row=r, column=1, value="ORDER QTY").font = bold
    ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
    ws.cell(row=r, column=3, value="PCS").font = normal
    r += 1
    ws.cell(row=r, column=1, value="PACK QTY").font = bold
    pack_qty_row = r
    ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
    ws.cell(row=r, column=3, value="PCS").font = normal
    r += 1
    excess_short_row = r
    ws.cell(row=r, column=1, value="EXCESS/SHORT QTY").font = bold
    ws.cell(row=r, column=3, value="PCS").font = normal
    r += 1
    percentage_row = r
    ws.cell(row=r, column=1, value="PERCENTAGE").font = bold
    r += 1
    ws.cell(row=r, column=1, value="CRD").font = bold
    ws.cell(row=r, column=2, value=line.get("crd")).font = normal
    r += 1
    ws.cell(row=r, column=1, value="COUNTRY").font = bold
    ws.cell(row=r, column=2, value=line.get("destination_country")).font = normal
    r += 1
    ws.cell(row=r, column=1, value="DESCRIPTION").font = bold
    ws.cell(row=r, column=2, value=line["description"]).font = normal

    header_row = r + 2
    headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
              ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
               "total.Grs.wt", "total.net.wt"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = bold
        c.alignment = center
        c.border = box
        c.fill = grey
    n_cols = len(headers)

    pdf_max_by_size = {s["size"]: int(s.get("items_per_outer_pack") or 1) for s in sublines}
    data_row_start = header_row + 1
    carton_rows = compute_carton_rows_v3(sublines, pdf_max_by_size, mode="solid_mixed")
    for idx, row_data in enumerate(carton_rows):
        rr = data_row_start + idx
        ws.cell(row=rr, column=1, value=idx + 1).border = box

        color_val = row_data["color"] if row_data["color"] is not None else "MIXED"
        color_cell = ws.cell(row=rr, column=3, value=color_val)
        color_cell.border = box
        if row_data["mixed"]:
            color_cell.fill = yellow
        ws.cell(row=rr, column=4, value=row_data["upc"]).border = box
        for sc_idx, size_name in enumerate(sizes):
            col = 5 + sc_idx
            val = row_data["sizes"].get(size_name)
            cell = ws.cell(row=rr, column=col, value=val)
            cell.border = box
            cell.alignment = center

        ctn_pcs_col = 5 + n_sizes
        first_size_letter = get_column_letter(5)
        last_size_letter = get_column_letter(4 + n_sizes)
        ctn_cell = ws.cell(
            row=rr, column=ctn_pcs_col,
            value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
        )
        ctn_cell.border = box
        ctn_cell.alignment = center
        ctn_cell.font = bold

        total_ctn_col = ctn_pcs_col + 1
        total_pcs_col = ctn_pcs_col + 2
        total_ctn_letter = get_column_letter(total_ctn_col)

        tc_cell = ws.cell(row=rr, column=total_ctn_col, value=row_data["total_ctn"])
        tc_cell.border = box
        tc_cell.alignment = center

        if idx == 0:
            carton_formula = f"={total_ctn_letter}{rr}"
        else:
            carton_formula = f"={get_column_letter(2)}{rr - 1}+{total_ctn_letter}{rr}"
        cB = ws.cell(row=rr, column=2, value=carton_formula)
        cB.border = box
        cB.alignment = center

        ctn_pcs_letter = get_column_letter(ctn_pcs_col)
        tp_cell = ws.cell(
            row=rr, column=total_pcs_col,
            value=f"={total_ctn_letter}{rr}*{ctn_pcs_letter}{rr}"
        )
        tp_cell.border = box
        tp_cell.alignment = center

        grs_wt_col = total_pcs_col + 1
        nw_row = spec_header_row + 1
        empty_row = spec_header_row + 3
        first_spec_letter = get_column_letter(2)
        last_spec_letter = get_column_letter(1 + n_sizes)
        first_main_size_letter = get_column_letter(5)
        last_main_size_letter = get_column_letter(4 + n_sizes)

        grs_formula = (
            f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
            f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
            f"+{first_spec_letter}{empty_row}"
        )
        grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
        grs_cell.border = box
        grs_cell.alignment = center

        net_wt_col = grs_wt_col + 1
        grs_wt_letter_for_net = get_column_letter(grs_wt_col)
        net_cell = ws.cell(
            row=rr, column=net_wt_col,
            value=f"={grs_wt_letter_for_net}{rr}-{first_spec_letter}{empty_row}"
        )
        net_cell.border = box
        net_cell.alignment = center

        total_grs_col = net_wt_col + 1
        grs_wt_letter = get_column_letter(grs_wt_col)
        tg_cell = ws.cell(
            row=rr, column=total_grs_col,
            value=f"={total_ctn_letter}{rr}*{grs_wt_letter}{rr}"
        )
        tg_cell.border = box
        tg_cell.alignment = center

        total_net_col = total_grs_col + 1
        net_wt_letter = get_column_letter(net_wt_col)
        tn_cell = ws.cell(
            row=rr, column=total_net_col,
            value=f"={total_ctn_letter}{rr}*{net_wt_letter}{rr}"
        )
        tn_cell.border = box
        tn_cell.alignment = center

    total_row = data_row_start + len(carton_rows)
    ws.cell(row=total_row, column=3, value="TOTAL").font = bold
    total_ctn_letter_for_sum = get_column_letter(total_ctn_col)
    for sc_idx, size_name in enumerate(sizes):
        col = 5 + sc_idx
        col_letter = get_column_letter(col)
        cell = ws.cell(
            row=total_row, column=col,
            value=f"=SUMPRODUCT({col_letter}{data_row_start}:{col_letter}{total_row-1},"
                  f"{total_ctn_letter_for_sum}{data_row_start}:{total_ctn_letter_for_sum}{total_row-1})"
        )
        cell.font = bold
        cell.border = box

    total_ctn_col = ctn_pcs_col + 1
    total_pcs_col = ctn_pcs_col + 2
    total_ctn_letter = get_column_letter(total_ctn_col)
    total_pcs_letter = get_column_letter(total_pcs_col)

    ws.cell(row=total_row, column=total_ctn_col,
            value=f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
                  f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))').font = bold

    ws.cell(row=total_row, column=total_pcs_col,
            value=f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
                  f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))').font = bold

    grs_wt_col = total_pcs_col + 1
    net_wt_col = total_pcs_col + 2
    grs_wt_letter = get_column_letter(grs_wt_col)
    net_wt_letter = get_column_letter(net_wt_col)

    ws.cell(row=total_row, column=grs_wt_col,
            value=f"=SUMPRODUCT({grs_wt_letter}{data_row_start}:{grs_wt_letter}{total_row-1},"
                  f"{total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})").font = bold
    ws.cell(row=total_row, column=net_wt_col,
            value=f"=SUMPRODUCT({net_wt_letter}{data_row_start}:{net_wt_letter}{total_row-1},"
                  f"{total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})").font = bold

    total_grs_col = total_pcs_col + 3
    total_grs_letter = get_column_letter(total_grs_col)
    ws.cell(row=total_row, column=total_grs_col,
            value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
                  f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))').font = bold

    total_net_col = total_grs_col + 1
    total_net_letter = get_column_letter(total_net_col)
    ws.cell(row=total_row, column=total_net_col,
            value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
                  f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))').font = bold

    sum_header_row = total_row + 3
    ws.cell(row=sum_header_row, column=1, value="SIZE").font = bold
    for sc_idx, size_name in enumerate(sizes):
        ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name).font = bold
    ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL").font = bold

    gw_col = 3 + n_sizes
    nw_col = 4 + n_sizes
    ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
                    end_row=sum_header_row, end_column=gw_col)
    ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)").font = bold
    ws.cell(row=sum_header_row, column=gw_col).alignment = center
    ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
                    end_row=sum_header_row, end_column=nw_col)
    ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)").font = bold
    ws.cell(row=sum_header_row, column=nw_col).alignment = center

    rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
    for i, label in enumerate(rows_needed):
        rr = sum_header_row + 1 + i
        ws.cell(row=rr, column=1, value=label).font = bold
        for sc_idx, size_name in enumerate(sizes):
            col = 2 + sc_idx
            col_letter = get_column_letter(col)
            cell = ws.cell(row=rr, column=col)
            if label == "ORDER QTY":
                qty = next(s["qty"] for s in sublines if s["size"] == size_name)
                cell.value = qty
            elif label == "CUT QTY":
                cell.fill = yellow
            elif label == "SHIP QTY":
                main_size_col = 5 + sc_idx
                main_size_letter = get_column_letter(main_size_col)
                cell.value = f"={main_size_letter}{total_row}"
            elif label == "EXS/SHT QTY":
                ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
                order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
                cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
            elif label == "PERCENTAGE":
                exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
                order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
                cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
                cell.number_format = "0.00%"
            cell.border = box
            cell.alignment = center

        gt_col = 2 + n_sizes
        gt_letter = get_column_letter(gt_col)
        first_letter = get_column_letter(2)
        last_letter = get_column_letter(1 + n_sizes)
        if label in ("CUT QTY", "EXS/SHT QTY"):
            ws.cell(row=rr, column=gt_col,
                    value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))').font = bold
        elif label in ("ORDER QTY", "SHIP QTY"):
            ws.cell(row=rr, column=gt_col,
                    value=f"=SUM({first_letter}{rr}:{last_letter}{rr})").font = bold
        else:
            exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
            order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
            ws.cell(row=rr, column=gt_col,
                    value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))').number_format = "0.00%"
        ws.cell(row=rr, column=gt_col).border = box

    exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
    order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
    gt_col_letter = get_column_letter(2 + n_sizes)

    ship_gt_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
    ws.cell(row=pack_qty_row, column=2, value=f'={gt_col_letter}{ship_gt_row}').font = normal
    ws.cell(row=excess_short_row, column=2,
            value=f'={gt_col_letter}{exs_gt_row}').font = normal
    ws.cell(row=percentage_row, column=2,
            value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
                  f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))'
            ).font = normal

    wt_row = sum_header_row + len(rows_needed) + 2
    total_grs_letter = get_column_letter(total_grs_col)
    total_net_letter = get_column_letter(total_net_col)
    total_ctn_letter_final = get_column_letter(total_ctn_col)

    ws.cell(row=wt_row, column=1, value="GROSS WEIGHT :").font = bold
    c = ws.cell(row=wt_row, column=2, value=f"={total_grs_letter}{total_row}")
    c.font = bold
    c.border = box
    c.alignment = center

    ws.cell(row=wt_row + 1, column=1, value="NET WEIGHT :").font = bold
    c = ws.cell(row=wt_row + 1, column=2, value=f"={total_net_letter}{total_row}")
    c.font = bold
    c.border = box
    c.alignment = center

    ws.cell(row=wt_row + 2, column=1, value="CTN MEAS :").font = bold
    ws.cell(row=wt_row + 2, column=2).fill = yellow
    ws.cell(row=wt_row + 2, column=2).border = box
    ws.cell(row=wt_row + 3, column=2).fill = yellow
    ws.cell(row=wt_row + 3, column=2).border = box

    ws.cell(row=wt_row + 4, column=1, value="CBM :").font = bold
    ws.cell(row=wt_row + 4, column=2).fill = yellow
    ws.cell(row=wt_row + 4, column=2).border = box
    cbm_result = ws.cell(
        row=wt_row + 4, column=3,
        value=f'=IF(B{wt_row + 4}="","",B{wt_row + 4}*{total_ctn_letter_final}{total_row})'
    )
    cbm_result.font = bold
    cbm_result.border = box
    cbm_result.alignment = center

    first_sum_row = sum_header_row + 1
    last_sum_row = sum_header_row + len(rows_needed)

    ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
                    end_row=last_sum_row, end_column=gw_col)
    gw_cell = ws.cell(row=first_sum_row, column=gw_col,
                       value=f'={total_grs_letter}{total_row}')
    gw_cell.alignment = center
    gw_cell.font = bold
    gw_cell.border = box

    ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
                    end_row=last_sum_row, end_column=nw_col)
    nw_cell = ws.cell(row=first_sum_row, column=nw_col,
                       value=f'={total_net_letter}{total_row}')
    nw_cell.alignment = center
    nw_cell.font = bold
    nw_cell.border = box

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    for i in range(n_sizes):
        ws.column_dimensions[get_column_letter(5 + i)].width = 10
    for col in range(5 + n_sizes, n_cols + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions[get_column_letter(gw_col)].width = 22
    ws.column_dimensions[get_column_letter(nw_col)].width = 22

    legend_row = wt_row + 6
    ws.cell(row=legend_row, column=1, value="Legend:").font = bold
    ws.cell(row=legend_row + 1, column=1,
            value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)"
            ).font = Font(name=FONT, size=9, italic=True)


def build_workbook(data):
    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()
    for line in data["lines"]:
        sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
        if not sublines:
            continue
        name = sheet_name_for(line)
        base_name, i = name, 1
        while name in used_names:
            i += 1
            name = f"{base_name[:28]}_{i}"
        used_names.add(name)
        ws = wb.create_sheet(title=name)
        build_sheet(ws, line, sublines)
    return wb


def build_edited_pdf(export_state):
    line = export_state["line"]
    full_sizes = export_state["full_sizes"]
    rows = export_state["rows"]

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=5 * mm, rightMargin=5 * mm, topMargin=10 * mm, bottomMargin=10 * mm,
    )
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("CREATIVE COLLECTIONS LTD", styles["Title"]))
    elements.append(Paragraph("PACKING LIST DETAILS", styles["Heading3"]))
    elements.append(Spacer(1, 6))

    info_data = [
        ["BUYER", "VANS", "LOT NO", line["style"], "P.O NO", line["po_line_no"]],
        ["ORDER QTY", line["order_qty"], "EXCESS/SHORT QTY", export_state["exs_total"],
         "CRD", line.get("crd") or ""],
        ["COUNTRY", line.get("destination_country") or "", "DESCRIPTION",
         line["description"], "", ""],
    ]
    info_table = Table(info_data, colWidths=[65, 90, 90, 140, 65, 90])
    info_table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.black),
        ("BACKGROUND", (0, 0), (0, -1), rl_colors.whitesmoke),
        ("BACKGROUND", (2, 0), (2, -1), rl_colors.whitesmoke),
        ("BACKGROUND", (4, 0), (4, -1), rl_colors.whitesmoke),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 8))

    headers = ["Ctn No", "Color", "UPC"] + full_sizes + [
        "CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs/ctn", "Net/ctn", "tot Grs", "tot Net"
    ]
    table_data = [headers]
    for rdata in rows:
        table_data.append([
            rdata["ctn_no"], rdata["color"], rdata["upc"],
            *[rdata["sizes"].get(sz, "") for sz in full_sizes],
            rdata["ctn_pcs"], rdata["total_ctn"], rdata["total_pcs"],
            f'{rdata["grs_per_ctn"]:.2f}', f'{rdata["net_per_ctn"]:.2f}',
            f'{rdata["total_grs"]:.2f}', f'{rdata["total_net"]:.2f}',
        ])
    total_row = (
        ["", "TOTAL", ""]
        + [export_state["size_packet_totals"].get(sz, 0) for sz in full_sizes]
        + ["", export_state["total_ctn"], export_state["total_pcs"], "", "",
           f'{export_state["total_grs"]:.2f}', f'{export_state["total_net"]:.2f}']
    )
    table_data.append(total_row)

    main_table = Table(table_data, repeatRows=1)
    main_table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.lightgrey),
        ("BACKGROUND", (0, -1), (-1, -1), rl_colors.lightgrey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(main_table)
    elements.append(Spacer(1, 10))

    def fmt_pct(v):
        return "" if v is None else f"{v * 100:.2f}%"

    summary_data = [["SIZE"] + full_sizes + ["G TOTAL"]]
    summary_data.append(
        ["CUT QTY"] + [export_state["cut_qty"].get(sz, "") for sz in full_sizes]
        + [export_state["cut_total"]]
    )
    summary_data.append(
        ["ORDER QTY"] + [export_state["order_qty_by_size"].get(sz, "") for sz in full_sizes]
        + [export_state["order_total"]]
    )
    summary_data.append(
        ["SHIP QTY"] + [export_state["ship_qty_by_size"].get(sz, "") for sz in full_sizes]
        + [export_state["ship_total"]]
    )
    summary_data.append(
        ["EXS/SHT QTY"] + [export_state["exs_by_size"].get(sz, "") for sz in full_sizes]
        + [export_state["exs_total"]]
    )
    summary_data.append(
        ["PERCENTAGE"] + [fmt_pct(export_state["pct_by_size"].get(sz)) for sz in full_sizes]
        + [fmt_pct(export_state["pct_total"])]
    )

    summary_table = Table(summary_data, repeatRows=1)
    summary_table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.lightgrey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 10))

    wt_data = [
        ["GROSS WEIGHT :", f'{export_state["total_grs"]:.2f}'],
        ["NET WEIGHT :", f'{export_state["total_net"]:.2f}'],
        ["CTN MEAS :", f'{export_state["ctn_meas_1"]} {export_state["ctn_meas_2"]}'.strip()],
        ["CBM :", f'{export_state["cbm_total"]:.8f}'],
    ]
    wt_table = Table(wt_data, colWidths=[100, 150])
    wt_table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
    ]))
    elements.append(wt_table)
    elements.append(Spacer(1, 20))

    sign_table = Table([["INCHARGE", "", "FM"]], colWidths=[120, 120, 120])
    sign_table.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (0, 0), 0.75, rl_colors.black),
        ("LINEABOVE", (2, 0), (2, 0), 0.75, rl_colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    elements.append(sign_table)

    doc.build(elements)
    buf.seek(0)
    return buf


# ============================================================
# PART 4 - STREAMLIT UI & PREVIEW
# ============================================================


@_fragment

# def render_packing_list_preview(line, sublines):
#     style_key = line["style"]
#     full_sizes = list(dict.fromkeys([s["size"] for s in sublines])) # deduplicate while preserving order
#     default_color = next((s["color"] for s in sublines if s.get("color")), "MIXED")
#     pdf_max_by_size = {s["size"]: int(s.get("items_per_outer_pack") or 1) for s in sublines}

#     # ---- Weight spec table (native data_editor) ----
#     st.markdown("##### Weight spec (N.W. / N.N.W. / EMPTY CTN)")
#     spec_df = pd.DataFrame(
#         [[1.0] * len(full_sizes) for _ in range(3)],
#         index=["N.W.", "N.N.W.", "EMPTY CTN"],
#         columns=full_sizes,
#     )
#     spec_edit = st.data_editor(
#         spec_df, use_container_width=True, key=f"spec_v10_{style_key}",
#     )
#     nw_by_size = spec_edit.loc["N.W."].to_dict()
#     nnw_by_size = spec_edit.loc["N.N.W."].to_dict()
#     empty_by_size = spec_edit.loc["EMPTY CTN"].to_dict()
#     empty_ctn_wt = float(next(iter(empty_by_size.values()), 0.0))

#     # ---- Order quantity by size (HORIZONTAL: sizes as columns, no visible headers) ----
#     st.markdown("##### Order quantity by size")

#     ZWSP = "\u200b"  # zero-width space — unique per column, invisible in UI
#     order_data = {
#         ZWSP: ["SIZE", "COLOR", "UPC", "ORDER QTY"],
#     }
#     for i, s in enumerate(sublines):
#         order_data[f"{ZWSP}{i + 1}"] = [s["size"], s["color"], s["upc"], s["qty"]]

#     order_df = pd.DataFrame(order_data)
#     st.dataframe(order_df, use_container_width=True, hide_index=True)

#     # ---- MAX per carton (native data_editor) ----
#     st.markdown("##### Packing qty (MAX per carton)")
#     max_df = pd.DataFrame(
#         [[pdf_max_by_size.get(sz, 1) for sz in full_sizes]],
#         index=["MAX"], columns=full_sizes,
#     )
#     max_edit = st.data_editor(
#         max_df, use_container_width=True, key=f"pack_max_v6_{style_key}",
#         column_config={sz: st.column_config.NumberColumn(min_value=1, step=1) for sz in full_sizes},
#     )
#     max_by_size = {sz: int(max_edit.loc["MAX", sz]) for sz in full_sizes}

#     # ---- Packing mode ----
#     pack_mode = st.radio(
#         "Packing mode",
#         ["solid_no_leftover", "solid_all", "solid_mixed"],
#         index=2,
#         format_func=lambda k: {
#             "solid_no_leftover": "Solid, no leftover",
#             "solid_all": "Solid-all",
#             "solid_mixed": "Solid + mixed (current)",
#         }[k],
#         key=f"pack_mode_v2_{style_key}",
#         horizontal=True,
#     )

#     # ---- Build row state ----
#     rows_key = f"carton_rows_v10_{style_key}"
#     max_sig_key = f"max_sig_v10_{style_key}"
#     max_sig = tuple(sorted(max_by_size.items())) + (pack_mode,)
#     if rows_key not in st.session_state or st.session_state.get(max_sig_key) != max_sig:
#         base = compute_carton_rows_v3(sublines, max_by_size, mode=pack_mode)
#         st.session_state[rows_key] = [
#             {
#                 "id": uuid.uuid4().hex,
#                 "color": r["color"] or default_color,
#                 "upc": r["upc"] if not r["mixed"] else "MIXED",
#                 "sizes": {sz: int(r["sizes"].get(sz, 0) or 0) for sz in full_sizes},
#                 "total_ctn": int(r["total_ctn"]),
#                 "mixed": bool(r["mixed"]),
#             }
#             for r in base
#         ]
#         st.session_state[max_sig_key] = max_sig

#     rows = st.session_state[rows_key]

#     # ---- Main carton table (single native data_editor) ----
#     st.markdown("##### Main carton table (editable)")

#     def to_df(rows):
#         data = []
#         for r in rows:
#             rec = {"Delete": False, "Color": r["color"], "UPC": r["upc"]}
#             for sz in full_sizes:
#                 rec[sz] = int(r["sizes"].get(sz, 0) or 0)
#             rec["TOTAL CTN"] = int(r["total_ctn"])
#             data.append(rec)
#         return pd.DataFrame(data)

#     df = to_df(rows)
#     col_cfg = {"Delete": st.column_config.CheckboxColumn(width="small")}
#     for sz in full_sizes:
#         col_cfg[sz] = st.column_config.NumberColumn(min_value=0, step=1, format="%d")
#     col_cfg["TOTAL CTN"] = st.column_config.NumberColumn(min_value=0, step=1, format="%d")

#     edited = st.data_editor(
#         df,
#         use_container_width=True,
#         hide_index=True,
#         num_rows="dynamic",
#         key=f"grid_v10_{style_key}",
#         column_config=col_cfg,
#     )

#     new_rows = []
#     for i, (_, rec) in enumerate(edited.iterrows()):
#         old_id = rows[i]["id"] if i < len(rows) else uuid.uuid4().hex
#         new_rows.append({
#             "id": old_id,
#             "color": rec["Color"],
#             "upc": rec["UPC"],
#             "sizes": {sz: int(rec[sz] or 0) for sz in full_sizes},
#             "total_ctn": int(rec["TOTAL CTN"] or 0),
#             "mixed": str(rec["UPC"]).strip().upper() == "MIXED",
#         })

#     keep = [r for r, (_, rec) in zip(new_rows, edited.iterrows()) if not bool(rec["Delete"])]
#     if len(keep) != len(new_rows) and keep:
#         st.session_state[rows_key] = keep
#         rows = keep
#     else:
#         st.session_state[rows_key] = new_rows
#         rows = new_rows

#     # ---- Computed totals ----
#     size_packet_totals = {
#         sz: sum(int(r["sizes"].get(sz, 0)) * int(r["total_ctn"]) for r in rows)
#         for sz in full_sizes
#     }
#     total_ctn = sum(int(r["total_ctn"]) for r in rows)
#     total_pcs = sum(
#         sum(int(r["sizes"].get(sz, 0)) for sz in full_sizes) * int(r["total_ctn"])
#         for r in rows
#     )
#     grs_list = [
#         sum(int(r["sizes"].get(sz, 0)) * float(nw_by_size.get(sz, 0.0)) for sz in full_sizes)
#         + empty_ctn_wt
#         for r in rows
#     ]
#     net_list = [
#         sum(int(r["sizes"].get(sz, 0)) * float(nw_by_size.get(sz, 0.0)) for sz in full_sizes)
#         for r in rows
#     ]
#     total_grs = sum(g * int(r["total_ctn"]) for g, r in zip(grs_list, rows))
#     total_net = sum(n * int(r["total_ctn"]) for n, r in zip(net_list, rows))

#     # ---- Summary table (native st.table) ----
#     order_by = {s["size"]: s["qty"] for s in sublines}
#     ship_by = size_packet_totals
#     exs_by = {sz: ship_by.get(sz, 0) - order_by.get(sz, 0) for sz in full_sizes}
#     pct_by = {
#         sz: (exs_by[sz] / order_by[sz] * 100 if order_by.get(sz) else 0)
#         for sz in full_sizes
#     }
#     st.markdown("##### Summary")
#     st.table(pd.DataFrame({
#         "SIZE": full_sizes,
#         "ORDER QTY": [order_by.get(sz, 0) for sz in full_sizes],
#         "SHIP QTY": [ship_by.get(sz, 0) for sz in full_sizes],
#         "EXS/SHT QTY": [exs_by[sz] for sz in full_sizes],
#         "PERCENTAGE %": [round(pct_by[sz], 2) for sz in full_sizes],
#     }))

#     # ---- Live totals ----
#     m1, m2, m3, m4 = st.columns(4)
#     m1.metric("Total Cartons", total_ctn)
#     m2.metric("Total Pieces", total_pcs)
#     m3.metric("Total Gross Wt", f"{total_grs:.2f}")
#     m4.metric("Total Net Wt", f"{total_net:.2f}")

#     _exs_now = total_pcs - sum(order_by.values())
#     if _exs_now > 0:
#         st.success(f"Over-ship: **+{_exs_now}** pieces")
#     elif _exs_now < 0:
#         st.error(f"Short-ship: **{_exs_now}** pieces")
#     else:
#         st.info("Ship qty matches order qty exactly.")

#     # ---- CTN MEAS / CBM inputs ----
#     with st.expander("CTN MEAS / CBM", expanded=False):
#         d1, d2, d3 = st.columns(3)
#         ctn_l = d1.number_input("Length (in)", min_value=0.0, value=24.0, step=0.5, key=f"cl_{style_key}")
#         ctn_w = d2.number_input("Width (in)", min_value=0.0, value=16.0, step=0.5, key=f"cw_{style_key}")
#         ctn_h = d3.number_input("Height (in)", min_value=0.0, value=10.0, step=0.5, key=f"ch_{style_key}")
#         ctn_meas_1 = ""
#         ctn_meas_2 = f"L {ctn_l:g}'' X W {ctn_w:g}'' X H {ctn_h:g}''"
#         cbm_per_ctn = (ctn_l * ctn_w * ctn_h) / 61000

#     cbm_total = cbm_per_ctn * total_ctn

#     # ---- Export state ----
#     rows_export = []
#     cum = 0
#     for r, g, n in zip(rows, grs_list, net_list):
#         s_ctn = cum + 1
#         cum += int(r["total_ctn"])
#         rows_export.append({
#             "ctn_no": f"{s_ctn}-{cum}",
#             "color": r["color"],
#             "upc": r["upc"],
#             "sizes": r["sizes"],
#             "ctn_pcs": sum(r["sizes"].values()),
#             "total_ctn": int(r["total_ctn"]),
#             "grs_per_ctn": g,
#             "net_per_ctn": n,
#             "total_pcs": int(r["total_ctn"]) * sum(r["sizes"].values()),
#             "total_grs": int(r["total_ctn"]) * g,
#             "total_net": int(r["total_ctn"]) * n,
#         })

#     order_total = sum(order_by.values())
#     export_state = {
#         "line": line, "full_sizes": full_sizes, "rows": rows_export,
#         "total_ctn": total_ctn, "total_pcs": total_pcs,
#         "total_grs": total_grs, "total_net": total_net,
#         "size_packet_totals": size_packet_totals,
#         "order_qty_by_size": order_by, "ship_qty_by_size": ship_by,
#         "exs_by_size": exs_by, "pct_by_size": {sz: v / 100 for sz, v in pct_by.items()},
#         "order_total": order_total, "ship_total": total_pcs,
#         "exs_total": total_pcs - order_total,
#         "pct_total": ((total_pcs - order_total) / order_total if order_total else None),
#         "cut_qty": {sz: 0 for sz in full_sizes}, "cut_total": 0,
#         "ctn_meas_1": ctn_meas_1, "ctn_meas_2": ctn_meas_2,
#         "cbm_per_ctn": cbm_per_ctn, "cbm_total": cbm_total,
#         "nw_by_size": nw_by_size, "nnw_by_size": nnw_by_size, "empty_by_size": empty_by_size,
#     }

#     st.markdown("##### Download")
#     c1, c2 = st.columns(2)
#     with c1:
#         if st.button("Prepare Excel (.xlsx)", key=f"prep_xlsx_v10_{style_key}"):
#             wb = build_workbook({"lines": [line], "sublines": sublines})
#             b = io.BytesIO()
#             wb.save(b)
#             b.seek(0)
#             st.session_state[f"xlsx_bytes_v10_{style_key}"] = b.getvalue()
#         if st.session_state.get(f"xlsx_bytes_v10_{style_key}"):
#             st.download_button(
#                 f"Download combined (.xlsx)",
#                 data=st.session_state[f"xlsx_bytes_v10_{style_key}"],
#                 file_name=f"packing_list_combined.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                 key=f"dl_xlsx_v10_{style_key}",
#             )
#     with c2:
#         if st.button("Prepare PDF", key=f"prep_pdf_v10_{style_key}"):
#             pb = build_edited_pdf(export_state)
#             st.session_state[f"pdf_bytes_v10_{style_key}"] = pb.getvalue()
#         if st.session_state.get(f"pdf_bytes_v10_{style_key}"):
#             st.download_button(
#                 f"Download combined (.pdf)",
#                 data=st.session_state[f"pdf_bytes_v10_{style_key}"],
#                 file_name=f"packing_list_combined.pdf",
#                 mime="application/pdf",
#                 key=f"dl_pdf_v10_{style_key}",
#             )


@_fragment
def render_packing_list_preview(line, sublines):
    style_key = line["style"]
    
    # Group and merge data for the SAME size from multiple POs
    grouped = {}
    for s in sublines:
        sz = s["size"]
        if sz not in grouped:
            grouped[sz] = {"colors": [], "upcs": [], "qty": 0}
        # Append if not already present
        if s["color"] not in grouped[sz]["colors"]:
            grouped[sz]["colors"].append(s["color"])
        if s["upc"] not in grouped[sz]["upcs"]:
            grouped[sz]["upcs"].append(s["upc"])
        grouped[sz]["qty"] += s["qty"]

    full_sizes = list(grouped.keys())  # Preserves order of appearance
    default_color = next((s["color"] for s in sublines if s.get("color")), "MIXED")
    
    # Get max pack by size (using the first occurrence or fallback to 1)
    pdf_max_by_size = {}
    for s in sublines:
        sz = s["size"]
        if sz not in pdf_max_by_size:
            pdf_max_by_size[sz] = int(s.get("items_per_outer_pack") or 1)

    # ---- Weight spec table (native data_editor) ----
    st.markdown("##### Weight spec (N.W. / N.N.W. / EMPTY CTN)")
    spec_df = pd.DataFrame(
        [[1.0] * len(full_sizes) for _ in range(3)],
        index=["N.W.", "N.N.W.", "EMPTY CTN"],
        columns=full_sizes,
    )
    spec_edit = st.data_editor(
        spec_df, use_container_width=True, key=f"spec_v10_{style_key}",
    )
    nw_by_size = spec_edit.loc["N.W."].to_dict()
    nnw_by_size = spec_edit.loc["N.N.W."].to_dict()
    empty_by_size = spec_edit.loc["EMPTY CTN"].to_dict()
    empty_ctn_wt = float(next(iter(empty_by_size.values()), 0.0))

    # ---- Order quantity by size (HORIZONTAL, MERGED) ----
    st.markdown("##### Order quantity by size")

    ZWSP = "\u200b"
    order_data = {
        ZWSP: ["SIZE", "COLOR", "UPC", "ORDER QTY"],
    }
    
    for i, sz in enumerate(full_sizes):
        colors = grouped[sz]["colors"]
        upcs = grouped[sz]["upcs"]
        
        # Merge multiple colors/upcs with commas
        color_str = ", ".join(colors)
        upc_str = ", ".join(upcs)
        
        order_data[f"{ZWSP}{i + 1}"] = [sz, color_str, upc_str, grouped[sz]["qty"]]

    order_df = pd.DataFrame(order_data)
    st.dataframe(order_df, use_container_width=True, hide_index=True)

    # ---- MAX per carton (native data_editor) ----
    st.markdown("##### Packing qty (MAX per carton)")
    max_df = pd.DataFrame(
        [[pdf_max_by_size.get(sz, 1) for sz in full_sizes]],
        index=["MAX"], columns=full_sizes,
    )
    max_edit = st.data_editor(
        max_df, use_container_width=True, key=f"pack_max_v6_{style_key}",
        column_config={sz: st.column_config.NumberColumn(min_value=1, step=1) for sz in full_sizes},
    )
    max_by_size = {sz: int(max_edit.loc["MAX", sz]) for sz in full_sizes}

    # ---- Packing mode ----
    pack_mode = st.radio(
        "Packing mode",
        ["solid_no_leftover", "solid_all", "solid_mixed"],
        index=2,
        format_func=lambda k: {
            "solid_no_leftover": "Solid, no leftover",
            "solid_all": "Solid-all",
            "solid_mixed": "Solid + mixed (current)",
        }[k],
        key=f"pack_mode_v2_{style_key}",
        horizontal=True,
    )

    # ---- Build row state ----
    rows_key = f"carton_rows_v10_{style_key}"
    max_sig_key = f"max_sig_v10_{style_key}"
    max_sig = tuple(sorted(max_by_size.items())) + (pack_mode,)
    
    # Note: We pass the merged sublines (or just rebuild based on max_by_size)
    if rows_key not in st.session_state or st.session_state.get(max_sig_key) != max_sig:
        # Reconstruct simplified sublines for the packing logic
        merged_sublines = [{"size": sz, "qty": grouped[sz]["qty"], "items_per_outer_pack": pdf_max_by_size[sz], "color": grouped[sz]["colors"][0], "upc": grouped[sz]["upcs"][0]} for sz in full_sizes]
        
        base = compute_carton_rows_v3(merged_sublines, max_by_size, mode=pack_mode)
        st.session_state[rows_key] = [
            {
                "id": uuid.uuid4().hex,
                "color": r["color"] or default_color,
                "upc": r["upc"] if not r["mixed"] else "MIXED",
                "sizes": {sz: int(r["sizes"].get(sz, 0) or 0) for sz in full_sizes},
                "total_ctn": int(r["total_ctn"]),
                "mixed": bool(r["mixed"]),
            }
            for r in base
        ]
        st.session_state[max_sig_key] = max_sig

    rows = st.session_state[rows_key]

    # ---- Main carton table (single native data_editor) ----
    st.markdown("##### Main carton table (editable)")

    def to_df(rows):
        data = []
        for r in rows:
            rec = {"Delete": False, "Color": r["color"], "UPC": r["upc"]}
            for sz in full_sizes:
                rec[sz] = int(r["sizes"].get(sz, 0) or 0)
            rec["TOTAL CTN"] = int(r["total_ctn"])
            data.append(rec)
        return pd.DataFrame(data)

    df = to_df(rows)
    col_cfg = {"Delete": st.column_config.CheckboxColumn(width="small")}
    for sz in full_sizes:
        col_cfg[sz] = st.column_config.NumberColumn(min_value=0, step=1, format="%d")
    col_cfg["TOTAL CTN"] = st.column_config.NumberColumn(min_value=0, step=1, format="%d")

    edited = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key=f"grid_v10_{style_key}",
        column_config=col_cfg,
    )

    new_rows = []
    for i, (_, rec) in enumerate(edited.iterrows()):
        old_id = rows[i]["id"] if i < len(rows) else uuid.uuid4().hex
        new_rows.append({
            "id": old_id,
            "color": rec["Color"],
            "upc": rec["UPC"],
            "sizes": {sz: int(rec[sz] or 0) for sz in full_sizes},
            "total_ctn": int(rec["TOTAL CTN"] or 0),
            "mixed": str(rec["UPC"]).strip().upper() == "MIXED",
        })

    keep = [r for r, (_, rec) in zip(new_rows, edited.iterrows()) if not bool(rec["Delete"])]
    if len(keep) != len(new_rows) and keep:
        st.session_state[rows_key] = keep
        rows = keep
    else:
        st.session_state[rows_key] = new_rows
        rows = new_rows

    # ---- Computed totals ----
    size_packet_totals = {
        sz: sum(int(r["sizes"].get(sz, 0)) * int(r["total_ctn"]) for r in rows)
        for sz in full_sizes
    }
    total_ctn = sum(int(r["total_ctn"]) for r in rows)
    total_pcs = sum(
        sum(int(r["sizes"].get(sz, 0)) for sz in full_sizes) * int(r["total_ctn"])
        for r in rows
    )
    grs_list = [
        sum(int(r["sizes"].get(sz, 0)) * float(nw_by_size.get(sz, 0.0)) for sz in full_sizes)
        + empty_ctn_wt
        for r in rows
    ]
    net_list = [
        sum(int(r["sizes"].get(sz, 0)) * float(nw_by_size.get(sz, 0.0)) for sz in full_sizes)
        for r in rows
    ]
    total_grs = sum(g * int(r["total_ctn"]) for g, r in zip(grs_list, rows))
    total_net = sum(n * int(r["total_ctn"]) for n, r in zip(net_list, rows))

    # ---- Summary table (native st.table) ----
    order_by = {sz: grouped[sz]["qty"] for sz in full_sizes}
    ship_by = size_packet_totals
    exs_by = {sz: ship_by.get(sz, 0) - order_by.get(sz, 0) for sz in full_sizes}
    pct_by = {
        sz: (exs_by[sz] / order_by[sz] * 100 if order_by.get(sz) else 0)
        for sz in full_sizes
    }
    st.markdown("##### Summary")
    st.table(pd.DataFrame({
        "SIZE": full_sizes,
        "ORDER QTY": [order_by.get(sz, 0) for sz in full_sizes],
        "SHIP QTY": [ship_by.get(sz, 0) for sz in full_sizes],
        "EXS/SHT QTY": [exs_by[sz] for sz in full_sizes],
        "PERCENTAGE %": [round(pct_by[sz], 2) for sz in full_sizes],
    }))

    # ---- Live totals ----
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Cartons", total_ctn)
    m2.metric("Total Pieces", total_pcs)
    m3.metric("Total Gross Wt", f"{total_grs:.2f}")
    m4.metric("Total Net Wt", f"{total_net:.2f}")

    _exs_now = total_pcs - sum(order_by.values())
    if _exs_now > 0:
        st.success(f"Over-ship: **+{_exs_now}** pieces")
    elif _exs_now < 0:
        st.error(f"Short-ship: **{_exs_now}** pieces")
    else:
        st.info("Ship qty matches order qty exactly.")

    # ---- CTN MEAS / CBM inputs ----
    with st.expander("CTN MEAS / CBM", expanded=False):
        d1, d2, d3 = st.columns(3)
        ctn_l = d1.number_input("Length (in)", min_value=0.0, value=24.0, step=0.5, key=f"cl_{style_key}")
        ctn_w = d2.number_input("Width (in)", min_value=0.0, value=16.0, step=0.5, key=f"cw_{style_key}")
        ctn_h = d3.number_input("Height (in)", min_value=0.0, value=10.0, step=0.5, key=f"ch_{style_key}")
        ctn_meas_1 = ""
        ctn_meas_2 = f"L {ctn_l:g}'' X W {ctn_w:g}'' X H {ctn_h:g}''"
        cbm_per_ctn = (ctn_l * ctn_w * ctn_h) / 61000

    cbm_total = cbm_per_ctn * total_ctn

    # ---- Export state ----
    rows_export = []
    cum = 0
    for r, g, n in zip(rows, grs_list, net_list):
        s_ctn = cum + 1
        cum += int(r["total_ctn"])
        rows_export.append({
            "ctn_no": f"{s_ctn}-{cum}",
            "color": r["color"],
            "upc": r["upc"],
            "sizes": r["sizes"],
            "ctn_pcs": sum(r["sizes"].values()),
            "total_ctn": int(r["total_ctn"]),
            "grs_per_ctn": g,
            "net_per_ctn": n,
            "total_pcs": int(r["total_ctn"]) * sum(r["sizes"].values()),
            "total_grs": int(r["total_ctn"]) * g,
            "total_net": int(r["total_ctn"]) * n,
        })

    order_total = sum(order_by.values())
    export_state = {
        "line": line, "full_sizes": full_sizes, "rows": rows_export,
        "total_ctn": total_ctn, "total_pcs": total_pcs,
        "total_grs": total_grs, "total_net": total_net,
        "size_packet_totals": size_packet_totals,
        "order_qty_by_size": order_by, "ship_qty_by_size": ship_by,
        "exs_by_size": exs_by, "pct_by_size": {sz: v / 100 for sz, v in pct_by.items()},
        "order_total": order_total, "ship_total": total_pcs,
        "exs_total": total_pcs - order_total,
        "pct_total": ((total_pcs - order_total) / order_total if order_total else None),
        "cut_qty": {sz: 0 for sz in full_sizes}, "cut_total": 0,
        "ctn_meas_1": ctn_meas_1, "ctn_meas_2": ctn_meas_2,
        "cbm_per_ctn": cbm_per_ctn, "cbm_total": cbm_total,
        "nw_by_size": nw_by_size, "nnw_by_size": nnw_by_size, "empty_by_size": empty_by_size,
    }

    # st.markdown("##### Download")
    # c1, c2 = st.columns(2)
    # with c1:
    #     if st.button("Prepare Excel (.xlsx)", key=f"prep_xlsx_v10_{style_key}"):
    #         wb = build_workbook({"lines": [line], "sublines": sublines})
    #         b = io.BytesIO()
    #         wb.save(b)
    #         b.seek(0)
    #         st.session_state[f"xlsx_bytes_v10_{style_key}"] = b.getvalue()
    #     if st.session_state.get(f"xlsx_bytes_v10_{style_key}"):
    #         st.download_button(
    #             f"Download combined (.xlsx)",
    #             data=st.session_state[f"xlsx_bytes_v10_{style_key}"],
    #             file_name=f"packing_list_combined.xlsx",
    #             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #             key=f"dl_xlsx_v10_{style_key}",
    #         )
    # with c2:
    #     if st.button("Prepare PDF", key=f"prep_pdf_v10_{style_key}"):
    #         pb = build_edited_pdf(export_state)
    #         st.session_state[f"pdf_bytes_v10_{style_key}"] = pb.getvalue()
    #     if st.session_state.get(f"pdf_bytes_v10_{style_key}"):
    #         st.download_button(
    #             f"Download combined (.pdf)",
    #             data=st.session_state[f"pdf_bytes_v10_{style_key}"],
    #             file_name=f"packing_list_combined.pdf",
    #             mime="application/pdf",
    #             key=f"dl_pdf_v10_{style_key}",
    #         )
    # st.markdown("##### Download")
    # c1, c2 = st.columns(2)
    # with c1:
    #     if st.button("Prepare Excel (.xlsx)", key=f"prep_xlsx_v10_{style_key}"):
    #         # ✅ NEW: Create a single workbook and a single sheet manually
    #         wb = Workbook()
    #         ws = wb.active
    #         ws.title = "Combined_Packing"  # Give it a valid name
            
    #         # Call build_sheet directly with the combined line and sublines
    #         build_sheet(ws, line, sublines)
            
    #         b = io.BytesIO()
    #         wb.save(b)
    #         b.seek(0)
    #         st.session_state[f"xlsx_bytes_v10_{style_key}"] = b.getvalue()
    #     if st.session_state.get(f"xlsx_bytes_v10_{style_key}"):
    #         st.download_button(
    #             f"Download combined (.xlsx)",
    #             data=st.session_state[f"xlsx_bytes_v10_{style_key}"],
    #             file_name=f"packing_list_combined.xlsx",
    #             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #             key=f"dl_xlsx_v10_{style_key}",
    #         )
    # with c2:
    #     if st.button("Prepare PDF", key=f"prep_pdf_v10_{style_key}"):
    #         pb = build_edited_pdf(export_state)
    #         st.session_state[f"pdf_bytes_v10_{style_key}"] = pb.getvalue()
    #     if st.session_state.get(f"pdf_bytes_v10_{style_key}"):
    #         st.download_button(
    #             f"Download combined (.pdf)",
    #             data=st.session_state[f"pdf_bytes_v10_{style_key}"],
    #             file_name=f"packing_list_combined.pdf",
    #             mime="application/pdf",
    #             key=f"dl_pdf_v10_{style_key}",
    #         )


def run(file=None):
    if file is None:
        st.title("PO to Packing List Converter")
        st.caption("Upload a VF/Vans Purchase Order PDF, pick one or more PO lines, edit the combined carton table, "
                   "then download the Excel or PDF.")
    uploaded_file = file if file is not None else st.file_uploader("Upload PO PDF", type=["pdf"])
    if uploaded_file is not None:
        with st.spinner("Reading PDF..."):
            try:
                data = extract_data_cached(uploaded_file.getvalue())
            except Exception as e:
                st.error(f"Failed to read the PDF: {e}")
                st.stop()
    
        if not data["lines"]:
            st.warning("No PO lines were found. This PDF's layout may differ from the expected format.")
            st.stop()
    
        _upload_sig = f"{uploaded_file.name}::{uploaded_file.size}"
        if st.session_state.get("_last_upload_sig") != _upload_sig:
            for _k in list(st.session_state.keys()):
                if _k.startswith(("carton_rows_v10_", "max_sig_v10_",
                                  "pack_max_v6_", "pack_mode_v2_",
                                  "xlsx_bytes_v10_", "pdf_bytes_v10_",
                                  "spec_v10_", "grid_v10_",
                                  "cl_", "cw_", "ch_")):
                    st.session_state.pop(_k, None)
            st.session_state.pop("po_lines_table", None)
            st.session_state.pop("preview_styles", None)
            st.session_state["_last_upload_sig"] = _upload_sig
    
        # ---- 1 summary ----
        st.subheader("1. PO summary")
        _tot_qty = sum(int(l.get("order_qty") or 0) for l in data["lines"])
        c1, c2, c3 = st.columns(3)
        c1.metric("PO lines", len(data["lines"]))
        c2.metric("Size / colour sublines", len(data["sublines"]))
        c3.metric("Order qty", f"{_tot_qty:,}")
    
        # ---- 2 choose a PO line (MULTI-SELECT) ----
        st.subheader("2. Choose PO line(s)")
        st.caption("Click the checkbox to select one or more PO lines. Their items will be combined into one single table below.")
    
        po_line_df = pd.DataFrame([
            {
                "PO Line No.": l["po_line_no"],
                "Style": l["style"],
                "Description": l["description"],
                "Order Qty": l["order_qty"],
                "Unit Price": l["unit_price"],
                "Line Cost": l["line_cost"],
                "CRD": l.get("crd"),
                "Country": l.get("destination_country"),
            }
            for l in data["lines"]
        ])
        
        selection = st.dataframe(
            po_line_df,
            use_container_width=True,
            on_select="rerun",
            selection_mode="multi-row", 
            key="po_lines_table",
        )
    
        selected_idxs = []
        if selection is not None and hasattr(selection, "selection"):
            selected_idxs = selection.selection.rows
    
        selected_lines = [data["lines"][i] for i in selected_idxs if i < len(data["lines"])]
        st.session_state["preview_styles"] = [l["style"] for l in selected_lines]
    
        if selected_lines:
            st.info(f"Selected **{len(selected_lines)}** PO line(s): {', '.join(l['style'] for l in selected_lines)}")
        else:
            st.caption("No PO lines selected yet.")
    
        # ---- Sublines expander ----
        with st.expander("Sublines (size / colour / UPC)", expanded=False):
            if selected_lines:
                filtered_sublines = [
                    s for s in data["sublines"] 
                    if s["style"] in st.session_state["preview_styles"]
                ]
                st.dataframe(pd.DataFrame(filtered_sublines), use_container_width=True)
            else:
                st.dataframe(pd.DataFrame(data["sublines"]), use_container_width=True)
    
        # ---- 3 single combined carton table ----
        if selected_lines:
            st.subheader("3. Carton table (combined)")
            
            # Create a dummy "line" object so the table renders once
            combined_line = {
                "style": " + ".join([l["style"] for l in selected_lines]),
                "po_line_no": " | ".join([l["po_line_no"] for l in selected_lines]),
                "description": " | ".join([l["description"] for l in selected_lines]),
                "order_qty": sum([l["order_qty"] for l in selected_lines]),
                "crd": selected_lines[0].get("crd"),
                "destination_country": selected_lines[0].get("destination_country"),
            }
            
            # Combine sublines from all selected POs
            combined_sublines = [
                s for s in data["sublines"] 
                if s["style"] in st.session_state["preview_styles"]
            ]
            
            # Render once
            render_packing_list_preview(combined_line, combined_sublines)
        else:
            st.subheader("3. Carton table")
            st.caption("Select one or more PO lines above to see and edit the combined carton table.")
     
        # ---- 4 download ----
        st.subheader("4. Download")
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            # ✅ FIX: Return the buffer object itself, not bytes
            def get_full_wb():
                wb = build_workbook(data)
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                return buf  
                
            st.download_button(
                "⬇ Download all PO lines (Excel)",
                data=get_full_wb(),
                file_name="packing_list_all_PO_lines.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_full_xlsx",
            )
                
        with col2:
            if selected_lines:
                # ✅ FIX: Return the buffer object itself
                def get_selected_wb():
                    selected_data = {
                        "lines": selected_lines,
                        "sublines": [s for s in data["sublines"] if s["style"] in st.session_state["preview_styles"]],
                    }
                    wb = build_workbook(selected_data)
                    buf = io.BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    return buf
                    
                st.download_button(
                    "⬇ Download selected PO line(s) (Excel)",
                    data=get_selected_wb(),
                    file_name="packing_list_selected.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_selected_xlsx",
                )
            else:
                st.caption("Select PO line(s) above to download just those sheets.")
                
        with col3:
            if selected_lines:
                combined_style_key = " + ".join([l["style"] for l in selected_lines])
                
                # ✅ FIX: Return the buffer object itself
                def get_combined_xlsx():
                    combined_line = {
                        "style": combined_style_key,
                        "po_line_no": " | ".join([l["po_line_no"] for l in selected_lines]),
                        "description": " | ".join([l["description"] for l in selected_lines]),
                        "order_qty": sum([l["order_qty"] for l in selected_lines]),
                        "crd": selected_lines[0].get("crd"),
                        "destination_country": selected_lines[0].get("destination_country"),
                    }
                    combined_sublines = [
                        s for s in data["sublines"] 
                        if s["style"] in st.session_state["preview_styles"]
                    ]
                    
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Combined_Packing"
                    build_sheet(ws, combined_line, combined_sublines)
                    
                    b = io.BytesIO()
                    wb.save(b)
                    b.seek(0)
                    return b
                    
                st.download_button(
                    "⬇ Download combined (.xlsx)",
                    data=get_combined_xlsx(),
                    file_name="packing_list_combined.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_xlsx_v10_{combined_style_key}",
                )
            else:
                st.caption("Select PO line(s) to prepare combined Excel.")
                
        with col4:
            if selected_lines:
                combined_style_key = " + ".join([l["style"] for l in selected_lines])
                
                # ✅ FIX: Return the buffer object itself
                def get_combined_pdf():
                    combined_line = {
                        "style": combined_style_key,
                        "po_line_no": " | ".join([l["po_line_no"] for l in selected_lines]),
                        "description": " | ".join([l["description"] for l in selected_lines]),
                        "order_qty": sum([l["order_qty"] for l in selected_lines]),
                        "crd": selected_lines[0].get("crd"),
                        "destination_country": selected_lines[0].get("destination_country"),
                    }
                    
                    export_state = {
                        "line": combined_line, 
                        "full_sizes": [s["size"] for s in data["sublines"] if s["style"] in st.session_state["preview_styles"]], 
                        "rows": [],
                        "total_ctn": 0, "total_pcs": 0, "total_grs": 0, "total_net": 0,
                        "size_packet_totals": {}, "order_qty_by_size": {}, "ship_qty_by_size": {},
                        "exs_by_size": {}, "pct_by_size": {}, "cut_qty": {}, "cut_total": 0,
                        "order_total": 0, "ship_total": 0, "exs_total": 0, "pct_total": 0,
                        "ctn_meas_1": "", "ctn_meas_2": "", "cbm_per_ctn": 0, "cbm_total": 0,
                        "nw_by_size": {}, "nnw_by_size": {}, "empty_by_size": {}
                    }
                    
                    pb = build_edited_pdf(export_state)
                    pb.seek(0)
                    return pb
                    
                st.download_button(
                    "⬇ Download combined (.pdf)",
                    data=get_combined_pdf(),
                    file_name="packing_list_combined.pdf",
                    mime="application/pdf",
                    key=f"dl_pdf_v10_{combined_style_key}",
                )
            else:
                st.caption("Select PO line(s) to prepare combined PDF.")
    
        with col5:
            st.download_button(
                label="⬇ Download extracted JSON",
                data=json.dumps(data, indent=2).encode("utf-8"),
                file_name=f"po_extracted_{len(data['lines'])}_lines.json",
                mime="application/json",
                key="dl_po_json",
            )
    else:
        st.info("Upload a PDF to get started.")
        st.stop()

if __name__ == "__main__":
    st.set_page_config(page_title="VF PO -> Packing List", layout="wide", page_icon="📦")
    run()
