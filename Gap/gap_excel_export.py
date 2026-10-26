"""
gap_excel_export.py
-------------------
ALL Excel (xlsx) generation for the GAP / Old Navy packing list, in one file.
No Streamlit / UI code here - safe to import from the app or from a plain script.

Two exporters, producing two SEPARATE workbooks:

  1. export_unified_excel(...)   ->  one sheet "Packing List" for ALL pack types
                                     (Bulk / Single / Multi Y / Multi N) with the
                                     N.W / N.N.W table, carton-tare table, colour
                                     summary and G. Total Summery.
  2. export_multi_n_excel(...)   ->  sheets "PKL" + "sum" for MULTI PACK ( N ) only
                                     (mirrors the reference "... Multi.xlsx").

All derived cells are live Excel formulas.

Shared defaults (weights, carton measures, sort order) are imported from
color_builders.py / multi_n_core.py so there is a single source of truth.
"""

from __future__ import annotations

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties

from color_builders import DEFAULT_CTN_MEAS, DEFAULT_NW, NET_DEDUCT, sort_sizes
from multi_n_core import DEFAULT_NNW, DEFAULT_TARE, size_family

__all__ = ["export_unified_excel", "export_multi_n_excel"]

FONT = "Courier New"


# ===========================================================================
# 1. Main workbook - all pack types on one sheet
# ===========================================================================
def export_unified_excel(
    data, bulk_secs, single_secs, multi_y_secs, multi_n_secs,
    nw: dict | None = None,
    nnw: dict | None = None,
    tare_by_code: dict | None = None,
    net_deduct: dict | None = None,
    order: list | None = None,
) -> bytes:
    """
    SINGLE-SHEET packing list for ALL GAP types (Bulk / Single / Multi Y / Multi N).

    Layout (top → bottom on one sheet "Packing List"):
      1. Company header
      2. Table A — Net weight by SIZE (N.W / N.N.W) — editable
      3. Table B — Empty carton weight by CTN CODE — editable (separate table)
      4. Each packing type section (with SL# serial numbers per section)
      5. Color-wise SUMMARY + G. Total Summery

    All derived cells are live Excel formulas.
    """
    nw = dict(nw or DEFAULT_NW)
    nnw = dict(nnw or DEFAULT_NNW)
    tare_by_code = dict(tare_by_code or DEFAULT_TARE)
    deduct_by_pack = {**NET_DEDUCT, **(net_deduct or {})}   # {"Bulk": .24, "Multi N": .42, ...}

    # Column widths by header name - edit here to adjust any column.
    # First columns (A-E) are widened so full headers / left labels show without cut-off.
    adjusted_widths = {
        "SL#": 11, "CTN FROM": 10, "CTN TO": 9, "CTN MES.": 10, "CTN QTY": 9,
        "COLOR": 14, "PREPACK STECKER": 12,
        "QTY PER CTN": 10, "TOTAL QTY": 10, "GROSS WEIGHT": 11, "NET WEIGHT": 11,
    }
    DEFAULT_COL_WIDTH = 7        # size columns and anything not listed above

    wb = Workbook()
    ws = wb.active
    ws.title = "Packing List"

    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF", name="Courier New", size=10)
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    yellow = PatternFill("solid", fgColor="FFFF00")
    summary_fill = PatternFill("solid", fgColor="DDEBF7")
    nw_fill = PatternFill("solid", fgColor="E2EFDA")
    type_strip_fill = PatternFill("solid", fgColor="FCE4D6")
    font_bold = Font(bold=True, name="Courier New", size=11)
    font_norm = Font(name="Courier New", size=10)
    font_title = Font(bold=True, name="Courier New", size=16, underline="single")
    font_addr = Font(bold=True, name="Courier New", size=12)
    font_sum_hdr = Font(bold=True, name="Courier New", size=12)
    ctr = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    header = data.get("PO_HEADER") or {}
    buyer = header.get("BUYER") or "OLD NAVY"
    style = header.get("STYLE") or ""
    po_no = str(header.get("P.O._#") or header.get("P.O. #") or "")
    oiqty = header.get("OIQTY") or 0

    # Section order on the sheet (default Bulk, Single, Multi Y, Multi N)
    _by_type = {"Bulk": list(bulk_secs), "Single": list(single_secs),
                "Multi Y": list(multi_y_secs), "Multi N": list(multi_n_secs)}
    _order = [t for t in (order or ["Bulk", "Single", "Multi Y", "Multi N"]) if t in _by_type]
    _order += [t for t in _by_type if t not in _order]
    all_secs = [sec for t in _order for sec in _by_type[t]]
    grand_ctn = sum(s.get("ctn_total", 0) for s in all_secs)

    # Collect all sizes across every section (for NW table + consistent columns)
    size_set = []
    for sec in all_secs:
        for s in (sec.get("sizes") or []):
            if s not in size_set:
                size_set.append(s)
    size_set = sort_sizes(size_set) if size_set else ["XS", "S", "M", "L", "XL", "XXL", "XXL+"]

    # Column layout matching reference packing list:
    # SL# | CTN NO (from) | CTN NO (to) | CTN MES. | CTN QTY | COLOR | PREPACK STECKER
    # | size cols | QTY PER CTN | TOTAL QTY | GROSS WEIGHT | NET WEIGHT
    base_names = ["SL#", "CTN FROM", "CTN TO", "CTN MES.", "CTN QTY", "COLOR", "PREPACK STECKER"]
    derived_names = ["QTY PER CTN", "TOTAL QTY", "GROSS WEIGHT", "NET WEIGHT"]
    all_col_names = base_names + size_set + derived_names
    n_cols = len(all_col_names)
    col_idx = {name: i + 1 for i, name in enumerate(all_col_names)}

    # Set column widths
    for name, c in col_idx.items():
        ws.column_dimensions[L(c)].width = adjusted_widths.get(name, DEFAULT_COL_WIDTH)

    r = 1  # current row

    # ========== 1. COMPANY HEADER ==========
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    cell = ws.cell(r, 1, "CREATIVE COLLECTIONS LTD")
    cell.font = font_title
    cell.alignment = ctr
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    cell = ws.cell(r, 1, "Nishat Nagar , Tongi , Gazipur .")
    cell.font = font_addr
    cell.alignment = ctr
    r += 2

    # ---- helpers: centre each reference table under the centred title ----
    LABEL_SPAN = 3                       # label cell = 3 merged columns (fits "N.N.W / N.N.WT")
    col_w = {c: float(ws.column_dimensions[L(c)].width or 8.43) for c in range(1, n_cols + 1)}

    def _centered_start(n_value_cols: int) -> int:
        """First column so that [label + values] sits in the middle of the sheet width."""
        span = LABEL_SPAN + n_value_cols
        if span >= n_cols:
            return 1
        best_c, best_gap = 1, None
        for c in range(1, n_cols - span + 2):
            left_w = sum(col_w[x] for x in range(1, c))
            right_w = sum(col_w[x] for x in range(c + span, n_cols + 1))
            gap = abs(left_w - right_w)
            if best_gap is None or gap < best_gap:
                best_c, best_gap = c, gap
        return best_c

    def _label(row, c0, text, header=False):
        """Merged, centred, bordered label cell (3 columns wide)."""
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c0 + LABEL_SPAN - 1)
        for cc in range(c0, c0 + LABEL_SPAN):
            cell = ws.cell(row, cc)
            cell.border = thin
            cell.alignment = ctr
            if header:
                cell.fill = header_fill
        cell = ws.cell(row, c0, text)
        cell.font = header_font if header else font_bold

    # ========== 2a. TABLE A — Net weight by SIZE (editable, centred) ==========
    #   SIZE | XS | S | M | L | XL | XXL | XXL+
    #   N.W  | values...
    #   N.N.W| values...
    nw_label_col = _centered_start(len(size_set))
    nw_start_col = nw_label_col + LABEL_SPAN      # first size column
    nw_header_row = r
    nw_row = r + 1
    nnw_row = r + 2

    _label(nw_header_row, nw_label_col, "SIZE", header=True)
    _label(nw_row, nw_label_col, "N.W / N.WT")
    _label(nnw_row, nw_label_col, "N.N.W / N.N.WT")

    for i, s in enumerate(size_set):
        c = nw_start_col + i
        cell = ws.cell(nw_header_row, c, s)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = ctr
        cell.border = thin

        cell = ws.cell(nw_row, c, float(nw.get(s, 0.5)))
        cell.number_format = "0.000"
        cell.fill = yellow
        cell.alignment = ctr
        cell.border = thin

        cell = ws.cell(nnw_row, c, float(nnw.get(s, round(float(nw.get(s, 0.5)) - 0.02, 3))))
        cell.number_format = "0.000"
        cell.fill = yellow
        cell.alignment = ctr
        cell.border = thin

    nw_cell_ref = {s: f"${L(nw_start_col + i)}${nw_row}" for i, s in enumerate(size_set)}
    r = nnw_row + 2

    # ========== 2b. TABLE B — Empty carton weight by CTN CODE (separate, editable, centred) ==========
    tare_codes = list(tare_by_code.keys()) or ["G81", "G82", "G84", "G85"]
    tare_label_col = _centered_start(len(tare_codes))     # centred on its own
    tare_start_col = tare_label_col + LABEL_SPAN
    tare_header_row = r
    tare_val_row = r + 1

    _label(tare_header_row, tare_label_col, "CTN CODE", header=True)
    _label(tare_val_row, tare_label_col, "EMPTY CTN WT")

    tare_cell_ref = {}
    for i, code in enumerate(tare_codes):
        c = tare_start_col + i
        cell = ws.cell(tare_header_row, c, code)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = ctr
        cell.border = thin

        cell = ws.cell(tare_val_row, c, float(tare_by_code.get(code, 0.71)))
        cell.number_format = "0.00"
        cell.fill = yellow
        cell.alignment = ctr
        cell.border = thin
        tare_cell_ref[code] = f"${L(c)}${tare_val_row}"

    tare_hdr_rng = f"${L(tare_start_col)}${tare_header_row}:${L(tare_start_col + len(tare_codes) - 1)}${tare_header_row}"
    tare_val_rng = f"${L(tare_start_col)}${tare_val_row}:${L(tare_start_col + len(tare_codes) - 1)}${tare_val_row}"
    r = tare_val_row + 2

    # ========== 2c. TABLE C — Net deduct by PACK TYPE (editable, centred) ==========
    #   Net weight = Gross weight - NET DEDUCT
    pack_names = [("Bulk", "BULK"), ("Single", "SINGLE"), ("Multi Y", "MULTI Y"), ("Multi N", "MULTI N")]
    ded_label_col = _centered_start(len(pack_names))
    ded_start_col = ded_label_col + LABEL_SPAN
    ded_header_row, ded_val_row = r, r + 1
    _label(ded_header_row, ded_label_col, "PACK TYPE", header=True)
    _label(ded_val_row, ded_label_col, "NET DEDUCT")
    deduct_ref = {}
    for i, (pk, nm) in enumerate(pack_names):
        c = ded_start_col + i
        cell = ws.cell(ded_header_row, c, nm)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        cell.border = thin
        cell = ws.cell(ded_val_row, c, float(deduct_by_pack.get(pk, 0.24)))
        cell.number_format = "0.00"
        cell.fill = yellow
        cell.alignment = ctr
        cell.border = thin
        deduct_ref[pk] = f"${L(c)}${ded_val_row}"
    r = ded_val_row + 2

    # NW range laid out exactly like the size columns of the packing table
    nw_rng = f"${L(nw_start_col)}${nw_row}:${L(nw_start_col + len(size_set) - 1)}${nw_row}"

    # ========== 3. PACKING SECTIONS ==========
    section_summary_refs = []  # for grand summary

    def _write_one_section(sec, start_r, serial_start=1):
        """Write one colour section. Returns (next_row, summary_ref_dict)."""
        pack = sec["pack_type"]
        sizes = list(sec.get("sizes") or [])
        rows_data = sec.get("rows") or []
        nonlocal_r = start_r

        # --- type strip ---
        label = {
            "Bulk": "BULK PACK",
            "Single": "SINGLE PACK",
            "Multi Y": "MULTY PACK ( Y )",
            "Multi N": "MULTY PACK ( N )",
        }.get(pack, pack.upper())
        ws.merge_cells(start_row=nonlocal_r, start_column=1,
                       end_row=nonlocal_r, end_column=n_cols)
        cell = ws.cell(nonlocal_r, 1, label)
        cell.fill = yellow
        cell.font = Font(bold=True, name="Courier New", size=12)
        cell.alignment = ctr
        nonlocal_r += 1

        # --- PO header block (left) ---
        hdr0 = nonlocal_r
        left_labels = ["BUYER", "STYLE", "P.O. #", "O/QTY", "SHIP QTY", "EX/SHORT", "CTN QTY"]
        for i, lab in enumerate(left_labels):
            ws.cell(hdr0 + i, 1, lab).font = font_bold
            ws.cell(hdr0 + i, 2, ":")
        ws.cell(hdr0 + 0, 3, buyer)
        ws.cell(hdr0 + 1, 3, sec.get("style_full") or style)
        ws.cell(hdr0 + 2, 3, po_no)
        oq_cell = ws.cell(hdr0 + 3, 3, sec.get("qty_ordered", 0))
        oq_cell.number_format = '#,##0 "PCS"'
        ship_cell = ws.cell(hdr0 + 4, 3)   # formula later
        ship_cell.number_format = '#,##0 "PCS"'
        ex_cell = ws.cell(hdr0 + 5, 3)
        ex_cell.number_format = '#,##0 "PCS"'
        ctn_hdr_cell = ws.cell(hdr0 + 6, 3)
        ctn_hdr_cell.number_format = '#,##0 "CTNS"'

        meas_list = sec.get("ctn_meas_list") or list(DEFAULT_CTN_MEAS)
        for i, m in enumerate(meas_list):
            ws.cell(hdr0 + 7 + i, 1, "CTN MEAS.").font = font_bold
            ws.cell(hdr0 + 7 + i, 2, ":")
            ws.cell(hdr0 + 7 + i, 3, m)

        # ----- RIGHT TOP block (exactly top-right of this section) -----
        # Labels + values aligned with BUYER / STYLE / P.O. / O/QTY rows
        # right edge of the block = right edge of the main table (NET WEIGHT column)
        rc_lab = col_idx["GROSS WEIGHT"]   # label column  (second-last table column)
        rc_val = col_idx["NET WEIGHT"]     # value column  (last table column)
        ctn_from0 = rows_data[0].get("ctn_from", 1) if rows_data else 1
        right_items = [
            ("PO #", po_no),
            ("SKU / Item", sec.get("style_full") or style),
            ("Unit / Prepack", str(sec.get("units_per_prepack", 1))),
            ("CARTON", f"{int(ctn_from0):02d} of {grand_ctn}"),
        ]
        left_fit = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
        for i, (lab, val) in enumerate(right_items):
            ws.cell(hdr0 + i, rc_lab, lab).font = font_bold
            ws.cell(hdr0 + i, rc_lab).alignment = left_fit      # long labels shrink to fit the cell
            cell = ws.cell(hdr0 + i, rc_val, val)
            cell.alignment = left_fit
            cell.font = font_norm
            cell.border = Border(bottom=Side(style="thin"))

        nonlocal_r = hdr0 + 7 + len(meas_list) + 1

        # --- column headers ---
        header_row = nonlocal_r
        for c, name in enumerate(all_col_names, 1):
            cell = ws.cell(nonlocal_r, c, name)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin
            cell.alignment = ctr
        nonlocal_r += 1

        first_data = nonlocal_r
        size_letters = [L(col_idx[s]) for s in size_set if s in col_idx]

        for k, rowd in enumerate(rows_data):
            sl = serial_start + k
            # SL#
            ws.cell(nonlocal_r, col_idx["SL#"], sl).alignment = ctr
            # CTN FROM / CTN TO (two columns, matching reference)
            ctn_from_val = rowd.get("ctn_from", k + 1)
            ctn_to_val = rowd.get("ctn_to") or ctn_from_val
            ws.cell(nonlocal_r, col_idx["CTN FROM"], ctn_from_val).alignment = ctr
            ws.cell(nonlocal_r, col_idx["CTN TO"], ctn_to_val).alignment = ctr
            # CTN MES. — prefer short code (G81); extract if full dimension given
            mes = rowd.get("ctn_mes", "") or ""
            if " " in str(mes):
                parts = str(mes).replace(".", " ").split()
                code_cand = [p for p in parts if p.upper().startswith("G") and any(ch.isdigit() for ch in p)]
                if code_cand:
                    mes = code_cand[-1]
            ws.cell(nonlocal_r, col_idx["CTN MES."], mes).alignment = ctr
            # CTN QTY (editable input)
            cell = ws.cell(nonlocal_r, col_idx["CTN QTY"], rowd.get("ctn_qty", 0))
            cell.alignment = ctr
            cell.number_format = "0"
            cell.fill = yellow
            # COLOR / PREPACK STECKER
            ws.cell(nonlocal_r, col_idx["COLOR"],
                    rowd.get("color") or sec.get("color", "")).alignment = ctr
            ws.cell(nonlocal_r, col_idx["PREPACK STECKER"],
                    rowd.get("prepack") or sec.get("prepack", "")).alignment = ctr

            # size units (editable)
            if pack in ("Single", "Bulk"):
                for s in size_set:
                    val = rowd.get("size_val") if s == rowd.get("size") else None
                    cell = ws.cell(nonlocal_r, col_idx[s], val if val not in ("", None) else None)
                    cell.alignment = ctr
                    if val not in ("", None):
                        cell.number_format = "0"
                        cell.fill = yellow
            else:
                size_vals = rowd.get("size_vals") or sec.get("units") or {}
                for s in size_set:
                    val = size_vals.get(s)
                    cell = ws.cell(nonlocal_r, col_idx[s], val if val not in ("", None) else None)
                    cell.alignment = ctr
                    if val not in ("", None):
                        cell.number_format = "0"
                        cell.fill = yellow

            # ---- FORMULAS ----
            # QTY PER CTN = sum of size columns (× prepacks for Multi when needed)
            ppc = rowd.get("prepacks_per_ctn") or sec.get("per_ctn") or 1
            if size_letters:
                if pack not in ("Bulk", "Single") and isinstance(ppc, (int, float)) and ppc != 1:
                    qty_per_f = f"=SUM({size_letters[0]}{nonlocal_r}:{size_letters[-1]}{nonlocal_r})*{ppc}"
                else:
                    qty_per_f = f"=SUM({size_letters[0]}{nonlocal_r}:{size_letters[-1]}{nonlocal_r})"
            else:
                qty_per_f = rowd.get("qty_per_ctn") or 0
            ws.cell(nonlocal_r, col_idx["QTY PER CTN"], qty_per_f).alignment = ctr

            # TOTAL QTY = QTY PER CTN * CTN QTY
            ws.cell(nonlocal_r, col_idx["TOTAL QTY"],
                    f"={L(col_idx['QTY PER CTN'])}{nonlocal_r}*{L(col_idx['CTN QTY'])}{nonlocal_r}"
                    ).alignment = ctr

            # GROSS / NET weights — LIVE formulas (change a size qty, N.W, CTN MES
            # or a table value and they recalculate).
            #   Gross = ROUND( SUMPRODUCT(size qty row, N.W row) x mult + empty carton wt , 2)
            #   Net   = ROUND( Gross - NET DEDUCT[pack type] , 2)
            # mult: Multi N = prepacks per carton, Single = units per prepack
            #       (size cell holds prepacks), Bulk / Multi Y = 1
            if size_letters:
                if pack == "Multi N":
                    mult = f"*{int(ppc)}" if isinstance(ppc, (int, float)) and ppc != 1 else ""
                elif pack == "Single":
                    upp_s = int(sec.get("units_per_prepack") or 1)
                    mult = f"*{upp_s}" if upp_s != 1 else ""
                else:
                    mult = ""
                rr_ = nonlocal_r
                mes_ref = f"{L(col_idx['CTN MES.'])}{rr_}"
                size_rng = f"{size_letters[0]}{rr_}:{size_letters[-1]}{rr_}"
                gw_f = (f"=ROUND(SUMPRODUCT({size_rng},{nw_rng}){mult}"
                        f"+IFERROR(INDEX({tare_val_rng},MATCH({mes_ref},{tare_hdr_rng},0)),0),2)")
                gw_ref = f"{L(col_idx['GROSS WEIGHT'])}{rr_}"
                nt_f = f"=ROUND({gw_ref}-{deduct_ref.get(pack, deduct_ref['Bulk'])},2)"
                ws.cell(nonlocal_r, col_idx["GROSS WEIGHT"], gw_f)
                ws.cell(nonlocal_r, col_idx["NET WEIGHT"], nt_f)
            else:
                ws.cell(nonlocal_r, col_idx["GROSS WEIGHT"], rowd.get("gross"))
                ws.cell(nonlocal_r, col_idx["NET WEIGHT"], rowd.get("net"))
            for wc in ("GROSS WEIGHT", "NET WEIGHT"):
                ws.cell(nonlocal_r, col_idx[wc]).alignment = ctr
                ws.cell(nonlocal_r, col_idx[wc]).number_format = "0.00"

            for c in range(1, n_cols + 1):
                ws.cell(nonlocal_r, c).border = thin
                ws.cell(nonlocal_r, c).font = font_norm
            nonlocal_r += 1

        last_data = nonlocal_r - 1

        # --- TOTAL row ---
        tr = nonlocal_r
        ws.cell(tr, col_idx["SL#"], "").border = thin
        ws.cell(tr, col_idx["CTN FROM"], "TOTAL").font = font_bold
        ctn_letter = L(col_idx["CTN QTY"])
        tot_letter = L(col_idx["TOTAL QTY"])
        qp_letter = L(col_idx["QTY PER CTN"])

        def _row_product_sum(col_letter: str) -> str:
            parts = [f"{col_letter}{rr}*{ctn_letter}{rr}"
                     for rr in range(first_data, last_data + 1)]
            return "=" + "+".join(parts) if parts else "=0"

        if first_data <= last_data:
            ws.cell(tr, col_idx["CTN QTY"],
                    f"=SUM({ctn_letter}{first_data}:{ctn_letter}{last_data})")
            for s in size_set:
                cl = L(col_idx[s])
                if pack in ("Bulk", "Single"):
                    ws.cell(tr, col_idx[s], _row_product_sum(cl))
                else:
                    # Multi: size units * CTN QTY  (QTY PER CTN already embeds prepacks)
                    parts = [
                        f"{cl}{rr}*{ctn_letter}{rr}"
                        for rr in range(first_data, last_data + 1)
                    ]
                    ws.cell(tr, col_idx[s], "=" + "+".join(parts) if parts else "=0")
            ws.cell(tr, col_idx["TOTAL QTY"],
                    f"=SUM({tot_letter}{first_data}:{tot_letter}{last_data})")
            for wcol in ("GROSS WEIGHT", "NET WEIGHT"):
                wl = L(col_idx[wcol])
                ws.cell(tr, col_idx[wcol],
                        f"=SUM({wl}{first_data}:{wl}{last_data})")
        else:
            ws.cell(tr, col_idx["CTN QTY"], 0)
            ws.cell(tr, col_idx["TOTAL QTY"], 0)

        for c in range(1, n_cols + 1):
            cell = ws.cell(tr, c)
            cell.font = font_bold
            cell.fill = total_fill
            cell.border = thin
            cell.alignment = ctr

        # link PO header ship / ctn / excess to TOTAL
        ship_cell.value = f"={tot_letter}{tr}"
        ctn_hdr_cell.value = f"={ctn_letter}{tr}"
        ex_cell.value = f"={tot_letter}{tr}-C{hdr0 + 3}"

        nonlocal_r = tr + 2
        # order qty by size (best-effort from section data)
        order_by_size = {}
        excess = sec.get("excess_by_size") or {}
        units_map = sec.get("units") or {}
        if pack == "Bulk" and units_map:
            # units = ship_by_size for Bulk
            for s in size_set:
                ship_s = int(units_map.get(s, 0) or 0)
                ex_s = int(excess.get(s, 0) or 0)
                order_by_size[s] = max(0, ship_s - ex_s)
        elif pack in ("Multi Y", "Multi N") and units_map:
            # units = per-prepack; ship ≈ units * prepackets
            pp = int(sec.get("prepackets") or sec.get("ctn_total") or 1)
            for s in size_set:
                order_by_size[s] = int(units_map.get(s, 0) or 0) * pp
        else:
            # Single or fallback: distribute qty_ordered or use row totals
            for s in size_set:
                order_by_size[s] = 0
            if pack == "Single":
                for rowd in rows_data:
                    s = rowd.get("size")
                    if s in order_by_size:
                        order_by_size[s] += int(rowd.get("total_qty") or 0)

        return nonlocal_r, {
            "pack": pack,
            "color": sec.get("color", ""),
            "qty_ordered": sec.get("qty_ordered", 0),
            "total_row": tr,
            "ctn_formula": f"={ctn_letter}{tr}",
            "ship_formula": f"={tot_letter}{tr}",
            "size_formulas": {s: f"={L(col_idx[s])}{tr}" for s in size_set},
            "order_by_size": order_by_size,
        }

    # Write each type in order; serial numbers restart per section
    type_blocks = [(t, _by_type[t]) for t in _order]
    for type_label, secs in type_blocks:
        if not secs:
            continue
        for sec in secs:
            r, refs = _write_one_section(sec, r, serial_start=1)
            section_summary_refs.append(refs)

    # ========== 4. COLOR-WISE SUMMARY + G. TOTAL SUMMARY (same sheet) ==========
    # Layout matches reference:
    #   [COLOR NAME]
    #   SIZE | XS | S | M | ... | TOTAL
    #   ORDER QTY | ... | total
    #   SHIP QTY  | ... | total   (formulas → packing TOTAL rows)
    #   EX/SHORT  | ... | total
    #   %         | ... | total
    # then same block for each color, then "G. Total Summery"

    r += 2
    n_size = len(size_set)
    # columns used in summary tables: A=label, B..=sizes, last=TOTAL
    # use columns 1 .. n_size+2
    # centred: label (LABEL_SPAN merged cols) + n_size size cols + TOTAL col
    c0 = _centered_start(n_size + 1)          # first column of the summary tables
    sz0 = c0 + LABEL_SPAN                     # first size column
    sum_total_col = sz0 + n_size              # column index of TOTAL
    label_rows = []                           # rows whose label cell gets merged at the end

    def _style_sum_row(row, fill=None, bold=False):
        label_rows.append(row)
        for c in range(c0, sum_total_col + 1):
            cell = ws.cell(row, c)
            cell.border = thin
            cell.alignment = ctr
            if fill:
                cell.fill = fill
            if bold:
                cell.font = font_bold

    def _write_color_summary_block(title, order_map, ship_formula_map, start_r):
        """Write one color (or G.Total) summary block. Returns next free row.
        ship_formula_map: size -> formula string like '=H22' or a number.
        order_map: size -> int
        """
        rr = start_r
        # title row
        ws.merge_cells(start_row=rr, start_column=c0, end_row=rr, end_column=sum_total_col)
        cell = ws.cell(rr, c0, title)
        cell.font = font_bold
        cell.fill = summary_fill
        cell.alignment = ctr
        cell.border = thin
        for c in range(c0 + 1, sum_total_col + 1):
            ws.cell(rr, c).border = thin
            ws.cell(rr, c).fill = summary_fill
        rr += 1

        # SIZE header
        ws.cell(rr, c0, "SIZE").font = font_bold
        for j, s in enumerate(size_set):
            ws.cell(rr, sz0 + j, s).font = font_bold
        ws.cell(rr, sum_total_col, "TOTAL").font = font_bold
        _style_sum_row(rr, fill=header_fill, bold=True)
        for c in range(c0, sum_total_col + 1):
            ws.cell(rr, c).font = header_font
            ws.cell(rr, c).fill = header_fill
        rr += 1

        order_row = rr
        ship_row = rr + 1
        ex_row = rr + 2
        pct_row = rr + 3

        # ORDER QTY
        ws.cell(order_row, c0, "ORDER QTY").font = font_bold
        for j, s in enumerate(size_set):
            cell = ws.cell(order_row, sz0 + j, int(order_map.get(s, 0) or 0))
            cell.number_format = "#,##0"
            cell.fill = yellow  # editable
        # TOTAL = sum of size cols
        first_sz = L(sz0)
        last_sz = L(sz0 + n_size - 1)
        ws.cell(order_row, sum_total_col,
                f"=SUM({first_sz}{order_row}:{last_sz}{order_row})")
        _style_sum_row(order_row)

        # SHIP QTY (formulas from packing totals)
        ws.cell(ship_row, c0, "SHIP QTY").font = font_bold
        for j, s in enumerate(size_set):
            val = ship_formula_map.get(s, 0)
            cell = ws.cell(ship_row, sz0 + j, val if val != "" else 0)
            cell.number_format = "#,##0"
        ws.cell(ship_row, sum_total_col,
                f"=SUM({first_sz}{ship_row}:{last_sz}{ship_row})")
        _style_sum_row(ship_row)

        # EX/SHORT = SHIP - ORDER
        ws.cell(ex_row, c0, "EX/SHORT").font = font_bold
        for j in range(n_size):
            cl = L(sz0 + j)
            ws.cell(ex_row, sz0 + j, f"={cl}{ship_row}-{cl}{order_row}")
        ws.cell(ex_row, sum_total_col,
                f"={L(sum_total_col)}{ship_row}-{L(sum_total_col)}{order_row}")
        _style_sum_row(ex_row)

        # % = EX/SHORT / ORDER * 100
        ws.cell(pct_row, c0, "%").font = font_bold
        for j in range(n_size):
            cl = L(sz0 + j)
            ws.cell(pct_row, sz0 + j,
                    f"=IF({cl}{order_row}=0,0,{cl}{ex_row}/{cl}{order_row}*100)")
            ws.cell(pct_row, sz0 + j).number_format = "0.00"
        ws.cell(pct_row, sum_total_col,
                f"=IF({L(sum_total_col)}{order_row}=0,0,"
                f"{L(sum_total_col)}{ex_row}/{L(sum_total_col)}{order_row}*100)")
        ws.cell(pct_row, sum_total_col).number_format = "0.00"
        _style_sum_row(pct_row)

        return pct_row + 2, order_row, ship_row, ex_row, pct_row

    # Group section refs by color (aggregate if same color appears more than once)
    from collections import OrderedDict
    by_color = OrderedDict()
    for refs in section_summary_refs:
        color = refs["color"] or "UNKNOWN"
        if color not in by_color:
            by_color[color] = {
                "order": {s: 0 for s in size_set},
                "ship_parts": {s: [] for s in size_set},  # list of formula strings to sum
            }
        for s in size_set:
            by_color[color]["order"][s] = by_color[color]["order"].get(s, 0) + int(
                (refs.get("order_by_size") or {}).get(s, 0) or 0
            )
            f = (refs.get("size_formulas") or {}).get(s)
            if f:
                # strip leading '=' for joining
                by_color[color]["ship_parts"][s].append(f.lstrip("="))

    color_block_rows = []  # (order_row, ship_row, ex_row, pct_row) per color

    for color, agg in by_color.items():
        ship_map = {}
        for s in size_set:
            parts = agg["ship_parts"].get(s) or []
            if not parts:
                ship_map[s] = 0
            elif len(parts) == 1:
                ship_map[s] = "=" + parts[0]
            else:
                ship_map[s] = "=" + "+".join(parts)
        r, o_r, s_r, e_r, p_r = _write_color_summary_block(
            color, agg["order"], ship_map, r
        )
        color_block_rows.append((o_r, s_r, e_r, p_r))

    # ----- G. Total Summery -----
    # ORDER / SHIP for G.Total = SUM of the corresponding rows from each color block
    g_order = {s: 0 for s in size_set}
    g_ship = {s: "" for s in size_set}
    if color_block_rows:
        for j, s in enumerate(size_set):
            cl = L(sz0 + j)
            g_order[s] = f"={'+'.join(cl + str(o_r) for o_r, _, _, _ in color_block_rows)}"
            g_ship[s] = f"={'+'.join(cl + str(s_r) for _, s_r, _, _ in color_block_rows)}"
        # For order we want values that update if user edits yellow order cells
        # so use formulas referencing color order rows
    else:
        g_order = {s: 0 for s in size_set}
        g_ship = {s: 0 for s in size_set}

    # Write G.Total using formula maps (order cells of color blocks are yellow/editable)
    rr = r
    ws.merge_cells(start_row=rr, start_column=c0, end_row=rr, end_column=sum_total_col)
    cell = ws.cell(rr, c0, "G. Total Summery")
    cell.font = font_bold
    cell.fill = total_fill
    cell.alignment = ctr
    for c in range(c0, sum_total_col + 1):
        ws.cell(rr, c).border = thin
        ws.cell(rr, c).fill = total_fill
    rr += 1

    # SIZE header
    ws.cell(rr, c0, "SIZE").font = header_font
    for j, s in enumerate(size_set):
        ws.cell(rr, sz0 + j, s).font = header_font
        ws.cell(rr, sz0 + j).fill = header_fill
        ws.cell(rr, sz0 + j).border = thin
        ws.cell(rr, sz0 + j).alignment = ctr
    ws.cell(rr, sum_total_col, "TOTAL").font = header_font
    ws.cell(rr, sum_total_col).fill = header_fill
    ws.cell(rr, sum_total_col).border = thin
    for c in range(c0, sz0):
        ws.cell(rr, c).fill = header_fill
        ws.cell(rr, c).border = thin
        ws.cell(rr, c).alignment = ctr
    ws.cell(rr, c0).font = header_font
    label_rows.append(rr)
    rr += 1

    go_row = rr
    gs_row = rr + 1
    ge_row = rr + 2
    gp_row = rr + 3
    first_sz = L(sz0)
    last_sz = L(sz0 + n_size - 1)

    # ORDER QTY (sum of color order rows)
    ws.cell(go_row, c0, "ORDER QTY").font = font_bold
    for j, s in enumerate(size_set):
        cell = ws.cell(go_row, sz0 + j, g_order.get(s, 0))
        cell.number_format = "#,##0"
    ws.cell(go_row, sum_total_col, f"=SUM({first_sz}{go_row}:{last_sz}{go_row})")
    _style_sum_row(go_row, fill=total_fill, bold=True)

    # SHIP QTY
    ws.cell(gs_row, c0, "SHIP QTY").font = font_bold
    for j, s in enumerate(size_set):
        cell = ws.cell(gs_row, sz0 + j, g_ship.get(s, 0))
        cell.number_format = "#,##0"
    ws.cell(gs_row, sum_total_col, f"=SUM({first_sz}{gs_row}:{last_sz}{gs_row})")
    _style_sum_row(gs_row, fill=total_fill, bold=True)

    # EX/SHORT
    ws.cell(ge_row, c0, "EX/SHORT").font = font_bold
    for j in range(n_size):
        cl = L(sz0 + j)
        ws.cell(ge_row, sz0 + j, f"={cl}{gs_row}-{cl}{go_row}")
    ws.cell(ge_row, sum_total_col,
            f"={L(sum_total_col)}{gs_row}-{L(sum_total_col)}{go_row}")
    _style_sum_row(ge_row, fill=total_fill, bold=True)

    # %
    ws.cell(gp_row, c0, "%").font = font_bold
    for j in range(n_size):
        cl = L(sz0 + j)
        ws.cell(gp_row, sz0 + j,
                f"=IF({cl}{go_row}=0,0,{cl}{ge_row}/{cl}{go_row}*100)")
        ws.cell(gp_row, sz0 + j).number_format = "0.00"
    ws.cell(gp_row, sum_total_col,
            f"=IF({L(sum_total_col)}{go_row}=0,0,"
            f"{L(sum_total_col)}{ge_row}/{L(sum_total_col)}{go_row}*100)")
    ws.cell(gp_row, sum_total_col).number_format = "0.00"
    _style_sum_row(gp_row, fill=total_fill, bold=True)

    for lr in label_rows:
        ws.merge_cells(start_row=lr, start_column=c0, end_row=lr, end_column=sz0 - 1)

    # Force Excel to recalculate on open
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.calcOnSave = True
    except Exception:
        pass

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ===========================================================================
# 2. Multi N workbook - sheets 'PKL' + 'sum'
# ===========================================================================
def export_multi_n_excel(
    sections: list,
    nw: dict,
    nnw: dict,
    tare_by_code: dict,
    net_deduct: float,
    ctn_meas: list,
    buyer: str,
    po_no: str,
    grand_total_ctn: int,
    season: str = "Sum'26",
) -> bytes:
    thin = Side(style="thin", color="000000")
    med = Side(style="medium", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    f_title = Font(name=FONT, size=18, bold=True, underline="single")
    f_addr = Font(name=FONT, size=13, bold=True)
    f_bold = Font(name=FONT, size=11, bold=True)
    f_norm = Font(name=FONT, size=11)
    ctr = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    fill_h = PatternFill("solid", fgColor="D9D9D9")
    fill_t = PatternFill("solid", fgColor="F2F2F2")

    all_sizes = [s for sec in sections for s in sec["sizes"]]
    cols = size_family(all_sizes)[:7]          # sheet size columns (G..M)
    SC0 = 7                                    # first size column (G)
    SC1 = SC0 + 6                              # last size column  (M)
    cPB, cX, cP, cQ, cR, cS, cT = 14, 15, 16, 17, 18, 19, 20   # N..T

    wb = Workbook()
    ws = wb.active
    ws.title = "PKL"

    for col, w in {"A": 7, "B": 7, "C": 8, "D": 10, "E": 20, "F": 18,
                   "G": 8, "H": 8, "I": 8, "J": 8, "K": 8, "L": 8, "M": 8,
                   "N": 12, "O": 4, "P": 8, "Q": 12, "R": 12, "S": 14, "T": 14}.items():
        ws.column_dimensions[col].width = w

    # ---- title ----------------------------------------------------------
    ws.merge_cells("A1:T1")
    ws["A1"] = "CREATIVE COLLECTIONS LTD"
    ws["A1"].font, ws["A1"].alignment = f_title, ctr
    ws.merge_cells("A2:T2")
    ws["A2"] = "Nishat Nagar , Tongi , Gazipur ."
    ws["A2"].font, ws["A2"].alignment = f_addr, ctr

    # ---- weight reference table (rows 4-7) ------------------------------
    for r, (a, b) in zip((4, 5, 6, 7), (("SIZE", "SIZE"), ("N.W", "N.WT"),
                                         ("N.N.W", "N.N.WT"), ("EMPTY CTN WET", ""))):
        ws.cell(r, 4, a).font = f_bold
        ws.cell(r, 6, b).font = f_bold
        for c in range(4, SC1 + 1):
            ws.cell(r, c).border = box
    for i, s in enumerate(cols):
        c = SC0 + i
        ws.cell(4, c, s).font = f_bold
        ws.cell(4, c).alignment = ctr
        ws.cell(5, c, nw.get(s)).number_format = "0.000"
        ws.cell(6, c, nnw.get(s)).number_format = "0.000"
        ws.cell(5, c).alignment = ws.cell(6, c).alignment = ctr
    ws.merge_cells("D7:E7")

    # carton tare table  N4:T7  (same place as dummy_61480360_packing_list.xlsx)
    codes = list(tare_by_code)[:4]
    for rr_, lab_ in ((4, "CTN CODE"), (5, "EMPTY CTN WT"), (6, "GROSS-NET MULTI")):
        ws.merge_cells(start_row=rr_, start_column=14, end_row=rr_, end_column=16)
        ws.cell(rr_, 14, lab_).font = f_bold
    for i, code in enumerate(codes):
        ws.cell(4, 17 + i, code).font = f_bold
        ws.cell(5, 17 + i, tare_by_code[code]).number_format = "0.00"
        ws.cell(4, 17 + i).alignment = ws.cell(5, 17 + i).alignment = ctr
    ws.cell(6, 17, net_deduct).number_format = "0.00"   # $Q$6
    for r in (4, 5, 6):
        for c in range(14, 21):
            ws.cell(r, c).border = box
    TARE_CODES, TARE_VALS, NET_CELL = "$Q$4:$T$4", "$Q$5:$T$5", "$Q$6"

    # ---- sections -------------------------------------------------------
    r = 9
    prev_last_row = None
    sum_refs = []                      # for 'sum' sheet
    for sec in sections:
        # strip
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=20)
        c = ws.cell(r, 1, "MULTY PACK ( N )")
        c.font, c.alignment = Font(name=FONT, size=13, bold=True), ctr
        c.border = Border(top=med, bottom=med)
        strip = r
        r += 1

        hdr0 = r
        po_left = ["BUYER", "STYLE", "P. O. #", "O/QTY", "SHIP QTY", "EX/SHORT", "CTN QTY"]
        for k, lab in enumerate(po_left):
            ws.merge_cells(start_row=r + k, start_column=1, end_row=r + k, end_column=2)
            ws.cell(r + k, 1, lab).font = f_bold
            ws.cell(r + k, 3, ":").alignment = ctr
            ws.merge_cells(start_row=r + k, start_column=4, end_row=r + k, end_column=8)
            ws.cell(r + k, 4).alignment = left
            ws.cell(r + k, 4).font = f_norm
        r += len(po_left)
        for m in ctn_meas:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws.cell(r, 1, "CTN MEAS.").font = f_bold
            ws.cell(r, 3, ":").alignment = ctr
            ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
            ws.cell(r, 4, m).font = f_norm
            r += 1

        # table header
        th = r
        for i, h in enumerate(["CTN", "", "CTN", "CTN QTY", "COLOR", "PREPACK STECKER"]):
            ws.cell(th, 1 + i, h)
        ws.merge_cells(start_row=th, start_column=SC0, end_row=th, end_column=SC1)
        ws.cell(th, SC0, "SIZE")
        for c, h in ((cPB, "PER BLST"), (cQ, "QTY PER CTN"), (cR, "TOTAL QTY"),
                     (cS, "GROSS WEIGHT"), (cT, "NET WEIGHT")):
            ws.cell(th, c, h)
        ws.merge_cells(start_row=th, start_column=cPB, end_row=th + 1, end_column=cP)
        ws.cell(th + 1, 1, "NO")
        ws.cell(th + 1, 3, "MES:")
        ws.cell(th + 1, 5, "COLOR")
        for i in range(7):
            ws.cell(th + 1, SC0 + i, f"={L(SC0 + i)}$4")
        for rr in (th, th + 1):
            for c in range(1, cT + 1):
                cell = ws.cell(rr, c)
                cell.font, cell.fill, cell.border = f_bold, fill_h, box
                cell.alignment = ctr
        r = th + 2

        # data rows
        first_data = r
        for k, row in enumerate(sec["rows"]):
            # CTN NO from / to
            if k == 0 and prev_last_row is None:
                ws.cell(r, 1, row["ctn_from"])
            elif k == 0:
                ws.cell(r, 1, f"=B{prev_last_row}+1")
            else:
                ws.cell(r, 1, f"=B{r - 1}+1")
            ws.cell(r, 2, f"=A{r}+D{r}-1")
            ws.cell(r, 3, row["ctn_mes"])
            ws.cell(r, 4, row["ctn_qty"])
            if k == 0:
                ws.cell(r, 5, sec["color"])
                ws.cell(r, 6, int(sec["prepack"]))
            for i, s in enumerate(cols):
                u = sec["units"].get(s)
                if u:
                    ws.cell(r, SC0 + i, u)
            ws.cell(r, cPB, f"=SUM({L(SC0)}{r}:{L(SC1)}{r})")
            ws.cell(r, cX, "X")
            ws.cell(r, cP, row["prepacks_per_ctn"])
            ws.cell(r, cQ, f"={L(cPB)}{r}*{L(cP)}{r}")
            ws.cell(r, cR, f"={L(cQ)}{r}*D{r}")
            ws.cell(r, cS,
                    f"=SUMPRODUCT({L(SC0)}{r}:{L(SC1)}{r},$G$5:$M$5)*{L(cP)}{r}"
                    f"+IFERROR(INDEX({TARE_VALS},MATCH(C{r},{TARE_CODES},0)),0)")
            ws.cell(r, cT, f"={L(cS)}{r}-{NET_CELL}")
            ws.cell(r, cS).number_format = ws.cell(r, cT).number_format = "0.00"
            for c in range(1, cT + 1):
                ws.cell(r, c).border = box
                ws.cell(r, c).alignment = ctr
                ws.cell(r, c).font = f_norm
            r += 1
        last_data = r - 1
        prev_last_row = last_data

        # TOTAL row
        tr = r
        ws.cell(tr, 1, "TOTAL")
        ws.cell(tr, 4, f"=SUM(D{first_data}:D{last_data})")
        ws.cell(tr, 5, "CTNS")
        for i, s in enumerate(cols):
            if sec["units"].get(s):
                col = L(SC0 + i)
                ws.cell(tr, SC0 + i,
                        f"=SUMPRODUCT({col}{first_data}:{col}{last_data},"
                        f"{L(cP)}{first_data}:{L(cP)}{last_data},"
                        f"D{first_data}:D{last_data})")
        ws.cell(tr, cR, f"=SUM({L(cR)}{first_data}:{L(cR)}{last_data})")
        ws.cell(tr, cS, "PCS")
        for c in range(1, cT + 1):
            cell = ws.cell(tr, c)
            cell.font, cell.fill, cell.border, cell.alignment = f_bold, fill_t, box, ctr

        # header values (need total-row refs)
        vals = [
            buyer,
            sec["style_full"],
            po_no,
            f"{sec['qty_ordered']}",
            f"=R{tr}",
            f"=D{hdr0 + 4}-D{hdr0 + 3}",
            f"=D{tr}",
        ]
        # put O/QTY as number + unit text via number format
        for k, v in enumerate(vals):
            cell = ws.cell(hdr0 + k, 4)
            cell.value = int(v) if k == 3 else v
            if k >= 3:
                cell.number_format = '#,##0 "PCS"' if k < 6 else '#,##0 "CTNS"'
            cell.font, cell.alignment = f_norm, left
        # right block (PO # .. CARTON) beside O/QTY..CTN QTY lines
        carton_label = (f'=TEXT(A{first_data},"00")&" of {grand_total_ctn}"')
        right = [("PO #", f"=D{hdr0 + 2}"),
                 ("SKU / Item", f"=D{hdr0 + 1}"),
                 ("Unit / Prepack", f"{sec['units_per_prepack']}/{sec['per_ctn']}"),
                 ("CARTON", carton_label)]
        for k, (lab, val) in enumerate(right):
            rr = hdr0 + 3 + k
            ws.merge_cells(start_row=rr, start_column=11, end_row=rr, end_column=12)
            ws.cell(rr, 11, lab).font = f_bold
            ws.cell(rr, 13, ":").alignment = ctr
            ws.merge_cells(start_row=rr, start_column=14, end_row=rr, end_column=17)
            ws.cell(rr, 14, val).font = f_norm
            ws.cell(rr, 14).alignment = left

        sum_refs.append({
            "color": sec["color"], "qty": sec["qty_ordered"],
            "ship": f"=PKL!R{tr}", "ctns": f"=PKL!D{tr}",
        })
        r = tr + 2

    # ---- sheet 'sum' ----------------------------------------------------
    s2 = wb.create_sheet("sum")
    for col, w in zip("ABCDEFGHIJ", (8, 14, 10, 20, 12, 12, 10, 10, 10, 10)):
        s2.column_dimensions[col].width = w
    s2.merge_cells("A3:J3")
    s2["A3"] = "CREATIVE COLLECTIONS LTD"
    s2["A3"].font, s2["A3"].alignment = f_addr, ctr
    s2.merge_cells("A4:J4")
    s2["A4"] = "Nishatnagar,Tongi,Gazipur"
    s2["A4"].alignment = ctr
    s2.merge_cells("A5:J5")
    s2["A5"] = "PACKING LIST SUMMARY"
    s2["A5"].font, s2["A5"].alignment = f_bold, ctr
    s2["A7"], s2["B7"] = "BUYER", buyer
    s2["A8"], s2["B8"] = "STYLE", sections[0]["style_no"] if sections else ""
    s2["A7"].font = s2["A8"].font = f_bold
    for i, h in enumerate(["DC", "PO NO", "SEASON", "COLOR/CODE", "ORD/QTY", "SHIP/QTY",
                           "EXCES", "SHORT", "PER (%)", "CTNS"]):
        c = s2.cell(10, 1 + i, h)
        c.font, c.fill, c.border, c.alignment = f_bold, fill_h, box, ctr
    rr = 11
    for ref in sum_refs:
        s2.cell(rr, 1, "BDC")
        s2.cell(rr, 2, po_no)
        s2.cell(rr, 3, season)
        s2.cell(rr, 4, ref["color"])
        s2.cell(rr, 5, ref["qty"])
        s2.cell(rr, 6, ref["ship"])
        s2.cell(rr, 7, f"=MAX(F{rr}-E{rr},0)")
        s2.cell(rr, 8, f"=MAX(E{rr}-F{rr},0)")
        s2.cell(rr, 9, f"=IF(E{rr}=0,0,(F{rr}-E{rr})/E{rr})")
        s2.cell(rr, 9).number_format = "0.00%"
        s2.cell(rr, 10, ref["ctns"])
        for c in range(1, 11):
            s2.cell(rr, c).border, s2.cell(rr, c).alignment = box, ctr
        rr += 1
    s2.cell(rr, 4, "TOTAL")
    for c, col in ((5, "E"), (6, "F"), (7, "G"), (8, "H"), (10, "J")):
        s2.cell(rr, c, f"=SUM({col}11:{col}{rr - 1})")
    s2.cell(rr, 9, f"=IF(E{rr}=0,0,(F{rr}-E{rr})/E{rr})")
    s2.cell(rr, 9).number_format = "0.00%"
    for c in range(1, 11):
        cell = s2.cell(rr, c)
        cell.font, cell.fill, cell.border, cell.alignment = f_bold, fill_t, box, ctr

    # print setup: landscape, one page wide
    for sh in (ws, s2):
        sh.page_setup.orientation = "landscape"
        sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
