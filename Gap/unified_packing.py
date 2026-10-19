"""
unified_packing.py
------------------
GAP / Old Navy PO → Packing List (Unified)

Logic  : GAP (Bulk / Single / Multi Y / Multi N) — merge same STYLE by COLOR
UI     : Interactive preview (yellow strips, PO header, live totals, Excel)

Run:
    streamlit run unified_packing.py
"""

from __future__ import annotations

import io
import json
from pathlib import Path

import sys
import hashlib

import pandas as pd
import streamlit as st

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:                      # pl_kit.py lives in the project root
    sys.path.insert(0, _ROOT)
import pl_kit as kit                           # noqa: E402

import parser_lib as parser
from color_builders import (
    BULK_CTN_CAP,
    DEFAULT_CTN_CAP,
    DEFAULT_CTN_MEAS,
    DEFAULT_EXCESS_PCT,
    DEFAULT_NW,
    build_bulk_by_color,
    build_multi_y_by_color,
    build_single_by_color,
    carton_code,
    carton_weights,
    detect_table_type,
    rebuild_one_size_section,
    section_to_display_rows,
    sort_sizes,
)
from multi_n_core import (
    DEFAULT_CTN_CODE,
    DEFAULT_NET_DEDUCT,
    DEFAULT_NNW,
    DEFAULT_PREPACKS_PER_CTN,
    DEFAULT_TARE,
    build_multi_n,
)
from multi_n_export import export_multi_n_excel

PACK_TYPES = ("Bulk", "Single", "Multi Y", "Multi N")


def detect_all(data: dict) -> dict[str, list[str]]:
    found = {t: [] for t in PACK_TYPES}
    for tname, table in (data.get("tables") or {}).items():
        ptype = detect_table_type(table)
        if ptype:
            found[ptype].append(tname)
    return found


def inject_css():
    st.markdown(
        """
<style>
  .plist-wrap { font-family: 'Courier New', monospace; font-size: 13px; }
  .cc-strip {
    text-align: center; font-weight: bold; font-size: 15px;
    letter-spacing: 2px; background: #FFFF00; border: 1px solid #000;
    padding: 6px 0; margin: 14px 0 8px 0;
  }
  .po-grid {
    display: grid; grid-template-columns: 1.3fr 1fr; gap: 18px;
    margin-bottom: 8px; line-height: 1.45;
  }
  .po-line { display: flex; }
  .po-label { width: 110px; font-weight: bold; }
  .po-colon { width: 14px; text-align: center; }
  table.plist {
    border-collapse: collapse; width: 100%; margin: 6px 0 14px 0;
    font-family: 'Courier New', monospace; font-size: 12.5px;
  }
  table.plist th, table.plist td {
    border: 1px solid #333; padding: 3px 6px; text-align: center;
  }
  table.plist th { background: #4472C4; color: #fff; font-weight: bold; }
  table.plist tr.total td { background: #FFF2CC; font-weight: bold; }
</style>
""",
        unsafe_allow_html=True,
    )


def render_section_strip(title: str):
    st.markdown(f'<div class="cc-strip">{title}</div>', unsafe_allow_html=True)


def render_po_header_html(sec: dict, buyer: str, po_no: str, grand_ctn: int):
    meas = "".join(
        f'<div class="po-line"><span class="po-label">CTN MEAS.</span>'
        f'<span class="po-colon">:</span><span>{m}</span></div>'
        for m in (sec.get("ctn_meas_list") or DEFAULT_CTN_MEAS)
    )
    ctn_from = sec["rows"][0]["ctn_from"] if sec.get("rows") else 1
    exs = sec["ship_qty"] - sec["qty_ordered"]
    html = f"""
<div class="plist-wrap">
  <div class="po-grid">
    <div>
      <div class="po-line"><span class="po-label">BUYER</span><span class="po-colon">:</span><span>{buyer}</span></div>
      <div class="po-line"><span class="po-label">STYLE</span><span class="po-colon">:</span><span>{sec.get('style_full','')}</span></div>
      <div class="po-line"><span class="po-label">P.O. #</span><span class="po-colon">:</span><span>{po_no}</span></div>
      <div class="po-line"><span class="po-label">O/QTY</span><span class="po-colon">:</span><span>{sec['qty_ordered']} PCS</span></div>
      <div class="po-line"><span class="po-label">SHIP QTY</span><span class="po-colon">:</span><span>{sec['ship_qty']} PCS</span></div>
      <div class="po-line"><span class="po-label">EX/SHORT</span><span class="po-colon">:</span><span>{exs} PCS</span></div>
      <div class="po-line"><span class="po-label">CTN QTY</span><span class="po-colon">:</span><span>{sec['ctn_total']} CTNS</span></div>
      {meas}
    </div>
    <div style="margin-top:48px">
      <div class="po-line"><span class="po-label">PO #</span><span class="po-colon">:</span><span>{po_no}</span></div>
      <div class="po-line"><span class="po-label">SKU / Item</span><span class="po-colon">:</span><span>{sec.get('style_full','')}</span></div>
      <div class="po-line"><span class="po-label">Unit / Prepack</span><span class="po-colon">:</span><span>{sec.get('units_per_prepack',1)}</span></div>
      <div class="po-line"><span class="po-label">CARTON</span><span class="po-colon">:</span><span>{ctn_from:02d} of {grand_ctn}</span></div>
    </div>
  </div>
</div>
"""
    st.markdown(html, unsafe_allow_html=True)


def render_plist_table(sec: dict):
    recs = section_to_display_rows(sec)
    if not recs:
        st.warning("No rows")
        return
    cols = list(recs[0].keys())
    th = "".join(f"<th>{c}</th>" for c in cols)
    body = ""
    for rec in recs:
        cls = ' class="total"' if str(rec.get("CTN NO", "")).upper() == "TOTAL" else ""
        tds = "".join(f"<td>{rec.get(c, '')}</td>" for c in cols)
        body += f"<tr{cls}>{tds}</tr>"
    st.markdown(
        f'<div class="plist-wrap"><table class="plist"><tr>{th}</tr>{body}</table></div>',
        unsafe_allow_html=True,
    )


def export_unified_excel(
    data, bulk_secs, single_secs, multi_y_secs, multi_n_secs,
    nw: dict | None = None,
    nnw: dict | None = None,
    tare_by_code: dict | None = None,
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
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L

    from color_builders import DEFAULT_NW, sort_sizes
    try:
        from multi_n_core import DEFAULT_NNW, DEFAULT_TARE
    except Exception:
        DEFAULT_NNW = {s: round(v - 0.02, 3) for s, v in DEFAULT_NW.items()}
        DEFAULT_TARE = {"G81": 0.71, "G82": 0.71, "G84": 0.24, "G85": 0.24}

    nw = dict(nw or DEFAULT_NW)
    nnw = dict(nnw or DEFAULT_NNW)
    tare_by_code = dict(tare_by_code or DEFAULT_TARE)

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

    all_secs = list(bulk_secs) + list(single_secs) + list(multi_y_secs) + list(multi_n_secs)
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
    widths = {
        "SL#": 5, "CTN FROM": 8, "CTN TO": 8, "CTN MES.": 8, "CTN QTY": 8,
        "COLOR": 14, "PREPACK STECKER": 12,
        "QTY PER CTN": 10, "TOTAL QTY": 10, "GROSS WEIGHT": 11, "NET WEIGHT": 11,
    }
    for name, c in col_idx.items():
        ws.column_dimensions[L(c)].width = widths.get(name, 7)

    r = 1  # current row

    # ========== 1. COMPANY HEADER ==========
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    cell = ws.cell(r, 1, "CREATIVE COLLECTIONS LTD-1A.")
    cell.font = font_title
    cell.alignment = ctr
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    cell = ws.cell(r, 1, "Nishat Nagar , Tongi , Gazipur .")
    cell.font = font_addr
    cell.alignment = ctr
    r += 2

    # ========== 2a. TABLE A — Net weight by SIZE (editable) ==========
    #   SIZE | XS | S | M | L | XL | XXL | XXL+
    #   N.W  | values...
    #   N.N.W| values...
    nw_start_col = 2  # sizes start at column B
    nw_header_row = r
    nw_row = r + 1
    nnw_row = r + 2

    ws.cell(nw_header_row, 1, "SIZE").font = header_font
    ws.cell(nw_header_row, 1).fill = header_fill
    ws.cell(nw_header_row, 1).border = thin
    ws.cell(nw_header_row, 1).alignment = ctr
    ws.cell(nw_row, 1, "N.W / N.WT").font = font_bold
    ws.cell(nw_row, 1).border = thin
    ws.cell(nnw_row, 1, "N.N.W / N.N.WT").font = font_bold
    ws.cell(nnw_row, 1).border = thin

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

    # ========== 2b. TABLE B — Empty carton weight by CTN CODE (separate, editable) ==========
    tare_codes = list(tare_by_code.keys()) or ["G81", "G82", "G84", "G85"]
    tare_header_row = r
    tare_val_row = r + 1

    ws.cell(tare_header_row, 1, "CTN CODE").font = header_font
    ws.cell(tare_header_row, 1).fill = header_fill
    ws.cell(tare_header_row, 1).border = thin
    ws.cell(tare_header_row, 1).alignment = ctr
    ws.cell(tare_val_row, 1, "EMPTY CTN WT").font = font_bold
    ws.cell(tare_val_row, 1).border = thin

    tare_cell_ref = {}
    for i, code in enumerate(tare_codes):
        c = 2 + i
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

    r = tare_val_row + 2

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
        rc_lab = max(n_cols - 3, 10)   # label column on the right
        rc_val = rc_lab + 1            # value column
        ctn_from0 = rows_data[0].get("ctn_from", 1) if rows_data else 1
        right_items = [
            ("PO #", po_no),
            ("SKU / Item", sec.get("style_full") or style),
            ("Unit / Prepack", str(sec.get("units_per_prepack", 1))),
            ("CARTON", f"{int(ctn_from0):02d} of {grand_ctn}"),
        ]
        for i, (lab, val) in enumerate(right_items):
            ws.cell(hdr0 + i, rc_lab, lab).font = font_bold
            ws.cell(hdr0 + i, rc_lab).alignment = left
            cell = ws.cell(hdr0 + i, rc_val, val)
            cell.alignment = left
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

            # GROSS / NET weights — formula using NW table when possible
            # For Bulk/Single: gross ≈ size_unit * NW[size] * ctn_qty + tare (simplified)
            gw = rowd.get("gross")
            nw_val = rowd.get("net")
            if gw not in ("", None):
                ws.cell(nonlocal_r, col_idx["GROSS WEIGHT"], gw).alignment = ctr
            else:
                ws.cell(nonlocal_r, col_idx["GROSS WEIGHT"], None).alignment = ctr
            if nw_val not in ("", None):
                ws.cell(nonlocal_r, col_idx["NET WEIGHT"], nw_val).alignment = ctr
            else:
                ws.cell(nonlocal_r, col_idx["NET WEIGHT"], None).alignment = ctr

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
    type_blocks = [
        ("Bulk", bulk_secs),
        ("Single", single_secs),
        ("Multi Y", multi_y_secs),
        ("Multi N", multi_n_secs),
    ]
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
    sum_total_col = n_size + 2  # column index of TOTAL

    def _style_sum_row(row, fill=None, bold=False):
        for c in range(1, sum_total_col + 1):
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
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=sum_total_col)
        cell = ws.cell(rr, 1, title)
        cell.font = font_bold
        cell.fill = summary_fill
        cell.alignment = ctr
        cell.border = thin
        for c in range(2, sum_total_col + 1):
            ws.cell(rr, c).border = thin
            ws.cell(rr, c).fill = summary_fill
        rr += 1

        # SIZE header
        ws.cell(rr, 1, "SIZE").font = font_bold
        for j, s in enumerate(size_set):
            ws.cell(rr, 2 + j, s).font = font_bold
        ws.cell(rr, sum_total_col, "TOTAL").font = font_bold
        _style_sum_row(rr, fill=header_fill, bold=True)
        for c in range(1, sum_total_col + 1):
            ws.cell(rr, c).font = header_font
            ws.cell(rr, c).fill = header_fill
        rr += 1

        order_row = rr
        ship_row = rr + 1
        ex_row = rr + 2
        pct_row = rr + 3

        # ORDER QTY
        ws.cell(order_row, 1, "ORDER QTY").font = font_bold
        for j, s in enumerate(size_set):
            cell = ws.cell(order_row, 2 + j, int(order_map.get(s, 0) or 0))
            cell.number_format = "#,##0"
            cell.fill = yellow  # editable
        # TOTAL = sum of size cols
        first_sz = L(2)
        last_sz = L(1 + n_size)
        ws.cell(order_row, sum_total_col,
                f"=SUM({first_sz}{order_row}:{last_sz}{order_row})")
        _style_sum_row(order_row)

        # SHIP QTY (formulas from packing totals)
        ws.cell(ship_row, 1, "SHIP QTY").font = font_bold
        for j, s in enumerate(size_set):
            val = ship_formula_map.get(s, 0)
            cell = ws.cell(ship_row, 2 + j, val if val != "" else 0)
            cell.number_format = "#,##0"
        ws.cell(ship_row, sum_total_col,
                f"=SUM({first_sz}{ship_row}:{last_sz}{ship_row})")
        _style_sum_row(ship_row)

        # EX/SHORT = SHIP - ORDER
        ws.cell(ex_row, 1, "EX/SHORT").font = font_bold
        for j in range(n_size):
            cl = L(2 + j)
            ws.cell(ex_row, 2 + j, f"={cl}{ship_row}-{cl}{order_row}")
        ws.cell(ex_row, sum_total_col,
                f"={L(sum_total_col)}{ship_row}-{L(sum_total_col)}{order_row}")
        _style_sum_row(ex_row)

        # % = EX/SHORT / ORDER * 100
        ws.cell(pct_row, 1, "%").font = font_bold
        for j in range(n_size):
            cl = L(2 + j)
            ws.cell(pct_row, 2 + j,
                    f"=IF({cl}{order_row}=0,0,{cl}{ex_row}/{cl}{order_row}*100)")
            ws.cell(pct_row, 2 + j).number_format = "0.00"
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
            cl = L(2 + j)
            g_order[s] = f"={'+'.join(cl + str(o_r) for o_r, _, _, _ in color_block_rows)}"
            g_ship[s] = f"={'+'.join(cl + str(s_r) for _, s_r, _, _ in color_block_rows)}"
        # For order we want values that update if user edits yellow order cells
        # so use formulas referencing color order rows
    else:
        g_order = {s: 0 for s in size_set}
        g_ship = {s: 0 for s in size_set}

    # Write G.Total using formula maps (order cells of color blocks are yellow/editable)
    rr = r
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=sum_total_col)
    cell = ws.cell(rr, 1, "G. Total Summery")
    cell.font = font_bold
    cell.fill = total_fill
    cell.alignment = ctr
    for c in range(1, sum_total_col + 1):
        ws.cell(rr, c).border = thin
        ws.cell(rr, c).fill = total_fill
    rr += 1

    # SIZE header
    ws.cell(rr, 1, "SIZE").font = header_font
    for j, s in enumerate(size_set):
        ws.cell(rr, 2 + j, s).font = header_font
        ws.cell(rr, 2 + j).fill = header_fill
        ws.cell(rr, 2 + j).border = thin
        ws.cell(rr, 2 + j).alignment = ctr
    ws.cell(rr, sum_total_col, "TOTAL").font = header_font
    ws.cell(rr, sum_total_col).fill = header_fill
    ws.cell(rr, sum_total_col).border = thin
    ws.cell(rr, 1).fill = header_fill
    ws.cell(rr, 1).border = thin
    ws.cell(rr, 1).font = header_font
    rr += 1

    go_row = rr
    gs_row = rr + 1
    ge_row = rr + 2
    gp_row = rr + 3
    first_sz = L(2)
    last_sz = L(1 + n_size)

    # ORDER QTY (sum of color order rows)
    ws.cell(go_row, 1, "ORDER QTY").font = font_bold
    for j, s in enumerate(size_set):
        cell = ws.cell(go_row, 2 + j, g_order.get(s, 0))
        cell.number_format = "#,##0"
    ws.cell(go_row, sum_total_col, f"=SUM({first_sz}{go_row}:{last_sz}{go_row})")
    _style_sum_row(go_row, fill=total_fill, bold=True)

    # SHIP QTY
    ws.cell(gs_row, 1, "SHIP QTY").font = font_bold
    for j, s in enumerate(size_set):
        cell = ws.cell(gs_row, 2 + j, g_ship.get(s, 0))
        cell.number_format = "#,##0"
    ws.cell(gs_row, sum_total_col, f"=SUM({first_sz}{gs_row}:{last_sz}{gs_row})")
    _style_sum_row(gs_row, fill=total_fill, bold=True)

    # EX/SHORT
    ws.cell(ge_row, 1, "EX/SHORT").font = font_bold
    for j in range(n_size):
        cl = L(2 + j)
        ws.cell(ge_row, 2 + j, f"={cl}{gs_row}-{cl}{go_row}")
    ws.cell(ge_row, sum_total_col,
            f"={L(sum_total_col)}{gs_row}-{L(sum_total_col)}{go_row}")
    _style_sum_row(ge_row, fill=total_fill, bold=True)

    # %
    ws.cell(gp_row, 1, "%").font = font_bold
    for j in range(n_size):
        cl = L(2 + j)
        ws.cell(gp_row, 2 + j,
                f"=IF({cl}{go_row}=0,0,{cl}{ge_row}/{cl}{go_row}*100)")
        ws.cell(gp_row, 2 + j).number_format = "0.00"
    ws.cell(gp_row, sum_total_col,
            f"=IF({L(sum_total_col)}{go_row}=0,0,"
            f"{L(sum_total_col)}{ge_row}/{L(sum_total_col)}{go_row}*100)")
    ws.cell(gp_row, sum_total_col).number_format = "0.00"
    _style_sum_row(gp_row, fill=total_fill, bold=True)

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
# Streamlit UI
# ===========================================================================
st.set_page_config(page_title="GAP PO → Packing List", layout="wide", page_icon="📦")
inject_css()
st.title("📦 GAP / Old Navy — PO → Packing List")
st.caption(
    "Upload PDF → extract → auto-detect **Bulk / Single / Multi Y / Multi N** → "
    "merge same STYLE by **COLOR** → interactive packing list + Excel"
)

uploaded = st.file_uploader("Upload Destination Purchase Order PDF", type=["pdf"])
with st.expander("Or load an existing JSON"):
    json_file = st.file_uploader("JSON file", type=["json"], key="json_up")

data = None
source_name = ""

if uploaded is not None:
    source_name = uploaded.name
    sig = f"{uploaded.name}::{uploaded.size}"
    if st.session_state.get("_upload_sig") != sig:
        for k in list(st.session_state.keys()):
            if k.startswith(("xlsx_", "mn_xlsx_")):
                st.session_state.pop(k, None)
        st.session_state["_upload_sig"] = sig
    with st.spinner("Reading PDF…"):
        try:
            text = parser.extract_text(io.BytesIO(uploaded.getvalue()))
        except Exception as e:
            st.error(f"PDF read failed: {e}")
            st.stop()
    with st.spinner("Parsing PO structure…"):
        try:
            data = parser.parse_text(text)
        except Exception as e:
            st.error(f"Parse failed: {e}")
            st.stop()
    st.success(
        f"Found **{len(data.get('tables', {}))}** table(s) "
        f"from `{uploaded.name}` ({len(text):,} characters)"
    )
elif json_file is not None:
    source_name = json_file.name
    try:
        data = json.loads(json_file.getvalue().decode("utf-8"))
        st.success(f"Loaded JSON `{json_file.name}`")
    except Exception as e:
        st.error(f"Invalid JSON: {e}")
        st.stop()

if data is None:
    st.info("Upload a GAP / Old Navy Destination PO PDF to get started.")
    st.markdown("""
### Logic (GAP)
| Type | Detection | Grouping |
|------|-----------|----------|
| **Bulk** | `PrePack Type == Bulk` | One section **per color** |
| **Single** | `PrePack Type == Single` | One section **per color** (prepacks merged) |
| **Multi Y** | `Multi` + `Full Carton == Y` | One section **per color** (all sizes one row) |
| **Multi N** | `Multi` + `Full Carton == N` | Per prepack (shared cartons) |

Same **STYLE** → tables merged by **COLOR** only.
""")
    st.stop()

header = data.get("PO_HEADER") or {}
buyer = header.get("BUYER") or "OLD NAVY"
style = header.get("STYLE") or ""
po_no = str(header.get("P.O._#") or header.get("P.O. #") or "")
oiqty = header.get("OIQTY") or 0

detected = detect_all(data)
active_types = [t for t in PACK_TYPES if detected[t]]

kit.step(1, "PO summary")
kit.summary_cards([("Buyer", str(buyer)), ("Style", str(style)), ("PO #", po_no or "-"), ("Order qty", f"{oiqty} pcs")])
dcols = st.columns(4)
for i, ptype in enumerate(PACK_TYPES):
    n = len(detected[ptype])
    if n:
        dcols[i].success(f"**{ptype}** — {n} table(s)")
    else:
        dcols[i].info(f"{ptype} — none")

if not active_types:
    st.error("This PDF has no GAP / Old Navy packing tables (PrePack Type Bulk / Single / Multi), "
             "so there is nothing to build here.")
    st.info("Is it another buyer's PO (for example Tally Weijl)? Choose the buyer in the router sidebar "
            "under **Pick a buyer if auto-detection fails**.")
    with st.expander("Technical details - what the GAP parser read"):
        st.json(data)
    st.stop()

kit.step(2, "Variables", "Packing defaults. Changing them rebuilds the carton tables below.")
with st.container(border=True):
    v1, v2, v3 = st.columns(3)
    prepacks_per_ctn = v1.number_input("Prepacks / carton (Multi N)", min_value=1, max_value=50, value=DEFAULT_PREPACKS_PER_CTN)
    single_max_pp = v2.number_input(
        "Max prepacks / carton (Single) — 0 = auto by size", min_value=0, max_value=50, value=0)
    excess_pct = v3.number_input(
        "Max excess % per size (Bulk)", min_value=0.0, max_value=10.0,
        value=float(DEFAULT_EXCESS_PCT), step=0.5,
        help="Extra pieces shipped with Bulk = floor(this % of each size's total order).")
    _sizes = sort_sizes(sorted({
        str(r.get("Size Desc") or "")
        for t in (data.get("tables") or {}).values() for r in (t.get("rows") or [])
    } - {""}))
    caps = {}
    with st.expander("Carton capacity (pcs per carton)"):
        ccols = st.columns(min(max(len(_sizes), 1), 8))
        for j, _s in enumerate(_sizes):
            caps[_s] = ccols[j % len(ccols)].number_input(
                _s, min_value=1, max_value=500,
                value=int(BULK_CTN_CAP.get(_s, DEFAULT_CTN_CAP)), key=f"cap_{_s}")
    w1, w2 = st.columns(2)
    net_deduct = w1.number_input("Net deduct (Multi N)", value=float(DEFAULT_NET_DEDUCT), step=0.01)
    ctn_code = w2.text_input("Carton code (Multi N)", value=DEFAULT_CTN_CODE)

    # ----- Table A: Net weight by SIZE — 3-row table (header / N.W / N.N.W) -----
    st.markdown("**Net weight by size (N.W / N.N.W)** — editable")
    _nw_sizes = _sizes or ["XS", "S", "M", "L", "XL", "XXL", "XXL+"]
    _nw_row = {s: float(DEFAULT_NW.get(s, 0.5)) for s in _nw_sizes}
    _nnw_row = {
        s: float(DEFAULT_NNW.get(s, round(float(DEFAULT_NW.get(s, 0.5)) - 0.02, 3)))
        for s in _nw_sizes
    }
    nw_df = pd.DataFrame(
        [_nw_row, _nnw_row],
        index=["N.W", "N.N.W"],
        columns=_nw_sizes,
    )
    edited_nw = st.data_editor(
        nw_df,
        key="nw_table",
        use_container_width=True,
        hide_index=False,
        num_rows="fixed",          # no add/delete rows
        disabled=False,
    )
    # Read back edited values (plain numbers, no steppers)
    user_nw = {s: float(edited_nw.loc["N.W", s] or 0) for s in _nw_sizes}
    user_nnw = {s: float(edited_nw.loc["N.N.W", s] or 0) for s in _nw_sizes}

    # ----- Table B: Empty carton weight by CTN CODE (separate 2-row table) -----
    st.markdown("**Empty carton weight by CTN CODE** — editable (separate table)")
    _tare_codes = list(DEFAULT_TARE.keys()) or ["G81", "G82", "G84", "G85"]
    _tare_row = {c: float(DEFAULT_TARE.get(c, 0.71)) for c in _tare_codes}
    tare_df = pd.DataFrame(
        [_tare_row],
        index=["EMPTY CTN WT"],
        columns=_tare_codes,
    )
    edited_tare = st.data_editor(
        tare_df,
        key="tare_table",
        use_container_width=True,
        hide_index=False,
        num_rows="fixed",
        disabled=False,
    )
    user_tare = {c: float(edited_tare.loc["EMPTY CTN WT", c] or 0) for c in _tare_codes}

sig = hashlib.md5(repr((prepacks_per_ctn, single_max_pp, excess_pct, sorted(caps.items()),
                        net_deduct, ctn_code,
                        sorted(user_nw.items()), sorted(user_nnw.items()),
                        sorted(user_tare.items()))).encode()).hexdigest()[:10]

# ---- automatic plan (before the user's edits) --------------------------------
auto_bulk = build_bulk_by_color(data, start_ctn=1, caps=caps, excess_pct=excess_pct) if detected["Bulk"] else []
auto_single = build_single_by_color(data, start_ctn=1, max_pp=int(single_max_pp) or None, caps=caps) if detected["Single"] else []

kit.step(3, "Carton tables",
         "Bulk and Single: the full packing table is editable — change any cell, add a row with +, "
         "tick 🗑 on rows and press “Remove ticked rows”, or “Reset table to PO”. "
         "Multi Y and Multi N follow the PO rules and are shown read-only. "
         "Changing the variables above rebuilds the tables.")
totals_box = st.container()

tab_labels, tab_types = [], []
for ptype in PACK_TYPES:
    if detected[ptype]:
        tab_labels.append(f"{ptype} ({len(detected[ptype])})")
        tab_types.append(ptype)
tab_labels += ["Raw tables", "JSON"]
tabs = st.tabs(tab_labels)
tab_of = {p: tabs[i] for i, p in enumerate(tab_types)}

STRIP = {"Bulk": "BULK PACK", "Single": "SINGLE PACK", "Multi Y": "MULTY PACK (Y)", "Multi N": "MULTY PACK (N)"}


def _editable_sections(ptype, autos, next_ctn):
    """Main packing table: fully editable via st.data_editor (direct, no cache tricks)."""
    out = []

    def _to_int(x, default=0):
        if x is None:
            return default
        try:
            if isinstance(x, float) and pd.isna(x):
                return default
        except Exception:
            pass
        s = str(x).strip().lower()
        if s in ("", "none", "nan", "nat"):
            return default
        try:
            return int(float(x))
        except Exception:
            return default

    for si, base in enumerate(autos):
        render_section_strip(STRIP[ptype])
        color = base.get("color", "")
        st.markdown(f"**{color}**  ·  {base.get('style_full', '')}")
        sizes = list(base.get("sizes") or [])

        # ---- build rows from real data ----
        rows_out = []
        for r in base.get("rows") or []:
            ctn_from = r.get("ctn_from", "")
            ctn_to = r.get("ctn_to", ctn_from)
            if ctn_to not in ("", None) and ctn_to != ctn_from:
                ctn_no = f"{ctn_from}-{ctn_to}"
            else:
                ctn_no = str(ctn_from) if ctn_from not in ("", None) else ""
            ctn_qty = _to_int(r.get("ctn_qty"), 0)
            qty_per = _to_int(r.get("qty_per_ctn"), 0)
            total_qty = _to_int(r.get("total_qty"), ctn_qty * qty_per)
            rec = {
                "🗑": False,
                "CTN NO": ctn_no,
                "CTN MES": str(r.get("ctn_mes") or ""),
                "CTN QTY": ctn_qty,
                "COLOR": str(r.get("color") or color or ""),
                "PREPACK STECKER": str(r.get("prepack") or base.get("prepack") or ""),
            }
            for s in sizes:
                if ptype in ("Bulk", "Single"):
                    rec[s] = _to_int(r.get("size_val"), 0) if s == r.get("size") else None
                else:
                    sv = (r.get("size_vals") or base.get("units") or {}).get(s)
                    rec[s] = _to_int(sv, 0) if sv not in (None, "", 0) else None
            rec["QTY PER CTN"] = qty_per
            rec["TOTAL QTY"] = total_qty
            gw, nw = r.get("gross"), r.get("net")
            try:
                rec["GROSS WEIGHT"] = float(gw) if gw not in (None, "") else None
            except Exception:
                rec["GROSS WEIGHT"] = None
            try:
                rec["NET WEIGHT"] = float(nw) if nw not in (None, "") else None
            except Exception:
                rec["NET WEIGHT"] = None
            rows_out.append(rec)

        cols = ["🗑", "CTN NO", "CTN MES", "CTN QTY", "COLOR", "PREPACK STECKER"] + sizes + [
            "QTY PER CTN", "TOTAL QTY", "GROSS WEIGHT", "NET WEIGHT"
        ]
        df = pd.DataFrame(rows_out, columns=cols)

        # Session state for this grid (simple, always editable)
        state_key = f"gap_edit_{ptype}_{si}_{color}_{sig}"
        data_key = f"{state_key}__data"
        nonce_key = f"{state_key}__nonce"
        if nonce_key not in st.session_state:
            st.session_state[nonce_key] = 0
        if data_key not in st.session_state:
            st.session_state[data_key] = df.copy()

        # Editable inputs vs calculated outputs
        # QTY PER CTN = sum of size columns (XS+S+M+…)
        # TOTAL QTY   = CTN QTY × QTY PER CTN
        # GROSS/NET   = from size NW + empty carton weight
        CALC_COLS = ["QTY PER CTN", "TOTAL QTY", "GROSS WEIGHT", "NET WEIGHT"]
        cfg = {
            "🗑": st.column_config.CheckboxColumn("🗑", default=False, width="small",
                                                  help="Tick to delete row"),
            "CTN NO": st.column_config.TextColumn("CTN NO"),
            "CTN MES": st.column_config.TextColumn("CTN MES"),
            "CTN QTY": st.column_config.NumberColumn("CTN QTY", min_value=0, step=1, format="%d"),
            "COLOR": st.column_config.TextColumn("COLOR"),
            "PREPACK STECKER": st.column_config.TextColumn("PREPACK STECKER"),
            "QTY PER CTN": st.column_config.NumberColumn(
                "QTY PER CTN", format="%d",
                help="Auto: sum of size columns (XS+S+M+…)"),
            "TOTAL QTY": st.column_config.NumberColumn(
                "TOTAL QTY", format="%d",
                help="Auto: CTN QTY × QTY PER CTN"),
            "GROSS WEIGHT": st.column_config.NumberColumn(
                "GROSS WEIGHT", format="%.2f",
                help="Auto: pcs × N.W + empty carton wt"),
            "NET WEIGHT": st.column_config.NumberColumn(
                "NET WEIGHT", format="%.2f",
                help="Auto: gross − net deduct"),
        }
        for s in sizes:
            cfg[s] = st.column_config.NumberColumn(s, min_value=0, step=1, format="%d")

        edited = st.data_editor(
            st.session_state[data_key],
            key=f"{state_key}__ed{st.session_state[nonce_key]}",
            num_rows="dynamic",
            hide_index=True,
            column_config=cfg,
            # calculated columns are read-only; user edits CTN QTY + sizes
            disabled=CALC_COLS,
            use_container_width=True,
        )

        # ---- Always recompute calculated columns from CTN QTY + size cells ----
        def _apply_formulas(frame: pd.DataFrame) -> pd.DataFrame:
            out = frame.copy()
            for idx, row in out.iterrows():
                size = None
                size_sum = 0
                for s in sizes:
                    v = _to_int(row.get(s), 0)
                    size_sum += v
                    if v > 0 and size is None:
                        size = s
                ctn = _to_int(row.get("CTN QTY"), 0)
                # FORMULA: QTY PER CTN = sum of all size columns
                qty_per = size_sum
                # FORMULA: TOTAL QTY = CTN QTY × QTY PER CTN
                total_qty = ctn * qty_per
                code = str(row.get("CTN MES") or "").strip()
                if not code and size and qty_per > 0:
                    try:
                        code = carton_code(size, qty_per)
                    except Exception:
                        code = "G81"
                gross = net = None
                if size and qty_per > 0:
                    try:
                        gross, net = carton_weights(size, qty_per, code or "G81", ptype)
                    except Exception:
                        pass
                out.at[idx, "QTY PER CTN"] = int(qty_per)
                out.at[idx, "TOTAL QTY"] = int(total_qty)
                if gross is not None:
                    out.at[idx, "GROSS WEIGHT"] = float(gross)
                if net is not None:
                    out.at[idx, "NET WEIGHT"] = float(net)
                if code:
                    out.at[idx, "CTN MES"] = code
            return out

        recalc = _apply_formulas(edited)

        # Detect if calculated values differ from what the grid is showing
        def _sig(frame):
            parts = []
            for _, row in frame.iterrows():
                parts.append((
                    _to_int(row.get("CTN QTY"), 0),
                    tuple(_to_int(row.get(s), 0) for s in sizes),
                    _to_int(row.get("QTY PER CTN"), 0),
                    _to_int(row.get("TOTAL QTY"), 0),
                ))
            return tuple(parts)

        if _sig(recalc) != _sig(edited):
            st.session_state[data_key] = recalc.reset_index(drop=True)
            st.session_state[nonce_key] = int(st.session_state.get(nonce_key, 0)) + 1
            st.rerun()
        # keep data_key synced even when sig matches (e.g. after stable load)
        st.session_state[data_key] = recalc.reset_index(drop=True)

        b1, b2, b3 = st.columns([1.4, 1.4, 4])
        n_tick = int(recalc["🗑"].fillna(False).astype(bool).sum()) if "🗑" in recalc.columns else 0
        if b1.button(f"🗑 Remove ticked rows ({n_tick})", key=f"{state_key}__rm", disabled=n_tick == 0):
            keep = recalc.loc[~recalc["🗑"].fillna(False).astype(bool)].copy()
            keep["🗑"] = False
            st.session_state[data_key] = keep.reset_index(drop=True)
            st.session_state[nonce_key] += 1
            st.rerun()
        if b2.button("↺ Reset table to PO", key=f"{state_key}__rst"):
            st.session_state[data_key] = df.copy()
            st.session_state[nonce_key] += 1
            st.rerun()

        # ---- build section from recalculated rows ----
        recs = []
        for _, row in recalc.iterrows():
            if bool(row.get("🗑", False)):
                continue
            size = None
            size_val = 0
            size_sum = 0
            for s in sizes:
                v = _to_int(row.get(s), 0)
                size_sum += v
                if v > 0 and size is None:
                    size, size_val = s, v
            ctn = _to_int(row.get("CTN QTY"), 0)
            pcs = size_sum if size_sum > 0 else _to_int(row.get("QTY PER CTN"), 0)
            if pcs <= 0 and size_val <= 0 and ctn <= 0:
                continue
            if not size:
                size = sizes[0] if sizes else ""
            pcs = pcs or size_val
            ctn = ctn or 1
            recs.append({
                "size": size,
                "pcs": pcs,
                "ctn": ctn,
                "prepack": str(row.get("PREPACK STECKER") or base.get("prepack") or ""),
            })

        sec, probs = rebuild_one_size_section(base, recs, next_ctn)
        kit.problems(probs)
        next_ctn += sec["ctn_total"]
        out.append(sec)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ship qty", f"{sec['ship_qty']} PCS")
        m2.metric("Cartons", sec["ctn_total"])
        m3.metric("Order qty", f"{sec['qty_ordered']} PCS")
        diff = sec["ship_qty"] - sec["qty_ordered"]
        m4.metric("Excess / short", f"{diff:+d}")
        if ptype == "Bulk":
            order = {
                sz: int(sec["units"].get(sz, 0)) - int(sec["excess_by_size"].get(sz, 0))
                for sz in sec["sizes"]
            }
            kit.check_by_size(list(sec["sizes"]), order, sec["units"], over_pct=float(excess_pct))
        elif diff == 0:
            st.success("Ship qty matches the PO.")
        elif diff > 0:
            st.warning(f"Ship qty is {diff} pcs above the PO.")
        else:
            st.error(f"Ship qty is {-diff} pcs below the PO.")
        st.divider()
    return out, next_ctn



next_ctn = 1
bulk_secs, single_secs = [], []
if auto_bulk:
    with tab_of["Bulk"]:
        bulk_secs, next_ctn = _editable_sections("Bulk", auto_bulk, next_ctn)
if auto_single:
    with tab_of["Single"]:
        single_secs, next_ctn = _editable_sections("Single", auto_single, next_ctn)

multi_y_secs = build_multi_y_by_color(data, start_ctn=next_ctn) if detected["Multi Y"] else []
next_ctn += sum(s["ctn_total"] for s in multi_y_secs)

multi_n_secs = []
multi_n_xlsx = None
if detected["Multi N"]:
    try:
        multi_n_secs = build_multi_n(
            data, nw=DEFAULT_NW, tare_by_code=DEFAULT_TARE, net_deduct=net_deduct,
            default_per_ctn=prepacks_per_ctn, ctn_code=ctn_code, start_ctn_no=next_ctn,
        )
        grand = next_ctn - 1 + sum(s.get("ctn_total", 0) for s in multi_n_secs)
        multi_n_xlsx = export_multi_n_excel(
            multi_n_secs, DEFAULT_NW, DEFAULT_NNW, DEFAULT_TARE, net_deduct,
            list(DEFAULT_CTN_MEAS), buyer, po_no, grand,
        )
    except Exception as e:
        st.error(f"Multi N build failed: {e}")

all_secs = bulk_secs + single_secs + multi_y_secs + multi_n_secs
grand_ctn = sum(s.get("ctn_total", 0) for s in all_secs)
grand_pcs = sum(s.get("ship_qty", 0) for s in all_secs)

for ptype, secs in (("Multi Y", multi_y_secs), ("Multi N", multi_n_secs)):
    if not secs:
        continue
    with tab_of[ptype]:
        for sec in secs:
            render_section_strip(STRIP[ptype])
            render_po_header_html(sec, buyer, po_no, grand_ctn)
            if ptype == "Multi N":
                sizes = sec.get("sizes") or []
                recs = []
                for r in sec.get("rows") or []:
                    rec = {
                        "CTN NO": (f"{r['ctn_from']}-{r['ctn_to']}" if r.get("ctn_to") != r.get("ctn_from") else str(r.get("ctn_from"))),
                        "CTN MES": r.get("ctn_mes", ""), "CTN QTY": r.get("ctn_qty"),
                        "COLOR": sec.get("color"), "PREPACK": sec.get("prepack"),
                    }
                    for sz in sizes:
                        rec[sz] = (sec.get("units") or {}).get(sz, "")
                    rec["PER BLST"] = r.get("per_blst", "")
                    rec["QTY PER CTN"] = r.get("qty_per_ctn", "")
                    rec["TOTAL QTY"] = r.get("total_qty", "")
                    rec["GROSS"] = r.get("gross", "")
                    rec["NET"] = r.get("net", "")
                    recs.append(rec)
                st.dataframe(pd.DataFrame(recs), hide_index=True, **kit.STRETCH)
            else:
                render_plist_table(sec)
            c1, c2, c3 = st.columns(3)
            c1.metric("Ship qty", f"{sec['ship_qty']} PCS")
            c2.metric("Cartons", sec["ctn_total"])
            c3.metric("Color", sec.get("color", ""))
            st.divider()

with totals_box:
    exs = grand_pcs - oiqty
    kit.summary_cards([("Total cartons", f"{grand_ctn:,}"), ("Total pieces", f"{grand_pcs:,}"),
                       ("Order qty", f"{oiqty}"), ("Over / short", f"{exs:+,}")])

with tabs[len(tab_types)]:
    for tname, table in (data.get("tables") or {}).items():
        ptype = detect_table_type(table) or "?"
        rows = table.get("rows") or []
        with st.expander(f"{tname} — {ptype} — {len(rows)} rows"):
            if rows:
                st.dataframe(rows, **kit.STRETCH)
            if table.get("total"):
                st.json(table["total"])

with tabs[len(tab_types) + 1]:
    st.json(data)
    st.download_button(
        "⬇ Download extracted JSON",
        data=json.dumps(data, indent=2, ensure_ascii=False),
        file_name=f"{po_no or Path(source_name).stem}.json",
        mime="application/json",
    )

# ---- 4 check ----------------------------------------------------------------------
kit.step(4, "Check against the PO")
if grand_pcs == oiqty:
    st.success(f"All sections together ship {grand_pcs:,} pcs = the PO order quantity.")
elif grand_pcs > oiqty:
    st.warning(f"All sections together ship {grand_pcs - oiqty:,} pcs above the PO ({grand_pcs:,} vs {oiqty}). "
               f"Bulk may carry up to {excess_pct:g}% excess per size.")
else:
    st.error(f"All sections together ship {oiqty - grand_pcs:,} pcs below the PO ({grand_pcs:,} vs {oiqty}).")

# ---- 5 download ---------------------------------------------------------------------
kit.step(5, "Download", "The Excel is rebuilt on every change, so it always has your latest edits.")
try:
    packing_all = export_unified_excel(
        data, bulk_secs, single_secs, multi_y_secs, multi_n_secs,
        nw=user_nw, nnw=user_nnw, tare_by_code=user_tare,
    )
except Exception as e:
    packing_all = None
    st.error(f"Excel build failed: {e}")

dl1, dl2 = st.columns(2)
with dl1:
    if packing_all:
        kit.download_block(
            "⬇ Download packing list for all",
            packing_all,
            f"packing_list_ALL_{po_no or 'PO'}.xlsx",
            f"Single sheet · Bulk+Single+Multi · {grand_ctn:,} ctns · {grand_pcs:,} pcs · formulas",
        )
with dl2:
    if multi_n_xlsx:
        kit.download_block(
            "⬇ Multi N template Excel only",
            multi_n_xlsx,
            f"packing_list_multi_n_{po_no}.xlsx",
        )
    else:
        st.caption("Multi N template appears when Multi N tables are present.")