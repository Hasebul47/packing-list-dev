# """
# JCPenney / Thereabouts — PO (.docx) to Packing List (.xlsx)
# ===========================================================
# Reads a JCPenney PO (.docx) and produces the factory's standard
# JCP PC import packing list (one sheet, one row per pack line).

# Run standalone:  streamlit run jcp_app.py
# """
# import io
# import re
# import zipfile
# from datetime import datetime

# import pandas as pd
# import streamlit as st
# from lxml import etree
# from openpyxl import Workbook
# from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# # ============================================================
# # Size mapping
# # ============================================================
# SIZE_LENGTH = {"07": "10.5", "08": "12.5", "09": "14.5",
#                "10": "16.5", "11": "18.5", "12": "20.5"}
# LENGTHS = ["10.5", "12.5", "14.5", "16.5", "18.5", "20.5"]

# DEFAULT_NETNET = 10.10
# DEFAULT_NET    = 10.80
# DEFAULT_GROSS  = 12.60
# DEFAULT_CBM    = 0.0632

# W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
# W_P  = f"{{{W_NS}}}p"
# W_T  = f"{{{W_NS}}}t"
# W_TBL = f"{{{W_NS}}}tbl"
# W_TR  = f"{{{W_NS}}}tr"
# W_TC  = f"{{{W_NS}}}tc"


# # ============================================================
# # 1. DOCX -> text  (raw XML walk, most reliable)
# # ============================================================
# def _para_text(el):
#     return "".join(t.text or "" for t in el.iter(W_T)).strip()


# def _table_lines(tbl):
#     out = []
#     for tr in tbl.iter(W_TR):
#         # only direct <w:tc> children of this <w:tr>
#         cells = [tc for tc in tr if tc.tag == W_TC]
#         row = []
#         for tc in cells:
#             txt = " ".join(_para_text(p) for p in tc.iter(W_P))
#             row.append(re.sub(r"\s+", " ", txt).strip())
#         # drop trailing empty cells
#         while row and not row[-1]:
#             row.pop()
#         if any(row):
#             out.append(" | ".join(row))
#     return out


# def read_docx_text(docx_bytes: bytes) -> str:
#     with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
#         xml = z.read("word/document.xml")
#     root = etree.fromstring(xml)
#     body = root.find(f"{{{W_NS}}}body")
#     lines = []
#     for el in body:
#         if el.tag == W_P:
#             t = _para_text(el)
#             if t:
#                 lines.append(t)
#         elif el.tag == W_TBL:
#             lines.extend(_table_lines(el))
#     return "\n".join(lines)


# # ============================================================
# # 2. Parse
# # ============================================================
# def _c(s):
#     return re.sub(r"\s+", " ", str(s or "")).strip()


# def _g(text, pat, default=""):
#     m = re.search(pat, text, re.I | re.S)
#     return _c(m.group(1)) if m else default


# def parse_po(text):
#     head = {
#         "item_number":  _g(text, r"Item\s+Number\s*[^A-Za-z0-9]{1,5}(\d+)"),
#         "sequence":     _g(text, r"Sequence\s+Number\s*[^A-Za-z0-9]{1,5}(\d+)"),
#         "order_no":     _g(text, r"Order#\s*[^0-9]{0,5}(\d{10,})"),
#         "brand":        _g(text, r"Brand\s*[^A-Za-z0-9]{0,5}([0-9A-Za-z\-]+)"),
#         "del_id":       _g(text, r"Del\s+ID\s*[^0-9]{0,5}(\d+)"),
#         "broker_id":    _g(text, r"BrokerID\s*[^A-Za-z0-9]{0,5}([0-9A-Za-z\-]+)"),
#         "suffix":       _g(text, r"Suffix\s*[^A-Za-z]{0,5}([A-Z]+)"),
#         "supply_chain": _g(text, r"Supply\s+Chain\s*[^A-Za-z0-9]{0,5}(\S+)"),
#         "tsn":          _g(text, r"TSN\s*[^0-9]{0,5}(\d+)"),
#         "retail_price": _g(text, r"Retail\s+Price\s*[^0-9$]{0,5}\$?\s*([\d.]+)"),
#         "carton_dims":  _g(text, r"Carton\s+Dimensions\s*[^A-Za-z0-9]{0,5}([^\n]+?)(?:\s+Carton\b|\n|$)"),
#         "carton_wt":    _g(text, r"Carton\s+Weight\s*[^0-9]{0,5}([\d.]+)"),
#         "style":        _g(text, r"Style#?\s*[^0-9]{0,3}(\d{6})"),
#         "dcd_date":     _g(text, r"(\d{4}-\d{2}-\d{2})"),
#         "ship_to":      _g(text, r"\b(9661-\d)\b"),
#     }

#     m = re.search(r"([\d,]{3,})\s+EA\b", text)
#     head["order_qty"] = int(m.group(1).replace(",", "")) if m else 0

#     # ---- PREPACK blocks ----
#     # find every "Pack Type ... PREPACK" occurrence
#     starts = [m.start() for m in re.finditer(r"Pack\s+Type", text, re.I)]
#     packs = []

#     for i, s in enumerate(starts):
#         # slice from this "Pack Type" to the next "Pack Type" (or end)
#         end = starts[i + 1] if i + 1 < len(starts) else len(text)
#         chunk = text[s:end]

#         # is it PREPACK?
#         if not re.search(r"PREPACK", chunk, re.I):
#             continue

#         # pieces per carton
#         mc = re.search(
#             r"No\.\s*of\s+Mfg\s+units\s+in\s+a\s+Carton[^0-9]{0,5}(\d+)",
#             chunk, re.I)
#         pce_per_ctn = int(mc.group(1)) if mc else 0

#         # data row after the table header
#         # looking for:  total_pkg | order_line | prepack_code | color | pp# | size_group
#         dm = None
#         for pat in (
#             # 6 cells joined by " | "
#             r"(\d{1,5})\s*\|\s*(\d{3,4})\s*\|\s*([*A-Z]+)\s*\|\s*"
#             r"([A-Z][A-Z ]+?)\s*\|\s*(\d{6})\s*\|\s*(\d{4})",
#             # space-separated fallback
#             r"(\d{1,5})\s+(\d{3,4})\s+([*A-Z]+)\s+"
#             r"([A-Z][A-Z ]+?)\s+(\d{6})\s+(\d{4})",
#         ):
#             dm = re.search(pat, chunk, re.S)
#             if dm:
#                 break
#         if not dm:
#             continue

#         # size ratios: rows like "08 | 12.5 | 1" or "08 12.5 1"
#         sizes = []
#         for sm in re.finditer(
#             r"\b(0[7-9]|1[0-2])\b\s*[\|]?\s*(\d{1,2}(?:\.\d+)?)\s*[\|]?\s*(\d{1,3})\b",
#             chunk
#         ):
#             sz = sm.group(1)
#             if any(x["size"] == sz for x in sizes):
#                 continue
#             sizes.append({"size": sz, "length": sm.group(2),
#                           "ratio": int(sm.group(3))})
#             if len(sizes) >= 6:
#                 break
#         if len(sizes) < 3:
#             continue

#         packs.append({
#             "pce_per_ctn":  pce_per_ctn,
#             "total_pkg":    int(dm.group(1)),
#             "order_line":   dm.group(2),
#             "prepack_code": dm.group(3),
#             "color":        _c(dm.group(4)),
#             "pp":           dm.group(5),
#             "size_group":   dm.group(6),
#             "sizes":        sizes,
#         })

#     return head, packs


# # ============================================================
# # 3. XLSX
# # ============================================================
# FONT = "Calibri"
# thin = Side(style="thin", color="000000")
# _box = Border(left=thin, right=thin, top=thin, bottom=thin)
# _CEN = Alignment(horizontal="center", vertical="center")
# _LEFT = Alignment(horizontal="left",   vertical="center")
# # _bold = Font(name=FONT, size=10, bold=True)
# # _norm = Font(name=FONT, size=10)
# # _head = Font(name=FONT, size=12, bold=True)
# # _title = Font(name=FONT, size=14, bold=True)
# _bold = Font(name=FONT, size=10, bold=True)
# _norm = Font(name=FONT, size=10)
# _head = Font(name=FONT, size=12, bold=True)
# _title = Font(name=FONT, size=14, bold=True)
# _yellow = PatternFill("solid", fgColor="FFFF00")   # <-- এই লাইন যোগ করুন

# def _put(ws, r, c, v=None, font=_norm, al=_CEN, fmt=None, fill=None):
#     cell = ws.cell(row=r, column=c)
#     if v is not None:
#         cell.value = v
#     cell.font = font
#     cell.alignment = al
#     if fmt:
#         cell.number_format = fmt
#     if fill:
#         cell.fill = fill
#     return cell


# def _merge(ws, r1, c1, r2, c2):
#     if (r1, c1) != (r2, c2):
#         ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


# # def build_xlsx(head, packs, netnet, net, gross, cbm):
# #     wb = Workbook()
# #     ws = wb.active
# #     ws.title = "Sheet1"

# #     widths = {"A": 8, "B": 8, "C": 18, "D": 10,
# #               "E": 7, "F": 7, "G": 7, "H": 7, "I": 7, "J": 7,
# #               "K": 10, "L": 12, "M": 12, "N": 12, "O": 12,
# #               "P": 10, "Q": 10, "R": 10, "S": 10}
# #     for col, w in widths.items():
# #         ws.column_dimensions[col].width = w

# #     _merge(ws, 1, 1, 1, 19)
# #     _put(ws, 1, 1,
# #          "JCP PURCHASING CORPORATION, (JCPPC) IMPORT PACKING LIST",
# #          _title, _LEFT)

# #     _merge(ws, 3, 1, 3, 6)
# #     _put(ws, 3, 1,
# #          f"JCPPC SUB/CONTRACT/SUFFIX : {head.get('del_id','')}/"
# #          f"{head.get('order_no','')}{head.get('suffix','')}",
# #          _bold, _LEFT)
# #     _merge(ws, 3, 17, 3, 19)
# #     _put(ws, 3, 17, "PACKAGE TYPE        : 0         RETAIL", _norm, _LEFT)

# #     _merge(ws, 5, 1, 5, 6)
# #     _put(ws, 5, 1, "DESTINATION      : 9661-0, 5011 FIRESTORE PL", _norm, _LEFT)
# #     _merge(ws, 6, 1, 6, 6)
# #     _put(ws, 6, 1,
# #          "                              SOUTH GATE CA90280 U.S.A. BY SEA.",
# #          _norm, _LEFT)
# #     _merge(ws, 5, 12, 5, 19)
# #     _put(ws, 5, 12,
# #          "UCC/EAN MANUFACTURER/LABELER INDENTIFICATION NUMBER : 764746",
# #          _norm, _LEFT)

# #     _merge(ws, 8, 1, 8, 4)
# #     _put(ws, 8, 1, f"STYLE# {head.get('style','')}", _head, _LEFT)

# #     _merge(ws, 9, 1, 9, 4)
# #     _put(ws, 9, 1, f"PO# {head.get('order_no','')}", _norm, _LEFT)
# #     _merge(ws, 9, 5, 9, 8)
# #     _put(ws, 9, 5, f"DI# {head.get('del_id','')}", _norm, _LEFT)
# #     dcd = head.get("dcd_date", "")
# #     try:
# #         dcd_fmt = datetime.strptime(dcd, "%Y-%m-%d").strftime("%m/%d/%y")
# #     except Exception:
# #         dcd_fmt = dcd
# #     _merge(ws, 9, 12, 9, 19)
# #     _put(ws, 9, 12, f"DCD DATE : {dcd_fmt}", _norm, _LEFT)

# #     # Header
# #     R_HDR, R_HDR2 = 11, 12

# #     _merge(ws, R_HDR, 1, R_HDR, 4)
# #     _put(ws, R_HDR, 1, "PREPACK", _bold)
# #     _merge(ws, R_HDR, 5, R_HDR, 10)
# #     _put(ws, R_HDR, 5, "SIZE RATIO", _bold)

# #     for i, (h1, h2, c) in enumerate([
# #         ("PCE PER",   "CARTON",       11),
# #         ("NET NET",   "WEIGHT/CTN",   12),
# #         ("NET WEIGHT","/CARTON",      13),
# #         ("GR.WGT.",   "/CARTON",      14),
# #         ("CBMs /",    "",             15),
# #         ("TOTAL",     "PCS.",         16),
# #         ("TOTAL",     "CARTONS",      17),
# #         ("FIRST",     "CARTON#",      18),
# #         ("LAST",      "CARTON#",      19),
# #     ]):
# #         _merge(ws, R_HDR, c, R_HDR2, c)
# #         _put(ws, R_HDR, c, h1, _bold)
# #         if h2:
# #             _put(ws, R_HDR2, c, h2, _bold)

# #     _put(ws, R_HDR2, 1, "LOT #", _bold)
# #     _put(ws, R_HDR2, 2, "LINE", _bold)
# #     _put(ws, R_HDR2, 3, "COLOR", _bold)
# #     _put(ws, R_HDR2, 4, "CODE", _bold)
# #     for i, ln in enumerate(LENGTHS):
# #         _put(ws, R_HDR2, 5 + i, ln, _bold)

# #     for c in range(1, 20):
# #         for rr in (R_HDR, R_HDR2):
# #             ws.cell(row=rr, column=c).border = _box

# #     # Data
# #     r = R_HDR2 + 1
# #     first_ctn = 1
# #     sum_pcs = sum_ctn = 0
# #     sum_netnet = sum_net = sum_gross = sum_cbm = 0.0

# #     for p in packs:
# #         pcs_per_ctn = p["pce_per_ctn"]
# #         ctn_qty = p["total_pkg"]
# #         pcs = pcs_per_ctn * ctn_qty

# #         nnet_line = round(ctn_qty * netnet, 1)
# #         net_line  = round(ctn_qty * net,    1)
# #         gross_ln  = round(ctn_qty * gross,  1)
# #         cbm_line  = round(ctn_qty * cbm,    6)

# #         _put(ws, r, 1, head.get("item_number", ""))
# #         _put(ws, r, 2, p["order_line"])
# #         _put(ws, r, 3, p["color"])
# #         _put(ws, r, 4, p["prepack_code"])

# #         ratio_by_len = {ln: 0 for ln in LENGTHS}
# #         for s in p["sizes"]:
# #             ln = SIZE_LENGTH.get(s["size"]) or s["length"]
# #             if ln in ratio_by_len:
# #                 ratio_by_len[ln] = s["ratio"]
# #         for i, ln in enumerate(LENGTHS):
# #             _put(ws, r, 5 + i, ratio_by_len[ln], fmt="0")

# #         _put(ws, r, 11, pcs_per_ctn, fmt="0")
# #         _put(ws, r, 12, nnet_line, fmt="0.0")
# #         _put(ws, r, 13, net_line,  fmt="0.0")
# #         _put(ws, r, 14, gross_ln,  fmt="0.0")
# #         _put(ws, r, 15, cbm_line,  fmt="0.000000")
# #         _put(ws, r, 16, pcs,       fmt="0")
# #         _put(ws, r, 17, ctn_qty,   fmt="0")
# #         _put(ws, r, 18, first_ctn, fmt="0")
# #         _put(ws, r, 19, first_ctn + ctn_qty - 1, fmt="0")

# #         for c in range(1, 20):
# #             ws.cell(row=r, column=c).border = _box

# #         first_ctn += ctn_qty
# #         sum_pcs += pcs; sum_ctn += ctn_qty
# #         sum_netnet += nnet_line; sum_net += net_line
# #         sum_gross += gross_ln; sum_cbm += cbm_line
# #         r += 1

# #     _put(ws, r, 11, "TOTAL", _bold)
# #     _put(ws, r, 12, round(sum_netnet, 1), _bold, fmt="0.0")
# #     _put(ws, r, 13, round(sum_net,    1), _bold, fmt="0.0")
# #     _put(ws, r, 14, round(sum_gross,  1), _bold, fmt="0.0")
# #     _put(ws, r, 15, round(sum_cbm,    6), _bold, fmt="0.000000")
# #     _put(ws, r, 16, sum_pcs, _bold, fmt="0")
# #     _put(ws, r, 17, sum_ctn, _bold, fmt="0")
# #     for c in range(1, 20):
# #         ws.cell(row=r, column=c).border = _box
# #     r += 2

# #     _put(ws, r, 1, "CARTON MEASUREMENT : 24.5 X 21 X7.5 INCH-18 PCS", _norm, _LEFT)
# #     r += 2
# #     _put(ws, r, 1, "THAT'S IT GARMENTS LTD.", _bold, _LEFT)
# #     _put(ws, r, 11, "CARTON TOTAL :", _bold, _LEFT)
# #     r += 1
# #     _put(ws, r, 1, "NORASHING PUR , ASHULIA", _norm, _LEFT)
# #     _put(ws, r, 11, "CARTONS     :", _bold, _LEFT); _put(ws, r, 13, sum_ctn)
# #     r += 1
# #     _put(ws, r, 1, "DHAKA, BANGLADESH.", _norm, _LEFT)
# #     _put(ws, r, 11, "PICEES         :", _bold, _LEFT); _put(ws, r, 13, sum_pcs)
# #     r += 1
# #     _put(ws, r, 1, "TEL : 02-880-8811149, 8823448, FAX : 02-880-8826090", _norm, _LEFT)
# #     _put(ws, r, 11, "NET NET-KG:", _bold, _LEFT); _put(ws, r, 13, round(sum_netnet, 1))
# #     r += 1
# #     _put(ws, r, 1, "EMAIL : RUBEL@HAMEEMGROUP.COM", _norm, _LEFT)
# #     _put(ws, r, 11, "NET WT-KG  :", _bold, _LEFT); _put(ws, r, 13, round(sum_net, 1))
# #     r += 1
# #     _put(ws, r, 11, "GR.WT-KG    :", _bold, _LEFT); _put(ws, r, 13, round(sum_gross, 1))
# #     r += 1
# #     _put(ws, r, 11, "CBM               :", _bold, _LEFT); _put(ws, r, 13, round(sum_cbm, 6))

# #     buf = io.BytesIO()
# #     wb.save(buf)
# #     return buf.getvalue()


# # def build_xlsx(head, packs, netnet, net, gross, cbm):
# #     wb = Workbook()
# #     ws = wb.active
# #     ws.title = "Sheet1"

# #     widths = {"A": 8, "B": 8, "C": 18, "D": 10,
# #               "E": 7, "F": 7, "G": 7, "H": 7, "I": 7, "J": 7,
# #               "K": 10, "L": 12, "M": 12, "N": 12, "O": 12,
# #               "P": 10, "Q": 10, "R": 10, "S": 10}
# #     for col, w in widths.items():
# #         ws.column_dimensions[col].width = w

# #     # ---- Title ----
# #     _merge(ws, 1, 1, 1, 19)
# #     _put(ws, 1, 1,
# #          "JCP PURCHASING CORPORATION, (JCPPC) IMPORT PACKING LIST",
# #          _title, _LEFT)

# #     # ---- Sub-contract / suffix ----
# #     _merge(ws, 3, 1, 3, 6)
# #     _put(ws, 3, 1,
# #          f"JCPPC SUB/CONTRACT/SUFFIX : {head.get('del_id','')}/"
# #          f"{head.get('order_no','')}{head.get('suffix','')}",
# #          _bold, _LEFT)
# #     _merge(ws, 3, 17, 3, 19)
# #     _put(ws, 3, 17, "PACKAGE TYPE        : 0         RETAIL", _norm, _LEFT)

# #     # ---- Destination + UCC ----
# #     _merge(ws, 5, 1, 5, 6)
# #     _put(ws, 5, 1, "DESTINATION      : 9661-0, 5011 FIRESTORE PL", _norm, _LEFT)
# #     _merge(ws, 6, 1, 6, 6)
# #     _put(ws, 6, 1,
# #          "                              SOUTH GATE CA90280 U.S.A. BY SEA.",
# #          _norm, _LEFT)
# #     _merge(ws, 5, 12, 5, 19)
# #     _put(ws, 5, 12,
# #          "UCC/EAN MANUFACTURER/LABELER INDENTIFICATION NUMBER : 764746",
# #          _norm, _LEFT)

# #     # ---- STYLE ----
# #     _merge(ws, 8, 1, 8, 4)
# #     _put(ws, 8, 1, f"STYLE# {head.get('style','')}", _head, _LEFT)

# #     # ---- PO# / DI# / DCD DATE ----
# #     _merge(ws, 9, 1, 9, 4)
# #     _put(ws, 9, 1, f"PO# {head.get('order_no','')}", _norm, _LEFT)
# #     _merge(ws, 9, 5, 9, 8)
# #     _put(ws, 9, 5, f"DI# {head.get('del_id','')}", _norm, _LEFT)
# #     dcd = head.get("dcd_date", "")
# #     try:
# #         dcd_fmt = datetime.strptime(dcd, "%Y-%m-%d").strftime("%m/%d/%y")
# #     except Exception:
# #         dcd_fmt = dcd
# #     _merge(ws, 9, 12, 9, 19)
# #     _put(ws, 9, 12, f"DCD DATE : {dcd_fmt}", _norm, _LEFT)

# #     # ============================================================
# #     # Header — TWO rows, NO vertical merges
# #     # ============================================================
# #     R_HDR, R_HDR2 = 11, 12

# #     # Row 1: PREPACK (A-D merged), SIZE RATIO (E-J merged), K..S single cells
# #     _merge(ws, R_HDR, 1, R_HDR, 4)
# #     _put(ws, R_HDR, 1, "PREPACK", _bold)

# #     _merge(ws, R_HDR, 5, R_HDR, 10)
# #     _put(ws, R_HDR, 5, "SIZE RATIO", _bold)

# #     for h1, c in [
# #         ("PCE PER",    11),
# #         ("NET NET",    12),
# #         ("NET WEIGHT", 13),
# #         ("GR.WGT.",    14),
# #         ("CBMs /",     15),
# #         ("TOTAL",      16),
# #         ("TOTAL",      17),
# #         ("FIRST",      18),
# #         ("LAST",       19),
# #     ]:
# #         _put(ws, R_HDR, c, h1, _bold)

# #     # Row 2: labels under each column
# #     _put(ws, R_HDR2, 1, "LOT #", _bold)
# #     _put(ws, R_HDR2, 2, "LINE",  _bold)
# #     _put(ws, R_HDR2, 3, "COLOR", _bold)
# #     _put(ws, R_HDR2, 4, "CODE",  _bold)
# #     for i, ln in enumerate(LENGTHS):
# #         _put(ws, R_HDR2, 5 + i, ln, _bold)

# #     for h2, c in [
# #         ("CARTON",      11),
# #         ("WEIGHT/CTN",  12),
# #         ("/CARTON",     13),
# #         ("/CARTON",     14),
# #         ("",            15),
# #         ("PCS.",        16),
# #         ("CARTONS",     17),
# #         ("CARTON#",     18),
# #         ("CARTON#",     19),
# #     ]:
# #         if h2:
# #             _put(ws, R_HDR2, c, h2, _bold)

# #     # borders on both header rows
# #     for c in range(1, 20):
# #         for rr in (R_HDR, R_HDR2):
# #             ws.cell(row=rr, column=c).border = _box

# #     # ============================================================
# #     # Data rows
# #     # ============================================================
# #     r = R_HDR2 + 1
# #     first_ctn = 1
# #     sum_pcs = sum_ctn = 0
# #     sum_netnet = sum_net = sum_gross = sum_cbm = 0.0

# #     for p in packs:
# #         pcs_per_ctn = p["pce_per_ctn"]
# #         ctn_qty = p["total_pkg"]
# #         pcs = pcs_per_ctn * ctn_qty

# #         nnet_line = round(ctn_qty * netnet, 1)
# #         net_line  = round(ctn_qty * net,    1)
# #         gross_ln  = round(ctn_qty * gross,  1)
# #         cbm_line  = round(ctn_qty * cbm,    6)

# #         _put(ws, r, 1, head.get("item_number", ""))
# #         _put(ws, r, 2, p["order_line"])
# #         _put(ws, r, 3, p["color"])
# #         _put(ws, r, 4, p["prepack_code"])

# #         ratio_by_len = {ln: 0 for ln in LENGTHS}
# #         for s in p["sizes"]:
# #             ln = SIZE_LENGTH.get(s["size"]) or s["length"]
# #             if ln in ratio_by_len:
# #                 ratio_by_len[ln] = s["ratio"]
# #         for i, ln in enumerate(LENGTHS):
# #             _put(ws, r, 5 + i, ratio_by_len[ln], fmt="0")

# #         _put(ws, r, 11, pcs_per_ctn, fmt="0")
# #         _put(ws, r, 12, nnet_line, fmt="0.0")
# #         _put(ws, r, 13, net_line,  fmt="0.0")
# #         _put(ws, r, 14, gross_ln,  fmt="0.0")
# #         _put(ws, r, 15, cbm_line,  fmt="0.000000")
# #         _put(ws, r, 16, pcs,       fmt="0")
# #         _put(ws, r, 17, ctn_qty,   fmt="0")
# #         _put(ws, r, 18, first_ctn, fmt="0")
# #         _put(ws, r, 19, first_ctn + ctn_qty - 1, fmt="0")

# #         for c in range(1, 20):
# #             ws.cell(row=r, column=c).border = _box

# #         first_ctn += ctn_qty
# #         sum_pcs += pcs; sum_ctn += ctn_qty
# #         sum_netnet += nnet_line; sum_net += net_line
# #         sum_gross += gross_ln; sum_cbm += cbm_line
# #         r += 1

# #     # ---- Total row ----
# #     _put(ws, r, 11, "TOTAL", _bold)
# #     _put(ws, r, 12, round(sum_netnet, 1), _bold, fmt="0.0")
# #     _put(ws, r, 13, round(sum_net,    1), _bold, fmt="0.0")
# #     _put(ws, r, 14, round(sum_gross,  1), _bold, fmt="0.0")
# #     _put(ws, r, 15, round(sum_cbm,    6), _bold, fmt="0.000000")
# #     _put(ws, r, 16, sum_pcs, _bold, fmt="0")
# #     _put(ws, r, 17, sum_ctn, _bold, fmt="0")
# #     for c in range(1, 20):
# #         ws.cell(row=r, column=c).border = _box
# #     r += 2

# #     # ---- Footer ----
# #     _put(ws, r, 1, "CARTON MEASUREMENT : 24.5 X 21 X7.5 INCH-18 PCS", _norm, _LEFT)
# #     r += 2
# #     _put(ws, r, 1, "THAT'S IT GARMENTS LTD.", _bold, _LEFT)
# #     _put(ws, r, 11, "CARTON TOTAL :", _bold, _LEFT)
# #     r += 1
# #     _put(ws, r, 1, "NORASHING PUR , ASHULIA", _norm, _LEFT)
# #     _put(ws, r, 11, "CARTONS     :", _bold, _LEFT)
# #     _put(ws, r, 13, sum_ctn)
# #     r += 1
# #     _put(ws, r, 1, "DHAKA, BANGLADESH.", _norm, _LEFT)
# #     _put(ws, r, 11, "PICEES         :", _bold, _LEFT)
# #     _put(ws, r, 13, sum_pcs)
# #     r += 1
# #     _put(ws, r, 1,
# #          "TEL : 02-880-8811149, 8823448, FAX : 02-880-8826090", _norm, _LEFT)
# #     _put(ws, r, 11, "NET NET-KG:", _bold, _LEFT)
# #     _put(ws, r, 13, round(sum_netnet, 1))
# #     r += 1
# #     _put(ws, r, 1, "EMAIL : RUBEL@HAMEEMGROUP.COM", _norm, _LEFT)
# #     _put(ws, r, 11, "NET WT-KG  :", _bold, _LEFT)
# #     _put(ws, r, 13, round(sum_net, 1))
# #     r += 1
# #     _put(ws, r, 11, "GR.WT-KG    :", _bold, _LEFT)
# #     _put(ws, r, 13, round(sum_gross, 1))
# #     r += 1
# #     _put(ws, r, 11, "CBM               :", _bold, _LEFT)
# #     _put(ws, r, 13, round(sum_cbm, 6))

# #     buf = io.BytesIO()
# #     wb.save(buf)
# #     return buf.getvalue()

# def build_xlsx(head, packs, netnet, net, gross, cbm):
#     from openpyxl.utils import get_column_letter as L

#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Sheet1"

#     widths = {"A": 8, "B": 8, "C": 18, "D": 10,
#               "E": 7, "F": 7, "G": 7, "H": 7, "I": 7, "J": 7,
#               "K": 10, "L": 12, "M": 12, "N": 12, "O": 12,
#               "P": 10, "Q": 10, "R": 10, "S": 10}
#     for col, w in widths.items():
#         ws.column_dimensions[col].width = w

#     # ============================================================
#     # Title block
#     # ============================================================
#     _merge(ws, 1, 1, 1, 19)
#     _put(ws, 1, 1,
#          "JCP PURCHASING CORPORATION, (JCPPC) IMPORT PACKING LIST",
#          _title, _LEFT)

#     _merge(ws, 3, 1, 3, 6)
#     _put(ws, 3, 1,
#          f"JCPPC SUB/CONTRACT/SUFFIX : {head.get('del_id','')}/"
#          f"{head.get('order_no','')}{head.get('suffix','')}",
#          _bold, _LEFT)
#     _merge(ws, 3, 17, 3, 19)
#     _put(ws, 3, 17, "PACKAGE TYPE        : 0         RETAIL", _norm, _LEFT)

#     _merge(ws, 5, 1, 5, 6)
#     _put(ws, 5, 1, "DESTINATION      : 9661-0, 5011 FIRESTORE PL", _norm, _LEFT)
#     _merge(ws, 6, 1, 6, 6)
#     _put(ws, 6, 1,
#          "                              SOUTH GATE CA90280 U.S.A. BY SEA.",
#          _norm, _LEFT)
#     _merge(ws, 5, 12, 5, 19)
#     _put(ws, 5, 12,
#          "UCC/EAN MANUFACTURER/LABELER INDENTIFICATION NUMBER : 764746",
#          _norm, _LEFT)

#     _merge(ws, 8, 1, 8, 4)
#     _put(ws, 8, 1, f"STYLE# {head.get('style','')}", _head, _LEFT)

#     _merge(ws, 9, 1, 9, 4)
#     _put(ws, 9, 1, f"PO# {head.get('order_no','')}", _norm, _LEFT)
#     _merge(ws, 9, 5, 9, 8)
#     _put(ws, 9, 5, f"DI# {head.get('del_id','')}", _norm, _LEFT)
#     dcd = head.get("dcd_date", "")
#     try:
#         dcd_fmt = datetime.strptime(dcd, "%Y-%m-%d").strftime("%m/%d/%y")
#     except Exception:
#         dcd_fmt = dcd
#     _merge(ws, 9, 12, 9, 19)
#     _put(ws, 9, 12, f"DCD DATE : {dcd_fmt}", _norm, _LEFT)

#     # ============================================================
#     # Rate row  (row 10) — edit L10, M10, N10, O10 in Excel
#     # ============================================================
#     R_RATE = 10
#     _merge(ws, R_RATE, 1, R_RATE, 10)
#     _put(ws, R_RATE, 1,
#          "PER CARTON RATE  →   edit the yellow cells; every line recalculates",
#          _bold, _LEFT)
#     _put(ws, R_RATE, 12, round(netnet, 4), _bold, fmt="0.0000", fill=_yellow)
#     _put(ws, R_RATE, 13, round(net,    4), _bold, fmt="0.0000", fill=_yellow)
#     _put(ws, R_RATE, 14, round(gross,  4), _bold, fmt="0.0000", fill=_yellow)
#     _put(ws, R_RATE, 15, round(cbm,    6), _bold, fmt="0.000000", fill=_yellow)

#     # ============================================================
#     # Header (rows 11-12)
#     # ============================================================
#     R_HDR, R_HDR2 = 11, 12

#     _merge(ws, R_HDR, 1, R_HDR, 4)
#     _put(ws, R_HDR, 1, "PREPACK", _bold)

#     _merge(ws, R_HDR, 5, R_HDR, 10)
#     _put(ws, R_HDR, 5, "SIZE RATIO", _bold)

#     for h1, c in [
#         ("PCE PER",    11), ("NET NET",    12), ("NET WEIGHT", 13),
#         ("GR.WGT.",    14), ("CBMs /",     15), ("TOTAL",      16),
#         ("TOTAL",      17), ("FIRST",      18), ("LAST",       19),
#     ]:
#         _put(ws, R_HDR, c, h1, _bold)

#     _put(ws, R_HDR2, 1, "LOT #", _bold)
#     _put(ws, R_HDR2, 2, "LINE",  _bold)
#     _put(ws, R_HDR2, 3, "COLOR", _bold)
#     _put(ws, R_HDR2, 4, "CODE",  _bold)
#     for i, ln in enumerate(LENGTHS):
#         _put(ws, R_HDR2, 5 + i, ln, _bold)

#     for h2, c in [
#         ("CARTON",    11), ("WEIGHT/CTN", 12), ("/CARTON", 13),
#         ("/CARTON",   14), ("",           15), ("PCS.",    16),
#         ("CARTONS",   17), ("CARTON#",    18), ("CARTON#", 19),
#     ]:
#         if h2:
#             _put(ws, R_HDR2, c, h2, _bold)

#     for c in range(1, 20):
#         for rr in (R_HDR, R_HDR2):
#             ws.cell(row=rr, column=c).border = _box

#     # ============================================================
#     # Data rows — live formulas
#     # ============================================================
#     r = R_HDR2 + 1
#     data_first = r

#     for p in packs:
#         pcs_per_ctn = p["pce_per_ctn"]
#         ctn_qty = int(p["total_pkg"])

#         _put(ws, r, 1, head.get("item_number", ""))
#         _put(ws, r, 2, p["order_line"])
#         _put(ws, r, 3, p["color"])
#         _put(ws, r, 4, p["prepack_code"])

#         ratio_by_len = {ln: 0 for ln in LENGTHS}
#         for s in p["sizes"]:
#             ln = SIZE_LENGTH.get(s["size"]) or s["length"]
#             if ln in ratio_by_len:
#                 ratio_by_len[ln] = s["ratio"]
#         for i, ln in enumerate(LENGTHS):
#             _put(ws, r, 5 + i, ratio_by_len[ln], fmt="0")

#         # K = pcs per carton (static)
#         _put(ws, r, 11, pcs_per_ctn, fmt="0")

#         # L..O — formula referencing the rate row (row 10) × Q (cartons)
#         _put(ws, r, 12, f"=Q{r}*$L${R_RATE}", fmt="0.0000")
#         _put(ws, r, 13, f"=Q{r}*$M${R_RATE}", fmt="0.0000")
#         _put(ws, r, 14, f"=Q{r}*$N${R_RATE}", fmt="0.0000")
#         _put(ws, r, 15, f"=Q{r}*$O${R_RATE}", fmt="0.000000")

#         # P = K × Q (total pcs)
#         _put(ws, r, 16, f"=K{r}*Q{r}", fmt="0")

#         # Q = cartons (static)
#         _put(ws, r, 17, ctn_qty, fmt="0")

#         # R / S — first / last carton numbers
#         if r == data_first:
#             _put(ws, r, 18, 1, fmt="0")
#         else:
#             _put(ws, r, 18, f"=S{r-1}+1", fmt="0")
#         _put(ws, r, 19, f"=R{r}+Q{r}-1", fmt="0")

#         for c in range(1, 20):
#             ws.cell(row=r, column=c).border = _box

#         r += 1

#     data_last = r - 1

#     # ============================================================
#     # TOTAL row — SUM formulas
#     # ============================================================
#     total_row = r
#     _put(ws, r, 11, "TOTAL", _bold)
#     for col, fmt in ((12, "0.0000"), (13, "0.0000"), (14, "0.0000"),
#                      (15, "0.000000"), (16, "0"), (17, "0")):
#         cl = L(col)
#         _put(ws, r, col,
#              f"=SUM({cl}{data_first}:{cl}{data_last})",
#              _bold, fmt=fmt)
#     for c in range(1, 20):
#         ws.cell(row=r, column=c).border = _box
#     r += 2

#     # ============================================================
#     # Footer — references the total row
#     # ============================================================
#     _put(ws, r, 1, "CARTON MEASUREMENT : 24.5 X 21 X7.5 INCH-18 PCS",
#          _norm, _LEFT)
#     r += 2

#     _put(ws, r, 1, "THAT'S IT GARMENTS LTD.", _bold, _LEFT)
#     _put(ws, r, 11, "CARTON TOTAL :", _bold, _LEFT)
#     r += 1
#     _put(ws, r, 1, "NORASHING PUR , ASHULIA", _norm, _LEFT)
#     _put(ws, r, 11, "CARTONS     :", _bold, _LEFT)
#     _put(ws, r, 13, f"=Q{total_row}", fmt="0")
#     r += 1
#     _put(ws, r, 1, "DHAKA, BANGLADESH.", _norm, _LEFT)
#     _put(ws, r, 11, "PICEES         :", _bold, _LEFT)
#     _put(ws, r, 13, f"=P{total_row}", fmt="0")
#     r += 1
#     _put(ws, r, 1,
#          "TEL : 02-880-8811149, 8823448, FAX : 02-880-8826090", _norm, _LEFT)
#     _put(ws, r, 11, "NET NET-KG:", _bold, _LEFT)
#     _put(ws, r, 13, f"=L{total_row}", fmt="0.0000")
#     r += 1
#     _put(ws, r, 1, "EMAIL : RUBEL@HAMEEMGROUP.COM", _norm, _LEFT)
#     _put(ws, r, 11, "NET WT-KG  :", _bold, _LEFT)
#     _put(ws, r, 13, f"=M{total_row}", fmt="0.0000")
#     r += 1
#     _put(ws, r, 11, "GR.WT-KG    :", _bold, _LEFT)
#     _put(ws, r, 13, f"=N{total_row}", fmt="0.0000")
#     r += 1
#     _put(ws, r, 11, "CBM               :", _bold, _LEFT)
#     _put(ws, r, 13, f"=O{total_row}", fmt="0.000000")

#     # ---- Force Excel to recalculate all formulas on open ----
#     wb.calculation.fullCalcOnLoad = True

#     buf = io.BytesIO()
#     wb.save(buf)
#     return buf.getvalue()


# # ============================================================
# # 4. UI
# # ============================================================
# st.set_page_config(page_title="JCP PO -> Packing List",
#                    layout="wide", page_icon="📦")
# st.title("📦 JCPenney — PO to Packing List")
# st.caption("Upload a JCPenney PO (.docx) and download the JCP PC import packing list.")

# up = st.file_uploader("Upload PO .docx", type=["docx"])

# with st.sidebar:
#     st.markdown("### Carton weights (per carton)")
#     netnet = st.number_input("NET NET kg/carton", 0.0, 100.0,
#                              DEFAULT_NETNET, 0.01, format="%.2f")
#     net    = st.number_input("NET kg/carton",     0.0, 100.0,
#                              DEFAULT_NET,    0.01, format="%.2f")
#     gross  = st.number_input("GROSS kg/carton",   0.0, 100.0,
#                              DEFAULT_GROSS,  0.01, format="%.2f")
#     cbm    = st.number_input("CBM /carton",       0.0, 10.0,
#                              DEFAULT_CBM,    0.0001, format="%.4f")

# if up is None:
#     st.info("Upload a .docx to get started.")
#     st.stop()

# try:
#     text = read_docx_text(up.getvalue())
# except Exception as e:
#     st.error(f"Could not read the DOCX: {e}")
#     st.stop()

# # Show extracted text
# with st.expander("🔍 Extracted text from DOCX (first 250 lines)", expanded=False):
#     st.text("\n".join(text.splitlines()[:250]))

# try:
#     head, packs = parse_po(text)
# except Exception as e:
#     st.error(f"Parse failed: {e}")
#     st.stop()

# if not packs:
#     st.warning("No PREPACK blocks were found.")
#     with st.expander("Full extracted text", expanded=True):
#         st.code(text[:6000])
#     st.stop()

# st.success(f"PO **{head.get('order_no','')}** · "
#            f"Style **{head.get('style','')}** · {len(packs)} pack line(s)")

# rows = []
# for p in packs:
#     row = {"Order Line": p["order_line"], "Color": p["color"],
#            "Prepack Code": p["prepack_code"],
#            "Pce/Ctn": p["pce_per_ctn"], "Cartons": p["total_pkg"]}
#     for ln in LENGTHS:
#         ratio = 0
#         for s in p["sizes"]:
#             if SIZE_LENGTH.get(s["size"]) == ln:
#                 ratio = s["ratio"]
#         row[ln] = ratio
#     rows.append(row)
# st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# st.divider()
# if st.button("Prepare packing list (.xlsx)", type="primary"):
#     st.session_state["jcp_xlsx"] = build_xlsx(
#         head, packs, netnet, net, gross, cbm)

# if st.session_state.get("jcp_xlsx"):
#     st.download_button(
#         "⬇ Download packing list",
#         data=st.session_state["jcp_xlsx"],
#         file_name=f"packing_list_{head.get('order_no','')}.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     )


"""
JCPenney / Thereabouts — PO (.docx) to Packing List (.xlsx)
===========================================================
Reads a JCPenney PO (.docx) and produces the factory's standard
JCP PC import packing list (one sheet, one row per pack line).

Run standalone:  streamlit run jcp_app.py
"""
import io
import re
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L


# ============================================================
# Size mapping
# ============================================================
SIZE_LENGTH = {"07": "10.5", "08": "12.5", "09": "14.5",
               "10": "16.5", "11": "18.5", "12": "20.5"}
LENGTHS = ["10.5", "12.5", "14.5", "16.5", "18.5", "20.5"]

DEFAULT_NETNET = 10.10
DEFAULT_NET    = 10.80
DEFAULT_GROSS  = 12.60
DEFAULT_CBM    = 0.0632

W_NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P   = f"{{{W_NS}}}p"
W_T   = f"{{{W_NS}}}t"
W_TBL = f"{{{W_NS}}}tbl"
W_TR  = f"{{{W_NS}}}tr"
W_TC  = f"{{{W_NS}}}tc"


# ============================================================
# 1. DOCX -> text
# ============================================================
def _para_text(el):
    return "".join(t.text or "" for t in el.iter(W_T)).strip()


def _table_lines(tbl):
    out = []
    for tr in tbl.iter(W_TR):
        cells = [tc for tc in tr if tc.tag == W_TC]
        row = []
        for tc in cells:
            txt = " ".join(_para_text(p) for p in tc.iter(W_P))
            row.append(re.sub(r"\s+", " ", txt).strip())
        while row and not row[-1]:
            row.pop()
        if any(row):
            out.append(" | ".join(row))
    return out


def read_docx_text(docx_bytes: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{{{W_NS}}}body")
    lines = []
    for el in body:
        if el.tag == W_P:
            t = _para_text(el)
            if t:
                lines.append(t)
        elif el.tag == W_TBL:
            lines.extend(_table_lines(el))
    return "\n".join(lines)


# ============================================================
# 2. Parse
# ============================================================
def _c(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _g(text, pat, default=""):
    m = re.search(pat, text, re.I | re.S)
    return _c(m.group(1)) if m else default


def parse_po(text):
    head = {
        "item_number":  _g(text, r"Item\s+Number\s*[^A-Za-z0-9]{1,5}(\d+)"),
        "sequence":     _g(text, r"Sequence\s+Number\s*[^A-Za-z0-9]{1,5}(\d+)"),
        "order_no":     _g(text, r"Order#\s*[^0-9]{0,5}(\d{10,})"),
        "brand":        _g(text, r"Brand\s*[^A-Za-z0-9]{0,5}([0-9A-Za-z\-]+)"),
        "del_id":       _g(text, r"Del\s+ID\s*[^0-9]{0,5}(\d+)"),
        "broker_id":    _g(text, r"BrokerID\s*[^A-Za-z0-9]{0,5}([0-9A-Za-z\-]+)"),
        "suffix":       _g(text, r"Suffix\s*[^A-Za-z]{0,5}([A-Z]+)"),
        "supply_chain": _g(text, r"Supply\s+Chain\s*[^A-Za-z0-9]{0,5}(\S+)"),
        "tsn":          _g(text, r"TSN\s*[^0-9]{0,5}(\d+)"),
        "retail_price": _g(text, r"Retail\s+Price\s*[^0-9$]{0,5}\$?\s*([\d.]+)"),
        "carton_dims":  _g(text, r"Carton\s+Dimensions\s*[^A-Za-z0-9]{0,5}([^\n]+?)(?:\s+Carton\b|\n|$)"),
        "carton_wt":    _g(text, r"Carton\s+Weight\s*[^0-9]{0,5}([\d.]+)"),
        "style":        _g(text, r"Style#?\s*[^0-9]{0,3}(\d{6})"),
        "dcd_date":     _g(text, r"(\d{4}-\d{2}-\d{2})"),
        "ship_to":      _g(text, r"\b(9661-\d)\b"),
    }

    m = re.search(r"([\d,]{3,})\s+EA\b", text)
    head["order_qty"] = int(m.group(1).replace(",", "")) if m else 0

    starts = [m.start() for m in re.finditer(r"Pack\s+Type", text, re.I)]
    packs = []

    for i, s in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(text)
        chunk = text[s:end]

        if not re.search(r"PREPACK", chunk, re.I):
            continue

        mc = re.search(
            r"No\.\s*of\s+Mfg\s+units\s+in\s+a\s+Carton[^0-9]{0,5}(\d+)",
            chunk, re.I)
        pce_per_ctn = int(mc.group(1)) if mc else 0

        dm = None
        for pat in (
            r"(\d{1,5})\s*\|\s*(\d{3,4})\s*\|\s*([*A-Z]+)\s*\|\s*"
            r"([A-Z][A-Z ]+?)\s*\|\s*(\d{6})\s*\|\s*(\d{4})",
            r"(\d{1,5})\s+(\d{3,4})\s+([*A-Z]+)\s+"
            r"([A-Z][A-Z ]+?)\s+(\d{6})\s+(\d{4})",
        ):
            dm = re.search(pat, chunk, re.S)
            if dm:
                break
        if not dm:
            continue

        sizes = []
        for sm in re.finditer(
            r"\b(0[7-9]|1[0-2])\b\s*[\|]?\s*(\d{1,2}(?:\.\d+)?)\s*[\|]?\s*(\d{1,3})\b",
            chunk
        ):
            sz = sm.group(1)
            if any(x["size"] == sz for x in sizes):
                continue
            sizes.append({"size": sz, "length": sm.group(2),
                          "ratio": int(sm.group(3))})
            if len(sizes) >= 6:
                break
        if len(sizes) < 3:
            continue

        packs.append({
            "pce_per_ctn":  pce_per_ctn,
            "total_pkg":    int(dm.group(1)),
            "order_line":   dm.group(2),
            "prepack_code": dm.group(3),
            "color":        _c(dm.group(4)),
            "pp":           dm.group(5),
            "size_group":   dm.group(6),
            "sizes":        sizes,
        })

    return head, packs


# ============================================================
# 3. XLSX
# ============================================================
FONT = "Calibri"
thin = Side(style="thin", color="000000")
_box  = Border(left=thin, right=thin, top=thin, bottom=thin)
_CEN  = Alignment(horizontal="center", vertical="center")
_LEFT = Alignment(horizontal="left",   vertical="center")
_bold = Font(name=FONT, size=10, bold=True)
_norm = Font(name=FONT, size=10)
_head = Font(name=FONT, size=12, bold=True)
_title = Font(name=FONT, size=14, bold=True)
_yellow = PatternFill("solid", fgColor="FFFF00")


def _put(ws, r, c, v=None, font=_norm, al=_CEN, fmt=None, fill=None):
    cell = ws.cell(row=r, column=c)
    if v is not None:
        cell.value = v
    cell.font = font
    cell.alignment = al
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


def _merge(ws, r1, c1, r2, c2):
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _ratio_by_length(p):
    """Return {length: ratio} for a pack, filling missing lengths with 0."""
    out = {ln: 0 for ln in LENGTHS}
    for s in p["sizes"]:
        ln = SIZE_LENGTH.get(s["size"]) or s["length"]
        if ln in out:
            out[ln] = s["ratio"]
    return out


def build_xlsx(head, packs, netnet, net, gross, cbm):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    widths = {"A": 8, "B": 8, "C": 18, "D": 10,
              "E": 7, "F": 7, "G": 7, "H": 7, "I": 7, "J": 7,
              "K": 10, "L": 12, "M": 12, "N": 12, "O": 12,
              "P": 10, "Q": 10, "R": 10, "S": 10}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ----- Title -----
    _merge(ws, 1, 1, 1, 19)
    _put(ws, 1, 1,
         "JCP PURCHASING CORPORATION, (JCPPC) IMPORT PACKING LIST",
         _title, _LEFT)

    _merge(ws, 3, 1, 3, 6)
    _put(ws, 3, 1,
         f"JCPPC SUB/CONTRACT/SUFFIX : {head.get('del_id','')}/"
         f"{head.get('order_no','')}{head.get('suffix','')}",
         _bold, _LEFT)
    _merge(ws, 3, 17, 3, 19)
    _put(ws, 3, 17, "PACKAGE TYPE        : 0         RETAIL", _norm, _LEFT)

    _merge(ws, 5, 1, 5, 6)
    _put(ws, 5, 1, "DESTINATION      : 9661-0, 5011 FIRESTORE PL", _norm, _LEFT)
    _merge(ws, 6, 1, 6, 6)
    _put(ws, 6, 1,
         "                              SOUTH GATE CA90280 U.S.A. BY SEA.",
         _norm, _LEFT)
    _merge(ws, 5, 12, 5, 19)
    _put(ws, 5, 12,
         "UCC/EAN MANUFACTURER/LABELER INDENTIFICATION NUMBER : 764746",
         _norm, _LEFT)

    _merge(ws, 8, 1, 8, 4)
    _put(ws, 8, 1, f"STYLE# {head.get('style','')}", _head, _LEFT)

    _merge(ws, 9, 1, 9, 4)
    _put(ws, 9, 1, f"PO# {head.get('order_no','')}", _norm, _LEFT)
    _merge(ws, 9, 5, 9, 8)
    _put(ws, 9, 5, f"DI# {head.get('del_id','')}", _norm, _LEFT)
    dcd = head.get("dcd_date", "")
    try:
        dcd_fmt = datetime.strptime(dcd, "%Y-%m-%d").strftime("%m/%d/%y")
    except Exception:
        dcd_fmt = dcd
    _merge(ws, 9, 12, 9, 19)
    _put(ws, 9, 12, f"DCD DATE : {dcd_fmt}", _norm, _LEFT)

    # ----- Rate row (row 10) -----
    R_RATE = 10
    _merge(ws, R_RATE, 1, R_RATE, 10)
    _put(ws, R_RATE, 1,
         "PER CARTON RATE  →   edit the yellow cells; every line recalculates",
         _bold, _LEFT)
    _put(ws, R_RATE, 12, round(netnet, 4), _bold, fmt="0.0000", fill=_yellow)
    _put(ws, R_RATE, 13, round(net,    4), _bold, fmt="0.0000", fill=_yellow)
    _put(ws, R_RATE, 14, round(gross,  4), _bold, fmt="0.0000", fill=_yellow)
    _put(ws, R_RATE, 15, round(cbm,    6), _bold, fmt="0.000000", fill=_yellow)

    # ----- Header rows -----
    R_HDR, R_HDR2 = 11, 12

    _merge(ws, R_HDR, 1, R_HDR, 4)
    _put(ws, R_HDR, 1, "PREPACK", _bold)
    _merge(ws, R_HDR, 5, R_HDR, 10)
    _put(ws, R_HDR, 5, "SIZE RATIO", _bold)

    for h1, c in [
        ("PCE PER",    11), ("NET NET",    12), ("NET WEIGHT", 13),
        ("GR.WGT.",    14), ("CBMs /",     15), ("TOTAL",      16),
        ("TOTAL",      17), ("FIRST",      18), ("LAST",       19),
    ]:
        _put(ws, R_HDR, c, h1, _bold)

    _put(ws, R_HDR2, 1, "LOT #", _bold)
    _put(ws, R_HDR2, 2, "LINE",  _bold)
    _put(ws, R_HDR2, 3, "COLOR", _bold)
    _put(ws, R_HDR2, 4, "CODE",  _bold)
    for i, ln in enumerate(LENGTHS):
        _put(ws, R_HDR2, 5 + i, ln, _bold)

    for h2, c in [
        ("CARTON",    11), ("WEIGHT/CTN", 12), ("/CARTON", 13),
        ("/CARTON",   14), ("",           15), ("PCS.",    16),
        ("CARTONS",   17), ("CARTON#",    18), ("CARTON#", 19),
    ]:
        if h2:
            _put(ws, R_HDR2, c, h2, _bold)

    for c in range(1, 20):
        for rr in (R_HDR, R_HDR2):
            ws.cell(row=rr, column=c).border = _box

    # ----- Data rows -----
    r = R_HDR2 + 1
    data_first = r

    for p in packs:
        pcs_per_ctn = p["pce_per_ctn"]
        ctn_qty = int(p["total_pkg"])

        _put(ws, r, 1, head.get("item_number", ""))
        _put(ws, r, 2, p["order_line"])
        _put(ws, r, 3, p["color"])
        _put(ws, r, 4, p["prepack_code"])

        rbl = _ratio_by_length(p)
        for i, ln in enumerate(LENGTHS):
            _put(ws, r, 5 + i, rbl[ln], fmt="0")

        _put(ws, r, 11, pcs_per_ctn, fmt="0")
        _put(ws, r, 12, f"=Q{r}*$L${R_RATE}", fmt="0.0000")
        _put(ws, r, 13, f"=Q{r}*$M${R_RATE}", fmt="0.0000")
        _put(ws, r, 14, f"=Q{r}*$N${R_RATE}", fmt="0.0000")
        _put(ws, r, 15, f"=Q{r}*$O${R_RATE}", fmt="0.000000")
        _put(ws, r, 16, f"=K{r}*Q{r}", fmt="0")
        _put(ws, r, 17, ctn_qty, fmt="0")

        if r == data_first:
            _put(ws, r, 18, 1, fmt="0")
        else:
            _put(ws, r, 18, f"=S{r-1}+1", fmt="0")
        _put(ws, r, 19, f"=R{r}+Q{r}-1", fmt="0")

        for c in range(1, 20):
            ws.cell(row=r, column=c).border = _box
        r += 1

    data_last = r - 1

    # ----- TOTAL row -----
    total_row = r
    _put(ws, r, 11, "TOTAL", _bold)
    for col, fmt in ((12, "0.0000"), (13, "0.0000"), (14, "0.0000"),
                     (15, "0.000000"), (16, "0"), (17, "0")):
        cl = L(col)
        _put(ws, r, col,
             f"=SUM({cl}{data_first}:{cl}{data_last})",
             _bold, fmt=fmt)
    for c in range(1, 20):
        ws.cell(row=r, column=c).border = _box
    r += 2

    # ----- Footer -----
    _put(ws, r, 1, "CARTON MEASUREMENT : 24.5 X 21 X7.5 INCH-18 PCS",
         _norm, _LEFT)
    r += 2

    _put(ws, r, 1, "THAT'S IT GARMENTS LTD.", _bold, _LEFT)
    _put(ws, r, 11, "CARTON TOTAL :", _bold, _LEFT)
    r += 1
    _put(ws, r, 1, "NORASHING PUR , ASHULIA", _norm, _LEFT)
    _put(ws, r, 11, "CARTONS     :", _bold, _LEFT)
    _put(ws, r, 13, f"=Q{total_row}", fmt="0")
    r += 1
    _put(ws, r, 1, "DHAKA, BANGLADESH.", _norm, _LEFT)
    _put(ws, r, 11, "PICEES         :", _bold, _LEFT)
    _put(ws, r, 13, f"=P{total_row}", fmt="0")
    r += 1
    _put(ws, r, 1,
         "TEL : 02-880-8811149, 8823448, FAX : 02-880-8826090", _norm, _LEFT)
    _put(ws, r, 11, "NET NET-KG:", _bold, _LEFT)
    _put(ws, r, 13, f"=L{total_row}", fmt="0.0000")
    r += 1
    _put(ws, r, 1, "EMAIL : RUBEL@HAMEEMGROUP.COM", _norm, _LEFT)
    _put(ws, r, 11, "NET WT-KG  :", _bold, _LEFT)
    _put(ws, r, 13, f"=M{total_row}", fmt="0.0000")
    r += 1
    _put(ws, r, 11, "GR.WT-KG    :", _bold, _LEFT)
    _put(ws, r, 13, f"=N{total_row}", fmt="0.0000")
    r += 1
    _put(ws, r, 11, "CBM               :", _bold, _LEFT)
    _put(ws, r, 13, f"=O{total_row}", fmt="0.000000")

    wb.calculation.fullCalcOnLoad = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================
# 4. Streamlit render helpers
# ============================================================
def _render_header_block(head):
    """Mirror the Excel header block on the web page."""
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### Order / Style")
        st.markdown(f"**PO #:**  {head.get('order_no','')}")
        st.markdown(f"**Style #:**  {head.get('style','')}")
        st.markdown(f"**Item #:**  {head.get('item_number','')}")
        st.markdown(f"**DI #:**  {head.get('del_id','')}")
        st.markdown(f"**Sub/Contract/Suffix:**  "
                    f"{head.get('del_id','')}/{head.get('order_no','')}{head.get('suffix','')}")
        dcd = head.get("dcd_date", "")
        try:
            dcd_fmt = datetime.strptime(dcd, "%Y-%m-%d").strftime("%m/%d/%y")
        except Exception:
            dcd_fmt = dcd
        st.markdown(f"**DCD Date:**  {dcd_fmt}")
        if head.get("order_qty"):
            st.markdown(f"**Order Qty (EA):**  {head['order_qty']:,}")

    with c2:
        st.markdown("##### Destination / Package")
        st.markdown("**Destination:**  9661-0, 5011 Firestone Pl, "
                    "South Gate CA90280 U.S.A. by Sea")
        st.markdown("**Package Type:**  0  RETAIL")
        st.markdown("**UCC/EAN Mfr/Labeler ID:**  764746")
        st.markdown(f"**Brand:**  {head.get('brand','')}")
        st.markdown(f"**Supply Chain:**  {head.get('supply_chain','')}")
        if head.get("carton_dims"):
            st.markdown(f"**Carton Dimensions:**  {head['carton_dims']}")


def _render_order_qty_table(packs):
    """Compact Order Qty: lengths across the top, QTY row below, TOTAL at end."""
    st.markdown("##### Order Quantity (total pieces per length)")

    totals = {ln: 0 for ln in LENGTHS}
    for p in packs:
        rbl = _ratio_by_length(p)
        for ln in LENGTHS:
            totals[ln] += rbl[ln] * int(p["total_pkg"])

    header_row = ["QTY"] + [totals[ln] for ln in LENGTHS] + [sum(totals.values())]
    df = pd.DataFrame([header_row], columns=["SIZE"] + LENGTHS + ["TOTAL"])

    styler = (
        df.style
        .set_properties(subset=["SIZE"],
                        **{"text-align": "left", "font-weight": "bold"})
        .set_properties(
            subset=[c for c in df.columns if c != "SIZE"],
            **{"text-align": "center"},
        )
    )
    st.dataframe(styler, use_container_width=True, hide_index=True)


def _render_carton_detail(packs, netnet, net, gross, cbm, head):
    """Full carton detail — mirror of the Excel data rows + total."""
    st.markdown("##### Carton Details")

    rows = []
    first_ctn = 1
    for p in packs:
        ctn_qty = int(p["total_pkg"])
        pcs_per_ctn = p["pce_per_ctn"]

        row = {
            "LOT #":  head.get("item_number", ""),
            "LINE":   p["order_line"],
            "COLOR":  p["color"],
            "CODE":   p["prepack_code"],
        }
        rbl = _ratio_by_length(p)
        for ln in LENGTHS:
            row[ln] = rbl[ln]
        row["PCE PER CARTON"]   = pcs_per_ctn
        row["NET NET /CTN"]     = round(ctn_qty * netnet, 4)
        row["NET WT /CTN"]      = round(ctn_qty * net,    4)
        row["GR.WT. /CTN"]      = round(ctn_qty * gross,  4)
        row["CBMs"]             = round(ctn_qty * cbm,    6)
        row["TOTAL PCS"]        = pcs_per_ctn * ctn_qty
        row["TOTAL CARTONS"]    = ctn_qty
        row["FIRST CARTON#"]    = first_ctn
        row["LAST CARTON#"]     = first_ctn + ctn_qty - 1
        rows.append(row)
        first_ctn += ctn_qty

    # TOTAL row
    total_row = {
        "LOT #":  "TOTAL",
        "LINE":   "",
        "COLOR":  "",
        "CODE":   "",
    }
    for ln in LENGTHS:
        total_row[ln] = sum(r[ln] for r in rows)
    total_row["PCE PER CARTON"] = ""
    total_row["NET NET /CTN"]   = round(sum(r["NET NET /CTN"] for r in rows), 4)
    total_row["NET WT /CTN"]    = round(sum(r["NET WT /CTN"]  for r in rows), 4)
    total_row["GR.WT. /CTN"]    = round(sum(r["GR.WT. /CTN"]  for r in rows), 4)
    total_row["CBMs"]           = round(sum(r["CBMs"]         for r in rows), 6)
    total_row["TOTAL PCS"]      = sum(r["TOTAL PCS"]     for r in rows)
    total_row["TOTAL CARTONS"]  = sum(r["TOTAL CARTONS"] for r in rows)
    total_row["FIRST CARTON#"]  = ""
    total_row["LAST CARTON#"]   = ""
    rows.append(total_row)

    df = pd.DataFrame(rows)
    left_cols = ["LOT #", "LINE", "COLOR", "CODE"]
    styler = (
        df.style
        .set_properties(subset=left_cols, **{"text-align": "left"})
        .set_properties(
            subset=[c for c in df.columns if c not in left_cols],
            **{"text-align": "center"},
        )
        .apply(lambda r: [
            "font-weight: bold; background-color: #F0F0F0"
            if r["LOT #"] == "TOTAL" else ""
            for _ in r
        ], axis=1)
    )
    st.dataframe(styler, use_container_width=True, hide_index=True)


def _render_summary(packs, netnet, net, gross, cbm):
    """Summary totals — matches the Excel footer."""
    st.markdown("##### Summary")

    total_cartons = sum(int(p["total_pkg"]) for p in packs)
    total_pcs     = sum(p["pce_per_ctn"] * int(p["total_pkg"]) for p in packs)
    total_netnet  = round(total_cartons * netnet, 4)
    total_net     = round(total_cartons * net,    4)
    total_gross   = round(total_cartons * gross,  4)
    total_cbm     = round(total_cartons * cbm,    6)

    summary = pd.DataFrame([
        {"Description": "CARTON TOTAL",      "Value": f"{total_cartons:,}"},
        {"Description": "TOTAL PIECES",      "Value": f"{total_pcs:,}"},
        {"Description": "NET NET (kg)",      "Value": f"{total_netnet:,.4f}"},
        {"Description": "NET WEIGHT (kg)",   "Value": f"{total_net:,.4f}"},
        {"Description": "GROSS WEIGHT (kg)", "Value": f"{total_gross:,.4f}"},
        {"Description": "TOTAL CBM (m³)",    "Value": f"{total_cbm:,.6f}"},
    ])

    styler = (
        summary.style
        .set_properties(subset=["Description"],
                        **{"text-align": "left", "font-weight": "bold"})
        .set_properties(subset=["Value"], **{"text-align": "center"})
    )
    st.dataframe(styler, use_container_width=True, hide_index=True)


# ============================================================
# 5. UI
# ============================================================
st.set_page_config(page_title="JCP PO -> Packing List",
                   layout="wide", page_icon="📦")
st.title("📦 JCPenney — PO to Packing List")
st.caption("Upload a JCPenney PO (.docx) and download the JCP PC import packing list.")

up = st.file_uploader("Upload PO .docx", type=["docx"])

with st.sidebar:
    st.markdown("### Carton weights (per carton)")
    netnet = st.number_input("NET NET kg/carton", 0.0, 100.0,
                             DEFAULT_NETNET, 0.01, format="%.2f")
    net    = st.number_input("NET kg/carton",     0.0, 100.0,
                             DEFAULT_NET,    0.01, format="%.2f")
    gross  = st.number_input("GROSS kg/carton",   0.0, 100.0,
                             DEFAULT_GROSS,  0.01, format="%.2f")
    cbm    = st.number_input("CBM /carton",       0.0, 10.0,
                             DEFAULT_CBM,    0.0001, format="%.4f")

if up is None:
    st.info("Upload a .docx to get started.")
    st.stop()

try:
    text = read_docx_text(up.getvalue())
except Exception as e:
    st.error(f"Could not read the DOCX: {e}")
    st.stop()

with st.expander("🔍 Extracted text from DOCX (first 250 lines)", expanded=False):
    st.text("\n".join(text.splitlines()[:250]))

try:
    head, packs = parse_po(text)
except Exception as e:
    st.error(f"Parse failed: {e}")
    st.stop()

if not packs:
    st.warning("No PREPACK blocks were found.")
    with st.expander("Full extracted text", expanded=True):
        st.code(text[:6000])
    st.stop()

st.success(f"PO **{head.get('order_no','')}** · "
           f"Style **{head.get('style','')}** · {len(packs)} pack line(s)")

# ---- 1) HEADER BLOCK ----
_render_header_block(head)
st.markdown("---")

# ---- 2) ORDER QTY TABLE (BEFORE main) ----
_render_order_qty_table(packs)
st.markdown("---")

# ---- 3) CARTON DETAIL TABLE (MAIN) ----
_render_carton_detail(packs, netnet, net, gross, cbm, head)

# ---- 4) SUMMARY TABLE (AFTER main) ----
_render_summary(packs, netnet, net, gross, cbm)

st.divider()
if st.button("Prepare packing list (.xlsx)", type="primary"):
    st.session_state["jcp_xlsx"] = build_xlsx(
        head, packs, netnet, net, gross, cbm)

if st.session_state.get("jcp_xlsx"):
    st.download_button(
        "⬇ Download packing list",
        data=st.session_state["jcp_xlsx"],
        file_name=f"packing_list_{head.get('order_no','')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )