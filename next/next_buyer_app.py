# """
# Next Retail Ltd. / Clover Collections — Packing List generator
# ==============================================================

# Input:  Contract Confirmation Report XLSX (from Next ePLM).
# Output: NEXT standard packing list XLSX (bulk carton, single size per carton).

# Standalone:  streamlit run next_buyer_app.py
# Module use:  from next_buyer_app import run;  run(xlsx_file)
# """
# import io
# import re
# from collections import OrderedDict
# from datetime import date

# import openpyxl
# import pandas as pd
# import streamlit as st
# from openpyxl import Workbook
# from openpyxl.styles import Alignment, Border, Font, Side
# from openpyxl.utils import get_column_letter as L

# # ============================================================
# # DEFAULTS (editable in the UI)
# # ============================================================
# DEFAULT_PACK_PER_CTN = 9
# DEFAULT_FULL_L, DEFAULT_FULL_W, DEFAULT_FULL_H = 60, 30, 40
# DEFAULT_PART_L, DEFAULT_PART_W, DEFAULT_PART_H = 60, 30, 20
# DEFAULT_GROSS_PER_CTN = 6.1
# DEFAULT_TARE = 1.8
# DEFAULT_PACK_OF = "5 PK"

# # ============================================================
# # PARSER — Contract Confirmation Report
# # ============================================================
# SIZE_RE = re.compile(r"^\s*(\d+)\s*Yrs?\s*$", re.I)


# def _next_nonempty(row, start):
#     for j in range(start, len(row)):
#         v = row[j]
#         if v is not None and str(v).strip() != "":
#             return v
#     return None


# def _size_sort_key(lbl):
#     m = re.match(r"(\d+)", str(lbl) or "")
#     return int(m.group(1)) if m else 9999


# def parse_next_xlsx(xlsx_bytes: bytes):
#     """Return list of item dicts: {item, desc, contract, sizes: {label: qty}}."""
#     wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
#     ws = wb[wb.sheetnames[0]]
#     rows = [list(r) for r in ws.iter_rows(values_only=True)]

#     items, current = [], None
#     for i, row in enumerate(rows):
#         for c, cell in enumerate(row):
#             if not isinstance(cell, str):
#                 continue
#             key = cell.strip()
#             if key == "Item Number":
#                 val = _next_nonempty(row, c + 1)
#                 if not val:
#                     continue
#                 val_s = str(val).strip()
#                 if current and current.get("item"):
#                     items.append(current)
#                 current = {"item": val_s, "desc": "",
#                            "contract": "", "sizes": OrderedDict()}
#             elif key == "Item Description" and current is not None:
#                 val = _next_nonempty(row, c + 1)
#                 current["desc"] = str(val or "").strip()

#         if current is None:
#             continue

#         for c, cell in enumerate(row):
#             if isinstance(cell, str) and cell.strip() == "Contract No":
#                 if i + 1 < len(rows):
#                     nxt = rows[i + 1]
#                     if c < len(nxt) and nxt[c]:
#                         v = str(nxt[c]).strip()
#                         if v and not current["contract"]:
#                             current["contract"] = v

#         size_cols = {}
#         for c, cell in enumerate(row):
#             if isinstance(cell, str):
#                 m = SIZE_RE.match(cell)
#                 if m:
#                     size_cols[c] = f"{int(m.group(1))} Yrs"
#         if size_cols:
#             for j in range(i + 1, min(i + 6, len(rows))):
#                 qrow = rows[j]
#                 if any(isinstance(x, str) and x.strip() == "Quantity" for x in qrow):
#                     for c, lbl in size_cols.items():
#                         v = qrow[c] if c < len(qrow) else None
#                         if isinstance(v, (int, float)):
#                             current["sizes"][lbl] = int(v)
#                     break

#     if current and current.get("item"):
#         items.append(current)

#     for it in items:
#         it["sizes"] = OrderedDict(
#             sorted(it["sizes"].items(), key=lambda kv: _size_sort_key(kv[0])))
#     return [it for it in items if it["sizes"]]


# # ============================================================
# # CARTON PLAN
# # ============================================================
# def build_plan(item, pack, gross, tare, full_dims, part_dims):
#     rows, cno = [], 1
#     Lf, Wf, Hf = full_dims
#     Lp, Wp, Hp = part_dims
#     for sz, qty in item["sizes"].items():
#         p = max(int(pack.get(sz, DEFAULT_PACK_PER_CTN)), 1)
#         full, rem = divmod(int(qty), p)
#         g_full = float(gross.get(sz, DEFAULT_GROSS_PER_CTN))
#         t      = float(tare.get(sz, DEFAULT_TARE))

#         if full:
#             rows.append({
#                 "from": cno, "to": cno + full - 1, "size": sz,
#                 "n_ctn": full, "pcs_per_ctn": p,
#                 "L": Lf, "W": Wf, "H": Hf,
#                 "m3_per_ctn": Lf * Wf * Hf / 1_000_000.0,
#                 "gross_per_ctn": g_full,
#                 "net_per_ctn": round(g_full - t, 2),
#             })
#             cno += full

#         if rem:
#             g_part = round(g_full - (p - rem) * 0.07, 2)
#             g_part = max(g_part, max(round(t + 0.1, 2), 1.0))
#             rows.append({
#                 "from": cno, "to": cno, "size": sz,
#                 "n_ctn": 1, "pcs_per_ctn": rem,
#                 "L": Lp, "W": Wp, "H": Hp,
#                 "m3_per_ctn": Lp * Wp * Hp / 1_000_000.0,
#                 "gross_per_ctn": g_part,
#                 "net_per_ctn": round(max(g_part - t, 0.1), 2),
#             })
#             cno += 1
#     return rows


# # ============================================================
# # XLSX — write one packing list into an existing worksheet
# # ============================================================
# def fill_sheet(ws, head, item, rows):
#     """Fill `ws` with the NEXT standard bulk-carton packing list."""
#     TNR, ARI = "Times New Roman", "Arial"
#     thin = Side(style="thin")
#     CEN  = Alignment(horizontal="center", vertical="center", wrap_text=True)
#     LEFT = Alignment(horizontal="left",   vertical="center", wrap_text=True)

#     def put(r, c, v=None, name=ARI, sz=10, al=CEN, fmt=None, bold=False):
#         cell = ws.cell(row=r, column=c)
#         if v is not None:
#             cell.value = v
#         cell.font = Font(name=name, size=sz, bold=bold)
#         cell.alignment = al
#         if fmt:
#             cell.number_format = fmt
#         return cell

#     def merge(r1, c1, r2, c2):
#         if (r1, c1) != (r2, c2):
#             ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

#     def box(r1, c1, r2, c2):
#         for r in range(r1, r2 + 1):
#             for c in range(c1, c2 + 1):
#                 ws.cell(row=r, column=c).border = Border(left=thin, right=thin,
#                                                          top=thin, bottom=thin)

#     # ---------- header ----------
#     put(1, 1, "PACKING LIST", TNR, 13, al=LEFT, bold=True)
#     put(3, 1, "Agent / Supplier:", bold=True); put(3, 3, head["supplier"])
#     put(4, 1, "Factory name: (Full name)", bold=True); put(4, 3, head["factory"])
#     put(5, 1, "Factory address:", bold=True)
#     put(5, 3, head["factory_address"], al=LEFT); merge(5, 3, 6, 8)

#     put(3, 10, "Consignee Full name:", bold=True); put(3, 12, "Next Retail Limited")
#     put(4, 10, "Consignee Full address:", bold=True)
#     put(4, 12, "Desford Road, Enderby, Leicester, LE19 4AT, UK")
#     put(5, 10, "Delivery to:", bold=True)
#     put(5, 12, "Nominated forwarder in origin territory")

#     put(8, 1, "NEXT Contract Number:", bold=True); put(8, 4, head["contract"])
#     put(9, 1, "NEXT Item Number:", bold=True);     put(9, 4, item["item"])

#     if head.get("pack_of"):
#         put(10, 1, f"Pieces per Pack : {head['pack_of']}", bold=True)

#     put(8, 10, "PACKING METHOD", bold=True)
#     put(9, 10, "BULK CARTON - Single size/option per carton", bold=True)
#     put(9, 15, "√", bold=True)
#     put(10, 10, "PLEASE COMPLETE ALL COLUMNS & SIZE SUMMARY", bold=True)

#     put(12, 1, "Full Description and composition:", bold=True)
#     put(13, 1, "Product type :", bold=True);     put(13, 3, head["product_type"])
#     put(14, 1, "Item Composition:", bold=True);  put(14, 3, head["composition"])
#     put(15, 1, "Knitted or Woven:", bold=True);  put(15, 3, head["knit_woven"])

#     total_pcs = sum(r["n_ctn"] * r["pcs_per_ctn"] for r in rows)
#     total_net = sum(r["n_ctn"] * r["net_per_ctn"] for r in rows)

#     put(12, 10, "Total Net Weight (kg):", bold=True);   put(12, 12, round(total_net, 2))
#     put(13, 10, "Total pieces (Column G):", bold=True); put(13, 12, total_pcs)

#     put(17, 1, "Country of Origin:", bold=True); put(17, 3, head["origin"])
#     put(18, 1, "Preference available for this item?", bold=True); put(18, 3, "Yes")
#     put(17, 10, "Date of Issue:", bold=True);     put(17, 12, head["date"])

#     # ---------- carton table ----------
#     R_HDR = 20
#     R_D0  = R_HDR + 2

#     put(R_HDR, 1, "Carton Numbers", bold=True); merge(R_HDR, 1, R_HDR, 2)
#     put(R_HDR, 3, "Option", bold=True)
#     put(R_HDR, 4, "Size", bold=True)
#     put(R_HDR, 5, "Number of Cartons", bold=True)
#     put(R_HDR, 6, "Pieces* per Carton", bold=True)
#     put(R_HDR, 7, "Total pieces* supplied", bold=True)
#     put(R_HDR, 8, "Carton Dimensions", bold=True); merge(R_HDR, 8, R_HDR, 10)
#     put(R_HDR, 11, "Cubic metres (m3)", bold=True)
#     put(R_HDR, 12, "Gross Weight (kg)", bold=True)
#     put(R_HDR + 1, 1, "From", bold=True); put(R_HDR + 1, 2, "To", bold=True)
#     put(R_HDR + 1, 8, "Length (cm)", bold=True)
#     put(R_HDR + 1, 9, "Width (cm)",  bold=True)
#     put(R_HDR + 1, 10, "Height (cm)", bold=True)
#     box(R_HDR, 1, R_HDR + 1, 12)

#     r = R_D0
#     for row in rows:
#         put(r, 1, row["from"], fmt="0")
#         put(r, 2, row["to"],   fmt="0")
#         put(r, 3, item["item"])
#         put(r, 4, row["size"])
#         put(r, 5, row["n_ctn"], fmt="0")
#         put(r, 6, row["pcs_per_ctn"], fmt="0")
#         put(r, 7, row["n_ctn"] * row["pcs_per_ctn"], fmt="0")
#         put(r, 8, row["L"], fmt="0")
#         put(r, 9, row["W"], fmt="0")
#         put(r, 10, row["H"], fmt="0")
#         put(r, 11, round(row["m3_per_ctn"] * row["n_ctn"], 4), fmt="0.0000")
#         put(r, 12, round(row["gross_per_ctn"] * row["n_ctn"], 2), fmt="0.00")
#         r += 1

#     R_TOT = r
#     put(R_TOT, 1, "TOTALS", bold=True); merge(R_TOT, 1, R_TOT, 4)
#     put(R_TOT, 5,  f"=SUM(E{R_D0}:E{R_TOT-1})", bold=True)
#     put(R_TOT, 7,  f"=SUM(G{R_D0}:G{R_TOT-1})", bold=True)
#     put(R_TOT, 11, f"=SUM(K{R_D0}:K{R_TOT-1})", bold=True, fmt="0.0000")
#     put(R_TOT, 12, f"=SUM(L{R_D0}:L{R_TOT-1})", bold=True, fmt="0.00")
#     box(R_HDR, 1, R_TOT, 12)

#     # ---------- size summary ----------
#     R_SUM = R_TOT + 3
#     put(R_SUM, 1, "SIZE", bold=True)
#     sizes_ord = list(item["sizes"].keys())
#     for i, sz in enumerate(sizes_ord):
#         put(R_SUM, 3 + i, sz, bold=True)
#     put(R_SUM, 3 + len(sizes_ord), "TOTAL", bold=True)

#     r_ord, r_ship, r_diff, r_pct = R_SUM + 1, R_SUM + 2, R_SUM + 3, R_SUM + 4
#     put(r_ord,  1, "Order Qty",    bold=True, al=LEFT)
#     put(r_ship, 1, "Shipped Qty",  bold=True, al=LEFT)
#     put(r_diff, 1, "Ex/Short Qty", bold=True, al=LEFT)
#     put(r_pct,  1, "% Variance",   bold=True, al=LEFT)

#     shipped = {sz: 0 for sz in sizes_ord}
#     for row in rows:
#         shipped[row["size"]] += row["n_ctn"] * row["pcs_per_ctn"]

#     for i, sz in enumerate(sizes_ord):
#         c = 3 + i
#         put(r_ord,  c, item["sizes"][sz])
#         put(r_ship, c, shipped[sz])
#         put(r_diff, c, f"={L(c)}{r_ship}-{L(c)}{r_ord}")
#         put(r_pct,  c, f"=IFERROR({L(c)}{r_diff}/{L(c)}{r_ord},0)", fmt="0.0%")
#     last = 2 + len(sizes_ord)
#     put(r_ord,  last + 1, f"=SUM(C{r_ord}:{L(last)}{r_ord})",   bold=True)
#     put(r_ship, last + 1, f"=SUM(C{r_ship}:{L(last)}{r_ship})", bold=True)
#     put(r_diff, last + 1, f"=SUM(C{r_diff}:{L(last)}{r_diff})", bold=True)
#     box(R_SUM, 1, r_pct, last + 1)

#     # ---------- signature ----------
#     R_SIG = r_pct + 3
#     put(R_SIG, 1, "Signed by (Factory)", bold=True, al=LEFT)
#     put(R_SIG, 8, "Factory stamp", bold=True, al=LEFT)
#     put(R_SIG + 2, 1, "Town/City, Country", bold=True, al=LEFT)
#     put(R_SIG + 4, 1, "Date", bold=True, al=LEFT)

#     widths = {1: 22, 2: 6, 3: 14, 4: 8, 5: 12, 6: 12, 7: 14,
#               8: 10, 9: 10, 10: 10, 11: 12, 12: 14}
#     for c, w in widths.items():
#         ws.column_dimensions[L(c)].width = w

#     ws.page_setup.orientation = "landscape"
#     ws.page_setup.paperSize = 9
#     ws.page_setup.fitToWidth = 1
#     ws.page_setup.fitToHeight = 1
#     ws.sheet_properties.pageSetUpPr.fitToPage = True
#     ws.page_margins.left = ws.page_margins.right = 0.25
#     ws.page_margins.top  = ws.page_margins.bottom = 0.25


# def build_workbook(head, item, rows):
#     """Convenience wrapper: one-item workbook (kept for standalone use)."""
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "PACKING LIST"
#     fill_sheet(ws, head, item, rows)
#     return wb


# # ============================================================
# # Streamlit UI
# # ============================================================
# @st.cache_data(show_spinner=False, max_entries=5)
# def _parse_cached(b: bytes):
#     return parse_next_xlsx(b)


# def _item_ui(item, key):
#     st.markdown(f"#### {item['item']} — {item['desc']}")
#     tot = sum(item["sizes"].values())
#     st.caption(f"Contract {item.get('contract') or '?'}  ·  {tot:,} pcs  ·  "
#                f"{len(item['sizes'])} sizes")

#     sizes = list(item["sizes"].keys())

#     with st.expander("Packing configuration", expanded=True):
#         cols = st.columns(min(len(sizes), 6))
#         pack, gross, tare = {}, {}, {}
#         for i, sz in enumerate(sizes):
#             with cols[i % len(cols)]:
#                 st.markdown(f"**{sz}**")
#                 pack[sz]  = st.number_input("Pcs/ctn", 1, 500,
#                                             DEFAULT_PACK_PER_CTN, 1,
#                                             key=f"{key}_pk_{sz}")
#                 gross[sz] = st.number_input("GW/ctn kg", 0.5, 50.0,
#                                             DEFAULT_GROSS_PER_CTN, 0.1,
#                                             key=f"{key}_gr_{sz}")
#                 tare[sz]  = st.number_input("Tare kg", 0.0, 10.0,
#                                             DEFAULT_TARE, 0.1,
#                                             key=f"{key}_tr_{sz}")

#         c1, c2 = st.columns(2)
#         with c1:
#             st.markdown("**Full carton dims (cm)**")
#             fL = st.number_input("Full L", 10, 200, DEFAULT_FULL_L, 1, key=f"{key}_fL")
#             fW = st.number_input("Full W", 10, 200, DEFAULT_FULL_W, 1, key=f"{key}_fW")
#             fH = st.number_input("Full H", 10, 200, DEFAULT_FULL_H, 1, key=f"{key}_fH")
#         with c2:
#             st.markdown("**Partial carton dims (cm)**")
#             pL = st.number_input("Part L", 10, 200, DEFAULT_PART_L, 1, key=f"{key}_pL")
#             pW = st.number_input("Part W", 10, 200, DEFAULT_PART_W, 1, key=f"{key}_pW")
#             pH = st.number_input("Part H", 10, 200, DEFAULT_PART_H, 1, key=f"{key}_pH")

#         c1, c2, c3 = st.columns(3)
#         with c1:
#             product_type = st.text_input("Product type", "Boys School Shirt",
#                                          key=f"{key}_pt")
#         with c2:
#             composition = st.text_input("Composition",
#                                         "65% Recycled polyester, 35% Cotton",
#                                         key=f"{key}_comp")
#         with c3:
#             knit_woven = st.text_input("Knitted or Woven", "Woven", key=f"{key}_kw")

#         c4, c5 = st.columns(2)
#         with c4:
#             pack_of = st.text_input(
#                 "Pieces per Pack (printed on the sheet; clear to hide)",
#                 DEFAULT_PACK_OF, key=f"{key}_pop")
#         with c5:
#             contract_suffix = st.text_input(
#                 "Contract suffix (appended after contract #)",
#                 "-1", key=f"{key}_cs")

#     rows = build_plan(item, pack, gross, tare,
#                       (fL, fW, fH), (pL, pW, pH))

#     df = pd.DataFrame([{
#         "From": r["from"], "To": r["to"], "Size": r["size"],
#         "Cartons": r["n_ctn"], "Pcs/ctn": r["pcs_per_ctn"],
#         "Total pcs": r["n_ctn"] * r["pcs_per_ctn"],
#         "L": r["L"], "W": r["W"], "H": r["H"],
#         "m3": round(r["m3_per_ctn"] * r["n_ctn"], 4),
#         "GW kg": round(r["gross_per_ctn"] * r["n_ctn"], 2),
#     } for r in rows])
#     st.markdown("##### Carton plan")
#     st.dataframe(df, use_container_width=True, hide_index=True)

#     tot_ctn = sum(r["n_ctn"] for r in rows)
#     tot_pcs = sum(r["n_ctn"] * r["pcs_per_ctn"] for r in rows)
#     tot_m3  = sum(r["n_ctn"] * r["m3_per_ctn"]   for r in rows)
#     tot_gw  = sum(r["n_ctn"] * r["gross_per_ctn"] for r in rows)
#     tot_nw  = sum(r["n_ctn"] * r["net_per_ctn"]  for r in rows)
#     m = st.columns(5)
#     m[0].metric("Cartons", tot_ctn)
#     m[1].metric("Pieces",  f"{tot_pcs:,}")
#     m[2].metric("m3",      f"{tot_m3:.3f}")
#     m[3].metric("Gross kg", f"{tot_gw:.2f}")
#     m[4].metric("Net kg",   f"{tot_nw:.2f}")

#     if tot_pcs != tot:
#         st.warning(f"Packed {tot_pcs:,} vs PO {tot:,} pcs (diff {tot_pcs - tot:+d})")

#     head = {
#         "supplier":         "Clover Collections",
#         "factory":          "CREATIVE COLLECTIONS LTD.",
#         "factory_address":  "69,107 NISHAT NAGAR, TONGI, GAZIPUR-1711, BANGLADESH",
#         "contract":         f"{item.get('contract','')}{contract_suffix}",
#         "product_type":     product_type,
#         "composition":      composition,
#         "knit_woven":       knit_woven,
#         "origin":           "Bangladesh",
#         "date":             date.today(),
#         "pack_of":          pack_of.strip() or None,
#     }
#     return head, item, rows


# def run(xlsx_file=None):
#     st.title("Next Retail / Clover — Packing List")
#     if xlsx_file is None:
#         st.info("Upload a Contract Confirmation Report (.xlsx) from Next ePLM.")
#         return

#     try:
#         items = _parse_cached(xlsx_file.getvalue())
#     except Exception as e:
#         st.error(f"Could not read the XLSX: {e}")
#         return

#     if not items:
#         st.warning("No item blocks found — check the file is a Contract Confirmation Report.")
#         return

#     st.success(f"Parsed **{len(items)}** item(s) from `{xlsx_file.name}`")

#     results = []
#     tabs = st.tabs([f"{it['item']} — {it['desc'][:22]}" for it in items])
#     for tab, it in zip(tabs, items):
#         with tab:
#             results.append(_item_ui(it, key=f"nx_{it['item']}"))

#     # build one workbook, one sheet per item — no cross-workbook copying
#     combined = Workbook()
#     combined.remove(combined.active)
#     for head, item, rows in results:
#         sheet_name = re.sub(r"[\[\]:*?/\\]", "-", str(item["item"]))[:31] or "Sheet1"
#         ws = combined.create_sheet(title=sheet_name)
#         fill_sheet(ws, head, item, rows)

#     buf = io.BytesIO(); combined.save(buf)
#     st.download_button(
#         "Download packing list (.xlsx)", buf.getvalue(),
#         file_name=f"packing_list_{items[0].get('contract','contract')}.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         type="primary", key="nx_dl")


# if __name__ == "__main__":
#     st.set_page_config(page_title="Next / Clover Packing List", layout="wide")
#     run(st.file_uploader("Upload Contract Confirmation Report (.xlsx)", type=["xlsx"]))


"""
Next Retail Ltd. / Clover Collections — Packing List generator
==============================================================

Input:  Contract Confirmation Report XLSX (from Next ePLM).
Output: NEXT standard packing list XLSX (bulk carton, single size per carton).

Standalone:  streamlit run next_buyer_app.py
Module use:  from next_buyer_app import run;  run(xlsx_file)
"""
import io
import re
from collections import OrderedDict
from datetime import date

import openpyxl
import pandas as pd
import sys as _sys
from pathlib import Path as _Path
_ROOT = str(_Path(__file__).resolve().parent.parent)
if _ROOT not in _sys.path:                     # pl_kit.py lives in the project root
    _sys.path.insert(0, _ROOT)
import pl_kit as kit                           # noqa: E402
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter as L

# ============================================================
# DEFAULTS (editable in the UI)
# ============================================================
DEFAULT_PACK_PER_CTN = 9
DEFAULT_FULL_L, DEFAULT_FULL_W, DEFAULT_FULL_H = 60, 30, 40
DEFAULT_PART_L, DEFAULT_PART_W, DEFAULT_PART_H = 60, 30, 20
DEFAULT_GROSS_PER_CTN = 6.1
DEFAULT_TARE = 1.8
DEFAULT_PACK_OF = "5 PK"

# ============================================================
# PARSER — Contract Confirmation Report
# ============================================================
SIZE_RE = re.compile(r"^\s*(\d+)\s*Yrs?\s*$", re.I)


def _next_nonempty(row, start):
    for j in range(start, len(row)):
        v = row[j]
        if v is not None and str(v).strip() != "":
            return v
    return None


def _size_sort_key(lbl):
    m = re.match(r"(\d+)", str(lbl) or "")
    return int(m.group(1)) if m else 9999


def parse_next_xlsx(xlsx_bytes: bytes):
    """Return list of item dicts: {item, desc, contract, sizes: {label: qty}}."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]

    items, current = [], None
    for i, row in enumerate(rows):
        for c, cell in enumerate(row):
            if not isinstance(cell, str):
                continue
            key = cell.strip()
            if key == "Item Number":
                val = _next_nonempty(row, c + 1)
                if not val:
                    continue
                val_s = str(val).strip()
                if current and current.get("item"):
                    items.append(current)
                current = {"item": val_s, "desc": "",
                           "contract": "", "sizes": OrderedDict()}
            elif key == "Item Description" and current is not None:
                val = _next_nonempty(row, c + 1)
                current["desc"] = str(val or "").strip()

        if current is None:
            continue

        for c, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip() == "Contract No":
                if i + 1 < len(rows):
                    nxt = rows[i + 1]
                    if c < len(nxt) and nxt[c]:
                        v = str(nxt[c]).strip()
                        if v and not current["contract"]:
                            current["contract"] = v

        size_cols = {}
        for c, cell in enumerate(row):
            if isinstance(cell, str):
                m = SIZE_RE.match(cell)
                if m:
                    size_cols[c] = f"{int(m.group(1))} Yrs"
        if size_cols:
            for j in range(i + 1, min(i + 6, len(rows))):
                qrow = rows[j]
                if any(isinstance(x, str) and x.strip() == "Quantity" for x in qrow):
                    for c, lbl in size_cols.items():
                        v = qrow[c] if c < len(qrow) else None
                        if isinstance(v, (int, float)):
                            current["sizes"][lbl] = int(v)
                    break

    if current and current.get("item"):
        items.append(current)

    for it in items:
        it["sizes"] = OrderedDict(
            sorted(it["sizes"].items(), key=lambda kv: _size_sort_key(kv[0])))
    return [it for it in items if it["sizes"]]


# ============================================================
# CARTON PLAN
# ============================================================
def build_plan(item, pack, gross, tare, full_dims, part_dims):
    rows, cno = [], 1
    Lf, Wf, Hf = full_dims
    Lp, Wp, Hp = part_dims
    for sz, qty in item["sizes"].items():
        p = max(int(pack.get(sz, DEFAULT_PACK_PER_CTN)), 1)
        full, rem = divmod(int(qty), p)
        g_full = float(gross.get(sz, DEFAULT_GROSS_PER_CTN))
        t      = float(tare.get(sz, DEFAULT_TARE))

        if full:
            rows.append({
                "from": cno, "to": cno + full - 1, "size": sz,
                "n_ctn": full, "pcs_per_ctn": p,
                "L": Lf, "W": Wf, "H": Hf,
                "m3_per_ctn": Lf * Wf * Hf / 1_000_000.0,
                "gross_per_ctn": g_full,
                "net_per_ctn": round(g_full - t, 2),
            })
            cno += full

        if rem:
            g_part = round(g_full - (p - rem) * 0.07, 2)
            g_part = max(g_part, max(round(t + 0.1, 2), 1.0))
            rows.append({
                "from": cno, "to": cno, "size": sz,
                "n_ctn": 1, "pcs_per_ctn": rem,
                "L": Lp, "W": Wp, "H": Hp,
                "m3_per_ctn": Lp * Wp * Hp / 1_000_000.0,
                "gross_per_ctn": g_part,
                "net_per_ctn": round(max(g_part - t, 0.1), 2),
            })
            cno += 1
    return rows


def rows_from_edit(recs, item):
    """Edited carton table -> carton rows (numbered). recs keys: size, ctn, pcs, L, W, H, gross, net."""
    rows, problems, cno = [], [], 1
    for n, r in enumerate(recs, 1):
        sz = r.get("size")
        try:
            ctn, pcs = int(float(r.get("ctn"))), int(float(r.get("pcs")))
            L, W, H = (float(r.get(k)) for k in ("L", "W", "H"))
            gross, net = float(r.get("gross")), float(r.get("net"))
        except (TypeError, ValueError):
            if sz:
                problems.append(f"Row {n}: cartons, pieces, size and weights must be numbers - row ignored.")
            continue
        if sz not in item["sizes"]:
            problems.append(f"Row {n}: pick a size from the list - row ignored.")
            continue
        if ctn <= 0 or pcs <= 0 or min(L, W, H) <= 0:
            problems.append(f"Row {n}: cartons, pieces and measurements must be above 0 - row ignored.")
            continue
        rows.append({"from": cno, "to": cno + ctn - 1, "size": sz, "n_ctn": ctn, "pcs_per_ctn": pcs,
                     "L": L, "W": W, "H": H, "m3_per_ctn": L * W * H / 1_000_000.0,
                     "gross_per_ctn": gross, "net_per_ctn": net})
        cno += ctn
    return rows, problems


# ============================================================
# XLSX — write one packing list into an existing worksheet
# ============================================================
def fill_sheet(ws, head, item, rows):
    """Fill `ws` with the NEXT standard bulk-carton packing list."""
    TNR, ARI = "Times New Roman", "Arial"
    thin = Side(style="thin")
    CEN  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    def put(r, c, v=None, name=ARI, sz=10, al=CEN, fmt=None, bold=False):
        cell = ws.cell(row=r, column=c)
        if v is not None:
            cell.value = v
        cell.font = Font(name=name, size=sz, bold=bold)
        cell.alignment = al
        if fmt:
            cell.number_format = fmt
        return cell

    def merge(r1, c1, r2, c2):
        if (r1, c1) != (r2, c2):
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    def box(r1, c1, r2, c2):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).border = Border(left=thin, right=thin,
                                                         top=thin, bottom=thin)

    # ---------- header ----------
    put(1, 1, "PACKING LIST", TNR, 13, al=LEFT, bold=True)
    put(3, 1, "Agent / Supplier:", bold=True); put(3, 3, head["supplier"])
    put(4, 1, "Factory name: (Full name)", bold=True); put(4, 3, head["factory"])
    put(5, 1, "Factory address:", bold=True)
    put(5, 3, head["factory_address"], al=LEFT); merge(5, 3, 6, 8)

    put(3, 10, "Consignee Full name:", bold=True); put(3, 12, "Next Retail Limited")
    put(4, 10, "Consignee Full address:", bold=True)
    put(4, 12, "Desford Road, Enderby, Leicester, LE19 4AT, UK")
    put(5, 10, "Delivery to:", bold=True)
    put(5, 12, "Nominated forwarder in origin territory")

    put(8, 1, "NEXT Contract Number:", bold=True); put(8, 4, head["contract"])
    put(9, 1, "NEXT Item Number:", bold=True);     put(9, 4, item["item"])

    if head.get("pack_of"):
        put(10, 1, f"Pieces per Pack : {head['pack_of']}", bold=True)

    put(8, 10, "PACKING METHOD", bold=True)
    put(9, 10, "BULK CARTON - Single size/option per carton", bold=True)
    put(9, 15, "√", bold=True)
    put(10, 10, "PLEASE COMPLETE ALL COLUMNS & SIZE SUMMARY", bold=True)

    put(12, 1, "Full Description and composition:", bold=True)
    put(13, 1, "Product type :", bold=True);     put(13, 3, head["product_type"])
    put(14, 1, "Item Composition:", bold=True);  put(14, 3, head["composition"])
    put(15, 1, "Knitted or Woven:", bold=True);  put(15, 3, head["knit_woven"])

    total_pcs = sum(r["n_ctn"] * r["pcs_per_ctn"] for r in rows)
    total_net = sum(r["n_ctn"] * r["net_per_ctn"] for r in rows)

    put(12, 10, "Total Net Weight (kg):", bold=True);   put(12, 12, round(total_net, 2))
    put(13, 10, "Total pieces (Column G):", bold=True); put(13, 12, total_pcs)

    put(17, 1, "Country of Origin:", bold=True); put(17, 3, head["origin"])
    put(18, 1, "Preference available for this item?", bold=True); put(18, 3, "Yes")
    put(17, 10, "Date of Issue:", bold=True);     put(17, 12, head["date"])

    # ---------- carton table ----------
    R_HDR = 20
    R_D0  = R_HDR + 2

    put(R_HDR, 1, "Carton Numbers", bold=True); merge(R_HDR, 1, R_HDR, 2)
    put(R_HDR, 3, "Option", bold=True)
    put(R_HDR, 4, "Size", bold=True)
    put(R_HDR, 5, "Number of Cartons", bold=True)
    put(R_HDR, 6, "Pieces* per Carton", bold=True)
    put(R_HDR, 7, "Total pieces* supplied", bold=True)
    put(R_HDR, 8, "Carton Dimensions", bold=True); merge(R_HDR, 8, R_HDR, 10)
    put(R_HDR, 11, "Cubic metres (m3)", bold=True)
    put(R_HDR, 12, "Gross Weight (kg)", bold=True)
    put(R_HDR + 1, 1, "From", bold=True); put(R_HDR + 1, 2, "To", bold=True)
    put(R_HDR + 1, 8, "Length (cm)", bold=True)
    put(R_HDR + 1, 9, "Width (cm)",  bold=True)
    put(R_HDR + 1, 10, "Height (cm)", bold=True)
    box(R_HDR, 1, R_HDR + 1, 12)

    r = R_D0
    for row in rows:
        put(r, 1, row["from"], fmt="0")
        put(r, 2, row["to"],   fmt="0")
        put(r, 3, item["item"])
        put(r, 4, row["size"])
        put(r, 5, row["n_ctn"], fmt="0")
        put(r, 6, row["pcs_per_ctn"], fmt="0")
        put(r, 7, row["n_ctn"] * row["pcs_per_ctn"], fmt="0")
        put(r, 8, row["L"], fmt="0")
        put(r, 9, row["W"], fmt="0")
        put(r, 10, row["H"], fmt="0")
        put(r, 11, round(row["m3_per_ctn"] * row["n_ctn"], 4), fmt="0.0000")
        put(r, 12, round(row["gross_per_ctn"] * row["n_ctn"], 2), fmt="0.00")
        r += 1

    R_TOT = r
    put(R_TOT, 1, "TOTALS", bold=True); merge(R_TOT, 1, R_TOT, 4)
    put(R_TOT, 5,  f"=SUM(E{R_D0}:E{R_TOT-1})", bold=True)
    put(R_TOT, 7,  f"=SUM(G{R_D0}:G{R_TOT-1})", bold=True)
    put(R_TOT, 11, f"=SUM(K{R_D0}:K{R_TOT-1})", bold=True, fmt="0.0000")
    put(R_TOT, 12, f"=SUM(L{R_D0}:L{R_TOT-1})", bold=True, fmt="0.00")
    box(R_HDR, 1, R_TOT, 12)

    # ---------- size summary ----------
    R_SUM = R_TOT + 3
    put(R_SUM, 1, "SIZE", bold=True)
    sizes_ord = list(item["sizes"].keys())
    for i, sz in enumerate(sizes_ord):
        put(R_SUM, 3 + i, sz, bold=True)
    put(R_SUM, 3 + len(sizes_ord), "TOTAL", bold=True)

    r_ord, r_ship, r_diff, r_pct = R_SUM + 1, R_SUM + 2, R_SUM + 3, R_SUM + 4
    put(r_ord,  1, "Order Qty",    bold=True, al=LEFT)
    put(r_ship, 1, "Shipped Qty",  bold=True, al=LEFT)
    put(r_diff, 1, "Ex/Short Qty", bold=True, al=LEFT)
    put(r_pct,  1, "% Variance",   bold=True, al=LEFT)

    shipped = {sz: 0 for sz in sizes_ord}
    for row in rows:
        shipped[row["size"]] += row["n_ctn"] * row["pcs_per_ctn"]

    for i, sz in enumerate(sizes_ord):
        c = 3 + i
        put(r_ord,  c, item["sizes"][sz])
        put(r_ship, c, shipped[sz])
        put(r_diff, c, f"={L(c)}{r_ship}-{L(c)}{r_ord}")
        put(r_pct,  c, f"=IFERROR({L(c)}{r_diff}/{L(c)}{r_ord},0)", fmt="0.0%")
    last = 2 + len(sizes_ord)
    put(r_ord,  last + 1, f"=SUM(C{r_ord}:{L(last)}{r_ord})",   bold=True)
    put(r_ship, last + 1, f"=SUM(C{r_ship}:{L(last)}{r_ship})", bold=True)
    put(r_diff, last + 1, f"=SUM(C{r_diff}:{L(last)}{r_diff})", bold=True)
    box(R_SUM, 1, r_pct, last + 1)

    # ---------- signature ----------
    R_SIG = r_pct + 3
    put(R_SIG, 1, "Signed by (Factory)", bold=True, al=LEFT)
    put(R_SIG, 8, "Factory stamp", bold=True, al=LEFT)
    put(R_SIG + 2, 1, "Town/City, Country", bold=True, al=LEFT)
    put(R_SIG + 4, 1, "Date", bold=True, al=LEFT)

    widths = {1: 22, 2: 6, 3: 14, 4: 8, 5: 12, 6: 12, 7: 14,
              8: 10, 9: 10, 10: 10, 11: 12, 12: 14}
    for c, w in widths.items():
        ws.column_dimensions[L(c)].width = w

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top  = ws.page_margins.bottom = 0.25


def build_workbook(head, item, rows):
    """Convenience wrapper: one-item workbook (kept for standalone use)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "PACKING LIST"
    fill_sheet(ws, head, item, rows)
    return wb


# ============================================================
# STREAMLIT RENDER HELPERS
# ============================================================
def _render_order_table(item):
    """Compact Order Qty table: sizes as columns, single QTY row."""
    st.markdown("**Order quantity**")

    sizes_ord = list(item["sizes"].keys())
    header_row = ["QTY"] + [item["sizes"][sz] for sz in sizes_ord] + \
                 [sum(item["sizes"].values())]
    df = pd.DataFrame(
        [header_row],
        columns=["SIZE"] + sizes_ord + ["TOTAL"],
    )

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


def _render_header_block(head, item, rows):
    """Show the header block from the Excel sheet on the web page."""
    total_pcs = sum(r["n_ctn"] * r["pcs_per_ctn"] for r in rows)
    total_net = sum(r["n_ctn"] * r["net_per_ctn"] for r in rows)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### Supplier / Factory")
        st.markdown(f"**Agent / Supplier:**  {head['supplier']}")
        st.markdown(f"**Factory name:**  {head['factory']}")
        st.markdown(f"**Factory address:**  {head['factory_address']}")

        st.markdown("##### Order")
        st.markdown(f"**NEXT Contract Number:**  {head['contract']}")
        st.markdown(f"**NEXT Item Number:**  {item['item']}")
        if head.get("pack_of"):
            st.markdown(f"**Pieces per Pack:**  {head['pack_of']}")

        st.markdown("##### Product")
        st.markdown(f"**Product type:**  {head['product_type']}")
        st.markdown(f"**Item Composition:**  {head['composition']}")
        st.markdown(f"**Knitted or Woven:**  {head['knit_woven']}")

    with c2:
        st.markdown("##### Consignee")
        st.markdown("**Consignee Full name:**  Next Retail Limited")
        st.markdown("**Consignee Full address:**  Desford Road, Enderby, "
                    "Leicester, LE19 4AT, UK")
        st.markdown("**Delivery to:**  Nominated forwarder in origin territory")

        st.markdown("##### Packing Method")
        st.markdown("**BULK CARTON – Single size/option per carton**  ✓")
        st.caption("Please complete all columns & size summary.")

        st.markdown("##### Summary")
        st.markdown(f"**Total Net Weight (kg):**  {round(total_net, 2)}")
        st.markdown(f"**Total pieces:**  {total_pcs:,}")
        st.markdown(f"**Country of Origin:**  {head['origin']}")
        st.markdown(f"**Date of Issue:**  {head['date']}")


def _render_carton_table(item, rows):
    """Carton table with Excel-style multi-tier headers."""
    st.markdown("##### Carton Details")

    table_rows = []
    for r in rows:
        table_rows.append({
            "From":       r["from"],
            "To":         r["to"],
            "Option":     item["item"],
            "Size":       r["size"],
            "No. Ctns":   r["n_ctn"],
            "Pcs/Ctn":    r["pcs_per_ctn"],
            "Total Pcs":  r["n_ctn"] * r["pcs_per_ctn"],
            "L (cm)":     r["L"],
            "W (cm)":     r["W"],
            "H (cm)":     r["H"],
            "m³":         round(r["m3_per_ctn"] * r["n_ctn"], 4),
            "GW (kg)":    round(r["gross_per_ctn"] * r["n_ctn"], 2),
        })

    table_rows.append({
        "From":       "TOTALS",
        "To":         "",
        "Option":     "",
        "Size":       "",
        "No. Ctns":   sum(r["n_ctn"] for r in rows),
        "Pcs/Ctn":    "",
        "Total Pcs":  sum(r["n_ctn"] * r["pcs_per_ctn"] for r in rows),
        "L (cm)":     "",
        "W (cm)":     "",
        "H (cm)":     "",
        "m³":         round(sum(r["m3_per_ctn"] * r["n_ctn"] for r in rows), 4),
        "GW (kg)":    round(sum(r["gross_per_ctn"] * r["n_ctn"] for r in rows), 2),
    })

    df = pd.DataFrame(table_rows)
    left_cols = ["From", "To", "Option", "Size"]
    styler = (
        df.style
        .set_properties(subset=left_cols, **{"text-align": "left"})
        .set_properties(
            subset=[c for c in df.columns if c not in left_cols],
            **{"text-align": "center"},
        )
        .apply(lambda r: [
            "font-weight: bold; background-color: #F0F0F0"
            if r["From"] == "TOTALS" else ""
            for _ in r
        ], axis=1)
    )
    st.dataframe(styler, use_container_width=True, hide_index=True)


def _render_size_summary(item, rows):
    """SIZE SUMMARY table below the carton detail."""
    st.markdown("##### Size Summary")

    sizes_ord = list(item["sizes"].keys())
    shipped = {sz: 0 for sz in sizes_ord}
    for r in rows:
        shipped[r["size"]] += r["n_ctn"] * r["pcs_per_ctn"]

    order_row    = {"SIZE": "Order Qty"}
    ship_row     = {"SIZE": "Shipped Qty"}
    diff_row     = {"SIZE": "Ex/Short Qty"}
    pct_row      = {"SIZE": "% Variance"}

    for sz in sizes_ord:
        order_row[sz] = item["sizes"][sz]
        ship_row[sz]  = shipped[sz]
        diff_row[sz]  = shipped[sz] - item["sizes"][sz]
        o = item["sizes"][sz]
        pct_row[sz] = 0.0 if o == 0 else round((shipped[sz] - o) / o, 4)

    for row in (order_row, ship_row, diff_row, pct_row):
        row["TOTAL"] = sum(row[sz] for sz in sizes_ord)

    df = pd.DataFrame([order_row, ship_row, diff_row, pct_row])

    display_df = df.copy()
    for sz in sizes_ord + ["TOTAL"]:
        display_df[sz] = df[sz].apply(
            lambda v: f"{v:.2%}" if isinstance(v, float) else str(v)
        )

    styler = (
        display_df.style
        .set_properties(subset=["SIZE"],
                        **{"text-align": "left", "font-weight": "bold"})
        .set_properties(
            subset=[c for c in display_df.columns if c != "SIZE"],
            **{"text-align": "center"},
        )
    )
    st.dataframe(styler, use_container_width=True, hide_index=True)


# ============================================================
# Streamlit UI
# ============================================================
@st.cache_data(show_spinner=False, max_entries=5)
def _parse_cached(b: bytes):
    return parse_next_xlsx(b)


def _item_ui(item, key):
    tot = sum(item["sizes"].values())
    kit.step(1, f"Item {item['item']}")
    kit.summary_cards([("Item", str(item["item"])), ("Contract", str(item.get("contract") or "?")),
                       ("Sizes", f"{len(item['sizes'])}"), ("Pieces", f"{tot:,}")])
    st.caption(str(item["desc"]))
    _render_order_table(item)

    sizes = list(item["sizes"].keys())

    kit.step(2, "Variables", "Pieces per carton, weights, carton sizes and header texts.")
    with st.expander("Packing configuration", expanded=True):
        cols = st.columns(min(len(sizes), 6))
        pack, gross, tare = {}, {}, {}
        for i, sz in enumerate(sizes):
            with cols[i % len(cols)]:
                st.markdown(f"**{sz}**")
                pack[sz]  = st.number_input("Pcs/ctn", 1, 500,
                                            DEFAULT_PACK_PER_CTN, 1,
                                            key=f"{key}_pk_{sz}")
                gross[sz] = st.number_input("GW/ctn kg", 0.5, 50.0,
                                            DEFAULT_GROSS_PER_CTN, 0.1,
                                            key=f"{key}_gr_{sz}")
                tare[sz]  = st.number_input("Tare kg", 0.0, 10.0,
                                            DEFAULT_TARE, 0.1,
                                            key=f"{key}_tr_{sz}")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Full carton dims (cm)**")
            fL = st.number_input("Full L", 10, 200, DEFAULT_FULL_L, 1, key=f"{key}_fL")
            fW = st.number_input("Full W", 10, 200, DEFAULT_FULL_W, 1, key=f"{key}_fW")
            fH = st.number_input("Full H", 10, 200, DEFAULT_FULL_H, 1, key=f"{key}_fH")
        with c2:
            st.markdown("**Partial carton dims (cm)**")
            pL = st.number_input("Part L", 10, 200, DEFAULT_PART_L, 1, key=f"{key}_pL")
            pW = st.number_input("Part W", 10, 200, DEFAULT_PART_W, 1, key=f"{key}_pW")
            pH = st.number_input("Part H", 10, 200, DEFAULT_PART_H, 1, key=f"{key}_pH")

        c1, c2, c3 = st.columns(3)
        with c1:
            product_type = st.text_input("Product type", "Boys School Shirt",
                                         key=f"{key}_pt")
        with c2:
            composition = st.text_input("Composition",
                                        "65% Recycled polyester, 35% Cotton",
                                        key=f"{key}_comp")
        with c3:
            knit_woven = st.text_input("Knitted or Woven", "Woven", key=f"{key}_kw")

        c4, c5 = st.columns(2)
        with c4:
            pack_of = st.text_input(
                "Pieces per Pack (printed on the sheet; clear to hide)",
                DEFAULT_PACK_OF, key=f"{key}_pop")
        with c5:
            contract_suffix = st.text_input(
                "Contract suffix (appended after contract #)",
                "-1", key=f"{key}_cs")

    rows = build_plan(item, pack, gross, tare,
                      (fL, fW, fH), (pL, pW, pH))

    head = {
        "supplier":         "Clover Collections",
        "factory":          "CREATIVE COLLECTIONS LTD.",
        "factory_address":  "69,107 NISHAT NAGAR, TONGI, GAZIPUR-1711, BANGLADESH",
        "contract":         f"{item.get('contract','')}{contract_suffix}",
        "product_type":     product_type,
        "composition":      composition,
        "knit_woven":       knit_woven,
        "origin":           "Bangladesh",
        "date":             date.today(),
        "pack_of":          pack_of.strip() or None,
    }

    # ---- 3 carton table (editable)
    kit.step(3, "Carton table",
             "Made from the settings above. Edit any cell, add a row with +, or tick 🗑 Delete and press "
             "“Remove ticked rows”. Changing the settings above rebuilds this table.")
    import hashlib
    sig = hashlib.md5(repr((sorted(pack.items()), sorted(gross.items()), sorted(tare.items()),
                            (fL, fW, fH), (pL, pW, pH))).encode()).hexdigest()[:10]
    base = pd.DataFrame([{"Size": r["size"], "Ctns": r["n_ctn"], "Pcs/Ctn": r["pcs_per_ctn"], "L (cm)": r["L"],
                          "W (cm)": r["W"], "H (cm)": r["H"], "Gross/Ctn": r["gross_per_ctn"],
                          "Net/Ctn": r["net_per_ctn"]} for r in rows])
    cfg = {"Size": st.column_config.SelectboxColumn("Size", options=sizes, required=True),
           "Ctns": st.column_config.NumberColumn("Ctns", min_value=1, step=1, format="%d"),
           "Pcs/Ctn": st.column_config.NumberColumn("Pcs/Ctn", min_value=1, step=1, format="%d"),
           "L (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
           "W (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
           "H (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
           "Gross/Ctn": st.column_config.NumberColumn(min_value=0.0, step=0.1, format="%.2f"),
           "Net/Ctn": st.column_config.NumberColumn(min_value=0.0, step=0.1, format="%.2f")}
    ed = kit.editable_grid(base, f"{key}_ct_{sig}", column_config=cfg)
    recs = [{"size": r["Size"], "ctn": r["Ctns"], "pcs": r["Pcs/Ctn"], "L": r["L (cm)"], "W": r["W (cm)"],
             "H": r["H (cm)"], "gross": r["Gross/Ctn"], "net": r["Net/Ctn"]} for r in ed.to_dict("records")]
    rows, probs = rows_from_edit(recs, item)
    kit.problems(probs)
    if not rows:
        st.error("The carton table is empty. Press “Reset table to PO” or add a row.")
        return head, item, []
    _render_carton_table(item, rows)

    # ---- 4 check
    kit.step(4, "Check against the PO")
    packed = {sz: sum(r["n_ctn"] * r["pcs_per_ctn"] for r in rows if r["size"] == sz) for sz in sizes}
    kit.check_by_size(sizes, dict(item["sizes"]), packed)
    tot_ctn = sum(r["n_ctn"] for r in rows)
    tot_pcs = sum(r["n_ctn"] * r["pcs_per_ctn"] for r in rows)
    tot_m3  = sum(r["n_ctn"] * r["m3_per_ctn"]   for r in rows)
    tot_gw  = sum(r["n_ctn"] * r["gross_per_ctn"] for r in rows)
    tot_nw  = sum(r["n_ctn"] * r["net_per_ctn"]  for r in rows)
    kit.summary_cards([("Cartons", f"{tot_ctn:,}"), ("Pieces", f"{tot_pcs:,}"), ("m³", f"{tot_m3:.3f}"),
                       ("Gross kg", f"{tot_gw:,.1f}"), ("Net kg", f"{tot_nw:,.1f}")])
    with st.expander("Size summary"):
        _render_size_summary(item, rows)

    return head, item, rows


def run(xlsx_file=None):
    st.title("Next Retail / Clover — Packing List")
    if xlsx_file is None:
        st.info("Upload a Contract Confirmation Report (.xlsx) from Next ePLM.")
        return

    try:
        items = _parse_cached(xlsx_file.getvalue())
    except Exception as e:
        st.error(f"Could not read the XLSX: {e}")
        return

    if not items:
        st.warning("No item blocks found — check the file is a Contract Confirmation Report.")
        return

    st.success(f"Parsed **{len(items)}** item(s) from `{xlsx_file.name}`")

    results = []
    tabs = st.tabs([f"{it['item']} — {it['desc'][:22]}" for it in items])
    for tab, it in zip(tabs, items):
        with tab:
            results.append(_item_ui(it, key=f"nx_{it['item']}"))

    combined = Workbook()
    combined.remove(combined.active)
    for head, item, rows in results:
        sheet_name = re.sub(r"[\[\]:*?/\\]", "-", str(item["item"]))[:31] or "Sheet1"
        ws = combined.create_sheet(title=sheet_name)
        fill_sheet(ws, head, item, rows)

    kit.step(5, "Download", "One Excel file with a sheet for every item.")
    buf = io.BytesIO(); combined.save(buf)
    kit.download_block("⬇ Download packing list (Excel)", buf.getvalue(),
                       f"packing_list_{items[0].get('contract','contract')}.xlsx")


if __name__ == "__main__":
    st.set_page_config(page_title="Next / Clover Packing List", layout="wide")
    run(st.file_uploader("Upload Contract Confirmation Report (.xlsx)", type=["xlsx"]))