"""
Justice (Sports Products of America / Adjmi) - buyer page, shared design (pl_kit).

  1 PO summary -> 2 Packing table (editable, add / delete rows) -> 3 Variables
  -> 4 Live check (order vs packed) -> 5 Download
"""
from __future__ import annotations

import dataclasses
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:                      # pl_kit.py lives in the project root
    sys.path.insert(0, _ROOT)

import pl_kit as kit                           # noqa: E402
import justice_packing as jp                   # noqa: E402


@st.cache_data(show_spinner="Reading PO…")
def _parse(data: bytes):
    return jp.parse_po(data)


def render(default_file=None) -> None:
    st.header("Justice (SPA / Adjmi) — PO → Packing list")

    data = None
    if default_file is not None:
        if hasattr(default_file, "getvalue"):
            data = default_file.getvalue()
        elif isinstance(default_file, (bytes, bytearray)):
            data = bytes(default_file)
        elif isinstance(default_file, (str, Path)) and Path(default_file).exists():
            data = Path(default_file).read_bytes()
    if data is None:
        up = st.file_uploader("Justice PO (.pdf)", type=["pdf"], key="jus_upload")
        if up is not None:
            data = up.getvalue()
    if data is None:
        st.info("Upload the Justice PO PDF to start.")
        return

    try:
        po = _parse(data)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the PO: {e}")
        return
    kit.problems(po.warnings)
    if not po.lines:
        st.error("No PO lines found. Confirm this is a Justice / SPA purchase order PDF.")
        return

    k = f"jus_{po.po_no}_{po.ref_no}"
    sizes = po.sizes

    # ---------------------------------------------------------------- 1
    kit.step(1, "PO summary")
    kit.summary_cards([
        ("PO", po.po_no), ("Style", po.style), ("Ship date", po.ship_date),
        ("Cartons", f"{po.total_ctn:,}"), ("Pieces", f"{po.total_units:,}"),
    ])
    st.caption(f"{po.desc} · {po.color} · ref {po.ref_no} · {po.ship_type}")

    # ---------------------------------------------------------------- 2
    kit.step(2, "Packing table",
             "Click any cell to change it. Use the + row at the bottom to add a pack. "
             "To delete: tick 🗑 Delete on the row(s), then press “Remove ticked rows”. "
             "Size columns = pieces per carton.")
    base = pd.DataFrame(jp.rows_from_po(po))
    cfg = {
        jp.COL_LABEL: st.column_config.TextColumn("Label", width="small"),
        jp.COL_PACK: st.column_config.TextColumn("Pack", width="small"),
        jp.COL_CTNS: st.column_config.NumberColumn("Ctns", min_value=0, step=1, format="%d"),
        jp.COL_L: st.column_config.NumberColumn(min_value=1.0, step=1.0),
        jp.COL_W: st.column_config.NumberColumn(min_value=1.0, step=1.0),
        jp.COL_H: st.column_config.NumberColumn(min_value=1.0, step=1.0),
        jp.COL_NET: st.column_config.NumberColumn(min_value=0.0, step=0.01, format="%.3f"),
    }
    for s in sizes:
        cfg[s] = st.column_config.NumberColumn(s, min_value=0, step=1, format="%d")
    edited = kit.editable_grid(base, k, column_config=cfg)

    lines, dims, net_pc, probs = jp.lines_from_rows(edited.to_dict("records"), sizes)
    kit.problems(probs)
    if not lines:
        st.error("The table has no valid rows. Press “Reset table to PO” or add a row.")
        return

    # ---------------------------------------------------------------- 3
    kit.step(3, "Variables", "Used for the carton plan and the Excel header. Dimensions and "
                             "net weight per piece are in the table above (yellow cells in Excel).")
    c1, c2, c3 = st.columns([1, 2, 3])
    tare = float(c1.number_input("Carton tare (kg)", min_value=0.0, value=float(jp.DEF_TARE),
                                 step=0.1, key=f"{k}_tare"))
    company = c2.text_input("Company header", key=f"{k}_co", placeholder="Factory name (optional)")
    address = c3.text_input("Address", key=f"{k}_addr", placeholder="Factory address (optional)")

    po_e = dataclasses.replace(po, lines=lines)           # order_qty stays as printed on the PO
    plan = jp.build_carton_plan(po_e, dims, net_pc, tare)

    # ---------------------------------------------------------------- carton plan (read-only result)
    st.markdown("**Carton plan**  (calculated from your table)")
    rows = []
    for r in plan:
        rows.append({"From": r.ctn_from, "To": r.ctn_to, "Ctns": r.ctns,
                     "Label": f"{r.line.label} ({r.line.dim_pk})", "Pcs/Ctn": r.pcs_per_ctn,
                     "Total pcs": r.total_pcs, "CBM": r.cbm,
                     "Net/Ctn": r.nw_ctn, "Gross/Ctn": r.gw_ctn})
    st.dataframe(pd.DataFrame(rows), hide_index=True, **kit.STRETCH)

    tot_ctn = sum(r.ctns for r in plan)
    tot_pcs = sum(r.total_pcs for r in plan)
    tot_cbm = round(sum(r.cbm for r in plan), 3)
    tot_nt = round(sum(r.nw_ctn * r.ctns for r in plan), 2)
    tot_gr = round(sum(r.gw_ctn * r.ctns for r in plan), 2)
    kit.summary_cards([
        ("Cartons", f"{tot_ctn:,}"), ("Pieces", f"{tot_pcs:,}"),
        ("CBM", f"{tot_cbm:,}"), ("Net kg", f"{tot_nt:,.0f}"), ("Gross kg", f"{tot_gr:,.0f}"),
    ])

    # ---------------------------------------------------------------- 4
    kit.step(4, "Check against the PO")
    ok = kit.check_by_size(po_e.sizes, po.order_by_size, po_e.packed_by_size())

    # ---------------------------------------------------------------- 5
    kit.step(5, "Download")
    if not ok:
        st.caption("The table differs from the PO. You can still download; the Excel size summary shows the difference.")
    xlsx = jp.make_packing_excel(po_e, plan, company=company, address=address)
    kit.download_block("⬇ Download packing list (Excel)", xlsx,
                       f"PACKING_LIST_{po.style}_{po.po_no}.xlsx",
                       f"{tot_ctn:,} cartons · {tot_pcs:,} pcs · CBM {tot_cbm}")
