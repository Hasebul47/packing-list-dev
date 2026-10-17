"""
Justice (Sports Products of America / Adjmi) buyer page for the unified router.

Page order:
  1. PO lines
  2. Carton inputs (measurement / weight)  -> main packing table
  3. Summary by size
  4. Download Excel
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

import justice_packing as jp

_V = tuple(int(x) for x in st.__version__.split(".")[:2])
_STRETCH = {"width": "stretch"} if _V >= (1, 50) else {"use_container_width": True}


@st.cache_data(show_spinner="Reading PO…")
def _parse(data: bytes):
    return jp.parse_po(data)


def render(default_file=None) -> None:
    st.header("Justice (SPA / Adjmi) — PO (PDF) → Packing list")

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
    for w in po.warnings:
        st.warning(w)
    if not po.lines:
        st.error("No PO lines found. Confirm this is a Justice / SPA purchase order PDF.")
        return

    k = f"{po.po_no}_{po.ref_no}"
    st.caption(
        f"PO **{po.po_no}** · ref **{po.ref_no}** · style **{po.style}** · {po.desc} · "
        f"{po.color} · {po.total_ctn:,} ctn · {po.total_units:,} pcs · ship {po.ship_date}"
    )

    # ------------------------------------------------------------ 1 PO lines
    st.subheader("1 · PO lines")
    sizes = po.sizes
    rows = []
    for ln in po.lines:
        r = {"Label": ln.label, "Pack": ln.dim_pk, "Type": ln.kind,
             "Pcs/Ctn": ln.pcs_per_ctn, "Ctns": ln.ctns, "Units": ln.units}
        for s in sizes:
            r[s] = ln.ratio[ln.sizes.index(s)] if s in ln.sizes else None
        rows.append(r)
    st.dataframe(pd.DataFrame(rows), hide_index=True,
                 height=min(320, 40 + 35 * len(rows)), **_STRETCH)

    # ------------------------------------------------------------ 2 inputs
    st.subheader("2 · Carton measurement & weight")
    c1, c2 = st.columns(2)
    tare = float(c1.number_input("Carton tare (kg)", min_value=0.0, value=float(jp.DEF_TARE),
                                 step=0.1, key=f"jus_tare_{k}"))
    company = c2.text_input("Company header", value="", key=f"jus_co_{k}",
                            placeholder="Factory name (optional)")
    address = st.text_input("Address", value="", key=f"jus_addr_{k}",
                            placeholder="Factory address (optional)")

    inp = pd.DataFrame([{
        "Label": ln.label, "Pack": ln.dim_pk, "Ctns": ln.ctns,
        "L (cm)": jp.DEF_DIM[0], "W (cm)": jp.DEF_DIM[1], "H (cm)": jp.DEF_DIM[2],
        "Net kg / pc": jp.DEF_NET,
    } for ln in po.lines])
    ed = st.data_editor(
        inp, hide_index=True, key=f"jus_inp_{k}",
        disabled=["Label", "Pack", "Ctns"],
        column_config={
            "L (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
            "W (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
            "H (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
            "Net kg / pc": st.column_config.NumberColumn(min_value=0.0, step=0.01, format="%.3f"),
        },
        **_STRETCH,
    )
    dims = {i: (float(r["L (cm)"]), float(r["W (cm)"]), float(r["H (cm)"]))
            for i, r in ed.reset_index(drop=True).iterrows()}
    net_pc = {i: float(r["Net kg / pc"]) for i, r in ed.reset_index(drop=True).iterrows()}
    plan = jp.build_carton_plan(po, dims, net_pc, tare)

    # ------------------------------------------------------------ main table
    st.subheader("3 · Packing table")
    table = []
    for r in plan:
        t = {"From": r.ctn_from, "To": r.ctn_to, "Ctns": r.ctns,
             "Label": f"{r.line.label} ({r.line.dim_pk})"}
        for s in sizes:
            t[s] = r.line.ratio[r.line.sizes.index(s)] if s in r.line.sizes else None
        t.update({"Pcs/Ctn": r.pcs_per_ctn, "Total Pcs": r.total_pcs,
                  "Meas (cm)": f"{r.L:g}×{r.W:g}×{r.H:g}", "CBM": r.cbm,
                  "Nt/Ctn": r.nw_ctn, "Gr/Ctn": r.gw_ctn,
                  "Total Nt": round(r.nw_ctn * r.ctns, 2),
                  "Total Gr": round(r.gw_ctn * r.ctns, 2)})
        table.append(t)
    st.dataframe(pd.DataFrame(table), hide_index=True, **_STRETCH)

    # ------------------------------------------------------------ summary
    st.subheader("4 · Summary")
    order = po.order_by_size
    packed = {s: sum(r.ctns * r.line.ratio[r.line.sizes.index(s)]
                     for r in plan if s in r.line.sizes) for s in sizes}
    tot_ctn = sum(r.ctns for r in plan)
    tot_pcs = sum(r.total_pcs for r in plan)
    tot_cbm = round(sum(r.cbm for r in plan), 3)
    tot_nt = round(sum(r.nw_ctn * r.ctns for r in plan), 2)
    tot_gr = round(sum(r.gw_ctn * r.ctns for r in plan), 2)
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Order qty", f"{po.total_units:,}")
    m2.metric("Packed qty", f"{tot_pcs:,}", delta=f"{tot_pcs - po.total_units:+d}")
    m3.metric("Cartons", f"{tot_ctn:,}")
    m4.metric("CBM", f"{tot_cbm:,}")
    m5.metric("Net / Gross kg", f"{tot_nt:,.0f} / {tot_gr:,.0f}")
    st.dataframe(pd.DataFrame({
        "Size": sizes,
        "Order": [order[s] for s in sizes],
        "Packed": [packed[s] for s in sizes],
        "Excess": [packed[s] - order[s] for s in sizes],
    }), hide_index=True, **_STRETCH)

    # ------------------------------------------------------------ download
    st.subheader("5 · Download")
    xlsx = jp.make_packing_excel(po, plan, company=company, address=address)
    fname = f"PACKING_LIST_{po.style}_{po.po_no}.xlsx"
    st.download_button(
        "⬇ Download packing list (Excel)", data=xlsx, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption(f"{tot_ctn:,} cartons · {tot_pcs:,} pcs · CBM {tot_cbm} · file `{fname}`")
