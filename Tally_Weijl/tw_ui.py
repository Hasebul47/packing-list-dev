"""
Tally Weijl buyer page (standalone or for the unified router).

    from tw_ui import render
    render()

Shared design (pl_kit): 1 PO summary -> 2 Quantities to pack -> 3 Variables -> 4 Carton table (editable)
-> 5 Check vs PO -> 6 Download Excel
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

import hashlib
import sys

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:                      # pl_kit.py lives in the project root
    sys.path.insert(0, _ROOT)

import pl_kit as kit                           # noqa: E402
import tw_packing as tw                        # noqa: E402

_V = tuple(int(x) for x in st.__version__.split(".")[:2])
_STRETCH = {"width": "stretch"} if _V >= (1, 50) else {"use_container_width": True}


@st.cache_data(show_spinner="Reading Order Form…")
def _parse(data: bytes):
    return tw.parse_po(data)


def render(default_file=None) -> None:
    st.header("Tally Weijl — Order Form (PDF) → Packing list")
    data = None
    if default_file is not None:
        if hasattr(default_file, "getvalue"):
            data = default_file.getvalue()
        elif isinstance(default_file, (bytes, bytearray)):
            data = bytes(default_file)
        elif isinstance(default_file, (str, Path)) and Path(default_file).exists():
            data = Path(default_file).read_bytes()
    if data is None:
        up = st.file_uploader("Tally Weijl Order Form (.pdf)", type=["pdf"], key="tw_upload")
        if up is not None:
            data = up.getvalue()
    if data is None:
        st.info("Upload the Tally Weijl Order Form PDF to start.")
        return
    try:
        po = _parse(data)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the Order Form: {e}")
        return
    kit.problems(po.warnings)
    if not po.lines:
        st.error("No size lines found. Confirm this is a Tally Weijl Order Form PDF.")
        return

    k = f"tw_{po.order_no}"
    sizes = po.sizes

    # ------------------------------------------------------------ 1
    kit.step(1, "PO summary")
    kit.summary_cards([("PO", po.order_no), ("Style", po.style), ("Ship date", po.ship_date or "—"),
                       ("Incoterm", po.incoterm or "—"), ("Pieces", f"{po.total_qty:,}")])
    st.caption(f"{po.color_code} {po.color_name} · {po.item_desc or ''} · {po.destination}")
    with st.expander("PO lines (EAN, price)"):
        st.dataframe(pd.DataFrame([{"Size": ln.size, "Color code": ln.color_code, "Color": ln.color_name,
                                    "EAN13": ln.ean, "Qty": ln.qty, "USD": ln.unit_usd} for ln in po.lines]),
                     hide_index=True, height=min(280, 40 + 35 * len(po.lines)), **kit.STRETCH)

    # ------------------------------------------------------------ 2
    kit.step(2, "Quantities to pack", "Change “Ship qty” to pack more or less than ordered.")
    sdf = pd.DataFrame({"Size": [ln.size for ln in po.lines], "Order qty": [ln.qty for ln in po.lines],
                        "Ship qty": [ln.qty for ln in po.lines]})
    sed = st.data_editor(sdf, hide_index=True, key=f"{k}_ship", disabled=["Size", "Order qty"], **kit.STRETCH,
                         column_config={"Ship qty": st.column_config.NumberColumn(min_value=0, step=1, format="%d")})
    ship = {str(r["Size"]): (int(r["Ship qty"]) if not pd.isna(r["Ship qty"]) else 0) for _, r in sed.iterrows()}

    # ------------------------------------------------------------ 3
    kit.step(3, "Variables")
    c1, c2, c3 = st.columns([1, 1, 3])
    std = int(c1.number_input("Pcs per full carton", min_value=1, value=int(tw.DEF_STD), step=1, key=f"{k}_std"))
    tare = float(c2.number_input("Carton tare (kg)", min_value=0.0, value=float(tw.DEF_TARE), step=0.1, key=f"{k}_tare"))
    c3.markdown(f"**Item**  \n{po.item_desc or '—'}")
    w1, w2 = st.columns([2, 3])
    with w1:
        wed = st.data_editor(pd.DataFrame({"Size": sizes, "Net wt / pc (kg)": [tw.DEF_NET] * len(sizes)}),
                             hide_index=True, key=f"{k}_w", disabled=["Size"], **kit.STRETCH,
                             column_config={"Net wt / pc (kg)": st.column_config.NumberColumn(
                                 min_value=0.0, step=0.01, format="%.3f")})
    with w2:
        st.caption("Full carton L × W × H, cm")
        d1, d2, d3 = st.columns(3)
        Lb = d1.number_input("L", min_value=1.0, value=tw.DEF_BIG[0], key=f"{k}_Lb")
        Wb = d2.number_input("W", min_value=1.0, value=tw.DEF_BIG[1], key=f"{k}_Wb")
        Hb = d3.number_input("H", min_value=1.0, value=tw.DEF_BIG[2], key=f"{k}_Hb")
        st.caption("Small (remainder) carton L × W × H, cm")
        s1, s2, s3 = st.columns(3)
        Ls = s1.number_input("L ", min_value=1.0, value=tw.DEF_BIG[0], key=f"{k}_Ls")
        Ws = s2.number_input("W ", min_value=1.0, value=tw.DEF_BIG[1], key=f"{k}_Ws")
        Hs = s3.number_input("H ", min_value=1.0, value=20.0, key=f"{k}_Hs")
    net_map = {str(r["Size"]): float(r["Net wt / pc (kg)"]) for _, r in wed.iterrows()}
    e1, e2 = st.columns(2)
    company = e1.text_input("Company header", value="THAT'S IT SPORT'S WEAR LTD.     ( UNIT = 04 )", key=f"{k}_co")
    address = e2.text_input("Address", value="Narashighpur, Ashulia, Dhaka-1341.", key=f"{k}_addr")

    auto = tw.build_carton_plan(tw.with_ship_qty(po, ship), std, (Lb, Wb, Hb), (Ls, Ws, Hs))

    # ------------------------------------------------------------ 4
    kit.step(4, "Carton table",
             "Made from the quantities above. Edit any cell, add a row with +, or tick 🗑 Delete and press "
             "“Remove ticked rows”. Changing the quantities or variables above rebuilds this table.")
    sig = hashlib.md5(repr((sorted(ship.items()), std, tare, (Lb, Wb, Hb), (Ls, Ws, Hs))).encode()).hexdigest()[:10]
    base = pd.DataFrame([{"Size": r.size, "Ctns": r.ctns, "Pcs/Ctn": r.pcs_per_ctn, "L (cm)": r.L, "W (cm)": r.W,
                          "H (cm)": r.H, "Carton": "Small" if r.is_remainder else "Full"} for r in auto])
    cfg = {
        "Size": st.column_config.SelectboxColumn("Size", options=list(sizes), required=True),
        "Ctns": st.column_config.NumberColumn("Ctns", min_value=1, step=1, format="%d"),
        "Pcs/Ctn": st.column_config.NumberColumn("Pcs/Ctn", min_value=1, step=1, format="%d"),
        "L (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
        "W (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
        "H (cm)": st.column_config.NumberColumn(min_value=1.0, step=1.0),
        "Carton": st.column_config.SelectboxColumn("Carton", options=["Full", "Small"]),
    }
    ed = kit.editable_grid(base, f"{k}_{sig}", column_config=cfg)
    rows_in = [{"size": r["Size"], "ctns": r["Ctns"], "pcs": r["Pcs/Ctn"], "L": r["L (cm)"], "W": r["W (cm)"],
                "H": r["H (cm)"], "small": r["Carton"] == "Small"} for r in ed.to_dict("records")]
    plan, probs = tw.plan_from_rows(po, rows_in)
    kit.problems(probs)
    if not plan:
        st.error("The carton table is empty. Press “Reset table to PO” or add a row.")
        return

    rows, ctn_from = [], 1
    for r in plan:
        ctn_to = ctn_from + r.ctns - 1
        net_g = net_map.get(r.size, tw.DEF_NET)
        net_c = round(net_g * r.pcs_per_ctn, 2)
        gr_c = round(net_c + tare, 2)
        rows.append({"From": ctn_from, "To": ctn_to, "Ctns": r.ctns, "Meas (cm)": f"{r.L:g}×{r.W:g}×{r.H:g}",
                     "CBM": round(r.L * r.W * r.H * r.ctns / 1e6, 3), "Size": r.size, "Pcs/Ctn": r.pcs_per_ctn,
                     "Total Pcs": r.total_pcs, "Nt/Ctn": net_c, "Gr/Ctn": gr_c,
                     "Total Nt": round(net_c * r.ctns, 2), "Total Gr": round(gr_c * r.ctns, 2)})
        ctn_from = ctn_to + 1
    st.dataframe(pd.DataFrame(rows), hide_index=True, **kit.STRETCH)

    # ------------------------------------------------------------ 5
    kit.step(5, "Check against the PO")
    order = {ln.size: ln.qty for ln in po.lines}
    packed = {s: sum(r.total_pcs for r in plan if r.size == s) for s in sizes}
    kit.check_by_size(sizes, order, packed)
    tot_pcs = sum(r.total_pcs for r in plan)
    tot_ctn = sum(r.ctns for r in plan)
    tot_cbm = round(sum(r.L * r.W * r.H * r.ctns / 1e6 for r in plan), 3)
    tot_net = round(sum(net_map.get(r.size, tw.DEF_NET) * r.total_pcs for r in plan), 2)
    tot_gr = round(sum((net_map.get(r.size, tw.DEF_NET) * r.pcs_per_ctn + tare) * r.ctns for r in plan), 2)
    kit.summary_cards([("Cartons", f"{tot_ctn:,}"), ("Pieces", f"{tot_pcs:,}"), ("CBM", f"{tot_cbm}"),
                       ("Net kg", f"{tot_net:,.0f}"), ("Gross kg", f"{tot_gr:,.0f}")])

    # ------------------------------------------------------------ 6
    kit.step(6, "Download")
    xlsx = tw.make_packing_excel(po, plan, net_map, tare, company=company, address=address)
    fname = f"PACKING_LIST_{po.style}_{po.order_no}.xlsx"
    kit.download_block("⬇ Download packing list (Excel)", xlsx, fname,
                       f"{tot_ctn:,} cartons · {tot_pcs:,} pcs · CBM {tot_cbm}")
