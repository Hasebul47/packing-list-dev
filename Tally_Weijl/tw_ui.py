"""
Tally Weijl buyer page (standalone or for the unified router).

    from tw_ui import render
    render()

Page order:
  1. PO lines
  2. Main packing table (preview)
  3. Summary / abstract
  4. Calculations (std, weights, dims)
  5. Download Excel
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

import tw_packing as tw

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
        up = st.file_uploader(
            "Tally Weijl Order Form (.pdf)",
            type=["pdf"],
            key="tw_upload",
        )
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

    for w in po.warnings:
        st.warning(w)
    if not po.lines:
        st.error("No size lines found. Confirm this is a Tally Weijl Order Form PDF.")
        return

    st.caption(
        f"PO **{po.order_no}** · style **{po.style}** · {po.color_code} {po.color_name} · "
        f"{po.total_qty:,} pcs · {po.destination} · {po.incoterm} · ship {po.ship_date or '—'}"
    )

    # ------------------------------------------------------------------ 1 PO lines
    st.subheader("1 · PO lines")
    lines_df = pd.DataFrame(
        [
            {
                "Size": ln.size,
                "Color code": ln.color_code,
                "Color": ln.color_name,
                "EAN13": ln.ean,
                "Qty": ln.qty,
                "USD": ln.unit_usd,
            }
            for ln in po.lines
        ]
    )
    st.dataframe(
        lines_df,
        hide_index=True,
        height=min(280, 40 + 35 * len(lines_df)),
        **_STRETCH,
    )

    # containers filled after inputs
    st.subheader("2 · Main table")
    main_box = st.container()
    st.subheader("3 · Summary")
    sum_box = st.container()
    st.subheader("4 · Calculations")
    calc_box = st.container()
    st.subheader("5 · Download")
    dl_box = st.container()

    sizes = po.sizes
    k = po.order_no

    with calc_box:
        c1, c2, c3 = st.columns(3)
        std = int(
            c1.number_input(
                "Pcs per full carton",
                min_value=1,
                value=int(tw.DEF_STD),
                step=1,
                key=f"tw_std_{k}",
            )
        )
        tare = float(
            c2.number_input(
                "Carton tare (kg)",
                min_value=0.0,
                value=float(tw.DEF_TARE),
                step=0.1,
                key=f"tw_tare_{k}",
            )
        )
        c3.markdown(f"**Item**  \n{po.item_desc or '—'}")

        st.markdown("**Net weight per garment (kg)**")
        wdf = pd.DataFrame(
            {"Size": sizes, "Net wt / pc (kg)": [tw.DEF_NET] * len(sizes)}
        )
        wed = st.data_editor(
            wdf,
            hide_index=True,
            key=f"tw_w_{k}",
            disabled=["Size"],
            column_config={
                "Net wt / pc (kg)": st.column_config.NumberColumn(
                    min_value=0.0, step=0.01, format="%.3f"
                )
            },
            **_STRETCH,
        )
        net_map = {
            str(r["Size"]): float(r["Net wt / pc (kg)"]) for _, r in wed.iterrows()
        }

        st.markdown("**Carton size (cm)**")
        d1, d2, d3 = st.columns(3)
        Lb = d1.number_input("Full L", min_value=1.0, value=tw.DEF_BIG[0], key=f"tw_Lb_{k}")
        Wb = d2.number_input("Full W", min_value=1.0, value=tw.DEF_BIG[1], key=f"tw_Wb_{k}")
        Hb = d3.number_input("Full H", min_value=1.0, value=tw.DEF_BIG[2], key=f"tw_Hb_{k}")
        s1, s2, s3 = st.columns(3)
        Ls = s1.number_input("Small L", min_value=1.0, value=tw.DEF_BIG[0], key=f"tw_Ls_{k}")
        Ws = s2.number_input("Small W", min_value=1.0, value=tw.DEF_BIG[1], key=f"tw_Ws_{k}")
        Hs = s3.number_input("Small H", min_value=1.0, value=20.0, key=f"tw_Hs_{k}")

        company = st.text_input(
            "Company header",
            value="THAT'S IT SPORT'S WEAR LTD.     ( UNIT = 04 )",
            key=f"tw_co_{k}",
        )
        address = st.text_input(
            "Address",
            value="Narashighpur, Ashulia, Dhaka-1341.",
            key=f"tw_addr_{k}",
        )

    plan = tw.build_carton_plan(
        po, std, (Lb, Wb, Hb), (Ls, Ws, Hs)
    )

    # ------------------------------------------------------------------ 2 main table
    with main_box:
        if not plan:
            st.warning("No cartons to show.")
        else:
            rows = []
            ctn_from = 1
            for r in plan:
                ctn_to = ctn_from + r.ctns - 1
                net_g = net_map.get(r.size, tw.DEF_NET)
                net_c = round(net_g * r.pcs_per_ctn, 2)
                gr_c = round(net_c + tare, 2)
                rows.append(
                    {
                        "From": ctn_from,
                        "To": ctn_to,
                        "Ctns": r.ctns,
                        "Meas (cm)": f"{r.L:g}×{r.W:g}×{r.H:g}",
                        "CBM": round(r.L * r.W * r.H * r.ctns / 1e6, 3),
                        "Size": r.size,
                        "Pcs/Ctn": r.pcs_per_ctn,
                        "Total Pcs": r.total_pcs,
                        "Nt/pc": net_g,
                        "Nt/Ctn": net_c,
                        "Gr/Ctn": gr_c,
                        "Total Nt": round(net_c * r.ctns, 2),
                        "Total Gr": round(gr_c * r.ctns, 2),
                    }
                )
                ctn_from = ctn_to + 1
            st.dataframe(pd.DataFrame(rows), hide_index=True, **_STRETCH)

    # ------------------------------------------------------------------ 3 summary
    with sum_box:
        order_qty = po.total_qty
        packed = sum(r.total_pcs for r in plan)
        excess = packed - order_qty
        total_ctns = sum(r.ctns for r in plan)
        total_cbm = round(
            sum(r.L * r.W * r.H * r.ctns / 1e6 for r in plan), 3
        )
        total_net = round(
            sum(net_map.get(r.size, tw.DEF_NET) * r.total_pcs for r in plan), 2
        )
        total_gr = round(
            sum(
                (net_map.get(r.size, tw.DEF_NET) * r.pcs_per_ctn + tare) * r.ctns
                for r in plan
            ),
            2,
        )

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Order qty", f"{order_qty:,}")
        m2.metric("Packed qty", f"{packed:,}", delta=f"{excess:+d}")
        m3.metric("Cartons", f"{total_ctns}")
        m4.metric("CBM", f"{total_cbm}")
        m5.metric("Net / Gross kg", f"{total_net} / {total_gr}")

        by_size = []
        for ln in po.lines:
            pk = sum(r.total_pcs for r in plan if r.size == ln.size)
            by_size.append(
                {
                    "Size": ln.size,
                    "Order": ln.qty,
                    "Packed": pk,
                    "Excess": pk - ln.qty,
                    "%": round((pk - ln.qty) / ln.qty * 100, 2) if ln.qty else 0,
                }
            )
        st.dataframe(pd.DataFrame(by_size), hide_index=True, **_STRETCH)

    # ------------------------------------------------------------------ 5 download
    with dl_box:
        xlsx = tw.make_packing_excel(
            po, plan, net_map, tare, company=company, address=address
        )
        fname = f"PACKING_LIST_{po.style}_{po.order_no}.xlsx"
        st.download_button(
            "⬇ Download packing list (Excel)",
            data=xlsx,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(
            f"{total_ctns} cartons · {packed:,} pcs · CBM {total_cbm} · "
            f"file `{fname}`"
        )
