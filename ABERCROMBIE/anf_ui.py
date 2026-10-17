"""
A&F (Abercrombie & Fitch) buyer page for the unified router.

    from anf_ui import render as render_anf
    ...
    elif buyer == "A&F":
        render_anf()

Page order: 1 PO lines  ->  2 Main table  ->  3 Summary  ->  4 Calculations  ->  5 Download Excel / PDF.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

import anf_packing as ap

_V = tuple(int(x) for x in st.__version__.split(".")[:2])
_STRETCH = {"width": "stretch"} if _V >= (1, 50) else {"use_container_width": True}


@st.cache_data(show_spinner="Reading PO…")
def _parse(data: bytes):
    return ap.parse_po(data)


def render(default_file=None) -> None:
    st.header("A&F - PO (PDF) → Packing list")
    data = None
    if default_file is not None:
        if hasattr(default_file, "getvalue"):
            data = default_file.getvalue()
        elif isinstance(default_file, (bytes, bytearray)):
            data = bytes(default_file)
        elif isinstance(default_file, (str, Path)) and Path(default_file).exists():
            data = Path(default_file).read_bytes()

    if data is None:
        up = st.file_uploader("A&F purchase order (.pdf)", type=["pdf"], key="anf_upload")
        if up is not None:
            data = up.getvalue()

    if data is None:
        st.info("Upload the A&F PO PDF to start.")
        return
    try:
        po = _parse(data)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the PO: {e}")
        return
    for w in po.warnings:
        st.warning(w)

    st.caption(f"PO **{po.po}** · code {po.po_code} · tech pack {po.tech_pack} · pack {po.pack_type} · "
               f"carton std {po.carton_std} · ship date {po.ship_date} · {po.total_units:,} pcs")
    st.subheader("1 · PO lines")
    st.dataframe(po.lines[["color", "size", "qty", "sku", "short_sku", "cost"]].rename(columns=str.title),
                 hide_index=True, height=min(320, 40 + 35 * len(po.lines)), **_STRETCH)

    st.subheader("2 · Main table")
    main_box = st.container()
    st.subheader("3 · Summary")
    sum_box = st.container()
    st.subheader("4 · Calculations")
    calc_box = st.container()
    st.subheader("5 · Download")
    dl_box = st.container()

    sizes = ap.ordered_sizes(po.lines)
    k = po.po
    with calc_box:
        c1, c2, c3 = st.columns(3)
        style = c1.text_input("Style (packing list)", value=po.tech_pack, key=f"anf_style_{k}")
        desc = c2.text_input("Description", value=po.po_code, key=f"anf_desc_{k}")
        std = int(c3.number_input("Pcs per full carton", min_value=1, value=int(po.carton_std or ap.DEF_STD),
                                  step=1, key=f"anf_std_{k}"))

        st.markdown("**Table 2 · Weights and carton size**")
        w1, w2 = st.columns([3, 2])
        with w1:
            wed = st.data_editor(
                pd.DataFrame({"Size": sizes, "Net wt per pc (kg)": [ap.DEF_NET_PER_PC] * len(sizes)}),
                hide_index=True, key=f"anf_w_{k}", disabled=["Size"], **_STRETCH,
                column_config={"Net wt per pc (kg)": st.column_config.NumberColumn(min_value=0.0, step=0.01, format="%.3f")})
        with w2:
            tare = st.number_input("Carton weight (kg)", min_value=0.0, value=ap.DEF_TARE, step=0.1, key=f"anf_tare_{k}")
            st.caption("Big carton (full) L × W × H, cm")
            b1, b2, b3 = st.columns(3)
            big = (b1.number_input("L", value=ap.DEF_BIG[0], key=f"anf_bl_{k}"),
                   b2.number_input("W", value=ap.DEF_BIG[1], key=f"anf_bw_{k}"),
                   b3.number_input("H", value=ap.DEF_BIG[2], key=f"anf_bh_{k}"))
            st.caption("Small carton (remainder) L × W × H, cm")
            s1, s2, s3 = st.columns(3)
            small = (s1.number_input("L ", value=ap.DEF_SMALL[0], key=f"anf_sl_{k}"),
                     s2.number_input("W ", value=ap.DEF_SMALL[1], key=f"anf_sw_{k}"),
                     s3.number_input("H ", value=ap.DEF_SMALL[2], key=f"anf_sh_{k}"))
        wt = {r["Size"]: float(r["Net wt per pc (kg)"] or 0) for _, r in wed.iterrows()}

        st.markdown("**Table 3 · Order and ship quantity** (edit *Ship qty* to pack more or less than ordered)")
        sdf = pd.DataFrame({"Color": po.lines["color"], "Size": po.lines["size"],
                            "Order qty": po.lines["qty"].astype(int), "Ship qty": po.lines["qty"].astype(int)})
        sdf = sdf.sort_values(by=["Color", "Size"], key=lambda s: s.map(ap.size_key) if s.name == "Size" else s,
                              kind="stable").reset_index(drop=True)
        sed = st.data_editor(sdf, hide_index=True, key=f"anf_s_{k}", disabled=["Color", "Size", "Order qty"], **_STRETCH,
                             column_config={"Ship qty": st.column_config.NumberColumn(min_value=0, step=1, format="%d")})
        ship = {(r["Color"], r["Size"]): (int(r["Ship qty"]) if not pd.isna(r["Ship qty"]) else 0)
                for _, r in sed.iterrows()}

    plan = ap.make_plan(po, ship, wt, float(tare), std, big, small, style=style, description=desc)
    with main_box:
        st.dataframe(ap.main_df(plan), hide_index=True, height=min(620, 40 + 35 * (len(plan.rows) + 1)), **_STRETCH)
    with sum_box:
        st.dataframe(ap.summary_df(plan), hide_index=True, **_STRETCH)
        for color in plan.colors:
            st.caption(f"Colour & size-wise breakdown · {color}")
            st.dataframe(ap.breakdown_df(plan, color), **_STRETCH)
    with dl_box:
        stem = ap.file_stem(plan)
        d1, d2, _ = st.columns([1, 1, 3])
        d1.download_button("⬇ Download Excel", data=ap.build_xlsx(plan), file_name=f"{stem}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key=f"anf_dlx_{k}", **_STRETCH)
        d2.download_button("⬇ Download PDF", data=ap.build_pdf(plan), file_name=f"{stem}.pdf",
                           mime="application/pdf", key=f"anf_dlp_{k}", **_STRETCH)
