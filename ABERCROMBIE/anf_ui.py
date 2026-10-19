"""
A&F (Abercrombie & Fitch) buyer page for the unified router.

    from anf_ui import render as render_anf
    ...
    elif buyer == "A&F":
        render_anf()

Shared design (pl_kit): 1 PO summary -> 2 Quantities to pack -> 3 Variables -> 4 Carton table (editable)
-> 5 Check vs PO -> 6 Download Excel / PDF.
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
import anf_packing as ap                       # noqa: E402

_V = tuple(int(x) for x in st.__version__.split(".")[:2])
_STRETCH = {"width": "stretch"} if _V >= (1, 50) else {"use_container_width": True}


@st.cache_data(show_spinner="Reading PO…")
def _parse(data: bytes):
    return ap.parse_po(data)


def render(default_file=None) -> None:
    st.header("A&F — PO (PDF) → Packing list")
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
    kit.problems(po.warnings)

    k = f"anf_{po.po}"
    sizes = ap.ordered_sizes(po.lines)

    # ------------------------------------------------------------ 1
    kit.step(1, "PO summary")
    kit.summary_cards([("PO", po.po), ("Tech pack", po.tech_pack), ("Pack", po.pack_type or "-"),
                       ("Ship date", po.ship_date), ("Pieces", f"{po.total_units:,}")])
    st.caption(f"code {po.po_code} · carton standard {po.carton_std}")
    with st.expander("PO lines (SKU, EAN, cost)"):
        st.dataframe(po.lines[["color", "size", "qty", "sku", "short_sku", "cost"]].rename(columns=str.title),
                     hide_index=True, height=min(320, 40 + 35 * len(po.lines)), **kit.STRETCH)

    # ------------------------------------------------------------ 2
    kit.step(2, "Quantities to pack", "Change “Ship qty” to pack more or less than ordered.")
    sdf = pd.DataFrame({"Color": po.lines["color"], "Size": po.lines["size"],
                        "Order qty": po.lines["qty"].astype(int), "Ship qty": po.lines["qty"].astype(int)})
    sdf = sdf.sort_values(by=["Color", "Size"], key=lambda s: s.map(ap.size_key) if s.name == "Size" else s,
                          kind="stable").reset_index(drop=True)
    sed = st.data_editor(sdf, hide_index=True, key=f"{k}_s", disabled=["Color", "Size", "Order qty"], **kit.STRETCH,
                         column_config={"Ship qty": st.column_config.NumberColumn(min_value=0, step=1, format="%d")})
    ship = {(r["Color"], r["Size"]): (int(r["Ship qty"]) if not pd.isna(r["Ship qty"]) else 0)
            for _, r in sed.iterrows()}

    # ------------------------------------------------------------ 3
    kit.step(3, "Variables")
    c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
    style = c1.text_input("Style (packing list)", value=po.tech_pack, key=f"{k}_style")
    desc = c2.text_input("Description", value=po.po_code, key=f"{k}_desc")
    std = int(c3.number_input("Pcs per full carton", min_value=1, value=int(po.carton_std or ap.DEF_STD),
                              step=1, key=f"{k}_std"))
    tare = float(c4.number_input("Carton weight (kg)", min_value=0.0, value=float(ap.DEF_TARE), step=0.1, key=f"{k}_tare"))
    w1, w2 = st.columns([3, 2])
    with w1:
        wed = st.data_editor(
            pd.DataFrame({"Size": sizes, "Net wt per pc (kg)": [ap.DEF_NET_PER_PC] * len(sizes)}),
            hide_index=True, key=f"{k}_w", disabled=["Size"], **kit.STRETCH,
            column_config={"Net wt per pc (kg)": st.column_config.NumberColumn(min_value=0.0, step=0.01, format="%.3f")})
    with w2:
        st.caption("Big carton (full)  L × W × H, cm")
        b1, b2, b3 = st.columns(3)
        big = (b1.number_input("L", value=ap.DEF_BIG[0], key=f"{k}_bl"),
               b2.number_input("W", value=ap.DEF_BIG[1], key=f"{k}_bw"),
               b3.number_input("H", value=ap.DEF_BIG[2], key=f"{k}_bh"))
        st.caption("Small carton (remainder)  L × W × H, cm")
        s1, s2, s3 = st.columns(3)
        small = (s1.number_input("L ", value=ap.DEF_SMALL[0], key=f"{k}_sl"),
                 s2.number_input("W ", value=ap.DEF_SMALL[1], key=f"{k}_sw"),
                 s3.number_input("H ", value=ap.DEF_SMALL[2], key=f"{k}_sh"))
    wt = {r["Size"]: float(r["Net wt per pc (kg)"] or 0) for _, r in wed.iterrows()}

    auto = ap.make_plan(po, ship, wt, tare, std, big, small, style=style, description=desc)

    # ------------------------------------------------------------ 4
    kit.step(4, "Carton table",
             "Made from the quantities above. Edit any cell, add a row with +, or tick 🗑 Delete and press "
             "“Remove ticked rows”. Changing the quantities or variables above rebuilds this table.")
    sig = hashlib.md5(repr((sorted(ship.items()), std, tare, sorted(wt.items()), big, small)).encode()).hexdigest()[:10]
    base = pd.DataFrame([{"Color": r["color"], "Size": r["size"], "Qty/Ctn": r["qty"], "Ctns": r["ctns"],
                          "Carton": "Small" if r["small"] else "Big"} for r in auto.rows])
    cfg = {
        "Color": st.column_config.SelectboxColumn("Color", options=list(auto.colors), required=True),
        "Size": st.column_config.SelectboxColumn("Size", options=list(auto.sizes), required=True),
        "Qty/Ctn": st.column_config.NumberColumn("Qty/Ctn", min_value=1, step=1, format="%d"),
        "Ctns": st.column_config.NumberColumn("Ctns", min_value=1, step=1, format="%d"),
        "Carton": st.column_config.SelectboxColumn("Carton", options=["Big", "Small"]),
    }
    ed = kit.editable_grid(base, f"{k}_{sig}", column_config=cfg)
    rows_in = [{"color": r["Color"], "size": r["Size"], "qty": r["Qty/Ctn"], "ctns": r["Ctns"],
                "small": r["Carton"] == "Small"} for r in ed.to_dict("records")]
    plan, probs = ap.plan_from_rows(auto, rows_in)
    kit.problems(probs)
    if not plan.rows:
        st.error("The carton table is empty. Press “Reset table to PO” or add a row.")
        return
    st.dataframe(ap.main_df(plan), hide_index=True, height=min(620, 40 + 35 * (len(plan.rows) + 1)), **kit.STRETCH)

    # ------------------------------------------------------------ 5
    kit.step(5, "Check against the PO")
    multi = len(plan.colors) > 1
    keys = [(c, s) for c in plan.colors for s in plan.sizes if (c, s) in plan.order]
    lab = (lambda c, s: f"{c} / {s}") if multi else (lambda c, s: s)
    kit.check_by_size([lab(c, s) for c, s in keys],
                      {lab(c, s): plan.order[(c, s)] for c, s in keys},
                      {lab(c, s): plan.ship.get((c, s), 0) for c, s in keys})
    S = plan.summary
    kit.summary_cards([("Cartons", f"{S['ctns']:,}"), ("Big / small", f"{S['big']:,} / {S['small']:,}"),
                       ("Pieces", f"{S['ship']:,}"), ("CBM", f"{S['cbm']:,.3f}"),
                       ("Gross kg", f"{S['gross']:,.0f}")])
    with st.expander("Summary and colour / size breakdown"):
        st.dataframe(ap.summary_df(plan), hide_index=True, **kit.STRETCH)
        for color in plan.colors:
            st.caption(f"Colour & size-wise breakdown · {color}")
            st.dataframe(ap.breakdown_df(plan, color), **kit.STRETCH)

    # ------------------------------------------------------------ 6
    kit.step(6, "Download")
    stem = ap.file_stem(plan)
    d1, d2, _ = st.columns([1, 1, 3])
    d1.download_button("⬇ Download Excel", data=ap.build_xlsx(plan), file_name=f"{stem}.xlsx", type="primary",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=f"{k}_dlx", **kit.STRETCH)
    d2.download_button("⬇ Download PDF", data=ap.build_pdf(plan), file_name=f"{stem}.pdf",
                       mime="application/pdf", key=f"{k}_dlp", **kit.STRETCH)
