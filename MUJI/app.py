"""MUJI PO -> Packing List  (Streamlit).  Run:  streamlit run app.py
Needs muji_packing_list.py in the same folder."""
import io
import openpyxl
import pandas as pd
import streamlit as st
import sys
from pathlib import Path

import muji_packing_list as m

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:                      # pl_kit.py lives in the project root
    sys.path.insert(0, _ROOT)
import pl_kit as kit                           # noqa: E402


def run(xlsx_file=None):
    st.title("📦 MUJI PO → Packing List")
    st.caption("Creative Collections Ltd. · upload the MUJI PO sheet (発注一覧) and get packing lists in the reference format.")

    if xlsx_file is not None:
        po = xlsx_file
    else:
        po = st.file_uploader("PO sheet (.xlsx)", type=["xlsx"], key="muji_upload")

    if not po:
        st.info("Upload the PO Excel to begin.")
        return

    try:
        rows = m.read_po(po)
        groups, issues = m.build_groups(rows)
    except Exception as e:
        st.error(f"Could not read this file as a MUJI PO sheet: {e}")
        return
    if not groups:
        st.error("No PO rows found (expected the header on row 2 and data from row 3).")
        return

    # ---- overview of what the PO contains
    def pack_stats(g):
        assort_ctn = sum(next(iter(v)) for v in g["acount"].values())
        solid_pcs = sum(q for _, q in g["solid"].values())
        solid_ctn = sum(-(-q // m.CARTON_PCS) for _, q in g["solid"].values())
        return assort_ctn, solid_ctn, sum(g["order"].values())

    overview = pd.DataFrame(
        [dict(Destination=d, Style=s, Colour=f"{c}:{m.COLOR_EN.get(c, '')}",
              **dict(zip(["Assort ctns", "Solid ctns", "PO qty (pcs)"], pack_stats(g))))
         for (d, s, c), g in groups.items()])
    overview["Total ctns"] = overview["Assort ctns"] + overview["Solid ctns"]

    kit.step(1, "PO summary")
    kit.summary_cards([("Styles", f"{overview['Style'].nunique()}"),
                       ("Destinations", f"{overview['Destination'].nunique()}"),
                       ("Cartons", f"{int(overview['Total ctns'].sum()):,}"),
                       ("PO pieces", f"{int(overview['PO qty (pcs)'].sum()):,}")])
    kit.step(2, "Select what to pack", "Choose the styles and destinations to include in the Excel file.")
    if issues:
        st.warning(f"{len(issues)} data check(s) failed - review before using the packing list.")
        st.write("\n".join(f"- {i}" for i in issues))
    else:
        st.success("PO is consistent: solid + assort = PO total for every size, every assort mix = 24 pcs.")

    c1, c2 = st.columns(2)
    styles = sorted(overview["Style"].unique(), key=lambda s: (s != "BE43WA7S", s))
    dests = sorted(overview["Destination"].unique())
    sel_s = c1.multiselect("Style", styles, default=styles, key="muji_sel_s")
    sel_d = c2.multiselect("Destination", dests, default=dests, key="muji_sel_d")
    view = overview[overview["Style"].isin(sel_s) & overview["Destination"].isin(sel_d)]
    st.dataframe(view, hide_index=True, **kit.STRETCH)

    if any(s != "BE43WA7S" for s in sel_s):
        st.caption("⚠ Weights for BE43AA7S (culottes) have no reference - those cells are left yellow/blank. "
                   "Skirt XS/XXL and non-reference assort mixes are estimated (yellow).")

    if view.empty:
        return

    kit.step(3, "Selection",
             f"{len(view)} style / colour / destination line(s) · {int(view['Total ctns'].sum()):,} cartons · "
             f"{int(view['PO qty (pcs)'].sum()):,} pcs")
    kit.step(4, "Download", "Packing lists follow the MUJI reference format (assort mixes and solid cartons).")
    if st.button("Generate packing lists", type="primary", key="muji_gen_btn"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        keys = [k for k in groups if k[0] in sel_d and k[1] in sel_s]
        keys.sort(key=lambda k: (k[1] != "BE43WA7S", k[1], k[0] not in ("Vietnam", "Malaysia"), k[0]))
        for d, s, c in keys:
            m.make_sheet(wb, d, s, c, groups[(d, s, c)], f"{d} {s}")
        m.make_notes(wb)
        wb.calculation.fullCalcOnLoad = True
        buf = io.BytesIO()
        wb.save(buf)
        st.session_state["muji_xlsx"] = buf.getvalue()
        st.session_state["muji_n"] = len(keys)

    if "muji_xlsx" in st.session_state:
        st.success(f"{st.session_state['muji_n']} packing list sheet(s) ready (formulas calculate when opened in Excel).")
        kit.download_block("⬇ Download packing list (Excel)", st.session_state["muji_xlsx"],
                           "Muji_Packing_Lists.xlsx")


if __name__ == "__main__":
    st.set_page_config(page_title="MUJI Packing List", page_icon="📦", layout="wide")
    run()
