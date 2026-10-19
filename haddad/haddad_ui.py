"""
Haddad buyer page for the unified router.

    from haddad_ui import render as render_haddad
    ...
    elif buyer == "Haddad":
        render_haddad()

Page order (top to bottom):
  1. Ref list (Table 1)            - click a row to select the ref
  2. Main table (Table 4)          - packing-list preview
  3. Summary (Table 5)
  4. Calculations                  - weights / carton size (Table 2), order qty + excess % (Table 3)
  5. Download Excel / Download PDF

The calculation widgets sit BELOW the outputs, so the outputs are placed in containers
that are created first and filled after the inputs have been read.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st

import hashlib
import sys

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:                      # pl_kit.py lives in the project root
    sys.path.insert(0, _ROOT)

import pl_kit as kit                           # noqa: E402
import haddad_packing as hp                    # noqa: E402

# Streamlit >= 1.50 wants width="stretch"; older versions only know use_container_width.
_V = tuple(int(x) for x in st.__version__.split(".")[:2])
_STRETCH = {"width": "stretch"} if _V >= (1, 50) else {"use_container_width": True}


@st.cache_data(show_spinner="Reading status sheet…")
def _load(data: bytes, name: str) -> pd.DataFrame:
    return hp.load_status(data, name)


def render(default_file=None) -> None:
    st.header("Haddad - PO status sheet → Packing list")

    data, name = None, None
    if default_file is not None:
        if hasattr(default_file, "getvalue"):
            data = default_file.getvalue()
            name = getattr(default_file, "name", "status.xls")
        elif isinstance(default_file, (bytes, bytearray)):
            data = bytes(default_file)
            name = "status.xls"
        elif isinstance(default_file, (str, Path)) and Path(default_file).exists():
            data, name = Path(default_file).read_bytes(), Path(default_file).name

    if data is None:
        up = st.file_uploader("Haddad PO status sheet (.xls / .xlsx)", type=["xls", "xlsx"], key="hd_upload")
        if up is not None:
            data, name = up.getvalue(), up.name

    if data is None:
        st.info("Upload the Haddad PO status sheet to start.")
        return

    try:
        df = _load(data, name)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the file: {e}")
        return
    if df.attrs.get("dropped"):
        st.warning(f"{df.attrs['dropped']} line(s) skipped (missing PO qty or box qty).")

    # ------------------------------------------------------------------ Table 1: refs
    kit.step(1, "Select ref")
    reft = hp.ref_table(df)
    refs = reft["Ref#"].tolist()
    if st.session_state.get("hd_ref") not in refs:
        st.session_state["hd_ref"] = refs[0]
    try:
        ev = st.dataframe(reft, hide_index=True, **_STRETCH, height=min(380, 40 + 35 * len(reft)),
                          on_select="rerun", selection_mode="single-row", key="hd_reftbl")
        if ev.selection.rows:
            st.session_state["hd_ref"] = refs[ev.selection.rows[0]]
    except TypeError:          # older Streamlit without row selection
        st.session_state["hd_ref"] = st.selectbox("Ref#", refs, index=refs.index(st.session_state["hd_ref"]))
    ref = st.session_state["hd_ref"]

    sub = df[df["Ref#"] == ref]
    c1, c2, c3 = st.columns([2, 3, 2])
    with c1:
        st.markdown(f"**Selected: {ref}**")
    with c2:
        dates = sorted(sub["Ship Date"].unique(), key=lambda d: pd.to_datetime(d, format="%m/%d/%y", errors="coerce"))
        ship_sel = st.multiselect("Ship date", dates, default=dates, key=f"hd_ship_{ref}")
    with c3:
        has_bk = (sub["Status"].str.upper() == "PO BK").any()
        incl_bk = st.checkbox("Include PO BK lines", value=True, key=f"hd_bk_{ref}", disabled=not has_bk)

    lines = hp.select_lines(df, ref, ship_sel, incl_bk)
    if lines.empty:
        st.warning("No lines for this selection.")
        return
    sizes = hp.size_union(lines)

    kit.summary_cards([("Ref", ref), ("POs", f"{lines['PO#'].nunique()}"),
                       ("Pieces", f"{int(lines['PO Qty'].sum()):,}"), ("Sizes", f"{len(sizes)}")])

    # ------------------------------------------------------------ 2 variables
    kit.step(2, "Variables")
    with st.expander("Header details (optional)", expanded=False):
        h1, h2, h3, h4 = st.columns(4)
        date_v = h1.date_input("Date", value=dt.date.today(), key=f"hd_date_{ref}")
        invoice = h2.text_input("Invoice no", key=f"hd_inv_{ref}")
        container = h3.text_input("Container no", key=f"hd_con_{ref}")
        warehouse = h4.text_input("Warehouse code", key=f"hd_wh_{ref}")
    restart = st.radio("Carton numbering", ["Running (1 … N across POs)", "Restart at 1 for every PO"],
                       horizontal=True, key=f"hd_num_{ref}") != "Running (1 … N across POs)"
    w1, w2 = st.columns([3, 2])
    with w1:
        wdf = pd.DataFrame({"Size": sizes, "Net wt per pc (kg)": [hp.DEF_NET_PER_PC] * len(sizes)})
        wed = st.data_editor(
            wdf, hide_index=True, **kit.STRETCH, key=f"hd_w_{ref}", disabled=["Size"],
            column_config={"Net wt per pc (kg)": st.column_config.NumberColumn(
                min_value=0.0, step=0.001, format="%.5f")})
    with w2:
        tare = st.number_input("Carton tare (kg)", min_value=0.0, value=hp.DEF_TARE, step=0.05, key=f"hd_tare_{ref}")
        d1, d2, d3 = st.columns(3)
        L = d1.number_input("L (in)", min_value=0.0, value=hp.DEF_DIMS[0], step=0.5, key=f"hd_L_{ref}")
        W = d2.number_input("W (in)", min_value=0.0, value=hp.DEF_DIMS[1], step=0.5, key=f"hd_W_{ref}")
        H = d3.number_input("H (in)", min_value=0.0, value=hp.DEF_DIMS[2], step=0.5, key=f"hd_H_{ref}")
    wt = {r["Size"]: float(r["Net wt per pc (kg)"] or 0) for _, r in wed.iterrows()}

    st.markdown("**Order quantity and excess**  (edit *Excess %*, or type an *Order qty* to override)")
    odf = pd.DataFrame({
        "PO#": lines["PO#"], "Color": lines["Fashion Color"] + " - " + lines["Color"],
        "Ship qty": lines["PO Qty"].astype(int), "Excess %": 0.0, "Order qty (override)": None,
    })
    oed = st.data_editor(
        odf, hide_index=True, **kit.STRETCH, key=f"hd_o_{ref}",
        disabled=["PO#", "Color", "Ship qty"],
        column_config={
            "Ship qty": st.column_config.NumberColumn(format="%d"),
            "Excess %": st.column_config.NumberColumn(min_value=-50.0, max_value=100.0, step=0.1, format="%.2f"),
            "Order qty (override)": st.column_config.NumberColumn(min_value=0, step=1, format="%d"),
        })
    excess = {r["PO#"]: (0.0 if pd.isna(r["Excess %"]) else float(r["Excess %"])) for _, r in oed.iterrows()}
    override = {r["PO#"]: r["Order qty (override)"] for _, r in oed.iterrows()
                if not pd.isna(r["Order qty (override)"])}

    try:
        auto = hp.compute(lines, sizes, excess, override, wt, float(tare), (float(L), float(W), float(H)), restart)
    except ValueError as e:
        st.error(str(e))
        return
    kit.problems(auto.warnings)
    header = {"date": date_v, "invoice": invoice, "container": container, "warehouse": warehouse}

    # ------------------------------------------------------------ 3 carton table
    kit.step(3, "Carton table",
             "Made from the sheet and the variables above. Edit any cell (cartons, pieces per size, carton size, "
             "order qty), add a row with +, or tick 🗑 Delete and press “Remove ticked rows”. "
             "Changing the variables above rebuilds this table.")
    sig = hashlib.md5(repr((ref, list(ship_sel), incl_bk, sorted(excess.items()), sorted(override.items()),
                            sorted(wt.items()), tare, (L, W, H), restart)).encode()).hexdigest()[:10]
    recs = []
    for x in auto.rows:
        d = {"PO#": x["po"], "Ctns": x["ctns"]}
        for sz in sizes:
            d[sz] = x["pcs"].get(sz, 0)
        d.update({"L (in)": x["dims"][0], "W (in)": x["dims"][1], "H (in)": x["dims"][2], "Order qty": x["order"]})
        recs.append(d)
    base = pd.DataFrame(recs)
    cfg = {"PO#": st.column_config.SelectboxColumn("PO#", options=list(lines["PO#"].unique()), required=True),
           "Ctns": st.column_config.NumberColumn("Ctns", min_value=1, step=1, format="%d"),
           "Order qty": st.column_config.NumberColumn("Order qty", min_value=0, step=1, format="%d")}
    for sz in sizes:
        cfg[sz] = st.column_config.NumberColumn(sz, min_value=0, step=1, format="%d")
    for c in ("L (in)", "W (in)", "H (in)"):
        cfg[c] = st.column_config.NumberColumn(c, min_value=0.5, step=0.5)
    ed = kit.editable_grid(base, f"hd_{ref}_{sig}", column_config=cfg)
    rows_in = [{"po": r["PO#"], "ctns": r["Ctns"], "pcs": {sz: r.get(sz) for sz in sizes},
                "dims": (r["L (in)"], r["W (in)"], r["H (in)"]), "order": r["Order qty"]}
               for r in ed.to_dict("records")]
    res, probs = hp.compute_rows(lines, sizes, rows_in, wt, float(tare), restart)
    kit.problems(probs)
    if not res.rows:
        st.error("The carton table is empty. Press “Reset table to PO” or add a row.")
        return
    st.dataframe(hp.main_df(res), hide_index=True, **kit.STRETCH, height=min(600, 40 + 35 * (len(res.rows) + 1)))

    # ------------------------------------------------------------ 4 check
    kit.step(4, "Check against the sheet")
    S = res.summary
    pcs_sheet = {}
    for _, r in lines.iterrows():
        for sz, q in hp.pack_pcs(r["Size Configuration"], r["Pack Ratio"], int(r["Master Box Quantity"]), r["PO#"]).items():
            pcs_sheet[sz] = pcs_sheet.get(sz, 0) + q * int(round(r["PO Qty"] / int(r["Master Box Quantity"])))
    packed = {sz: sum(x["pcs"].get(sz, 0) * x["ctns"] for x in res.rows) for sz in sizes}
    kit.check_by_size(sizes, pcs_sheet, packed)
    kit.summary_cards([("Cartons", f"{S['ctns']:,}"), ("Ship qty", f"{S['total_ship']:,}"),
                       ("Excess / short", f"{S['excess_qty']:+,}"), ("CBM", f"{S['cbm']:,.2f}"),
                       ("Gross lbs", f"{S['gross_lbs']:,.0f}")])
    with st.expander("Full summary"):
        st.dataframe(hp.summary_df(res), hide_index=True, **kit.STRETCH)

    # ------------------------------------------------------------ 5 download
    kit.step(5, "Download")
    stem = hp.file_stem(res, lines)
    b1, b2, _ = st.columns([1, 1, 3])
    b1.download_button("⬇ Download Excel", data=hp.build_xlsx(res, header), file_name=f"{stem}.xlsx", type="primary",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       **kit.STRETCH, key=f"hd_dlx_{ref}")
    b2.download_button("⬇ Download PDF", data=hp.build_pdf(res, header), file_name=f"{stem}.pdf",
                       mime="application/pdf", **kit.STRETCH, key=f"hd_dlp_{ref}")
