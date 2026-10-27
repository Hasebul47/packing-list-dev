"""
unified_packing.py
------------------
GAP / Old Navy PO → Packing List (Unified)

Logic  : GAP (Bulk / Single / Multi Y / Multi N) — merge same STYLE by COLOR
UI     : Interactive preview (yellow strips, PO header, live totals, Excel)

Run:
    streamlit run unified_packing.py
"""

from __future__ import annotations

import io
import json
from pathlib import Path

import sys
import hashlib

import pandas as pd
import streamlit as st

_ROOT = str(Path(__file__).resolve().parent.parent)
if _ROOT not in sys.path:                      # pl_kit.py lives in the project root
    sys.path.insert(0, _ROOT)
import pl_kit as kit                           # noqa: E402

import parser_lib as parser
from color_builders import (
    BULK_CTN_CAP,
    DEFAULT_CTN_CAP,
    DEFAULT_CTN_MEAS,
    DEFAULT_EXCESS_PCT,
    DEFAULT_NW,
    build_bulk_by_color,
    build_multi_y_by_color,
    build_single_by_color,
    carton_code,
    carton_weights,
    detect_table_type,
    rebuild_one_size_section,
    section_to_display_rows,
    sort_sizes,
)
from multi_n_core import (
    DEFAULT_CTN_CODE,
    DEFAULT_NET_DEDUCT,
    DEFAULT_NNW,
    DEFAULT_PREPACKS_PER_CTN,
    DEFAULT_TARE,
    build_multi_n,
)
from gap_excel_export import default_box_size, export_multi_n_excel, export_unified_excel

PACK_TYPES = ("Bulk", "Single", "Multi Y", "Multi N")


def detect_all(data: dict) -> dict[str, list[str]]:
    found = {t: [] for t in PACK_TYPES}
    for tname, table in (data.get("tables") or {}).items():
        ptype = detect_table_type(table)
        if ptype:
            found[ptype].append(tname)
    return found


def inject_css():
    st.markdown(
        """
<style>
  .plist-wrap { font-family: 'Courier New', monospace; font-size: 13px; }
  .cc-strip {
    text-align: center; font-weight: bold; font-size: 15px;
    letter-spacing: 2px; background: #FFFF00; border: 1px solid #000;
    padding: 6px 0; margin: 14px 0 8px 0;
  }
  .po-grid {
    display: grid; grid-template-columns: 1.3fr 1fr; gap: 18px;
    margin-bottom: 8px; line-height: 1.45;
  }
  .po-line { display: flex; }
  .po-label { width: 110px; font-weight: bold; }
  .po-colon { width: 14px; text-align: center; }
  table.plist {
    border-collapse: collapse; width: 100%; margin: 6px 0 14px 0;
    font-family: 'Courier New', monospace; font-size: 12.5px;
  }
  table.plist th, table.plist td {
    border: 1px solid #333; padding: 3px 6px; text-align: center;
  }
  table.plist th { background: #4472C4; color: #fff; font-weight: bold; }
  table.plist tr.total td { background: #FFF2CC; font-weight: bold; }
</style>
""",
        unsafe_allow_html=True,
    )


def render_section_strip(title: str):
    st.markdown(f'<div class="cc-strip">{title}</div>', unsafe_allow_html=True)


def render_po_header_html(sec: dict, buyer: str, po_no: str, grand_ctn: int):
    ctn_from = sec["rows"][0]["ctn_from"] if sec.get("rows") else 1
    exs = sec["ship_qty"] - sec["qty_ordered"]
    html = f"""
<div class="plist-wrap">
  <div class="po-grid">
    <div>
      <div class="po-line"><span class="po-label">BUYER</span><span class="po-colon">:</span><span>{buyer}</span></div>
      <div class="po-line"><span class="po-label">STYLE</span><span class="po-colon">:</span><span>{sec.get('style_full','')}</span></div>
      <div class="po-line"><span class="po-label">P.O. #</span><span class="po-colon">:</span><span>{po_no}</span></div>
      <div class="po-line"><span class="po-label">O/QTY</span><span class="po-colon">:</span><span>{sec['qty_ordered']} PCS</span></div>
      <div class="po-line"><span class="po-label">SHIP QTY</span><span class="po-colon">:</span><span>{sec['ship_qty']} PCS</span></div>
      <div class="po-line"><span class="po-label">EX/SHORT</span><span class="po-colon">:</span><span>{exs} PCS</span></div>
      <div class="po-line"><span class="po-label">CTN QTY</span><span class="po-colon">:</span><span>{sec['ctn_total']} CTNS</span></div>
    </div>
    <div style="margin-top:48px">
      <div class="po-line"><span class="po-label">PO #</span><span class="po-colon">:</span><span>{po_no}</span></div>
      <div class="po-line"><span class="po-label">SKU / Item</span><span class="po-colon">:</span><span>{sec.get('style_full','')}</span></div>
      <div class="po-line"><span class="po-label">Unit / Prepack</span><span class="po-colon">:</span><span>{sec.get('units_per_prepack',1)}</span></div>
      <div class="po-line"><span class="po-label">CARTON</span><span class="po-colon">:</span><span>{ctn_from:02d} of {grand_ctn}</span></div>
    </div>
  </div>
</div>
"""
    st.markdown(html, unsafe_allow_html=True)


def render_plist_table(sec: dict):
    recs = section_to_display_rows(sec)
    if not recs:
        st.warning("No rows")
        return
    cols = list(recs[0].keys())
    th = "".join(f"<th>{c}</th>" for c in cols)
    body = ""
    for rec in recs:
        cls = ' class="total"' if str(rec.get("CTN NO", "")).upper() == "TOTAL" else ""
        tds = "".join(f"<td>{rec.get(c, '')}</td>" for c in cols)
        body += f"<tr{cls}>{tds}</tr>"
    st.markdown(
        f'<div class="plist-wrap"><table class="plist"><tr>{th}</tr>{body}</table></div>',
        unsafe_allow_html=True,
    )


# NOTE: Excel generation lives in gap_excel_export.py (export_unified_excel / export_multi_n_excel).




# ===========================================================================
# Streamlit UI
# ===========================================================================
def run(file=None):
    inject_css()
    uploaded = file if file is not None else st.file_uploader("Upload Destination Purchase Order PDF", type=["pdf"])
    with st.expander("Or load an existing JSON"):
        json_file = st.file_uploader("JSON file", type=["json"], key="json_up")
    
    data = None
    source_name = ""
    
    if uploaded is not None:
        source_name = uploaded.name
        sig = f"{uploaded.name}::{uploaded.size}"
        if st.session_state.get("_upload_sig") != sig:
            for k in list(st.session_state.keys()):
                if k.startswith(("xlsx_", "mn_xlsx_")):
                    st.session_state.pop(k, None)
            st.session_state["_upload_sig"] = sig
        with st.spinner("Reading PDF…"):
            try:
                text = parser.extract_text(io.BytesIO(uploaded.getvalue()))
            except Exception as e:
                st.error(f"PDF read failed: {e}")
                st.stop()
        with st.spinner("Parsing PO structure…"):
            try:
                data = parser.parse_text(text)
            except Exception as e:
                st.error(f"Parse failed: {e}")
                st.stop()
        st.success(
            f"Found **{len(data.get('tables', {}))}** table(s) "
            f"from `{uploaded.name}` ({len(text):,} characters)"
        )
    elif json_file is not None:
        source_name = json_file.name
        try:
            data = json.loads(json_file.getvalue().decode("utf-8"))
            st.success(f"Loaded JSON `{json_file.name}`")
        except Exception as e:
            st.error(f"Invalid JSON: {e}")
            st.stop()
    
    if data is None:
        st.info("Upload a GAP / Old Navy Destination PO PDF to get started.")
        st.markdown("""
    ### Logic (GAP)
    | Type | Detection | Grouping |
    |------|-----------|----------|
    | **Bulk** | `PrePack Type == Bulk` | One section **per color** |
    | **Single** | `PrePack Type == Single` | One section **per color** (prepacks merged) |
    | **Multi Y** | `Multi` + `Full Carton == Y` | One section **per color** (all sizes one row) |
    | **Multi N** | `Multi` + `Full Carton == N` | Per prepack (shared cartons) |
    
    Same **STYLE** → tables merged by **COLOR** only.
    """)
        st.stop()
    
    header = data.get("PO_HEADER") or {}
    buyer = header.get("BUYER") or "OLD NAVY"
    style = header.get("STYLE") or ""
    po_no = str(header.get("P.O._#") or header.get("P.O. #") or "")
    oiqty = header.get("OIQTY") or 0
    
    detected = detect_all(data)
    active_types = [t for t in PACK_TYPES if detected[t]]
    
    kit.step(1, "PO summary")
    kit.summary_cards([("Buyer", str(buyer)), ("Style", str(style)), ("PO #", po_no or "-"), ("Order qty", f"{oiqty} pcs")])
    dcols = st.columns(4)
    for i, ptype in enumerate(PACK_TYPES):
        n = len(detected[ptype])
        if n:
            dcols[i].success(f"**{ptype}** — {n} table(s)")
        else:
            dcols[i].info(f"{ptype} — none")
    
    if not active_types:
        st.error("This PDF has no GAP / Old Navy packing tables (PrePack Type Bulk / Single / Multi), "
                 "so there is nothing to build here.")
        st.info("Is it another buyer's PO (for example Tally Weijl)? Choose the buyer in the router sidebar "
                "under **Pick a buyer if auto-detection fails**.")
        with st.expander("Technical details - what the GAP parser read"):
            st.json(data)
        st.stop()
    
    kit.step(2, "Variables", "Packing defaults. Changing them rebuilds the carton tables below.")
    with st.container(border=True):
        v1, v2, v3 = st.columns(3)
        prepacks_per_ctn = v1.number_input("Prepacks / carton (Multi N)", min_value=1, max_value=50, value=DEFAULT_PREPACKS_PER_CTN)
        single_max_pp = v2.number_input(
            "Max prepacks / carton (Single) — 0 = auto by size", min_value=0, max_value=50, value=0)
        excess_pct = v3.number_input(
            "Max excess % per size (Bulk)", min_value=0.0, max_value=10.0,
            value=float(DEFAULT_EXCESS_PCT), step=0.5,
            help="Extra pieces shipped with Bulk = floor(this % of each size's total order).")
        _sizes = sort_sizes(sorted({
            str(r.get("Size Desc") or "")
            for t in (data.get("tables") or {}).values() for r in (t.get("rows") or [])
        } - {""}))
        caps = {}
        with st.expander("Carton capacity (pcs per carton)"):
            ccols = st.columns(min(max(len(_sizes), 1), 8))
            for j, _s in enumerate(_sizes):
                caps[_s] = ccols[j % len(ccols)].number_input(
                    _s, min_value=1, max_value=500,
                    value=int(BULK_CTN_CAP.get(_s, DEFAULT_CTN_CAP)), key=f"cap_{_s}")
        w1, w2 = st.columns(2)
        net_deduct = w1.number_input("Net deduct (Multi N)", value=float(DEFAULT_NET_DEDUCT), step=0.01)
        ctn_code = w2.text_input("Carton code (Multi N)", value=DEFAULT_CTN_CODE)
    
        # ----- Table A: Net weight by SIZE — 3-row table (header / N.W / N.N.W) -----
        st.markdown("**Net weight by size (N.W / N.N.W)** — editable")
        _nw_sizes = _sizes or ["XS", "S", "M", "L", "XL", "XXL", "XXL+"]
        _nw_row = {s: float(DEFAULT_NW.get(s, 0.5)) for s in _nw_sizes}
        _nnw_row = {
            s: float(DEFAULT_NNW.get(s, round(float(DEFAULT_NW.get(s, 0.5)) - 0.02, 3)))
            for s in _nw_sizes
        }
        nw_df = pd.DataFrame(
            [_nw_row, _nnw_row],
            index=["N.W", "N.N.W"],
            columns=_nw_sizes,
        )
        edited_nw = st.data_editor(
            nw_df,
            key="nw_table",
            use_container_width=True,
            hide_index=False,
            num_rows="fixed",          # no add/delete rows
            disabled=False,
        )
        # Read back edited values (plain numbers, no steppers)
        user_nw = {s: float(edited_nw.loc["N.W", s] or 0) for s in _nw_sizes}
        user_nnw = {s: float(edited_nw.loc["N.N.W", s] or 0) for s in _nw_sizes}
    
        # ----- Table B: Box size / CTN CODE / Empty carton weight — ONE common table -----
        st.markdown("**Box size, CTN CODE and empty carton weight** — editable (shown once in the packing list)")
        _tare_codes = list(DEFAULT_TARE.keys()) or ["G81", "G82", "G84", "G85"]
        _box_def = default_box_size(_tare_codes)
        tare_df = pd.DataFrame(
            [
                {c: _box_def.get(c, "") for c in _tare_codes},
                {c: f"{float(DEFAULT_TARE.get(c, 0.71)):.2f}" for c in _tare_codes},
            ],
            index=["BOX SIZE", "EMPTY CTN WT"],
            columns=_tare_codes,
        )
        tare_df.columns.name = "CTN CODE"
        edited_tare = st.data_editor(
            tare_df,
            key="tare_table",
            use_container_width=True,
            hide_index=False,
            num_rows="fixed",
            disabled=False,
        )

        def _num(v, d):
            try:
                return float(str(v).replace(",", "").strip())
            except (TypeError, ValueError):
                return d
        user_tare = {c: _num(edited_tare.loc["EMPTY CTN WT", c], float(DEFAULT_TARE.get(c, 0.71)))
                     for c in _tare_codes}
        user_box = {c: str(edited_tare.loc["BOX SIZE", c] or "").strip() for c in _tare_codes}

    sig = hashlib.md5(repr((prepacks_per_ctn, single_max_pp, excess_pct, sorted(caps.items()),
                            net_deduct, ctn_code,
                            sorted(user_nw.items()), sorted(user_nnw.items()),
                            sorted(user_tare.items()), sorted(user_box.items()))).encode()).hexdigest()[:10]
    
    # ---- automatic plan (before the user's edits) --------------------------------
    auto_bulk = build_bulk_by_color(data, start_ctn=1, caps=caps, excess_pct=excess_pct) if detected["Bulk"] else []
    auto_single = build_single_by_color(data, start_ctn=1, max_pp=int(single_max_pp) or None, caps=caps) if detected["Single"] else []
    
    kit.step(3, "Carton tables",
             "Bulk and Single: the full packing table is editable — change any cell, add a row with +, "
             "tick 🗑 on rows and press “Remove ticked rows”, or “Reset table to PO”. "
             "Multi Y and Multi N follow the PO rules and are shown read-only. "
             "Changing the variables above rebuilds the tables.")
    ORDER_FWD = "Bulk, Single, Multi"
    ORDER_REV = "Multi, Single, Bulk"
    order_pick = st.radio(
        "Carton serial order",
        [ORDER_FWD, ORDER_REV],
        index=0,
        horizontal=True,
        key="gap_serial_order",
        help="Sets the order of the carton serial numbers, the tabs below and the Excel sections. "
             "Multi Y always comes before Multi N.",
    )
    PACK_ORDER = ("Bulk", "Single", "Multi Y", "Multi N") if order_pick == ORDER_FWD \
        else ("Multi Y", "Multi N", "Single", "Bulk")
    totals_box = st.container()
    
    tab_labels, tab_types = [], []
    for ptype in PACK_ORDER:
        if detected[ptype]:
            tab_labels.append(f"{ptype} ({len(detected[ptype])})")
            tab_types.append(ptype)
    tab_labels += ["Raw tables", "JSON"]
    tabs = st.tabs(tab_labels)
    tab_of = {p: tabs[i] for i, p in enumerate(tab_types)}
    
    STRIP = {"Bulk": "BULK PACK", "Single": "SINGLE PACK", "Multi Y": "MULTY PACK (Y)", "Multi N": "MULTY PACK (N)"}
    
    
    def _editable_sections(ptype, autos, next_ctn):
        """Main packing table: fully editable via st.data_editor (direct, no cache tricks)."""
        out = []
    
        def _to_int(x, default=0):
            if x is None:
                return default
            try:
                if isinstance(x, float) and pd.isna(x):
                    return default
            except Exception:
                pass
            s = str(x).strip().lower()
            if s in ("", "none", "nan", "nat"):
                return default
            try:
                return int(float(x))
            except Exception:
                return default
    
        for si, base in enumerate(autos):
            render_section_strip(STRIP[ptype])
            color = base.get("color", "")
            st.markdown(f"**{color}**  ·  {base.get('style_full', '')}")
            sizes = list(base.get("sizes") or [])
    
            # ---- build rows from real data ----
            rows_out = []
            for r in base.get("rows") or []:
                ctn_from = r.get("ctn_from", "")
                ctn_to = r.get("ctn_to", ctn_from)
                if ctn_to not in ("", None) and ctn_to != ctn_from:
                    ctn_no = f"{ctn_from}-{ctn_to}"
                else:
                    ctn_no = str(ctn_from) if ctn_from not in ("", None) else ""
                ctn_qty = _to_int(r.get("ctn_qty"), 0)
                qty_per = _to_int(r.get("qty_per_ctn"), 0)
                total_qty = _to_int(r.get("total_qty"), ctn_qty * qty_per)
                rec = {
                    "🗑": False,
                    "CTN NO": ctn_no,
                    "CTN MES": str(r.get("ctn_mes") or ""),
                    "CTN QTY": ctn_qty,
                    "COLOR": str(r.get("color") or color or ""),
                    "PREPACK STECKER": str(r.get("prepack") or base.get("prepack") or ""),
                }
                for s in sizes:
                    if ptype in ("Bulk", "Single"):
                        rec[s] = _to_int(r.get("size_val"), 0) if s == r.get("size") else None
                    else:
                        sv = (r.get("size_vals") or base.get("units") or {}).get(s)
                        rec[s] = _to_int(sv, 0) if sv not in (None, "", 0) else None
                rec["QTY PER CTN"] = qty_per
                rec["TOTAL QTY"] = total_qty
                gw, nw = r.get("gross"), r.get("net")
                try:
                    rec["GROSS WEIGHT"] = float(gw) if gw not in (None, "") else None
                except Exception:
                    rec["GROSS WEIGHT"] = None
                try:
                    rec["NET WEIGHT"] = float(nw) if nw not in (None, "") else None
                except Exception:
                    rec["NET WEIGHT"] = None
                rows_out.append(rec)
    
            cols = ["🗑", "CTN NO", "CTN MES", "CTN QTY", "COLOR", "PREPACK STECKER"] + sizes + [
                "QTY PER CTN", "TOTAL QTY", "GROSS WEIGHT", "NET WEIGHT"
            ]
            df = pd.DataFrame(rows_out, columns=cols)
    
            # Session state for this grid (simple, always editable)
            state_key = f"gap_edit_{ptype}_{si}_{color}_{sig}"
            data_key = f"{state_key}__data"
            nonce_key = f"{state_key}__nonce"
            if nonce_key not in st.session_state:
                st.session_state[nonce_key] = 0
            if data_key not in st.session_state:
                st.session_state[data_key] = df.copy()
    
            # Editable inputs vs calculated outputs
            # QTY PER CTN = sum of size columns (XS+S+M+…)
            # TOTAL QTY   = CTN QTY × QTY PER CTN
            # GROSS/NET   = from size NW + empty carton weight
            CALC_COLS = ["QTY PER CTN", "TOTAL QTY", "GROSS WEIGHT", "NET WEIGHT"]
            cfg = {
                "🗑": st.column_config.CheckboxColumn("🗑", default=False, width="small",
                                                      help="Tick to delete row"),
                "CTN NO": st.column_config.TextColumn("CTN NO"),
                "CTN MES": st.column_config.SelectboxColumn("CTN MES", options=list(user_tare.keys()), required=False,
                                                             help="Pick the carton code — carton weight (Gross/Net) follows it"),
                "CTN QTY": st.column_config.NumberColumn("CTN QTY", min_value=0, step=1, format="%d"),
                "COLOR": st.column_config.TextColumn("COLOR"),
                "PREPACK STECKER": st.column_config.TextColumn("PREPACK STECKER"),
                "QTY PER CTN": st.column_config.NumberColumn(
                    "QTY PER CTN", format="%d",
                    help="Auto: sum of size columns (XS+S+M+…)"),
                "TOTAL QTY": st.column_config.NumberColumn(
                    "TOTAL QTY", format="%d",
                    help="Auto: CTN QTY × QTY PER CTN"),
                "GROSS WEIGHT": st.column_config.NumberColumn(
                    "GROSS WEIGHT", format="%.2f",
                    help="Auto: pcs × N.W + empty carton wt"),
                "NET WEIGHT": st.column_config.NumberColumn(
                    "NET WEIGHT", format="%.2f",
                    help="Auto: gross − net deduct"),
            }
            for s in sizes:
                cfg[s] = st.column_config.NumberColumn(s, min_value=0, step=1, format="%d")
    
            edited = st.data_editor(
                st.session_state[data_key],
                key=f"{state_key}__ed{st.session_state[nonce_key]}",
                num_rows="dynamic",
                hide_index=True,
                column_config=cfg,
                # calculated columns are read-only; user edits CTN QTY + sizes
                disabled=CALC_COLS,
                use_container_width=True,
            )
    
            # ---- Always recompute calculated columns from CTN QTY + size cells ----
            def _apply_formulas(frame: pd.DataFrame) -> pd.DataFrame:
                out = frame.copy()
                for idx, row in out.iterrows():
                    size = None
                    size_sum = 0
                    for s in sizes:
                        v = _to_int(row.get(s), 0)
                        size_sum += v
                        if v > 0 and size is None:
                            size = s
                    ctn = _to_int(row.get("CTN QTY"), 0)
                    # FORMULA: QTY PER CTN = sum of all size columns
                    qty_per = size_sum
                    # FORMULA: TOTAL QTY = CTN QTY × QTY PER CTN
                    total_qty = ctn * qty_per
                    code = str(row.get("CTN MES") or "").strip()
                    if not code and size and qty_per > 0:
                        try:
                            code = carton_code(size, qty_per)
                        except Exception:
                            code = "G81"
                    gross = net = None
                    if size and qty_per > 0:
                        try:
                            gross, net = carton_weights(size, qty_per, code or "G81", ptype, tare=user_tare)
                        except Exception:
                            pass
                    out.at[idx, "QTY PER CTN"] = int(qty_per)
                    out.at[idx, "TOTAL QTY"] = int(total_qty)
                    if gross is not None:
                        out.at[idx, "GROSS WEIGHT"] = float(gross)
                    if net is not None:
                        out.at[idx, "NET WEIGHT"] = float(net)
                    if code:
                        out.at[idx, "CTN MES"] = code
                return out
    
            recalc = _apply_formulas(edited)
    
            # Detect if calculated values differ from what the grid is showing
            def _sig(frame):
                parts = []
                for _, row in frame.iterrows():
                    parts.append((
                        _to_int(row.get("CTN QTY"), 0),
                        tuple(_to_int(row.get(s), 0) for s in sizes),
                        _to_int(row.get("QTY PER CTN"), 0),
                        _to_int(row.get("TOTAL QTY"), 0),
                    ))
                return tuple(parts)
    
            if _sig(recalc) != _sig(edited):
                st.session_state[data_key] = recalc.reset_index(drop=True)
                st.session_state[nonce_key] = int(st.session_state.get(nonce_key, 0)) + 1
                st.rerun()
            # keep data_key synced even when sig matches (e.g. after stable load)
            st.session_state[data_key] = recalc.reset_index(drop=True)
    
            b1, b2, b3 = st.columns([1.4, 1.4, 4])
            n_tick = int(recalc["🗑"].fillna(False).astype(bool).sum()) if "🗑" in recalc.columns else 0
            if b1.button(f"🗑 Remove ticked rows ({n_tick})", key=f"{state_key}__rm", disabled=n_tick == 0):
                keep = recalc.loc[~recalc["🗑"].fillna(False).astype(bool)].copy()
                keep["🗑"] = False
                st.session_state[data_key] = keep.reset_index(drop=True)
                st.session_state[nonce_key] += 1
                st.rerun()
            if b2.button("↺ Reset table to PO", key=f"{state_key}__rst"):
                st.session_state[data_key] = df.copy()
                st.session_state[nonce_key] += 1
                st.rerun()
    
            # ---- build section from recalculated rows ----
            recs = []
            for _, row in recalc.iterrows():
                if bool(row.get("🗑", False)):
                    continue
                size = None
                size_val = 0
                size_sum = 0
                for s in sizes:
                    v = _to_int(row.get(s), 0)
                    size_sum += v
                    if v > 0 and size is None:
                        size, size_val = s, v
                ctn = _to_int(row.get("CTN QTY"), 0)
                pcs = size_sum if size_sum > 0 else _to_int(row.get("QTY PER CTN"), 0)
                if pcs <= 0 and size_val <= 0 and ctn <= 0:
                    continue
                if not size:
                    size = sizes[0] if sizes else ""
                pcs = pcs or size_val
                ctn = ctn or 1
                recs.append({
                    "size": size,
                    "pcs": pcs,
                    "ctn": ctn,
                    "prepack": str(row.get("PREPACK STECKER") or base.get("prepack") or ""),
                })
    
            sec, probs = rebuild_one_size_section(base, recs, next_ctn)
            kit.problems(probs)
            next_ctn += sec["ctn_total"]
            out.append(sec)
    
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Ship qty", f"{sec['ship_qty']} PCS")
            m2.metric("Cartons", sec["ctn_total"])
            m3.metric("Order qty", f"{sec['qty_ordered']} PCS")
            diff = sec["ship_qty"] - sec["qty_ordered"]
            m4.metric("Excess / short", f"{diff:+d}")
            if ptype == "Bulk":
                order = {
                    sz: int(sec["units"].get(sz, 0)) - int(sec["excess_by_size"].get(sz, 0))
                    for sz in sec["sizes"]
                }
                kit.check_by_size(list(sec["sizes"]), order, sec["units"], over_pct=float(excess_pct))
            elif diff == 0:
                st.success("Ship qty matches the PO.")
            elif diff > 0:
                st.warning(f"Ship qty is {diff} pcs above the PO.")
            else:
                st.error(f"Ship qty is {-diff} pcs below the PO.")
            st.divider()
        return out, next_ctn
    
    
    
    next_ctn = 1
    bulk_secs, single_secs, multi_y_secs, multi_n_secs = [], [], [], []
    multi_n_xlsx = None
    for ptype in PACK_ORDER:
        if ptype == "Bulk" and auto_bulk:
            with tab_of["Bulk"]:
                bulk_secs, next_ctn = _editable_sections("Bulk", auto_bulk, next_ctn)
        elif ptype == "Single" and auto_single:
            with tab_of["Single"]:
                single_secs, next_ctn = _editable_sections("Single", auto_single, next_ctn)
        elif ptype == "Multi Y" and detected["Multi Y"]:
            multi_y_secs = build_multi_y_by_color(data, start_ctn=next_ctn)
            next_ctn += sum(s["ctn_total"] for s in multi_y_secs)
        elif ptype == "Multi N" and detected["Multi N"]:
            try:
                multi_n_secs = build_multi_n(
                    data, nw=DEFAULT_NW, tare_by_code=DEFAULT_TARE, net_deduct=net_deduct,
                    default_per_ctn=prepacks_per_ctn, ctn_code=ctn_code, start_ctn_no=next_ctn,
                )
                next_ctn += sum(s.get("ctn_total", 0) for s in multi_n_secs)
            except Exception as e:
                st.error(f"Multi N build failed: {e}")
    
    _secs_of = {"Bulk": bulk_secs, "Single": single_secs, "Multi Y": multi_y_secs, "Multi N": multi_n_secs}
    all_secs = [sec for _t in PACK_ORDER for sec in _secs_of[_t]]
    grand_ctn = sum(s.get("ctn_total", 0) for s in all_secs)
    grand_pcs = sum(s.get("ship_qty", 0) for s in all_secs)
    
    if multi_n_secs:
        try:
            multi_n_xlsx = export_multi_n_excel(
                multi_n_secs, DEFAULT_NW, DEFAULT_NNW, DEFAULT_TARE, net_deduct,
                list(DEFAULT_CTN_MEAS), buyer, po_no, grand_ctn,
            )
        except Exception as e:
            st.error(f"Multi N build failed: {e}")
    
    for ptype in PACK_ORDER:
        if ptype not in ("Multi Y", "Multi N"):
            continue
        secs = _secs_of[ptype]
        if not secs:
            continue
        with tab_of[ptype]:
            for sec in secs:
                render_section_strip(STRIP[ptype])
                render_po_header_html(sec, buyer, po_no, grand_ctn)
                if ptype == "Multi N":
                    sizes = sec.get("sizes") or []
                    recs = []
                    for r in sec.get("rows") or []:
                        rec = {
                            "CTN NO": (f"{r['ctn_from']}-{r['ctn_to']}" if r.get("ctn_to") != r.get("ctn_from") else str(r.get("ctn_from"))),
                            "CTN MES": r.get("ctn_mes", ""), "CTN QTY": r.get("ctn_qty"),
                            "COLOR": sec.get("color"), "PREPACK": sec.get("prepack"),
                        }
                        for sz in sizes:
                            rec[sz] = (sec.get("units") or {}).get(sz, "")
                        rec["PER BLST"] = r.get("per_blst", "")
                        rec["QTY PER CTN"] = r.get("qty_per_ctn", "")
                        rec["TOTAL QTY"] = r.get("total_qty", "")
                        rec["GROSS"] = r.get("gross", "")
                        rec["NET"] = r.get("net", "")
                        recs.append(rec)
                    st.dataframe(pd.DataFrame(recs), hide_index=True, **kit.STRETCH)
                else:
                    render_plist_table(sec)
                c1, c2, c3 = st.columns(3)
                c1.metric("Ship qty", f"{sec['ship_qty']} PCS")
                c2.metric("Cartons", sec["ctn_total"])
                c3.metric("Color", sec.get("color", ""))
                st.divider()
    
    with totals_box:
        exs = grand_pcs - oiqty
        kit.summary_cards([("Total cartons", f"{grand_ctn:,}"), ("Total pieces", f"{grand_pcs:,}"),
                           ("Order qty", f"{oiqty}"), ("Over / short", f"{exs:+,}")])
    
    with tabs[len(tab_types)]:
        for tname, table in (data.get("tables") or {}).items():
            ptype = detect_table_type(table) or "?"
            rows = table.get("rows") or []
            with st.expander(f"{tname} — {ptype} — {len(rows)} rows"):
                if rows:
                    st.dataframe(rows, **kit.STRETCH)
                if table.get("total"):
                    st.json(table["total"])
    
    with tabs[len(tab_types) + 1]:
        st.json(data)
        st.download_button(
            "⬇ Download extracted JSON",
            data=json.dumps(data, indent=2, ensure_ascii=False),
            file_name=f"{po_no or Path(source_name).stem}.json",
            mime="application/json",
        )
    
    # ---- 4 check ----------------------------------------------------------------------
    kit.step(4, "Check against the PO")
    if grand_pcs == oiqty:
        st.success(f"All sections together ship {grand_pcs:,} pcs = the PO order quantity.")
    elif grand_pcs > oiqty:
        st.warning(f"All sections together ship {grand_pcs - oiqty:,} pcs above the PO ({grand_pcs:,} vs {oiqty}). "
                   f"Bulk may carry up to {excess_pct:g}% excess per size.")
    else:
        st.error(f"All sections together ship {oiqty - grand_pcs:,} pcs below the PO ({grand_pcs:,} vs {oiqty}).")
    
    # ---- 5 download ---------------------------------------------------------------------
    kit.step(5, "Download", "The Excel is rebuilt on every change, so it always has your latest edits.")
    try:
        packing_all = export_unified_excel(
            data, bulk_secs, single_secs, multi_y_secs, multi_n_secs,
            nw=user_nw, nnw=user_nnw, tare_by_code=user_tare, ctn_box=user_box,
            net_deduct={"Multi N": net_deduct},
            order=list(PACK_ORDER),
        )
    except Exception as e:
        packing_all = None
        st.error(f"Excel build failed: {e}")
    
    dl1, dl2 = st.columns(2)
    with dl1:
        if packing_all:
            kit.download_block(
                "⬇ Download packing list for all",
                packing_all,
                f"packing_list_ALL_{po_no or 'PO'}.xlsx",
                f"Single sheet · Bulk+Single+Multi · {grand_ctn:,} ctns · {grand_pcs:,} pcs · formulas",
            )
    with dl2:
        if multi_n_xlsx:
            kit.download_block(
                "⬇ Multi N template Excel only",
                multi_n_xlsx,
                f"packing_list_multi_n_{po_no}.xlsx",
            )
        else:
            st.caption("Multi N template appears when Multi N tables are present.")


if __name__ == "__main__":
    st.set_page_config(page_title="GAP PO -> Packing List", layout="wide", page_icon="📦")
    st.title("📦 GAP / Old Navy — PO → Packing List")
    st.caption(
        "Upload PDF → extract → auto-detect **Bulk / Single / Multi Y / Multi N** → "
        "merge same STYLE by **COLOR** → interactive packing list + Excel"
    )
    run()
