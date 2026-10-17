"""
grid_editor.py - Reusable Interactive Carton Packing Grid
=========================================================
Extracted and generalized from the VF Vans gold-standard implementation.

Features:
- Sticky left frozen columns (+, -, #, Ctn No, plus custom buyer columns like Color, UPC, Ref)
- Horizontally scrollable dynamic size columns with in-cell numeric editing (st.number_input)
- Trailing calculation columns (CTN PCS, TOTAL CTN [editable], TOTAL PCS, Grs/ctn, Net/ctn, tot Grs, tot Net)
- Sticky top pinned rows: HEADER and ORDER QTY
- Sticky bottom pinned rows: TOTAL and REMAIN (color-coded live balance)
- Row insertion (＋) below any row, and row deletion (❌)
- Reset button (🔄 Reset) restoring rows to PO defaults
- Green border dirty-state tracking (mark_if_dirty) for modified cells
- Sub-second performance via Streamlit fragment (@st.fragment)
- Synchronized return state for Excel/PDF export
"""
from __future__ import annotations

import copy
import re
import uuid
from html import escape
from typing import Any, Callable

import streamlit as st

# ============================================================
# Fragment helper for fast, local reruns without page flash
# ============================================================
_HAS_FRAGMENT = hasattr(st, "fragment")


def _fragment(fn: Callable) -> Callable:
    return st.fragment(fn) if _HAS_FRAGMENT else fn


def _rerun_preview() -> None:
    """Rerun only the preview fragment (fast); fall back to full rerun."""
    if _HAS_FRAGMENT:
        try:
            st.rerun(scope="fragment")
            return
        except Exception:
            pass
    st.rerun()


def _set_state(key: str, value: Any) -> None:
    st.session_state[key] = value


# ============================================================
# Green-border CSS + marker helper
# ============================================================
DIRTY_RING_COLOR = "#16a34a"


def inject_dirty_ring_css() -> None:
    st.markdown(
        f"""
        <style>
        [data-testid="stMarkdownContainer"]:has(> .dirty-flag),
        [data-testid="stElementContainer"]:has(> [data-testid="stMarkdownContainer"] > .dirty-flag) {{
            display: none !important;
            height: 0 !important;
            min-height: 0 !important;
            max-height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
            overflow: hidden !important;
        }}
        .dirty-flag {{ display: none !important; }}

        [data-testid="stColumn"]:has(.dirty-flag) input[type="number"],
        [data-testid="stColumn"]:has(.dirty-flag) input[type="text"] {{
            border-color: {DIRTY_RING_COLOR} !important;
            box-shadow: inset 0 0 0 1px {DIRTY_RING_COLOR} !important;
        }}
        [data-testid="stColumn"]:has(.dirty-flag) [data-baseweb="input"],
        [data-testid="stColumn"]:has(.dirty-flag) [data-baseweb="base-input"],
        [data-testid="stColumn"]:has(.dirty-flag) [data-baseweb="input"] > div,
        [data-testid="stColumn"]:has(.dirty-flag) [data-baseweb="base-input"] > div {{
            border-color: {DIRTY_RING_COLOR} !important;
            box-shadow: 0 0 0 1px {DIRTY_RING_COLOR} !important;
        }}
        [data-testid="stColumn"]:has(.dirty-flag) [data-testid="stNumberInput"] > div,
        [data-testid="stColumn"]:has(.dirty-flag) [data-testid="stNumberInput"] > div > div,
        [data-testid="stColumn"]:has(.dirty-flag) [data-testid="stTextInput"] > div,
        [data-testid="stColumn"]:has(.dirty-flag) [data-testid="stTextInput"] > div > div {{
            border-color: {DIRTY_RING_COLOR} !important;
            box-shadow: 0 0 0 1px {DIRTY_RING_COLOR} !important;
            overflow: visible !important;
        }}
        [data-testid="stColumn"]:has(.dirty-flag) [data-testid="stNumberInput"],
        [data-testid="stColumn"]:has(.dirty-flag) [data-testid="stTextInput"] {{
            overflow: visible !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def mark_if_dirty(is_dirty: bool) -> None:
    if is_dirty:
        st.markdown(
            '<span class="dirty-flag" aria-hidden="true">*</span>',
            unsafe_allow_html=True,
        )


def _grid_key(key_prefix: str) -> str:
    """Unique, CSS-safe container key per grid (several grids can be on one page,
    e.g. one per colour in Banana Republic). Streamlit turns it into the class st-key-<key>."""
    return "carton_grid_" + re.sub(r"[^A-Za-z0-9_-]", "_", str(key_prefix))


def _sticky_sel(prefix: str, gk: str) -> str:
    p = f'[class*="st-key-{prefix}"]'
    g = f".st-key-{gk}"
    return (
        f'{g} > *:has(> {p}), '
        f'{g} > [data-testid="stVerticalBlock"] > *:has(> {p}), '
        f'{g} > {p}, '
        f'{g} > [data-testid="stVerticalBlock"] > {p}'
    )


# ============================================================
# Core Grid Editor Renderer
# ============================================================
@_fragment
def render_carton_grid(
    key_prefix: str,
    sizes: list[str],
    order_qty: dict[str, int],
    initial_rows: list[dict],
    frozen_columns: list[dict] | None = None,
    net_weight_by_size: dict[str, float] | None = None,
    empty_carton_weight: float = 0.0,
    show_weights: bool = True,
    show_metrics: bool = True,
    title: str = "Main carton table (editable)",
    caption: str = (
        "Main carton table is fully editable. Left columns are frozen. "
        "Header, ORDER QTY, TOTAL and REMAIN stay pinned vertically. "
        "Edited cells get a green border."
    ),
    render_below: Callable[[dict], None] | None = None,
) -> dict:
    """Render an interactive, frozen-column carton packing grid.

    Parameters:
      key_prefix: Unique string namespace for session_state.
      sizes: Ordered list of size names (e.g. ['S', 'M', 'L'] or ['28', '30']).
      order_qty: Dict mapping size -> ordered pieces count.
      initial_rows: List of dicts representing initial carton breakdown:
          [{"color": "Navy", "upc": "...", "sizes": {"S": 12, "M": 12}, "total_ctn": 10}, ...]
      frozen_columns: Optional list of additional frozen left column definitions:
          [{"header": "Color", "key": "color", "width": 130},
           {"header": "UPC",   "key": "upc",   "width": 130}]
      net_weight_by_size: Dict mapping size -> kg net weight per piece.
      empty_carton_weight: Empty carton tare weight (kg).
      show_weights: Whether to display Grs/Net weight columns.
      show_metrics: Whether to render metric cards (Total Cartons, Pieces, Wt) and over/short badge.

    Returns:
      Dict with edited carton rows, totals, remainders, and dirty status.
    """
    inject_dirty_ring_css()
    GK = _grid_key(key_prefix)

    full_sizes = list(sizes)
    n_sizes = len(full_sizes)
    order_by_size = {sz: int(order_qty.get(sz, 0) or 0) for sz in full_sizes}
    order_total = sum(order_by_size.values())

    nw_by_size = net_weight_by_size or {}
    empty_ctn_wt = float(empty_carton_weight or 0.0)

    # Default frozen columns if not specified
    if frozen_columns is None:
        frozen_columns = [
            {"header": "Color", "key": "color", "width": 130},
            {"header": "UPC",   "key": "upc",   "width": 130},
        ]

    # Session state keys
    rows_key = f"{key_prefix}_carton_rows"
    defaults_key = f"{key_prefix}_carton_defaults"
    insert_key = f"{key_prefix}_insert_idx"
    delete_key = f"{key_prefix}_delete_idx"

    # Initialize state if not present
    if rows_key not in st.session_state:
        clean_rows = []
        for r in initial_rows:
            r_id = r.get("id") or uuid.uuid4().hex
            row_dict = {
                "id": r_id,
                "sizes": {sz: int(r.get("sizes", {}).get(sz, 0) or 0) for sz in full_sizes},
                "total_ctn": int(r.get("total_ctn", 1)),
                "mixed": bool(r.get("mixed", False)),
            }
            # copy custom frozen column values
            for fc in frozen_columns:
                k = fc["key"]
                row_dict[k] = r.get(k, "")
            clean_rows.append(row_dict)

        st.session_state[rows_key] = clean_rows
        st.session_state[defaults_key] = {
            r["id"]: {
                "sizes": dict(r["sizes"]),
                "total_ctn": int(r["total_ctn"]),
            }
            for r in clean_rows
        }

    carton_rows = st.session_state[rows_key]
    _all_defaults = st.session_state.get(defaults_key, {})

    # Handle row insertion
    if st.session_state.get(insert_key) is not None:
        idx = st.session_state[insert_key]
        new_id = uuid.uuid4().hex
        prev_row = carton_rows[idx] if 0 <= idx < len(carton_rows) else {}
        new_row = {
            "id": new_id,
            "sizes": {sz: 0 for sz in full_sizes},
            "total_ctn": 1,
            "mixed": True,
        }
        for fc in frozen_columns:
            k = fc["key"]
            new_row[k] = prev_row.get(k, "")

        carton_rows.insert(idx + 1, new_row)
        st.session_state[rows_key] = carton_rows
        _all_defaults[new_id] = {
            "sizes": dict(new_row["sizes"]),
            "total_ctn": int(new_row["total_ctn"]),
        }
        st.session_state[defaults_key] = _all_defaults
        st.session_state[insert_key] = None

    # Handle row deletion
    if st.session_state.get(delete_key) is not None:
        idx = st.session_state[delete_key]
        if len(carton_rows) > 1 and 0 <= idx < len(carton_rows):
            removed_id = carton_rows[idx]["id"]
            carton_rows.pop(idx)
            st.session_state[rows_key] = carton_rows
            _all_defaults.pop(removed_id, None)
            st.session_state[defaults_key] = _all_defaults
        st.session_state[delete_key] = None

    n_rows = len(carton_rows)

    # Layout constants
    ACTION_WIDTH = 42
    NUM_WIDTH = 50
    CTN_WIDTH = 85
    CELL_WIDTH = 75
    CELL_HEIGHT = 42
    HEADER_HEIGHT = 44
    TOTAL_HEIGHT = 44
    GRID_MAX_HEIGHT = 520

    # Build frozen column widths and offsets
    # Base columns: [Add(1), Del(2), #(3), Ctn No(4)] + Custom frozen cols(5...)
    fixed_widths = [ACTION_WIDTH, ACTION_WIDTH, NUM_WIDTH, CTN_WIDTH]
    fixed_widths.extend(int(fc.get("width", 120)) for fc in frozen_columns)
    num_frozen_cols = len(fixed_widths)

    offsets = [0]
    for w in fixed_widths[:-1]:
        offsets.append(offsets[-1] + w)

    # Column specifications for Streamlit st.columns
    # Frozen cols + Size cols + Trailing calculation cols
    trailing_headers = ["CTN PCS", "TOTAL CTN", "TOTAL PCS"]
    if show_weights:
        trailing_headers.extend(["Grs/ctn", "Net/ctn", "tot Grs", "tot Net"])

    col_spec = fixed_widths + [CELL_WIDTH] * n_sizes + [CELL_WIDTH] * len(trailing_headers)

    # Generate CSS for scroller, frozen columns, and sticky headers/footers
    frozen_css_rules = []
    for c_idx in range(num_frozen_cols):
        nth = c_idx + 1
        w_px = fixed_widths[c_idx]
        off_px = offsets[c_idx]
        shadow_prop = "box-shadow: 2px 0 0 0 #94a3b8 !important;" if nth == num_frozen_cols else ""
        frozen_css_rules.append(
            f"""
            .st-key-{GK} [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child({nth}) {{
                position: sticky !important;
                z-index: 6 !important;
                background: white !important;
                left: {off_px}px !important;
                flex: 0 0 {w_px}px !important;
                width: {w_px}px !important;
                min-width: {w_px}px !important;
                max-width: {w_px}px !important;
                {shadow_prop}
            }}
            """
        )

    SEL_HDR = _sticky_sel(f"grid_hdr_{key_prefix}", GK)
    SEL_ORD = _sticky_sel(f"grid_ord_{key_prefix}", GK)
    SEL_TOT = _sticky_sel(f"grid_tot_{key_prefix}", GK)
    SEL_REM = _sticky_sel(f"grid_rem_{key_prefix}", GK)

    grid_css = f"""
    <style>
    .st-key-{GK} {{
        width: 100% !important;
        max-width: 100% !important;
        max-height: {GRID_MAX_HEIGHT}px !important;
        overflow: auto !important;
        box-sizing: border-box !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
        background: white !important;
        position: relative !important;
        display: block !important;
        gap: 0 !important;
        row-gap: 0 !important;
    }}
    .st-key-{GK} [data-testid="stVerticalBlock"],
    .st-key-{GK} [data-testid="stElementContainer"],
    .st-key-{GK} [data-testid="stHorizontalBlock"],
    .st-key-{GK} [data-testid="stColumn"] {{
        overflow: visible !important;
    }}
    .st-key-{GK} > [data-testid="stVerticalBlock"] {{
        display: contents !important;
    }}
    .st-key-{GK} > *,
    .st-key-{GK} > [data-testid="stVerticalBlock"] > * {{
        margin: 0 !important;
    }}
    /* Streamlit gives every markdown block a -1rem bottom margin, which made each header /
       ORDER / TOTAL / REMAIN row 16px shorter than its drawn height, so the rows overlapped
       and the carton rows were hidden behind them. Cancel it inside the grid. */
    .st-key-{GK} [data-testid="stMarkdownContainer"],
    .st-key-{GK} [data-testid="stMarkdown"] {{
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }}
    .st-key-{GK} [data-testid="stMarkdownContainer"] p {{
        margin: 0 !important;
    }}
    .st-key-{GK} [data-testid="stHorizontalBlock"] {{
        display: flex !important;
        flex-wrap: nowrap !important;
        gap: 0 !important;
        width: max-content !important;
        min-width: max-content !important;
    }}
    .st-key-{GK} [data-testid="stColumn"] {{
        flex-shrink: 0 !important;
        box-sizing: border-box !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
        margin: 0 !important;
    }}
    .st-key-{GK} [data-testid="stColumn"] > [data-testid="stVerticalBlock"] {{
        display: block !important;
        gap: 0 !important;
        row-gap: 0 !important;
    }}
    .st-key-{GK} [data-testid="stColumn"] > [data-testid="stVerticalBlock"] > div {{
        margin: 0 !important;
        padding: 0 !important;
    }}

    {"".join(frozen_css_rules)}

    /* Non-frozen columns */
    .st-key-{GK} [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(n+{num_frozen_cols + 1}) {{
        flex: 0 0 {CELL_WIDTH}px !important;
        width: {CELL_WIDTH}px !important;
        min-width: {CELL_WIDTH}px !important;
        max-width: {CELL_WIDTH}px !important;
    }}

    /* Pinned Top and Bottom rows */
    {SEL_HDR} {{
        position: sticky !important;
        top: 0 !important;
        z-index: 30 !important;
        background: white !important;
    }}
    {SEL_ORD} {{
        position: sticky !important;
        top: {HEADER_HEIGHT}px !important;
        z-index: 29 !important;
        background: white !important;
    }}
    {SEL_TOT} {{
        position: sticky !important;
        bottom: {TOTAL_HEIGHT}px !important;
        z-index: 29 !important;
        background: white !important;
    }}
    {SEL_REM} {{
        position: sticky !important;
        bottom: 0 !important;
        z-index: 30 !important;
        background: white !important;
    }}

    /* Header styling */
    .grid-header {{
        height: {HEADER_HEIGHT}px !important;
        min-height: {HEADER_HEIGHT}px !important;
        max-height: {HEADER_HEIGHT}px !important;
        box-sizing: border-box !important;
        display: flex; align-items: center; justify-content: center;
        background: #1f2937; color: white;
        border: 1px solid #4b5563 !important;
        font-size: 14px; font-weight: 600;
        white-space: nowrap; overflow: hidden;
    }}

    /* Cell styling */
    .data-cell {{
        height: {CELL_HEIGHT}px !important;
        min-height: {CELL_HEIGHT}px !important;
        max-height: {CELL_HEIGHT}px !important;
        box-sizing: border-box !important;
        display: flex; align-items: center; justify-content: center;
        border: 1px solid #d1d5db !important; background: #fff;
        font-size: 15px; white-space: nowrap; overflow: hidden;
    }}
    .data-cell.mixed {{ background: #FFFF66; }}

    /* Inputs in grid */
    .st-key-{GK} [data-testid="stNumberInput"] {{
        width: {CELL_WIDTH}px !important;
        min-width: {CELL_WIDTH}px !important;
        max-width: {CELL_WIDTH}px !important;
        margin: 0 !important; padding: 0 !important;
        overflow: visible !important;
    }}
    .st-key-{GK} [data-testid="stNumberInput"] > div,
    .st-key-{GK} [data-testid="stNumberInput"] > div > div,
    .st-key-{GK} [data-testid="stNumberInput"] [data-baseweb="input"],
    .st-key-{GK} [data-testid="stNumberInput"] [data-baseweb="base-input"] {{
        overflow: visible !important;
        border-radius: 0 !important;
    }}
    .st-key-{GK} [data-testid="stNumberInput"] input {{
        width: {CELL_WIDTH}px !important;
        min-width: {CELL_WIDTH}px !important;
        max-width: {CELL_WIDTH}px !important;
        height: {CELL_HEIGHT}px !important;
        min-height: {CELL_HEIGHT}px !important;
        max-height: {CELL_HEIGHT}px !important;
        box-sizing: border-box !important; border-radius: 0 !important;
        margin: 0 !important; padding: 0 4px !important;
        font-size: 15px !important; text-align: center !important;
    }}
    .st-key-{GK} [data-testid="stNumberInput"] label {{ display: none !important; }}

    /* Buttons in grid */
    .st-key-{GK} [data-testid="stButton"] {{
        width: {ACTION_WIDTH}px !important;
        min-width: {ACTION_WIDTH}px !important;
        max-width: {ACTION_WIDTH}px !important;
        margin: 0 !important; padding: 0 !important;
    }}
    .st-key-{GK} [data-testid="stButton"] button {{
        width: {ACTION_WIDTH}px !important;
        min-width: {ACTION_WIDTH}px !important;
        max-width: {ACTION_WIDTH}px !important;
        height: {CELL_HEIGHT}px !important;
        min-height: {CELL_HEIGHT}px !important;
        max-height: {CELL_HEIGHT}px !important;
        padding: 0 !important; margin: 0 !important;
        border-radius: 0 !important; font-size: 15px !important;
    }}

    /* Order Row */
    .order-cell {{
        height: {CELL_HEIGHT}px !important;
        box-sizing: border-box !important;
        display: flex; align-items: center; justify-content: center;
        background: #D9D9D9; color: #000;
        border: 1px solid #6b7280; font-weight: 700; font-size: 14px;
        white-space: nowrap; overflow: hidden;
    }}

    /* Total Row */
    .total-cell, .total-action-cell {{
        height: {TOTAL_HEIGHT}px !important;
        box-sizing: border-box !important;
        display: flex; align-items: center; justify-content: center;
        background: #D9D9D9; color: #000;
        border: 1px solid #6b7280; font-weight: 700; font-size: 14px;
        white-space: nowrap; overflow: hidden;
    }}

    /* Remain Row */
    .remain-cell, .remain-action-cell {{
        height: {TOTAL_HEIGHT}px !important;
        box-sizing: border-box !important;
        display: flex; align-items: center; justify-content: center;
        background: #fde68a; color: #78350f;
        border: 1px solid #6b7280; font-weight: 700; font-size: 14px;
        white-space: nowrap; overflow: hidden;
    }}
    .remain-cell.matched {{ background: #dcfce7; color: #166534; }}
    .remain-cell.overship {{ background: #fee2e2; color: #991b1b; }}
    </style>
    """
    st.markdown(grid_css, unsafe_allow_html=True)

    # Top Control Bar (Title + Reset)
    rc1, rc2 = st.columns([6, 1])
    with rc1:
        st.markdown(f"##### {title}")
    with rc2:
        if st.button("🔄 Reset", key=f"reset_{key_prefix}", help="Restore rows to PO defaults"):
            st.session_state.pop(rows_key, None)
            st.session_state.pop(defaults_key, None)
            for k in list(st.session_state.keys()):
                if k.startswith(f"sz_{key_prefix}_") or k.startswith(f"tctn_{key_prefix}_"):
                    st.session_state.pop(k, None)
            _rerun_preview()

    if caption:
        st.caption(caption)

    # Running calculations
    edited_sizes = []
    edited_total_ctn = []
    edited_ctn_pcs = []
    row_total_pcs = []
    grs_wts, net_wts = [], []
    ctn_ranges = []
    cum_ctn = 0
    is_any_dirty = False

    with st.container(key=GK):
        # 1. HEADER ROW
        with st.container(key=f"grid_hdr_{key_prefix}"):
            hdr_cols = st.columns(col_spec, gap="small")
            with hdr_cols[0]:
                st.markdown('<div class="grid-header">+</div>', unsafe_allow_html=True)
            with hdr_cols[1]:
                st.markdown('<div class="grid-header">−</div>', unsafe_allow_html=True)
            with hdr_cols[2]:
                st.markdown('<div class="grid-header">#</div>', unsafe_allow_html=True)
            with hdr_cols[3]:
                st.markdown('<div class="grid-header">Ctn No</div>', unsafe_allow_html=True)

            # Custom frozen column headers
            for fc_idx, fc in enumerate(frozen_columns):
                with hdr_cols[4 + fc_idx]:
                    st.markdown(f'<div class="grid-header">{escape(fc["header"])}</div>', unsafe_allow_html=True)

            # Size headers
            for s_idx, sz in enumerate(full_sizes):
                with hdr_cols[num_frozen_cols + s_idx]:
                    st.markdown(f'<div class="grid-header">{escape(sz)}</div>', unsafe_allow_html=True)

            # Trailing headers
            for t_idx, th in enumerate(trailing_headers):
                with hdr_cols[num_frozen_cols + n_sizes + t_idx]:
                    st.markdown(f'<div class="grid-header">{escape(th)}</div>', unsafe_allow_html=True)

        # 2. ORDER QTY ROW
        with st.container(key=f"grid_ord_{key_prefix}"):
            ord_cols = st.columns(col_spec, gap="small")
            for c_i in range(num_frozen_cols):
                with ord_cols[c_i]:
                    txt = "ORDER QTY" if c_i == num_frozen_cols - 1 else ""
                    st.markdown(f'<div class="order-cell">{txt}</div>', unsafe_allow_html=True)

            for s_idx, sz in enumerate(full_sizes):
                with ord_cols[num_frozen_cols + s_idx]:
                    st.markdown(f'<div class="order-cell">{order_by_size.get(sz, 0)}</div>', unsafe_allow_html=True)

            with ord_cols[num_frozen_cols + n_sizes]:
                st.markdown('<div class="order-cell"></div>', unsafe_allow_html=True)
            with ord_cols[num_frozen_cols + n_sizes + 1]:
                st.markdown('<div class="order-cell"></div>', unsafe_allow_html=True)
            with ord_cols[num_frozen_cols + n_sizes + 2]:
                st.markdown(f'<div class="order-cell">{order_total}</div>', unsafe_allow_html=True)

            if show_weights:
                for w_i in range(3, len(trailing_headers)):
                    with ord_cols[num_frozen_cols + n_sizes + w_i]:
                        st.markdown('<div class="order-cell"></div>', unsafe_allow_html=True)

        # 3. DATA ROWS
        for row_index, row in enumerate(carton_rows):
            row_id = row["id"]
            row_defaults = _all_defaults.get(row_id, {})
            def_sizes = row_defaults.get("sizes", {})
            def_tctn = int(row_defaults.get("total_ctn", 0))

            row_cols = st.columns(col_spec, gap="small")

            # Actions
            with row_cols[0]:
                st.button("＋", key=f"add_{key_prefix}_{row_id}", help="Add row below",
                          on_click=_set_state, args=(insert_key, row_index))
            with row_cols[1]:
                st.button("❌", key=f"del_{key_prefix}_{row_id}", help="Delete this row",
                          disabled=n_rows <= 1, on_click=_set_state, args=(delete_key, row_index))
            with row_cols[2]:
                st.markdown(f'<div class="data-cell">{row_index + 1}</div>', unsafe_allow_html=True)

            # Sizes numeric inputs
            current_sizes = {}
            for s_idx, sz in enumerate(full_sizes):
                with row_cols[num_frozen_cols + s_idx]:
                    sz_key = f"sz_{key_prefix}_{row_id}_{sz}"
                    curr_val = int(row["sizes"].get(sz, 0) or 0)
                    widget_val = st.session_state.get(sz_key)
                    effective = int(widget_val) if widget_val is not None else curr_val
                    def_val = int(def_sizes.get(sz, 0))
                    dirty = (effective != def_val)
                    if dirty:
                        is_any_dirty = True
                    mark_if_dirty(dirty)
                    val = st.number_input(
                        f"sz_{row_id}_{sz}", min_value=0, value=curr_val, step=1,
                        key=sz_key, label_visibility="collapsed",
                    )
                    current_sizes[sz] = int(val)
                    carton_rows[row_index]["sizes"][sz] = int(val)

            # TOTAL CTN input
            total_ctn_col_idx = num_frozen_cols + n_sizes + 1
            with row_cols[total_ctn_col_idx]:
                tctn_key = f"tctn_{key_prefix}_{row_id}"
                curr_tctn = int(row.get("total_ctn") or 0)
                widget_tctn = st.session_state.get(tctn_key)
                effective_tctn = int(widget_tctn) if widget_tctn is not None else curr_tctn
                dirty_tctn = (effective_tctn != def_tctn)
                if dirty_tctn:
                    is_any_dirty = True
                mark_if_dirty(dirty_tctn)
                new_tctn = st.number_input(
                    f"tctn_{row_id}", min_value=0, value=curr_tctn, step=1,
                    key=tctn_key, label_visibility="collapsed",
                )
                carton_rows[row_index]["total_ctn"] = int(new_tctn)
                row_tctn_val = int(new_tctn)

            # Calculations per row
            size_sum = sum(current_sizes.values())
            ctn_start = cum_ctn + 1
            cum_ctn += row_tctn_val
            ctn_end = cum_ctn
            ctn_range_str = f"{ctn_start}-{ctn_end}" if row_tctn_val > 0 else "-"
            tot_pcs = row_tctn_val * size_sum

            # Weights
            piece_wt = sum(
                float(current_sizes.get(sz, 0) or 0) * float(nw_by_size.get(sz, 0.0) or 0.0)
                for sz in full_sizes
            )
            grs = piece_wt + empty_ctn_wt if size_sum > 0 else 0.0
            net = piece_wt

            edited_sizes.append(current_sizes)
            edited_total_ctn.append(row_tctn_val)
            edited_ctn_pcs.append(size_sum)
            row_total_pcs.append(tot_pcs)
            grs_wts.append(grs)
            net_wts.append(net)
            ctn_ranges.append(ctn_range_str)

            # Render Ctn No
            with row_cols[3]:
                st.markdown(f'<div class="data-cell">{ctn_range_str}</div>', unsafe_allow_html=True)

            # Render custom frozen columns
            mixed_cls = " mixed" if row.get("mixed") else ""
            for fc_idx, fc in enumerate(frozen_columns):
                with row_cols[4 + fc_idx]:
                    val_str = str(row.get(fc["key"], "") or "")
                    st.markdown(f'<div class="data-cell{mixed_cls}">{escape(val_str)}</div>', unsafe_allow_html=True)

            # Render CTN PCS & TOTAL PCS
            with row_cols[num_frozen_cols + n_sizes]:
                st.markdown(f'<div class="data-cell"><b>{size_sum}</b></div>', unsafe_allow_html=True)
            with row_cols[num_frozen_cols + n_sizes + 2]:
                st.markdown(f'<div class="data-cell"><b>{tot_pcs}</b></div>', unsafe_allow_html=True)

            # Render weights
            if show_weights:
                with row_cols[num_frozen_cols + n_sizes + 3]:
                    st.markdown(f'<div class="data-cell">{grs:.2f}</div>', unsafe_allow_html=True)
                with row_cols[num_frozen_cols + n_sizes + 4]:
                    st.markdown(f'<div class="data-cell">{net:.2f}</div>', unsafe_allow_html=True)
                with row_cols[num_frozen_cols + n_sizes + 5]:
                    st.markdown(f'<div class="data-cell">{row_tctn_val * grs:.2f}</div>', unsafe_allow_html=True)
                with row_cols[num_frozen_cols + n_sizes + 6]:
                    st.markdown(f'<div class="data-cell">{row_tctn_val * net:.2f}</div>', unsafe_allow_html=True)

        st.session_state[rows_key] = carton_rows

        # Compute Grand Totals
        total_ctn = sum(edited_total_ctn)
        total_pcs = sum(row_total_pcs)
        total_grs = sum(tc * g for tc, g in zip(edited_total_ctn, grs_wts))
        total_net = sum(tc * n for tc, n in zip(edited_total_ctn, net_wts))
        size_packet_totals = {
            sz: sum(int(edited_sizes[i].get(sz, 0) or 0) * edited_total_ctn[i] for i in range(n_rows))
            for sz in full_sizes
        }

        # 4. TOTAL ROW
        with st.container(key=f"grid_tot_{key_prefix}"):
            tot_cols = st.columns(col_spec, gap="small")
            for c_i in range(num_frozen_cols):
                with tot_cols[c_i]:
                    txt = "TOTAL" if c_i == num_frozen_cols - 1 else ""
                    st.markdown(f'<div class="total-cell">{txt}</div>', unsafe_allow_html=True)

            for s_idx, sz in enumerate(full_sizes):
                with tot_cols[num_frozen_cols + s_idx]:
                    st.markdown(f'<div class="total-cell">{size_packet_totals.get(sz, 0)}</div>', unsafe_allow_html=True)

            with tot_cols[num_frozen_cols + n_sizes]:
                st.markdown('<div class="total-cell"></div>', unsafe_allow_html=True)
            with tot_cols[num_frozen_cols + n_sizes + 1]:
                st.markdown(f'<div class="total-cell">{total_ctn}</div>', unsafe_allow_html=True)
            with tot_cols[num_frozen_cols + n_sizes + 2]:
                st.markdown(f'<div class="total-cell">{total_pcs}</div>', unsafe_allow_html=True)

            if show_weights:
                with tot_cols[num_frozen_cols + n_sizes + 3]:
                    st.markdown('<div class="total-cell"></div>', unsafe_allow_html=True)
                with tot_cols[num_frozen_cols + n_sizes + 4]:
                    st.markdown('<div class="total-cell"></div>', unsafe_allow_html=True)
                with tot_cols[num_frozen_cols + n_sizes + 5]:
                    st.markdown(f'<div class="total-cell">{total_grs:.2f}</div>', unsafe_allow_html=True)
                with tot_cols[num_frozen_cols + n_sizes + 6]:
                    st.markdown(f'<div class="total-cell">{total_net:.2f}</div>', unsafe_allow_html=True)

        # 5. REMAIN ROW
        remain_by_size = {
            sz: order_by_size.get(sz, 0) - size_packet_totals.get(sz, 0)
            for sz in full_sizes
        }
        remain_total = order_total - total_pcs

        with st.container(key=f"grid_rem_{key_prefix}"):
            rem_cols = st.columns(col_spec, gap="small")
            for c_i in range(num_frozen_cols):
                with rem_cols[c_i]:
                    txt = "REMAIN" if c_i == num_frozen_cols - 1 else ""
                    st.markdown(f'<div class="remain-action-cell">{txt}</div>', unsafe_allow_html=True)

            for s_idx, sz in enumerate(full_sizes):
                diff = remain_by_size.get(sz, 0)
                cls = " matched" if diff == 0 else (" overship" if diff < 0 else "")
                with rem_cols[num_frozen_cols + s_idx]:
                    st.markdown(f'<div class="remain-cell{cls}">{diff}</div>', unsafe_allow_html=True)

            with rem_cols[num_frozen_cols + n_sizes]:
                st.markdown('<div class="remain-cell"></div>', unsafe_allow_html=True)
            with rem_cols[num_frozen_cols + n_sizes + 1]:
                st.markdown('<div class="remain-cell"></div>', unsafe_allow_html=True)
            tot_rem_cls = " matched" if remain_total == 0 else (" overship" if remain_total < 0 else "")
            with rem_cols[num_frozen_cols + n_sizes + 2]:
                st.markdown(f'<div class="remain-cell{tot_rem_cls}">{remain_total}</div>', unsafe_allow_html=True)

            if show_weights:
                for w_i in range(3, len(trailing_headers)):
                    with rem_cols[num_frozen_cols + n_sizes + w_i]:
                        st.markdown('<div class="remain-cell"></div>', unsafe_allow_html=True)

    # 6. LIVE METRIC CARDS & ALERTS
    diff_total = total_pcs - order_total
    if show_metrics:
        st.markdown("**Live totals**")
        if show_weights:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Cartons", f"{total_ctn:,}")
            m2.metric("Total Pieces", f"{total_pcs:,}")
            m3.metric("Total Gross Wt (kg)", f"{total_grs:,.2f}")
            m4.metric("Total Net Wt (kg)", f"{total_net:,.2f}")
        else:
            m1, m2 = st.columns(2)
            m1.metric("Total Cartons", f"{total_ctn:,}")
            m2.metric("Total Pieces", f"{total_pcs:,}")

        if diff_total > 0:
            st.success(f"✅ Over-ship: **+{diff_total:,}** pieces")
        elif diff_total < 0:
            st.error(f"⚠️ Short-ship: **{diff_total:,}** pieces (remain: {abs(diff_total):,})")
        else:
            st.info("✅ Ship quantity matches order quantity exactly.")

    # Return exported state dictionary
    exported_rows = []
    for i in range(n_rows):
        r_dict = {
            "ctn_no": ctn_ranges[i],
            "sizes": edited_sizes[i],
            "ctn_pcs": edited_ctn_pcs[i],
            "total_ctn": edited_total_ctn[i],
            "total_pcs": row_total_pcs[i],
            "grs_per_ctn": grs_wts[i] if show_weights else 0.0,
            "net_per_ctn": net_wts[i] if show_weights else 0.0,
            "total_grs": edited_total_ctn[i] * grs_wts[i] if show_weights else 0.0,
            "total_net": edited_total_ctn[i] * net_wts[i] if show_weights else 0.0,
        }
        for fc in frozen_columns:
            k = fc["key"]
            r_dict[k] = carton_rows[i].get(k, "")
        exported_rows.append(r_dict)

    res = {
        "rows": exported_rows,
        "size_totals": size_packet_totals,
        "remain_by_size": remain_by_size,
        "total_ctn": total_ctn,
        "total_pcs": total_pcs,
        "total_grs": total_grs,
        "total_net": total_net,
        "order_total": order_total,
        "ship_total": total_pcs,
        "diff_total": diff_total,
        "is_dirty": is_any_dirty,
    }

    if render_below is not None:
        render_below(res)

    return res
