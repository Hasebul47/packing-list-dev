"""
pl_kit.py - shared UI building blocks for EVERY buyer page
===========================================================
One flow, one look:

    1  PO summary        summary_cards()
    2  Packing table     editable_grid()  (edit cells, + add row, tick-and-remove rows, reset)
    3  Variables         (buyer page: carton size, weights, header ...)
    4  Live check        check_by_size()  (order vs packed, green / orange / red)
    5  Download          download_block()

Only stable Streamlit APIs are used (st.data_editor, st.metric, st.container).
Buyer modules keep their own parsing + Excel rules; this file only draws the screen.
"""
from __future__ import annotations

import pandas as pd
import streamlit as st

_V = tuple(int(x) for x in st.__version__.split(".")[:2])
STRETCH = {"width": "stretch"} if _V >= (1, 50) else {"use_container_width": True}


def step(n: int, title: str, hint: str = "") -> None:
    st.markdown(f"#### {n} · {title}")
    if hint:
        st.caption(hint)


def summary_cards(items: list[tuple[str, str]]) -> None:
    """items = [(label, value), ...] shown as a row of cards."""
    with st.container(border=True):
        cols = st.columns(len(items))
        for col, (label, value) in zip(cols, items):
            col.metric(label, value)


DEL = "🗑 Delete"


def editable_grid(df: pd.DataFrame, key: str, column_config: dict | None = None,
                  disabled: list[str] | None = None, dynamic: bool = True) -> pd.DataFrame:
    """
    Editable table with VISIBLE row controls:
      * edit any cell
      * "+" row at the bottom of the table adds a row
      * tick the first column (🗑 Delete) on any rows, then press "Remove ticked rows"
      * "Reset table to PO" restores the original table
    Rows that are ticked are ignored by the returned table straight away.
    """
    base_key, nonce_key = f"{key}__base", f"{key}__nonce"
    st.session_state.setdefault(nonce_key, 0)
    if base_key not in st.session_state:
        st.session_state[base_key] = df.reset_index(drop=True).copy()
    base = st.session_state[base_key].copy()
    if DEL in base.columns:
        base = base.drop(columns=[DEL])
    base.insert(0, DEL, False)

    cfg = {DEL: st.column_config.CheckboxColumn(DEL, help="Tick the rows to remove", default=False, width="small")}
    cfg.update(column_config or {})
    edited = st.data_editor(
        base, key=f"{key}__ed{st.session_state[nonce_key]}", hide_index=True,
        num_rows="dynamic" if dynamic else "fixed",
        column_config=cfg, disabled=disabled or [], **STRETCH,
    )
    ticked = edited[DEL].fillna(False).astype(bool)
    n_ticked = int(ticked.sum())

    c1, c2, c3 = st.columns([1.3, 1.3, 4])
    if c1.button(f"🗑 Remove ticked rows ({n_ticked})", key=f"{key}__rm", disabled=n_ticked == 0):
        keep = edited[~ticked].drop(columns=[DEL]).reset_index(drop=True)
        st.session_state[base_key] = keep
        st.session_state[nonce_key] += 1
        st.rerun()
    if c2.button("↺ Reset table to PO", key=f"{key}__reset", help="Discard all edits and go back to what the PO says"):
        st.session_state[base_key] = df.reset_index(drop=True).copy()
        st.session_state[nonce_key] += 1
        st.rerun()
    if n_ticked:
        c3.caption(f"{n_ticked} ticked row(s) are ignored below. Press “Remove ticked rows” to delete them for good.")
    return edited[~ticked].drop(columns=[DEL]).reset_index(drop=True)


def check_by_size(sizes: list[str], order: dict, packed: dict,
                  over_pct: float = 0.0, under_pct: float = 0.0) -> bool:
    """Order vs packed per size, with a plain-language verdict. True when everything is OK.
    over_pct / under_pct = ship tolerance of the PO in % (0 = must match exactly)."""
    tol = over_pct > 0 or under_pct > 0
    diff = {s: int(packed.get(s, 0)) - int(order.get(s, 0)) for s in sizes}

    def pct(s):
        o = int(order.get(s, 0))
        return diff[s] / o * 100 if o else 0.0

    def ok(s):
        return diff[s] == 0 or (tol and -under_pct <= pct(s) <= over_pct)

    def status(s):
        if diff[s] == 0:
            return "✅ OK"
        if ok(s):
            return f"✅ {diff[s]:+d} (within tolerance)"
        return f"🟠 {diff[s]:+d} extra" if diff[s] > 0 else f"🔴 {-diff[s]} short"

    df = pd.DataFrame({
        "Size": sizes,
        "Order": [int(order.get(s, 0)) for s in sizes],
        "Packed": [int(packed.get(s, 0)) for s in sizes],
        "Difference": [diff[s] for s in sizes],
        "%": [round(pct(s), 2) for s in sizes],
        "Status": [status(s) for s in sizes],
    })
    st.dataframe(df, hide_index=True, **STRETCH)
    short = [s for s in sizes if not ok(s) and diff[s] < 0]
    extra = [s for s in sizes if not ok(s) and diff[s] > 0]
    if not short and not extra:
        st.success("All sizes match the PO" + (f" (tolerance +{over_pct:g}% / -{under_pct:g}%)." if tol else "."))
        return True
    if short:
        st.error("Short vs PO: " + ", ".join(f"{s} ({-diff[s]})" for s in short))
    if extra:
        st.warning("Extra vs PO: " + ", ".join(f"{s} (+{diff[s]})" for s in extra))
    return False


def download_block(label: str, data: bytes, file_name: str, caption: str = "") -> None:
    st.download_button(
        label, data=data, file_name=file_name, type="primary",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if caption:
        st.caption(caption)


def problems(msgs: list[str]) -> None:
    for m in msgs:
        st.warning(m)
