"""
Unified PO -> Packing List router
=================================
TWO upload boxes:

  📄  PDF   → VF/Vans · GAP/Old Navy · Kontoor/Lee · s.Oliver · Carter's/OshKosh
  📊  XLSX  → Next Retail / Clover

Only one box at a time.  Both filled → warning, nothing runs.
Nothing filled → "upload a file" message.

Each buyer module exposes run(file).  Top-level scripts (like Vans) are
executed via runpy with st.set_page_config / st.file_uploader faked out.

Run:  streamlit run app.py
"""
import importlib
import io
import runpy
import sys
from pathlib import Path

import pdfplumber
import streamlit as st

BASE = Path(__file__).resolve().parent

# key -> (folder, filename, human label, "pdf" | "xlsx")
APPS = {
    "VF":  ("VF Vans",    "VF_Packing_V3_2_app.py", "VF (Vans)",                "pdf"),
    "GAP": ("Gap",        "unified_packing.py",     "GAP (Old Navy)",           "pdf"),
    "KT":  ("kontor",     "kontoor_app.py",         "Kontoor (Lee / Wrangler)", "pdf"),
    "SO":  ("s_oliver",   "s_oliver_app.py",        "s.Oliver",                 "pdf"),
    "CT":  ("carters",    "carters_app.py",         "Carter's / OshKosh",       "pdf"),
    "NX":  ("next_buyer", "next_buyer_app.py",      "Next Retail / Clover",     "xlsx"),
}


# ============================================================
# Locate the module script
# ============================================================
def _find_script(folder_name: str, filename: str):
    folder = BASE / folder_name

    # 1) exact path candidates
    for c in (folder / filename,
              BASE / filename,
              BASE.parent / folder_name / filename):
        if c.is_file():
            return c, c.parent

    # 2) anywhere inside that folder with matching stem
    if folder.is_dir():
        stem = Path(filename).stem
        for c in folder.rglob(f"{stem}.py"):
            if c.is_file():
                return c, c.parent

    # 3) recursive under BASE
    for c in BASE.rglob(filename):
        if c.is_file():
            return c, c.parent

    return None, None


# ============================================================
# Buyer detection (PDF only)
# ============================================================
def detect_buyer_pdf(pdf_bytes: bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        t = (pdf.pages[0].extract_text() or "").upper()

    if ("CARTERS" in t or "WILLIAM CARTER" in t or "OSHKOSH" in t
            or "E- PLM" in t or "E-PLM" in t):
        return "CT"
    if "S.OLIVER" in t or "S.OLIVER BERND FREIER" in t:
        return "SO"
    if "KONTOOR" in t:
        return "KT"
    if "VF OUTDOOR" in t or ("VANS" in t and "BUYER CODE" in t):
        return "VF"
    if "OLD NAVY" in t or "GAP INTL" in t or "DESTINATION PURCHASE ORDER" in t:
        return "GAP"
    return None


# ============================================================
# Dispatch
# ============================================================
def run_app(key: str, up_file):
    folder_name, filename, label, _ = APPS[key]
    script, folder = _find_script(folder_name, filename)

    if script is None:
        st.error(
            f"Could not find **{filename}** anywhere under `{BASE}`.\n\n"
            f"Looked in:\n"
            f"- `{BASE / folder_name / filename}`\n"
            f"- `{BASE / filename}`\n"
            f"- anything reachable under `{BASE}`"
        )
        st.stop()

    # fresh state when the buyer changes
    if st.session_state.get("_active_buyer") != key:
        for k in list(st.session_state.keys()):
            if not k.startswith(("_active_buyer", "router_")):
                st.session_state.pop(k, None)
        st.session_state["_active_buyer"] = key

    if str(folder) not in sys.path:
        sys.path.insert(0, str(folder))

    src = script.read_text(encoding="utf-8")

    # preferred: the module defines run(...)
    if "def run(" in src:
        mod = importlib.import_module(script.stem)
        mod.run(up_file)
        return

    # fallback: top-level Streamlit script (e.g. Vans)
    real_cfg, real_up = st.set_page_config, st.file_uploader

    def fake_uploader(label, type=None, *a, **kw):
        # Only the PO file box is fed by the router. Any other uploader a module
        # shows (e.g. GAP's "load an existing JSON") stays a real uploader.
        kinds = [str(t).lower().lstrip(".") for t in (type or [])]
        if kinds and set(kinds) <= {"json"}:
            return real_up(label, type=type, *a, **kw)
        return up_file

    st.set_page_config = lambda *a, **kw: None       # already set by router
    st.file_uploader  = fake_uploader                 # feed the router's file
    try:
        runpy.run_path(str(script), run_name="__main__")
    finally:
        st.set_page_config, st.file_uploader = real_cfg, real_up


# ============================================================
# Page
# ============================================================
st.set_page_config(page_title="PO -> Packing List", layout="wide", page_icon="📦")
st.title("📦 PO to Packing List")

col_pdf, col_xlsx = st.columns(2)
with col_pdf:
    pdf_up = st.file_uploader(
        "📄 Upload PO PDF",
        type=["pdf"],
        key="router_pdf",
        help="VF/Vans · GAP/Old Navy · Kontoor/Lee · s.Oliver · Carter's/OshKosh",
    )
with col_xlsx:
    xlsx_up = st.file_uploader(
        "📊 Upload XL file",
        type=["xlsx"],
        key="router_xlsx",
        help="Next Retail / Clover — Contract Confirmation Report (.xlsx; save old .xls as .xlsx first)",
    )

with st.sidebar:
    st.markdown("### Buyer (optional override)")
    override = st.selectbox(
        "Pick a buyer if auto-detection fails",
        ["Auto-detect",
         "VF (Vans)",
         "GAP (Old Navy)",
         "Kontoor (Lee / Wrangler)",
         "s.Oliver",
         "Carter's / OshKosh",
         "Next Retail / Clover"],
        key="router_override",
    )


# ============================================================
# 1) Both uploaded -> warn, do nothing
# ============================================================
if pdf_up is not None and xlsx_up is not None:
    st.warning("⚠️ **You uploaded both a PDF and an Excel file.** "
               "Please clear one box and upload a single file at a time.")
    st.stop()


# ============================================================
# 2) Nothing uploaded
# ============================================================
if pdf_up is None and xlsx_up is None:
    if override == "GAP (Old Navy)":
        # lets the GAP app's "load existing JSON" path work without a PDF
        st.divider()
        run_app("GAP", None)
    else:
        st.info("Upload **one** file — a PDF or an Excel file — to get started.")
        st.caption("PDF buyers: VF/Vans · GAP/Old Navy · Kontoor/Lee · "
                   "s.Oliver · Carter's/OshKosh")
        st.caption("Excel buyer: Next Retail / Clover")
    st.stop()


# ============================================================
# 3) PDF uploaded
# ============================================================
if pdf_up is not None:
    drop = {
        "VF (Vans)":                 "VF",
        "GAP (Old Navy)":            "GAP",
        "Kontoor (Lee / Wrangler)":  "KT",
        "s.Oliver":                  "SO",
        "Carter's / OshKosh":        "CT",
    }.get(override)

    if drop is None:                       # auto-detect
        try:
            drop = detect_buyer_pdf(pdf_up.getvalue())
        except Exception as e:
            st.error(f"Could not read the PDF: {e}")
            st.stop()

    if drop is None:
        st.warning("Could not recognise the PDF buyer. "
                   "Pick the buyer in the sidebar and try again.")
        st.stop()

    st.success(
        f"Detected buyer: **{APPS[drop][2]}**" if override == "Auto-detect"
        else f"Buyer set manually: **{APPS[drop][2]}**"
    )
    st.divider()
    run_app(drop, pdf_up)
    st.stop()


# ============================================================
# 4) XLSX uploaded
# ============================================================
if xlsx_up is not None:
    if override in ("Auto-detect", "Next Retail / Clover"):
        st.success("Excel file detected → **Next Retail / Clover**")
        st.divider()
        run_app("NX", xlsx_up)
    else:
        st.error(
            f"The **{override}** module expects a PDF, but you uploaded an Excel file. "
            f"Either pick **Next Retail / Clover** in the sidebar, or upload a PDF."
        )
    st.stop()