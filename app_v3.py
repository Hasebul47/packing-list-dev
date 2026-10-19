"""
Unified PO -> Packing List router
=================================
FOUR upload boxes (the first three build a packing list, the green one stores data):

  📄  PDF   → VF/Vans · GAP/Old Navy · Banana Republic · Kontoor/Lee
              · s.Oliver · Carter's/OshKosh
  📊  XLSX  → Next Retail / Clover
  📝  DOCX  → JCPenney / Thereabouts
  🟩  XL (green) → FINAL completed packing-list Excel files are stored in the
              database (final_db.py). Independent of the three boxes above.

Only one of the first three boxes at a time.  More than one filled → warning, nothing runs.
Nothing filled → info message.

Each buyer module exposes run(file) OR is a top-level Streamlit script;
the router handles both via runpy fallback.

Run:  streamlit run app.py
"""
import importlib
import io
import re
import runpy
import sys
import zipfile
from pathlib import Path

import pdfplumber
import streamlit as st

BASE = Path(__file__).resolve().parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

import theme                      # shared palette / smooth-UI css / xlsx polish
import final_db                   # green box: final completed XL files -> database

# ============================================================
# Buyer registry
# key -> (folder, filename, human label, file type)
# ============================================================
APPS = {
    "VF":  ("VF Vans",    "VF_Packing_V3_2_app.py", "VF (Vans)",                "pdf"),
    "GAP": ("Gap",        "unified_packing.py",     "GAP (Old Navy)",           "pdf"),
    "BR":  ("br",         "br_packing.py",          "Banana Republic",          "pdf"),
    "KT":  ("kontor",     "kontoor_app.py",         "Kontoor (Lee / Wrangler)", "pdf"),
    "SO":  ("s_oliver",   "s_oliver_app.py",        "s.Oliver",                 "pdf"),
    "CT":  ("carters",    "carters_app.py",         "Carter's / OshKosh",       "pdf"),
    "ANF": ("ABERCROMBIE", "anf_app.py",            "Abercrombie & Fitch",      "pdf"),
    "TW":  ("Tally_Weijl", "tw_app.py",             "Tally Weijl",              "pdf"),
    "JU":  ("justice",    "justice_app.py",         "Justice (SPA / Adjmi)",    "pdf"),
    "NX":  ("next",       "next_buyer_app.py",      "Next Retail / Clover",     "xlsx"),
    "HD":  ("haddad",     "haddad_app.py",          "Haddad (Levi's Kids)",     "xlsx"),
    "MJ":  ("MUJI",       "app.py",                 "MUJI",                     "xlsx"),
    "JC":  ("jcp",        "jcp_app.py",             "JCPenney / Thereabouts",   "docx"),
}

# label -> key  (for the dropdown override)
LABEL_TO_KEY = {v[2]: k for k, v in APPS.items()}


# ============================================================
# Locate the module script (very tolerant)
# ============================================================
def _find_script(folder_name: str, filename: str):
    folder = BASE / folder_name
    stem = Path(filename).stem

    # 1) exact path candidates
    for c in (folder / filename,
              BASE / filename,
              BASE.parent / folder_name / filename):
        if c.is_file():
            return c, c.parent

    # 2) anywhere inside that folder with matching stem
    if folder.is_dir():
        for c in folder.rglob(f"{stem}.py"):
            if c.is_file():
                return c, c.parent

    # 3) recursive anywhere
    for c in BASE.rglob(filename):
        if c.is_file():
            return c, c.parent

    return None, None


# ============================================================
# Buyer detection — PDF
# ============================================================
@st.cache_data(show_spinner=False, max_entries=32)
def detect_buyer_pdf(pdf_bytes: bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        t = (pdf.pages[0].extract_text() or "").upper()
        # buyer names can sit on page 2 or be written "Tally-Weijl" / "TALLY WEIJL":
        # look at the first two pages, ignoring spaces, hyphens and punctuation
        t2 = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:2]).upper()
    squashed = re.sub(r"[^A-Z0-9]", "", t2)

    # Abercrombie & Fitch
    if ("ABERCROMBIE" in t or "FITCH PATH" in t):
        return "ANF"
    # Tally Weijl
    if ("TALLYWEIJL" in squashed or "TSWL" in squashed or "VIADUKTSTRASSE" in squashed
            or ("ORDERFORM" in squashed and "SUPPLIERSTALLY" in squashed)):
        return "TW"
    # Justice - Sports Products of America / Adjmi Apparel Group
    if ("SPORTSPRODUCTSOFAMERICA" in squashed or "ADJMIAPPAREL" in squashed):
        return "JU"
    # Carter's / OshKosh
    if ("CARTERS" in t or "WILLIAM CARTER" in t or "OSHKOSH" in t
            or "E- PLM" in t or "E-PLM" in t):
        return "CT"
    # s.Oliver
    if "S.OLIVER" in t or "S.OLIVER BERND FREIER" in t:
        return "SO"
    # Kontoor / Lee / Wrangler
    if "KONTOOR" in t:
        return "KT"
    # VF / Vans
    if "VF OUTDOOR" in t or ("VANS" in t and "BUYER CODE" in t):
        return "VF"
    # Banana Republic / GAP Canada — MUST come before GAP check
    if ("BANANA REPUBLIC" in t or "GAP (CANADA)" in t
            or "GAP(CANADA)" in t or "PURCHASER" in t and "GAP (CANADA)" in t):
        return "BR"
    # GAP / Old Navy
    if "OLD NAVY" in t or "GAP INTL" in t or "DESTINATION PURCHASE ORDER" in t:
        return "GAP"
    return None


# ============================================================
# Buyer detection — XLSX (Next, Haddad, MUJI)
# ============================================================
@st.cache_data(show_spinner=False, max_entries=32)
def detect_buyer_xlsx(xlsx_bytes: bytes, file_name: str = ""):
    fn = (file_name or "").upper()
    if "MUJI" in fn or "発注" in fn:
        return "MJ"
    if "HADDAD" in fn:
        return "HD"
    if "NEXT" in fn or "CLOVER" in fn:
        return "NX"

    # Try openpyxl
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
        for sn in wb.sheetnames:
            sn_u = sn.upper()
            if "MUJI" in sn_u or "発注" in sn_u:
                return "MJ"
            if "STATUS" in sn_u or "HADDAD" in sn_u:
                return "HD"
        ws = wb.active or wb[wb.sheetnames[0]]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > 30:
                break
            row_str = " ".join(str(c or "") for c in row).upper()
            if "REF#" in row_str and "PO#" in row_str:
                return "HD"
            if "仕向地" in row_str or "発注一覧" in row_str:
                return "MJ"
            if "ITEM NUMBER" in row_str or "CONTRACT NO" in row_str:
                return "NX"
    except Exception:
        pass

    # Try xlrd for .xls
    try:
        import xlrd
        book = xlrd.open_workbook(file_contents=xlsx_bytes)
        sheet = book.sheet_by_index(0)
        for r in range(min(30, sheet.nrows)):
            row_vals = " ".join(str(sheet.cell_value(r, c)) for c in range(sheet.ncols)).upper()
            if "REF#" in row_vals and "PO#" in row_vals:
                return "HD"
            if "ITEM NUMBER" in row_vals or "CONTRACT NO" in row_vals:
                return "NX"
    except Exception:
        pass

    return "NX"


# ============================================================
# Buyer detection — DOCX (JCPenney)
# ============================================================
@st.cache_data(show_spinner=False, max_entries=32)
def detect_buyer_docx(docx_bytes: bytes):
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="ignore").upper()
    except Exception:
        return None
    if ("THEREABOUTS" in xml or "JCPPC" in xml
            or "JCP PURCHASING" in xml or "JCPENNEY" in xml):
        return "JC"
    return None


# ============================================================
# Dispatch  (module / compiled code are cached by file mtime, so a widget
# click no longer re-reads, re-compiles and re-imports the buyer script)
# ============================================================
@st.cache_resource(show_spinner=False)
def _load_module(path_str: str, mtime: float):
    path = Path(path_str)
    sys.modules.pop(path.stem, None)
    return importlib.import_module(path.stem)


@st.cache_resource(show_spinner=False)
def _load_code(path_str: str, mtime: float):
    src = Path(path_str).read_text(encoding="utf-8")
    return compile(src, path_str, "exec"), ("def run(" in src)


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

    # fresh session state when the buyer changes
    if st.session_state.get("_active_buyer") != key:
        for k in list(st.session_state.keys()):
            if not k.startswith(("_active_buyer", "router_")):
                st.session_state.pop(k, None)
        st.session_state["_active_buyer"] = key

    if str(folder) not in sys.path:
        sys.path.insert(0, str(folder))

    mtime = script.stat().st_mtime
    code, has_run = _load_code(str(script), mtime)

    # preferred: the module defines run(...)
    if has_run:
        _load_module(str(script), mtime).run(up_file)
        return

    # fallback: top-level Streamlit script
    real_cfg, real_up = st.set_page_config, st.file_uploader

    def fake_uploader(label, type=None, *a, **kw):
        return up_file

    st.set_page_config = lambda *a, **kw: None       # already set by router
    st.file_uploader  = fake_uploader                 # feed the router's file
    try:
        exec(code, {"__name__": "__main__", "__file__": str(script),
                    "__builtins__": __builtins__})
    finally:
        st.set_page_config, st.file_uploader = real_cfg, real_up


# ============================================================
# Page
# ============================================================
st.set_page_config(page_title="PO -> Packing List", layout="wide", page_icon="📦")
theme.inject_theme()
theme.install_xlsx_polish()
theme.install_dataframe_safe()
theme.hero("PO to Packing List", "Upload a buyer PO (PDF, Excel or DOCX) and get a ready-to-ship packing list", icon="📦")

col_pdf, col_xlsx, col_docx, col_db = st.columns(4)

with col_pdf:
    pdf_up = st.file_uploader(
        "📄 Upload PO PDF",
        type=["pdf"],
        key="router_pdf",
        help="VF/Vans · GAP/Old Navy · Banana Republic · Kontoor/Lee\n"
             "· s.Oliver · Carter's/OshKosh · A&F · Tally Weijl · Justice",
    )
with col_xlsx:
    xlsx_up = st.file_uploader(
        "📊 Upload XL file",
        type=["xlsx", "xls"],
        key="router_xlsx",
        help="Next Retail / Clover · Haddad · MUJI",
    )
with col_docx:
    docx_up = st.file_uploader(
        "📝 Upload DOCX",
        type=["docx"],
        key="router_docx",
        help="JCPenney / Thereabouts PO",
    )
with col_db:
    final_files = final_db.uploader_box()          # 🟩 green box - stores final XL files in the database

final_db.render_panel(final_files)                 # preview + "Save to database" + stored files

with st.sidebar:
    st.markdown("### Buyer (optional override)")
    override = st.selectbox(
        "Pick a buyer if auto-detection fails",
        ["Auto-detect"] + [APPS[k][2] for k in APPS],
        key="router_override",
    )

    st.divider()
    st.markdown("### Supported buyers")
    for k, (_, _, label, ftype) in APPS.items():
        st.caption(f"**{label}** — {ftype.upper()}")


# ============================================================
# Which boxes are filled?
# ============================================================
filled = [x for x in (pdf_up, xlsx_up, docx_up) if x is not None]

# 1) More than one → warn
if len(filled) > 1:
    st.warning("⚠️ **You uploaded more than one file.** "
               "Please clear the other boxes and upload a single file at a time.")
    st.stop()

# 2) Nothing uploaded
if not filled:
    if override == "GAP (Old Navy)":
        st.divider()
        run_app("GAP", None)
    else:
        if not final_files:
            st.info("Upload **one** file — PDF, Excel or DOCX — to get started.")
            st.caption("PDF: VF/Vans · GAP/Old Navy · Banana Republic · "
                       "Kontoor/Lee · s.Oliver · Carter's/OshKosh · A&F · Tally Weijl · Justice")
            st.caption("XLSX: Next Retail / Clover · Haddad · MUJI")
            st.caption("DOCX: JCPenney / Thereabouts")
            st.caption("🟩 Green box: final completed Excel files → saved in the database")
    st.stop()

# ============================================================
# 3) PDF uploaded
# ============================================================
if pdf_up is not None:
    key = LABEL_TO_KEY.get(override)

    # if the user picked a non-PDF buyer (JCP or Next) → error
    if key and APPS[key][3] != "pdf":
        st.error(f"The **{APPS[key][2]}** module expects a "
                 f"`.{APPS[key][3]}` file, but you uploaded a PDF.")
        st.stop()

    if key is None:                          # auto-detect
        try:
            key = detect_buyer_pdf(pdf_up.getvalue())
        except Exception as e:
            st.error(f"Could not read the PDF: {e}")
            st.stop()

    if key is None:
        st.warning("Could not recognise the PDF buyer. "
                   "Pick the buyer in the sidebar and try again.")
        st.stop()

    st.success(
        f"Detected buyer: **{APPS[key][2]}**" if override == "Auto-detect"
        else f"Buyer set manually: **{APPS[key][2]}**"
    )
    st.divider()
    run_app(key, pdf_up)
    st.stop()

# ============================================================
# 4) XLSX uploaded
# ============================================================
if xlsx_up is not None:
    key = LABEL_TO_KEY.get(override)

    if key and APPS[key][3] not in ("xlsx", "xls"):
        st.error(f"The **{APPS[key][2]}** module expects a "
                 f"`.{APPS[key][3]}` file, but you uploaded an Excel file.")
        st.stop()

    if key is None:
        try:
            key = detect_buyer_xlsx(xlsx_up.getvalue(), getattr(xlsx_up, "name", ""))
        except Exception:
            key = "NX"

    st.success(f"Detected buyer: **{APPS[key][2]}**"
               if override == "Auto-detect"
               else f"Buyer set manually: **{APPS[key][2]}**")
    st.divider()
    run_app(key, xlsx_up)
    st.stop()

# ============================================================
# 5) DOCX uploaded
# ============================================================
if docx_up is not None:
    key = LABEL_TO_KEY.get(override)

    if key and APPS[key][3] != "docx":
        st.error(f"The **{APPS[key][2]}** module expects a "
                 f"`.{APPS[key][3]}` file, but you uploaded a DOCX.")
        st.stop()

    if key is None:                          # auto-detect
        try:
            key = detect_buyer_docx(docx_up.getvalue())
        except Exception as e:
            st.error(f"Could not read the DOCX: {e}")
            st.stop()

    if key is None:
        st.warning("Could not recognise the DOCX buyer. "
                   "Pick the buyer in the sidebar and try again.")
        st.stop()

    st.success(
        f"Detected buyer: **{APPS[key][2]}**" if override == "Auto-detect"
        else f"Buyer set manually: **{APPS[key][2]}**"
    )
    st.divider()
    run_app(key, docx_up)
    st.stop()
