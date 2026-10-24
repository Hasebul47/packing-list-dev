



# """
# Unified PO -> Packing List router
# =================================
# Three upload boxes build a packing list (only one at a time):

#   PDF   -> VF/Vans, GAP/Old Navy, Banana Republic, Kontoor/Lee, s.Oliver,
#            Carter's/OshKosh, Abercrombie & Fitch, Tally Weijl, Justice
#   XLSX  -> Next Retail / Clover, Haddad, MUJI
#   DOCX  -> JCPenney / Thereabouts

# The green box (final_db) stores FINAL completed packing-list Excel files in the
# database. It is independent of the three boxes above.

# More than one box filled -> warning, nothing runs.
# Nothing filled           -> welcome card.

# Each buyer module exposes run(file) OR is a top-level Streamlit script;
# the router handles both.

# Run:  streamlit run app.py
# """
# import importlib
# import io
# import re
# import runpy
# import sys
# import zipfile
# from pathlib import Path

# import pdfplumber
# import streamlit as st

# BASE = Path(__file__).resolve().parent
# if str(BASE) not in sys.path:
#     sys.path.insert(0, str(BASE))

# import theme                      # shared palette / smooth-UI css / xlsx polish
# import final_db                   # green box: completed XL files -> database

# # ============================================================
# # Buyer registry
# # key -> (folder, filename, human label, file type)
# # ============================================================
# APPS = {
#     "VF":  ("VF Vans",    "VF_Packing_V3_2_app.py", "VF (Vans)",                "pdf"),
#     "GAP": ("Gap",        "unified_packing.py",     "GAP (Old Navy)",           "pdf"),
#     "BR":  ("br",         "br_packing.py",          "Banana Republic",          "pdf"),
#     "KT":  ("kontor",     "kontoor_app.py",         "Kontoor (Lee / Wrangler)", "pdf"),
#     "SO":  ("s_oliver",   "s_oliver_app.py",        "s.Oliver",                 "pdf"),
#     "CT":  ("carters",    "carters_app.py",         "Carter's / OshKosh",       "pdf"),
#     "ANF": ("ABERCROMBIE", "anf_app.py",            "Abercrombie & Fitch",      "pdf"),
#     "TW":  ("Tally_Weijl", "tw_app.py",             "Tally Weijl",              "pdf"),
#     "JU":  ("justice",    "justice_app.py",         "Justice (SPA / Adjmi)",    "pdf"),
#     "NX":  ("next",       "next_buyer_app.py",      "Next Retail / Clover",     "xlsx"),
#     "HD":  ("haddad",     "haddad_app.py",          "Haddad (Levi's Kids)",     "xlsx"),
#     "MJ":  ("MUJI",       "app.py",                 "MUJI",                     "xlsx"),
#     "JC":  ("jcp",        "jcp_app.py",             "JCPenney / Thereabouts",   "docx"),
# }

# # label -> key  (for the dropdown override)
# LABEL_TO_KEY = {v[2]: k for k, v in APPS.items()}

# # file type -> (icon, card title, short buyer line shown on the card)
# TYPE_META = {
#     "pdf":  ("📄", "PO PDF", "VF · GAP · Banana Republic · Kontoor · s.Oliver · Carter's · A&F · Tally Weijl · Justice"),
#     "xlsx": ("📊", "Excel",  "Next / Clover · Haddad · MUJI"),
#     "docx": ("📝", "DOCX",   "JCPenney / Thereabouts"),
# }


# # ============================================================
# # Locate the module script (very tolerant)
# # ============================================================
# def _find_script(folder_name: str, filename: str):
#     folder = BASE / folder_name
#     stem = Path(filename).stem

#     for c in (folder / filename,
#               BASE / filename,
#               BASE.parent / folder_name / filename):
#         if c.is_file():
#             return c, c.parent

#     if folder.is_dir():
#         for c in folder.rglob(f"{stem}.py"):
#             if c.is_file():
#                 return c, c.parent

#     for c in BASE.rglob(filename):
#         if c.is_file():
#             return c, c.parent

#     return None, None


# # ============================================================
# # Buyer detection — PDF
# # ============================================================
# @st.cache_data(show_spinner=False, max_entries=32)
# def detect_buyer_pdf(pdf_bytes: bytes):
#     with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
#         t = (pdf.pages[0].extract_text() or "").upper()
#         t2 = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:2]).upper()
#     squashed = re.sub(r"[^A-Z0-9]", "", t2)

#     if ("ABERCROMBIE" in t or "FITCH PATH" in t or "ABFICO" in squashed or "ANFCORP" in squashed
#             or ("HOLLISTERCO" in squashed and "POSKUSUMMARY" in squashed)):
#         return "ANF"
#     if ("TALLYWEIJL" in squashed or "TSWL" in squashed or "VIADUKTSTRASSE" in squashed
#             or ("ORDERFORM" in squashed and "SUPPLIERSTALLY" in squashed)):
#         return "TW"
#     if ("SPORTSPRODUCTSOFAMERICA" in squashed or "ADJMIAPPAREL" in squashed):
#         return "JU"
#     if ("CARTERS" in t or "WILLIAM CARTER" in t or "OSHKOSH" in t
#             or "E- PLM" in t or "E-PLM" in t):
#         return "CT"
#     if "S.OLIVER" in t or "S.OLIVER BERND FREIER" in t:
#         return "SO"
#     if "KONTOOR" in t:
#         return "KT"
#     if "VF OUTDOOR" in t or ("VANS" in t and "BUYER CODE" in t):
#         return "VF"
#     if ("BANANA REPUBLIC" in t or "GAP (CANADA)" in t
#             or "GAP(CANADA)" in t or "PURCHASER" in t and "GAP (CANADA)" in t):
#         return "BR"
#     if "OLD NAVY" in t or "GAP INTL" in t or "DESTINATION PURCHASE ORDER" in t:
#         return "GAP"
#     return None


# # ============================================================
# # Buyer detection — XLSX (Next, Haddad, MUJI)
# # ============================================================
# @st.cache_data(show_spinner=False, max_entries=32)
# def detect_buyer_xlsx(xlsx_bytes: bytes, file_name: str = ""):
#     fn = (file_name or "").upper()
#     if "MUJI" in fn or "発注" in fn:
#         return "MJ"
#     if "HADDAD" in fn:
#         return "HD"
#     if "NEXT" in fn or "CLOVER" in fn:
#         return "NX"

#     try:
#         import openpyxl
#         wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
#         for sn in wb.sheetnames:
#             sn_u = sn.upper()
#             if "MUJI" in sn_u or "発注" in sn_u:
#                 return "MJ"
#             if "STATUS" in sn_u or "HADDAD" in sn_u:
#                 return "HD"
#         ws = wb.active or wb[wb.sheetnames[0]]
#         for i, row in enumerate(ws.iter_rows(values_only=True)):
#             if i > 30:
#                 break
#             row_str = " ".join(str(c or "") for c in row).upper()
#             if "REF#" in row_str and "PO#" in row_str:
#                 return "HD"
#             if "仕向地" in row_str or "発注一覧" in row_str:
#                 return "MJ"
#             if "ITEM NUMBER" in row_str or "CONTRACT NO" in row_str:
#                 return "NX"
#     except Exception:
#         pass

#     try:
#         import xlrd
#         book = xlrd.open_workbook(file_contents=xlsx_bytes)
#         sheet = book.sheet_by_index(0)
#         for r in range(min(30, sheet.nrows)):
#             row_vals = " ".join(str(sheet.cell_value(r, c)) for c in range(sheet.ncols)).upper()
#             if "REF#" in row_vals and "PO#" in row_vals:
#                 return "HD"
#             if "ITEM NUMBER" in row_vals or "CONTRACT NO" in row_vals:
#                 return "NX"
#     except Exception:
#         pass

#     return "NX"


# # ============================================================
# # Buyer detection — DOCX (JCPenney)
# # ============================================================
# @st.cache_data(show_spinner=False, max_entries=32)
# def detect_buyer_docx(docx_bytes: bytes):
#     try:
#         with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
#             xml = z.read("word/document.xml").decode("utf-8", errors="ignore").upper()
#     except Exception:
#         return None
#     if ("THEREABOUTS" in xml or "JCPPC" in xml
#             or "JCP PURCHASING" in xml or "JCPENNEY" in xml):
#         return "JC"
#     return None


# # ============================================================
# # Dispatch
# # ============================================================
# @st.cache_resource(show_spinner=False)
# def _load_module(path_str: str, mtime: float):
#     path = Path(path_str)
#     sys.modules.pop(path.stem, None)
#     return importlib.import_module(path.stem)


# @st.cache_resource(show_spinner=False)
# def _load_code(path_str: str, mtime: float):
#     src = Path(path_str).read_text(encoding="utf-8")
#     return compile(src, path_str, "exec"), ("def run(" in src)


# def run_app(key: str, up_file):
#     folder_name, filename, label, _ = APPS[key]
#     script, folder = _find_script(folder_name, filename)

#     if script is None:
#         st.error(
#             f"Could not find **{filename}** anywhere under `{BASE}`.\n\n"
#             f"Looked in:\n"
#             f"- `{BASE / folder_name / filename}`\n"
#             f"- `{BASE / filename}`\n"
#             f"- anything reachable under `{BASE}`"
#         )
#         st.stop()

#     if st.session_state.get("_active_buyer") != key:
#         for k in list(st.session_state.keys()):
#             if not k.startswith(("_active_buyer", "router_")):
#                 st.session_state.pop(k, None)
#         st.session_state["_active_buyer"] = key

#     if str(folder) not in sys.path:
#         sys.path.insert(0, str(folder))

#     mtime = script.stat().st_mtime
#     code, has_run = _load_code(str(script), mtime)

#     if has_run:
#         _load_module(str(script), mtime).run(up_file)
#         return

#     real_cfg, real_up = st.set_page_config, st.file_uploader

#     def fake_uploader(label, type=None, *a, **kw):
#         return up_file

#     st.set_page_config = lambda *a, **kw: None
#     st.file_uploader  = fake_uploader
#     try:
#         exec(code, {"__name__": "__main__", "__file__": str(script),
#                     "__builtins__": __builtins__})
#     finally:
#         st.set_page_config, st.file_uploader = real_cfg, real_up


# # ============================================================
# # UI helpers
# # ============================================================
# UI_CSS = """
# <style>
# /* ---------- tokens (follow the active Streamlit light/dark theme) ---------- */
# :root {
#     --po-line:   color-mix(in srgb, var(--text-color, #1f2937) 14%, transparent);
#     --po-line-2: color-mix(in srgb, var(--text-color, #1f2937) 8%, transparent);
#     --po-soft:   color-mix(in srgb, var(--text-color, #1f2937) 4%, transparent);
#     --po-accent: var(--primary-color, #ff4b4b);
#     --po-radius: 12px;
#     --po-field-h: 2.6rem;
# }

# /* ---------- page rhythm ---------- */
# .block-container { padding-top: 2rem; padding-bottom: 3rem; }
# div[data-testid="stAlert"] { border-radius: var(--po-radius); }

# /* ---------- section titles ---------- */
# .po-section {
#     display: flex; align-items: baseline; justify-content: space-between;
#     gap: 1rem; margin: 1.6rem 0 0.7rem;
#     padding-bottom: 0.45rem; border-bottom: 1px solid var(--po-line-2);
# }
# .po-section b   { font-size: 1.02rem; font-weight: 650; letter-spacing: -0.005em; }
# .po-section span { font-size: 0.82rem; opacity: 0.6; }

# /* ---------- upload cards ---------- */
# .st-key-po_upload [data-testid="stVerticalBlockBorderWrapper"] {
#     border-radius: var(--po-radius);
#     border-color: var(--po-line);
#     padding: 0.35rem 0.4rem 0.2rem;
# }
# .st-key-po_upload [data-testid="stHorizontalBlock"] { align-items: stretch; }
# .st-key-po_upload :is([data-testid="stColumn"], [data-testid="column"]) > div,
# .st-key-po_upload :is([data-testid="stColumn"], [data-testid="column"]) [data-testid="stVerticalBlockBorderWrapper"] { height: 100%; }

# .po-card-head { display: flex; align-items: center; gap: 0.7rem; min-height: 3.1rem; margin-bottom: 0.2rem; }
# .po-ico {
#     flex: 0 0 2.3rem; height: 2.3rem; display: grid; place-items: center;
#     border-radius: 10px; font-size: 1.15rem; background: var(--po-soft);
# }
# .po-head-text { flex: 1 1 auto; min-width: 0; }
# .po-title { font-weight: 650; font-size: 0.98rem; line-height: 1.2; }
# .po-sub   { font-size: 0.76rem; opacity: 0.6; line-height: 1.35; margin-top: 0.1rem; }
# .po-pill {
#     flex: 0 0 auto; font-size: 0.7rem; font-weight: 650; padding: 0.15rem 0.55rem;
#     border-radius: 999px; letter-spacing: 0.02em;
# }
# .po-pill.pdf  { background: rgba(239, 68, 68, .14); color: #ef4444; }
# .po-pill.xlsx { background: rgba(34, 197, 94, .15); color: #16a34a; }
# .po-pill.docx { background: rgba(59, 130, 246, .15); color: #3b82f6; }

# /* ---------- uploaders inside the three cards ---------- */
# .st-key-po_upload div[data-testid="stFileUploader"] { margin: 0; padding: 0; width: 100%; }
# .st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) {
#     display: flex; flex-direction: column; align-items: stretch; justify-content: center;
#     gap: 0.6rem; box-sizing: border-box; width: 100%; max-width: 100%; overflow: hidden;
#     min-height: 5.6rem; padding: 0.8rem 0.9rem;
#     border: 1.5px dashed var(--po-line); border-radius: 10px;
#     background: transparent; transition: border-color .15s, background .15s;
# }
# .st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) > div {
#     display: flex; align-items: center; justify-content: center; gap: 0.5rem;
#     min-width: 0; max-width: 100%; text-align: center;
# }
# .st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]):hover {
#     border-color: var(--po-accent);
#     background: color-mix(in srgb, var(--po-accent) 5%, transparent);
# }
# .st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) span { font-size: 0.82rem; font-weight: 500; }
# .st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) small { display: none; }
# .st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) svg { width: 20px; height: 20px; }
# .st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) button {
#     margin: 0; width: 100%; white-space: nowrap;
#     height: 2.1rem; min-height: 2.1rem; padding: 0 0.9rem;
#     font-size: 0.8rem; line-height: 1; border-radius: 8px;
# }
# .st-key-po_upload [data-testid="stFileUploaderFile"] { font-size: 0.82rem; }

# /* ---------- green database box: keep content inside the box, equal padding ---------- */
# .st-key-po_db, .st-key-po_db * { box-sizing: border-box !important; }
# .st-key-po_db [data-testid="stElementContainer"],
# .st-key-po_db .element-container { width: 100% !important; max-width: 100% !important; }
# /* the uploader (green box): fixed, even padding on all four sides */
# .st-key-po_db div[data-testid="stFileUploader"] {
#     width: 100% !important; max-width: 100% !important; min-width: 0 !important;
#     margin: 0 !important; left: auto !important; right: auto !important;
#     transform: none !important; padding: 0.8rem 1rem !important; overflow: hidden;
# }
# /* everything inside it: no extra offsets, never wider than the box */
# .st-key-po_db div[data-testid="stFileUploader"] > *,
# .st-key-po_db div[data-testid="stFileUploader"] section,
# .st-key-po_db :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]),
# .st-key-po_db [data-testid="stWidgetLabel"] {
#     width: 100% !important; max-width: 100% !important; min-width: 0 !important;
#     margin-left: 0 !important; margin-right: 0 !important;
#     left: auto !important; right: auto !important; transform: none !important;
# }
# .st-key-po_db div[data-testid="stFileUploader"] > * { padding-left: 0 !important; padding-right: 0 !important; }

# /* ---------- search panel background ---------- */
# .st-key-po_search {
#     background: #e0f2fe;
#     border: 1px solid #bae6fd;
#     border-radius: var(--po-radius);
#     padding: 0.8rem 0.9rem;
# }
# .st-key-po_search div[data-baseweb="input"],
# .st-key-po_search div[data-baseweb="base-input"] { background: var(--background-color, #fff); }

# /* ---------- search row: input and button share one height ---------- */
# .st-key-po_search div[data-baseweb="input"],
# .st-key-po_search div[data-baseweb="base-input"] { min-height: var(--po-field-h); border-radius: 10px; }
# .st-key-po_search input { height: var(--po-field-h); font-size: 0.92rem; }
# .st-key-po_search div.stButton > button {
#     height: var(--po-field-h); min-height: var(--po-field-h);
#     border-radius: 10px; font-size: 0.9rem; font-weight: 600; padding: 0 1rem; margin: 0;
# }
# .st-key-po_search [data-testid="stHorizontalBlock"] { align-items: center; }
# .st-key-po_search div[data-testid="stTextInput"],
# .st-key-po_search div.stButton { margin: 0; }

# /* ---------- sidebar ---------- */
# section[data-testid="stSidebar"] .po-group { margin: 0.9rem 0 0.2rem; font-size: 0.85rem; font-weight: 650; }
# section[data-testid="stSidebar"] .po-list { margin: 0 0 0.3rem; padding-left: 1.1rem; }
# section[data-testid="stSidebar"] .po-list li { font-size: 0.85rem; line-height: 1.35; margin: 0.18rem 0; opacity: 0.85; }
# </style>
# """


# def _scope(key: str, **kwargs):
#     """A keyed container (CSS hook `.st-key-<key>`); plain container on old Streamlit."""
#     try:
#         return st.container(key=key, **kwargs)
#     except TypeError:
#         return st.container(**kwargs)


# def _section(title: str, hint: str = ""):
#     hint_html = f"<span>{hint}</span>" if hint else ""
#     st.markdown(f'<div class="po-section"><b>{title}</b>{hint_html}</div>', unsafe_allow_html=True)


# def _card_head(ftype: str):
#     icon, title, sub = TYPE_META[ftype]
#     st.markdown(
#         f'<div class="po-card-head">'
#         f'<div class="po-ico">{icon}</div>'
#         f'<div class="po-head-text"><div class="po-title">{title}</div><div class="po-sub">{sub}</div></div>'
#         f'<span class="po-pill {ftype}">{ftype.upper()}</span>'
#         f'</div>',
#         unsafe_allow_html=True,
#     )


# def _buyer_list(ftype: str) -> str:
#     items = [v[2] for v in APPS.values() if v[3] == ftype]
#     return '<ul class="po-list">' + "".join(f"<li>{i}</li>" for i in items) + "</ul>"


# def _buyer_banner(key: str):
#     st.success(f"Detected buyer: **{APPS[key][2]}**" if override == "Auto-detect"
#                else f"Buyer set manually: **{APPS[key][2]}**")


# # ============================================================
# # Page & Layout
# # ============================================================
# st.set_page_config(page_title="PO -> Packing List", layout="wide", page_icon="📦")
# theme.inject_theme()
# theme.install_xlsx_polish()
# theme.install_dataframe_safe()
# st.markdown(UI_CSS, unsafe_allow_html=True)

# theme.hero(
#     "PO to Packing List",
#     "Upload a buyer PO (PDF, Excel or DOCX) and get a ready-to-ship packing list",
#     icon="📦",
# )


# # ---- 2. Search (input + button, same height, same baseline) ----
# _section("Search")
# with _scope("po_search"):
#     sc_input, sc_btn = st.columns([8, 2], gap="small")
#     with sc_input:
#         st.text_input(
#             "Search",
#             placeholder="Search packing lists, POs, styles, buyers...",
#             label_visibility="collapsed",
#             key="router_search",
#         )
#     with sc_btn:
#         st.button("Search", key="router_search_btn", use_container_width=True)


# # ---- 1. Upload cards (equal width, equal height) ----
# _section("Upload a PO", "One file at a time")
# with _scope("po_upload"):
#     col_pdf, col_xlsx, col_docx = st.columns(3, gap="medium")

#     with col_pdf:
#         with st.container(border=True):
#             _card_head("pdf")
#             pdf_up = st.file_uploader(
#                 "PO PDF", type=["pdf"], key="router_pdf",
#                 label_visibility="collapsed",
#             )
#     with col_xlsx:
#         with st.container(border=True):
#             _card_head("xlsx")
#             xlsx_up = st.file_uploader(
#                 "Excel", type=["xlsx", "xls"], key="router_xlsx",
#                 label_visibility="collapsed",
#             )
#     with col_docx:
#         with st.container(border=True):
#             _card_head("docx")
#             docx_up = st.file_uploader(
#                 "DOCX", type=["docx"], key="router_docx",
#                 label_visibility="collapsed",
#             )


# # ---- 3. Green database box + preview panel ----
# _section("Final packing-list database", "Save completed Excel files")
# with _scope("po_db"):
#     final_files = final_db.uploader_box()
#     final_db.render_panel(final_files)

# # ---- Sidebar ----
# with st.sidebar:
#     st.markdown("### Buyer")
#     override = st.selectbox(
#         "Pick a buyer if auto-detection fails",
#         ["Auto-detect"] + [APPS[k][2] for k in APPS],
#         key="router_override",
#     )
#     st.divider()
#     st.markdown("### Supported buyers")
#     for ft in ("pdf", "xlsx", "docx"):
#         st.markdown(
#             f'<div class="po-group">{TYPE_META[ft][1]}</div>'
#             f'{_buyer_list(ft)}',
#             unsafe_allow_html=True,
#         )


# # ============================================================
# # Logic: which boxes are filled?
# # ============================================================
# filled = [x for x in (pdf_up, xlsx_up, docx_up) if x is not None]

# if len(filled) > 1:
#     st.warning("⚠️ **You uploaded more than one file.** "
#                "Please clear the other boxes and upload a single file at a time.")
#     st.stop()

# # ---- Nothing uploaded ----
# if not filled:
#     if override == "GAP (Old Navy)":
#         st.divider()
#         run_app("GAP", None)
#     st.stop()

# # ============================================================
# # Routing logic (unchanged)
# # ============================================================
# if pdf_up is not None:
#     key = LABEL_TO_KEY.get(override)
#     if key and APPS[key][3] != "pdf":
#         st.error(f"The **{APPS[key][2]}** module expects a `.{APPS[key][3]}` file, but you uploaded a PDF.")
#         st.stop()
#     if key is None:
#         try:
#             key = detect_buyer_pdf(pdf_up.getvalue())
#         except Exception as e:
#             st.error(f"Could not read the PDF: {e}")
#             st.stop()
#     if key is None:
#         st.warning("Could not recognise the PDF buyer. Pick the buyer in the sidebar and try again.")
#         st.stop()
#     _buyer_banner(key)
#     st.divider()
#     run_app(key, pdf_up)
#     st.stop()

# if xlsx_up is not None:
#     key = LABEL_TO_KEY.get(override)
#     if key and APPS[key][3] not in ("xlsx", "xls"):
#         st.error(f"The **{APPS[key][2]}** module expects a `.{APPS[key][3]}` file, but you uploaded an Excel file.")
#         st.stop()
#     if key is None:
#         try:
#             key = detect_buyer_xlsx(xlsx_up.getvalue(), getattr(xlsx_up, "name", ""))
#         except Exception:
#             key = "NX"
#     _buyer_banner(key)
#     st.divider()
#     run_app(key, xlsx_up)
#     st.stop()

# if docx_up is not None:
#     key = LABEL_TO_KEY.get(override)
#     if key and APPS[key][3] != "docx":
#         st.error(f"The **{APPS[key][2]}** module expects a `.{APPS[key][3]}` file, but you uploaded a DOCX.")
#         st.stop()
#     if key is None:
#         try:
#             key = detect_buyer_docx(docx_up.getvalue())
#         except Exception as e:
#             st.error(f"Could not read the DOCX: {e}")
#             st.stop()
#     if key is None:
#         st.warning("Could not recognise the DOCX buyer. Pick the buyer in the sidebar and try again.")
#         st.stop()
#     _buyer_banner(key)
#     st.divider()
#     run_app(key, docx_up)
#     st.stop()



"""
Unified PO -> Packing List router
=================================
Three upload boxes build a packing list (only one at a time):

  PDF   -> VF/Vans, GAP/Old Navy, Banana Republic, Kontoor/Lee, s.Oliver,
           Carter's/OshKosh, Abercrombie & Fitch, Tally Weijl, Justice
  XLSX  -> Next Retail / Clover, Haddad, MUJI
  DOCX  -> JCPenney / Thereabouts

The green box (final_db) stores FINAL completed packing-list Excel files in the
database. It is independent of the three boxes above.

More than one box filled -> warning, nothing runs.
Nothing filled           -> welcome card.

Each buyer module exposes run(file) OR is a top-level Streamlit script;
the router handles both.

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
import final_db                   # green box: completed XL files -> database

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

# file type -> (icon, card title, short buyer line shown on the card)
TYPE_META = {
    "pdf":  ("📄", "PO PDF", "VF · GAP · Banana Republic · Kontoor · s.Oliver · Carter's · A&F · Tally Weijl · Justice"),
    "xlsx": ("📊", "Excel",  "Next / Clover · Haddad · MUJI"),
    "docx": ("📝", "DOCX",   "JCPenney / Thereabouts"),
}


# ============================================================
# Locate the module script (very tolerant)
# ============================================================
def _find_script(folder_name: str, filename: str):
    folder = BASE / folder_name
    stem = Path(filename).stem

    for c in (folder / filename,
              BASE / filename,
              BASE.parent / folder_name / filename):
        if c.is_file():
            return c, c.parent

    if folder.is_dir():
        for c in folder.rglob(f"{stem}.py"):
            if c.is_file():
                return c, c.parent

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
        t2 = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:2]).upper()
    squashed = re.sub(r"[^A-Z0-9]", "", t2)

    if ("ABERCROMBIE" in t or "FITCH PATH" in t or "ABFICO" in squashed or "ANFCORP" in squashed
            or ("HOLLISTERCO" in squashed and "POSKUSUMMARY" in squashed)):
        return "ANF"
    if ("TALLYWEIJL" in squashed or "TSWL" in squashed or "VIADUKTSTRASSE" in squashed
            or ("ORDERFORM" in squashed and "SUPPLIERSTALLY" in squashed)):
        return "TW"
    if ("SPORTSPRODUCTSOFAMERICA" in squashed or "ADJMIAPPAREL" in squashed):
        return "JU"
    if ("CARTERS" in t or "WILLIAM CARTER" in t or "OSHKOSH" in t
            or "E- PLM" in t or "E-PLM" in t):
        return "CT"
    if "S.OLIVER" in t or "S.OLIVER BERND FREIER" in t:
        return "SO"
    if "KONTOOR" in t:
        return "KT"
    if "VF OUTDOOR" in t or ("VANS" in t and "BUYER CODE" in t):
        return "VF"
    if ("BANANA REPUBLIC" in t or "GAP (CANADA)" in t
            or "GAP(CANADA)" in t or "PURCHASER" in t and "GAP (CANADA)" in t):
        return "BR"
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
# Dispatch
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

    if st.session_state.get("_active_buyer") != key:
        for k in list(st.session_state.keys()):
            if not k.startswith(("_active_buyer", "router_")):
                st.session_state.pop(k, None)
        st.session_state["_active_buyer"] = key

    if str(folder) not in sys.path:
        sys.path.insert(0, str(folder))

    mtime = script.stat().st_mtime
    code, has_run = _load_code(str(script), mtime)

    if has_run:
        _load_module(str(script), mtime).run(up_file)
        return

    real_cfg, real_up = st.set_page_config, st.file_uploader

    def fake_uploader(label, type=None, *a, **kw):
        return up_file

    st.set_page_config = lambda *a, **kw: None
    st.file_uploader  = fake_uploader
    try:
        exec(code, {"__name__": "__main__", "__file__": str(script),
                    "__builtins__": __builtins__})
    finally:
        st.set_page_config, st.file_uploader = real_cfg, real_up


# ============================================================
# UI helpers
# ============================================================
UI_CSS = """
<style>
/* ---------- tokens (follow the active Streamlit light/dark theme) ---------- */
:root {
    --po-line:   color-mix(in srgb, var(--text-color, #1f2937) 14%, transparent);
    --po-line-2: color-mix(in srgb, var(--text-color, #1f2937) 8%, transparent);
    --po-soft:   color-mix(in srgb, var(--text-color, #1f2937) 4%, transparent);
    --po-accent: var(--primary-color, #ff4b4b);
    --po-radius: 12px;
    --po-field-h: 2.6rem;
}

/* ---------- page rhythm ---------- */
.block-container { padding-top: 2rem; padding-bottom: 3rem; }
div[data-testid="stAlert"] { border-radius: var(--po-radius); }

/* ---------- section titles ---------- */
.po-section {
    display: flex; align-items: baseline; justify-content: space-between;
    gap: 1rem; margin: 1.6rem 0 0.7rem;
    padding-bottom: 0.45rem; border-bottom: 1px solid var(--po-line-2);
}
.po-section b   { font-size: 1.02rem; font-weight: 650; letter-spacing: -0.005em; }
.po-section span { font-size: 0.82rem; opacity: 0.6; }

/* ---------- upload cards ---------- */
.st-key-po_upload [data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: var(--po-radius);
    border-color: var(--po-line);
    padding: 0.35rem 0.4rem 0.2rem;
}
.st-key-po_upload [data-testid="stHorizontalBlock"] { align-items: stretch; }
.st-key-po_upload :is([data-testid="stColumn"], [data-testid="column"]) > div,
.st-key-po_upload :is([data-testid="stColumn"], [data-testid="column"]) [data-testid="stVerticalBlockBorderWrapper"] { height: 100%; }

.po-card-head { display: flex; align-items: center; gap: 0.7rem; min-height: 3.1rem; margin-bottom: 0.2rem; }
.po-ico {
    flex: 0 0 2.3rem; height: 2.3rem; display: grid; place-items: center;
    border-radius: 10px; font-size: 1.15rem; background: var(--po-soft);
}
.po-head-text { flex: 1 1 auto; min-width: 0; }
.po-title { font-weight: 650; font-size: 0.98rem; line-height: 1.2; }
.po-sub   { font-size: 0.76rem; opacity: 0.6; line-height: 1.35; margin-top: 0.1rem; }
.po-pill {
    flex: 0 0 auto; font-size: 0.7rem; font-weight: 650; padding: 0.15rem 0.55rem;
    border-radius: 999px; letter-spacing: 0.02em;
}
.po-pill.pdf  { background: rgba(239, 68, 68, .14); color: #ef4444; }
.po-pill.xlsx { background: rgba(34, 197, 94, .15); color: #16a34a; }
.po-pill.docx { background: rgba(59, 130, 246, .15); color: #3b82f6; }

/* ---------- uploaders inside the three cards ---------- */
.st-key-po_upload div[data-testid="stFileUploader"] { margin: 0; padding: 0; width: 100%; }
.st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) {
    display: flex; flex-direction: column; align-items: stretch; justify-content: center;
    gap: 0.6rem; box-sizing: border-box; width: 100%; max-width: 100%; overflow: hidden;
    min-height: 5.6rem; padding: 0.8rem 0.9rem;
    border: 1.5px dashed var(--po-line); border-radius: 10px;
    background: transparent; transition: border-color .15s, background .15s;
}
.st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) > div {
    display: flex; align-items: center; justify-content: center; gap: 0.5rem;
    min-width: 0; max-width: 100%; text-align: center;
}
.st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]):hover {
    border-color: var(--po-accent);
    background: color-mix(in srgb, var(--po-accent) 5%, transparent);
}
.st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) span { font-size: 0.82rem; font-weight: 500; }
.st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) small { display: none; }
.st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) svg { width: 20px; height: 20px; }
.st-key-po_upload :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]) button {
    margin: 0; width: 100%; white-space: nowrap;
    height: 2.1rem; min-height: 2.1rem; padding: 0 0.9rem;
    font-size: 0.8rem; line-height: 1; border-radius: 8px;
}
.st-key-po_upload [data-testid="stFileUploaderFile"] { font-size: 0.82rem; }

/* ---------- green database box: keep content inside the box, equal padding ---------- */
.st-key-po_db, .st-key-po_db * { box-sizing: border-box !important; }
.st-key-po_db [data-testid="stElementContainer"],
.st-key-po_db .element-container { width: 100% !important; max-width: 100% !important; }
/* the uploader (green box): fixed, even padding on all four sides */
.st-key-po_db div[data-testid="stFileUploader"] {
    width: 100% !important; max-width: 100% !important; min-width: 0 !important;
    margin: 0 !important; left: auto !important; right: auto !important;
    transform: none !important; padding: 0.8rem 1rem !important; overflow: hidden;
}
/* everything inside it: no extra offsets, never wider than the box */
.st-key-po_db div[data-testid="stFileUploader"] > *,
.st-key-po_db div[data-testid="stFileUploader"] section,
.st-key-po_db :is([data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"]),
.st-key-po_db [data-testid="stWidgetLabel"] {
    width: 100% !important; max-width: 100% !important; min-width: 0 !important;
    margin-left: 0 !important; margin-right: 0 !important;
    left: auto !important; right: auto !important; transform: none !important;
}
.st-key-po_db div[data-testid="stFileUploader"] > * { padding-left: 0 !important; padding-right: 0 !important; }

/* ---------- search panel background ---------- */
.st-key-po_search {
    background: #e0f2fe;
    border: 1px solid #bae6fd;
    border-radius: var(--po-radius);
    padding: 0.8rem 0.9rem;
}
.st-key-po_search div[data-baseweb="input"],
.st-key-po_search div[data-baseweb="base-input"] { background: var(--background-color, #fff); }

/* ---------- search row: input and button share one height ---------- */
.st-key-po_search div[data-baseweb="input"],
.st-key-po_search div[data-baseweb="base-input"] { min-height: var(--po-field-h); border-radius: 10px; }
.st-key-po_search input { height: var(--po-field-h); font-size: 0.92rem; }
.st-key-po_search div.stButton > button {
    height: var(--po-field-h); min-height: var(--po-field-h);
    border-radius: 10px; font-size: 0.9rem; font-weight: 600; padding: 0 1rem; margin: 0;
}
.st-key-po_search [data-testid="stHorizontalBlock"] { align-items: center; }
.st-key-po_search div[data-testid="stTextInput"],
.st-key-po_search div.stButton { margin: 0; }

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"] .po-group { margin: 0.9rem 0 0.2rem; font-size: 0.85rem; font-weight: 650; }
section[data-testid="stSidebar"] .po-list { margin: 0 0 0.3rem; padding-left: 1.1rem; }
section[data-testid="stSidebar"] .po-list li { font-size: 0.85rem; line-height: 1.35; margin: 0.18rem 0; opacity: 0.85; }
</style>
"""


def _scope(key: str, **kwargs):
    """A keyed container (CSS hook `.st-key-<key>`); plain container on old Streamlit."""
    try:
        return st.container(key=key, **kwargs)
    except TypeError:
        return st.container(**kwargs)


def _section(title: str, hint: str = ""):
    hint_html = f"<span>{hint}</span>" if hint else ""
    st.markdown(f'<div class="po-section"><b>{title}</b>{hint_html}</div>', unsafe_allow_html=True)


def _card_head(ftype: str):
    icon, title, sub = TYPE_META[ftype]
    st.markdown(
        f'<div class="po-card-head">'
        f'<div class="po-ico">{icon}</div>'
        f'<div class="po-head-text"><div class="po-title">{title}</div><div class="po-sub">{sub}</div></div>'
        f'<span class="po-pill {ftype}">{ftype.upper()}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )


def _buyer_list(ftype: str) -> str:
    items = [v[2] for v in APPS.values() if v[3] == ftype]
    return '<ul class="po-list">' + "".join(f"<li>{i}</li>" for i in items) + "</ul>"


def _buyer_banner(key: str):
    st.success(f"Detected buyer: **{APPS[key][2]}**" if override == "Auto-detect"
               else f"Buyer set manually: **{APPS[key][2]}**")


# ============================================================
# Page & Layout
# ============================================================
st.set_page_config(page_title="PO -> Packing List", layout="wide", page_icon="📦",
                   initial_sidebar_state="collapsed")
theme.inject_theme()
theme.install_xlsx_polish()
theme.install_dataframe_safe()
st.markdown(UI_CSS, unsafe_allow_html=True)

theme.hero(
    "PO to Packing List",
    "Upload a buyer PO (PDF, Excel or DOCX) and get a ready-to-ship packing list",
    icon="📦",
)

# ---- 1. Upload cards (equal width, equal height) ----
_section("Upload a PO", "One file at a time")
with _scope("po_upload"):
    col_pdf, col_xlsx, col_docx = st.columns(3, gap="medium")

    with col_pdf:
        with st.container(border=True):
            _card_head("pdf")
            pdf_up = st.file_uploader(
                "PO PDF", type=["pdf"], key="router_pdf",
                label_visibility="collapsed",
            )
    with col_xlsx:
        with st.container(border=True):
            _card_head("xlsx")
            xlsx_up = st.file_uploader(
                "Excel", type=["xlsx", "xls"], key="router_xlsx",
                label_visibility="collapsed",
            )
    with col_docx:
        with st.container(border=True):
            _card_head("docx")
            docx_up = st.file_uploader(
                "DOCX", type=["docx"], key="router_docx",
                label_visibility="collapsed",
            )

# ---- 2. Search (input + button, same height, same baseline) ----
_section("Search")
with _scope("po_search"):
    sc_input, sc_btn = st.columns([8, 2], gap="small")
    with sc_input:
        st.text_input(
            "Search",
            placeholder="Search packing lists, POs, styles, buyers...",
            label_visibility="collapsed",
            key="router_search",
        )
    with sc_btn:
        st.button("Search", key="router_search_btn", use_container_width=True)

# ---- 3. Green database box + preview panel ----
_section("Final packing-list database", "Save completed Excel files")
with _scope("po_db"):
    final_files = final_db.uploader_box()
    final_db.render_panel(final_files)

# ---- Sidebar ----
with st.sidebar:
    st.markdown("### Buyer")
    override = st.selectbox(
        "Pick a buyer if auto-detection fails",
        ["Auto-detect"] + [APPS[k][2] for k in APPS],
        key="router_override",
    )
    st.divider()
    st.markdown("### Supported buyers")
    for ft in ("pdf", "xlsx", "docx"):
        st.markdown(
            f'<div class="po-group">{TYPE_META[ft][1]}</div>'
            f'{_buyer_list(ft)}',
            unsafe_allow_html=True,
        )


# ============================================================
# Logic: which boxes are filled?
# ============================================================
filled = [x for x in (pdf_up, xlsx_up, docx_up) if x is not None]

if len(filled) > 1:
    st.warning("⚠️ **You uploaded more than one file.** "
               "Please clear the other boxes and upload a single file at a time.")
    st.stop()

# ---- Nothing uploaded ----
if not filled:
    if override == "GAP (Old Navy)":
        st.divider()
        run_app("GAP", None)
    st.stop()

# ============================================================
# Routing logic (unchanged)
# ============================================================
if pdf_up is not None:
    key = LABEL_TO_KEY.get(override)
    if key and APPS[key][3] != "pdf":
        st.error(f"The **{APPS[key][2]}** module expects a `.{APPS[key][3]}` file, but you uploaded a PDF.")
        st.stop()
    if key is None:
        try:
            key = detect_buyer_pdf(pdf_up.getvalue())
        except Exception as e:
            st.error(f"Could not read the PDF: {e}")
            st.stop()
    if key is None:
        st.warning("Could not recognise the PDF buyer. Pick the buyer in the sidebar and try again.")
        st.stop()
    _buyer_banner(key)
    st.divider()
    run_app(key, pdf_up)
    st.stop()

if xlsx_up is not None:
    key = LABEL_TO_KEY.get(override)
    if key and APPS[key][3] not in ("xlsx", "xls"):
        st.error(f"The **{APPS[key][2]}** module expects a `.{APPS[key][3]}` file, but you uploaded an Excel file.")
        st.stop()
    if key is None:
        try:
            key = detect_buyer_xlsx(xlsx_up.getvalue(), getattr(xlsx_up, "name", ""))
        except Exception:
            key = "NX"
    _buyer_banner(key)
    st.divider()
    run_app(key, xlsx_up)
    st.stop()

if docx_up is not None:
    key = LABEL_TO_KEY.get(override)
    if key and APPS[key][3] != "docx":
        st.error(f"The **{APPS[key][2]}** module expects a `.{APPS[key][3]}` file, but you uploaded a DOCX.")
        st.stop()
    if key is None:
        try:
            key = detect_buyer_docx(docx_up.getvalue())
        except Exception as e:
            st.error(f"Could not read the DOCX: {e}")
            st.stop()
    if key is None:
        st.warning("Could not recognise the DOCX buyer. Pick the buyer in the sidebar and try again.")
        st.stop()
    _buyer_banner(key)
    st.divider()
    run_app(key, docx_up)
    st.stop()