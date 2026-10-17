# What changed (pipeline / logic / rules untouched)

## Speed
- VF/Vans: `render_packing_list_preview` is now a Streamlit fragment, so ＋ / ❌ / cell edits
  rerun only the grid + totals + preview (not the whole page, tables and JSON download).
  "Reset" reruns the fragment only. Needs Streamlit >= 1.37 (requirements.txt updated).
- Router (`app.py`): buyer detection is cached (no more pdfplumber re-open on every click);
  buyer scripts are compiled/imported once per file-mtime instead of on every rerun.

## Look (Streamlit == Excel)
- `theme.py` holds ONE palette/font: yellow FFFF00, grey D9D9D9, Arial. The VF preview
  yellow now equals the Excel fill (was FFFF66). No grey-out/flicker while a rerun runs.
- Excel files: print setup only (landscape, A4, fit to 1 page wide, centred, narrow margins).
  Cell values, formulas, fonts, fills, borders, merges, widths are NOT touched. Workbooks
  with images/drawings are skipped.

## UI beautify + centred cells (pipeline / logic / rules untouched)
- `theme.py` only (+ one line in `app.py`: `st.title` -> `theme.hero(...)`).
- Streamlit: gradient header banner, dashed "drop card" uploaders, rounded accent buttons / inputs /
  alerts / expanders / tabs / metrics, tidier sidebar. Every HTML preview table
  (`table.plist`, `.order-tbl`, `.sidebar-qty`, incl. description / row-label cells) is centred.
  Inputs are centred. `st.dataframe` text is centred through a display-only Styler.
- Excel: every cell centred (horizontal + vertical), one font (Arial), gridlines off, tab colour,
  rows never shorter than 16pt, widths only added for columns that had none and would clip.
  Text that overflowed its narrow cell uses "Center Across Selection" (or shrink-to-fit if a
  neighbour is in the way). Values, formulas, fills, borders, merges, rotations, wrap flags and
  explicit widths are NOT changed.
- Workbooks with images (e.g. Haddad logo) used to be skipped; now only `xl/styles.xml` + sheet
  views are patched (centre / font / gridlines), every other part is copied byte-for-byte.

## Docker (port 8003)
- `Dockerfile`, `docker-compose.yml`: Streamlit now served on 8003.
- `.dockerignore` no longer drops `*.xlsx/*.pdf/*.docx` (that removed `anf_template.xlsx` and
  `haddad_template.xlsx`, which those buyers need). Only sample output files are ignored.
- `requirements.txt`: streamlit>=1.37 (fragments), + xlrd, Pillow.

## UI + routing fixes (second pass)
- `grid_editor.py`: (1) unique container key per grid (StreamlitDuplicateElementKey); (2) NEW - the carton rows were hidden
  behind the pinned header / ORDER / TOTAL / REMAIN rows because Streamlit gives markdown blocks a -1rem bottom margin
  (each row was 16px shorter than it was drawn). The margin is cancelled inside the grid.
- `app.py`: Tally Weijl detection looks at the first two pages and ignores spaces / hyphens ("Tally-Weijl"), so it no longer
  falls through to another buyer.
- `Gap/unified_packing.py`: when the PDF has no GAP tables the page shows a clear message (pick the right buyer) instead of
  dumping raw JSON; the JSON is in a collapsed "Technical details" box.
- `br/br_packing.py`: Excel prints one colour per page (was one tiny page), wider columns so labels are not cut,
  trailing columns follow the number of sizes, case numbers / header totals / summary are formulas.
- `kontor/kontoor_app.py`: up to 40 sizes. Duplicate download key in Gap fixed.

## v3
- VF/Vans: sidebar "Order qty by size" no longer raises `StreamlitAPIException` (st.sidebar inside st.fragment);
  the sidebar slot is created outside the fragment and the fragment writes into it.
- NEW buyer: **Justice (SPA / Adjmi)** - folder `justice/` (`justice_packing.py`, `justice_ui.py`, `justice_app.py`).
  Router (`app.py`): key `JU`, auto-detected from "Sports Products of America" / "Adjmi Apparel" (first 2 pages).
  Handles PREPACK lines (W1-W5, WM size ratios) and SOLID lines (WL, 12 pcs of one size).
  Checks units = ctn x pcs/ctn, ratio sum = pcs/ctn and PO totals. Excel: carton ranges, size-ratio columns,
  CBM / weights by formula, size summary (order vs packed). Yellow cells = measurement / weights to enter.
