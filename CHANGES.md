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

## v3.1 - new shared UI (step 1 of the redesign)
- NEW `pl_kit.py`: shared screen blocks for all buyers (summary cards, editable grid with add/delete row,
  "Reset to PO", live order-vs-packed check, download button).
- Justice page rebuilt on it: editable packing table (size ratios, cartons, carton size, net kg/pc),
  variables, live check, Excel built from the edited table. PO order quantities are never editable.
- Next: Tally, Banana Republic, Kontoor, s.Oliver, Carter's, Abercrombie, Haddad, Next, MUJI, JCPenney, then VF and GAP.
- v3.1b: delete rows is now a visible 🗑 Delete tick-box column + "Remove ticked rows" button (the built-in row-select box was not found by users).

## v3.3 - shared UI rolled out (Justice, A&F, Tally Weijl, Haddad)
Same 5-6 step page on all four: summary cards -> quantities -> variables -> editable carton table
(edit cells, + add row, tick 🗑 Delete + "Remove ticked rows", "Reset table to PO") -> check vs PO -> download.
- A&F: `plan_from_rows()` rebuilds the plan from the edited carton table (Excel + PDF use it).
- Tally: `with_ship_qty()` + `plan_from_rows()`.
- Haddad: `compute_rows()`; new rows pick a PO# and take style / label / ref / colour from it.
- Changing quantities or variables above the table rebuilds the table (edits to it are replaced).
- Not yet converted: Banana Republic, Kontoor, s.Oliver, Carter's, Next, MUJI, JCPenney, GAP, VF (unchanged, still work as before).

## v3.4
- s.Oliver: editable carton table (colour / inseam / waist / pcs / cartons, add + tick-delete rows); Excel uses it.
- Kontoor: same step layout; visible delete; tolerance-aware check (+/- % from the PO) via `pl_kit.check_by_size(over_pct, under_pct)`.
- Banana Republic: same step layout (settings moved from the sidebar to the page); keeps its interactive carton grid;
  new order-vs-packed check and summary cards.
- MUJI: same step layout (packs follow the fixed MUJI reference mixes, so no free row editing).
- Still on the old screens: Carter's, Next, JCPenney, GAP, VF.

## v3.5
- Carter's: same step layout around its interactive carton grid; new order-vs-packed check; the Excel is now rebuilt
  on every run, so it always has your latest grid edits (the old "Prepare" button served a stale first copy).
- Next / Clover: editable carton table (size, cartons, pcs, carton size, gross / net per carton); Excel uses it.
- JCPenney: editable carton table (length ratios, pcs per carton, cartons, code, colour); weights moved from the
  sidebar to a "Variables" step; check by length.
- Still on the old screens: GAP, VF.

## v3.6 - all 12 buyers on the shared layout
- GAP / Old Navy: settings moved from the sidebar to a "Variables" step; Bulk and Single sections now have editable
  carton tables (size, pcs/ctn, cartons, prepack; + row, tick-delete, reset). Carton numbers, weights and carton codes
  are recalculated from your rows (`color_builders.rebuild_one_size_section`); Multi Y / Multi N stay read-only.
  Bulk check uses the "max excess %" as tolerance. The Excel is rebuilt on every run (no stale "Prepare" copy).
- VF / Vans: same step layout; picking a PO line now shows its editable carton table straight away (no need to press
  "Build sheet" first); download buttons no longer disappear after the next click. The carton grid itself is unchanged.

## v3.7
- VF / Vans carton grid: the two "＋ / −" columns are gone. Same controls as the other buyers now:
  a 🗑 tick box as the first column, and under the grid "＋ Add row", "🗑 Remove ticked rows (n)" and
  "↺ Reset table to PO". TOTAL and REMAIN labels moved into the "#" column; frozen columns are now 🗑, #, Ctn No, Color, UPC.

## v3.8
- VF / Vans carton grid: header, ORDER QTY, TOTAL and REMAIN rows now stay pinned.
  Cause: Streamlit's -1rem bottom margin on text cells made each pinned row 16px shorter than drawn, so the rows no
  longer touched and carton rows scrolled through the gaps. The same margin cancel that fixed Banana Republic is now in the VF grid CSS.

## v8.9
- VF / Vans carton grid: header, ORDER QTY, TOTAL and REMAIN rows were still scrolling away on some Streamlit versions.
  Cause: the pinning CSS expected the keyed row container ONE level under the scroller; older Streamlit puts it two levels
  down (BorderWrapper > div > keyed block), so nothing was pinned. `_sticky_sel` now matches at any depth, and the
  "display: contents" flatten rule no longer touches the keyed row blocks. Tested in a real browser on Streamlit 1.42.2
  and 1.66 (vertical + horizontal scrolling). Same selector fix applied to `grid_editor.py` (Banana Republic etc.).
- NEW green upload box (4th box on the router page): upload the FINAL completed Excel files (.xlsx / .xlsm / .xls, several at
  once). Page shows what was read (sheets, rows, buyer / PO / style), then "Save to database" stores them in SQLite
  (`data/packing_list.db`, env `PL_DB_PATH`): the original file, every sheet and every non-empty row. Same file twice is not
  stored twice; same name with new content becomes the next version. A "Stored files" browser shows the rows and downloads
  the original. Warns when formulas have no saved value (file never opened in Excel). New file `final_db.py`.
- Docker: `data/` is a named volume (`po_data`) so the database survives rebuilds; `.dockerignore` excludes `data/`.

## v9.10 – GAP unified Excel: formulas + same-sheet summary
- **Every derived cell is now an Excel formula** in the GAP unified packing-list workbook
  (`export_unified_excel` in `Gap/unified_packing.py`):
  - PER BLST = SUM of size columns on the row
  - QTY PER CTN = PER BLST (× prepacks when applicable)
  - TOTAL QTY = QTY PER CTN × CTN QTY
  - TOTAL row = SUM / SUMPRODUCT of the data rows
  - Header SHIP QTY / CTN QTY / EX/SHORT link to the TOTAL row
- **Summary table sits on the same sheet**, directly after the packing-list rows
  (COLOR | ORDER QTY | SHIP QTY | EX/SHORT | CTN QTY | size breakdown).  
  All summary numbers are formulas that reference the TOTAL rows above.
- **All packing types (Bulk, Single, Multi Y, Multi N) stay in one single XLSX**;
  each type that has data gets its own sheet; an Overview sheet indexes them.
- Changing any CTN QTY or size-unit cell automatically recalculates related totals.
