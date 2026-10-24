"""
Tally Weijl (TSWL) buyer: Order Form PDF → factory packing list (Excel).

Pure logic, no Streamlit.  UI lives in tw_ui.py.

Packing rule (from reference sheets):
  SOLID pack, one size per carton group.
  Full cartons hold `std` pcs (user-set, default 28 or 60 by style).
  Remainder of each size goes into one extra (small) carton.
  Carton dims default 60×40×30 cm (full) / 60×40×20 or 60×40×10 (small).
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

HERE = Path(__file__).parent

# defaults (kg / cm)
DEF_STD = 28
DEF_NET = 0.35          # kg per garment (placeholder; user overrides per size)
DEF_TARE = 1.50         # kg per carton
DEF_BIG = (60.0, 40.0, 30.0)
DEF_SMALL_H = {          # height by remainder size band
    "tiny": 10.0,        # < 15 pcs
    "small": 20.0,       # 15–std/2
    "mid": 30.0,         # otherwise use full height
}

THIN = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
YELLOW = PatternFill("solid", fgColor="FFFF00")
BOLD = Font(name="Calibri", size=12, bold=True)
TITLE = Font(name="Calibri", size=26, bold=True)
NORMAL = Font(name="Calibri", size=11)
ARIAL10 = Font(name="Arial", size=10)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")


# --------------------------------------------------------------------------- data
@dataclass
class SizeLine:
    size: str
    color_code: str
    color_name: str
    ean: str
    qty: int
    unit_usd: float = 0.0


@dataclass
class PoData:
    order_no: str
    style: str
    item_desc: str
    material: str
    hs_code: str
    destination: str
    warehouse: str
    ship_date: str
    incoterm: str
    supplier_code: str
    total_qty: int
    lines: list[SizeLine] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def sizes(self) -> list[str]:
        # preserve order of appearance
        seen, out = set(), []
        for ln in self.lines:
            if ln.size not in seen:
                seen.add(ln.size)
                out.append(ln.size)
        return out

    @property
    def color_name(self) -> str:
        return self.lines[0].color_name if self.lines else ""

    @property
    def color_code(self) -> str:
        return self.lines[0].color_code if self.lines else ""


# --------------------------------------------------------------------------- parse PDF
_SIZE_TOKEN = re.compile(
    r"(?P<style>[A-Z0-9]+)\s+"
    r"(?P<code>[A-Z0-9]+),(?P<size>[A-Z0-9/]+)\s+"
    r"(?P<color>.+?)\s+"
    r"(?P<ean>\d{13})\s+"
    r"(?P<eur>[\d.,]+)\s*EUR\s+"
    r"(?:(?P<chf>[\d.,]+)\s*CHF\s+)?"
    r"(?P<qty>\d{1,3}(?:\s\d{3})*|\d+)\s+"
    r"(?P<usd>[\d]+[.,][\d]+|\d+)\s+"
    r"(?P<amt>[\d\s.,]+)",
    re.I,
)

# fallback: looser line with style + size code + qty
_LOOSE = re.compile(
    r"(?P<style>[A-Z]{2,}[A-Z0-9]+)\s+"
    r"(?P<code>[A-Z]{2,}),(?P<size>[A-Z0-9/]+)\s+"
    r"(?P<color>[A-Za-z][A-Za-z0-9 %/\-]+?)\s+"
    r"(?P<ean>\d{13})",
    re.I,
)


def _page_texts(data: bytes) -> list[str]:
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return [p.extract_text() or "" for p in pdf.pages]


def parse_po(data: bytes) -> PoData:
    pages = _page_texts(data)
    text = "\n".join(pages)
    upper = text.upper()

    def grab(pat: str, default: str = "") -> str:
        m = re.search(pat, text, re.I | re.M)
        return m.group(1).strip() if m else default

    order_no = grab(r"ORDER FORM\s*\(1\)\s*N[o°º]?\s*(\d+)", "")
    if not order_no:
        order_no = grab(r"No\s*(\d{6,})", "")

    warehouse = grab(r"WAREHOUSE\s+([A-ZÄÖÜß /\-]+)", "")
    supplier_code = grab(r"Supplier.?s Tally-?Weijl code\s*\(1\)\s*(\d+)", "126198")
    ship_date = grab(r"(\d{2}/\d{2}/\d{4})\s*,?\s*FCA", "")
    incoterm = grab(r"(FCA\s+\w+)", "FCA Dhaka")
    destination = grab(r"Country of destination\s+(\w+)", "Germany")
    if not destination or destination.upper() == "OF":
        destination = "Germany"

    # material / HS from first style block
    material = grab(r"Main material[^:]*:\s*([^\n]+)", "")
    hs_code = grab(r"HS CODE[:\s]*(\d+)", "")

    # item description from page 2 style block
    item_desc = grab(r"Commercial description\s+([^\n]+)", "")
    if not item_desc:
        item_desc = grab(r"woven\s+([^\n]+)", "")

    lines: list[SizeLine] = []
    style = ""
    # Parse line-by-line (joining pages breaks color boundaries)
    line_re = re.compile(
        r"^(?P<style>[A-Z][A-Z0-9]{4,})\s+"
        r"(?P<code>[A-Z0-9]+),(?P<size>[A-Z0-9/]+)\s+"
        r"(?P<color>.+?)\s+"
        r"(?P<ean>\d{13})\s+"
        r"(?P<eur>[\d.,]+)\s*EUR\s+"
        r"(?:(?P<chf>[\d.,]+)\s*CHF\s+)?"
        r"(?P<qty>\d{1,3}(?:\s\d{3})*|\d+)\s+"
        r"(?P<usd>[\d]+[.,][\d]+|\d+)",
        re.I,
    )
    for raw in text.splitlines():
        ln = re.sub(r"\s+", " ", raw).strip()
        m = line_re.match(ln)
        if not m:
            continue
        if not style:
            style = m.group("style").upper()
        size = m.group("size").strip().upper()
        color = m.group("color").strip()
        color = re.sub(r",?\s*" + re.escape(size) + r"\s*$", "", color, flags=re.I).strip()
        if color and "wash" not in color.lower() and "black used" in color.lower():
            color = color + " Wash"
        qty = int(re.sub(r"\s", "", m.group("qty")))
        try:
            usd = float(m.group("usd").replace(",", "."))
        except ValueError:
            usd = 0.0
        lines.append(SizeLine(
            size=size,
            color_code=m.group("code").upper(),
            color_name=color,
            ean=m.group("ean"),
            qty=qty,
            unit_usd=usd,
        ))

    if not style:
        for cand in re.findall(r"\b(S[A-Z]{6,15})\b", text):
            if cand not in ("SWITZERLAND", "SPECIFICATIONS"):
                style = cand
                break
    if not style:
        style = "UNKNOWN"

    warnings = []
    if not lines:
        warnings.append("No size lines parsed – check that this is a Tally Weijl Order Form PDF.")
    if any(ln.qty <= 0 for ln in lines):
        warnings.append("Some sizes have qty 0 – verify numbers on the PO.")

    total = sum(ln.qty for ln in lines)
    return PoData(
        order_no=order_no or "UNKNOWN",
        style=style or "UNKNOWN",
        item_desc=item_desc or "",
        material=material,
        hs_code=hs_code,
        destination=destination,
        warehouse=warehouse,
        ship_date=ship_date,
        incoterm=incoterm,
        supplier_code=supplier_code,
        total_qty=total,
        lines=lines,
        warnings=warnings,
    )


# --------------------------------------------------------------------------- carton plan
@dataclass
class CartonRow:
    size: str
    ctns: int
    pcs_per_ctn: int
    total_pcs: int
    L: float
    W: float
    H: float
    is_remainder: bool


def build_carton_plan(
    po: PoData,
    std: int,
    dims_big: tuple[float, float, float],
    dims_small: tuple[float, float, float] | None = None,
) -> list[CartonRow]:
    """One-size-solid plan: full cartons of `std` + one remainder carton per size."""
    Lb, Wb, Hb = dims_big
    rows: list[CartonRow] = []
    for ln in po.lines:
        if ln.qty <= 0:
            continue
        full = ln.qty // std
        rem = ln.qty % std
        if full:
            rows.append(CartonRow(ln.size, full, std, full * std, Lb, Wb, Hb, False))
        if rem:
            if dims_small:
                Ls, Ws, Hs = dims_small
            else:
                Hs = DEF_SMALL_H["tiny"] if rem < 15 else (
                    DEF_SMALL_H["small"] if rem <= std // 2 else Hb
                )
                Ls, Ws = Lb, Wb
            rows.append(CartonRow(ln.size, 1, rem, rem, Ls, Ws, Hs, True))
    return rows


def with_ship_qty(po: PoData, ship: dict[str, int]) -> PoData:
    """Copy of the PO whose line quantities are the quantities the user wants to ship (for planning only)."""
    import dataclasses
    lines = [dataclasses.replace(ln, qty=int(ship.get(ln.size, ln.qty))) for ln in po.lines]
    return dataclasses.replace(po, lines=lines)


def plan_from_rows(po: PoData, rows: list[dict]) -> tuple[list[CartonRow], list[str]]:
    """User-edited carton table -> carton plan.
    rows keys: size, ctns, pcs, L, W, H, small (bool). Invalid rows are skipped and reported."""
    out, problems = [], []
    for n, r in enumerate(rows, 1):
        size = r.get("size")
        try:
            ctns, pcs = int(float(r.get("ctns"))), int(float(r.get("pcs")))
            L, W, H = (float(r.get(k)) for k in ("L", "W", "H"))
        except (TypeError, ValueError):
            if size:
                problems.append(f"Row {n}: cartons, pcs and size must be numbers - row ignored.")
            continue
        if size not in po.sizes:
            problems.append(f"Row {n}: pick a size from the list - row ignored.")
            continue
        if ctns <= 0 or pcs <= 0 or min(L, W, H) <= 0:
            problems.append(f"Row {n}: cartons, pcs per carton and measurements must be above 0 - row ignored.")
            continue
        out.append(CartonRow(size, ctns, pcs, ctns * pcs, L, W, H, bool(r.get("small"))))
    return out, problems


# --------------------------------------------------------------------------- Excel
def _style_header_row(ws, row: int, cols: int):
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.border = THIN
        cell.alignment = CENTER


def make_packing_excel(
    po: PoData,
    plan: list[CartonRow],
    net_per_size: dict[str, float],
    tare: float,
    company: str = "THAT'S IT SPORT'S WEAR LTD.     ( UNIT = 04 )",
    address: str = "Narashighpur, Ashulia, Dhaka-1341.",
) -> bytes:
    """Build factory packing-list workbook matching the reference layout."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{po.style}_{po.order_no}"[:31]

    sizes = po.sizes
    # column map
    # A Carton from | B Carton to | C Total Ctns | D Measurements | E CBM
    # F Destination | G COLOR | H.. size cols | then Pcs/Ctn | Total Pcs |
    # Nt.Wt/Gmt | Nt.Wt/Ctn | Gr.Wt/Ctn | Total Nt | Total Gr
    size_start = 8  # column H
    n_sizes = len(sizes)
    col_pcs = size_start + n_sizes          # Pcs / Ctn
    col_tot = col_pcs + 1                   # Total Pcs
    col_netg = col_tot + 1                  # Nt. Wt / Gmt
    col_netc = col_netg + 1                 # Nt. Wt / Ctn
    col_grc = col_netc + 1                  # Gr.Wt / Ctn
    col_tnet = col_grc + 1                  # Total Nt. Wt
    col_tgr = col_tnet + 1                  # Total Gr. Wt
    last_col = col_tgr

    # ---- title block
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws["A1"] = company
    ws["A1"].font = TITLE
    ws["A1"].alignment = LEFT

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws["A2"] = address
    ws["A2"].font = ARIAL10

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)
    ws["A3"] = "PACKING LIST"
    ws["A3"].font = Font(name="Arial", size=12, bold=True)
    ws["A3"].alignment = CENTER

    # ---- meta row 4-7
    order_qty = po.total_qty
    packed_qty = sum(r.total_pcs for r in plan)
    excess = packed_qty - order_qty

    meta = [
        (4, "Date", "", "Color", po.color_name, "Order Qty", order_qty),
        (5, "Invoice No.", "", "Fabric", po.material or "", "Packed Qty", packed_qty),
        (6, "PO Number", po.order_no, "Item", po.item_desc or "", "Excess / Shortage", excess),
        (7, "Style No", po.style, "Net Weight", "", "Gross Weight", ""),
    ]
    for row, a, b, f, g, m, p in meta:
        ws.cell(row, 1, a).font = BOLD
        ws.cell(row, 4, b).font = BOLD
        ws.cell(row, 6, f).font = BOLD
        ws.cell(row, 8, g).font = BOLD
        ws.cell(row, 13, m).font = BOLD
        ws.cell(row, 16, p).font = BOLD

    # excess %
    if order_qty:
        ws.cell(6, 17, round(excess / order_qty, 6)).font = BOLD

    # ---- header row 8
    headers = {
        1: "Carton Nos", 3: "Total Ctns", 4: "Carton Measurements", 5: "CBM",
        6: "Destination", 7: "COLOR",
    }
    for i, s in enumerate(sizes):
        headers[size_start + i] = s
    headers[col_pcs] = "Pcs / Ctn"
    headers[col_tot] = "Total Pcs"
    headers[col_netg] = "Nt. Wt / Gmt"
    headers[col_netc] = "Nt. Wt / Ctn"
    headers[col_grc] = "Gr.Wt / Ctn"
    headers[col_tnet] = "Total Nt. Wt"
    headers[col_tgr] = "Total Gr. Wt"
    for c, v in headers.items():
        cell = ws.cell(8, c, v)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.border = THIN
        cell.alignment = CENTER

    # ---- body
    body0 = 9
    ctn_from = 1
    for i, r in enumerate(plan):
        row = body0 + i
        ctn_to = ctn_from + r.ctns - 1
        cbm = r.L * r.W * r.H * r.ctns / 1_000_000
        net_g = float(net_per_size.get(r.size, DEF_NET))
        net_c = round(net_g * r.pcs_per_ctn, 2)
        gr_c = round(net_c + tare, 2)
        tot_net = round(net_c * r.ctns, 2)
        tot_gr = round(gr_c * r.ctns, 2)

        ws.cell(row, 1, 1 if i == 0 else f"=B{row - 1}+1").border = THIN
        ws.cell(row, 2, f"=A{row}+C{row}-1").border = THIN
        ws.cell(row, 3, r.ctns).border = THIN
        ws.cell(row, 4, f"{r.L:g} X {r.W:g} X {r.H:g} CM").border = THIN
        ws.cell(row, 5, round(cbm, 3)).border = THIN
        ws.cell(row, 6, po.destination if i == 0 else "").border = THIN
        ws.cell(row, 7, po.color_name if i == 0 else "").border = THIN

        for si, s in enumerate(sizes):
            cell = ws.cell(row, size_start + si)
            cell.border = THIN
            cell.alignment = CENTER
            if s == r.size:
                cell.value = r.pcs_per_ctn

        sz0, sz1 = CL(size_start), CL(size_start + n_sizes - 1)
        # LIVE formulas: change a size qty / ctns / Nt.Wt-per-Gmt and every weight recalculates
        ws.cell(row, col_pcs, f"=SUM({sz0}{row}:{sz1}{row})").border = THIN
        ws.cell(row, col_tot, f"=C{row}*{CL(col_pcs)}{row}").border = THIN
        ws.cell(row, col_netg, net_g).border = THIN
        ws.cell(row, col_netc, f"=ROUND({CL(col_netg)}{row}*{CL(col_pcs)}{row},2)").border = THIN
        ws.cell(row, col_grc, f"=ROUND({CL(col_netc)}{row}+{tare:g},2)").border = THIN
        ws.cell(row, col_tnet, f"=ROUND({CL(col_netc)}{row}*C{row},2)").border = THIN
        ws.cell(row, col_tgr, f"=ROUND({CL(col_grc)}{row}*C{row},2)").border = THIN
        for _c in (col_netc, col_grc, col_tnet, col_tgr):
            ws.cell(row, _c).number_format = "0.00"

        for c in range(1, last_col + 1):
            ws.cell(row, c).alignment = CENTER
            ws.cell(row, c).font = NORMAL

        ctn_from = ctn_to + 1

    body1 = body0 + len(plan) - 1
    sub_row = body1 + 1

    # SUBTOTAL
    ws.cell(sub_row, 1, "SUBTOTAL").font = BOLD
    b0, b1 = body0, body1
    ws.cell(sub_row, 3, f"=SUM(C{b0}:C{b1})").font = BOLD
    ws.cell(sub_row, 5, f"=ROUND(SUM(E{b0}:E{b1}),3)").font = BOLD
    # size totals = sum of (size cell x ctns) over the body rows
    for si, s in enumerate(sizes):
        c_ = CL(size_start + si)
        ws.cell(sub_row, size_start + si, f"=SUMPRODUCT({c_}{b0}:{c_}{b1},$C${b0}:$C${b1})").font = BOLD
    ws.cell(sub_row, col_tot, f"=SUM({CL(col_tot)}{b0}:{CL(col_tot)}{b1})").font = BOLD
    ws.cell(sub_row, col_tnet, f"=ROUND(SUM({CL(col_tnet)}{b0}:{CL(col_tnet)}{b1}),2)").font = BOLD
    ws.cell(sub_row, col_tgr, f"=ROUND(SUM({CL(col_tgr)}{b0}:{CL(col_tgr)}{b1}),2)").font = BOLD
    total_net = sum(float(net_per_size.get(r.size, DEF_NET)) * r.total_pcs for r in plan)
    total_gr = sum((float(net_per_size.get(r.size, DEF_NET)) * r.pcs_per_ctn + tare) * r.ctns for r in plan)
    for c in range(1, last_col + 1):
        ws.cell(sub_row, c).border = THIN
        ws.cell(sub_row, c).fill = HEADER_FILL

    # fill Net / Gross weight in header
    ws.cell(7, 8, f"={CL(col_tnet)}{sub_row}").font = BOLD
    ws.cell(7, 16, f"={CL(col_tgr)}{sub_row}").font = BOLD

    # ---- GRAND TOTAL
    grand = sub_row + 2
    ws.cell(grand, 1, "GRAND Total").font = BOLD
    ws.cell(grand, 3, f"=C{sub_row}").font = BOLD
    ws.cell(grand, 4, "TOTAL CBM").font = BOLD
    ws.cell(grand, 5, f"=E{sub_row}").font = BOLD
    for si, s in enumerate(sizes):
        ws.cell(grand, size_start + si, f"={CL(size_start + si)}{sub_row}").font = BOLD
    ws.cell(grand, col_tot, f"={CL(col_tot)}{sub_row}").font = BOLD
    ws.cell(grand, col_tnet, f"={CL(col_tnet)}{sub_row}").font = BOLD
    ws.cell(grand, col_tgr, f"={CL(col_tgr)}{sub_row}").font = BOLD

    # ---- Abstract
    abs0 = grand + 2
    ws.cell(abs0, 3, "Abstract").font = BOLD
    ws.cell(abs0, 6, "Destination").font = BOLD
    for si, s in enumerate(sizes):
        ws.cell(abs0, size_start + si, s).font = BOLD
        ws.cell(abs0, size_start + si).fill = HEADER_FILL
    ws.cell(abs0, col_pcs, "Total").font = BOLD

    order_by_size = {ln.size: ln.qty for ln in po.lines}
    packed_by_size = {s: sum(r.total_pcs for r in plan if r.size == s) for s in sizes}

    sz_a, sz_z = CL(size_start), CL(size_start + n_sizes - 1)
    r_ord, r_pk, r_ex = abs0 + 1, abs0 + 2, abs0 + 3

    ws.cell(r_ord, 3, "Order Qty").font = BOLD
    ws.cell(r_ord, 6, po.destination)
    ws.cell(r_ord, 7, po.color_name)
    for si, s_ in enumerate(sizes):
        ws.cell(r_ord, size_start + si, order_by_size.get(s_, 0))          # order qty per size = PO input
    ws.cell(r_ord, col_pcs, f"=SUM({sz_a}{r_ord}:{sz_z}{r_ord})")

    # LIVE: packed qty per size comes straight from the SUBTOTAL row, so it follows every edit above
    ws.cell(r_pk, 3, "Packed Qty").font = BOLD
    for si, s_ in enumerate(sizes):
        ws.cell(r_pk, size_start + si, f"={CL(size_start + si)}{sub_row}")
    ws.cell(r_pk, col_pcs, f"=SUM({sz_a}{r_pk}:{sz_z}{r_pk})")

    ws.cell(r_ex, 3, "Excess / Short").font = BOLD
    for si, s_ in enumerate(sizes):
        c_ = CL(size_start + si)
        ws.cell(r_ex, size_start + si, f"={c_}{r_pk}-{c_}{r_ord}")
    ws.cell(r_ex, col_pcs, f"={CL(col_pcs)}{r_pk}-{CL(col_pcs)}{r_ord}")

    # color summary box
    ws.cell(abs0 + 1, col_tot, "COLOR CODE").font = BOLD
    ws.cell(abs0 + 1, col_netg, "ORDER QTY").font = BOLD
    ws.cell(abs0 + 1, col_netc, "SHIP QTY").font = BOLD
    ws.cell(abs0 + 1, col_grc, "EXCESS/SHORT").font = BOLD
    ws.cell(abs0 + 1, col_tnet, "EXCESS/SHORT %").font = BOLD

    ws.cell(abs0 + 2, col_tot, po.color_code)
    ws.cell(abs0 + 2, col_netg, f"={CL(col_pcs)}{r_ord}")
    ws.cell(abs0 + 2, col_netc, f"={CL(col_pcs)}{r_pk}")
    ws.cell(abs0 + 2, col_grc, f"={CL(col_netc)}{abs0 + 2}-{CL(col_netg)}{abs0 + 2}")
    ws.cell(abs0 + 2, col_tnet,
            f"=IF({CL(col_netg)}{abs0 + 2}=0,0,{CL(col_grc)}{abs0 + 2}/{CL(col_netg)}{abs0 + 2})")

    # header block (rows 4-6, right side) follows the same live numbers
    ws.cell(4, 16, f"={CL(col_pcs)}{r_ord}")
    ws.cell(5, 16, f"={CL(col_pcs)}{r_pk}")
    ws.cell(6, 16, f"=P5-P4")
    ws.cell(6, 17, "=IF(P4=0,0,P6/P4)")

    # column widths
    widths = {
        1: 12, 2: 12, 3: 11, 4: 26, 5: 10, 6: 12, 7: 20,
    }
    for i in range(n_sizes):
        widths[size_start + i] = 10
    widths[col_pcs] = 11
    widths[col_tot] = 11
    widths[col_netg] = 13
    widths[col_netc] = 12
    widths[col_grc] = 12
    widths[col_tnet] = 12
    widths[col_tgr] = 12
    for c, w in widths.items():
        ws.column_dimensions[CL(c)].width = w

    ws.row_dimensions[1].height = 32
    ws.print_title_rows = "1:8"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
