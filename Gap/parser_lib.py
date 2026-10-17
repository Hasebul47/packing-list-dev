import io
import re
import json
import math

import pdfplumber


# ===========================================================================
# 1. Shared column schema (updated)
# ===========================================================================
COLUMNS = [
    "BUYER",
    "STYLE",
    "P.O. #",
    "Dept",
    "Style No",
    "Full Carton",
    "PrePack",
    "PrePack Type",
    "PrePack Pack Factor",
    "SKU / ITEM",
    "Units per PrePack",
    "UNIT / PREPACK",
    "Style Description",
    "Item Carton Pack Factor",
    "ORDER QTY",
    "SHIP QTY",
    "EX-SHORT",
    "CTN QTY",
    "CARTON QTY",
    "#PrePackets Ordered",
    "Universal CC #Color Desc",
    "Size Desc",
    "Unit Cost",
    "Total Cost",
]


# ===========================================================================
# 2. Lines that must never be treated as a color name
# ===========================================================================
IGNORE_PATTERNS = [
    "Print Date", "OLD NAVY", "Global Reference", "Market:",
    "Market Channel", "Department:", "Destination Purchase Order", "Vendor",
    "Factory", "Agent Name", "Payment Type", "Payment Terms", "Payment Method",
    "Purchaser Currency", "Contract Ship", "Ship Cancel", "Terms of Sale",
    "Freight Paid By", "Transfer Point", "Special Instructions", "Ship Mode",
    "Country of Origin", "Country of Destination", "Powered by Infor Nexus",
    "Dept Style No", "Total Units", "Total BULK", "Total SINGLE", "Total MULTIPLE",
    "Total ASSORTED", "GRAND TOTAL", "THIS PURCHASE ORDER", "1 THIS EDI",
    "2 VENDOR SHALL", "3 NOTHWITHSTANDING", "4 PLEASE PUT", "5 THE TRACEABILITY",
    "6 A VARIANCE", "NOTE: FOR VENDORS", "=====", "BUYER", "STYLE", "P.O.",
    "OIQTY", "SHIP QTY", "EX/SHORT", "CTN QTY", "CTN MEAS",
]


# Fit words that mean "standard fit" and are not part of the size label
DEFAULT_FIT_WORDS = {"REG", "REGULAR"}

# Any item line: <4-digit dept> <6-digit style> ...  (was hard-coded to dept 3340)
ITEM_LINE_RE = re.compile(r"^\d{4}\s+\d{6,}\s+")  # 6+ digit style (e.g. 1183764)


def _to_int(tok: str) -> int:
    """'2,223' -> 2223"""
    return int(tok.replace(",", ""))


def normalize_size(size_tok: str, next_tok: str | None = None) -> tuple[str, bool]:
    """
    Normalize size tokens from PDF text.
      '5'            -> ('5', False)   consumed_next=False
      '10HUSKY'      -> ('10H', False)
      '8' + 'HUSKY'  -> ('8H', True)   consumed_next=True
      '12-18M'       -> ('12-18M', False)
      'XXL+'         -> ('XXL+', False)
    """
    s = (size_tok or "").strip()
    nxt = (next_tok or "").strip().upper()
    if not s:
        return s, False
    # glued husky: 10HUSKY / 8HUSKY
    m = re.match(r"^(\d+)HUSKY$", s, re.I)
    if m:
        return f"{m.group(1)}H", False
    # split: 8 HUSKY
    if nxt in ("HUSKY", "H"):
        if re.fullmatch(r"\d+", s):
            return f"{s}H", True
    # already 8H / 10H
    m = re.match(r"^(\d+)H$", s, re.I)
    if m:
        return f"{m.group(1)}H", False
    return s, False




def is_color_name(line: str) -> bool:
    """Return True if the line looks like an all-uppercase color description."""
    if not line or not line.strip():
        return False
    s = line.strip()
    if not s.isupper():
        return False
    if not all(c.isalpha() or c.isspace() for c in s):
        return False
    for pat in IGNORE_PATTERNS:
        if pat in s:
            return False
    return True


# ===========================================================================
# 3. PO header parser
# ===========================================================================
def parse_po_header(text: str) -> dict:
    """
    Extract PO-level fields that are FIXED for all tables:
        BUYER, STYLE, P.O. #
    plus derived totals:
        OIQTY, SHIP_QTY, EX_SHORT, CTN_QTY
    """
    header = {
        "BUYER": None,
        "STYLE": None,
        "P.O._#": None,
        "OIQTY": 0,
        "SHIP_QTY": 0,
        "EX_SHORT": 0,
        "CTN_QTY": None,
    }

    # ---- BUYER ----
    m = re.search(r"BUYER\s*[:\-]\s*([A-Z][A-Za-z0-9 &\.\-']+)", text)
    if m:
        header["BUYER"] = m.group(1).strip()

    # ---- STYLE ----
    m = re.search(r"STYLE\s*[:\-]\s*([0-9A-Za-z\-]+)", text)
    if m:
        header["STYLE"] = m.group(1).strip()

    # ---- P.O. # ----
    m = re.search(r"P\.?\s*O\.?\s*#?\s*[:\-]\s*([0-9A-Za-z\-]+)", text)
    if m:
        header["P.O._#"] = m.group(1).strip()

    # ---- OIQTY / SHIP QTY / EX/SHORT / CTN QTY ----
    m = re.search(r"OIQTY\s*[:\-]?\s*([\d,]+)", text)
    if m:
        header["OIQTY"] = int(m.group(1).replace(",", ""))

    m = re.search(r"SHIP\s*QTY\s*[:\-]?\s*([\d,]+)", text)
    if m:
        header["SHIP_QTY"] = int(m.group(1).replace(",", ""))

    m = re.search(r"EX\s*/?\s*SHORT\s*[:\-]?\s*([\d,]+)", text)
    if m:
        header["EX_SHORT"] = int(m.group(1).replace(",", ""))

    m = re.search(r"CTN\s*QTY\s*[:\-]?\s*([\d,]+)", text)
    if m:
        header["CTN_QTY"] = int(m.group(1).replace(",", ""))

    # ---- FIX 2: the PO PDF has no "BUYER:" / "P.O.:" labels — use its own labels ----
    if not header["BUYER"]:
        m = re.search(r"Brand:\s*([A-Z ]+?)\s+Global Reference", text)
        header["BUYER"] = m.group(1).strip() if m else None
    if not header["P.O._#"]:
        m = re.search(r"Destination Purchase Order # (\d+)", text)
        header["P.O._#"] = m.group(1) if m else None
    if not header["STYLE"]:
        m = re.search(r"^\d{4} (\d{6,}) ", text, re.M)
        header["STYLE"] = m.group(1) if m else None

    # Fallback: if OIQTY not found, sum from all "Total" lines later
    return header


# ===========================================================================
# 4. Item row parser
# ===========================================================================
def parse_line_item(line: str):
    """Parse a single '3340 ...' line into a row dict (new format)."""
    tokens = line.split()
    dept = tokens[0]
    style_no = tokens[1]
    idx = 2

    full_carton = None
    if idx < len(tokens) and tokens[idx] in ("N", "Y"):
        full_carton = tokens[idx]
        idx += 1

    prepack = tokens[idx]; idx += 1
    prepack_type = tokens[idx]; idx += 1

    # FIX 1: there is no pack-factor column on these lines. The next number is
    # the 13-digit SKU, then units per prepack (2 / 1, or "Bulk" for Bulk lines).
    prepack_pack_factor = None
    sku = tokens[idx]; idx += 1

    units_per_prepack = tokens[idx]
    if units_per_prepack.isdigit():
        units_per_prepack = int(units_per_prepack)
    idx += 1

    # Style Description: collect alphabetic tokens until a numeric token
    style_desc_tokens = []
    while idx < len(tokens):
        tok = tokens[idx]
        if re.fullmatch(r'\d+', tok.replace(',', '')):
            break
        style_desc_tokens.append(tok)
        idx += 1
    style_description = " ".join(style_desc_tokens)

    item_carton_pack_factor = None
    qty_ordered = None

    # The colour code ("001183764-000") splits the line:
    #   [icpf] qty  <colour code>  <size words...>  unit_cost  total_cost
    # Counting numbers before it is more reliable than counting all tokens,
    # because a size can be several words ("M REG", "10 HUSKY", "L TALL").
    cc_idx = next((j for j in range(idx, len(tokens))
                   if re.fullmatch(r"\d+-\d+", tokens[j])), None)
    if cc_idx is not None:
        pre = tokens[idx:cc_idx]
        if len(pre) >= 2:
            item_carton_pack_factor = int(pre[0].replace(",", ""))
            qty_ordered = int(pre[1].replace(",", ""))
        else:
            qty_ordered = int(pre[0].replace(",", ""))
        idx = cc_idx
    else:                                   # unexpected layout: old token-count rule
        remaining = len(tokens) - idx
        if remaining >= 6:
            item_carton_pack_factor = int(tokens[idx].replace(",", ""))
            qty_ordered = int(tokens[idx + 1].replace(",", ""))
            idx += 2
        else:
            qty_ordered = int(tokens[idx].replace(",", ""))
            idx += 1

    color_code = tokens[idx]; idx += 1

    # Everything between the colour code and the last two numbers is the size.
    size_tokens = tokens[idx:-2]
    if not size_tokens:
        raise ValueError(f"size not found in line: {line!r}")
    # "5", "10HUSKY", "8 HUSKY" -> existing husky rules
    size, consumed = normalize_size(size_tokens[0],
                                    size_tokens[1] if len(size_tokens) > 1 else None)
    extra = size_tokens[1 + (1 if consumed else 0):]
    # "REG"/"REGULAR" is the default fit: drop it so "M REG" matches size "M".
    # Other fit words (TALL, SLIM, PETITE ...) are kept so sizes stay distinct.
    extra = [t for t in extra if t.upper() not in DEFAULT_FIT_WORDS]
    if extra:
        size = f"{size} {' '.join(extra)}"
    unit_cost = float(tokens[-2].replace(',', ''))
    total_cost = float(tokens[-1].replace(',', ''))

    # ---- derived fields ----
    order_qty = qty_ordered or 0
    ship_qty = order_qty           # placeholder — refine when ship data exists
    ex_short = max(order_qty - ship_qty, 0)

    # UNIT / PREPACK
    unit_prepack = units_per_prepack if isinstance(units_per_prepack, int) else 1

    # FIX 3: carton count depends on how prepacks are split into cartons
    # (packing-list step), so it is not guessed here.
    ctn_qty = None
    carton_qty = None

    return {
        "BUYER": None,                    # filled from PO header later
        "STYLE": None,                    # filled from PO header later
        "P.O. #": None,                   # filled from PO header later
        "Dept": dept,
        "Style No": style_no,
        "Full Carton": full_carton,
        "PrePack": prepack,
        "PrePack Type": prepack_type,
        "PrePack Pack Factor": prepack_pack_factor,
        "SKU / ITEM": sku,
        "Units per PrePack": units_per_prepack,
        "UNIT / PREPACK": unit_prepack,
        "Style Description": style_description,
        "Item Carton Pack Factor": item_carton_pack_factor,
        "ORDER QTY": order_qty,
        "SHIP QTY": ship_qty,
        "EX-SHORT": ex_short,
        "CTN QTY": ctn_qty,
        "CARTON QTY": carton_qty,
        "#PrePackets Ordered": None,
        "Universal CC #Color Desc": color_code,
        "Size Desc": size,
        "Unit Cost": unit_cost,
        "Total Cost": total_cost,
    }


# ===========================================================================
# 5. Total row parser
# ===========================================================================
def parse_total(line: str):
    """
    5 tokens → Bulk total   : style filled, PrePack = null
    6 tokens → PrePack total: style = null, PrePack filled
    """
    tokens = line.split()
    if len(tokens) < 2 or not tokens[1].replace(",", "").isdigit():
        return None

    if len(tokens) == 5:
        return {
            "style": _to_int(tokens[1]),
            "PrePack": None,
            "PrePack Pack Factor": None,
            "ORDER QTY": _to_int(tokens[2]),
            "SHIP QTY": _to_int(tokens[2]),
            "EX-SHORT": 0,
            "CTN QTY": None,
            "CARTON QTY": None,
            "#PrePackets Ordered": None,
            "Universal CC #Color Desc": tokens[3],
            "Total Cost": float(tokens[4].replace(',', '')),
        }
    elif len(tokens) == 6:
        return {
            "style": None,
            "PrePack": _to_int(tokens[1]),
            "PrePack Pack Factor": _to_int(tokens[2]),
            "ORDER QTY": _to_int(tokens[3]),
            "SHIP QTY": _to_int(tokens[3]),
            "EX-SHORT": 0,
            "CTN QTY": None,
            "CARTON QTY": None,
            "#PrePackets Ordered": _to_int(tokens[4]),
            "Universal CC #Color Desc": None,
            "Total Cost": float(tokens[5].replace(',', '')),
        }
    return None


# ===========================================================================
# 6. Main parser
# ===========================================================================
def parse_text(text: str) -> dict:
    lines = text.splitlines()
    po_header = parse_po_header(text)

    tables = []
    current_rows = []
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        if ITEM_LINE_RE.match(line):
            row = parse_line_item(line)
            # look ahead for color name
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if is_color_name(nxt):
                    row["Universal CC #Color Desc"] = (
                        row["Universal CC #Color Desc"] + " " + nxt
                    )
                    i += 1
            current_rows.append(row)

        elif line.startswith("Total "):
            total = parse_total(line)
            if total is not None:
                # look ahead for color name (Bulk totals only)
                if i + 1 < len(lines):
                    nxt = lines[i + 1].strip()
                    if is_color_name(nxt) and total["Universal CC #Color Desc"] is not None:
                        total["Universal CC #Color Desc"] = (
                            total["Universal CC #Color Desc"] + " " + nxt
                        )
                        i += 1

                # propagate #PrePackets Ordered into Single/Multi rows
                for row in current_rows:
                    if row["PrePack Type"] in ("Single", "Multi"):
                        if total.get("#PrePackets Ordered") is not None:
                            row["#PrePackets Ordered"] = total["#PrePackets Ordered"]

                # inject PO header fields into every row
                for row in current_rows:
                    row["BUYER"] = po_header["BUYER"]
                    row["STYLE"] = po_header["STYLE"]
                    row["P.O. #"] = po_header["P.O._#"]

                tables.append({"rows": current_rows, "total": total})
                current_rows = []

        i += 1

    # Fallback: if OIQTY missing from header, sum from totals
    if not po_header["OIQTY"]:
        po_header["OIQTY"] = sum(t["total"]["ORDER QTY"] for t in tables)
    if not po_header["SHIP_QTY"]:
        po_header["SHIP_QTY"] = sum(t["total"]["SHIP QTY"] for t in tables)
    po_header["EX_SHORT"] = max(
        po_header["OIQTY"] - po_header["SHIP_QTY"], 0
    )

    return {
        "PO_HEADER": po_header,
        "columns": COLUMNS,
        "tables": {f"table {idx + 1}": t for idx, t in enumerate(tables)},
    }


# ===========================================================================
# 7. PDF → Text
# ===========================================================================
def extract_text(file_obj) -> str:
    with pdfplumber.open(file_obj) as pdf:
        return "\n".join(
            f"===== PAGE {i} =====\n{page.extract_text() or ''}\n"
            for i, page in enumerate(pdf.pages, start=1)
        )


# ===========================================================================
