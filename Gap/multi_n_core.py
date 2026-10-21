"""
multi_n_core.py
---------------
Pure-python logic for MULTI PACK ( N ) packing lists (no Streamlit import, so it
can be shared by multi_n_packing.py, single_pack.py and bulk_pack.py).

What is a "Multi N" table?
--------------------------
In the parsed PO JSON a table whose first row has
        PrePack Type == "Multi"   and   Full Carton == "N"
is a MULTI PACK ( N ) prepack: the prepack is NOT a full carton, so several
prepacks share one carton (Multi Y = exactly 1 prepack per carton).

Example (PO 61429740, prepack 8655475, 247 prepacks, 9 pcs/prepack):
    prepacks per carton = 4
    247 = 61 cartons x 4  +  1 carton x 3
        -> row 1 : 61 CTNS, PER BLST = 9 X 4, QTY/CTN = 36, TOTAL = 2196
        -> row 2 :  1 CTN , PER BLST = 9 X 3, QTY/CTN = 27, TOTAL =   27
                                                            ---------
                                                 TOTAL QTY  =  2223 PCS
"""

from __future__ import annotations

import json
from pathlib import Path

# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------
TODDLER_SIZES = ["12-18M", "18-24M", "2T", "3T", "4T", "5T", "6T"]
ALPHA_SIZES   = ["XS", "S", "M", "L", "XL", "XXL", "XXL+"]
SIZE_ORDER    = TODDLER_SIZES + ALPHA_SIZES

# toddler weights  = old_navy_size_weight_values.json  (N.WT)
# N.N.W            = N.W - 0.02   (same rule the reference sheet uses)
DEFAULT_NW = {
    "12-18M": 0.17, "18-24M": 0.19, "2T": 0.21, "3T": 0.22,
    "4T": 0.25, "5T": 0.26, "6T": 0.28,
    "XS": 0.510, "S": 0.520, "M": 0.550, "L": 0.560,
    "XL": 0.610, "XXL": 0.650, "XXL+": 0.690,
}
DEFAULT_NNW = {
    **{s: round(DEFAULT_NW[s] - 0.02, 3) for s in TODDLER_SIZES},
    "XS": 0.490, "S": 0.500, "M": 0.530, "L": 0.540,
    "XL": 0.590, "XXL": 0.630, "XXL+": 0.670,
}

# Empty-carton weight per carton code.
# G81 = 0.71 is the value used by the Multi-N reference sheet
# ("Style - 908546 PO - 61429740 ... Multi.xlsx"); the others come from
# dummy_61480360_packing_list.xlsx.  All editable in the UI / sheet.
DEFAULT_TARE = {"G81": 0.71, "G82": 0.71, "G84": 0.24, "G85": 0.24}
DEFAULT_NET_DEDUCT = 0.42           # GROSS - NET for MULTI cartons
DEFAULT_CTN_CODE = "G81"            # carton code used in the Multi-N reference
DEFAULT_PREPACKS_PER_CTN = 4        # reference: 61 x 4 + 1 x 3 = 247

DEFAULT_CTN_MEAS = [
    "58.67 X 38.48 X 29.71 CM.G8. SL",
    "58.67 X 38.48 X 14.86 CM.G8S. SL",
    "38.48 X 29.33 X 14.86 CM.G-8M",
    "29.33 X 19.25 X 14.86 CM.G-8XS",
    "55.88 X 38.1  X 15.24 CM G13",
]


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def load_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def color_display(raw: str) -> str:
    """'000908546-001 ESPRESSO BARK' -> 'ESPRESSO BARK'."""
    if not raw:
        return ""
    parts = raw.split(" ", 1)
    return parts[1].strip() if len(parts) > 1 else raw


def color_code(raw: str) -> str:
    """'000908546-001 ESPRESSO BARK' -> '001'."""
    if not raw:
        return "000"
    first = raw.split()[0]
    return first.split("-")[-1] if "-" in first else "000"


def int_or_none(v):
    if v is None:
        return None
    try:
        return int(float(str(v).replace(",", "").strip().split()[0]))
    except (TypeError, ValueError, IndexError):
        return None


def is_multi_n(table: dict) -> bool:
    rows = table.get("rows") or []
    if not rows:
        return False
    first = rows[0]
    return first.get("PrePack Type") == "Multi" and first.get("Full Carton") == "N"


def sort_sizes(sizes) -> list:
    known = [s for s in SIZE_ORDER if s in sizes]
    other = [s for s in sizes if s not in SIZE_ORDER]
    return known + other


def size_family(sizes) -> list:
    """The 7 size columns shown on the sheet (reference always shows 7)."""
    if any(s in TODDLER_SIZES for s in sizes):
        return list(TODDLER_SIZES)
    if any(s in ALPHA_SIZES for s in sizes):
        return list(ALPHA_SIZES)
    return sort_sizes(sizes)


def split_cartons(prepackets: int, per_ctn: int) -> list[tuple[int, int]]:
    """
    247 prepacks, 4 per carton -> [(61, 4), (1, 3)]
    Returns a list of (ctn_qty, prepacks_per_ctn).
    """
    per_ctn = max(int(per_ctn), 1)
    prepackets = max(int(prepackets), 0)
    full, rem = divmod(prepackets, per_ctn)
    out = []
    if full:
        out.append((full, per_ctn))
    if rem:
        out.append((1, rem))
    return out


def carton_weights(units: dict, nw: dict, prepacks: int, tare: float,
                   net_deduct: float) -> tuple[float, float]:
    """Gross = sum(units x N.W) x prepacks + empty carton ; Net = Gross - deduct."""
    pcs_weight = sum(u * nw.get(s, 0.0) for s, u in units.items())
    gross = round(pcs_weight * prepacks + tare, 2)
    net = round(gross - net_deduct, 2)
    return gross, net


# ---------------------------------------------------------------------------
# Main builder
# ---------------------------------------------------------------------------
def build_multi_n(
    data: dict,
    nw: dict,
    tare_by_code: dict,
    net_deduct: float = DEFAULT_NET_DEDUCT,
    default_per_ctn: int = DEFAULT_PREPACKS_PER_CTN,
    per_ctn_overrides: dict | None = None,   # {prepack_no: prepacks per carton}
    ctn_code: str = DEFAULT_CTN_CODE,
    start_ctn_no: int = 1,
) -> list[dict]:
    """
    One section per Multi-N table.  Carton numbers run continuously
    across sections (start_ctn_no, ...).
    """
    per_ctn_overrides = per_ctn_overrides or {}
    po_header = data.get("PO_HEADER") or {}
    sections: list[dict] = []
    next_no = start_ctn_no

    for tname, table in (data.get("tables") or {}).items():
        if not is_multi_n(table):
            continue
        trows = table["rows"]
        total = table.get("total") or {}
        first = trows[0]

        prepack = str(first.get("PrePack"))
        per_ctn = int(per_ctn_overrides.get(prepack, default_per_ctn))

        units = {r["Size Desc"]: int(r.get("Units per PrePack") or 0) for r in trows}
        sizes = sort_sizes(list(units))
        units_per_prepack = sum(units.values())

        prepackets = (
            int_or_none(total.get("#PrePackets Ordered"))
            or int_or_none(first.get("#PrePackets Ordered"))
            or 0
        )
        qty_ordered = (
            int_or_none(total.get("Qty Ordered (eaches)"))
            or int_or_none(total.get("ORDER QTY"))
            or sum(int_or_none(r.get("Qty Ordered (eaches)") or r.get("ORDER QTY")) or 0
                   for r in trows)
            or prepackets * units_per_prepack
        )

        ccode = color_code(first.get("Universal CC #Color Desc", ""))
        style_no = str(first.get("Style No") or po_header.get("STYLE") or "")
        tare = float(tare_by_code.get(ctn_code, 0.0))

        rows = []
        for ctn_qty, p in split_cartons(prepackets, per_ctn):
            gross, net = carton_weights(units, nw, p, tare, net_deduct)
            qty_per_ctn = units_per_prepack * p
            rows.append({
                "ctn_from": next_no,
                "ctn_to": next_no + ctn_qty - 1,
                "ctn_mes": ctn_code,
                "ctn_qty": ctn_qty,
                "prepacks_per_ctn": p,
                "per_blst": f"{units_per_prepack} X {p}",
                "qty_per_ctn": qty_per_ctn,
                "total_qty": qty_per_ctn * ctn_qty,
                "gross": gross,
                "net": net,
            })
            next_no += ctn_qty

        sections.append({
            "pack_type": "Multi N",
            "table": tname,
            "color": color_display(first.get("Universal CC #Color Desc", "")) or "UNKNOWN",
            "color_code": ccode,
            "style_no": style_no,
            "style_full": f"{style_no}-{ccode[-2:]}-1",
            "prepack": prepack,
            "sizes": sizes,
            "units": units,
            "units_per_prepack": units_per_prepack,
            "prepackets": prepackets,
            "per_ctn": per_ctn,
            "qty_ordered": qty_ordered,
            "ship_qty": sum(r["total_qty"] for r in rows),
            "ctn_total": sum(r["ctn_qty"] for r in rows),
            "rows": rows,
        })

    return sections


def multi_n_carton_total(data: dict,
                         default_per_ctn: int = DEFAULT_PREPACKS_PER_CTN,
                         per_ctn_overrides: dict | None = None) -> int:
    """Total cartons used by all Multi-N tables (for 'NN of TTT' numbering)."""
    per_ctn_overrides = per_ctn_overrides or {}
    total = 0
    for _, table in (data.get("tables") or {}).items():
        if not is_multi_n(table):
            continue
        first = table["rows"][0]
        t = table.get("total") or {}
        prepackets = (int_or_none(t.get("#PrePackets Ordered"))
                      or int_or_none(first.get("#PrePackets Ordered")) or 0)
        p = int(per_ctn_overrides.get(str(first.get("PrePack")), default_per_ctn))
        total += sum(q for q, _ in split_cartons(prepackets, p))
    return total
