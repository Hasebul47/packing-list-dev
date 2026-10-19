
"""
Color-grouped packing builders matching example layout:
  - Same STYLE → merge into one table per COLOR
  - Multi Y: one row per color (all sizes), CTN = #PrePackets
  - Single: one section per color, rows per prepack with carton splits (max 10 pp/ctn)
  - Continuous carton numbers across the PO
"""
from __future__ import annotations
import math
from collections import defaultdict

ALPHA_SIZES = ["XS", "S", "M", "L", "XL", "XXL", "XXL+"]
TODDLER_SIZES = ["12-18M", "18-24M", "2T", "3T", "4T", "5T", "6T"]
# Boys denim numeric + husky (from PO 61835339 / style 1183764)
BOYS_SIZES = ["5", "6", "7", "8", "10", "12", "14", "16", "18",
              "8H", "10H", "12H", "14H", "16H", "18H", "20H"]
SIZE_ORDER = TODDLER_SIZES + ALPHA_SIZES + BOYS_SIZES

DEFAULT_CTN_MEAS = [
    "58.67 X 38.48 X 29.71 CM.G8. SL",
    "58.67 X 38.48 X 14.86 CM.G8S. SL",
    "38.48 X 29.33 X 14.86 CM.G-8M",
    "29.33 X 19.25 X 14.86 CM.G-8XS",
    "55.88 X 38.1  X 15.24 CM G13",
]

# Max prepacks per carton for Single (matches example ~10)
SINGLE_MAX_PP_PER_CTN = 10

# Carton code by size family (from example images)
SIZE_CTN_CODE = {
    "XS": "G81", "S": "G82", "M": "G84", "L": "G81", "XL": "G81",
    "XXL": "G82", "XXL+": "G82",
    "12-18M": "G81", "18-24M": "G81", "2T": "G81", "3T": "G81",
    "4T": "G81", "5T": "G81", "6T": "G81",
    # Boys denim — matched to reference Excel (61835339)
    "5": "G85", "6": "G84", "7": "G84", "8": "G84", "10": "G82",
    "12": "G81", "14": "G81", "16": "G81", "18": "G84",
    "8H": "G85", "10H": "G85", "12H": "G84", "14H": "G84",
    "16H": "G84", "18H": "G84", "20H": "G84",
}

DEFAULT_NW = {
    "12-18M": 0.17, "18-24M": 0.19, "2T": 0.21, "3T": 0.22,
    "4T": 0.25, "5T": 0.26, "6T": 0.28,
    "XS": 0.510, "S": 0.520, "M": 0.550, "L": 0.560,
    "XL": 0.610, "XXL": 0.650, "XXL+": 0.690,
    # Boys denim (from reference packing list)
    "5": 0.37, "6": 0.40, "7": 0.42, "8": 0.45, "10": 0.48,
    "12": 0.52, "14": 0.58, "16": 0.60, "18": 0.63,
    "8H": 0.47, "10H": 0.51, "12H": 0.54, "14H": 0.60,
    "16H": 0.62, "18H": 0.65, "20H": 0.67,
}
DEFAULT_EMPTY = {
    "XS": 0.49, "S": 1.00, "M": 1.00, "L": 1.00,
    "XL": 1.00, "XXL": 1.00, "XXL+": 1.00,
    "12-18M": 0.71, "18-24M": 0.71, "2T": 0.71, "3T": 0.71,
    "4T": 0.71, "5T": 0.71, "6T": 0.71,
    # Boys — empty carton by code family (approx from reference net/gross)
    "5": 0.42, "6": 0.24, "7": 0.24, "8": 0.24, "10": 0.24,
    "12": 0.71, "14": 0.71, "16": 0.71, "18": 0.24,
    "8H": 0.42, "10H": 0.42, "12H": 0.24, "14H": 0.24,
    "16H": 0.24, "18H": 0.24, "20H": 0.24,
}


# ---------------------------------------------------------------------------
# Packing rules derived from the reference packing lists (POs 61480360/61429740)
# ---------------------------------------------------------------------------
# Max PCS per carton for Bulk (full cartons), by size. Alpha values match the
# reference exactly; 12-18M / 6T never fill a carton in the samples (assumed).
BULK_CTN_CAP = {
    "XS": 20, "S": 20, "M": 20, "L": 20, "XL": 18, "XXL": 18, "XXL+": 15,
    "12-18M": 60, "18-24M": 60, "2T": 50, "3T": 50, "4T": 44, "5T": 44, "6T": 44,
}
DEFAULT_CTN_CAP = 20          # sizes not listed above (e.g. boys numeric) - check!
DEFAULT_EXCESS_PCT = 3.0      # max over-shipment per size, added to Bulk only
CARTON_TARE = {"G81": 0.71, "G82": 0.71, "G84": 0.24, "G85": 0.24}
NET_DEDUCT = {"Bulk": 0.24, "Single": 0.24, "Multi Y": 0.42, "Multi N": 0.42}


def carton_code(size: str, qty: int) -> str:
    """Pick the carton type from how full the carton is (reference behaviour)."""
    if size in TODDLER_SIZES:
        if qty >= 40:
            return "G81"
        if qty >= 22:
            return "G82"
        if qty >= 9:
            return "G84"
        return "G85"
    net = qty * DEFAULT_NW.get(size, 0.5)
    if net >= 8.5:
        return "G81"
    if net >= 4.5:
        return "G82"
    if net >= 2.5:
        return "G84"
    return "G85"


def carton_weights(size: str, qty: int, code: str, pack_type: str) -> tuple[float, float]:
    gross = round(qty * DEFAULT_NW.get(size, 0.5) + CARTON_TARE.get(code, 0.71), 2)
    net = round(gross - NET_DEDUCT.get(pack_type, 0.24), 2)
    return gross, net


def _int_or(v, default=0) -> int:
    if v is None:
        return default
    try:
        return int(float(str(v).strip().replace(",", "").split()[0]))
    except (TypeError, ValueError, IndexError):
        return default


def color_display(raw: str) -> str:
    """'000905518-002 NAVY CAPTAIN' -> 'NAVY CAPTAIN'"""
    if not raw:
        return ""
    parts = str(raw).split()
    if parts and any(c.isdigit() for c in parts[0]):
        return " ".join(parts[1:]) if len(parts) > 1 else parts[0]
    return str(raw)


def color_code(raw: str) -> str:
    if not raw:
        return "000"
    first = raw.split()[0] if raw.split() else ""
    if "-" in first:
        return first.split("-")[-1]
    return "000"


def sort_sizes(sizes) -> list:
    known = [s for s in SIZE_ORDER if s in sizes]
    other = [s for s in sizes if s not in SIZE_ORDER]
    return known + other


def detect_table_type(table: dict) -> str | None:
    rows = table.get("rows") or []
    if not rows:
        return None
    first = rows[0]
    ptype = (first.get("PrePack Type") or "").strip()
    fc = (first.get("Full Carton") or "").strip().upper()
    if ptype == "Bulk":
        return "Bulk"
    if ptype == "Single":
        return "Single"
    if ptype == "Multi":
        if fc == "Y":
            return "Multi Y"
        if fc == "N":
            return "Multi N"
        return "Multi Y"
    return None


def split_prepacks(n: int, max_per: int = SINGLE_MAX_PP_PER_CTN) -> list[tuple[int, int]]:
    """Return list of (num_cartons, prepacks_in_each) covering n prepacks."""
    if n <= 0:
        return []
    if n <= max_per:
        return [(1, n)]
    full = n // max_per
    rem = n % max_per
    out = []
    if full:
        out.append((full, max_per))
    if rem:
        out.append((1, rem))
    return out


def split_prepacks_merge(n: int, max_per: int) -> list[tuple[int, int]]:
    """Like split_prepacks, but a lone leftover prepack rides in the last full
    carton (reference: 101 prepacks @10 -> 9x10 + 1x11, not 10x10 + 1x1)."""
    parts = split_prepacks(n, max_per)
    if len(parts) == 2 and parts[1] == (1, 1) and parts[0][0] >= 1:
        full, per = parts[0]
        out = [(full - 1, per)] if full > 1 else []
        out.append((1, per + 1))
        return out
    return parts


def bulk_excess(data: dict, pct: float = DEFAULT_EXCESS_PCT) -> dict:
    """{color: {size: extra_pcs}} = floor(pct% of that size's ORDER qty over ALL
    pack types of the colour). The extra pieces are shipped inside Bulk."""
    tot: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for table in (data.get("tables") or {}).values():
        for r in table.get("rows") or []:
            color = color_display(r.get("Universal CC #Color Desc", "")) or "UNKNOWN"
            tot[color][str(r.get("Size Desc") or "")] += _int_or(r.get("ORDER QTY"))
    return {c: {s: int(q * pct / 100 + 1e-9) for s, q in sizes.items()}
            for c, sizes in tot.items()}


def style_full(base_style: str, color_cd: str) -> str:
    """905518 + 000 -> 905518-00-1  (example format)."""
    base = str(base_style or "").split("-")[0]
    # color code like 000 / 002 -> 00 / 02
    cd = str(color_cd or "000").zfill(3)
    short = cd[-2:] if len(cd) >= 2 else cd
    return f"{base}-{short}-1"


# ---------------------------------------------------------------------------
# MULTI Y  — one section per COLOR (same style merged)
# ---------------------------------------------------------------------------
def build_multi_y_by_color(data: dict, start_ctn: int = 1) -> list[dict]:
    header = data.get("PO_HEADER") or {}
    base_style = str(header.get("STYLE") or "")
    # group tables by color
    by_color: dict[str, list] = defaultdict(list)
    for tname, table in (data.get("tables") or {}).items():
        if detect_table_type(table) != "Multi Y":
            continue
        rows = table.get("rows") or []
        if not rows:
            continue
        first = rows[0]
        color = color_display(first.get("Universal CC #Color Desc", "")) or "UNKNOWN"
        by_color[color].append((tname, table))

    sections = []
    next_ctn = start_ctn
    for color, tables in by_color.items():
        # Merge all Multi-Y tables of this color (same style)
        units: dict[str, int] = {}
        prepack = ""
        color_raw = ""
        total_order = 0
        total_ship = 0
        total_pp = 0
        for _, table in tables:
            trows = table["rows"]
            tot = table.get("total") or {}
            first = trows[0]
            prepack = str(first.get("PrePack") or prepack)
            color_raw = first.get("Universal CC #Color Desc") or color_raw
            for r in trows:
                sz = r.get("Size Desc") or ""
                units[sz] = _int_or(r.get("Units per PrePack"), 1)
            total_order += _int_or(tot.get("ORDER QTY"))
            total_ship += _int_or(tot.get("SHIP QTY") or tot.get("ORDER QTY"))
            total_pp += _int_or(tot.get("#PrePackets Ordered") or first.get("#PrePackets Ordered"))

        sizes = sort_sizes(list(units.keys()))
        upp = sum(units.values()) or 1
        if total_pp <= 0 and upp:
            total_pp = math.ceil(total_ship / upp)
        if total_ship <= 0:
            total_ship = total_pp * upp
        if total_order <= 0:
            total_order = total_ship

        ctn_qty = total_pp  # Multi Y: Full Carton = Y → 1 prepack = 1 carton
        cd = color_code(color_raw)
        ctn_code = "G82"
        rows_out = [{
            "ctn_from": next_ctn,
            "ctn_to": next_ctn + ctn_qty - 1 if ctn_qty else next_ctn,
            "ctn_qty": ctn_qty,
            "ctn_mes": ctn_code,
            "color": color,
            "prepack": prepack,
            "size_vals": dict(units),  # size -> units in prepack
            "per_blst": f"{upp} X 1",
            "qty_per_ctn": upp,
            "total_qty": total_ship,
            "gross": round(sum(DEFAULT_NW.get(s, 0.5) * units.get(s, 0) for s in sizes)
                           + CARTON_TARE[ctn_code], 2),
            "net": round(sum(DEFAULT_NW.get(s, 0.5) * units.get(s, 0) for s in sizes)
                         + CARTON_TARE[ctn_code] - NET_DEDUCT["Multi Y"], 2),
        }]
        next_ctn += ctn_qty

        sections.append({
            "pack_type": "Multi Y",
            "color": color,
            "color_code": cd,
            "prepack": prepack,
            "style_no": base_style,
            "style_full": style_full(base_style, cd),
            "sizes": sizes,
            "units": units,
            "units_per_prepack": upp,
            "qty_ordered": total_order,
            "ship_qty": total_ship,
            "prepackets": total_pp,
            "ctn_total": ctn_qty,
            "rows": rows_out,
            "ctn_meas_list": list(DEFAULT_CTN_MEAS),
            "carton_label": f"{rows_out[0]['ctn_from']:02d} of ",  # filled later
        })
    return sections


# ---------------------------------------------------------------------------
# SINGLE  — one section per COLOR, all prepacks merged into one table
# ---------------------------------------------------------------------------
def build_single_by_color(data: dict, start_ctn: int = 1,
                          max_pp: int | None = None,
                          caps: dict | None = None) -> list[dict]:
    """max_pp: force one prepacks-per-carton limit for every size (None/0 = auto,
    i.e. the per-size pcs cap in BULK_CTN_CAP divided by units per prepack)."""
    caps = {**BULK_CTN_CAP, **(caps or {})}
    header = data.get("PO_HEADER") or {}
    base_style = str(header.get("STYLE") or "")
    by_color: dict[str, list] = defaultdict(list)

    for tname, table in (data.get("tables") or {}).items():
        if detect_table_type(table) != "Single":
            continue
        rows = table.get("rows") or []
        if not rows:
            continue
        color = color_display(rows[0].get("Universal CC #Color Desc", "")) or "UNKNOWN"
        by_color[color].append((tname, table))

    sections = []
    next_ctn = start_ctn
    for color, tables in by_color.items():
        rows_out = []
        total_order = total_ship = total_pp = 0
        sizes_seen = set()
        color_raw = first_prepack = ""
        sec_upp = None

        for _, table in tables:
            trows = table["rows"]
            tot = table.get("total") or {}
            first = trows[0]
            prepack = str(first.get("PrePack") or "")
            first_prepack = first_prepack or prepack
            color_raw = first.get("Universal CC #Color Desc") or color_raw
            size = first.get("Size Desc") or ""
            sizes_seen.add(size)
            upp = _int_or(first.get("Units per PrePack") or first.get("UNIT / PREPACK"), 2) or 2
            sec_upp = sec_upp or upp
            order = _int_or(tot.get("ORDER QTY") or first.get("ORDER QTY"))
            ship = _int_or(tot.get("SHIP QTY") or order)
            pp = _int_or(tot.get("#PrePackets Ordered") or first.get("#PrePackets Ordered"))
            if pp <= 0 and upp:
                pp = math.ceil(order / upp)
            total_order += order
            total_ship += ship
            total_pp += pp

            cap_pp = int(max_pp) if max_pp else max(1, caps.get(size, DEFAULT_CTN_CAP) // upp)
            for n_ctn, pp_in_ctn in split_prepacks_merge(pp, cap_pp):
                qty_per = pp_in_ctn * upp
                code = carton_code(size, qty_per)
                gross, net = carton_weights(size, qty_per, code, "Single")
                rows_out.append({
                    "ctn_from": next_ctn,
                    "ctn_to": next_ctn + n_ctn - 1,
                    "ctn_qty": n_ctn,
                    "ctn_mes": code,
                    "color": "",
                    "prepack": prepack,
                    "size": size,
                    "size_val": pp_in_ctn,  # shown in size column (prepacks in ctn)
                    "per_blst": "",
                    "qty_per_ctn": qty_per,
                    "total_qty": qty_per * n_ctn,
                    "gross": gross,
                    "net": net,
                })
                next_ctn += n_ctn

        if rows_out:
            rows_out[0]["color"] = color
        cd = color_code(color_raw)
        sections.append({
            "pack_type": "Single",
            "color": color,
            "color_code": cd,
            "prepack": first_prepack,
            "style_no": base_style,
            "style_full": style_full(base_style, cd),
            "sizes": sort_sizes(list(sizes_seen)),
            "units": {},
            "units_per_prepack": sec_upp or 2,
            "qty_ordered": total_order,
            "ship_qty": total_ship,
            "prepackets": total_pp,
            "ctn_total": sum(r["ctn_qty"] for r in rows_out),
            "rows": rows_out,
            "ctn_meas_list": list(DEFAULT_CTN_MEAS),
        })
    return sections


# ---------------------------------------------------------------------------
# BULK — one section per COLOR
# ---------------------------------------------------------------------------
def build_bulk_by_color(data: dict, start_ctn: int = 1,
                        caps: dict | None = None,
                        excess_pct: float = DEFAULT_EXCESS_PCT) -> list[dict]:
    """
    Bulk packing, one section per COLOR, one block of rows per SIZE:
      ship qty = order qty + excess   (excess = floor(excess_pct% of the size's
                 order over all pack types; see bulk_excess)
      cartons  = full cartons of `caps[size]` pcs + one remainder carton
    (reference: Style 905518 / PO 61480360, Teakwood = 80 cartons, 1,505 pcs).
    """
    caps = {**BULK_CTN_CAP, **(caps or {})}
    header = data.get("PO_HEADER") or {}
    base_style = str(header.get("STYLE") or "")
    excess = bulk_excess(data, excess_pct)
    by_color: dict[str, list] = defaultdict(list)

    for tname, table in (data.get("tables") or {}).items():
        if detect_table_type(table) != "Bulk":
            continue
        rows = table.get("rows") or []
        if not rows:
            continue
        color = color_display(rows[0].get("Universal CC #Color Desc", "")) or "UNKNOWN"
        by_color[color].append((tname, table))

    sections = []
    next_ctn = start_ctn
    for color, tables in by_color.items():
        order_by_size: dict[str, int] = {}
        color_raw = ""
        for _, table in tables:
            color_raw = table["rows"][0].get("Universal CC #Color Desc") or color_raw
            for r in table["rows"]:
                sz = str(r.get("Size Desc") or "")
                order_by_size[sz] = order_by_size.get(sz, 0) + _int_or(
                    r.get("ORDER QTY") or r.get("SHIP QTY"))

        sizes = sort_sizes([s for s, q in order_by_size.items() if q > 0])
        cd = color_code(color_raw)
        rows_out = []
        ship_by_size, ex_by_size = {}, {}
        for sz in sizes:
            cap = int(caps.get(sz, DEFAULT_CTN_CAP))
            order = order_by_size[sz]
            ex = excess.get(color, {}).get(sz, 0)
            # Don't let the excess alone spill a tiny extra carton: if the order
            # already fits in full cartons and the spill is only a few pcs,
            # trim the excess back to the last full carton (reference: 4T 176, 5T 132).
            spill = (order + ex) % cap
            # (reference keeps 2-pc spill cartons on alpha sizes, so only cap >= 40)
            if cap >= 40 and 0 < spill <= math.ceil(cap * 0.1) and order <= (order + ex) - spill:
                ex -= spill
            ship = order + ex
            ship_by_size[sz], ex_by_size[sz] = ship, ex
            for n_ctn, per in split_prepacks(ship, cap):
                code = carton_code(sz, per)
                gross, net = carton_weights(sz, per, code, "Bulk")
                rows_out.append({
                    "ctn_from": next_ctn,
                    "ctn_to": next_ctn + n_ctn - 1,
                    "ctn_qty": n_ctn,
                    "ctn_mes": code,
                    "color": "",
                    "prepack": "",
                    "size": sz,
                    "size_val": per,          # pcs per carton, shown in size column
                    "size_vals": {sz: per},
                    "per_blst": "",
                    "qty_per_ctn": per,
                    "total_qty": per * n_ctn,
                    "gross": gross,
                    "net": net,
                })
                next_ctn += n_ctn
        if rows_out:
            rows_out[0]["color"] = color

        total_order = sum(order_by_size[s] for s in sizes)
        sections.append({
            "pack_type": "Bulk",
            "color": color,
            "color_code": cd,
            "prepack": "Bulk",
            "style_no": base_style,
            "style_full": style_full(base_style, cd),
            "sizes": sizes,
            "units": ship_by_size,
            "excess_by_size": ex_by_size,
            "units_per_prepack": 1,
            "qty_ordered": total_order,
            "ship_qty": sum(ship_by_size.values()),
            "prepackets": 0,
            "ctn_total": sum(r["ctn_qty"] for r in rows_out),
            "rows": rows_out,
            "ctn_meas_list": list(DEFAULT_CTN_MEAS),
        })
    return sections


def section_to_display_rows(sec: dict) -> list[dict]:
    """Flat rows for dataframe display matching example columns."""
    pack = sec["pack_type"]
    sizes = sec.get("sizes") or []
    recs = []
    for r in sec["rows"]:
        rec = {
            "CTN NO": (
                f"{r['ctn_from']}-{r['ctn_to']}"
                if r.get("ctn_to") and r["ctn_to"] != r["ctn_from"]
                else str(r["ctn_from"])
            ),
            "CTN MES": r.get("ctn_mes", ""),
            "CTN QTY": r["ctn_qty"],
            "COLOR": r.get("color") or sec.get("color", ""),
            "PREPACK STECKER": r.get("prepack") or sec.get("prepack", ""),
        }
        if pack in ("Single", "Bulk"):
            # one size column filled per row
            for s in sizes:
                rec[s] = r.get("size_val", "") if s == r.get("size") else ""
        else:
            size_vals = r.get("size_vals") or sec.get("units") or {}
            for s in sizes:
                rec[s] = size_vals.get(s, "")
        rec["PER BLST"] = r.get("per_blst", "")
        rec["QTY PER CTN"] = r.get("qty_per_ctn", "")
        rec["TOTAL QTY"] = r.get("total_qty", "")
        rec["GROSS WEIGHT"] = r.get("gross", "")
        rec["NET WEIGHT"] = r.get("net", "")
        recs.append(rec)

    # TOTAL row
    tot = {
        "CTN NO": "TOTAL",
        "CTN MES": "",
        "CTN QTY": f"{sec['ctn_total']} CTNS",
        "COLOR": "",
        "PREPACK STECKER": "",
        "PER BLST": "",
        "QTY PER CTN": "",
        "TOTAL QTY": f"{sec['ship_qty']} PCS",
        "GROSS WEIGHT": "",
        "NET WEIGHT": "",
    }
    if pack in ("Single", "Bulk"):
        for s in sizes:
            tot[s] = sum(
                (r.get("size_val") or 0) * r["ctn_qty"]
                for r in sec["rows"] if r.get("size") == s
            )
    else:
        for s in sizes:
            tot[s] = (sec.get("units") or {}).get(s, "")
    recs.append(tot)
    return recs
