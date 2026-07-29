#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate.py — Customer order -> Định mức + YCSX Excel files.

ARCHITECTURE (hybrid, byte-faithful):
  * Clone a template .xlsx and fill ONLY the input cells (cell_map.json).
  * Use xlsxpatch.XlsxPatch so embedded images/drawings/styles are preserved
    (openpyxl would drop them). Output opens identically to the examples.
  * The ONE interactive question is "số màu in?" (number of print colors) —
    every other value is parsed from the order or computed from formulas.json.

USAGE:
  python generate.py --order order.json            # fill from a JSON order
  python generate.py --ycsx path/to/YCSX.xlsx      # parse fields from a YCSX file
  python generate.py --sample                      # run the built-in 40KG sample
  python generate.py --order order.json --colors 3 # pre-answer the only question
"""
from __future__ import annotations
import argparse, json, os, sys, datetime, re, copy, zipfile
from xlsxpatch import XlsxPatch

try:  # Windows console defaults to cp1252; force utf-8 so Vietnamese prints survive
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
TEMPLATES = os.path.join(HERE, "templates")
OUTDIR = os.path.join(HERE, "output")


def template_path(key):
    return os.path.join(TEMPLATES, {
        "opp": "Định mức - OPP (Bao BOPP in ống đồng).xlsx",
        "paper_kp": "Định mức - Bao giấy (KP - in offset).xlsx",
        "ycsx": "YCSX.xlsx",
    }[key])


def load_json(name):
    with open(os.path.join(DATA, name), encoding="utf-8") as f:
        return json.load(f)


CONST = load_json("constants.json")
CELLMAP = load_json("cell_map.json")
INK = load_json("ink_master.json")
RULES = load_json("customer_rules.json")
BAG = load_json("bag_type_map.json")
DIMS = load_json("standard_dims.json")


# ---------------------------------------------------------------- parse order
def parse_ycsx(path):
    """Parse a YCSX .xlsx (read-only): header + ALL product line items.

    Merge-aware: a cell covered by a merged range returns that range's top-left
    value, so a shared spec such as F13:H15 applies to every product row beneath
    it. Returns an order dict with header fields, a ``products`` list (one entry
    per product row that carries a qty), and top-level single-product fields set
    to the first product (backward-compat with --order / --sample callers).
    """
    from openpyxl import load_workbook
    from openpyxl.utils import range_boundaries, coordinate_to_tuple
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    def cellv(coord):
        """Value at coord, falling back to the top-left of a containing merge."""
        val = ws[coord].value
        if val is not None:
            return val
        cr, cc = coordinate_to_tuple(coord)
        for mr in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(mr))
            if min_row <= cr <= max_row and min_col <= cc <= max_col:
                return ws.cell(row=min_row, column=min_col).value
        return None

    # header labels live in column B as "N. Label: value"
    g = {c.coordinate: c.value.strip()
         for row in ws.iter_rows(values_only=False)
         for c in row if isinstance(c.value, str)}

    def bval(label):
        for val in g.values():
            if val and val.startswith(label):
                return val.split(":", 1)[-1].strip() if ":" in val else ""
        return ""

    order = {
        "customer": bval("2. Khách hàng") or g.get("B7", ""),
        "address": bval("3. Địa chỉ"),
        "customer_code": bval("4. Mã khách hàng"),
        "order_id": bval("5. Số đơn hàng"),
        "ngay_yc": bval("1. Ngày yêu cầu"),
    }

    # product rows start at 13: include any row with a numeric qty (col J) and a
    # code/name. Stage-marker rows (MÀNH/IN/TRÁNG/...) carry no qty -> skipped.
    products = []
    for r in range(13, 41):
        code, name = cellv(f"C{r}"), cellv(f"D{r}")
        qty = _to_num(cellv(f"J{r}"))
        if qty is None or not (code or name):
            continue
        products.append({
            "product_code": str(code).strip() if code else "",
            "product_name": str(name).strip() if name else "",
            "ma_code": str(cellv(f"E{r}") or "").strip(),
            "qty": qty,
            "spec": str(cellv(f"F{r}") or "").strip().strip('"'),
        })
    if not products:  # legacy fallback: take row 13 verbatim
        products = [{
            "product_code": str(cellv("C13") or "").strip(),
            "product_name": str(cellv("D13") or "").strip(),
            "ma_code": str(cellv("E13") or "").strip(),
            "qty": _to_num(cellv("J13")) or 0,
            "spec": str(cellv("F13") or "").strip().strip('"'),
        }]
    order["products"] = products

    # top-level single-product fields = first product (backward compat)
    p0 = products[0]
    order.update({
        "product_code": p0["product_code"], "product_name": p0["product_name"],
        "ma_code": p0["ma_code"], "qty": p0["qty"], "spec": p0["spec"],
    })
    order.update(parse_spec(p0["spec"]))
    return order


def parse_spec(spec):
    """Extract dimensions / structure / bag hints from the Chi tiết kỹ thuật text."""
    out = {}
    if not spec:
        return out
    s = spec.lower()
    # dimensions like "(42+8) cm x 82cm" (optional parens) or "50x92 cm"
    m = re.search(r"\(?\s*(\d+)\s*\+\s*(\d+)\s*\)?\s*cm\s*x\s*(\d+(?:\.\d+)?)\s*cm", spec)
    if m:
        out["width_cm"] = int(m.group(1)); out["gusset_cm"] = int(m.group(2))
        out["width_plus_gusset_m"] = (int(m.group(1)) + int(m.group(2))) / 100
        out["bag_length_m"] = float(m.group(3)) / 100
    else:
        m = re.search(r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*cm", spec)
        if m:
            out["width_plus_gusset_m"] = float(m.group(1)) / 100
            out["bag_length_m"] = float(m.group(2)) / 100
    return out


def _to_num(v):
    if v is None or v == "":
        return None
    try:
        return int(v)
    except Exception:
        try:
            return float(v)
        except Exception:
            return None


# ------------------------------------------------------------ detect bag family
def detect_family(order):
    text = " ".join([
        str(order.get("customer", "")),
        str(order.get("product_name", "")),
        str(order.get("spec", "")),
    ]).lower()
    opp = BAG["families"]["opp"]
    kp = BAG["families"]["paper_kp"]
    if any(k.lower() in text for k in opp["name_keywords"] + opp["spec_keywords"]):
        return "opp"
    if any(k.lower() in text for k in kp["name_keywords"] + kp["spec_keywords"]):
        return "paper_kp"
    # fallback by weight token
    mw = re.search(r"(\d{2,3})\s*kgs?", text)
    if mw:
        w = int(mw.group(1))
        return "opp" if w >= 35 else "paper_kp"
    raise ValueError("Cannot detect bag family from order. Set order['bag_family'] manually.")


# ----------------------------------------------------------- customer overrides
def apply_rules(order, family):
    text = (str(order.get("customer", "")) + " " + str(order.get("spec", ""))).lower()
    ov = {}
    for r in RULES["rules"]:
        m = r["match"]
        if "customer_contains_any" in m and any(k.lower() in text for k in m["customer_contains_any"]):
            ov.update(r["effect"])
        if "spec_contains_any" in m and any(k.lower() in text for k in m["spec_contains_any"]):
            ov.update(r["effect"])
        if m.get("bag_family") == family:
            ov.update(r["effect"])
    return ov


# -------------------------------------------------------------------- compute
def compute(order, family, so_mau_in):
    ov = apply_rules(order, family)
    qty = float(order.get("qty") or 0)
    L = float(order.get("bag_length_m") or 0)
    W = float(order.get("width_plus_gusset_m") or 0)
    ibw = float(order.get("inner_bag_weight_kg") or 0)
    dung_sai = float(order.get("tolerance", CONST["dung_sai_default"]))

    kho_manh = round(W * 2 + 0.06, 3)
    kho_giay = round(W * 2 + 0.02, 3)
    kho_mang = float(order.get("kho_mang") or _dims_lookup(order, "mang_in_cm"))
    dl_manh = ov.get("dinh_luong_manh_trang_gm2", CONST["dinh_luong_manh_trang_gm2"])

    # Per-family conventions (verified against both examples):
    #   OPP   : sl_in = qty*L*(1+dung_sai) + phế chồng màu ; thành phẩm = qty
    #   paper : sl_in = qty*L (no tolerance on metres)      ; thành phẩm = qty*(1+dung_sai)
    if family == "opp":
        phe = CONST["phe_chong_mau_m"].get(str(so_mau_in), 0)
        sl_in_default = qty * L * (1 + dung_sai) + phe
        thanh_default = int(round(qty))
    else:
        sl_in_default = qty * L
        thanh_default = int(round(qty * (1 + dung_sai)))
    sl_in = order.get("sl_in_thuc_te_m")
    if sl_in in (None, ""):
        sl_in = sl_in_default
    sl_in = round(float(sl_in), 1)

    fields = {
        "customer": order.get("customer", ""), "product_name": order.get("product_name", ""),
        "product_code": order.get("product_code", ""), "order_id": order.get("order_id", ""),
        "so_phieu_sx": order.get("so_phieu_sx", "SX" + str(order.get("order_id", ""))[-6:]),
        "qty": int(qty) if qty == int(qty) else qty,
        "bag_length_m": L, "width_plus_gusset_m": W,
        "sl_in_thuc_te_m": sl_in, "so_mau_in": so_mau_in,
        "kho_manh": kho_manh, "kho_mang": kho_mang, "kho_giay": kho_giay,
        "inner_bag_weight_kg": ibw,
        "thanh_pham_du_kien": thanh_default,
        "bag_type_label": "Bao OPP" if family == "opp" else "Bao KP",
        "size_sx_mm": _size_mm(order, W, L, tp=False),
        "size_tp_mm": _size_mm(order, W, L, tp=True),
        "sl_theo_po_m": sl_in,
    }

    # ---- material norms
    if family == "opp":
        c = CONST["in_opp"]
        fields["mang_bopp_kg"] = round(sl_in * kho_mang * c["mang_bopp_coeff"], 4)
        fields["dung_moai_opp_kg"] = round(c["dung_moai_opp_coeff"] * sl_in * c["dung_moai_opp_share"], 4)
        fields["dung_moai_ea_kg"] = round(c["dung_moai_opp_coeff"] * sl_in * c["dung_moai_ea_share"], 4)
    else:
        gk = sl_in * kho_giay * (CONST["dinh_luong_giay_kraft_gm2"] / 1000)
        fields["giay_kraft_kg"] = round(gk, 4)
        fields["giay_kraft_code"] = order.get("giay_kraft_code", "GN07010200001")

    # ---- Dán (glue) — written into Dán 2 rows 18-20
    coeff = ov.get("dan_keo_coeff", CONST["dan"]["keo_coeff_opp"] if family == "opp" else CONST["dan"]["keo_coeff_kp"])
    glue_total = qty * L * coeff / 1000
    if ov.get("glue") == "9415_60_vistamax_40":
        fields["glue_rows"] = [
            ("9415", "NHPPFC9415G01", 0.6, round(glue_total * 0.6, 4)),
            ("Vistamaxx", "NHVITAMAX0001", 0.4, round(glue_total * 0.4, 4)),
        ]
    else:
        fields["glue_rows"] = [
            ("F801C", "NHPPNSF801C001", 0.5, round(glue_total * 0.5, 4)),
            ("Vistamaxx", "NHVITAMAX0001", 0.5, round(glue_total * 0.5, 4)),
        ]
    fields["glue_total_kg"] = round(glue_total, 4)
    return fields


def _dims_lookup(order, key):
    w = order.get("width_cm"); g = order.get("gusset_cm")
    if not w or not g:
        return 0
    for row in DIMS["rows"]:
        if row["width_cm"] == int(w) and row["gusset_cm"] == int(g):
            return row[key] / 100
    return 0


def _size_mm(order, W, L, tp):
    wmm = order.get("width_cm"); gmm = order.get("gusset_cm"); lmm = order.get("bag_length_m")
    if wmm and gmm:
        wmm, gmm = int(wmm), int(gmm)
        lmm = int(round(L * 1000))
        return f"({wmm*10} + {gmm*10}) x {lmm}"
    return f"{int(round(W*1000))} x {int(round(L*1000))}"


# -------------------------------------------------------------------- fillers
def _set(xp, sheet, ref, value):
    if value is None or value == "":
        return
    xp.set_value(sheet, ref, value)


def fill_dinh_muc(template_path, family, fields, so_mau_in, out_path):
    xp = XlsxPatch(template_path)
    cm = CELLMAP["dinh_muc"]

    # --- In sheet: header + both info tables ---
    for region in ("header", "info_table_M", "info_table_C"):
        for ref, field in cm["In"].get(region, {}).items():
            _set(xp, "In", ref, fields.get(field))
    _set(xp, "In", "G29", fields.get("thanh_pham_du_kien"))

    # --- In sheet: materials ---
    if family == "opp":
        m = cm["In"]["materials_opp"]
        for row, spec in m.items():
            if row == "22_26" or row.startswith("_"):
                continue
            for col, val in spec.items():
                if val in fields:
                    _set(xp, "In", f"{col}{row}", fields[val])
                else:
                    _set(xp, "In", f"{col}{row}", val)
        # ink color rows 22..(22+N-1). Names/codes set; kg LEFT BLANK (design-specific,
        # unknown until production — that's why "số màu in" is the only question).
        inks = INK["opp"]
        for i in range(5):  # rows 22-26
            r = 22 + i
            xp.set_value("In", f"G{r}", "")  # always clear kg
            if i < so_mau_in:
                _set(xp, "In", f"C{r}", inks[i]["name"])
                _set(xp, "In", f"D{r}", inks[i]["code"])
                _set(xp, "In", f"B{r}", 4 + i)
                _set(xp, "In", f"F{r}", "Kg")
            else:
                xp.set_value("In", f"C{r}", ""); xp.set_value("In", f"D{r}", "")
    else:
        m = cm["In"]["materials_paper_kp"]
        for row, spec in m.items():
            if row == "20_up" or row.startswith("_"):
                continue
            for col, val in spec.items():
                if val in fields:
                    _set(xp, "In", f"{col}{row}", fields[val])
                elif "{" in str(val):
                    pass  # placeholder name, leave template
                else:
                    _set(xp, "In", f"{col}{row}", val)
        inks = INK["flexo"]
        for i in range(4):  # rows 20-23
            r = 20 + i
            xp.set_value("In", f"G{r}", "")  # kg filled at production
            if i < so_mau_in:
                _set(xp, "In", f"C{r}", inks[i]["name"])
                _set(xp, "In", f"D{r}", inks[i]["code"])
                _set(xp, "In", f"B{r}", 2 + i)
                _set(xp, "In", f"F{r}", "Kg")
            else:
                xp.set_value("In", f"C{r}", ""); xp.set_value("In", f"D{r}", "")

    # --- Stage sheets: headers + stage values ---
    for sheet in ("May 1", "Chia biên 2", "Dán 2"):
        if sheet not in xp.sheet_paths:
            continue
        for region in ("header", "stage"):
            for ref, field in cm.get(sheet, {}).get(region, {}).items():
                if ref.startswith("_"):
                    continue
                _set(xp, sheet, ref, fields.get(field))
    # Dán 2 glue rows 18-20
    if "Dán 2" in xp.sheet_paths:
        for i, (name, code, ratio, kg) in enumerate(fields["glue_rows"]):
            r = 18 + i
            _set(xp, "Dán 2", f"B{r}", name)
            _set(xp, "Dán 2", f"D{r}", code)
            _set(xp, "Dán 2", f"E{r}", ratio)
            _set(xp, "Dán 2", f"F{r}", "Kg")
            _set(xp, "Dán 2", f"G{r}", kg)

    xp.save(out_path)
    return out_path


def _single_product(order):
    """Collapse a flat (single-product) order dict into one product record."""
    return {
        "product_code": order.get("product_code", ""),
        "product_name": order.get("product_name", ""),
        "ma_code": order.get("ma_code", ""),
        "qty": order.get("qty"),
        "spec": order.get("spec", ""),
    }


def _per_product_order(order, product):
    """Build a compute()-ready order for one product (inherits header + shared dims)."""
    po = {k: v for k, v in order.items() if k != "products"}
    po.update(product)
    po.update(parse_spec(product.get("spec", "")))
    return po


def fill_ycsx(template_path, order, out_path):
    """Fill the YCSX form: header (label + value) + ALL product line items.

    Stale template data (leftover products, stage markers, old header values,
    old giao-hàng/notes) is cleared so output never leaks a previous order. The
    shared spec is written into each product row's F cell; where F is the
    top-left of a merge (e.g. F13:H15) that sets the merged display.
    """
    xp = XlsxPatch(template_path)
    cm = CELLMAP["ycsx"]
    sheet = list(xp.sheet_paths)[0]

    # header — always rewrite as "label: value" so stale values can't survive
    labels = cm.get("header_labels") or {
        "B6": "1. Ngày yêu cầu", "B7": "2. Khách hàng", "B8": "3. Địa chỉ",
        "B9": "4. Mã khách hàng", "B10": "5. Số đơn hàng",
    }
    hfields = {"B6": "ngay_yc", "B7": "customer", "B8": "address",
               "B9": "customer_code", "B10": "order_id"}
    for ref, label in labels.items():
        val = order.get(hfields.get(ref, ""), "") or ""
        xp.set_value(sheet, ref, f"{label}: {val}".rstrip())

    # product line items from row 13
    products = order.get("products") or [_single_product(order)]
    last_data_row = 12
    for i, p in enumerate(products):
        r = 13 + i
        xp.set_value(sheet, f"B{r}", i + 1)
        xp.set_value(sheet, f"C{r}", p.get("product_code", ""))
        xp.set_value(sheet, f"D{r}", p.get("product_name", ""))
        xp.set_value(sheet, f"E{r}", p.get("ma_code", ""))
        xp.set_value(sheet, f"F{r}", p.get("spec", ""))
        xp.set_value(sheet, f"I{r}", "Cái")
        xp.set_value(sheet, f"J{r}", p.get("qty", ""))
        last_data_row = r

    # clear stale notes/schedule (K,L) across the whole product region, then
    # clear B-J beyond the products written (leftover products + stage markers
    # at 16-21 + giao-hàng at 23). Merged top-lefts (e.g. B23) clear on contact.
    for r in range(13, 24):
        xp.set_value(sheet, f"K{r}", "")
        xp.set_value(sheet, f"L{r}", "")
    for r in range(last_data_row + 1, 24):
        for col in "BCDEFGHIJ":
            xp.set_value(sheet, f"{col}{r}", "")

    xp.save(out_path)
    return out_path


# --------------------------------------------------------------------- sample
SAMPLE_40KG = {
    "customer": "CÔNG TY 4 ORANGES CO.,LTD",
    "product_name": "4O BTA140 SPEC CEO PUTTY FOR INTERIOR 40KG",
    "product_code": "XD1KH00565BP0011",
    "order_id": "26BA2HKH00565000014",
    "so_phieu_sx": "SX2607080",
    "qty": 3000, "bag_length_m": 0.82, "width_plus_gusset_m": 0.5,
    "width_cm": 42, "gusset_cm": 8, "inner_bag_weight_kg": 0.051,
    "kho_mang": 1.04, "spec": "1.Kích thước: (42+8) cm x 82cm\n3. Cấu trúc: Bao BOPP in ống đồng - OPP mờ",
}


def run(source, colors, outdir=None):
    """Generate one Định mức per product + one shared YCSX.

    source: ("sample", None) | ("order", path) | ("ycsx", path) | ("dict", order_dict)
    Returns {"family", "outdir", "outputs":[paths], "products":[{product_name, fields}]}.
    Reused by the CLI (main) and by the web/claude.ai code-execution wrapper.
    """
    kind = source[0]
    if kind == "sample":
        order = copy.deepcopy(SAMPLE_40KG)
    elif kind == "order":
        with open(source[1], encoding="utf-8") as f:
            order = json.load(f)
    elif kind == "ycsx":
        order = parse_ycsx(source[1])
    elif kind == "dict":
        order = copy.deepcopy(source[1])
    else:
        raise ValueError(f"bad source kind {kind!r}")

    products = order.get("products") or [_single_product(order)]
    order["products"] = products
    family = order.get("bag_family") or detect_family(order)

    outdir = outdir or os.path.join(OUTDIR, datetime.date.today().isoformat())
    os.makedirs(outdir, exist_ok=True)
    tmpl = template_path(family)

    outputs, summary = [], []
    for product in products:
        porder = _per_product_order(order, product)
        fields = compute(porder, family, colors)
        safe = re.sub(r"[^0-9A-Za-z]+", "-", product.get("product_name") or "order")[:40]
        dm_out = os.path.join(outdir, f"Định mức - {safe} - {order.get('order_id','')}.xlsx")
        fill_dinh_muc(tmpl, family, fields, colors, dm_out)
        outputs.append(dm_out)
        summary.append({"product_name": product.get("product_name"), "fields": fields})

    ycsx_safe = re.sub(r"[^0-9A-Za-z]+", "-",
                       order.get("product_name") or products[0].get("product_name") or "order")[:40]
    ycsx_out = os.path.join(outdir, f"YCSX - {ycsx_safe} - {order.get('order_id','')}.xlsx")
    fill_ycsx(template_path("ycsx"), order, ycsx_out)
    outputs.append(ycsx_out)

    zip_path = os.path.join(outdir, f"Dinh_muc_YCSX_{order.get('order_id','order')}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fpath in outputs:
            zf.write(fpath, os.path.basename(fpath))

    return {"family": family, "outdir": outdir, "outputs": [zip_path], "products": summary}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--order", help="path to order JSON")
    ap.add_argument("--ycsx", help="path to a YCSX .xlsx to parse")
    ap.add_argument("--sample", action="store_true", help="use built-in 40KG sample")
    ap.add_argument("--colors", type=int, help="pre-answer số màu in (skips the prompt)")
    ap.add_argument("--outdir", help="output directory (default: output/<today/>)")
    args = ap.parse_args()

    if args.sample:
        source = ("sample", None)
    elif args.order:
        source = ("order", args.order)
    elif args.ycsx:
        source = ("ycsx", args.ycsx)
    else:
        ap.error("provide --order, --ycsx, or --sample")

    if args.colors is not None:
        so_mau_in = args.colors
    else:
        so_mau_in = int(input("Số màu in? (number of print colors — the only unknown): ").strip() or "0")

    res = run(source, so_mau_in, outdir=args.outdir)
    print(f"Detected bag family: {res['family']}  (template: {BAG['families'][res['family']]['_aka']})")
    print(f"số màu in = {so_mau_in}\n")
    for p in res["products"]:
        f = p["fields"]
        print(f"== {p['product_name']} ==")
        for k in ("qty", "bag_length_m", "width_plus_gusset_m", "sl_in_thuc_te_m", "so_mau_in",
                  "kho_manh", "kho_mang", "kho_giay", "inner_bag_weight_kg"):
            print(f"  {k:24s} = {f.get(k)}")
        for k in ("mang_bopp_kg", "dung_moai_opp_kg", "dung_moai_ea_kg", "giay_kraft_kg", "glue_total_kg"):
            if k in f:
                print(f"  {k:24s} = {f.get(k)}")
    print("\nGenerated files:")
    for path in res["outputs"]:
        print("  ✓", path)


if __name__ == "__main__":
    main()
