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
import argparse, json, os, sys, datetime, re, copy
from xlsxpatch import XlsxPatch

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
OUTDIR = os.path.join(HERE, "output")


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
    """Pull header + first line item from a YCSX .xlsx using openpyxl (read-only)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    g = {}
    for row in ws.iter_rows(values_only=False):
        for c in row:
            if isinstance(c.value, str):
                g[c.coordinate] = c.value.strip()
    def find(prefix):
        for coord, val in g.items():
            if val and val.strip().lower().startswith(prefix.lower()):
                # value usually follows in the cell(s) to the right or after a ':'
                tail = val.split(":", 1)[-1].strip()
                return tail
        return ""
    # header labels live in column B
    def bval(label):
        for coord, val in g.items():
            if val and val.startswith(label):
                # take the part after the label/colon
                return val.split(":", 1)[-1].strip() if ":" in val else ""
        return ""
    order = {
        "customer": bval("2. Khách hàng") or g.get("B7", ""),
        "address": bval("3. Địa chỉ"),
        "customer_code": bval("4. Mã khách hàng"),
        "order_id": bval("5. Số đơn hàng"),
        "ngay_yc": bval("1. Ngày yêu cầu"),
        # line item (first data row, row 13)
        "product_code": g.get("C13", ""),
        "product_name": g.get("D13", ""),
        "ma_code": g.get("E13", ""),
        "spec": g.get("F13", ""),
        "qty": _to_num(g.get("J13")),
    }
    order.update(parse_spec(order.get("spec", "")))
    return order


def parse_spec(spec):
    """Extract dimensions / structure / bag hints from the Chi tiết kỹ thuật text."""
    out = {}
    if not spec:
        return out
    s = spec.lower()
    # dimensions like "(42+8) cm x 82cm" or "50x92 cm"
    m = re.search(r"(\d+)\s*[\+]\s*(\d+)\s*cm\s*x\s*(\d+(?:\.\d+)?)\s*cm", spec)
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
    text = " ".join([str(order.get("product_name", "")), str(order.get("spec", ""))]).lower()
    opp = BAG["families"]["opp"]
    kp = BAG["families"]["paper_kp"]
    if any(k.lower() in text for k in opp["name_keywords"] + opp["spec_keywords"]):
        return "opp"
    if any(k.lower() in text for k in kp["name_keywords"] + kp["spec_keywords"]):
        return "paper_kp"
    # fallback by weight token
    mw = re.search(r"(\d{2,3})\s*kg", text)
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


def fill_ycsx(template_path, order, fields, out_path):
    xp = XlsxPatch(template_path)
    cm = CELLMAP["ycsx"]
    sheet = xp.sheet_paths and list(xp.sheet_paths)[0]
    for ref, field in cm["header"].items():
        _set(xp, sheet, ref, order.get(field, fields.get(field)))
    # first line item row 13
    li = cm["line_item_template"]
    row = 13
    _set(xp, sheet, f"B{row}", 1)
    _set(xp, sheet, f"C{row}", order.get("product_code"))
    _set(xp, sheet, f"D{row}", order.get("product_name"))
    _set(xp, sheet, f"E{row}", order.get("ma_code"))
    _set(xp, sheet, f"F{row}", order.get("spec"))
    _set(xp, sheet, f"I{row}", "Cái")
    _set(xp, sheet, f"J{row}", order.get("qty"))
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--order", help="path to order JSON")
    ap.add_argument("--ycsx", help="path to a YCSX .xlsx to parse")
    ap.add_argument("--sample", action="store_true", help="use built-in 40KG sample")
    ap.add_argument("--colors", type=int, help="pre-answer số màu in (skips the prompt)")
    args = ap.parse_args()

    if args.sample:
        order = copy.deepcopy(SAMPLE_40KG)
    elif args.order:
        with open(args.order, encoding="utf-8") as f:
            order = json.load(f)
    elif args.ycsx:
        order = parse_ycsx(args.ycsx)
    else:
        ap.error("provide --order, --ycsx, or --sample")

    family = order.get("bag_family") or detect_family(order)
    print(f"Detected bag family: {family}  (template: {BAG['families'][family]['_aka']})")

    if args.colors is not None:
        so_mau_in = args.colors
    else:
        so_mau_in = int(input("Số màu in? (number of print colors — the only unknown): ").strip() or "0")
    print(f"số màu in = {so_mau_in}")

    fields = compute(order, family, so_mau_in)
    print("\nComputed values:")
    for k in ("qty", "bag_length_m", "width_plus_gusset_m", "sl_in_thuc_te_m", "so_mau_in",
              "kho_manh", "kho_mang", "kho_giay", "inner_bag_weight_kg"):
        print(f"  {k:24s} = {fields.get(k)}")
    for k in ("mang_bopp_kg", "dung_moai_opp_kg", "dung_moai_ea_kg", "giay_kraft_kg", "glue_total_kg"):
        if k in fields:
            print(f"  {k:24s} = {fields.get(k)}")

    today = datetime.date.today().isoformat()
    outdir = os.path.join(OUTDIR, today)
    os.makedirs(outdir, exist_ok=True)
    tmpl = os.path.join(HERE, "templates", {
        "opp": "Định mức - OPP (Bao BOPP in ống đồng).xlsx",
        "paper_kp": "Định mức - Bao giấy (KP - in offset).xlsx",
    }[family])
    safe = re.sub(r"[^0-9A-Za-z]+", "-", order.get("product_name", "order"))[:40]
    dm_out = os.path.join(outdir, f"Định mức - {safe} - {order.get('order_id','')}.xlsx")
    ycsx_out = os.path.join(outdir, f"YCSX - {safe} - {order.get('order_id','')}.xlsx")

    fill_dinh_muc(tmpl, family, fields, so_mau_in, dm_out)
    fill_ycsx(os.path.join(HERE, "templates", "YCSX.xlsx"), order, fields, ycsx_out)
    print(f"\n✓ Định mức: {dm_out}")
    print(f"✓ YCSX:     {ycsx_out}")


if __name__ == "__main__":
    main()
