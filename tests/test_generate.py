"""Tests: every expected cell filled, regression values, no stray paint."""
from __future__ import annotations

import json
import os
import sys
import tempfile

from openpyxl import load_workbook

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from generate import compute, detect_family, run

HERE = os.path.dirname(__file__)
SAMPLES = os.path.join(HERE, "..", "samples")


def _load_json(name):
    with open(os.path.join(SAMPLES, name), encoding="utf-8") as f:
        return json.load(f)


SAMPLE_25KG = _load_json("25kg_tan_chau.json")
SAMPLE_40KG = _load_json("40kg_4oranges.json")


def _num(ws, cell):
    v = ws[cell].value
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _str(ws, cell):
    v = ws[cell].value
    return str(v).strip() if v is not None else ""


def _fill_color(ws, cell):
    """Return the fill fgColor rgb of a cell, or None."""
    fill = ws[cell].fill
    if fill and fill.fgColor and fill.fgColor.rgb and fill.fgColor.rgb not in ("00000000",):
        return fill.fgColor.rgb
    return None


def _check_all_filled(ws, cells, label):
    """Assert every cell in the list is non-empty."""
    for cell in cells:
        v = _str(ws, cell)
        assert v != "", f"{label}: {cell} should be filled but is empty"


def _check_no_solid_fill(ws, cells, label):
    """Assert no cell has a solid fill colour (cleanliness guard)."""
    BAD = ("FFFF00", "FF0000")
    for cell in cells:
        c = ws[cell]
        if c.fill and c.fill.fgColor:
            rgb = c.fill.fgColor.rgb
            if rgb and rgb not in ("00000000",):
                assert rgb not in BAD, f"{label}: {cell} has forbidden fill {rgb}"


# ---- sheet: In ----
IN_HEADER = [f"C{i}" for i in range(5, 10)]
IN_INFO_M = [f"M{i}" for i in range(6, 15)]
IN_INFO_C = [f"C{i}" for i in range(39, 48)]
IN_THANH_PHAM = ["G29"]

# ---- material rows ----
PAPER_MAT_CELLS = ["C18", "D18", "F18", "G18", "F19", "G19"]
OPP_MAT_CELLS  = ["C18", "D18", "F18", "G18", "F19", "G19",
                  "C20", "D20", "F20", "G20", "C21", "D21", "F21", "G21"]

# ---- stage sheets: cells that should always be filled ----
STAGE_CELLS = {
    "May 1": [f"C{i}" for i in range(16, 21)],
    "Chia bien 2": ["C16", "G16"],
    "Dan 2": ["D16", "G18", "G19"],
}


def _stage_cells(xp, sheet_name):
    """Get filled stage cells for a given sheet if it exists."""
    if sheet_name in xp.sheetnames:
        return STAGE_CELLS[sheet_name]
    return []


# ---- stage-formula wire checks (data_only=False) ----
KP_STAGE_WIRE: dict[str, list[str]] = {
    "Tráng": ["G17", "G18", "G19"],
    "Dán":   ["G17", "G18"],
    "Thổi":  ["G17", "G19"],
    "May":   ["G17", "G21", "G22"],
}

OPP_STAGE_WIRE: dict[str, list[str]] = {
    "Tráng": ["G17", "G18", "G19", "G20"],
    "Dán":   ["G17", "G18"],
    "Thổi":  ["G17", "G19"],
    "May":   ["G17"],
}

LINKAGE_M_CELLS = [f"M{i}" for i in range(6, 10)]


def _check_stage_formula_wire(ws, sheet, g_cells, label):
    """Assert at least one M-cell in LINKAGE is a cross-sheet reference (contains ``!``)
    and every G-cell in g_cells is a formula (starts with ``=``)."""
    assert any(
        isinstance(ws[c].value, str) and "!" in ws[c].value
        for c in LINKAGE_M_CELLS
    ), f"{label}: {sheet} has no cross-sheet reference in M6:M9"
    for cell in g_cells:
        if cell not in ws:
            continue
        v = ws[cell].value
        assert isinstance(v, str) and v.startswith("="), \
            f"{label}: {sheet}!{cell} should be a formula, got {v!r}"


# ---------------------------------------------------------------- tests

class Test25KgKp:
    """Paper/KP 25kg Tan Chau -- 2 mau in."""

    @classmethod
    def setup_class(cls):
        cls.outdir = tempfile.mkdtemp(prefix="dm_test_25kg_")
        cls.order = dict(SAMPLE_25KG)
        cls.family = detect_family(cls.order)
        cls.fields = compute(cls.order, cls.family, 2)
        run(("dict", cls.order), 2, outdir=cls.outdir)
        xlsx_files = [os.path.join(cls.outdir, f) for f in os.listdir(cls.outdir)
                      if f.endswith(".xlsx") and f.startswith("Định mức")]
        assert xlsx_files, f"no Dinh muc xlsx in {os.listdir(cls.outdir)}"
        cls.dm_path = xlsx_files[0]
        cls.wb = load_workbook(cls.dm_path, data_only=True)
        cls.wb_f = load_workbook(cls.dm_path, data_only=False)
        cls.ws = cls.wb["In"]

    @classmethod
    def teardown_class(cls):
        cls.wb.close()
        cls.wb_f.close()
        import shutil
        shutil.rmtree(cls.outdir, ignore_errors=True)

    def test_completeness_header(self):
        _check_all_filled(self.ws, IN_HEADER, "In header")

    def test_completeness_info_m(self):
        _check_all_filled(self.ws, IN_INFO_M, "In info_M")

    def test_completeness_info_c(self):
        _check_all_filled(self.ws, IN_INFO_C, "In info_C")

    def test_completeness_thanh_pham(self):
        _check_all_filled(self.ws, IN_THANH_PHAM, "In thanh_pham")

    def test_completeness_materials(self):
        _check_all_filled(self.ws, PAPER_MAT_CELLS, "In materials_paper_kp")

    def test_kraft_name_filled(self):
        assert _str(self.ws, "C18") == "Giấy Kraft vàng Nhật K1020 ĐL70", \
            f"In!C18 got {_str(self.ws, 'C18')!r}"

    def test_stage_cells(self):
        for sheet_name in STAGE_CELLS:
            if sheet_name in self.wb.sheetnames:
                ws = self.wb[sheet_name]
                _check_all_filled(ws, STAGE_CELLS[sheet_name], f"{sheet_name} stage")

    def test_stage_formulas_wire_to_in(self):
        for sheet, g_cells in KP_STAGE_WIRE.items():
            if sheet not in self.wb_f.sheetnames:
                continue
            ws = self.wb_f[sheet]
            _check_stage_formula_wire(ws, sheet, g_cells, "KP")

    def test_regression_values(self):
        assert _num(self.ws, "M9") == 4600.0, f"M9={_num(self.ws, 'M9')}"
        assert _num(self.ws, "G18") == 328.44, f"G18={_num(self.ws, 'G18')}"
        assert _num(self.ws, "G29") == 5250.0, f"G29={_num(self.ws, 'G29')}"
        assert _num(self.ws, "M6") == 5000.0, f"M6={_num(self.ws, 'M6')}"
        assert _num(self.ws, "M14") == 0.041, f"M14={_num(self.ws, 'M14')}"

    def test_regression_glue(self):
        ws = self.wb["Dán 2"]
        assert abs(_num(ws, "D16") - 15.2444) < 0.01, f"Dán 2!D16={_num(ws, 'D16')}"
        assert abs(_num(ws, "G18") - 9.1466) < 0.01, f"Dán 2!G18={_num(ws, 'G18')}"
        assert abs(_num(ws, "G19") - 6.0978) < 0.01, f"Dán 2!G19={_num(ws, 'G19')}"

    def test_cleanliness_no_stray_fill(self):
        all_cells = IN_HEADER + IN_INFO_M + IN_INFO_C + IN_THANH_PHAM + PAPER_MAT_CELLS
        _check_no_solid_fill(self.ws, all_cells, "In")

    def test_ink_kg_blank(self):
        for r in range(20, 24):
            v = _num(self.ws, f"G{r}")
            assert v is None or v == 0, f"In!G{r}={v} should be blank"


class Test40KgOpp:
    """OPP 40kg 4 Oranges -- 3 mau in."""

    @classmethod
    def setup_class(cls):
        cls.outdir = tempfile.mkdtemp(prefix="dm_test_40kg_")
        cls.order = dict(SAMPLE_40KG)
        cls.family = detect_family(cls.order)
        cls.fields = compute(cls.order, cls.family, 3)
        run(("dict", cls.order), 3, outdir=cls.outdir)
        xlsx_files = [os.path.join(cls.outdir, f) for f in os.listdir(cls.outdir)
                      if f.endswith(".xlsx") and f.startswith("Định mức")]
        assert xlsx_files, f"no Dinh muc xlsx in {os.listdir(cls.outdir)}"
        cls.dm_path = xlsx_files[0]
        cls.wb = load_workbook(cls.dm_path, data_only=True)
        cls.wb_f = load_workbook(cls.dm_path, data_only=False)
        cls.ws = cls.wb["In"]

    @classmethod
    def teardown_class(cls):
        cls.wb.close()
        cls.wb_f.close()
        import shutil
        shutil.rmtree(cls.outdir, ignore_errors=True)

    def test_completeness_header(self):
        _check_all_filled(self.ws, IN_HEADER, "In header")

    def test_completeness_info_m(self):
        _check_all_filled(self.ws, IN_INFO_M, "In info_M")

    def test_completeness_info_c(self):
        _check_all_filled(self.ws, IN_INFO_C, "In info_C")

    def test_completeness_thanh_pham(self):
        _check_all_filled(self.ws, IN_THANH_PHAM, "In thanh_pham")

    def test_completeness_materials(self):
        _check_all_filled(self.ws, OPP_MAT_CELLS, "In materials_opp")

    def test_opp_material_name(self):
        assert _str(self.ws, "C18") == "Màng BOPP mờ K1040 18 mic", \
            f"In!C18 got {_str(self.ws, 'C18')!r}"

    def test_stage_cells(self):
        for sheet_name in STAGE_CELLS:
            if sheet_name in self.wb.sheetnames:
                ws = self.wb[sheet_name]
                _check_all_filled(ws, STAGE_CELLS[sheet_name], f"{sheet_name} stage")

    def test_stage_formulas_wire_to_in(self):
        for sheet, g_cells in OPP_STAGE_WIRE.items():
            if sheet not in self.wb_f.sheetnames:
                continue
            ws = self.wb_f[sheet]
            _check_stage_formula_wire(ws, sheet, g_cells, "OPP")

    def test_regression_values(self):
        assert _num(self.ws, "M9") == 2983.0, f"M9={_num(self.ws, 'M9')}"
        assert abs(_num(self.ws, "G18") - 49.1407) < 0.01, f"G18={_num(self.ws, 'G18')}"
        assert abs(_num(self.ws, "G20") - 12.5286) < 0.01, f"G20={_num(self.ws, 'G20')}"
        assert abs(_num(self.ws, "G21") - 5.3694) < 0.01, f"G21={_num(self.ws, 'G21')}"
        assert _num(self.ws, "G29") == 3000.0, f"G29={_num(self.ws, 'G29')}"

    def test_regression_glue(self):
        if "Dan 2" in self.wb.sheetnames:
            ws = self.wb["Dan 2"]
            assert abs(_num(ws, "D16") - 8.0964) < 0.01, f"Dan 2!D16={_num(ws, 'D16')}"

    def test_cleanliness_no_stray_fill(self):
        all_cells = IN_HEADER + IN_INFO_M + IN_INFO_C + IN_THANH_PHAM + OPP_MAT_CELLS
        _check_no_solid_fill(self.ws, all_cells, "In")

    def test_ink_kg_blank(self):
        for r in range(22, 27):
            v = _num(self.ws, f"G{r}")
            assert v is None or v == 0, f"In!G{r}={v} should be blank"
