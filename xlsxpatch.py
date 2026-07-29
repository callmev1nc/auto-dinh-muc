"""
xlsxpatch.py — byte-faithful in-place patcher for .xlsx files.

Why this exists: openpyxl drops embedded images/charts/drawings on save.
The Định mức / YCSX workbooks embed logos on every sheet, so we must NOT
round-trip them through openpyxl. Instead we edit the raw OOXML zip: we
touch ONLY xl/styles.xml + the specific sheet XML we want to change, and
copy every other zip entry (media, drawings, rels, other sheets) verbatim.

Public API:
    xp = XlsxPatch(path)          # load workbook into memory
    xp.set_fill(sheet, ref, rgb)  # paint a solid fill (e.g. "FFFFFF00")
    xp.set_value(sheet, ref, val) # write a number or string (inline)
    xp.get_fill(sheet, ref)       # read current fill rgb / None
    xp.save(out_path)             # write a new .xlsx
"""
from __future__ import annotations
import re, zipfile, shutil
import xml.etree.ElementTree as ET
from copy import deepcopy

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS   = "{%s}" % MAIN

# Keep prefixes stable on re-serialization so Excel reads the file happily.
ET.register_namespace("", MAIN)
for pfx, uri in [
    ("r", RNS),
    ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
    ("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"),
    ("xm", "http://schemas.microsoft.com/office/excel/2006/main"),
    ("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"),
    ("xsi", "http://www.w3.org/2001/XMLSchema-instance"),
    ("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
    ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
]:
    ET.register_namespace(pfx, uri)

_RE_CELL = re.compile(r"^([A-Z]+)(\d+)$")


def _ref(ref):
    m = _RE_CELL.match(ref)
    if not m:
        raise ValueError(f"bad cell ref {ref!r}")
    return m.group(1), int(m.group(2))


class XlsxPatch:
    def __init__(self, path: str):
        self.path = path
        self.zf = zipfile.ZipFile(path)
        self.names = self.zf.namelist()
        self._data = {n: self.zf.read(n) for n in self.names}
        self.zf.close()
        self._read_workbook()
        self._read_styles()

    # ---------- workbook / sheet path resolution ----------
    def _read_workbook(self):
        self.sheet_paths = {}  # sheet name -> zip path
        wb = ET.fromstring(self._data["xl/workbook.xml"])
        rels = {}
        if "xl/_rels/workbook.xml.rels" in self._data:
            rr = ET.fromstring(self._data["xl/_rels/workbook.xml.rels"])
            rels = {r.get("Id"): r.get("Target") for r in rr}
        for s in wb.find(NS + "sheets").findall(NS + "sheet"):
            name = s.get("name")
            tgt = rels.get(s.get("{%s}id" % RNS), "")
            if tgt and not tgt.startswith("xl/"):
                tgt = "xl/" + tgt
            if tgt.startswith("/"):
                tgt = tgt[1:]
            self.sheet_paths[name] = tgt

    # ---------- styles ----------
    def _read_styles(self):
        self.styles_root = ET.fromstring(self._data["xl/styles.xml"])
        self.fills_el = self.styles_root.find(NS + "fills")
        self.xfs_el = self.styles_root.find(NS + "cellXfs")

    def _ensure_fill(self, rgb: str) -> int:
        rgb = rgb.upper()
        for i, f in enumerate(self.fills_el.findall(NS + "fill")):
            pf = f.find(NS + "patternFill")
            if pf is not None and pf.get("patternType") == "solid":
                fg = pf.find(NS + "fgColor")
                if fg is not None and (fg.get("rgb") or "").upper() == rgb:
                    return i
        # create new solid fill
        f = ET.SubElement(self.fills_el, NS + "fill")
        pf = ET.SubElement(f, NS + "patternFill")
        pf.set("patternType", "solid")
        fg = ET.SubElement(pf, NS + "fgColor")
        fg.set("rgb", rgb)
        bg = ET.SubElement(pf, NS + "bgColor")
        bg.set("indexed", "64")
        self.fills_el.set("count", str(len(self.fills_el.findall(NS + "fill"))))
        return len(self.fills_el.findall(NS + "fill")) - 1

    def _xf_signature(self, xf: ET.Element, fill_id_override=None) -> str:
        # fill_id_override=None -> use the xf's OWN fillId (for comparing existing
        # styles); pass an int to pretend the xf has that fill (for the target).
        fillid = str(fill_id_override) if fill_id_override is not None else xf.get("fillId")
        attrs = {k: v for k, v in xf.attrib.items() if k != "fillId"}
        attrs["fillId"] = fillid
        children = "".join(
            ET.tostring(c, encoding="unicode") for c in xf.findall(NS + "alignment")
        )
        return repr(sorted(attrs.items())) + children

    def _xf_with_fill(self, src_index: int, fill_id: int) -> int:
        src = list(self.xfs_el.findall(NS + "xf"))[src_index]
        target_sig = self._xf_signature(src, fill_id)      # src attrs + NEW fill
        # reuse an existing xf only if ITS OWN fill already equals the target
        for i, xf in enumerate(self.xfs_el.findall(NS + "xf")):
            if self._xf_signature(xf) == target_sig:
                return i
        new = deepcopy(src)
        new.set("fillId", str(fill_id))
        new.set("applyFill", "1")
        self.xfs_el.append(new)
        self.xfs_el.set("count", str(len(self.xfs_el.findall(NS + "xf"))))
        return len(self.xfs_el.findall(NS + "xf")) - 1

    # ---------- cell helpers ----------
    def _sheet_root(self, sheet: str) -> ET.Element:
        return ET.fromstring(self._data[self.sheet_paths[sheet]])

    def _find_cell(self, sroot: ET.Element, ref: str):
        col, row = _ref(ref)
        for r in sroot.find(NS + "sheetData").findall(NS + "row"):
            if r.get("r") != str(row):
                continue
            for c in r.findall(NS + "c"):
                if c.get("r") == ref:
                    return c
        return None

    def _cell_style_index(self, c: ET.Element) -> int:
        return int(c.get("s") or 0)

    # ---------- public ops ----------
    def set_fill(self, sheet: str, ref: str, rgb: str):
        sroot = self._sheet_root(sheet)
        c = self._find_cell(sroot, ref)
        if c is None:
            return False
        fill_id = self._ensure_fill(rgb)
        new_xf = self._xf_with_fill(self._cell_style_index(c), fill_id)
        c.set("s", str(new_xf))
        self._data[self.sheet_paths[sheet]] = ET.tostring(sroot, encoding="utf-8", xml_declaration=True)
        return True

    def set_value(self, sheet: str, ref: str, value):
        sroot = self._sheet_root(sheet)
        c = self._find_cell(sroot, ref)
        if c is None:
            return False
        # strip existing value children
        for tag in ("v", "is"):
            e = c.find(NS + tag)
            if e is not None:
                c.remove(e)
        # drop formulas
        f = c.find(NS + "f")
        if f is not None:
            c.remove(f)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            c.attrib.pop("t", None)
            v = ET.SubElement(c, NS + "v")
            v.text = repr(value) if isinstance(value, float) else str(value)
        else:
            c.set("t", "inlineStr")
            is_el = ET.SubElement(c, NS + "is")
            t = ET.SubElement(is_el, NS + "t")
            t.text = " " + str(value) if str(value).startswith((" ", "\n")) else str(value)
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        self._data[self.sheet_paths[sheet]] = ET.tostring(sroot, encoding="utf-8", xml_declaration=True)
        return True

    def get_fill(self, sheet: str, ref: str):
        sroot = self._sheet_root(sheet)
        c = self._find_cell(sroot, ref)
        if c is None:
            return None
        sidx = self._cell_style_index(c)
        xf = list(self.xfs_el.findall(NS + "xf"))[sidx]
        fill_id = int(xf.get("fillId") or 0)
        fills = self.fills_el.findall(NS + "fill")
        if fill_id >= len(fills):
            return None
        pf = fills[fill_id].find(NS + "patternFill")
        if pf is None or pf.get("patternType") != "solid":
            return None
        fg = pf.find(NS + "fgColor")
        return fg.get("rgb") if fg is not None else None

    # ---------- save ----------
    def save(self, out_path: str):
        # re-serialize styles.xml (fills/xfs may have changed)
        self._data["xl/styles.xml"] = ET.tostring(self.styles_root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
            for name in self.names:
                zo.writestr(name, self._data[name])
