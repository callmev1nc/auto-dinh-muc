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
from xml.sax.saxutils import escape
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
# ET reserves these prefixes for its own generated substitutes (ns0, ns1, …)
# and the xml prefix is built-in; never try to register them.
_NS_RESERVED = re.compile(r"^(?:ns\d+|xml)$")


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
        self._register_doc_namespaces()
        self._read_workbook()
        self._read_styles()
        self._styles_dirty = False

    # ---------- namespaces ----------
    def _register_doc_namespaces(self):
        # Bind every xmlns prefix declared in this workbook to its original name
        # so ElementTree re-serializes with the SAME prefixes. Without this, ET
        # invents ns0:/ns1: substitutes for namespaces it doesn't know (xr,
        # x16r2, xcalcf, ...) while leaving mc:Ignorable="x14ac x16r2 xr xr9"
        # naming the originals -> Excel rejects the part ("part with XML error").
        for name, data in self._data.items():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            text = data.decode("utf-8", "ignore")
            for pfx, uri in re.findall(r'xmlns:([A-Za-z0-9_.\-]+)="([^"]+)"', text):
                if _NS_RESERVED.match(pfx):
                    continue
                try:
                    ET.register_namespace(pfx, uri)
                except ValueError:
                    pass

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
        # ponytail: still ET-round-trips this sheet + styles. Only the dev script
        # colour_25kg.py calls this; generate.py's production path never does.
        self._styles_dirty = True
        sroot = self._sheet_root(sheet)
        c = self._find_cell(sroot, ref)
        if c is None:
            return False
        fill_id = self._ensure_fill(rgb)
        new_xf = self._xf_with_fill(self._cell_style_index(c), fill_id)
        c.set("s", str(new_xf))
        self._data[self.sheet_paths[sheet]] = self._to_xml(sroot)
        return True

    def set_value(self, sheet: str, ref: str, value):
        # ponytail: byte-surgical splice — touch ONLY the target <c> element in
        # the raw sheet XML. Every other byte (xmlns declarations, mc:Ignorable,
        # images, merges) stays identical to the template, so Excel can never
        # reject the part (the prior ET round-trip renamed/dropped namespaces
        # like xr/xr2/xr9 and left mc:Ignorable dangling -> "part with XML error").
        # Ceiling: assumes one <c> per ref and a single-line <c ...> opening tag
        # (true for these Excel-generated templates). Not a general OOXML editor.
        key = self.sheet_paths[sheet]
        xml = self._data[key].decode("utf-8")
        target = 'r="%s"' % ref
        tag = None
        for m in re.finditer(r"<c\b[^>]*>", xml):
            if target in m.group(0):
                tag = m
                break
        if tag is None:
            return False
        opener = tag.group(0)
        start = tag.start()
        if opener.rstrip().endswith("/>"):          # self-closing cell
            end = tag.end()
        else:                                        # <c ...>...</c>
            close = xml.find("</c>", tag.end())
            if close == -1:
                return False
            end = close + 4
        sm = re.search(r'\bs="(\d+)"', opener)       # preserve cell style
        s_attr = ' s="%s"' % sm.group(1) if sm else ""
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            num = repr(value) if isinstance(value, float) else str(value)
            new_cell = '<c r="%s"%s><v>%s</v></c>' % (ref, s_attr, num)
        elif value == "" or value is None:           # clear the cell (keep style)
            new_cell = '<c r="%s"%s/>' % (ref, s_attr)
        else:                                        # string -> inlineStr
            new_cell = ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                        % (ref, s_attr, escape(str(value))))
        self._data[key] = (xml[:start] + new_cell + xml[end:]).encode("utf-8")
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

    # ---------- serialization ----------
    def _to_xml(self, root: ET.Element) -> bytes:
        """Serialize `root`, then strip markup-compatibility ghosts.

        ElementTree drops xmlns declarations for namespaces no element/attribute
        uses. mc:Ignorable="x14ac x16r2 xr xr9" (and mc:Choice/@Requires) still
        name those prefixes though — it's an attribute value, ET won't rewrite
        it — so Excel rejects the part ('part with XML error'). We drop any token
        whose prefix isn't actually declared in the part. Safe: an ignorable
        entry for a namespace that never appears in the part is meaningless.
        """
        return self._sanitize_mc(ET.tostring(root, encoding="utf-8", xml_declaration=True))

    @staticmethod
    def _sanitize_mc(xml_bytes: bytes) -> bytes:
        text = xml_bytes.decode("utf-8")
        declared = set(re.findall(r'xmlns:([A-Za-z0-9_.\-]+)=', text))

        def _keep(attr, value):
            toks = [t for t in value.split() if t in declared]
            return ('%s="%s"' % (attr, " ".join(toks))) if toks else ""

        text = re.sub(r'mc:Ignorable="([^"]*)"',
                      lambda m: _keep("mc:Ignorable", m.group(1)), text)
        text = re.sub(r'(?<= )Requires="([^"]*)"',
                      lambda m: _keep("Requires", m.group(1)), text)
        # collapse the double space left where an attribute was fully removed
        text = re.sub(r'  +', ' ', text).replace(' />', '/>').replace(' >', '>')
        return text.encode("utf-8")

    # ---------- save ----------
    def _strip_calc_chain(self):
        # Drop the cached calc-chain. generate.py rewrites the input cells that
        # the template's stage-sheet formulas depend on (e.g. In!G29/M9/M14), so
        # the copied calcChain.xml + the dependent cells' cached <v> go stale ->
        # Excel pops "Removed Records: Formula from /xl/calcChain.xml" and
        # "repairs" the file. Excel rebuilds the chain silently when the part is
        # absent (openpyxl never writes calcChain.xml for the same reason).
        # Byte-surgical regex splice (same technique as set_value) — no ET
        # round-trip, so no namespace re-serialization risk. No-op for YCSX.
        if "xl/calcChain.xml" not in self._data:
            return
        self._data.pop("xl/calcChain.xml", None)
        self.names = [n for n in self.names if n != "xl/calcChain.xml"]
        ct = self._data["[Content_Types].xml"].decode("utf-8")
        ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
        self._data["[Content_Types].xml"] = ct.encode("utf-8")
        wr = self._data["xl/_rels/workbook.xml.rels"].decode("utf-8")
        wr = re.sub(r'<Relationship [^>]*Target="calcChain\.xml"[^>]*/>', '', wr)
        self._data["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")
        wb = self._data["xl/workbook.xml"].decode("utf-8")
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r'<calcPr\b([^>]*)/>',
                        r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
            self._data["xl/workbook.xml"] = wb.encode("utf-8")

    def save(self, out_path: str):
        # Re-serialize styles.xml ONLY if set_fill changed it. generate.py never
        # calls set_fill, so styles.xml stays byte-identical to the template
        # (the prior unconditional re-serialize corrupted it on every save).
        if self._styles_dirty:
            self._data["xl/styles.xml"] = self._to_xml(self.styles_root)
        self._strip_calc_chain()
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
            for name in self.names:
                zo.writestr(name, self._data[name])
