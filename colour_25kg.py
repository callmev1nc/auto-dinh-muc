"""
colour_25kg.py — paint the 25KG Định mức's input cells yellow, mirroring
exactly the yellow (FFFFFF00) cells from the 40KG example.

The 40KG file marks every "value that changes per order" in yellow on its
input sheets (In, May 1, Chia biên 2, Dán 2). The 25KG file uses the same
company form template but was never coloured. This copies that highlight
set across, so the 25KG file is filled "the same way" as the 40KG one.

Output is a NEW file (original is never modified): "<25KG> - COLOURED.xlsx".
"""
import os, sys, zipfile
import xml.etree.ElementTree as ET
from xlsxpatch import XlsxPatch, NS, RNS

BASE = r"e:\Work\Thai-work"
F40 = os.path.join(BASE, "Định mức - 4O BTA140 SPEC CEO PUTTY FOR INTERIOR 40KG - 26BA2HKH00565000014.xlsx")
F25 = os.path.join(BASE, "Định mức - Bao giấy 25kg AM-20 Tân Châu - 26BA2HKH00563000029.xlsx")
OUT = F25.replace(".xlsx", " - COLOURED.xlsx")
YELLOW = "FFFFFF00"


def yellow_cells_per_sheet(path: str) -> dict:
    """Return {sheet_name: set(cell_refs)} for every FFFFFF00 cell in `path`."""
    z = zipfile.ZipFile(path)
    relmap = {}
    if "xl/_rels/workbook.xml.rels" in z.namelist():
        for rel in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
            relmap[rel.get("Id")] = rel.get("Target")
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    st = ET.fromstring(z.read("xl/styles.xml"))
    fills = []
    for f in st.find(NS + "fills").findall(NS + "fill"):
        pf = f.find(NS + "patternFill")
        rgb = None
        if pf is not None:
            fg = pf.find(NS + "fgColor")
            rgb = fg.get("rgb") if fg is not None else None
        fills.append((pf.get("patternType") if pf is not None else None, rgb))
    xfs = [x.get("fillId") for x in st.find(NS + "cellXfs").findall(NS + "xf")]
    out = {}
    for s in wb.find(NS + "sheets").findall(NS + "sheet"):
        name = s.get("name")
        tgt = relmap.get(s.get("{%s}id" % RNS), "")
        if tgt and not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        if tgt.startswith("/"):
            tgt = tgt[1:]
        sd = ET.fromstring(z.read(tgt))
        for row in sd.find(NS + "sheetData").findall(NS + "row"):
            for c in row.findall(NS + "c"):
                sidx = c.get("s")
                if sidx is None:
                    continue
                fid = int(xfs[int(sidx)])
                ptype, rgb = fills[fid]
                if rgb and rgb.upper() == YELLOW and ptype == "solid":
                    out.setdefault(name, set()).add(c.get("r"))
    z.close()
    return out


def main():
    yellow = yellow_cells_per_sheet(F40)
    print("40KG yellow input cells per sheet:")
    for sh, cells in sorted(yellow.items()):
        print(f"  {sh}: {len(cells)} cells")

    # copy 25KG -> OUT, then patch (XlsxPatch loads the file itself)
    import shutil
    shutil.copyfile(F25, OUT)
    xp = XlsxPatch(OUT)

    total = 0
    missing_sheets = []
    for sh, cells in yellow.items():
        if sh not in xp.sheet_paths:
            missing_sheets.append(sh)
            continue
        for ref in sorted(cells):
            ok = xp.set_fill(sh, ref, YELLOW)
            if ok:
                total += 1
        print(f"  painted {sh}: {len(cells)} cells")
    if missing_sheets:
        print("  (skipped sheets not present in 25KG):", missing_sheets)

    xp.save(OUT)
    print(f"\nDone. Painted {total} cells yellow.")
    print(f"Output: {OUT}")


if __name__ == "__main__":
    main()
