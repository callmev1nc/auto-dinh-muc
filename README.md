# auto-dinh-muc

Automate a packaging factory's per-order paperwork: given a **customer order**, generate
two Excel deliverables — a **Định mức** (material-norm, 23-sheet) workbook and a **YCSX**
(production-request) form — asking the user only **"số màu in?"** (how many print colors).

> ⚠️ **Private repo.** Contains real company data: customer names, product/order codes,
> the Quy trình formulas, and template logos. Do not make public without sanitizing.

## How it works (hybrid)
- A **claude.ai Project** holds the instructions + knowledge (the brain).
- `generate.py` (Python) runs in **Claude Code / your PC** to write the exact `.xlsx`.
- `xlsxpatch.py` edits the OOXML directly so every logo/merge/style is preserved
  (openpyxl is used only to *parse* incoming YCSX — it would drop images on save).

## Quickstart
```bash
python -m pip install openpyxl
python generate.py --sample --colors 3                              # 40KG OPP demo
python generate.py --order samples/25kg_tan_chau.json --colors 2    # 25KG paper demo
python generate.py --order my_order.json                           # asks "số màu in?"
python generate.py --ycsx "C:/path/YCSX.xlsx" --colors 4           # parse a YCSX
```
Output lands in `output/<YYYY-MM-DD>/`.

## Layout
```
generate.py / xlsxpatch.py    engine
data/                         the brain: constants, formulas, cell-map, rules, inks, dims
templates/                    clone sources (2 Định mức + YCSX)
knowledge/                    upload to the claude.ai Project as Knowledge
samples/                      reference orders
PROJECT_INSTRUCTIONS.md       paste into the Project's Custom Instructions
README_setup.md               full setup + daily-use guide
```

See **[README_setup.md](README_setup.md)** for the claude.ai Project setup and tuning.
