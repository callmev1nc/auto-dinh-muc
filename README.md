# auto-dinh-muc

Automate a packaging factory's per-order paperwork: given a **customer order**, generate
Excel deliverables — **Định mức** (material-norm, 23-sheet) workbooks and a **YCSX**
(production-request) form — asking the user only **"số màu in?"** (how many print colors).

> ⚠️ **Private repo.** Contains real company data: customer names, product/order codes,
> the Quy trình formulas, and template logos. Do not make public without sanitizing.

## How it works
- **Primary path (claude.ai Project — code execution):** co-workers upload the customer
  order to the Project; Claude reads it, asks *"số màu in?"*, runs the Python engine in
  its sandbox, and returns the finished `.xlsx` files (logo preserved) for download.
  Requires **Claude Pro**, no API key, no local Python.
- **Secondary path (local PC / Claude Code):** `generate.py` + `xlsxpatch.py` write the
  exact `.xlsx` by cloning a template and filling only input cells. Used for development
  or when the sandbox isn't available.
- `xlsxpatch.py` edits OOXML directly so every logo/merge/style is preserved
  (openpyxl is used only to *parse* incoming YCSX).

## Quickstart (local)
```bash
python -m pip install openpyxl
python generate.py --sample --colors 3                              # 40KG OPP demo
python generate.py --order samples/25kg_tan_chau.json --colors 2    # 25KG paper demo
python generate.py --order my_order.json                           # asks "số màu in?"
python generate.py --ycsx "C:/path/YCSX.xlsx" --colors 4           # parse a YCSX
```
Output lands in `output/<YYYY-MM-DD>/` (or `--outdir <dir>`).

## claude.ai Project setup
1. Create a Project at https://claude.ai → Projects → New.
2. **Custom Instructions**: paste `PROJECT_INSTRUCTIONS.md`.
3. **Knowledge**: upload `dinh_muc_kit.zip` + `knowledge/scenario_and_formulas.md`
   + the 6 `data/*.json` files.
4. Share the Project link with co-workers (Pro seats).

Usage: **Upload the YCSX .xlsx → answer "số màu in?" → download the finished Excel.**

## Layout
```
generate.py / xlsxpatch.py    engine
make_kit.py                   bundle dinh_muc_kit.zip for sandbox deployment
data/                         the brain: constants, formulas, cell-map, rules, inks, dims
templates/                    clone sources (2 Định mức + YCSX)
knowledge/                    upload to the claude.ai Project as Knowledge
samples/                      reference orders
PROJECT_INSTRUCTIONS.md       paste into the Project's Custom Instructions (Vietnamese, 2-turn)
don_hang.bat / tao_don.py     1-click launcher (PC): asks order file + "số màu in", runs engine
README_setup.md               full setup + daily-use guide
```

See **[README_setup.md](README_setup.md)** for detailed setup and tuning.

## Verification (regression-passing)
- OPP sample: sl_in=2983, màng BOPP=49.1407, dung môi OPP=12.5286/EA=5.3694.
- Paper sample: sl_in=4600, giấy kraft=328.44, glue=15.2444 (9415 rule).
- Generated Định mức preserves all embedded images and opens identically to the example.
