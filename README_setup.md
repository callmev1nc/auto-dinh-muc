# Setup guide — Định mức + YCSX automation

This kit automates: **customer order → Định mức workbook + YCSX form (.xlsx)**.
The only thing ever asked of you is **"số màu in?"** (how many print colors).

## Architecture
- **Primary: claude.ai Project with code execution.** The sandbox runs the Python engine
  directly — co-workers upload the order, answer one question, and download the finished
  `.xlsx` (with the company logo preserved). Requires **Claude Pro** (no API key, no local
  Python).
- **Secondary: local PC / Claude Code.** `generate.py` + `xlsxpatch.py` write the exact
  `.xlsx` by cloning templates and filling only input cells. Used for development or
  fallback.
- `xlsxpatch.py` edits OOXML directly so every logo, merge, and style is preserved.
  openpyxl is used only to *parse* incoming YCSX files.

## Files in this kit
```
auto-dinh-muc/
  generate.py            # the engine: order -> xlsx (asks only "số màu in?")
  xlsxpatch.py           # byte-faithful .xlsx editor (preserves images/styles)
  make_kit.py            # bundle dinh_muc_kit.zip for claude.ai sandbox deployment
  colour_25kg.py         # one-off: painted the 25KG file's inputs yellow
  PROJECT_INSTRUCTIONS.md# paste into the claude.ai Project's Custom Instructions
  don_hang.bat / tao_don.py # 1-click launcher (Option C): hỏi file đơn + số màu in
  README_setup.md        # this file
  data/                  # the brain: constants, formulas, cell-map, rules, inks, dims
    constants.json standard_dims.json cell_map.json
    customer_rules.json bag_type_map.json ink_master.json
  knowledge/             # upload these to the claude.ai Project as Knowledge
    scenario_and_formulas.md
  templates/             # clone sources (the 2 Định mức examples + YCSX)
  samples/               # reference orders (25kg_tan_chau.json, 40kg_4oranges.json)
  dinh_muc_kit.zip       # portable kit for sandbox (build with `python make_kit.py`)
  output/<date>/         # generated files land here
```

## One-time setup (claude.ai Project — primary path)
1. **Build the kit**: `python make_kit.py` → produces `dinh_muc_kit.zip`.
2. **Create the Project** at https://claude.ai → Projects → New:
   - **Custom Instructions**: paste the contents of `PROJECT_INSTRUCTIONS.md`.
   - **Knowledge (Project knowledge)**: upload `dinh_muc_kit.zip`,
     `knowledge/scenario_and_formulas.md`, and the 6 `data/*.json` files.
3. **Tell co-workers**: "Vào Project → upload file YCSX → trả lời số màu in → tải Excel về."

## One-time setup (local PC — secondary / development path)
1. `python -m pip install openpyxl`
2. Keep this `auto-dinh-muc/` folder on your PC (or in a Claude Code workspace).

## Daily use
**Option A — claude.ai Project (recommended — no Python needed):**
1. Open the Project on claude.ai.
2. Upload the customer order YCSX (.xlsx).
3. Claude reads it and asks *"Mặt hàng này in bao nhiêu màu?"*.
4. Answer *"N"*. Claude runs the engine in its sandbox and returns the `.xlsx` files
   (logo preserved) as a zip download. One message, done.

**Option B — Claude Code (files produced automatically):**
1. Drop the order into chat (or `samples/*.json` / a YCSX .xlsx).
2. Claude extracts fields, detects the bag family, asks *"Số màu in?"*.
3. Claude runs `python generate.py --order <file> --colors N` and returns the 2 `.xlsx`.

**Option C — one-click on your PC (no chat):**
Double-click **`don_hang.bat`**. It asks the order file (drag the YCSX/`.json` in) and
*"Mặt hàng này in bao nhiêu màu?"* — then runs `generate.py` and opens the output folder.

## Command reference
```
python generate.py --sample --colors 3                 # built-in 40KG OPP demo
python generate.py --order samples/25kg_tan_chau.json --colors 2   # 25KG paper demo
python generate.py --order my_order.json              # asks "số màu in?" interactively
python generate.py --ycsx "C:/path/YCSX.xlsx" --colors 4          # parse a YCSX file
python generate.py --ycsx "order.xlsx" --colors 3 --outdir out    # custom output dir
python make_kit.py                                    # rebuild dinh_muc_kit.zip
```

## Tuning (when a number doesn't match your plant)
All coefficients live in `data/constants.json`:
- `in_opp.mang_bopp_coeff` (film kg) — empirical; 0.01584 matches the 40KG example.
- `dan.keo_coeff_opp` (glue) — 3.2912 tuned to the 40KG example (was 3.114 in earlier docs).
- `dung_sai_default` (5%), định lượng mành/giấy, ratios — adjust per your records.
Add a new customer exception in `data/customer_rules.json`.

## Verification (regression-passing)
- OPP sample: sl_in=2983, màng BOPP=49.1407, dung môi OPP=12.5286/EA=5.3694.
- Paper sample: sl_in=4600, giấy kraft=328.44, glue=15.2444 (9415 rule).
- Generated Định mức preserves all 12 embedded images and opens identically to the example.
