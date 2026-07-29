# Setup guide — Định mức + YCSX automation

This kit automates: **customer order → Định mức workbook + YCSX form (.xlsx)**.
The only thing ever asked of you is **"số màu in?"** (how many print colors).

## Architecture (hybrid)
- A **claude.ai website Project** holds the instructions + knowledge (the "brain").
- A small Python engine (`generate.py` + `xlsxpatch.py`) **writes the exact .xlsx** by
  cloning the company template and filling only the input cells. It preserves every
  logo, merge, and style (openpyxl is NOT used to save — it would drop the images).
- The engine runs in **Claude Code or on your PC** (claude.ai chat can't write binary
  xlsx). The Project tells you the values + the exact command to run.

## Files in this kit
```
auto-dinh-muc/
  generate.py            # the engine: order -> 2 xlsx (asks only "số màu in?")
  xlsxpatch.py           # byte-faithful .xlsx editor (preserves images/styles)
  colour_25kg.py         # one-off: painted the 25KG file's inputs yellow
  PROJECT_INSTRUCTIONS.md# paste into the claude.ai Project's Custom Instructions
  don_hang.bat / tao_don.py # 1-cú nhấp (Option C): hỏi file đơn + số màu in, chạy generate.py
  README_setup.md        # this file
  data/                  # the brain: constants, formulas, cell-map, rules, inks, dims
    constants.json standard_dims.json cell_map.json
    customer_rules.json bag_type_map.json ink_master.json
  knowledge/             # upload these to the claude.ai Project as Knowledge
    scenario_and_formulas.md
  templates/             # clone sources (the 2 Định mức examples + YCSX)
  samples/               # reference orders (25kg_tan_chau.json) + the built-in 40KG sample
  output/<date>/         # generated files land here
```

## One-time setup
1. **Install Python 3 + openpyxl** (only used to *parse* an incoming YCSX):
   `python -m pip install openpyxl`
2. **Create the claude.ai Project** (https://claude.ai → Projects → New):
   - **Custom Instructions**: paste the contents of `PROJECT_INSTRUCTIONS.md`.
   - **Knowledge (Project knowledge)**: upload `knowledge/scenario_and_formulas.md`
     and the `data/*.json` files. (Keep each file < claude.ai's size limit.)
3. Keep this whole `auto-dinh-muc/` folder on your PC (or in a Claude Code workspace).

## Daily use
**Option A — in Claude Code (files produced automatically):**
1. Drop the order into chat (or a `samples/*.json` / a YCSX .xlsx).
2. Claude extracts fields, detects the bag family, asks **"Số màu in?"**.
3. Claude runs `python generate.py --order <file> --colors N` and returns the 2 `.xlsx`.

**Option B — in the claude.ai website Project (Vietnamese, 2 turns):**
The Project's Custom Instructions (`PROJECT_INSTRUCTIONS.md`) lock it to a strict 2-turn flow:
1. **Lượt 1:** you paste the order → Claude detects the bag family and asks **only**
   *"Mặt hàng này in bao nhiêu màu? (số màu in)"*.
2. **Lượt 2:** you answer N → Claude replies with **one** Vietnamese message: the inputs, the
   computed material norms, a downloadable `order.json`, and one ready command
   (`python generate.py --order order.json --colors N`). Then it stops — no extended chat.
3. On your PC run that command → collect the 2 `.xlsx` from `output/<date>/`.

> ⚠️ The website can't emit the binary `.xlsx` (it can't run code), so the `.xlsx` step happens
> on your PC via that one line. If you want zero typing, use `don_hang.bat` (Option C below).

**Option C — one-click on your PC (no typing, no chat):**
Double-click **`don_hang.bat`**. It asks two things in Vietnamese — the order file (drag the
YCSX/`.json` in) and *"Mặt hàng này in bao nhiêu màu?"* — then runs `generate.py` and opens the
output folder with the 2 `.xlsx`. This is the true "2 files, no conversation" path.

## Command reference
```
python generate.py --sample --colors 3                 # built-in 40KG OPP demo
python generate.py --order samples/25kg_tan_chau.json --colors 2   # 25KG paper demo
python generate.py --order my_order.json              # asks "số màu in?" interactively
python generate.py --ycsx "C:/path/YCSX.xlsx" --colors 4          # parse a YCSX file
```

## Tuning (when a number doesn't match your plant)
All coefficients live in `data/constants.json`:
- `in_opp.mang_bopp_coeff` (film kg) — empirical; 0.01584 matches the 40KG example.
- `dan.keo_coeff_opp` (glue) — 3.2912 tuned to the 40KG example (Quy trình text says 3.114).
- `dung_sai_default` (5%), định lượng mành/giấy, ratios — adjust per your records.
Add a new customer exception in `data/customer_rules.json`.

## Verification (already passing)
- OPP sample reproduces: sl_in 2983, màng BOPP 49.1407, dung môi OPP 12.5286 / EA 5.3694.
- Paper sample reproduces: sl_in 4600, giấy kraft 328.44, glue 15.2444 (9415 rule).
- Generated Định mức preserves all 12 embedded images and opens identically to the example.
- Coloured 25KG file: 294 input cells painted yellow (matches the 40KG highlight set).

## Phase 2 (not built yet)
- Email-in wiring: Base.vn/n8n webhook → headless `generate.py` → reply with the 2 files
  attached (the `HuongDan_AutoDonThuoc_BaseVN_n8n.docx` documents that plumbing).
