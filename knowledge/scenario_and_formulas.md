# Scenario & Formulas — Định mức / YCSX (Knowledge for the Claude Project)

## What the two output files are
- **Định mức** workbook: 23 sheets = one per production stage. The key input sheet is
  **"In"** (PHIẾU YÊU CẦU XUẤT KHO NVL — material requisition). Other input sheets:
  **May 1** (lệnh SX, stage May), **Chia biên 2** (chia cuộn), **Dán 2** (dán cắt).
  Sheets ending in **"2"** (In 2, Tráng 2…) hold machine tolerances from the TSVH files.
- **YCSX**: PHIẾU YÊU CẦU SẢN XUẤT (QTKSX-BM01) — header (khách hàng, mã KH, số đơn
  hàng, ngày) + line items (mã SP, tên SP, mã code, **Chi tiết kỹ thuật**, ĐVT, số lượng).

## Order input fields (from the customer order / YCSX)
customer · product_name · product_code · order_id (Số đơn hàng = Số YCSX) ·
qty (bao) · bag_length_m (chiều dài bao sx) · width_plus_gusset_m (chiều ngang+hông) ·
inner_bag_weight_kg · spec (Chi tiết kỹ thuật: Kích thước / Thành phẩm / Cấu trúc /
Quy cách gấp-may / Lồng túi PE / Đóng gói).
**Bag weight/type is read FROM the order** — in the product name (`…40KG`, `25kg`) and
in Cấu trúc (`Bao BOPP in ống đồng`, `Bao KP in Flexo`, `Bao in offset`).

## Derived widths
- khổ mành = (ngang + hông) × 2 + 0.06  (m)
- khổ giấy = (ngang + hông) × 2 + 0.02  (m)
- khổ màng: from the order, else the standard-dims lookup (cm/100).

## Số lượng in thực tế (print metres) — per family (verified on examples)
- **OPP**: qty × L × (1 + dung_sai 5%) + phế chồng màu
  (phế chồng màu: 1 màu=200 m, 2=300, 3=400, 4–6=500)
- **paper/KP**: qty × L  (no tolerance on metres)

## Thành phẩm dự kiến — per family
- **OPP**: = qty  (40KG example: 3000)
- **paper/KP**: = qty × (1 + dung_sai)  (25KG example: 5000 → 5250)

## Material norms per stage (from "Quy trình làm định mức")
**In — OPP / ống đồng:**
- Màng BOPP mờ = sl_in × khổ màng × 0.01584 kg  (coeff tuned; ~density×thickness)
- Dung môi OPP = 0.006 × sl_in × 70%
- Dung môi EA   = 0.006 × sl_in × 30%
- Ink: N color rows (N = số màu in) from the OPP ink master; kg left blank (print shop).

**In — flexo / paper:**
- Giấy kraft = sl_in × khổ giấy × (70 g/m² / 1000)
- Ink: N flexo ink rows; kg blank.

**Tráng:**
- Mành = qty × L × khổ mành × ĐL mành (trắng 70 / trong 72 / **4 Oranges 75** g/m²)
- F801C = qty × L × khổ mành × 0.90 (OPP 0.85) × 0.02
- Taical = qty × L × khổ mành × 0.10 × 0.02  (OPP thêm hạt màu × 0.05 × 0.02)

**Dán (glue — goes into Dán 2 rows 18–20):**
- total = qty × L × coeff / 1000 ; coeff = 3.314 (KP) / 3.2912 (OPP, tuned)
- Standard: F801C 50% + Vistamax 50%.
- **Tân Châu / Neo Nam Việt**: 9415 60% + Vistamax 40% (no F801C).

**Thổi túi lồng:**
- LDPE = qty × inner_bag_weight × dung_sai × 89.3%
- Taical EFPE = qty × inner_bag_weight × dung_sai × 10.7%
- **PE rin**: 100% LDPE, no taical.

**May:**
- Chỉ may = 0.6 kg / 1000 bao · Dây bó bao = 1 kg / 5000 bao
- Nẹp = chiều rộng nẹp × định lượng nẹp × (rộng bao + 12 cm)  (KP 0.16 / OPP 0.106)

## Standard dimensions lookup (cm) — when khổ màng/mành not stated
| ngang+hông | chiều dài | màng in | mành |
|---|---|---|---|
| 32+8 | 55–71 | 84 | 86 |
| 37+8 | 72–88 | 94 | 96 |
| 42+8 | — | 104 | 106 |
| 48+10 | 89–95 | 120 | 122 |
| 48+12 | 98–105 | 124 | 126 |
| 45+15 | — | 124 | 126 |

## Yellow input cells (the cells that change per order — "In" sheet, canonical)
- Header: C5=customer, C6=product_name, C7=product_code, C8=order_id, C9=so_phiếu_sx
- Info table (M col): M6=qty, M7=L, M8=W, M9=sl_in, **M10=số màu in**, M11=khổ mành,
  M12=khổ màng, M13=khổ giấy, M14=inner_bag_weight
- Same again in C39–C47; G29=thành phẩm dự kiến
- Material table A18:G29 (org/stt/name/code/lot/unit/qty)
- May 1 / Chia biên 2 / Dán 2: same header block (C5–C8) + stage values

## Verified regression (must reproduce)
- OPP 4O BTA140, qty 3000, 0.82×0.5 m, 3 màu → sl_in 2983, màng BOPP 49.1407,
  dung môi OPP 12.5286, EA 5.3694.
- Paper Tân Châu AM-20, qty 5000, 0.92×0.5 m, 2 màu → sl_in 4600, giấy kraft 328.44,
  glue total 15.2444 (9415 9.1466 / Vistamax 6.0978).
