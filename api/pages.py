# -*- coding: utf-8 -*-
"""Server-rendered HTML for the Định mức review board (Ecolar design system).

Two screens built on the same CSS variables as api/index.py:
  * board_page(orders)   — 5 columns mirroring the Base.vn workflow + import card
  * review_page(row)     — order info + computed BOM + warnings + Accept/Reject

Identity is lightweight: the reviewer's name is kept in a cookie and echoed
back into every action form so Accept/Reject carry an accountable name.
"""
from __future__ import annotations

import html
from typing import Any, Optional

import _logo_data  # same api/ package — LOGO_B64 keeps the header brand

LOGO_B64 = _logo_data.LOGO_B64
CURRENT_REVIEWER_COOKIE = "dinh_muc_reviewer"

# (key, title) — order of the columns on the board.
STAGES = [
    ("thong_tin", "Thông tin đơn hàng"),
    ("ke_toan", "Kế toán kiểm tra"),
    ("qc", "QC kiểm tra"),
    ("dinh_muc", "Lập định mức NVL"),
    ("chuan_bi", "Chuẩn bị NVL"),
]
STAGE_KEYS = [k for k, _ in STAGES]
STAGE_TITLES = dict(STAGES)

# BOM rows shown to kế toán on the review screen: (label, field, fmt).
# fmt "num" = right-aligned number, "text" = plain string.
_BOM_ROWS = [
    ("Số lượng in (m)", "sl_in_thuc_te_m", "num"),
    ("Số lượng tráng (m)", "trang_sl", "num"),
    ("Thành phẩm dự kiến", "thanh_pham_du_kien", "num"),
    ("Kích thước SX (mm)", "size_sx_mm", "text"),
    ("Kích thước TP (mm)", "size_tp_mm", "text"),
    ("Khổ mành (m)", "kho_manh", "num"),
    ("Khổ màng (m)", "kho_mang", "num"),
    ("Khổ giấy (m)", "kho_giay", "num"),
    ("Mành trắng (kg)", "manh_kg", "num"),
    ("Màng BOPP (kg)", "mang_bopp_kg", "num"),
    ("Dung môi OPP (kg)", "dung_moai_opp_kg", "num"),
    ("Dung môi EA (kg)", "dung_moai_ea_kg", "num"),
    ("Giấy Kraft (kg)", "giay_kraft_kg", "num"),
    ("Keo dán (kg)", "glue_total_kg", "num"),
    ("Chỉ may (kg)", "chi_may_kg", "num"),
    ("Dây bó bao (kg)", "day_bo_bao_kg", "num"),
    ("Túi PE LDPE (kg)", "tui_g17", "num"),
    ("Túi PE taical (kg)", "tui_g18", "num"),
    ("Bao/kiện", "bao_kien_text", "text"),
    ("Số màu in", "so_mau_in", "num"),
]

# Vật tư tên/mã pairs to help kế toán cross-check the BOM.
_MATERIAL_ROWS = [
    ("Mành trắng", "manh_ten", "manh_ma"),
    ("Màng BOPP", "mang_ten", "mang_ma"),
    ("Giấy Kraft", "giay_ten", "giay_ma"),
    ("Túi PE", "tui_ten", "tui_ma"),
]

# Info fields shown in the "Thông tin đơn hàng" card: (label, key).
_INFO_ROWS = [
    ("Khách hàng", "customer"),
    ("Số đơn hàng", "order_id"),
    ("Mã sản phẩm", "product_code"),
    ("Số lượng", "qty"),
    ("Số màu in", "so_mau_in"),
    ("Họ bao", "family"),
    ("Ngày tạo", "created_at"),
    ("Người tạo", "reviewer"),
]


def esc(value: Any) -> str:
    return html.escape(str(value))


def _fmt_num(value: Any) -> str:
    if value is None or value == "":
        return ""
    try:
        num = float(value)
    except (TypeError, ValueError):
        return esc(value)
    if num == int(num):
        return str(int(num))
    return f"{num:g}"


# ---------------------------------------------------------------- CSS (shared)
_CSS = """
:root{
  --primary:#A3D977; --primary-strong:#84B94E; --primary-deep:#4E7A2B;
  --primary-soft:#EAF5D8; --primary-wash:#F4F9EA; --honey:#F4B740;
  --surface:#FFFFFF; --surface-alt:#F7F9F4;
  --text:#141A14; --text-muted:#5C6B5C; --text-accent:#4E7A2B;
  --border:#E4E8DF;
  --success-bg:#EAF7EE; --success-ink:#1F6B3A;
  --danger-bg:#FBEDED; --danger-ink:#8A2424;
  --warning-bg:#FFF7E6; --warning-ink:#8A5A00; --warning-border:#F4D9A0;
  --ring:#A3D977;
  --s-1:4px; --s-2:8px; --s-3:12px; --s-4:16px; --s-5:20px; --s-6:24px; --s-8:32px; --s-10:40px;
  --shadow-card:0 1px 2px rgba(20,26,20,.04), 0 12px 32px -16px rgba(20,26,20,.14);
  --shadow-btn:0 1px 2px rgba(20,26,20,.10), inset 0 -1px 0 rgba(0,0,0,.06);
  --shadow-btn-hover:0 8px 18px -6px rgba(132,185,78,.50), inset 0 -1px 0 rgba(0,0,0,.06);
  --ring-focus:0 0 0 3px var(--surface), 0 0 0 5px var(--ring);
  --ease:cubic-bezier(.2,.65,.25,1); --dur-fast:140ms; --dur:200ms;
}
*,*::before,*::after{box-sizing:border-box}
html{font-family:'Be Vietnam Pro',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;font-feature-settings:'tnum' 1;color:var(--text);background:var(--surface-alt)}
body{margin:0;min-height:100vh;line-height:1.5;-webkit-font-smoothing:antialiased}
img{max-width:100%}
a{color:var(--text-accent)}
.container{width:100%;max-width:1120px;margin-inline:auto;padding-inline:20px}

/* HEADER */
.site-header{background:rgba(255,255,255,.82);backdrop-filter:blur(10px) saturate(140%);-webkit-backdrop-filter:blur(10px) saturate(140%);border-bottom:1px solid var(--border)}
.site-header__inner{display:flex;align-items:center;min-height:132px;padding:16px 0;padding-inline:clamp(20px,4vw,40px)}
.brand{display:flex;flex-direction:column;align-items:flex-start;gap:4px;text-decoration:none}
.brand__logo{height:120px;width:auto;display:block}
.brand__tagline{margin:0;font-size:11px;line-height:16px;font-weight:600;text-transform:uppercase;letter-spacing:.12em;color:var(--text-muted)}
.tagline__dot{display:inline-block;color:var(--honey);margin:0 2px;transform:translateY(-1px)}
.site-nav{margin-left:auto;display:flex;gap:8px;align-items:center}
.site-nav a{font-size:13px;line-height:18px;font-weight:600;color:var(--text-accent);text-decoration:none;border:1px solid var(--border);background:var(--surface);padding:8px 14px;border-radius:999px;transition:background var(--dur) var(--ease),border-color var(--dur) var(--ease)}
.site-nav a:hover,.site-nav a.is-active{background:var(--primary-soft);border-color:var(--primary)}

/* LAYOUT */
.board{padding-top:32px}
.board-head{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:20px}
.board-head h1{margin:0;font-size:24px;line-height:30px;font-weight:700;letter-spacing:-.015em}
.card__eyebrow{margin:0 0 4px;font-size:11px;line-height:16px;font-weight:700;text-transform:uppercase;letter-spacing:.10em;color:var(--text-accent)}
.card__sub{margin:6px 0 0;font-size:13.5px;line-height:20px;color:var(--text-muted);max-width:640px}
.reviewer-form{display:flex;flex-direction:column;gap:4px;font-size:12px;color:var(--text-muted)}
.reviewer-form input,.reviewer-form textarea{border:1px solid var(--border);border-radius:9px;padding:8px 10px;font:inherit;font-size:13px;background:var(--surface)}
.reviewer-form input{min-width:200px}
.reviewer-form input:focus,.reviewer-form textarea:focus{outline:0;box-shadow:var(--ring-focus)}

/* IMPORT CARD */
.import-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;box-shadow:var(--shadow-card);margin-bottom:24px}
.import-card summary{cursor:pointer;font-weight:600;font-size:14px;padding:14px 18px;list-style:none;display:flex;align-items:center;gap:8px}
.import-card summary::-webkit-details-marker{display:none}
.import-card summary::before{content:'+';color:var(--primary-deep);font-weight:700}
.import-card[open] summary::before{content:'–'}
.import-form{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:4px 18px 18px}
.import-form input[type=file]{font:inherit;font-size:13px}
.import-form input[type=number]{width:76px;border:1px solid var(--border);border-radius:9px;padding:8px 10px;font:inherit;font-size:13px}
.import-form label{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:600}
.btn{font:inherit;cursor:pointer;border:0;border-radius:10px;padding:10px 16px;font-size:14px;line-height:20px;font-weight:600;box-shadow:var(--shadow-btn);transition:background var(--dur) var(--ease),transform var(--dur-fast) var(--ease),box-shadow var(--dur) var(--ease)}
.btn:hover{transform:translateY(-1px)}
.btn:focus-visible{box-shadow:var(--ring-focus);outline:0}
.btn--primary{background:var(--primary-strong);color:var(--text)}
.btn--primary:hover{background:#78AC3F;box-shadow:var(--shadow-btn-hover)}
.btn--success{background:var(--success-bg);border:1px solid #BFE6CC;color:var(--success-ink)}
.btn--success:hover{background:#DDF4E4}
.btn--danger{background:var(--danger-bg);border:1px solid #F2C2C2;color:var(--danger-ink)}
.btn--danger:hover{background:#F7DBDB}
.btn--ghost{background:var(--surface);border:1px solid var(--border);color:var(--text-accent)}
.btn--ghost:hover{background:var(--primary-soft)}
.btn[disabled]{opacity:.55;cursor:not-allowed;transform:none}

/* BOARD COLUMNS */
.board-cols{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:14px;align-items:start}
.board-col{background:var(--surface-alt);border:1px solid var(--border);border-radius:14px;padding:12px;min-width:0}
.board-col__head{display:flex;align-items:center;gap:8px;margin:0 2px 10px}
.board-col__head h2{margin:0;font-size:13px;line-height:18px;font-weight:700}
.board-col__count{margin-left:auto;font-size:11px;font-weight:700;color:var(--text-muted);background:var(--surface);border:1px solid var(--border);border-radius:999px;min-width:22px;height:22px;display:flex;align-items:center;justify-content:center;padding:0 6px}
.order-card{display:block;text-decoration:none;background:var(--surface);border:1px solid var(--border);border-radius:11px;box-shadow:var(--shadow-card);padding:12px;margin-bottom:10px;transition:border-color var(--dur) var(--ease),transform var(--dur-fast) var(--ease)}
.order-card:hover{border-color:var(--primary);transform:translateY(-1px)}
.order-card__name{margin:0;font-size:13px;line-height:18px;font-weight:600;color:var(--text);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.order-card__meta{margin:5px 0 0;font-size:12px;line-height:16px;color:var(--text-muted)}
.order-card__meta b{color:var(--text)}
.order-card__badges{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
.badge{font-size:11px;line-height:16px;font-weight:600;border-radius:999px;padding:2px 8px}
.badge--accept{background:var(--success-bg);color:var(--success-ink)}
.badge--warn{background:var(--warning-bg);color:var(--warning-ink)}
.badge--reject{background:var(--danger-bg);color:var(--danger-ink)}

/* REVIEW PAGE */
.review{padding-top:32px}
.review-head{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:20px}
.review-head h1{margin:0;font-size:24px;line-height:30px;font-weight:700;letter-spacing:-.015em}
.back-link{font-size:13px;font-weight:600;text-decoration:none;display:inline-block;margin-bottom:10px}
.review-grid{display:grid;grid-template-columns:340px 1fr;gap:16px;align-items:start}
.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;box-shadow:var(--shadow-card);padding:20px;margin-bottom:16px}
.card h3{margin:0 0 14px;font-size:15px;line-height:20px;font-weight:700}
.card__kicker{margin:0 0 4px;font-size:11px;line-height:16px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--text-accent)}
.kv{width:100%;border-collapse:collapse}
.kv th,.kv td{text-align:left;vertical-align:top;padding:7px 8px;border-bottom:1px solid var(--border);font-size:13px;line-height:18px}
.kv th{width:44%;color:var(--text-muted);font-weight:500}
.kv td{font-weight:600;color:var(--text);word-break:break-word}
.kv td.num{text-align:right;font-variant-numeric:tabular-nums}
.stage-badge{font-size:12px;font-weight:700;border-radius:999px;padding:5px 12px;background:var(--primary-soft);color:var(--text-accent);border:1px solid var(--primary)}
.warnings{list-style:none;margin:0;padding:0}
.warnings li{position:relative;padding:9px 10px 9px 30px;background:var(--warning-bg);border:1px solid var(--warning-border);color:var(--warning-ink);border-radius:9px;margin-bottom:8px;font-size:12.5px;line-height:18px}
.warnings li::before{content:'⚠';position:absolute;left:10px;top:8px}
.reject-reason{margin:10px 0 0;padding:9px 10px;background:var(--danger-bg);border:1px solid #F2C2C2;color:var(--danger-ink);border-radius:9px;font-size:12.5px;line-height:18px}
.actions{display:flex;flex-direction:column;gap:10px}
.actions .reviewer-form{margin-bottom:2px}
.action-row{display:flex;gap:10px;flex-wrap:wrap}
.status-dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px;vertical-align:middle}
.status-dot--ok{background:var(--primary-strong)}
.status-dot--warn{background:var(--honey)}
.status-dot--reject{background:#C0392B}

@media(max-width:980px){
  .board-cols{grid-template-columns:repeat(2,minmax(0,1fr))}
  .review-grid{grid-template-columns:1fr}
}
@media(max-width:560px){
  .board-cols{grid-template-columns:1fr}
  .container{padding-inline:16px}
  .brand__logo{height:88px}
}
@media(prefers-reduced-motion:reduce){
  *{animation-duration:.001ms!important;transition-duration:.001ms!important}
  .btn:hover{transform:none}
}
"""

# Full CSS including the base design system above.
_SHELL = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>__CSS__</style>
</head>
<body>
  <header class="site-header">
    <div class="site-header__inner">
      <a class="brand" href="/" aria-label="Ecolar — trang chủ">
        <img class="brand__logo" src="data:image/png;base64,__LOGO__" alt="Ecolar" height="813" width="1200">
        <p class="brand__tagline">LỜI SỐNG XANH <span class="tagline__dot" aria-hidden="true">•</span> BỀN VỮNG</p>
      </a>
      <nav class="site-nav" aria-label="Điều hướng">
        <a href="/" __NAV_HOME__>Tạo Định mức</a>
        <a href="/board" __NAV_BOARD__>Bảng đơn hàng</a>
      </nav>
    </div>
  </header>
  <main class="container __MAIN_CLASS__" id="main">
__BODY__
  </main>
  <footer class="site-footer" style="border-top:1px solid var(--border);padding:24px 20px;margin-top:48px">
    <div class="container" style="text-align:center">
      <span style="font-size:12px;line-height:18px;color:var(--text-muted)">© 2026 Ecolar · Lời sống xanh bền vững</span>
    </div>
  </footer>
  <script>
  (function(){
    var KEY = '__COOKIE__';
    function readCookie(){
      var m = document.cookie.match(new RegExp('(?:^|; )' + KEY + '=([^;]*)'));
      return m ? decodeURIComponent(m[1]) : '';
    }
    var inputs = document.querySelectorAll('input[name="reviewer"]');
    for (var i = 0; i < inputs.length; i++){ inputs[i].value = readCookie(); }
    var visible = document.getElementById('reviewer-input');
    if (visible){
      visible.addEventListener('input', function(){
        var v = encodeURIComponent(visible.value);
        document.cookie = KEY + '=' + v + '; path=/; max-age=31536000; SameSite=Lax';
      });
    }
  })();
  </script>
</body>
</html>
"""


def _shell(title: str, body: str, active: str, main_class: str = "board") -> str:
    nav_home = 'class="is-active"' if active == "home" else ""
    nav_board = 'class="is-active"' if active == "board" else ""
    out = (_SHELL
           .replace("__TITLE__", title)
           .replace("__CSS__", _CSS)
           .replace("__LOGO__", LOGO_B64)
           .replace("__NAV_HOME__", nav_home)
           .replace("__NAV_BOARD__", nav_board)
           .replace("__MAIN_CLASS__", main_class)
           .replace("__BODY__", body)
           .replace("__COOKIE__", CURRENT_REVIEWER_COOKIE))
    return out


# ------------------------------------------------------------------ board page
def _stage_card(order: dict) -> str:
    qty = _fmt_num(order.get("qty"))
    colors = order.get("so_mau_in")
    badges = []
    if order.get("accepted_by"):
        badges.append('<span class="badge badge--accept">Đã duyệt</span>')
    if order.get("reject_reason"):
        badges.append('<span class="badge badge--reject">Bị trả lại</span>')
    if order.get("warnings"):
        n = len(order["warnings"]) if isinstance(order.get("warnings"), list) else 1
        badges.append(f'<span class="badge badge--warn">{n} cảnh báo</span>')
    return (f'<a class="order-card" href="/orders/{esc(order["id"])}">'
            f'<p class="order-card__name">{esc(order.get("product_name") or "—")}</p>'
            f'<p class="order-card__meta"><b>{esc(order.get("customer") or "—")}</b> · '
            f'{esc(order.get("order_id") or "—")}</p>'
            f'<p class="order-card__meta">SL <b>{esc(qty)}</b> · '
            f'số màu in <b>{esc(colors) if colors is not None else "—"}</b></p>'
            + (f'<div class="order-card__badges">{"".join(badges)}</div>' if badges else "")
            + "</a>")


def board_page(orders: list) -> str:
    by_stage: dict[str, list] = {k: [] for k in STAGE_KEYS}
    for order in orders:
        stage = order.get("stage") or STAGE_KEYS[0]
        by_stage.setdefault(stage, []).append(order)

    columns = []
    for key, title in STAGES:
        cards = "".join(_stage_card(o) for o in by_stage.get(key, []))
        columns.append(
            f'<section class="board-col" aria-label="{esc(title)}">'
            f'<div class="board-col__head"><h2>{esc(title)}</h2>'
            f'<span class="board-col__count">{len(by_stage.get(key, []))}</span></div>'
            f'{cards}</section>')

    import_form = (
        '<section class="import-card" aria-label="Tạo đơn mới">'
        '<details>'
        '<summary>Tạo đơn mới — tải file &amp; chọn số màu in</summary>'
        '<form class="import-form" method="post" action="/orders" enctype="multipart/form-data">'
        '<label>File đơn <input type="file" name="file" accept=".json,.xlsx" required></label>'
        '<label>Số màu in <input type="number" name="colors" min="0" max="12" value="3" required></label>'
        '<input type="hidden" name="reviewer" value="">'
        '<button class="btn btn--primary" type="submit">Import đơn</button>'
        '</form>'
        '</details>'
        '</section>')

    body = (
        '<div class="board-head">'
        '<div>'
        '<p class="card__eyebrow">Công cụ · Định mức</p>'
        '<h1>Bảng đơn hàng</h1>'
        '<p class="card__sub">Nhập đơn hàng ở cột đầu, kế toán kiểm tra định mức rồi duyệt '
        'vào “Lập định mức NVL”. Mỗi cột tương ứng một bước trong quy trình Base.vn.</p>'
        '</div>'
        '<label class="reviewer-form">Người kiểm tra'
        '<input id="reviewer-input" name="reviewer" type="text" autocomplete="name" '
        'placeholder="Nhập tên để ghi vào biên bản duyệt"></label>'
        '</div>'
        + import_form
        + f'<section class="board-cols">{"".join(columns)}</section>')
    return _shell("Bảng đơn hàng · Auto Định Mức", body, active="board")


# ---------------------------------------------------------------- review page
def _info_table(order: dict) -> str:
    rows = []
    for label, key in _INFO_ROWS:
        value = order.get(key)
        if key == "qty":
            value = _fmt_num(value)
        elif key == "family":
            value = "Bao BOPP (OPP)" if value == "opp" else "Bao giấy (KP)" if value == "paper_kp" else value
        elif value is None or value == "":
            continue
        rows.append(f'<tr><th>{esc(label)}</th><td>{esc(value)}</td></tr>')
    if order.get("reject_reason"):
        rows.append('<tr><th>Lý do trả lại</th>'
                    f'<td><span class="reject-reason" style="margin:0">{esc(order.get("reject_reason"))}</span></td></tr>')
    accepted = order.get("accepted_by")
    if accepted:
        rows.append(f'<tr><th>Đã duyệt bởi</th><td>{esc(accepted)} · {esc(order.get("accepted_at") or "")}</td></tr>')
    return f'<table class="kv">{"".join(rows)}</table>' if rows else ""


def _bom_table(fields: dict) -> str:
    rows = []
    for label, key, fmt in _BOM_ROWS:
        value = fields.get(key)
        if value is None or value == "":
            continue
        shown = _fmt_num(value) if fmt == "num" else esc(value)
        rows.append(f'<tr><th>{esc(label)}</th><td class="num">{shown}</td></tr>')
    return f'<table class="kv">{"".join(rows)}</table>' if rows else ""


def _material_table(fields: dict) -> str:
    rows = []
    for label, name_key, code_key in _MATERIAL_ROWS:
        name, code = fields.get(name_key), fields.get(code_key)
        if not name and not code:
            continue
        cells = [esc(name or "")] if name else []
        if code:
            cells.append(f'<code>{esc(code)}</code>')
        rows.append(f'<tr><th>{esc(label)}</th><td>{" · ".join(cells)}</td></tr>')
    return f'<table class="kv">{"".join(rows)}</table>' if rows else ""


def _action_forms(order: dict) -> str:
    stage = order.get("stage") or "thong_tin"
    idx = STAGE_KEYS.index(stage) if stage in STAGE_KEYS else 0
    reviewer_hidden = '<input type="hidden" name="reviewer" value="">'

    forms = []
    # Accept — the one-click finalise into "Lập định mức NVL".
    if stage != "dinh_muc":
        forms.append(
            f'<form method="post" action="/orders/{esc(order["id"])}/accept" class="action-row">'
            f'{reviewer_hidden}'
            '<button class="btn btn--success" type="submit">✔ Duyệt — Lập định mức NVL</button>'
            '</form>')
    # Reject (trả về Thông tin đơn hàng kèm lý do).
    forms.append(
        f'<form method="post" action="/orders/{esc(order["id"])}/reject" class="action-row">'
        f'{reviewer_hidden}'
        '<input type="text" name="reason" placeholder="Lý do trả lại (bắt buộc nếu trả)" style="flex:1;min-width:160px;border:1px solid var(--border);border-radius:9px;padding:8px 10px;font:inherit;font-size:13px">'
        '<button class="btn btn--danger" type="submit">Trả lại</button>'
        '</form>')
    # Generic stage navigation (qc / chuan_bi placeholders + back).
    nav = []
    if idx > 0:
        nav.append(
            f'<form method="post" action="/orders/{esc(order["id"])}/stage" class="action-row">'
            f'{reviewer_hidden}<input type="hidden" name="to" value="{esc(STAGE_KEYS[idx - 1])}">'
            '<button class="btn btn--ghost" type="submit">← Quay lại</button></form>')
    if idx < len(STAGE_KEYS) - 1:
        nav.append(
            f'<form method="post" action="/orders/{esc(order["id"])}/stage" class="action-row">'
            f'{reviewer_hidden}<input type="hidden" name="to" value="{esc(STAGE_KEYS[idx + 1])}">'
            '<button class="btn btn--ghost" type="submit">Chuyển tiếp →</button></form>')
    forms.extend(nav)

    download = (f'<a class="btn btn--primary" style="text-align:center;text-decoration:none" '
                f'href="/orders/{esc(order["id"])}/download">Tải ZIP định mức</a>')
    return (
        '<div class="actions">'
        '<label class="reviewer-form">Người duyệt'
        '<input id="reviewer-input" name="reviewer" type="text" autocomplete="name" '
        'placeholder="Tên sẽ ghi vào biên bản"></label>'
        + "".join(forms)
        + download
        + "</div>")


def review_page(order: dict) -> str:
    stage_title = STAGE_TITLES.get(order.get("stage") or "", order.get("stage") or "")
    products = order.get("fields_json") or []
    if not isinstance(products, list):
        products = []

    bom_cards = []
    for i, product in enumerate(products):
        fields = product.get("fields") if isinstance(product, dict) else {}
        fields = fields or {}
        name = (product.get("product_name") if isinstance(product, dict)
                else f"Sản phẩm {i + 1}") or f"Sản phẩm {i + 1}"
        bom_cards.append(
            f'<div class="card">'
            f'<p class="card__kicker">Định mức NVL · {esc(order.get("family") or "")}</p>'
            f'<h3>{esc(name)}</h3>'
            f'{_bom_table(fields)}'
            f'{_material_table(fields)}'
            '</div>')

    warnings = order.get("warnings") or []
    if not isinstance(warnings, list):
        warnings = []
    warnings_html = ""
    if warnings:
        items = "".join(f"<li>{esc(w)}</li>" for w in warnings)
        warnings_html = (
            f'<div class="card"><h3>⚠ Cảnh báo cần kiểm tra</h3>'
            f'<ul class="warnings">{items}</ul></div>')

    body = (
        '<div class="review-head">'
        '<div>'
        '<a class="back-link" href="/board">← Bảng đơn hàng</a>'
        f'<h1>{esc(order.get("product_name") or "Đơn hàng")}</h1>'
        f'<span class="stage-badge">{esc(stage_title)}</span>'
        '</div>'
        '</div>'
        '<div class="review-grid">'
        '<div>'
        f'<div class="card"><p class="card__kicker">Thông tin đơn hàng</p>{_info_table(order)}</div>'
        f'{_action_forms(order)}'
        '</div>'
        '<div>'
        + "".join(bom_cards)
        + warnings_html
        + "</div>"
        '</div>')
    return _shell("Đơn hàng · Auto Định Mức", body, active="board", main_class="review")