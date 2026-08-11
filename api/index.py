import os, sys, json, tempfile, traceback, shutil, base64
from pathlib import Path

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)  # so the sibling _logo_data module is importable

import generate
import base_columns
import base_vn
import dinh_muc_service
import _logo_data  # LOGO_B64 — base64 of logo.png, bundled alongside this function

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from starlette.background import BackgroundTask
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request

FUNCTION_PATH = "/api/index.py"

LOGO_B64 = _logo_data.LOGO_B64


class _VercelPathMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        path = request.scope.get("path", "")
        if path.startswith(FUNCTION_PATH):
            remainder = path[len(FUNCTION_PATH):]
            request.scope["path"] = remainder or "/"
        return await call_next(request)


app = FastAPI(title="Auto Định Mức")
app.add_middleware(_VercelPathMiddleware)

# Full HTML document with inline <style> and <script>. The CSS is brace-heavy,
# so the logo is injected with a plain .replace() on the {{LOGO_B64}} sentinel
# (NOT an f-string / .format, which would require escaping every CSS brace).
_TEMPLATE = """
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Auto Định Mức · Ecolar</title>
<meta name="description" content="Tạo file Định mức + YCSX từ đơn hàng (.json hoặc .xlsx) — Ecolar">
<meta name="theme-color" content="#A3D977">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{
  --primary:#A3D977; --primary-strong:#84B94E; --primary-deep:#4E7A2B;
  --primary-soft:#EAF5D8; --primary-wash:#F4F9EA; --honey:#F4B740;
  --surface:#FFFFFF; --surface-alt:#F7F9F4;
  --text:#141A14; --text-muted:#5C6B5C; --text-accent:#4E7A2B;
  --border:#E4E8DF;
  --success-bg:#EAF7EE; --success-ink:#1F6B3A;
  --danger-bg:#FBEDED; --danger-ink:#8A2424;
  --ring:#A3D977;
  --s-1:4px; --s-2:8px; --s-3:12px; --s-4:16px; --s-5:20px; --s-6:24px; --s-8:32px; --s-10:40px;
  --shadow-card:0 1px 2px rgba(20,26,20,.04), 0 12px 32px -16px rgba(20,26,20,.14);
  --shadow-btn:0 1px 2px rgba(20,26,20,.10), inset 0 -1px 0 rgba(0,0,0,.06);
  --shadow-btn-hover:0 8px 18px -6px rgba(132,185,78,.50), inset 0 -1px 0 rgba(0,0,0,.06);
  --shadow-drag:0 0 0 5px var(--primary-soft);
  --ring-focus:0 0 0 3px var(--surface), 0 0 0 5px var(--ring);
  --ease:cubic-bezier(.2,.65,.25,1); --dur-fast:140ms; --dur:200ms; --dur-slow:280ms;
}
*,*::before,*::after{box-sizing:border-box}
html{font-family:'Be Vietnam Pro',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;font-feature-settings:'tnum' 1;color:var(--text);background:var(--surface-alt)}
body{margin:0;min-height:100vh;line-height:1.5;-webkit-font-smoothing:antialiased}
img{max-width:100%}
.container{width:100%;max-width:720px;margin-inline:auto;padding-inline:20px}
.main{max-width:540px;padding-top:40px}

.visually-hidden{position:absolute!important;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0} /* focusable: NOT display:none */
.skip-link{position:fixed;top:-48px;left:8px;z-index:50;background:var(--surface);border:1px solid var(--border);padding:8px 12px;border-radius:8px;color:var(--text);text-decoration:none;font-size:13px;transition:top var(--dur) var(--ease)}
.skip-link:focus{top:8px;box-shadow:var(--ring-focus)}

/* HEADER */
.site-header{position:static;background:rgba(255,255,255,.82);backdrop-filter:blur(10px) saturate(140%);-webkit-backdrop-filter:blur(10px) saturate(140%);border-bottom:1px solid var(--border)}
.site-header__inner{display:flex;align-items:center;min-height:132px;padding:16px 0;padding-inline:clamp(20px,4vw,40px)}
.brand{display:flex;flex-direction:column;align-items:flex-start;gap:4px;text-decoration:none}
.brand__logo{height:120px;width:auto;display:block}
.brand__tagline{margin:0;font-size:11px;line-height:16px;font-weight:600;text-transform:uppercase;letter-spacing:.12em;color:var(--text-muted)}
.tagline__dot{display:inline-block;color:var(--honey);margin:0 2px;transform:translateY(-1px)}

/* CARD */
.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;box-shadow:var(--shadow-card);padding:28px;margin:0 auto}
.card__eyebrow{margin:0 0 6px;font-size:11px;line-height:16px;font-weight:700;text-transform:uppercase;letter-spacing:.10em;color:var(--text-accent)}
.card__title{margin:0 0 6px;font-size:20px;line-height:26px;font-weight:700;letter-spacing:-.015em}
.card__sub{margin:0 0 24px;font-size:13.5px;line-height:20px;color:var(--text-muted)}
.form{display:flex;flex-direction:column;gap:20px}
.field{display:flex;flex-direction:column;gap:8px}
.field__head{display:flex;align-items:baseline;justify-content:space-between;gap:8px}
.field__label{font-size:13px;line-height:18px;font-weight:600}
.field__hint{font-size:12px;line-height:16px;color:var(--text-muted)}

/* DROPZONE */
.dropzone{display:flex;flex-direction:column;align-items:center;text-align:center;gap:6px;padding:24px 20px;border:1.5px dashed var(--border);border-radius:12px;background:var(--surface-alt);cursor:pointer;color:var(--text-muted);transition:border-color var(--dur) var(--ease),background var(--dur) var(--ease),box-shadow var(--dur) var(--ease),transform var(--dur-fast) var(--ease)}
.dropzone__icon{color:var(--primary-deep)}
.dropzone__text{font-size:14px;line-height:20px;font-weight:500;color:var(--text)}
.dropzone__text strong{font-weight:600}
.dropzone__types{font-size:12px;line-height:16px}
.dropzone:hover,.dropzone:focus-within,.dropzone.is-dragover{border-color:var(--primary);background:var(--primary-soft)}
.dropzone.is-dragover{border-style:solid;box-shadow:var(--shadow-drag);transform:scale(1.004)}
.dropzone:focus-within{box-shadow:var(--ring-focus)}
.dropzone.has-file{flex-direction:row;align-items:center;justify-content:flex-start;gap:12px;padding:12px 14px;border:1px solid var(--primary);background:var(--surface);text-align:left}
.dropzone__filled{display:none;flex:1;align-items:center;gap:12px;min-width:0}
.dropzone.has-file .dropzone__filled{display:flex}
.dropzone.has-file .dropzone__idle{display:none}
.dropzone__file-icon{flex-shrink:0}
.dropzone__meta{display:flex;flex-direction:column;min-width:0;flex:1}
.dropzone__name{font-size:14px;line-height:20px;font-weight:500;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dropzone__size{font-size:12px;line-height:16px;color:var(--text-muted)}
.dropzone__remove{flex-shrink:0;width:28px;height:28px;border:0;background:transparent;color:var(--text-muted);font-size:18px;line-height:1;border-radius:8px;cursor:pointer;transition:background var(--dur-fast) var(--ease),color var(--dur-fast) var(--ease)}
.dropzone__remove:hover{background:var(--danger-bg);color:var(--danger-ink)}
.dropzone__remove:focus-visible{box-shadow:var(--ring-focus);outline:0}

/* STEPPER */
.stepper{display:flex;flex-direction:column;gap:12px}
.stepper__readout{width:100%;text-align:center;font-size:30px;line-height:34px;font-weight:700;letter-spacing:-.02em;color:var(--text);background:transparent;border:0;outline:0;padding:0;-moz-appearance:textfield;appearance:textfield}
.stepper__readout::-webkit-inner-spin-button,.stepper__readout::-webkit-outer-spin-button{-webkit-appearance:none;margin:0}
.stepper__slider{--pct:25;width:100%;height:8px;-webkit-appearance:none;appearance:none;background:linear-gradient(90deg,var(--primary) calc(var(--pct)*1%),var(--primary-soft) calc(var(--pct)*1%));border-radius:999px;outline:0;cursor:pointer;accent-color:var(--primary-strong)}
.stepper__slider::-webkit-slider-thumb{-webkit-appearance:none;width:18px;height:18px;border-radius:50%;background:var(--surface);border:2px solid var(--primary-strong);box-shadow:0 2px 6px rgba(94,148,52,.30);transition:transform var(--dur-fast) var(--ease)}
.stepper__slider::-moz-range-thumb{width:18px;height:18px;border-radius:50%;background:var(--surface);border:2px solid var(--primary-strong);box-shadow:0 2px 6px rgba(94,148,52,.30)}
.stepper__slider:active::-webkit-slider-thumb{transform:scale(1.12)}
.stepper__slider:focus-visible{box-shadow:var(--ring-focus)}

/* BUTTON */
.btn{font:inherit;cursor:pointer;border:0}
.btn--primary{position:relative;width:100%;height:46px;border-radius:11px;background:var(--primary-strong);color:var(--text);font-size:14.5px;line-height:20px;font-weight:600;letter-spacing:.01em;box-shadow:var(--shadow-btn);display:inline-flex;align-items:center;justify-content:center;gap:8px;transition:background var(--dur) var(--ease),transform var(--dur-fast) var(--ease),box-shadow var(--dur) var(--ease),opacity var(--dur) var(--ease)}
.btn--primary:hover{background:#78AC3F;transform:translateY(-1px);box-shadow:var(--shadow-btn-hover)}
.btn--primary:active{transform:translateY(0) scale(.99);background:#6C9E36}
.btn--primary:focus-visible{box-shadow:var(--ring-focus);outline:0}
.btn--primary[data-loading="true"]{opacity:.85;cursor:progress;transform:none}
.btn__spinner{display:none;width:16px;height:16px;border-radius:50%;border:2px solid rgba(20,26,20,.25);border-top-color:var(--text);animation:spin .7s linear infinite}
.btn--primary[data-loading="true"] .btn__spinner{display:inline-block}
@keyframes spin{to{transform:rotate(360deg)}}
.btn--primary.is-ping::after{content:'';position:absolute;inset:0;border-radius:11px;box-shadow:0 0 0 0 rgba(132,185,78,.55);animation:ping .4s var(--ease) forwards;pointer-events:none}
@keyframes ping{to{box-shadow:0 0 0 12px rgba(132,185,78,0);opacity:0}}

/* STATUS */
.status{min-height:20px;margin-top:4px;transition:opacity var(--dur) var(--ease),transform var(--dur) var(--ease)}
.status:empty{min-height:20px}
.status.is-success,.status.is-error{padding:12px 14px;border-radius:10px;font-size:13.5px;line-height:20px}
.status.is-success{background:var(--success-bg);border:1px solid #BFE6CC;color:var(--success-ink)}
.status.is-error{background:var(--danger-bg);border:1px solid #F2C2C2;color:var(--danger-ink);flex-direction:column;align-items:stretch}
.status__row{display:flex;gap:8px;align-items:flex-start}
.status__row svg{flex-shrink:0;margin-top:2px}
.status__detail{white-space:pre-wrap;word-break:break-word;font-family:ui-monospace,'SF Mono','Cascadia Code',Consolas,monospace;font-size:12px;line-height:18px;color:var(--danger-ink);max-height:180px;overflow:auto;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:10px;margin:8px 0 0}

/* STEPS */
.steps{margin-top:32px}
.steps__list{list-style:none;padding:0;margin:0;display:grid;grid-template-columns:repeat(3,1fr);gap:24px}
.step{display:flex;gap:12px;align-items:flex-start}
.step__num{flex-shrink:0;width:22px;height:22px;border-radius:999px;background:var(--primary-soft);color:var(--text-accent);font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center}
.step__title{margin:0;font-size:13px;line-height:18px;font-weight:600}
.step__desc{margin:2px 0 0;font-size:12px;line-height:16px;color:var(--text-muted)}

/* FOOTER */
.site-footer{border-top:1px solid var(--border);padding:24px 20px;margin-top:48px}
.site-footer__inner{text-align:center}
.footer__text{font-size:12px;line-height:18px;color:var(--text-muted)}

/* RESPONSIVE */
@media(max-width:560px){
  .main{padding-top:24px;padding-inline:16px}
  .container{padding-inline:16px}
  .card{padding:20px}
  .steps__list{grid-template-columns:1fr!important}
  .brand__logo{height:88px}
}

/* REDUCED MOTION */
@media(prefers-reduced-motion:reduce){
  *{animation-duration:.001ms!important;animation-iteration-count:1!important;transition-duration:.001ms!important;scroll-behavior:auto!important}
  .btn--primary:hover,.dropzone.is-dragover{transform:none}
  .btn__spinner{animation:none}
}
</style>
</head>
<body class="page">
  <a class="skip-link" href="#form">Bỏ qua tới biểu mẫu</a>

  <header class="site-header">
    <div class="site-header__inner">
      <a class="brand" href="/" aria-label="Ecolar — trang chủ">
        <img class="brand__logo" src="data:image/png;base64,{{LOGO_B64}}" alt="Ecolar" height="813" width="1200">
        <p class="brand__tagline">LỜI SỐNG XANH <span class="tagline__dot" aria-hidden="true">•</span> BỀN VỮNG</p>
      </a>
    </div>
  </header>

  <main class="container main" id="main">
    <section class="card" aria-labelledby="card-title">
      <p class="card__eyebrow">Công cụ · Định mức</p>
      <h1 class="card__title" id="card-title">Tạo Định mức + YCSX</h1>
      <p class="card__sub">Tải lên đơn hàng, chọn số màu in và nhận bộ file Excel qua file ZIP.</p>

      <form id="form" class="form" novalidate>
        <!-- FILE FIELD (dropzone wrapping the real native input) -->
        <div class="field">
          <div class="field__head">
            <label class="field__label" for="file">File đơn hàng</label>
            <span class="field__hint">.json hoặc .xlsx</span>
          </div>
          <label class="dropzone" id="dropzone" for="file">
            <input class="visually-hidden" id="file" name="file" type="file" accept=".json,.xlsx" required>
            <span class="dropzone__idle">
              <svg class="dropzone__icon" viewBox="0 0 24 24" width="26" height="26" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4m0 0l-5 5m5-5l5 5"/><path d="M4 17v2a1 1 0 001 1h14a1 1 0 001-1v-2"/></svg>
              <span class="dropzone__text"><strong>Kéo thả file vào đây</strong> hoặc bấm để chọn</span>
              <span class="dropzone__types">Hỗ trợ .json, .xlsx</span>
            </span>
            <span class="dropzone__filled" hidden>
              <svg class="dropzone__file-icon" viewBox="0 0 24 24" width="28" height="28" aria-hidden="true"><path d="M6 2h8l6 6v14H6z" fill="#EAF5D8" stroke="#84B94E" stroke-width="1.4" stroke-linejoin="round"/><path d="M14 2v6h6" fill="none" stroke="#84B94E" stroke-width="1.4" stroke-linejoin="round"/></svg>
              <span class="dropzone__meta">
                <span class="dropzone__name" data-file-name>—</span>
                <span class="dropzone__size" data-file-size>—</span>
              </span>
              <button type="button" class="dropzone__remove" data-remove aria-label="Bỏ file đã chọn">×</button>
            </span>
          </label>
        </div>

        <!-- COLORS FIELD (contract-safe: exact named number input + unnamed slider) -->
        <div class="field">
          <div class="field__head">
            <label class="field__label" for="colors">Số màu in</label>
            <span class="field__hint">0 = không in</span>
          </div>
          <div class="stepper">
            <input class="stepper__readout" id="colors" name="colors" type="number" min="0" max="12" value="3" required inputmode="numeric">
            <input class="stepper__slider" type="range" min="0" max="12" value="3" aria-label="Số màu in">
          </div>
        </div>

        <!-- SUBMIT -->
        <button type="submit" id="btn" class="btn btn--primary" data-rest-label="Tạo file ZIP &amp; tải xuống">
          <span class="btn__spinner" aria-hidden="true"></span>
          <span class="btn__label">Tạo file ZIP &amp; tải xuống</span>
        </button>
      </form>

      <div id="msg" class="status" role="status" aria-live="polite" aria-atomic="true"></div>
    </section>

    <section class="steps" aria-label="Hướng dẫn 3 bước">
      <ol class="steps__list">
        <li class="step"><span class="step__num">1</span><div><p class="step__title">Tải đơn hàng</p><p class="step__desc">File .json hoặc .xlsx từ khách hàng.</p></div></li>
        <li class="step"><span class="step__num">2</span><div><p class="step__title">Chọn số màu in</p><p class="step__desc">Từ 0 (không in) đến 12.</p></div></li>
        <li class="step"><span class="step__num">3</span><div><p class="step__title">Tải ZIP kết quả</p><p class="step__desc">Định mức vật tư + YCSX.</p></div></li>
      </ol>
    </section>
  </main>

  <footer class="site-footer">
    <div class="container site-footer__inner">
      <span class="footer__text">© 2026 Ecolar · Lời sống xanh bền vững</span>
    </div>
  </footer>

  <script>
    const form = document.getElementById('form');
    const btn  = document.getElementById('btn');
    const msg  = document.getElementById('msg');
    const fileInput = document.getElementById('file');
    const dropzone  = document.getElementById('dropzone');
    const colorsEl  = document.getElementById('colors');
    const slider    = document.querySelector('.stepper__slider');
    const btnLabel  = btn.querySelector('.btn__label');
    const REST_LABEL = btn.dataset.restLabel; // "Tạo file ZIP & tải xuống"

    const CHECK = '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>';
    const ALERT = '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4M12 17h.01M10.3 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.7 3.86a2 2 0 00-3.4 0z"/></svg>';

    function setMsg(cls, html){ msg.className = 'status' + (cls ? ' ' + cls : ''); msg.innerHTML = html; }
    function showError(text){ setMsg('is-error', '<div class="status__row">' + ALERT + '<span>Đã có lỗi xảy ra:</span></div><pre class="status__detail"></pre>'); msg.querySelector('.status__detail').textContent = text; }

    /* ---------- DROPZONE: filename + size + remove + DnD ---------- */
    function fmtSize(b){ if (b < 1024) return b + ' B'; if (b < 1024*1024) return (b/1024).toFixed(1) + ' KB'; return (b/1048576).toFixed(2) + ' MB'; }
    function renderFile(){
      const f = fileInput.files[0];
      const idle = dropzone.querySelector('.dropzone__idle');
      const filled = dropzone.querySelector('.dropzone__filled');
      if (f){
        dropzone.classList.add('has-file');
        idle.hidden = true; filled.hidden = false;
        dropzone.querySelector('[data-file-name]').textContent = f.name;
        dropzone.querySelector('[data-file-size]').textContent = fmtSize(f.size);
      } else {
        dropzone.classList.remove('has-file');
        idle.hidden = false; filled.hidden = true;
      }
    }
    fileInput.addEventListener('change', renderFile);
    dropzone.querySelector('[data-remove]').addEventListener('click', function(e){ e.preventDefault(); e.stopPropagation(); fileInput.value = ''; renderFile(); fileInput.focus(); });
    ['dragenter','dragover'].forEach(function(ev){ dropzone.addEventListener(ev, function(e){ e.preventDefault(); dropzone.classList.add('is-dragover'); }); });
    ['dragleave','dragend'].forEach(function(ev){ dropzone.addEventListener(ev, function(){ dropzone.classList.remove('is-dragover'); }); });
    dropzone.addEventListener('drop', function(e){ e.preventDefault(); dropzone.classList.remove('is-dragover'); if (e.dataTransfer.files.length){ fileInput.files = e.dataTransfer.files; renderFile(); } });

    /* ---------- STEPPER: two-way bind slider <-> #colors (slider has NO name) ---------- */
    function setPct(v){ slider.style.setProperty('--pct', (v/12*100)); }
    slider.addEventListener('input', function(){ colorsEl.value = slider.value; setPct(slider.value); });
    colorsEl.addEventListener('input', function(){ var v = Math.max(0, Math.min(12, +colorsEl.value || 0)); slider.value = v; setPct(v); });
    setPct(colorsEl.value); // init

    /* ---------- SUBMIT (existing fetch/ZIP/download logic preserved; +spinner, +VI labels, +status classes) ---------- */
    form.addEventListener('submit', async e => {
      e.preventDefault();
      btn.disabled = true; btn.dataset.loading = 'true'; btnLabel.textContent = 'Đang tạo…';
      setMsg('', '');
      const fd = new FormData(form);
      try {
        const res = await fetch('/generate', { method: 'POST', body: fd });
        if (!res.ok) {
          const err = await res.text();
          showError(err);
          return;
        }
        const blob = await res.blob();
        const cd = res.headers.get('Content-Disposition') || '';
        const name = cd.match(/filename="?(.+?)"?$/)?.[1] || 'dinh_muc.zip';
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = name;
        document.body.appendChild(a); a.click(); a.remove();
        URL.revokeObjectURL(url);
        setMsg('is-success', '<div class="status__row">' + CHECK + '<span>Đã tạo xong! File ZIP đang được tải xuống.</span></div>');
        btn.classList.add('is-ping'); setTimeout(() => btn.classList.remove('is-ping'), 400);
      } catch (err) {
        showError(err);
      } finally {
        btn.disabled = false; btn.dataset.loading = 'false'; btnLabel.textContent = REST_LABEL;
      }
    });
  </script>
</body>
</html>
"""

# Resolve the logo sentinel. Plain .replace (not f-string) so CSS braces are untouched.
HTML_FORM = _TEMPLATE.replace("{{LOGO_B64}}", LOGO_B64)


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_FORM


@app.post("/generate")
async def generate_endpoint(
    file: UploadFile = File(...),
    colors: int = Form(...),
):
    if colors < 0:
        raise HTTPException(400, "Số màu in không được âm")

    suffix = Path(file.filename).suffix.lower()
    if suffix not in (".json", ".xlsx"):
        raise HTTPException(400, f"Unsupported file type: {suffix}")

    tmp = tempfile.mkdtemp()
    try:
        tmp_path = os.path.join(tmp, file.filename)
        with open(tmp_path, "wb") as f:
            f.write(await file.read())

        source = ("ycsx", tmp_path) if suffix == ".xlsx" else ("order", tmp_path)
        res = generate.run(source, colors, outdir=os.path.join(tmp, "out"))

        zip_path = res["outputs"][0]
        zip_name = os.path.basename(zip_path)

        return FileResponse(
            zip_path, media_type="application/zip", filename=zip_name,
            background=BackgroundTask(shutil.rmtree, tmp),
        )
    except generate.InputValidationError as e:
        shutil.rmtree(tmp, ignore_errors=True)
        raise HTTPException(400, str(e))
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        raise HTTPException(500, traceback.format_exc())


# ---------------------------------------------------------------- Base.vn integration
def _job_id_from_payload(payload: dict):
    for container in (payload.get("data") if isinstance(payload.get("data"), dict) else {}, payload):
        for key in ("id", "job_id", "task_id"):
            value = container.get(key)
            if value not in (None, ""):
                return value
    return None


def _colors_from_payload(payload: dict):
    for key in ("colors", "so_mau_in", "so_mau", "custom_so_mau_in"):
        src = payload.get("data") if isinstance(payload.get("data"), dict) else {}
        value = payload.get(key, src.get(key))
        if value not in (None, ""):
            try:
                return int(str(value).strip())
            except (TypeError, ValueError):
                break
    default = os.environ.get("BASE_DEFAULT_COLORS", "3")
    try:
        return int(default)
    except ValueError:
        return 3


def _ycsx_url_from_payload(payload: dict):
    src = payload.get("data") if isinstance(payload.get("data"), dict) else {}
    for key in ("ycsx_url", "file_url", "ycsx", "custom_ycsx_url"):
        for container in (src, payload):
            value = container.get(key)
            if isinstance(value, str) and value:
                return value
    return None


def _ycsx_url_from_job(job: dict):
    """Pick the attached .xls/.xlsx (the YCSX file) from a fetched job."""
    job_obj = job.get("job") or job.get("data") or job
    for f in (job_obj.get("files") or []):
        if isinstance(f, dict):
            name = str(f.get("name") or "").lower()
            if f.get("url") and (name.endswith(".xls") or name.endswith(".xlsx")):
                return f["url"]
    return None


@app.post("/api/wf/receive", status_code=200)
async def wf_receive(request: Request):
    """Base Workflow webhook receiver (auto-execute).

    Body: the workflow's webhook output vars (flat, or job nested under
    ``data``). Resolution order for the order data:

    1. Attached YCSX file (payload ``ycsx_url``/``file_url`` or fetched from
       the job via ``job/get``) -> parsed with generate.parse_ycsx (accurate).
    2. Mapped order columns (``custom_*``) -> order dict path.
    """
    try:
        raw = await request.body()
        payload = json.loads(raw.decode("utf-8") or "{}")
    except (ValueError, UnicodeDecodeError):
        raise HTTPException(400, "Body phai la JSON hop le")

    job_id = _job_id_from_payload(payload)
    colors = _colors_from_payload(payload)
    tmp = tempfile.mkdtemp()
    client = base_vn.BaseVnClient()

    ycsx_url = _ycsx_url_from_payload(payload)
    fetched_job = None
    job_get = getattr(client, "get_job", None)
    if not ycsx_url and job_id and job_get is not None:
        try:
            fetched_job = job_get(job_id)
        except base_vn.BaseVnError:
            fetched_job = None
        if fetched_job:
            ycsx_url = _ycsx_url_from_job(fetched_job)
            if isinstance(fetched_job, dict):
                job_obj = fetched_job.get("job") or fetched_job.get("data") or {}
                merged = {k: v for k, v in job_obj.items() if not isinstance(v, dict)}
                payload = {**payload, **merged}

    try:
        if ycsx_url:
            local = client.download_file(ycsx_url, os.path.join(tmp, "ycsx.xlsx"))
            result = generate.run(("ycsx", local), colors, outdir=os.path.join(tmp, "out"))
            summary = dinh_muc_service.compute_summary(result)
        else:
            order = base_columns.order_from_webhook(payload)
            wrapped = dinh_muc_service.run_and_summarize(order, colors, outdir=os.path.join(tmp, "out"))
            summary = wrapped["summary"]
    except generate.InputValidationError as e:
        shutil.rmtree(tmp, ignore_errors=True)
        raise HTTPException(400, str(e))
    except Exception:
        shutil.rmtree(tmp, ignore_errors=True)
        raise HTTPException(500, traceback.format_exc())

    result_fields = base_columns.result_fields(
        dinh_muc_service.write_result_fields(summary))
    written_back = False
    if job_id:
        try:
            client.update_job(job_id, **result_fields)
            written_back = True
        except base_vn.BaseVnError as e:
            summary["error"] = f"write-back failed: {e}"

    moved_next = False
    auto_move = os.environ.get("BASE_AUTO_MOVE_NEXT", "0") == "1"
    if job_id and auto_move and summary.get("status") == "done":
        try:
            client.move_next(job_id)
            moved_next = True
        except base_vn.BaseVnError as e:
            summary.setdefault("warning", f"move-next failed: {e}")

    return {
        "ok": True,
        "job_id": job_id,
        "colors": colors,
        "ycsx_used": bool(ycsx_url),
        "summary": summary,
        "written_back": written_back,
        "moved_next": moved_next,
    }


@app.post("/api/wf/discover")
async def wf_discover():
    """Read-only probe of the workspace: workflows + their stages."""
    try:
        data = base_vn.BaseVnClient().discover()
    except base_vn.BaseVnError as e:
        raise HTTPException(400, str(e))
    return data
