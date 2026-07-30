import os, sys, json, tempfile, traceback
from pathlib import Path

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import generate

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
import uvicorn

app = FastAPI(title="Auto Định Mức")

HTML_FORM = """
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Auto Định Mức</title>
<style>
  * { box-sizing: border-box; font-family: system-ui, sans-serif; }
  body { max-width: 640px; margin: 2rem auto; padding: 1rem; }
  h1 { font-size: 1.5rem; }
  form { display: flex; flex-direction: column; gap: 1rem; }
  label { font-weight: 600; }
  input, button { padding: .5rem; font-size: 1rem; }
  button { background: #0055ff; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
  button:disabled { opacity: .5; }
  .error { color: #d00; white-space: pre-wrap; font-family: monospace; }
  .info { color: #080; }
</style>
</head>
<body>
  <h1>Auto Định Mức</h1>
  <p>Tạo file Định mức + YCSX từ đơn hàng (.json hoặc .xlsx).</p>
  <form id="form">
    <label for="file">File đơn hàng (JSON / XLSX)</label>
    <input type="file" name="file" accept=".json,.xlsx" required>

    <label for="colors">Số màu in</label>
    <input type="number" name="colors" min="1" max="12" value="3" required>

    <button type="submit" id="btn">Generate & Download ZIP</button>
  </form>
  <div id="msg"></div>
  <script>
    const form = document.getElementById('form');
    const btn = document.getElementById('btn');
    const msg = document.getElementById('msg');
    form.addEventListener('submit', async e => {
      e.preventDefault();
      btn.disabled = true; btn.textContent = 'Generating...';
      msg.className = ''; msg.textContent = '';
      const fd = new FormData(form);
      try {
        const res = await fetch('/generate', { method: 'POST', body: fd });
        if (!res.ok) {
          const err = await res.text();
          msg.className = 'error'; msg.textContent = err;
          return;
        }
        const blob = await res.blob();
        const cd = res.headers.get('Content-Disposition') || '';
        const name = cd.match(/filename="?(.+?)"?$/)?.[1] || 'dinh_muc.zip';
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = name;
        document.body.appendChild(a); a.click(); a.remove();
        URL.revokeObjectURL(url);
        msg.className = 'info'; msg.textContent = 'Download started!';
      } catch (err) {
        msg.className = 'error'; msg.textContent = err;
      } finally {
        btn.disabled = false; btn.textContent = 'Generate & Download ZIP';
      }
    });
  </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_FORM


@app.post("/generate")
async def generate_endpoint(
    file: UploadFile = File(...),
    colors: int = Form(...),
):
    if colors < 1:
        raise HTTPException(400, "Số màu in phải ≥ 1")

    suffix = Path(file.filename).suffix.lower()
    if suffix not in (".json", ".xlsx"):
        raise HTTPException(400, f"Unsupported file type: {suffix}")

    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = os.path.join(tmp, file.filename)
            with open(tmp_path, "wb") as f:
                f.write(await file.read())

            source = ("ycsx", tmp_path) if suffix == ".xlsx" else ("order", tmp_path)
            res = generate.run(source, colors, outdir=os.path.join(tmp, "out"))

            zip_path = res["outputs"][0]
            zip_name = os.path.basename(zip_path)

            return FileResponse(
                zip_path,
                media_type="application/zip",
                filename=zip_name,
            )
    except generate.InputValidationError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(500, traceback.format_exc())
