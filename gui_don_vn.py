# -*- coding: utf-8 -*-
"""gui_don_vn.py — GUI cho dòng chảy Base.vn: tạo file định mức + gửi lên Workflow.

Chạy như tao_don.py (hỏi file đơn + "số màu in"), tạo file xuống output/<ngay>/,
rồi gọi webhook tạo nhiệm vụ vào workflow "Quản lý đơn hàng sản xuất"
(cột "call định mức nvl").

Cờ dòng lệnh:
  --order PATH         file JSON đơn hàng (thay vì hỏi)
  --colors N           số màu in (thay vì hỏi)
  --outdir PATH        thư mục xuất (mặc định: output/<ngày>/)
  --dry-run            chỉ tạo file, KHÔNG gửi lên Base
  --discover           chỉ probe workflow, in JSON cột, không tạo gì
"""
from __future__ import annotations

import argparse
import datetime
import glob
import json
import os
import subprocess
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import base_columns
import base_vn
import dinh_muc_service
import generate


def ask(prompt):
    return input(prompt).strip().strip('"').strip("'")


def latest_output_dir():
    dirs = sorted(glob.glob(os.path.join(HERE, "output", "*")))
    return dirs[-1] if dirs else None


def discover(client: base_vn.BaseVnClient) -> int:
    print("Probing Base.vn workflows (read-only)...")
    data = client.discover()
    print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Tạo định mức + gửi lên Base Workflow")
    ap.add_argument("--order", default=None, help="đường dẫn file order JSON")
    ap.add_argument("--colors", type=int, default=None, help="số màu in")
    ap.add_argument("--outdir", default=None, help="thư mục xuất file")
    ap.add_argument("--dry-run", action="store_true", help="không gọi Base")
    ap.add_argument("--discover", action="store_true", help="probe workflow, không tạo gì")
    args = ap.parse_args(argv)

    if args.discover:
        return discover(base_vn.BaseVnClient())

    print("=" * 60)
    print(" GUI ĐỊNH MỨC + BASE WORKFLOW ".center(60, "="))
    print("=" * 60)

    path = args.order
    if not path:
        print("Kéo/thả file đơn hàng (JSON) vào đây rồi Enter, hoặc gõ đường dẫn.")
        path = ask("\nDuong dan file don hang: ")
    if not path or not os.path.isfile(path):
        print("KHONG tim thay file:", repr(path))
        return 1

    colors = args.colors
    if colors is None:
        try:
            colors = int(ask("Mat hang nay in bao nhieu mau? (so mau in): "))
        except ValueError:
            print("So mau in phai la so nguyen.")
            return 1
    if colors < 0:
        print("So mau in phai >= 0.")
        return 1
    if colors > generate.SO_MAU_IN_MAX:
        print(f"So mau in toi da la {generate.SO_MAU_IN_MAX}.")
        return 1

    try:
        with open(path, encoding="utf-8") as f:
            order = json.load(f)
    except (OSError, ValueError) as e:
        print("Khong doc duoc file JSON:", e)
        return 1

    outdir = args.outdir or os.path.join(HERE, "output", datetime.date.today().isoformat())
    print(f"\nDang tinh dinh muc (so mau in = {colors})...")
    try:
        wrapped = dinh_muc_service.run_and_summarize(order, colors, outdir=outdir)
    except Exception as e:
        print("\nLOI khi tinh dinh muc:", e)
        return 2
    result, summary = wrapped["result"], wrapped["summary"]
    print("XONG. File trong:", outdir)
    for f in summary.get("outputs") or []:
        print("  -", f)

    if args.dry_run:
        print("\n[dry-run] Khong gui len Base (dua --dry-run).")
    else:
        client = base_vn.BaseVnClient()
        colmap = base_columns.load_column_map()
        fields = base_columns.build_webhook_fields(order, colmap)
        print("\nDang tao nhiem vu tren Base Workflow...")
        resp = client.create_job_via_webhook(fields)
        print("Da tao xong. Phan hoi tu Base:")
        print(json.dumps(resp, ensure_ascii=False, indent=2))
        summary["status"] = "sent"

    out = latest_output_dir()
    if not os.environ.get("GUI_DON_NO_OPEN"):
        try:
            if sys.platform.startswith("win") and outdir:
                os.startfile(outdir)
        except Exception:
            pass
    print("\nXONG.")
    return 0


if __name__ == "__main__":
    sys.exit(main())