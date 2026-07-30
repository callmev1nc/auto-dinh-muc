#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""tao_don.py — tiện ích 1 cú nhấp: hỏi file đơn + "số màu in" rồi chạy generate.py.
Dùng kèm don_hang.bat (nhấp đôi trong Explorer để chạy). Chỉ hỏi đúng 1 câu, xong mở thư mục.
"""
import os, sys, subprocess, glob

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
HERE = os.path.dirname(os.path.abspath(__file__))


def ask(prompt):
    return input(prompt).strip().strip('"').strip("'")


def latest_output_dir():
    dirs = sorted(glob.glob(os.path.join(HERE, "output", "*")))
    return dirs[-1] if dirs else None


def main():
    print("=" * 60)
    print(" TẠO ĐỊNH MỨC + YCSX  (1 cú nhấp) ".center(60, "="))
    print("=" * 60)
    print("Kéo/thả file đơn hàng vào đây rồi Enter (hoặc gõ đường dẫn).")
    print("  • file YCSX .xlsx  -> sẽ phân tích phiếu YCSX")
    print("  • file order .json -> dùng trực tiếp")
    path = ask("\nDuong dan file don hang: ")
    if not path or not os.path.isfile(path):
        print("KHONG tim thay file:", repr(path))
        return 1
    colors = ask("Mat hang nay in bao nhieu mau? (so mau in): ")
    try:
        n = int(colors)
    except ValueError:
        print("So mau in phai la so nguyen.")
        return 1
    if n < 1:
        print("So mau in phai >= 1.")
        return 1

    flag = "--ycsx" if path.lower().endswith((".xlsx", ".xls")) else "--order"
    cmd = [sys.executable, os.path.join(HERE, "generate.py"), flag, path, "--colors", str(n)]
    print("\nDang tao file...\n  " + " ".join(cmd))
    rc = subprocess.call(cmd)
    if rc == 2:
        print("\n(Thieu du lieu dau vao — xem chi tiet phia tren.)")
        return rc
    if rc != 0:
        print("\nLoi khi chay generate.py (ma %d)." % rc)
        return rc

    out = latest_output_dir()
    print("\nXONG. File nam trong:", out or os.path.join(HERE, "output"))
    try:
        if out and sys.platform.startswith("win"):
            os.startfile(out)  # mo thu muc ket qua
    except Exception as e:
        print("(khong mo duoc thu muc:", e, ")")
    return 0


if __name__ == "__main__":
    sys.exit(main())
