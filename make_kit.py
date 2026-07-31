#!/usr/bin/env python3
"""make_kit.py — Bundle dinh_muc_kit.zip for claude.ai code-execution sandbox.

Usage:  python make_kit.py
Output: auto-dinh-muc/dinh_muc_kit.zip

The kit contains only what the sandbox needs: the engine, data, templates,
and samples. It can be uploaded to Project Knowledge or attached each session.
"""
from __future__ import annotations
import os, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "dinh_muc_kit.zip")

INCLUDE = [
    "generate.py",
    "xlsxpatch.py",
    "nvl_lookup.py",
    "tsvh_lookup.py",
]

INCLUDE_DIRS = {
    "data": [".json", ".xlsx"],
    "templates": [".xlsx"],
    "samples": [".json"],
}

EXCLUDE_SUFFIXES = (".bak",)
EXCLUDE_DIRS = {"output", "__pycache__", ".git", "__MACOSX", "tests"}
EXCLUDE_FILES = {"colour_25kg.py", "tao_don.py", "don_hang.bat", "make_kit.py",
                 "dinh_muc_kit.zip", "README.md", "README_setup.md", "PROJECT_INSTRUCTIONS.md"}


def main():
    if os.path.exists(OUT):
        os.remove(OUT)
    count = 0
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in INCLUDE:
            src = os.path.join(HERE, name)
            if os.path.exists(src):
                zf.write(src, name)
                count += 1
        for dirname, extensions in INCLUDE_DIRS.items():
            src_dir = os.path.join(HERE, dirname)
            if not os.path.isdir(src_dir):
                continue
            for root, dirs, files in os.walk(src_dir):
                rel_root = os.path.relpath(root, HERE)
                dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
                for f in files:
                    if f.endswith(EXCLUDE_SUFFIXES):
                        continue
                    if any(f.endswith(ext) for ext in extensions):
                        src = os.path.join(root, f)
                        arcname = os.path.join(rel_root, f)
                        zf.write(src, arcname)
                        count += 1
    print(f"Created {OUT} ({count} files)")


if __name__ == "__main__":
    main()
