# -*- coding: utf-8 -*-
"""Tests for gui_don_vn.py (PC sender). Offline: --dry-run only."""
from __future__ import annotations

import json
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import gui_don_vn

SAMPLES = os.path.join(os.path.dirname(__file__), "..", "samples")
SAMPLE = os.path.join(SAMPLES, "40kg_4oranges.json")


@pytest.fixture(autouse=True)
def no_explorer(monkeypatch):
    monkeypatch.setenv("GUI_DON_NO_OPEN", "1")


def test_dry_run_generates_files(tmp_path):
    rc = gui_don_vn.main(["--order", SAMPLE, "--colors", "3",
                          "--dry-run", "--outdir", str(tmp_path)])
    assert rc == 0
    zips = [f for f in os.listdir(tmp_path) if f.endswith(".zip")]
    assert len(zips) == 1


def test_missing_file_returns_1(tmp_path):
    rc = gui_don_vn.main(["--order", str(tmp_path / "nope.json"),
                          "--colors", "3", "--dry-run"])
    assert rc == 1


def test_zero_colors_allowed(tmp_path):
    rc = gui_don_vn.main(["--order", SAMPLE, "--colors", "0",
                          "--dry-run", "--outdir", str(tmp_path)])
    assert rc == 0


def test_bad_json_returns_1(tmp_path):
    bad = tmp_path / "bad.json"
    bad.write_text("{not json", encoding="utf-8")
    rc = gui_don_vn.main(["--order", str(bad), "--colors", "3", "--dry-run"])
    assert rc == 1