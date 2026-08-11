# -*- coding: utf-8 -*-
"""dinh_muc_service.py — shared wrapper around generate.run for the Base flow.

Used by the PC sender (gui_don_vn.py) and the server receiver
(api/index.py::/api/wf/receive). Produces the Excel files via generate.run and
a compact result summary that can be written back into workflow columns.
"""
from __future__ import annotations

import os
from typing import Any, Optional

import generate


def compute_summary(result: dict) -> dict:
    """Flatten generate.run() output into a compact write-back summary."""
    family = result.get("family")
    products = result.get("products") or []
    fields = (products[0] or {}).get("fields") if products else {}
    if not isinstance(fields, dict):
        fields = {}
    summary: dict[str, Any] = {
        "family": family,
        "outdir": result.get("outdir", ""),
        "outputs": [os.path.basename(p) for p in (result.get("outputs") or [])],
        "status": "done",
        "error": None,
    }
    for key in ("mang_bopp_kg", "dung_moai_opp_kg", "giay_kraft_kg",
                "bao_kien", "so_phieu_sx", "order_id"):
        if fields.get(key) is not None:
            summary[key] = fields.get(key)
    if products:
        summary["products"] = [p.get("product_name") for p in products]
    return summary


def run_and_summarize(order: dict, colors: int,
                      outdir: Optional[str] = None) -> dict:
    """Run generate.run on an order dict and return {"result", "summary"}."""
    result = generate.run(("dict", order), colors, outdir=outdir)
    summary = compute_summary(result)
    return {"result": result, "summary": summary}


def write_result_fields(summary: dict) -> dict:
    """Select the summary fields safe to push back as workflow columns."""
    interested = ("family", "status", "mang_bopp_kg", "dung_moai_opp_kg",
                  "giay_kraft_kg", "bao_kien")
    return {k: summary[k] for k in interested if summary.get(k) is not None}