# -*- coding: utf-8 -*-
"""base_columns.py — mapping between order JSON fields and Base Workflow columns.

Loaded from/saved to ``data/base_columns.json``. Keys are the logical order
fields on one side and the workflow's real column keys (``custom_<input_key>``)
on the other. ``confirmed`` flips to true once the discover step has validated
the real keys against the "Quản lý đơn hàng sản xuất" workflow.
"""
from __future__ import annotations

import json
import os
from typing import Any, Optional

HERE = os.path.dirname(os.path.abspath(__file__))
COLUMNS_PATH = os.path.join(HERE, "data", "base_columns.json")

# Logical order fields that map onto workflow columns.
ORDER_FIELDS = [
    "customer", "product_name", "product_code", "order_id", "so_phieu_sx",
    "qty", "bag_length_m", "width_plus_gusset_m", "width_cm", "gusset_cm",
    "inner_bag_weight_kg", "kho_mang", "spec",
]

DEFAULTS: dict[str, Any] = {
    "_comment": "Column mapping order-json <-> Base Workflow 'Quản lý đơn hàng sản xuất'. "
                "Unconfirmed best-guess keys; run the discover step to validate.",
    "confirmed": False,
    "workflow_id": None,
    "trigger": {
        "call_dinh_muc_nvl": "custom_call_dinh_muc_nvl",
        "trigger_value": "1",
    },
    "order_columns": {
        "customer": "custom_khach_hang",
        "product_name": "custom_ten_san_pham",
        "product_code": "custom_ma_san_pham",
        "order_id": "custom_ma_don_hang",
        "so_phieu_sx": "custom_so_phieu_sx",
        "qty": "custom_so_luong",
        "bag_length_m": "custom_chieu_dai_m",
        "width_plus_gusset_m": "custom_chieu_rong_cong_hong_m",
        "width_cm": "custom_chieu_rong_cm",
        "gusset_cm": "custom_hong_cm",
        "inner_bag_weight_kg": "custom_trong_luong_tui_pe_kg",
        "kho_mang": "custom_kho_mang",
        "spec": "custom_yeu_cau_ky_thuat",
    },
    "result_columns": {
        "family": "custom_loai_bao",
        "mang_bopp_kg": "custom_mang_bopp_kg",
        "dung_moai_opp_kg": "custom_dung_moai_opp_kg",
        "giay_kraft_kg": "custom_giay_kraft_kg",
        "bao_kien": "custom_so_bao_kien",
        "status": "custom_trang_thai",
        "error": "custom_loi",
    },
}

NUMERIC_FIELDS = [
    "qty", "bag_length_m", "width_plus_gusset_m", "width_cm", "gusset_cm",
    "inner_bag_weight_kg", "kho_mang",
]


def _deep_defaults() -> dict[str, Any]:
    return json.loads(json.dumps(DEFAULTS))


def load_column_map(path: Optional[str] = None) -> dict[str, Any]:
    path = path or COLUMNS_PATH
    colmap = _deep_defaults()
    try:
        with open(path, encoding="utf-8") as f:
            stored = json.load(f)
    except (OSError, ValueError):
        stored = {}
    for section in ("trigger", "order_columns", "result_columns"):
        if isinstance(stored.get(section), dict):
            colmap[section].update({k: v for k, v in stored[section].items() if v is not None})
    for key in ("confirmed", "workflow_id"):
        if key in stored:
            colmap[key] = stored[key]
    return colmap


def save_column_map(colmap: dict[str, Any], path: Optional[str] = None) -> str:
    path = path or COLUMNS_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(colmap, f, ensure_ascii=False, indent=2)
    return path


def _custom_key(colmap: dict[str, Any], section: str, field: str) -> Optional[str]:
    value = colmap.get(section, {}).get(field)
    return str(value) if value else None


def build_webhook_fields(order: dict, colmap: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    """Map an order dict -> workflow field keys for the create webhook.

    Sets the order columns, a readable name/content, and the trigger column
    ("call định mức nvl") so the job is created already marked to compute.
    """
    colmap = colmap or load_column_map()
    fields: dict[str, Any] = {}
    for field, key in (colmap.get("order_columns") or {}).items():
        if key and order.get(field) is not None and order.get(field) != "":
            fields[key] = order[field]
    name = order.get("product_name") or order.get("order_id") or "Đơn hàng"
    if order.get("order_id"):
        name = f"{order['order_id']} - {name}"
    fields.setdefault("name", name)
    fields.setdefault("content", json.dumps(order, ensure_ascii=False)[:2000])
    trigger_key = _custom_key(colmap, "trigger", "call_dinh_muc_nvl")
    if trigger_key:
        fields[trigger_key] = str(colmap.get("trigger", {}).get("trigger_value", "1"))
    return fields


def order_from_webhook(payload: dict, colmap: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    """Extract an order dict from a Base webhook output payload.

    Tolerant to flat payloads or payloads nesting the job under ``data``; reads
    each order column by its custom key, then by a case-insensitive key/title.
    """
    colmap = colmap or load_column_map()
    src: dict[str, Any] = {}
    if isinstance(payload.get("data"), dict):
        src.update(payload["data"])
    src.update(payload)

    lookup: dict[str, Any] = {}
    for k, v in src.items():
        if isinstance(v, dict):  # nested objects (assignee/creator) are ignored
            continue
        lookup[str(k).lower()] = v
        lookup[str(k).strip()] = v

    order: dict[str, Any] = {}
    for field in ORDER_FIELDS:
        key = _custom_key(colmap, "order_columns", field)
        value = None
        if key:
            value = src.get(key, lookup.get(key.lower()))
        if value is None and key is None:
            value = lookup.get(field.lower(), lookup.get(field))
        if value is None:
            value = lookup.get(field.replace("_", " ").lower())
        if value is None:
            continue
        if field in NUMERIC_FIELDS:
            try:
                value = float(str(value).replace(",", ".").strip())
                if field == "qty":
                    value = int(value)
            except (TypeError, ValueError):
                continue
        order[field] = value
    return order


def result_fields(result: dict, colmap: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    """Map a computed result summary -> workflow custom_* keys for job/edit."""
    colmap = colmap or load_column_map()
    fields: dict[str, Any] = {}
    for field, value in result.items():
        key = _custom_key(colmap, "result_columns", field)
        if key and value is not None:
            fields[key] = str(value)
    return fields
