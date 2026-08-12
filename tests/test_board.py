# -*- coding: utf-8 -*-
"""Tests for the Định mức review board (import → board → review → accept/reject).

No network: the store is stubbed with an in-memory FakeStore, and the endpoint
functions are invoked directly with a fake Request (mirrors tests/test_api.py).
"""
from __future__ import annotations

import io
import json
import os
import sys
import uuid

import pytest
from fastapi import HTTPException
from fastapi.responses import FileResponse, RedirectResponse
from starlette.datastructures import UploadFile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import api.index as api

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLE = os.path.join(HERE, "..", "samples", "40kg_4oranges.json")


class FakeReq:
    def __init__(self, cookies=None, form=None):
        self.cookies = cookies or {}
        self._form = form or {}

    async def form(self):
        return self._form

    async def body(self):
        return b"{}"


class FakeStore:
    """In-memory stand-in for db.get_store()."""

    def __init__(self):
        self.rows: dict = {}

    def create_order(self, data):
        row = dict(data)
        row.setdefault("id", str(uuid.uuid4()))
        row.setdefault("stage", "thong_tin")
        row.setdefault("created_at", "2026-01-01T00:00:00Z")
        row["updated_at"] = row.get("created_at")
        self.rows[row["id"]] = row
        return row

    def get_order(self, order_id):
        return self.rows.get(order_id)

    def list_orders(self):
        return sorted(self.rows.values(), key=lambda r: r.get("created_at", ""), reverse=True)

    def update_order(self, order_id, **fields):
        row = self.rows.get(order_id)
        if row is None:
            return None
        row.update(fields)
        row["updated_at"] = "2026-01-01T00:00:01Z"
        return row


@pytest.fixture(autouse=True)
def stub_store(monkeypatch):
    store = FakeStore()
    monkeypatch.setattr(api.db, "get_store", lambda: store)
    return store


def _upload(filename="order.json"):
    with open(SAMPLE, "rb") as f:
        return UploadFile(file=io.BytesIO(f.read()), filename=filename,
                          headers={"content-type": "application/json"})


async def _import_order(filename="order.json", colors=3, reviewer="Kế toán A"):
    return await api.orders_create(file=_upload(filename), colors=colors, reviewer=reviewer)


class TestBoardPage:
    def test_board_shows_stages_and_empty_columns(self):
        html = _run(api.board_page())
        assert "Bảng đơn hàng" in html
        for title in ("Thông tin đơn hàng", "Kế toán kiểm tra", "QC kiểm tra",
                      "Lập định mức NVL", "Chuẩn bị NVL"):
            assert title in html

    def test_board_lists_cards_with_order_data(self, stub_store):
        row = stub_store.create_order({
            "order_id": "O-1", "customer": "ACME", "product_name": "Bao 40kg",
            "qty": 3000, "so_mau_in": 3, "family": "opp", "stage": "thong_tin",
            "order_json": {}, "fields_json": [], "summary_json": {}, "warnings": [],
        })
        html = _run(api.board_page())
        assert "/orders/" + row["id"] in html
        assert "Bao 40kg" in html
        assert "ACME" in html


class TestImport:
    def test_import_json_creates_order_and_redirects(self, stub_store):
        resp = _run(_import_order())
        assert isinstance(resp, RedirectResponse)
        assert resp.status_code == 303
        assert resp.headers["location"] == "/board"
        rows = stub_store.list_orders()
        assert len(rows) == 1
        row = rows[0]
        assert row["order_id"] == "26BA2HKH00565000014"
        assert row["so_mau_in"] == 3
        assert row["stage"] == "thong_tin"
        assert row["family"] == "opp"
        assert row["reviewer"] == "Kế toán A"
        assert row["fields_json"], "BOM must be persisted"
        assert isinstance(row["order_json"], dict), "order_json round-trip source"

    def test_import_sets_reviewer_cookie(self, stub_store):
        resp = _run(_import_order(reviewer="Thu"))
        assert "dinh_muc_reviewer=Thu" in resp.headers.get("set-cookie", "")

    def test_import_rejects_negative_colors(self):
        with pytest.raises(HTTPException) as ei:
            _run(_import_order(colors=-1))
        assert ei.value.status_code == 400

    def test_import_rejects_bad_extension(self, stub_store):
        with pytest.raises(HTTPException) as ei:
            _run(_import_order(filename="order.txt"))
        assert ei.value.status_code == 400

    def test_import_ycsx_xlsx(self, stub_store):
        ycsx = os.path.join(HERE, "..", "samples", "YCSX_4oranges.xlsx")
        if not os.path.isfile(ycsx):
            pytest.skip("YCSX_4oranges.xlsx not found")
        with open(ycsx, "rb") as f:
            file = UploadFile(file=io.BytesIO(f.read()), filename="ycsx.xlsx")
        resp = _run(api.orders_create(file=file, colors=3, reviewer="Kế toán A"))
        assert isinstance(resp, RedirectResponse)
        row = stub_store.list_orders()[0]
        assert row["order_id"] == "26BA2HKH00565000046"
        assert len(row["fields_json"]) == 3, "multi-product YCSX -> 3 BOM rows"


class TestReview:
    def test_detail_shows_order_and_bom(self, stub_store):
        row = stub_store.create_order(_make_row())
        html = api.pages.review_page(row)
        assert "Thông tin đơn hàng" in html
        assert "Màng BOPP (kg)" in html
        assert "Duyệt" in html
        assert "/orders/" + row["id"] + "/download" in html

    def test_detail_404_for_unknown_order(self):
        with pytest.raises(HTTPException) as ei:
            _run(api.order_detail("nope"))
        assert ei.value.status_code == 404


class TestAcceptReject:
    def test_accept_locks_into_dinh_muc(self, stub_store):
        row = stub_store.create_order(_make_row(stage="ke_toan"))
        resp = _run(api.order_accept(FakeReq(form={"reviewer": "Kế toán B"}), row["id"]))
        assert resp.headers["location"].endswith("/orders/" + row["id"])
        updated = stub_store.get_order(row["id"])
        assert updated["stage"] == "dinh_muc"
        assert updated["accepted_by"] == "Kế toán B"
        assert updated["accepted_at"]

    def test_accept_defaults_reviewer_to_cookie(self, stub_store):
        row = stub_store.create_order(_make_row())
        _run(api.order_accept(FakeReq(cookies={"dinh_muc_reviewer": "Cookie Người"}), row["id"]))
        assert stub_store.get_order(row["id"])["accepted_by"] == "Cookie Người"

    def test_accept_defaults_reviewer_to_fallback(self, stub_store):
        row = stub_store.create_order(_make_row())
        _run(api.order_accept(FakeReq(), row["id"]))
        assert stub_store.get_order(row["id"])["accepted_by"] == "Người duyệt"

    def test_reject_returns_to_thong_tin_with_reason(self, stub_store):
        row = stub_store.create_order(_make_row(stage="dinh_muc"))
        _run(api.order_reject(
            FakeReq(form={"reason": "Sai định lượng màng", "reviewer": "Kế toán C"}), row["id"]))
        updated = stub_store.get_order(row["id"])
        assert updated["stage"] == "thong_tin"
        assert updated["reject_reason"] == "Sai định lượng màng"
        assert not updated["accepted_by"]

    def test_stage_moves_through_columns(self, stub_store):
        row = stub_store.create_order(_make_row(stage="thong_tin"))
        _run(api.order_stage(FakeReq(form={"to": "ke_toan"}), row["id"]))
        assert stub_store.get_order(row["id"])["stage"] == "ke_toan"
        _run(api.order_stage(FakeReq(form={"to": "chuan_bi"}), row["id"]))
        assert stub_store.get_order(row["id"])["stage"] == "chuan_bi"

    def test_stage_rejects_unknown_target(self, stub_store):
        row = stub_store.create_order(_make_row())
        with pytest.raises(HTTPException) as ei:
            _run(api.order_stage(FakeReq(form={"to": "nowhere"}), row["id"]))
        assert ei.value.status_code == 400

    def test_action_on_unknown_order_404(self):
        with pytest.raises(HTTPException) as ei:
            _run(api.order_accept(FakeReq(form={"reviewer": "X"}), "missing"))
        assert ei.value.status_code == 404


class TestDownload:
    def test_download_regen_zip_from_stored_order(self, stub_store):
        _run(_import_order())
        row = stub_store.list_orders()[0]
        resp = _run(api.order_download(row["id"]))
        assert isinstance(resp, FileResponse)
        assert resp.media_type == "application/zip"
        assert resp.filename and resp.filename.endswith(".zip")
        assert os.path.isfile(resp.path), "ZIP regenerated on disk"

    def test_download_missing_order_404(self):
        with pytest.raises(HTTPException) as ei:
            _run(api.order_download("missing"))
        assert ei.value.status_code == 404


class TestLocalStore:
    """The dev fallback store itself round-trips against a real temp file."""

    def test_round_trip(self, tmp_path):
        from api.db import LocalJsonStore
        store = LocalJsonStore(path=str(tmp_path / "orders.json"))
        created = store.create_order({
            "order_id": "O-9", "customer": "ACME", "product_name": "Bao 25kg",
            "qty": 5000, "so_mau_in": 2, "family": "paper_kp", "stage": "thong_tin",
            "order_json": {"order_id": "O-9"}, "fields_json": [{"product_name": "Bao 25kg"}],
            "summary_json": {}, "warnings": [], "reviewer": None,
        })
        assert created["id"]
        got = store.get_order(created["id"])
        assert got["customer"] == "ACME"
        assert got["order_json"]["order_id"] == "O-9"

        updated = store.update_order(created["id"], stage="dinh_muc", accepted_by="Thu")
        assert updated["stage"] == "dinh_muc"
        assert updated["accepted_by"] == "Thu"

        # reload from disk (new instance) — persisted, survives restart
        reloaded = LocalJsonStore(path=str(tmp_path / "orders.json"))
        assert reloaded.get_order(created["id"])["stage"] == "dinh_muc"
        assert reloaded.list_orders()[0]["product_name"] == "Bao 25kg"

    def test_update_unknown_id_returns_none(self, tmp_path):
        from api.db import LocalJsonStore
        store = LocalJsonStore(path=str(tmp_path / "orders.json"))
        assert store.update_order("missing", stage="qc") is None
        assert store.get_order("missing") is None


def _run(coro):
    import asyncio
    return asyncio.run(coro)


def _make_row(stage="thong_tin", **overrides):
    row = {
        "id": str(uuid.uuid4()),
        "order_id": "O-T", "customer": "ACME", "product_name": "Bao 40kg",
        "product_code": "P1", "qty": 3000, "so_mau_in": 3, "family": "opp",
        "stage": stage, "order_json": {"order_id": "O-T"},
        "fields_json": [{"product_name": "Bao 40kg", "fields": {
            "mang_bopp_kg": 49.14, "sl_in_thuc_te_m": 2983,
            "kho_mang": 1.04, "so_mau_in": 3, "bao_kien_text": "300 bao/kiện",
        }}],
        "summary_json": {}, "warnings": ["Không tìm thấy Mành trắng trong NVL"],
        "reviewer": None, "accepted_at": None, "accepted_by": None, "reject_reason": None,
        "created_at": "2026-01-01T00:00:00Z", "updated_at": "2026-01-01T00:00:00Z",
    }
    row.update(overrides)
    return row