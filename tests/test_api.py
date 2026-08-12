# -*- coding: utf-8 -*-
"""Tests for the Base.vn webhook endpoints in api/index.py.

No network / no TestClient: endpoints are invoked directly with a fake Request
and a stubbed BaseVnClient.
"""
from __future__ import annotations

import asyncio
import json
import os
import sys

import pytest
from fastapi import HTTPException

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import api.index as api

HERE = os.path.dirname(os.path.abspath(__file__))
YCSX_SAMPLE = os.path.join(HERE, "..", "samples", "YCSX_4oranges.xlsx")


class FakeReq:
    def __init__(self, body: bytes):
        self._body = body

    async def body(self):
        return self._body


def _run(coro):
    return asyncio.run(coro)


class _FakeBase:
    updated = []
    moved = []
    fetched = []
    discover_data = {"workflows": [{"id": 1, "name": "Quản lý đơn hàng sản xuất"}]}

    def __init__(self):
        pass

    def update_job(self, job_id, **custom):
        _FakeBase.updated.append((job_id, custom))
        return {"code": 1}

    def move_next(self, job_id):
        _FakeBase.moved.append(job_id)
        return {"code": 1}

    def get_job(self, job_id):
        _FakeBase.fetched.append(job_id)
        return {"job": {"id": job_id,
                        "files": [{"name": "YCSX.xlsx",
                                   "url": "https://cdn.example/ycsx.xlsx"}]}}

    def download_file(self, url, dest, timeout=60):
        if not os.path.isfile(YCSX_SAMPLE):
            raise api.base_vn.BaseVnError("sample YCSX missing")
        with open(YCSX_SAMPLE, "rb") as f:
            data = f.read()
        with open(dest, "wb") as f:
            f.write(data)
        return dest

    def discover(self):
        return self.discover_data


@pytest.fixture(autouse=True)
def stub_base(monkeypatch):
    _FakeBase.updated.clear()
    _FakeBase.moved.clear()
    _FakeBase.fetched.clear()
    monkeypatch.setattr(api.base_vn, "BaseVnClient", _FakeBase)


def _wf_payload(order_id="O-X", qty="1000"):
    return {
        "id": 99,
        "custom_khach_hang": "ACME",
        "custom_ma_don_hang": order_id,
        "custom_ma_san_pham": "XD1TP00001",
        "custom_so_luong": qty,
        "custom_chieu_dai_m": "0.82",
        "custom_chieu_rong_cong_hong_m": "0.5",
        "custom_chieu_rong_cm": "42",
        "custom_hong_cm": "8",
        "custom_trong_luong_tui_pe_kg": "0.051",
        "custom_kho_mang": "1.04",
        "custom_yeu_cau_ky_thuat": "1.Kích thước: (42+8) cm x 82cm\n3. Cấu trúc: Bao BOPP in ống đồng - OPP mờ",
        "colors": "3",
    }


class TestWfReceive:
    def test_generates_and_writes_back(self):
        resp = _run(api.wf_receive(FakeReq(json.dumps(_wf_payload()).encode())))
        assert resp["ok"] is True
        assert resp["job_id"] == 99
        assert resp["colors"] == 3
        assert resp["summary"]["status"] == "done"
        assert resp["written_back"] is True
        assert len(_FakeBase.updated) == 1
        job_id, custom = _FakeBase.updated[0]
        assert job_id == 99
        assert any(k.startswith("custom_") for k in custom)

    def test_write_back_failure_flagged(self, monkeypatch):
        class Broken(_FakeBase):
            def update_job(self, job_id, **custom):
                raise api.base_vn.BaseVnError("boom")

        monkeypatch.setattr(_FakeBase, "update_job", Broken.update_job)
        resp = _run(api.wf_receive(FakeReq(json.dumps(_wf_payload()).encode())))
        assert resp["ok"] is True
        assert resp["written_back"] is False
        assert "boom" in (resp["summary"].get("error") or "")

    def test_empty_order_400(self):
        with pytest.raises(HTTPException) as ei:
            _run(api.wf_receive(FakeReq(b'{"custom_khach_hang":"ACME"}')))
        assert ei.value.status_code == 400

    def test_invalid_json_400(self):
        with pytest.raises(HTTPException) as ei:
            _run(api.wf_receive(FakeReq(b"not-json")))
        assert ei.value.status_code == 400

    def test_nested_data_payload(self):
        payload = {"data": _wf_payload(order_id="O-NEST", qty="2000")}
        resp = _run(api.wf_receive(FakeReq(json.dumps(payload).encode())))
        assert resp["ok"] is True
        assert resp["job_id"] == 99
        assert resp["summary"]["status"] == "done"

    def test_no_job_id_skips_write_back(self):
        payload = _wf_payload()
        payload.pop("id")
        resp = _run(api.wf_receive(FakeReq(json.dumps(payload).encode())))
        assert resp["ok"] is True
        assert resp["written_back"] is False
        assert _FakeBase.updated == []

    def test_ycsx_url_payload(self):
        if not os.path.isfile(YCSX_SAMPLE):
            pytest.skip("YCSX_4oranges.xlsx not found")
        payload = {"id": 77, "ycsx_url": "https://cdn.example/YCSX.xlsx"}
        resp = _run(api.wf_receive(FakeReq(json.dumps(payload).encode())))
        assert resp["ok"] is True
        assert resp["ycsx_used"] is True
        assert resp["summary"]["status"] == "done"
        assert resp["written_back"] is True
        assert _FakeBase.updated[0][0] == 77

    def test_ycsx_from_fetched_job(self):
        if not os.path.isfile(YCSX_SAMPLE):
            pytest.skip("YCSX_4oranges.xlsx not found")
        payload = {"id": 88}
        resp = _run(api.wf_receive(FakeReq(json.dumps(payload).encode())))
        assert resp["ok"] is True
        assert resp["ycsx_used"] is True
        assert _FakeBase.fetched == [88]
        assert _FakeBase.updated[0][0] == 88

    def test_auto_move_next_when_enabled(self, monkeypatch):
        monkeypatch.setenv("BASE_AUTO_MOVE_NEXT", "1")
        resp = _run(api.wf_receive(FakeReq(json.dumps(_wf_payload()).encode())))
        assert resp["moved_next"] is True
        assert _FakeBase.moved == [99]

    def test_no_move_next_by_default(self):
        resp = _run(api.wf_receive(FakeReq(json.dumps(_wf_payload()).encode())))
        assert resp["moved_next"] is False
        assert _FakeBase.moved == []


class TestWfDiscover:
    def test_returns_probe_data(self):
        resp = _run(api.wf_discover())
        assert resp["workflows"][0]["name"] == "Quản lý đơn hàng sản xuất"

    def test_error_becomes_400(self, monkeypatch):
        class Boom:
            def discover(self):
                raise api.base_vn.BaseVnError("no token")

        monkeypatch.setattr(api.base_vn, "BaseVnClient", Boom)
        with pytest.raises(HTTPException) as ei:
            _run(api.wf_discover())
        assert ei.value.status_code == 400

class TestColorsCap:
    """Máy in tối đa 6 màu — chặn ở cả form web và webhook Base.vn."""

    def test_generate_endpoint_rejects_more_than_six(self):
        class FakeUpload:
            filename = "order.json"

            async def read(self):
                return b"{}"

        with pytest.raises(HTTPException) as exc:
            _run(api.generate_endpoint(file=FakeUpload(), colors=7))
        assert exc.value.status_code == 400
        assert "tối đa là 6" in exc.value.detail

    def test_generate_endpoint_still_rejects_negative(self):
        class FakeUpload:
            filename = "order.json"

            async def read(self):
                return b"{}"

        with pytest.raises(HTTPException) as exc:
            _run(api.generate_endpoint(file=FakeUpload(), colors=-1))
        assert exc.value.status_code == 400

    def test_payload_colors_clamped(self):
        assert api._colors_from_payload({"colors": 9}) == 6
        assert api._colors_from_payload({"colors": -3}) == 0
        assert api._colors_from_payload({"colors": 4}) == 4

    def test_payload_default_clamped(self, monkeypatch):
        monkeypatch.setenv("BASE_DEFAULT_COLORS", "10")
        assert api._colors_from_payload({}) == 6
