# -*- coding: utf-8 -*-
"""Tests for base_vn.py (Base client) and base_columns.py (column mapping).

Offline: HTTP transport is faked — no network calls.
"""
from __future__ import annotations

import io
import json
import os
import sys
import urllib.error
import urllib.parse

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import base_columns
import base_vn
from base_columns import DEFAULTS
from base_vn import BaseVnClient, BaseVnError, check_ok, parse_response


class _FakeResp:
    def __init__(self, raw: str):
        self._data = raw.encode("utf-8")

    def read(self):
        return self._data

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


def _client(workflow_token="tok-wf", wework_token="tok-ww", webhook="https://x/webhook/create/t"):
    c = BaseVnClient(workflow_token=workflow_token, wework_token=wework_token,
                     webhook_create_url=webhook)
    c.calls = []

    class Opener:
        def open(self, req, timeout):
            c.calls.append(((req.full_url, req.get_method(),
                            req.data.decode("utf-8") if req.data else None,
                            dict(req.header_items()))))
            return _FakeResp(c.next_raw)

    c._opener = Opener()
    c.next_raw = json.dumps({"code": 1, "data": {"ok": True}})
    return c


# ------------------------------------------------------------------ base_vn
class TestEnv:
    def test_parse_env_file(self, tmp_path):
        p = tmp_path / "x.env"
        p.write_text('A="hello world"\r\nB=2\n#C=3\nD="quoted"\n', encoding="utf-8")
        for k in ("A", "B", "D"):
            os.environ.pop(k, None)
        try:
            base_vn.load_env_file(str(p))
            assert os.environ["A"] == "hello world"
            assert os.environ["B"] == "2"
            assert os.environ["D"] == "quoted"
            assert "C" not in os.environ
        finally:
            for k in ("A", "B", "D"):
                os.environ.pop(k, None)

    def test_require_env_missing(self, monkeypatch):
        monkeypatch.delenv("BASE_WF_X", raising=False)
        with pytest.raises(BaseVnError, match="BASE_WF_X"):
            base_vn.require_env("BASE_WF_X")

    def test_require_env_default(self):
        assert base_vn.require_env("BASE_UNSET_Y", "fallback") == "fallback"


class TestClientTransport:
    def test_workflow_url_and_form(self):
        c = _client()
        c.workflow("job/create", workflow_id=9, name="A", custom_x="1")
        url, method, body, headers = c.calls[-1]
        assert url == "https://workflow.base.vn/extapi/v1/job/create"
        assert method == "POST"
        form = dict(x.split("=", 1) for x in body.split("&"))
        assert form["access_token_v2"] == "tok-wf"
        assert form["workflow_id"] == "9"
        assert form["custom_x"] == "1"
        assert headers and any("application/x-www-form-urlencoded" in v for _, v in headers.items())

    def test_wework_url(self):
        c = _client()
        c.wework("task/get", id=5)
        url, method, body, _ = c.calls[-1]
        assert url == "https://wework.base.vn/extapi/v3/task/get"
        assert "access_token_v2=tok-ww" in body

    def test_missing_token(self, monkeypatch):
        monkeypatch.setattr(base_vn, "load_env_local", lambda: None)
        monkeypatch.delenv("BASE_WORKFLOW_TOKEN", raising=False)
        c = BaseVnClient(workflow_token="", wework_token="", webhook_create_url="")
        with pytest.raises(BaseVnError, match="WORKFLOW_TOKEN"):
            c.workflow("workflow/get", id=1)

    def test_create_webhook_form_body(self):
        c = _client()
        c.next_raw = "<html>Workflow</html>"
        r = c.create_job_via_webhook({"name": "Đơn A", "custom_kh": "X"})
        url, method, body, headers = c.calls[-1]
        assert url == "https://x/webhook/create/t"
        assert urllib.parse.parse_qs(body) == {"name": ["Đơn A"], "custom_kh": ["X"]}
        assert any("application/x-www-form-urlencoded" in v for _, v in headers.items())
        assert r == {"ok": True, "html_response": True}

    def test_create_webhook_json_ok(self):
        c = _client()
        c.next_raw = json.dumps({"code": 1, "data": {"ok": True}})
        r = c.create_job_via_webhook({"name": "A"})
        assert r["data"]["ok"] is True

    def test_create_webhook_rejected(self):
        c = _client()
        c.next_raw = json.dumps({"code": 0, "message": "Invalid input (custom field): X"})
        with pytest.raises(BaseVnError, match="Invalid input"):
            c.create_job_via_webhook({"name": "A"})

    def test_create_webhook_missing_url(self, monkeypatch):
        monkeypatch.setattr(base_vn, "load_env_local", lambda: None)
        monkeypatch.delenv("BASE_WEBHOOK_CREATE_URL", raising=False)
        c = BaseVnClient(workflow_token="t", webhook_create_url="")
        with pytest.raises(BaseVnError, match="WEBHOOK"):
            c.create_job_via_webhook({"name": "x"})

    def test_api_error_code_raises(self):
        c = _client()
        c.next_raw = json.dumps({"code": 0, "msg": "bad token"})
        with pytest.raises(BaseVnError, match="bad token"):
            c.workflow("workflow/get", id=1)

    def test_http_error_raises(self, monkeypatch):
        c = _client()
        c.next_raw = ""

        def boom(req, timeout):
            raise urllib.error.HTTPError(req.full_url, 401, "Unauthorized", {}, io.BytesIO(b'Nope'))

        c._opener.open = boom
        with pytest.raises(BaseVnError, match="401"):
            c.workflow("workflow/get", id=1)

    def test_non_json_response(self):
        c = _client()
        c.next_raw = "not json"
        with pytest.raises(BaseVnError, match="Non-JSON"):
            c.workflow("workflow/get", id=1)

    def test_network_error(self, monkeypatch):
        c = _client()
        c.next_raw = ""

        def boom(req, timeout):
            raise urllib.error.URLError("Connection refused")

        c._opener.open = boom
        with pytest.raises(BaseVnError, match="network error"):
            c.workflow("workflow/get", id=1)


class TestResponse:
    def test_check_ok(self):
        assert check_ok({"code": 1})
        assert check_ok({"code": "success"})
        assert check_ok({"success": True})
        assert check_ok({"data": []})
        assert not check_ok({"code": 0})
        assert not check_ok({"success": False})

    def test_parse_response(self):
        assert parse_response('{"a": 1}') == {"a": 1}


# ------------------------------------------------------------- base_columns
def _colmap(**overrides):
    d = json.loads(json.dumps(DEFAULTS))
    d.update(overrides)
    return d


class TestColumnMap:
    def test_load_defaults_and_merge(self, tmp_path):
        p = tmp_path / "base_columns.json"
        p.write_text(json.dumps({"confirmed": True, "workflow_id": 42,
                                 "order_columns": {"customer": "custom_ten_kh"}}),
                     encoding="utf-8")
        m = base_columns.load_column_map(str(p))
        assert m["confirmed"] is True
        assert m["workflow_id"] == 42
        assert m["order_columns"]["customer"] == "custom_ten_kh"
        assert m["order_columns"]["qty"]  # default preserved

    def test_save_roundtrip(self, tmp_path):
        p = tmp_path / "bm.json"
        base_columns.save_column_map(_colmap(confirmed=True), str(p))
        m = base_columns.load_column_map(str(p))
        assert m["confirmed"] is True

    def test_default_confirmed_workflow_id(self):
        assert load_default()["confirmed"] is True
        assert load_default()["workflow_id"] == 1318
        assert load_default()["order_columns"]["order_id"] == "custom_so_phieu_ycsx"


def load_default():
    return base_columns.load_column_map(os.path.join(
        os.path.dirname(__file__), "..", "data", "base_columns.json"))


class TestBuildWebhookFields:
    def test_maps_order_columns_and_trigger(self):
        order = {"customer": "ACME", "order_id": "O1", "product_name": "Bao 40kg",
                 "qty": 3000, "spec": "BOPP"}
        m = _colmap()
        f = base_columns.build_webhook_fields(order, m)
        assert f["custom_khach_hang"] == "ACME"
        assert f["custom_ma_don_hang"] == "O1"
        assert f["custom_so_luong"] == 3000
        assert f["name"] == "O1 - Bao 40kg"
        assert f["custom_call_dinh_muc_nvl"] == "1"
        assert "custom_yeu_cau_ky_thuat" in f

    def test_skips_missing_fields(self):
        m = _colmap()
        f = base_columns.build_webhook_fields({"qty": 2}, m)
        assert "custom_khach_hang" not in f
        assert "custom_so_luong" in f

    def test_trigger_disabled_if_key_missing(self):
        m = _colmap()
        m["trigger"]["call_dinh_muc_nvl"] = ""
        f = base_columns.build_webhook_fields({"qty": 2}, m)
        assert "custom_call_dinh_muc_nvl" not in f

    def test_order_id_prefix_only_when_present(self):
        m = _colmap()
        f = base_columns.build_webhook_fields({"product_name": "Only name"}, m)
        assert f["name"] == "Only name"


class TestOrderFromWebhook:
    def test_flat_payload(self):
        m = _colmap()
        payload = {"custom_khach_hang": "ACME", "custom_so_luong": "3.000",
                   "custom_kho_mang": "1,04"}
        o = base_columns.order_from_webhook(payload, m)
        assert o["customer"] == "ACME"
        assert o["qty"] == 3
        assert abs(o["kho_mang"] - 1.04) < 1e-9

    def test_nested_data_payload(self):
        m = _colmap()
        payload = {"data": {"custom_ma_don_hang": "O9", "custom_hong_cm": "8"},
                   "user_id": 123}
        o = base_columns.order_from_webhook(payload, m)
        assert o["order_id"] == "O9"
        assert o["gusset_cm"] == 8.0
        assert "customer" not in o

    def test_bad_number_ignored(self):
        m = _colmap()
        o = base_columns.order_from_webhook({"custom_kho_mang": "abc"}, m)
        assert "kho_mang" not in o

    def test_fallback_plain_key(self):
        m = _colmap()
        o = base_columns.order_from_webhook({"customer": "FALLBACK"}, m)
        assert o["customer"] == "FALLBACK"

    def test_nested_objects_skipped(self):
        m = _colmap()
        o = base_columns.order_from_webhook({"assignee": {"id": 1}}, m)
        assert o == {}


class TestResultFields:
    def test_maps_result_to_custom_keys(self):
        m = _colmap()
        r = base_columns.result_fields({"mang_bopp_kg": 49.1407,
                                        "bao_kien": 300, "status": "done"}, m)
        assert r["custom_mang_bopp_kg"] == "49.1407"
        assert r["custom_so_bao_kien"] == "300"
        assert r["custom_trang_thai"] == "done"

    def test_ignores_none(self):
        m = _colmap()
        assert base_columns.result_fields({"mang_bopp_kg": None}, m) == {}