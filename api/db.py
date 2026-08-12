# -*- coding: utf-8 -*-
"""Thin Supabase (PostgREST) store backing the Định mức review board.

Two backends behind one small interface:

* **SupabaseStore**  — production. Talks directly to the PostgREST endpoint at
  ``SUPABASE_URL`` using ``SUPABASE_SERVICE_ROLE_KEY`` (trusted serverless fn).
  No ORM dependency — a ~60-line stdlib client keeps the Vercel bundle light.
* **LocalJsonStore** — dev fallback persisted to ``.board_db/orders.json`` so
  ``uvicorn`` / ``pytest`` work out of the box without env vars.

``get_store()`` returns whichever backend is configured (process-wide).
"""
from __future__ import annotations

import json
import os
import threading
import uuid
from datetime import datetime, timezone
from typing import Any, Optional
from urllib import error as urlerror, parse as urlparse, request as urlrequest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
LOCAL_DB_DIR = os.path.join(ROOT, ".board_db")
LOCAL_DB_PATH = os.path.join(LOCAL_DB_DIR, "orders.json")

TABLE = "dinh_muc_orders"

# Column projection shared by both backends (mirrors the Supabase table DDL in
# supabase/schema.sql). JSON-ish columns stay Python objects end-to-end.
_ORDER_COLUMNS = (
    "id", "order_id", "customer", "product_name", "product_code", "qty",
    "so_mau_in", "family", "stage", "order_json", "fields_json",
    "summary_json", "warnings", "reviewer", "accepted_at", "accepted_by",
    "reject_reason", "created_at", "updated_at",
)

# Columns mutable after import (everything else is set once at creation).
_UPDATEABLE = ("stage", "reviewer", "accepted_at", "accepted_by", "reject_reason")

DEFAULT_STAGE = "thong_tin"


def utcnow() -> str:
    """ISO-8601 UTC timestamp (Z suffix) for created_at/updated_at/accepted_at."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def _configured() -> bool:
    return bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_SERVICE_ROLE_KEY"))


class _BoardStore:
    def create_order(self, data: dict) -> dict:
        raise NotImplementedError

    def get_order(self, order_id: str) -> Optional[dict]:
        raise NotImplementedError

    def list_orders(self) -> list:
        raise NotImplementedError

    def update_order(self, order_id: str, **fields) -> Optional[dict]:
        raise NotImplementedError


class LocalJsonStore(_BoardStore):
    """File-backed store for local dev / tests (no SUPABASE_* env vars)."""

    def __init__(self, path: str = LOCAL_DB_PATH):
        self.path = path
        self._lock = threading.Lock()

    def _load(self) -> list:
        try:
            with open(self.path, encoding="utf-8") as f:
                rows = json.load(f)
            return rows if isinstance(rows, list) else []
        except (OSError, ValueError):
            return []

    def _save(self, rows: list) -> None:
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=1)

    @staticmethod
    def _public(row: dict) -> dict:
        return dict(row)

    def create_order(self, data: dict) -> dict:
        row = {c: data.get(c) for c in _ORDER_COLUMNS}
        row["id"] = data.get("id") or str(uuid.uuid4())
        row["stage"] = row.get("stage") or DEFAULT_STAGE
        row["created_at"] = row.get("created_at") or utcnow()
        row["updated_at"] = utcnow()
        with self._lock:
            rows = self._load()
            rows.insert(0, row)
            self._save(rows)
        return self._public(row)

    def get_order(self, order_id: str) -> Optional[dict]:
        with self._lock:
            rows = self._load()
        for row in rows:
            if row.get("id") == order_id:
                return self._public(row)
        return None

    def list_orders(self) -> list:
        with self._lock:
            rows = self._load()
        rows.sort(key=lambda r: r.get("created_at") or "", reverse=True)
        return [self._public(r) for r in rows]

    def update_order(self, order_id: str, **fields) -> Optional[dict]:
        patch = {k: v for k, v in fields.items() if k in _UPDATEABLE}
        if not patch:
            return self.get_order(order_id)
        patch["updated_at"] = utcnow()
        with self._lock:
            rows = self._load()
            for row in rows:
                if row.get("id") == order_id:
                    row.update(patch)
                    self._save(rows)
                    return self._public(row)
        return None


class SupabaseStore(_BoardStore):
    """Minimal PostgREST client for the ``dinh_muc_orders`` table (service role)."""

    def __init__(self, url: Optional[str] = None, key: Optional[str] = None):
        self.base = (url or os.environ["SUPABASE_URL"]).rstrip("/")
        self.key = key or os.environ["SUPABASE_SERVICE_ROLE_KEY"]
        self.endpoint = f"{self.base}/rest/v1/{TABLE}"

    def _headers(self, prefer: Optional[str] = None) -> dict:
        headers = {
            "apikey": self.key,
            "Authorization": f"Bearer {self.key}",
            "Content-Type": "application/json",
        }
        if prefer:
            headers["Prefer"] = prefer
        return headers

    def _request(self, method: str, url: str, body: Optional[dict] = None,
                 prefer: Optional[str] = None) -> Optional[Any]:
        payload = json.dumps(body).encode("utf-8") if body is not None else None
        req = urlrequest.Request(url, data=payload, headers=self._headers(prefer), method=method)
        try:
            with urlrequest.urlopen(req, timeout=30) as resp:
                raw = resp.read().decode("utf-8")
        except urlerror.HTTPError as e:
            detail = e.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Supabase {method} {self.endpoint} → HTTP {e.code}: {detail}")
        return json.loads(raw) if raw.strip() else None

    @staticmethod
    def _esc(value: str) -> str:
        return urlparse.quote(str(value), safe="")

    def create_order(self, data: dict) -> dict:
        body = {c: data[c] for c in _ORDER_COLUMNS if data.get(c) is not None}
        body["id"] = data.get("id") or str(uuid.uuid4())
        body["stage"] = body.get("stage") or DEFAULT_STAGE
        body["created_at"] = body.get("created_at") or utcnow()
        body["updated_at"] = utcnow()
        rows = self._request("POST", self.endpoint, body=body, prefer="return=representation")
        return (rows or [body])[0]

    def get_order(self, order_id: str) -> Optional[dict]:
        rows = self._request("GET", f"{self.endpoint}?select=*&id=eq.{self._esc(order_id)}")
        return (rows or [None])[0]

    def list_orders(self) -> list:
        rows = self._request("GET", f"{self.endpoint}?select=*&order=created_at.desc")
        return rows or []

    def update_order(self, order_id: str, **fields) -> Optional[dict]:
        patch = {k: v for k, v in fields.items() if k in _UPDATEABLE}
        if not patch:
            return self.get_order(order_id)
        patch["updated_at"] = utcnow()
        rows = self._request(
            "PATCH", f"{self.endpoint}?id=eq.{self._esc(order_id)}",
            body=patch, prefer="return=representation")
        return (rows or [None])[0]


_store: Optional[_BoardStore] = None


def get_store() -> _BoardStore:
    """Process-wide store: Supabase when env vars are set, local JSON otherwise."""
    global _store
    if _store is None:
        if _configured():
            _store = SupabaseStore()
        else:
            print("api.db: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY chưa đặt — "
                  f"dùng kho dữ liệu cục bộ (file {LOCAL_DB_PATH}) cho bảng đơn hàng.")
            _store = LocalJsonStore()
    return _store


def reset_store() -> None:
    """Forget the cached store (tests): next get_store() re-evaluates env."""
    global _store
    _store = None