# -*- coding: utf-8 -*-
"""base_vn.py — Base.vn Workflow + Wework Public API client and webhook-create sender.

Bases (form-urlencoded, ``access_token`` in body):
  * Workflow: POST https://workflow.{{domain}}/extapi/v1/<endpoint>
  * Wework:   POST https://wework.{{domain}}/extapi/v3/<endpoint>

The workflow *create webhook* is a separate URL
(https://workflow.{{domain}}/webhook/create/<token>) that accepts a JSON body of
field keys (``name``, ``content``, ``custom_<input_key>``, ...) to create a job.

Env vars (loaded from ``.env.local`` if present and not already set):
  * BASE_WORKFLOW_TOKEN      — workflow access token (workspace_id~token)
  * BASE_WEWORK_TOKEN        — wework access token (workspace_id~token)
  * BASE_WEBHOOK_CREATE_URL  — full create-webhook URL
  * BASE_DOMAIN              — default base.vn
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Mapping, Optional

BASE_DOMAIN = "base.vn"
TIMEOUT = 30

# Logical order fields that map onto workflow columns (filled by discover).
ORDER_FIELDS = [
    "customer", "product_name", "product_code", "order_id", "so_phieu_sx",
    "qty", "bag_length_m", "width_plus_gusset_m", "width_cm", "gusset_cm",
    "inner_bag_weight_kg", "kho_mang", "spec",
]


class BaseVnError(Exception):
    """Raised when a Base.vn call fails (network, auth, or API-level error)."""


def load_env_file(path: str) -> None:
    """Parse a dotenv-style file (KEY="value") into os.environ if unset."""
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val
    except OSError:
        pass


def load_env_local() -> None:
    """Load the project's .env.local (next to this module) into os.environ."""
    here = os.path.dirname(os.path.abspath(__file__))
    for name in (".env.local", ".env"):
        load_env_file(os.path.join(here, name))


def require_env(name: str, default: Optional[str] = None) -> str:
    val = os.environ.get(name) or default
    if not val:
        raise BaseVnError(
            f"Missing env var {name} — add it to .env.local "
            "(see README_setup.md)."
        )
    return val


def parse_response(raw: str) -> dict:
    try:
        return json.loads(raw)
    except (ValueError, TypeError):
        raise BaseVnError(f"Non-JSON response from Base.vn: {raw[:300]}")


def check_ok(parsed: dict) -> bool:
    if not isinstance(parsed, dict):
        return False
    if "code" in parsed:
        code = parsed["code"]
        return code in (1, "1", "success", "SUCCESS", True, "true")
    if "success" in parsed:
        return parsed["success"] in (True, "true", 1, "1")
    return True


class BaseVnClient:
    """Thin client over the Workflow/Wework public APIs + create webhook."""

    def __init__(
        self,
        workflow_token: Optional[str] = None,
        wework_token: Optional[str] = None,
        webhook_create_url: Optional[str] = None,
        domain: str = BASE_DOMAIN,
        timeout: int = TIMEOUT,
    ) -> None:
        load_env_local()
        self.workflow_token = workflow_token or os.environ.get("BASE_WORKFLOW_TOKEN") or ""
        self.wework_token = wework_token or os.environ.get("BASE_WEWORK_TOKEN") or ""
        self.webhook_create_url = webhook_create_url or os.environ.get("BASE_WEBHOOK_CREATE_URL") or ""
        self.domain = domain
        self.timeout = timeout
        self._opener = urllib.request.build_opener()

    # ------------------------------------------------------------- transport
    def _post(self, url: str, *, form: Optional[dict] = None, json_body: Optional[dict] = None) -> dict:
        headers = {"User-Agent": "auto-dinh-muc/1.0"}
        if json_body is not None:
            data = json.dumps(json_body, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
        else:
            data = urllib.parse.urlencode(form or {}).encode("utf-8")
            headers["Content-Type"] = "application/x-www-form-urlencoded"
        req = urllib.request.Request(url, data=data, headers=headers, method="POST")
        try:
            with self._opener.open(req, timeout=self.timeout) as resp:
                raw = resp.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")[:500]
            raise BaseVnError(f"Base.vn HTTP {e.code} on {url}: {body}")
        except urllib.error.URLError as e:
            raise BaseVnError(f"Base.vn network error on {url}: {e.reason}")
        parsed = parse_response(raw)
        if not check_ok(parsed):
            msg = parsed.get("msg") or parsed.get("message") or parsed.get("error") or raw[:300]
            raise BaseVnError(f"Base.vn API error on {url}: {msg}")
        return parsed

    def _require(self, token: str, label: str) -> str:
        if not token:
            raise BaseVnError(f"Missing {label} — add BASE_{label.upper()} to .env.local")
        return token

    # ------------------------------------------------------------- Workflow
    def workflow(self, endpoint: str, **params: Any) -> dict:
        """POST https://workflow.{domain}/extapi/v1/{endpoint} (form-urlencoded)."""
        token = self._require(self.workflow_token, "WORKFLOW_TOKEN")
        url = f"https://workflow.{self.domain}/extapi/v1/{endpoint}"
        return self._post(url, form={"access_token_v2": token, **params})

    def create_job_via_webhook(self, fields: Mapping[str, Any]) -> dict:
        """Create a job in the workflow via its /webhook/create/<token> URL.

        ``fields`` are the workflow field keys (name, content, custom_*, ...).
        Form-encoded POST; Base answers with JSON on errors and the workflow
        HTML page on success — the HTML response is treated as job-created.
        """
        if not self.webhook_create_url:
            raise BaseVnError("Missing BASE_WEBHOOK_CREATE_URL in .env.local")
        data = urllib.parse.urlencode(dict(fields)).encode("utf-8")
        headers = {"User-Agent": "auto-dinh-muc/1.0",
                   "Content-Type": "application/x-www-form-urlencoded"}
        req = urllib.request.Request(self.webhook_create_url, data=data,
                                     headers=headers, method="POST")
        try:
            with self._opener.open(req, timeout=self.timeout) as resp:
                raw = resp.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            raise BaseVnError(f"Webhook create HTTP {e.code}: {e.read()[:300]}")
        except urllib.error.URLError as e:
            raise BaseVnError(f"Webhook create network error: {e.reason}")
        if raw.lstrip().startswith("{"):
            parsed = parse_response(raw)
            if not check_ok(parsed):
                msg = parsed.get("msg") or parsed.get("message") or raw[:300]
                raise BaseVnError(f"Webhook create rejected: {msg}")
            return parsed
        return {"ok": True, "html_response": True}

    # ---------------------------------------------------------------- Wework
    def wework(self, endpoint: str, **params: Any) -> dict:
        """POST https://wework.{domain}/extapi/v3/{endpoint} (form-urlencoded)."""
        token = self._require(self.wework_token, "WEWORK_TOKEN")
        url = f"https://wework.{self.domain}/extapi/v3/{endpoint}"
        return self._post(url, form={"access_token_v2": token, **params})

    # -------------------------------------------------------------- shortcuts
    def get_job(self, job_id) -> dict:
        return self.workflow("job/get", id=job_id)

    def get_job_files(self, job_id) -> list:
        """Return the job's attached files (dicts with name + url)."""
        resp = self.get_job(job_id)
        job = resp.get("job") or resp.get("data") or resp
        files = job.get("files") or []
        return [f for f in files if isinstance(f, dict) and f.get("url")]

    def download_file(self, url: str, dest: str, timeout: int = 60) -> str:
        """Download a Base CDN file URL to disk; returns the local path."""
        req = urllib.request.Request(url, headers={"User-Agent": "auto-dinh-muc/1.0"})
        try:
            with self._opener.open(req, timeout=timeout) as resp:
                data = resp.read()
        except (urllib.error.HTTPError, urllib.error.URLError) as e:
            raise BaseVnError(f"Download failed for {url}: {e}")
        with open(dest, "wb") as f:
            f.write(data)
        return dest

    def update_job(self, job_id, **custom) -> dict:
        """Edit a job's fields; pass custom_* keys (e.g. custom_kq=value)."""
        return self.workflow("job/edit", id=job_id, **custom)

    def move_next(self, job_id) -> dict:
        """Advance a job to the next workflow stage (job/next)."""
        return self.workflow("job/next", id=job_id)

    def list_workflows(self) -> dict:
        return self.workflow("workflows/get", page_id=0)

    def get_workflow(self, workflow_id) -> dict:
        return self.workflow("workflow/get", id=workflow_id)

    def get_stages(self, workflow_id) -> dict:
        return self.workflow("workflow/stages", id=workflow_id)

    def list_jobs(self, workflow_id=None, **kw) -> dict:
        params = dict(kw)
        if workflow_id:
            params["workflow_id"] = workflow_id
        return self.workflow("jobs/get", **params)

    # -------------------------------------------------------------- discovery
    def discover(self) -> dict:
        """Read-only probe: workflows -> stages -> a sample job, to learn the
        real column input keys. Returns the raw API data for inspection."""
        out: dict[str, Any] = {}
        wf_resp = self.list_workflows()
        data = wf_resp.get("workflows") or wf_resp.get("data") or wf_resp
        out["workflows"] = data
        return out
