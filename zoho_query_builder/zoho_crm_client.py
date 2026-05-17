"""Zoho CRM API v8 用の薄いクライアント（Self Client refresh token 方式）。

主に商談モジュールへのカスタムフィールド追加で使用。
"""
from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests


THIS_DIR = Path(__file__).parent


def _load_env(env_path: Path | None = None) -> dict[str, str]:
    """`.env` を辞書として読み込む。`KEY=value # comment` 形式の行末コメントも除去。"""
    env: dict[str, str] = dict(os.environ)
    p = env_path or (THIS_DIR / ".env")
    if not p.exists():
        return env
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        # クォートを剥がしてから行末コメント除去
        v = v.strip().strip('"').strip("'")
        v = v.split("#", 1)[0].strip()
        env[k.strip()] = v
    return env


@dataclass
class ZohoCRMConfig:
    region: str            # "zoho.jp" / "zoho.com" / "zoho.eu" など
    client_id: str
    client_secret: str
    refresh_token: str


class ZohoCRMClient:
    def __init__(self, config: ZohoCRMConfig) -> None:
        if not config.region.startswith("zoho."):
            raise ValueError(
                f"ZohoCRMConfig.region は 'zoho.jp' 形式で指定してください: got {config.region!r}"
            )
        self.config = config
        self._access_token: str | None = None
        self._expiry: float = 0.0

    @classmethod
    def from_env(cls, env_path: Path | None = None) -> "ZohoCRMClient":
        env = _load_env(env_path)
        required = ("ZOHO_REGION", "ZOHO_CLIENT_ID", "ZOHO_CLIENT_SECRET", "ZOHO_CRM_REFRESH_TOKEN")
        missing = [k for k in required if not env.get(k)]
        if missing:
            raise RuntimeError(f".env に未設定のキーがあります: {', '.join(missing)}")
        cfg = ZohoCRMConfig(
            region=env["ZOHO_REGION"],
            client_id=env["ZOHO_CLIENT_ID"],
            client_secret=env["ZOHO_CLIENT_SECRET"],
            refresh_token=env["ZOHO_CRM_REFRESH_TOKEN"],
        )
        return cls(cfg)

    # ---------- auth ----------
    @property
    def accounts_base(self) -> str:
        return f"https://accounts.{self.config.region}"

    @property
    def api_base(self) -> str:
        # CRM API は region によって異なる（jp は www.zohoapis.jp）
        host = self.config.region.replace("zoho.", "zohoapis.")
        return f"https://www.{host}/crm/v8"

    def _ensure_token(self) -> str:
        if self._access_token and time.time() < self._expiry - 60:
            return self._access_token
        url = f"{self.accounts_base}/oauth/v2/token"
        params = {
            "refresh_token": self.config.refresh_token,
            "client_id": self.config.client_id,
            "client_secret": self.config.client_secret,
            "grant_type": "refresh_token",
        }
        resp = requests.post(url, params=params, timeout=30)
        try:
            data = resp.json()
        except json.JSONDecodeError as e:
            raise RuntimeError(f"CRM OAuth refresh failed (non-JSON): {resp.text[:300]}") from e
        if "access_token" not in data:
            raise RuntimeError(f"CRM OAuth refresh failed: {data}")
        self._access_token = data["access_token"]
        self._expiry = time.time() + int(data.get("expires_in", 3600))
        return self._access_token

    def _invalidate_token(self) -> None:
        self._access_token = None
        self._expiry = 0.0

    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Zoho-oauthtoken {self._ensure_token()}",
            "Accept": "application/json",
        }

    # ---------- low-level ----------
    def request(self, method: str, path: str, *, json_body: Any = None, params: dict | None = None) -> dict:
        """1 回だけ 401 を受けた場合に自動でトークン再発行＋リトライする。"""
        for attempt in (1, 2):
            url = f"{self.api_base}{path}"
            headers = self._headers()
            if json_body is not None:
                headers["Content-Type"] = "application/json"
            resp = requests.request(
                method, url, headers=headers,
                params=params,
                data=json.dumps(json_body, ensure_ascii=False).encode("utf-8") if json_body is not None else None,
                timeout=60,
            )
            if resp.status_code == 401 and attempt == 1:
                # access_token 失効を想定。1 回だけ refresh してリトライ
                self._invalidate_token()
                continue
            if resp.status_code >= 400:
                raise RuntimeError(f"CRM {method} {path} failed ({resp.status_code}): {resp.text[:600]}")
            try:
                return resp.json()
            except json.JSONDecodeError:
                return {"_raw": resp.text}
        # ここには到達しない（continue 後の 2 周目で必ず return か raise）
        raise RuntimeError(f"CRM {method} {path}: unreachable retry loop end")

    # ---------- high-level ----------
    def list_modules(self) -> list[dict[str, Any]]:
        return self.request("GET", "/settings/modules").get("modules", [])

    def list_fields(self, module: str) -> list[dict[str, Any]]:
        return self.request("GET", "/settings/fields", params={"module": module}).get("fields", [])

    def create_text_field(self, module: str, field_label: str, api_name: str, max_length: int = 100) -> dict:
        """Single Line Text のカスタムフィールドを追加。"""
        body = {
            "fields": [
                {
                    "field_label": field_label,
                    "api_name": api_name,
                    "data_type": "text",
                    "length": max_length,
                }
            ]
        }
        return self.request("POST", "/settings/fields", json_body=body, params={"module": module})
