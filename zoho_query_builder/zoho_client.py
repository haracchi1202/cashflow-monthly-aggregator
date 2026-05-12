"""Zoho Analytics API クライアント。

OAuth refresh + REST API 呼び出しの薄いラッパー。

公式リファレンス: https://www.zoho.com/analytics/api/v2/
"""
from __future__ import annotations

import json
import time
from dataclasses import dataclass
from typing import Any

import requests


@dataclass
class ZohoConfig:
    region: str           # "zoho.jp" / "zoho.com" / "zoho.eu" など
    client_id: str
    client_secret: str
    refresh_token: str
    org_id: str
    workspace_id: str


class ZohoAnalyticsClient:
    def __init__(self, config: ZohoConfig):
        self.config = config
        self._access_token: str | None = None
        self._token_expiry: float = 0

    # ----------------------------------------------------- auth
    @property
    def accounts_base(self) -> str:
        return f"https://accounts.{self.config.region}"

    @property
    def api_base(self) -> str:
        return f"https://analyticsapi.{self.config.region}/restapi/v2"

    def _ensure_token(self) -> str:
        """access_token を取得。期限切れ前なら再利用、期限切れなら refresh で再発行。"""
        if self._access_token and time.time() < self._token_expiry - 60:
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
            raise RuntimeError(f"OAuth refresh failed (non-JSON): {resp.text[:300]}") from e
        if "access_token" not in data:
            raise RuntimeError(f"OAuth refresh failed: {data}")
        self._access_token = data["access_token"]
        self._token_expiry = time.time() + int(data.get("expires_in", 3600))
        return self._access_token

    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Zoho-oauthtoken {self._ensure_token()}",
            "ZANALYTICS-ORGID": str(self.config.org_id),
            "Accept": "application/json",
        }

    # ----------------------------------------------------- low-level
    def request(self, method: str, path: str, *, config: dict | None = None, **kwargs) -> dict:
        url = f"{self.api_base}{path}"
        params = dict(kwargs.pop("params", {}) or {})
        data = kwargs.pop("data", None)
        if config is not None:
            cfg_str = json.dumps(config, ensure_ascii=False)
            # POST/PUT: SQL が長くなり URL 上限 (8KB) を超えるため、
            # フォームボディに CONFIG を入れる
            if method.upper() in ("POST", "PUT"):
                form_body = {"CONFIG": cfg_str}
                if isinstance(data, dict):
                    form_body.update(data)
                data = form_body
            else:
                params["CONFIG"] = cfg_str
        resp = requests.request(
            method,
            url,
            headers=self._headers(),
            params=params,
            data=data,
            timeout=60,
            **kwargs,
        )
        if resp.status_code >= 400:
            raise RuntimeError(
                f"Zoho API {method} {path} failed ({resp.status_code}): {resp.text[:500]}"
            )
        try:
            return resp.json()
        except json.JSONDecodeError:
            return {"_raw": resp.text}

    # ----------------------------------------------------- high-level
    def list_views(self) -> list[dict[str, Any]]:
        """ワークスペース内のすべての view (table / report / dashboard / query table) を返す。"""
        path = f"/workspaces/{self.config.workspace_id}/views"
        data = self.request("GET", path)
        return data.get("data", {}).get("views", [])

    def get_view_details(self, view_id: str) -> dict[str, Any]:
        """view の詳細（型・所属など）を返す。"""
        path = f"/views/{view_id}"
        data = self.request("GET", path)
        return data.get("data", {})

    def get_table_columns(self, view_id: str) -> list[str]:
        """table の列名一覧を Bulk Export API 経由で取得する。

        Zoho Analytics v2 には専用の columns 取得エンドポイントがないため、
        小さな CSV エクスポートのヘッダから列名を抽出する。
        """
        import io
        import csv
        import time

        # 1. エクスポートジョブ作成
        path = f"/bulk/workspaces/{self.config.workspace_id}/views/{view_id}/data"
        # GET メソッドで CONFIG 付きが正しい
        config = {"responseFormat": "csv"}
        data = self.request("GET", path, config=config)
        job_id = data.get("data", {}).get("jobId")
        if not job_id:
            raise RuntimeError(f"Bulk export job 作成に失敗: {data}")

        # 2. ポーリング (最大 120 秒)
        download_url: str | None = None
        for _ in range(60):
            status_path = f"/bulk/workspaces/{self.config.workspace_id}/exportjobs/{job_id}"
            st = self.request("GET", status_path)
            sd = st.get("data", {})
            status = sd.get("jobStatus")
            if status in ("JOB COMPLETED", "COMPLETED"):
                download_url = sd.get("downloadUrl")
                break
            if status in ("JOB FAILED", "FAILED"):
                raise RuntimeError(f"Export job failed: {sd}")
            time.sleep(2)
        if not download_url:
            raise RuntimeError(f"Export job タイムアウト: jobId={job_id}")

        # 3. CSV 取得 → ヘッダ行のみパース
        resp = requests.get(download_url, headers=self._headers(), timeout=60)
        if resp.status_code != 200:
            raise RuntimeError(f"CSV ダウンロード失敗 ({resp.status_code}): {resp.text[:300]}")
        # UTF-8 BOM を除去
        text = resp.content.decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(text))
        try:
            header = next(reader)
        except StopIteration:
            raise RuntimeError("CSV が空でヘッダを取得できませんでした")
        return [h.strip() for h in header]

    def create_query_table(self, name: str, sql: str, description: str = "") -> dict[str, Any]:
        """SQL を元に Query Table を作成する。"""
        path = f"/workspaces/{self.config.workspace_id}/querytables"
        config: dict[str, Any] = {
            "queryTableName": name,
            "sqlQuery": sql,
        }
        # description は description キーで送れる場合のみ追加（API バージョンで変動）
        if description:
            config["description"] = description
        return self.request("POST", path, config=config)

    def update_query_table(self, view_id: str, sql: str | None = None, name: str | None = None) -> dict[str, Any]:
        """既存 Query Table の SQL や名前を更新する。"""
        path = f"/workspaces/{self.config.workspace_id}/querytables/{view_id}"
        config: dict[str, Any] = {}
        if sql is not None:
            config["sqlQuery"] = sql
        if name is not None:
            config["queryTableName"] = name
        return self.request("PUT", path, config=config)

    def delete_view(self, view_id: str) -> dict[str, Any]:
        """指定 view (Query Table など) を削除する。"""
        path = f"/workspaces/{self.config.workspace_id}/views/{view_id}"
        return self.request("DELETE", path)

    def list_workspaces(self) -> list[dict[str, Any]]:
        """所属組織内のワークスペース一覧（org_id / workspace_id 探索用）。"""
        path = "/workspaces"
        data = self.request("GET", path)
        return data.get("data", {}).get("ownedWorkspaces", [])
