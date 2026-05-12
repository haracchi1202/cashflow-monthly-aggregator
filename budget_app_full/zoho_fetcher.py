"""Zoho Analytics から transaction Query Table を直接取得する。

提供する高レベル API:
    - load_zoho_config(): .env / st.secrets / 環境変数 から ZohoConfig を組み立て
    - ZohoFetcher: ワークスペースの view 名から CSV を取得して bytes を返す
    - fetch_transaction_xlsx_bytes(): view 名から Excel bytes に変換して返す
        （Streamlit の parse_transaction_file が読めるように pandas で Excel 化）
"""
from __future__ import annotations

import io
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from zoho_client import ZohoAnalyticsClient, ZohoConfig


@dataclass
class FetchedFile:
    """Zoho から取得した 1 つの view の結果。"""
    name: str            # view 名 (confirmed_transactions など)
    view_id: str
    filename: str        # "confirmed_transactions.xlsx" 等の表示用ファイル名
    xlsx_bytes: bytes
    row_count: int


# =========================================================
# 認証情報のロード
# =========================================================
def _load_env_file(env_path: Path) -> dict[str, str]:
    """.env を辞書として読み込み。失敗しても空を返す。"""
    out: dict[str, str] = {}
    if not env_path.exists():
        return out
    try:
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            out[k.strip()] = v.strip().strip('"').strip("'")
    except Exception:
        return {}
    return out


def load_zoho_config() -> tuple[ZohoConfig | None, str]:
    """ZohoConfig を構築。優先順位は (1) st.secrets (2) 環境変数 (3) .env。

    Returns
    -------
    (config, source)
        config: None なら認証情報なし
        source: どこから読み込んだか (説明用)
    """
    # 1) st.secrets （Streamlit Cloud / .streamlit/secrets.toml）
    try:
        import streamlit as st  # noqa: WPS433
        if hasattr(st, "secrets") and "zoho" in st.secrets:
            z = st.secrets["zoho"]
            cfg = ZohoConfig(
                region=str(z.get("region", "zoho.com")),
                client_id=str(z.get("client_id", "")),
                client_secret=str(z.get("client_secret", "")),
                refresh_token=str(z.get("refresh_token", "")),
                org_id=str(z.get("org_id", "")),
                workspace_id=str(z.get("workspace_id", "")),
            )
            if all([cfg.client_id, cfg.client_secret, cfg.refresh_token,
                    cfg.org_id, cfg.workspace_id]):
                return cfg, "st.secrets"
    except Exception:
        pass

    # 2) 環境変数 + .env
    env: dict[str, str] = {}
    env_path = Path(__file__).parent / ".env"
    env.update(_load_env_file(env_path))
    # 上位 .env も試す（zoho_query_builder/.env を再利用）
    parent_env_path = Path(__file__).parent.parent / "zoho_query_builder" / ".env"
    if parent_env_path.exists():
        for k, v in _load_env_file(parent_env_path).items():
            env.setdefault(k, v)
    # 環境変数 を最優先
    for k in ["ZOHO_REGION", "ZOHO_CLIENT_ID", "ZOHO_CLIENT_SECRET",
              "ZOHO_REFRESH_TOKEN", "ZOHO_ORG_ID", "ZOHO_WORKSPACE_ID"]:
        v = os.environ.get(k)
        if v:
            env[k] = v

    required = ["ZOHO_REGION", "ZOHO_CLIENT_ID", "ZOHO_CLIENT_SECRET",
                "ZOHO_REFRESH_TOKEN", "ZOHO_ORG_ID", "ZOHO_WORKSPACE_ID"]
    missing = [k for k in required if not env.get(k)]
    if missing:
        return None, f"認証情報が未設定: {', '.join(missing)}"
    cfg = ZohoConfig(
        region=env["ZOHO_REGION"],
        client_id=env["ZOHO_CLIENT_ID"],
        client_secret=env["ZOHO_CLIENT_SECRET"],
        refresh_token=env["ZOHO_REFRESH_TOKEN"],
        org_id=env["ZOHO_ORG_ID"],
        workspace_id=env["ZOHO_WORKSPACE_ID"],
    )
    src = ".env / 環境変数"
    return cfg, src


# =========================================================
# Fetcher
# =========================================================
class ZohoFetcher:
    def __init__(self, config: ZohoConfig):
        self.client = ZohoAnalyticsClient(config)
        self._views_cache: list[dict[str, Any]] | None = None

    def list_views(self) -> list[dict[str, Any]]:
        if self._views_cache is None:
            self._views_cache = self.client.list_views()
        return self._views_cache

    def find_view_id(self, view_name: str) -> str | None:
        """view 名から viewId を返す。完全一致優先、なければ部分一致。"""
        views = self.list_views()
        for v in views:
            if v.get("viewName") == view_name:
                return str(v.get("viewId"))
        for v in views:
            if view_name in (v.get("viewName") or ""):
                return str(v.get("viewId"))
        return None

    def fetch_csv_text(self, view_id: str) -> str:
        """view の現在の内容を CSV 文字列で取得する。"""
        import time
        import requests

        path = f"/bulk/workspaces/{self.client.config.workspace_id}/views/{view_id}/data"
        config = {"responseFormat": "csv"}
        data = self.client.request("GET", path, config=config)
        job_id = data.get("data", {}).get("jobId")
        if not job_id:
            raise RuntimeError(f"Bulk export job 作成に失敗: {data}")

        download_url: str | None = None
        for _ in range(60):
            status_path = (
                f"/bulk/workspaces/{self.client.config.workspace_id}"
                f"/exportjobs/{job_id}"
            )
            st = self.client.request("GET", status_path)
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

        resp = requests.get(download_url, headers=self.client._headers(), timeout=60)
        if resp.status_code != 200:
            raise RuntimeError(
                f"CSV ダウンロード失敗 ({resp.status_code}): {resp.text[:300]}"
            )
        return resp.content.decode("utf-8-sig", errors="replace")


def fetch_transaction_xlsx(
    fetcher: ZohoFetcher,
    view_name: str,
) -> FetchedFile:
    """view 名から CSV を取得し、parse_transaction_file が読める Excel bytes に変換する。"""
    view_id = fetcher.find_view_id(view_name)
    if not view_id:
        raise RuntimeError(
            f"Zoho ワークスペースに '{view_name}' が見つかりません。"
            "事前に zoho_query_builder で Query Table を作成してください。"
        )
    csv_text = fetcher.fetch_csv_text(view_id)
    # CSV を DataFrame にして、取引日列を datetime に変換
    df = pd.read_csv(io.StringIO(csv_text))
    df.columns = [str(c).strip() for c in df.columns]
    if "取引日" in df.columns:
        df["取引日"] = pd.to_datetime(df["取引日"], errors="coerce")
    # Excel bytes 化
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    bio.seek(0)
    return FetchedFile(
        name=view_name,
        view_id=view_id,
        filename=f"{view_name}.xlsx",
        xlsx_bytes=bio.read(),
        row_count=len(df),
    )
