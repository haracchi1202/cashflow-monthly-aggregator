"""Excel ファイル（確定入金 / 確定支払 / 予測入金 / 予測支払）の読み込みと正規化。

将来 Zoho API 連携などを追加しても再利用しやすいよう、入力ソースを問わず
共通の DataFrame スキーマに変換することを責務とする:

    columns = [
        "source_file",         # str  : 元ファイル名
        "source_group",        # str  : "確定入金" / "確定支払" / "予測入金" / "予測支払"
        "transaction_status",  # str  : "confirmed" / "forecast"
        "transaction_type",    # str  : "income" / "payment"
        "target_month",        # str  : "YYYY-MM"
        "transaction_date",    # str  : 元データの月セルそのまま（参考用）
        "amount",              # float: 円
        "client_name",         # str
        "deal_name",           # str
        "raw_row_index",       # int  : 元 Excel 上の行番号 (1始まり)
    ]
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import pandas as pd


# =========================================================
# 正規化ユーティリティ
# =========================================================
_JP_MONTH_RE = re.compile(r"(\d{1,2})\s*月")
_YEAR_RE = re.compile(r"(20\d{2}|19\d{2})")
_REIWA_RE = re.compile(r"令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月")
_AMOUNT_NOISE_RE = re.compile(r"[¥￥,，\s円]")


def normalize_month(value: Any) -> str | None:
    """様々な月表現を YYYY-MM に正規化。"""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None

    if isinstance(value, (pd.Timestamp, datetime)):
        try:
            return f"{value.year:04d}-{value.month:02d}"
        except Exception:
            return None

    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None

    m = re.match(r"^\s*(20\d{2}|19\d{2})[\s/\-.年]\s*(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"

    m = _REIWA_RE.search(s)
    if m:
        year = 2018 + int(m.group(1))
        return f"{year:04d}-{int(m.group(2)):02d}"

    m_month = _JP_MONTH_RE.search(s)
    m_year = _YEAR_RE.search(s)
    if m_month and m_year:
        return f"{int(m_year.group(1)):04d}-{int(m_month.group(1)):02d}"

    try:
        ts = pd.to_datetime(s, errors="raise")
        return f"{ts.year:04d}-{ts.month:02d}"
    except Exception:
        return None


def normalize_amount(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if pd.isna(value):
            return None
        return float(value)
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    if s.startswith("▲") or s.startswith("△"):
        neg = True
        s = s[1:]
    s = _AMOUNT_NOISE_RE.sub("", s)
    if s in ("", "-"):
        return None
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None


def format_yen(amount: float | int | None) -> str:
    if amount is None or (isinstance(amount, float) and pd.isna(amount)):
        return ""
    try:
        amt = float(amount)
    except (TypeError, ValueError):
        return ""
    sign = "-" if amt < 0 else ""
    return f"{sign}{abs(int(round(amt))):,}円"


def format_month_jp(ym: str | None) -> str:
    if not ym:
        return ""
    try:
        y, m = ym.split("-")
        return f"{int(y)}年{int(m)}月"
    except Exception:
        return ym


# =========================================================
# 列名の自動判定
# =========================================================
def detect_column(columns: list[str], keywords: list[str]) -> str | None:
    cols = [str(c) for c in columns]
    for kw in keywords:
        for c in cols:
            if c.strip() == kw:
                return c
    for kw in keywords:
        for c in cols:
            if kw in c:
                return c
    return None


# =========================================================
# データクラス
# =========================================================
GROUP_TO_STATUS_TYPE = {
    "確定入金": ("confirmed", "income"),
    "確定支払": ("confirmed", "payment"),
    "予測入金": ("forecast", "income"),
    "予測支払": ("forecast", "payment"),
}


@dataclass
class FileResult:
    file_name: str
    source_group: str           # "確定入金" など
    transaction_status: str     # "confirmed" / "forecast"
    transaction_type: str       # "income" / "payment"
    header_row: int
    month_col: str | None
    amount_col: str | None
    deal_col: str | None
    client_col: str | None
    detail_rows: list[dict[str, Any]] = field(default_factory=list)
    excluded_rows: list[dict[str, Any]] = field(default_factory=list)
    raw_head: pd.DataFrame | None = None
    error: str | None = None
    log: list[str] = field(default_factory=list)


# =========================================================
# Excel 読み込み
# =========================================================
def read_excel_with_header(
    file_bytes: bytes,
    header_row: int,
    sheet_name: int | str = 0,
) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)
    df = pd.read_excel(
        bio,
        sheet_name=sheet_name,
        header=header_row - 1,
        engine="openpyxl",
        dtype=object,
    )
    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_excel_preview(file_bytes: bytes, n: int = 30) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)
    df = pd.read_excel(bio, sheet_name=0, header=None, engine="openpyxl", dtype=object)
    return df.head(n)


# =========================================================
# 1ファイル解析
# =========================================================
def parse_file(
    file_name: str,
    file_bytes: bytes,
    source_group: str,
    header_row: int,
    config: dict[str, Any],
    overrides: dict[str, Any] | None = None,
) -> FileResult:
    """1ファイルを共通DataFrame行に変換した FileResult を返す。"""
    overrides = overrides or {}
    if source_group not in GROUP_TO_STATUS_TYPE:
        raise ValueError(f"unknown source_group: {source_group}")
    status, type_ = GROUP_TO_STATUS_TYPE[source_group]
    result = FileResult(
        file_name=file_name,
        source_group=source_group,
        transaction_status=status,
        transaction_type=type_,
        header_row=header_row,
        month_col=None,
        amount_col=None,
        deal_col=None,
        client_col=None,
    )

    try:
        df = read_excel_with_header(file_bytes, header_row)
    except Exception as e:
        result.error = f"Excel読み込み失敗: {e}"
        result.log.append(result.error)
        return result

    if df.empty:
        result.error = "データが空です"
        result.log.append(result.error)
        return result

    result.raw_head = df.head(30).copy()
    columns = list(df.columns)

    # ソースグループに応じたキーワード優先度
    type_specific_month = config.get("month_keywords_by_group", {}).get(source_group, [])
    type_specific_amount = config.get("amount_keywords_by_group", {}).get(source_group, [])

    month_keywords = type_specific_month + list(config["month_column_keywords"])
    amount_keywords = type_specific_amount + list(config["amount_column_keywords"])

    month_col = overrides.get("month_col") or detect_column(columns, month_keywords)
    amount_col = overrides.get("amount_col") or detect_column(columns, amount_keywords)
    deal_col = overrides.get("deal_col") or detect_column(columns, config["deal_column_keywords"])
    client_col = overrides.get("client_col") or detect_column(columns, config["client_column_keywords"])

    result.month_col = month_col
    result.amount_col = amount_col
    result.deal_col = deal_col
    result.client_col = client_col

    result.log.append(
        f"列判定: 月={month_col} / 金額={amount_col} / 商談={deal_col} / 顧客={client_col}"
    )

    if not month_col or not amount_col:
        result.error = "月列または金額列が判定できませんでした"
        result.log.append(result.error)
        return result

    exclude_keywords = list(config.get("exclude_keywords", [])) + list(
        overrides.get("extra_exclude_keywords", []) or []
    )

    # 予測ファイルのみ、消費税を自動加算（確定ファイルはそのまま）
    forecast_tax_rate = float(config.get("forecast_tax_rate", 0.10) or 0)
    apply_forecast_tax = (
        bool(config.get("apply_tax_to_forecast", True))
        and status == "forecast"
        and forecast_tax_rate > 0
    )
    if apply_forecast_tax:
        result.log.append(
            f"予測ファイルにつき消費税 {forecast_tax_rate * 100:.1f}% を自動加算 (税抜→税込)"
        )

    last_month_ym: str | None = None
    total_rows = 0
    included = 0
    excluded = 0

    for idx, row in df.iterrows():
        total_rows += 1
        excel_row_no = int(idx) + header_row + 1

        raw_month = row.get(month_col)
        raw_amount = row.get(amount_col)
        raw_deal = row.get(deal_col) if deal_col else None
        raw_client = row.get(client_col) if client_col else None

        ym = normalize_month(raw_month)
        if ym:
            last_month_ym = ym
        else:
            ym = last_month_ym

        amount_pretax = normalize_amount(raw_amount)
        if amount_pretax is not None and apply_forecast_tax:
            amount = round(amount_pretax * (1.0 + forecast_tax_rate))
        else:
            amount = amount_pretax
        deal_str = (
            ""
            if raw_deal is None or (isinstance(raw_deal, float) and pd.isna(raw_deal))
            else str(raw_deal).strip()
        )
        client_str = (
            ""
            if raw_client is None or (isinstance(raw_client, float) and pd.isna(raw_client))
            else str(raw_client).strip()
        )

        exclude_reason: str | None = None

        if ym is None:
            exclude_reason = "月なし"
        elif amount is None:
            exclude_reason = "金額なし"
        elif deal_col and not deal_str:
            exclude_reason = "商談名空欄"
        else:
            text_blob = " ".join(
                str(row.get(c, "")) for c in columns if not pd.isna(row.get(c))
            )
            for kw in exclude_keywords:
                if kw and kw in text_blob:
                    exclude_reason = f"合計行（{kw}）"
                    break

        record = {
            "source_file": file_name,
            "source_group": source_group,
            "transaction_status": status,
            "transaction_type": type_,
            "target_month": ym,
            "transaction_date": str(raw_month) if raw_month is not None else "",
            "amount": amount,                # 集計・予算上書きに使う値（予測は税込）
            "amount_pretax": amount_pretax,  # 元の値（参考表示用）
            "tax_applied": apply_forecast_tax,
            "client_name": client_str,
            "deal_name": deal_str,
            "raw_row_index": excel_row_no,
            "raw_month": raw_month,
            "raw_amount": raw_amount,
        }

        if exclude_reason:
            record["exclude_reason"] = exclude_reason
            result.excluded_rows.append(record)
            excluded += 1
        else:
            result.detail_rows.append(record)
            included += 1

    result.log.append(
        f"読み込み行数={total_rows} / 集計対象={included} / 除外={excluded}"
    )
    return result


# =========================================================
# Transaction 形式（縦持ち）ファイルのパース
# =========================================================
TRANSACTION_COLUMNS = [
    "商談名", "クライアント名", "取引日", "金額",
    "transaction_status", "transaction_type", "payment_round",
]
TRANSACTION_OPTIONAL_COLUMNS = ["案件番号"]  # 任意。あれば取り込む


def parse_transaction_file(
    file_name: str,
    file_bytes: bytes,
    config: dict[str, Any],
) -> list[FileResult]:
    """Zoho の confirmed/forecast_transactions エクスポート (transaction 形式 / 縦持ち)
    を読み込み、source_group ごとに FileResult を分離して返す。

    想定列:
        商談名, クライアント名, 取引日, 金額,
        transaction_status (confirmed|forecast),
        transaction_type (income|payment),
        payment_round (元項目名)

    戻り値: source_group ごとに 1つ の FileResult (確定入金 / 確定支払 / 予測入金 / 予測支払)
    """
    bio = io.BytesIO(file_bytes)
    try:
        df = pd.read_excel(bio, sheet_name=0, engine="openpyxl", dtype=object)
    except Exception as e:
        # 1つの FileResult にエラーを入れて返す
        err = FileResult(
            file_name=file_name, source_group="(不明)",
            transaction_status="unknown", transaction_type="unknown",
            header_row=1,
            month_col=None, amount_col=None, deal_col=None, client_col=None,
            error=f"Excel読み込み失敗: {e}",
        )
        return [err]

    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in TRANSACTION_COLUMNS if c not in df.columns]
    if missing:
        err = FileResult(
            file_name=file_name, source_group="(不明)",
            transaction_status="unknown", transaction_type="unknown",
            header_row=1,
            month_col=None, amount_col=None, deal_col=None, client_col=None,
            error=f"必須列が不足: {missing}。実際の列: {list(df.columns)}",
        )
        return [err]

    forecast_tax_rate = float(config.get("forecast_tax_rate", 0.10) or 0)
    apply_tax_for_forecast = bool(config.get("apply_tax_to_forecast", True))

    # source_group ごとに行を集める
    by_group: dict[str, list[dict[str, Any]]] = {}
    excluded_by_group: dict[str, list[dict[str, Any]]] = {}

    for idx, row in df.iterrows():
        excel_row_no = int(idx) + 2  # ヘッダ行が 1 なのでデータは 2 始まり
        status = str(row.get("transaction_status") or "").strip().lower()
        type_ = str(row.get("transaction_type") or "").strip().lower()
        if status not in ("confirmed", "forecast") or type_ not in ("income", "payment"):
            # 不正行はスキップ（除外行として保存）
            group = "(不明)"
            excluded_by_group.setdefault(group, []).append({
                "source_file": file_name,
                "source_group": group,
                "transaction_status": status,
                "transaction_type": type_,
                "raw_row_index": excel_row_no,
                "exclude_reason": "transaction_status/type が不正",
                "amount": None, "target_month": None,
                "deal_name": "", "client_name": "",
            })
            continue

        status_jp = "確定" if status == "confirmed" else "予測"
        type_jp = "入金" if type_ == "income" else "支払"
        group = status_jp + type_jp

        raw_date = row.get("取引日")
        raw_amount = row.get("金額")
        raw_deal = row.get("商談名")
        raw_client = row.get("クライアント名")
        raw_round = row.get("payment_round")
        raw_deal_id = row.get("案件番号") if "案件番号" in df.columns else None

        ym = normalize_month(raw_date)
        amount_pretax = normalize_amount(raw_amount)

        tax_applied = False
        if (
            status == "forecast"
            and apply_tax_for_forecast
            and amount_pretax is not None
            and forecast_tax_rate > 0
        ):
            amount = round(amount_pretax * (1.0 + forecast_tax_rate))
            tax_applied = True
        else:
            amount = amount_pretax

        deal_str = "" if raw_deal is None or (isinstance(raw_deal, float) and pd.isna(raw_deal)) else str(raw_deal).strip()
        client_str = "" if raw_client is None or (isinstance(raw_client, float) and pd.isna(raw_client)) else str(raw_client).strip()
        round_str = "" if raw_round is None or (isinstance(raw_round, float) and pd.isna(raw_round)) else str(raw_round).strip()
        deal_id_str = "" if raw_deal_id is None or (isinstance(raw_deal_id, float) and pd.isna(raw_deal_id)) else str(raw_deal_id).strip()

        exclude_reason: str | None = None
        if ym is None:
            exclude_reason = "月なし"
        elif amount is None:
            exclude_reason = "金額なし"
        elif amount == 0:
            exclude_reason = "金額ゼロ"
        elif not deal_str:
            exclude_reason = "商談名空欄"

        record = {
            "source_file": file_name,
            "source_group": group,
            "transaction_status": status,
            "transaction_type": type_,
            "target_month": ym,
            "transaction_date": str(raw_date) if raw_date is not None else "",
            "amount": amount,
            "amount_pretax": amount_pretax,
            "tax_applied": tax_applied,
            "client_name": client_str,
            "deal_name": deal_str,
            "deal_id": deal_id_str,
            "payment_round": round_str,
            "raw_row_index": excel_row_no,
            "raw_month": raw_date,
            "raw_amount": raw_amount,
        }

        if exclude_reason:
            record["exclude_reason"] = exclude_reason
            excluded_by_group.setdefault(group, []).append(record)
        else:
            by_group.setdefault(group, []).append(record)

    # source_group ごとに FileResult を作成
    results: list[FileResult] = []
    all_groups = set(by_group.keys()) | set(excluded_by_group.keys())
    for group in sorted(all_groups):
        status, type_ = GROUP_TO_STATUS_TYPE.get(group, ("unknown", "unknown"))
        fr = FileResult(
            file_name=file_name,
            source_group=group,
            transaction_status=status,
            transaction_type=type_,
            header_row=1,
            month_col="取引日",
            amount_col="金額",
            deal_col="商談名",
            client_col="クライアント名",
            detail_rows=by_group.get(group, []),
            excluded_rows=excluded_by_group.get(group, []),
        )
        included = len(fr.detail_rows)
        excluded = len(fr.excluded_rows)
        fr.log.append(f"transaction形式: 集計対象={included} / 除外={excluded}")
        if status == "forecast" and apply_tax_for_forecast and forecast_tax_rate > 0:
            fr.log.append(f"予測ファイルにつき消費税 {forecast_tax_rate*100:.1f}% を自動加算（税抜→税込）")
        results.append(fr)
    return results


def results_to_records(
    results: list[FileResult],
    restored_keys: set[str],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """FileResult リストを (集計対象レコード, 除外レコード) のリストに展開。
    復活キーが付いた除外行は集計対象に戻す。
    """
    detail_records: list[dict[str, Any]] = []
    excluded_records: list[dict[str, Any]] = []
    for r in results:
        for row in r.detail_rows:
            detail_records.append(row)
        for row in r.excluded_rows:
            key = f"{row['source_file']}::{row['raw_row_index']}"
            if (
                key in restored_keys
                and row.get("target_month")
                and row.get("amount") is not None
            ):
                restored = dict(row)
                restored["restored"] = True
                detail_records.append(restored)
            else:
                excluded_records.append(row)
    return detail_records, excluded_records
