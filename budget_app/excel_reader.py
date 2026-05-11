"""Excel 入金・出金ファイルの読み込み / 列判定 / 行抽出。

正規化ユーティリティ（月・金額）と FileResult データクラスもこのモジュールに集約する。
app.py からは parse_file() と各種ユーティリティを呼び出す。
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import pandas as pd


# =========================================================
# 月・金額・表示の正規化
# =========================================================
_JP_MONTH_RE = re.compile(r"(\d{1,2})\s*月")
_YEAR_RE = re.compile(r"(20\d{2}|19\d{2})")
_REIWA_RE = re.compile(r"令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月")
_AMOUNT_NOISE_RE = re.compile(r"[¥￥,，\s円]")


def normalize_month(value: Any) -> str | None:
    """様々な月表現を YYYY-MM に正規化。

    対応:
      - datetime / Timestamp
      - 'YYYY-MM' / 'YYYY/MM' / 'YYYY.MM' / 'YYYY年M月'
      - 'M月 YYYY' / 'M月YYYY (...)'
      - '令和7年9月'
      - 数値月のみ (4〜12, 1〜3) は年判定不能のため None
    """
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

    # 1) "YYYY-MM" / "YYYY/MM" / "YYYY.MM" / "YYYY年M月"
    m = re.match(r"^\s*(20\d{2}|19\d{2})[\s/\-.年]\s*(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"

    # 2) 令和YY年MM月
    m = _REIWA_RE.search(s)
    if m:
        year = 2018 + int(m.group(1))  # 令和1=2019 → 2018 + n
        return f"{year:04d}-{int(m.group(2)):02d}"

    # 3) "MM月 YYYY" / "M月YYYY (...)"
    m_month = _JP_MONTH_RE.search(s)
    m_year = _YEAR_RE.search(s)
    if m_month and m_year:
        return f"{int(m_year.group(1)):04d}-{int(m_month.group(1)):02d}"

    # 4) pandasの汎用日付パース
    try:
        ts = pd.to_datetime(s, errors="raise")
        return f"{ts.year:04d}-{ts.month:02d}"
    except Exception:
        return None


def normalize_amount(value: Any) -> float | None:
    """¥ / カンマ / 円 / 空白 を除いた数値を返す。括弧囲み・▲・△ はマイナスとして扱う。"""
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
    # 完全一致を優先
    for kw in keywords:
        for c in cols:
            if c.strip() == kw:
                return c
    # 部分一致
    for kw in keywords:
        for c in cols:
            if kw in c:
                return c
    return None


# =========================================================
# データクラス
# =========================================================
@dataclass
class FileResult:
    file_name: str
    kind: str  # "入金" or "出金"
    header_row: int
    month_col: str | None
    amount_col: str | None
    deal_col: str | None
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
    """header_row は 1始まりの行番号。pandas には 0始まりで渡す。"""
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
    kind: str,
    header_row: int,
    config: dict[str, Any],
    overrides: dict[str, Any] | None = None,
) -> FileResult:
    overrides = overrides or {}
    result = FileResult(
        file_name=file_name,
        kind=kind,
        header_row=header_row,
        month_col=None,
        amount_col=None,
        deal_col=None,
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

    month_col = overrides.get("month_col") or detect_column(columns, config["month_column_keywords"])
    amount_col = overrides.get("amount_col") or detect_column(columns, config["amount_column_keywords"])
    deal_col = overrides.get("deal_col") or detect_column(columns, config["deal_column_keywords"])

    result.month_col = month_col
    result.amount_col = amount_col
    result.deal_col = deal_col

    result.log.append(f"列判定: 月={month_col} / 金額={amount_col} / 商談={deal_col}")

    if not month_col or not amount_col:
        result.error = "月列または金額列が判定できませんでした"
        result.log.append(result.error)
        return result

    exclude_keywords = list(config.get("exclude_keywords", [])) + list(
        overrides.get("extra_exclude_keywords", []) or []
    )

    last_month_ym: str | None = None
    total_rows = 0
    included = 0
    excluded = 0

    for idx, row in df.iterrows():
        total_rows += 1
        excel_row_no = int(idx) + header_row + 1  # 元Excel上の行番号(1始まり)

        raw_month = row.get(month_col)
        raw_amount = row.get(amount_col)
        raw_deal = row.get(deal_col) if deal_col else None

        # 月の正規化（空欄なら前行を引き継ぐ）
        ym = normalize_month(raw_month)
        if ym:
            last_month_ym = ym
        else:
            ym = last_month_ym

        amount = normalize_amount(raw_amount)
        deal_str = (
            ""
            if raw_deal is None or (isinstance(raw_deal, float) and pd.isna(raw_deal))
            else str(raw_deal).strip()
        )

        exclude_reason: str | None = None

        # 1) 月が判定できない
        if ym is None:
            exclude_reason = "月判定不可"
        # 2) 金額が空欄
        elif amount is None:
            exclude_reason = "金額空欄"
        # 3) 商談名が空欄（= 小計の可能性が高い）
        elif deal_col and not deal_str:
            exclude_reason = "商談名空欄"
        else:
            # 4) 除外キーワードを含む
            text_blob = " ".join(
                str(row.get(c, "")) for c in columns if not pd.isna(row.get(c))
            )
            for kw in exclude_keywords:
                if kw and kw in text_blob:
                    exclude_reason = f"合計行（{kw}）"
                    break

        row_record = {
            "file_name": file_name,
            "kind": kind,
            "excel_row": excel_row_no,
            "month": ym,
            "amount": amount,
            "deal_name": deal_str,
            "raw_month": raw_month,
            "raw_amount": raw_amount,
        }

        if exclude_reason:
            row_record["exclude_reason"] = exclude_reason
            result.excluded_rows.append(row_record)
            excluded += 1
        else:
            result.detail_rows.append(row_record)
            included += 1

    result.log.append(
        f"読み込み行数={total_rows} / 集計対象={included} / 除外={excluded}"
    )
    return result
