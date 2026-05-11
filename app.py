"""月別 入金・出金 集計アプリ (Streamlit)

Excelファイルをアップロードして月別の入金・出金を集計するWebアプリ。
"""
from __future__ import annotations

import io
import json
import re
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

CONFIG_PATH = Path(__file__).parent / "config.json"

DEFAULT_CONFIG: dict[str, Any] = {
    "default_header_row": 7,
    "month_column_keywords": ["入金日", "支払日", "支払い日", "入金月", "支払月"],
    "amount_column_keywords": ["税込", "合計", "原価総額", "入金額", "金額"],
    "deal_column_keywords": ["商談名", "案件名", "取引先", "件名", "摘要"],
    "exclude_keywords": ["合計", "小計", "総計", "Grand Total", "Total", "計"],
    "file_overrides": {},
}


# =========================================================
# 設定の読み書き
# =========================================================
def load_config() -> dict[str, Any]:
    if CONFIG_PATH.exists():
        try:
            with CONFIG_PATH.open("r", encoding="utf-8") as f:
                cfg = json.load(f)
            merged = DEFAULT_CONFIG.copy()
            merged.update(cfg)
            return merged
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()


def save_config(cfg: dict[str, Any]) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# =========================================================
# 月・金額の正規化
# =========================================================
_JP_MONTH_RE = re.compile(r"(\d{1,2})\s*月")
_YEAR_RE = re.compile(r"(20\d{2}|19\d{2})")


def normalize_month(value: Any) -> str | None:
    """「9月 2025 ( 23 )」「2025/9/1」「2025-09-01」などを YYYY-MM に正規化。"""
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

    # 1) "YYYY-MM" / "YYYY/MM" / "YYYY.MM"
    m = re.match(r"^\s*(20\d{2}|19\d{2})[\s/\-.年]\s*(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"

    # 2) "MM月 YYYY" or "M月 YYYY (...)" or "M月YYYY"
    m_month = _JP_MONTH_RE.search(s)
    m_year = _YEAR_RE.search(s)
    if m_month and m_year:
        return f"{int(m_year.group(1)):04d}-{int(m_month.group(1)):02d}"

    # 3) pandasの汎用日付パース
    try:
        ts = pd.to_datetime(s, errors="raise")
        return f"{ts.year:04d}-{ts.month:02d}"
    except Exception:
        return None


_AMOUNT_NOISE_RE = re.compile(r"[¥￥,，\s円]")


def normalize_amount(value: Any) -> float | None:
    """¥ / カンマ / 円 / 空白 を除いた数値を返す。括弧囲みはマイナスとして扱う。"""
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
    sign = "-" if amount < 0 else ""
    return f"{sign}{abs(int(round(amount))):,}円"


def format_month_jp(ym: str) -> str:
    """'2025-09' -> '2025年9月'"""
    try:
        y, m = ym.split("-")
        return f"{int(y)}年{int(m)}月"
    except Exception:
        return ym


# =========================================================
# 列名の自動判定
# =========================================================
def detect_column(columns: list[str], keywords: list[str]) -> str | None:
    """カラム名の中から keywords を含む最初の列名を返す。"""
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
# ファイル読み込み・解析
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

    exclude_keywords = list(config.get("exclude_keywords", [])) + list(overrides.get("extra_exclude_keywords", []))

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
        deal_str = "" if raw_deal is None or (isinstance(raw_deal, float) and pd.isna(raw_deal)) else str(raw_deal).strip()

        # 主要列を抜粋
        major = {c: row.get(c) for c in columns[:8]}

        exclude_reason: str | None = None

        # 1) 月が判定できない
        if ym is None:
            exclude_reason = "月が判定できない"
        # 2) 金額が空欄
        elif amount is None:
            exclude_reason = "金額が空欄"
        # 3) 商談名が空欄（= 小計の可能性が高い）
        elif deal_col and not deal_str:
            exclude_reason = "商談名が空欄"
        else:
            # 4) 除外キーワードを含む
            text_blob = " ".join(
                str(row.get(c, "")) for c in columns if not pd.isna(row.get(c))
            )
            for kw in exclude_keywords:
                if kw and kw in text_blob:
                    exclude_reason = f"合計行・小計行の可能性 ({kw})"
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
            **{f"col::{c}": row.get(c) for c in columns[:8]},
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


# =========================================================
# 集計
# =========================================================
def build_dataframes(
    results: list[FileResult],
    restored_keys: set[str],
    month_from: str | None,
    month_to: str | None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """(月別集計, 入金明細, 出金明細, 除外行, ファイル別集計) を返す。"""
    detail_records: list[dict[str, Any]] = []
    excluded_records: list[dict[str, Any]] = []

    for r in results:
        for row in r.detail_rows:
            detail_records.append(row)
        for row in r.excluded_rows:
            key = f"{row['file_name']}::{row['excel_row']}"
            if key in restored_keys and row.get("month") and row.get("amount") is not None:
                # 復活分は明細に追加
                detail_records.append(row)
            else:
                excluded_records.append(row)

    detail_df = pd.DataFrame(detail_records)
    excluded_df = pd.DataFrame(excluded_records)

    if not detail_df.empty:
        # 期間フィルター
        if month_from:
            detail_df = detail_df[detail_df["month"] >= month_from]
        if month_to:
            detail_df = detail_df[detail_df["month"] <= month_to]

    income_df = detail_df[detail_df["kind"] == "入金"].copy() if not detail_df.empty else pd.DataFrame()
    expense_df = detail_df[detail_df["kind"] == "出金"].copy() if not detail_df.empty else pd.DataFrame()

    # 月別集計
    if detail_df.empty:
        monthly_df = pd.DataFrame(columns=["month", "入金合計", "出金合計", "差額"])
    else:
        inc = (
            income_df.groupby("month", as_index=False)["amount"].sum().rename(columns={"amount": "入金合計"})
            if not income_df.empty
            else pd.DataFrame(columns=["month", "入金合計"])
        )
        exp = (
            expense_df.groupby("month", as_index=False)["amount"].sum().rename(columns={"amount": "出金合計"})
            if not expense_df.empty
            else pd.DataFrame(columns=["month", "出金合計"])
        )
        monthly_df = pd.merge(inc, exp, on="month", how="outer").fillna(0.0)
        monthly_df["差額"] = monthly_df["入金合計"] - monthly_df["出金合計"]
        monthly_df = monthly_df.sort_values("month").reset_index(drop=True)

    # ファイル別集計
    file_records = []
    for r in results:
        included_amount = sum(
            (row.get("amount") or 0)
            for row in r.detail_rows
        )
        included_count = len(r.detail_rows)
        excluded_count = len(r.excluded_rows)
        # 復活分も足す
        restored_in_file = [
            row for row in r.excluded_rows
            if f"{row['file_name']}::{row['excel_row']}" in restored_keys
            and row.get("amount") is not None
        ]
        included_amount += sum((row.get("amount") or 0) for row in restored_in_file)
        included_count += len(restored_in_file)
        excluded_count -= len(restored_in_file)

        file_records.append({
            "ファイル名": r.file_name,
            "区分": r.kind,
            "合計金額": included_amount,
            "集計件数": included_count,
            "除外件数": excluded_count,
        })
    file_summary_df = pd.DataFrame(file_records)

    return monthly_df, income_df, expense_df, excluded_df, file_summary_df


# =========================================================
# Excel書き出し
# =========================================================
def build_excel(
    monthly_df: pd.DataFrame,
    income_df: pd.DataFrame,
    expense_df: pd.DataFrame,
    excluded_df: pd.DataFrame,
    file_summary_df: pd.DataFrame,
    log_lines: list[str],
    config: dict[str, Any],
) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        workbook = writer.book
        yen_fmt = workbook.add_format({"num_format": "#,##0\"円\""})
        neg_yen_fmt = workbook.add_format({"num_format": "#,##0\"円\";[Red]-#,##0\"円\""})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})

        # 月別集計
        m = monthly_df.copy()
        if not m.empty:
            m["月"] = m["month"].apply(format_month_jp)
            m = m[["月", "入金合計", "出金合計", "差額"]]
        else:
            m = pd.DataFrame(columns=["月", "入金合計", "出金合計", "差額"])
        m.to_excel(writer, sheet_name="月別集計", index=False)
        ws = writer.sheets["月別集計"]
        ws.set_column("A:A", 14)
        ws.set_column("B:C", 18, yen_fmt)
        ws.set_column("D:D", 18, neg_yen_fmt)
        for col_num, val in enumerate(m.columns):
            ws.write(0, col_num, val, header_fmt)

        # 入金明細
        _write_detail_sheet(writer, "入金明細", income_df, yen_fmt, header_fmt)
        # 出金明細
        _write_detail_sheet(writer, "出金明細", expense_df, yen_fmt, header_fmt)
        # 除外行一覧
        if not excluded_df.empty:
            ex = excluded_df.copy()
            keep = ["file_name", "excel_row", "month", "amount", "exclude_reason", "deal_name"]
            ex = ex[[c for c in keep if c in ex.columns]]
            ex.columns = ["ファイル名", "行番号", "月", "金額", "除外理由", "商談名"]
        else:
            ex = pd.DataFrame(columns=["ファイル名", "行番号", "月", "金額", "除外理由", "商談名"])
        ex.to_excel(writer, sheet_name="除外行一覧", index=False)
        ws = writer.sheets["除外行一覧"]
        ws.set_column("A:A", 36)
        ws.set_column("B:B", 8)
        ws.set_column("C:C", 10)
        ws.set_column("D:D", 16, yen_fmt)
        ws.set_column("E:E", 22)
        ws.set_column("F:F", 36)
        for col_num, val in enumerate(ex.columns):
            ws.write(0, col_num, val, header_fmt)

        # ファイル別集計
        fs = file_summary_df.copy() if not file_summary_df.empty else pd.DataFrame(
            columns=["ファイル名", "区分", "合計金額", "集計件数", "除外件数"]
        )
        fs.to_excel(writer, sheet_name="ファイル別集計", index=False)
        ws = writer.sheets["ファイル別集計"]
        ws.set_column("A:A", 36)
        ws.set_column("B:B", 8)
        ws.set_column("C:C", 18, yen_fmt)
        ws.set_column("D:E", 12)
        for col_num, val in enumerate(fs.columns):
            ws.write(0, col_num, val, header_fmt)

        # 設定・処理ログ
        log_df = pd.DataFrame({"ログ": log_lines})
        cfg_lines = [
            f"月列キーワード: {', '.join(config['month_column_keywords'])}",
            f"金額列キーワード: {', '.join(config['amount_column_keywords'])}",
            f"商談列キーワード: {', '.join(config['deal_column_keywords'])}",
            f"除外キーワード: {', '.join(config['exclude_keywords'])}",
            f"デフォルトヘッダー行: {config['default_header_row']}",
            f"出力日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        ]
        meta_df = pd.DataFrame({"設定": cfg_lines})
        meta_df.to_excel(writer, sheet_name="設定・処理ログ", index=False, startrow=0)
        log_df.to_excel(writer, sheet_name="設定・処理ログ", index=False, startrow=len(cfg_lines) + 3)
        ws = writer.sheets["設定・処理ログ"]
        ws.set_column("A:A", 80)

    bio.seek(0)
    return bio.read()


def _write_detail_sheet(writer, sheet_name, df, yen_fmt, header_fmt):
    if df.empty:
        out = pd.DataFrame(columns=["ファイル名", "行番号", "月", "金額", "商談名"])
    else:
        out = df.copy()
        keep = ["file_name", "excel_row", "month", "amount", "deal_name"]
        out = out[[c for c in keep if c in out.columns]]
        out.columns = ["ファイル名", "行番号", "月", "金額", "商談名"]
    out.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    ws.set_column("A:A", 36)
    ws.set_column("B:B", 8)
    ws.set_column("C:C", 10)
    ws.set_column("D:D", 16, yen_fmt)
    ws.set_column("E:E", 40)
    for col_num, val in enumerate(out.columns):
        ws.write(0, col_num, val, header_fmt)


# =========================================================
# UI
# =========================================================
def init_session() -> None:
    if "config" not in st.session_state:
        st.session_state.config = load_config()
    if "parsed_results" not in st.session_state:
        st.session_state.parsed_results = []  # list[FileResult]
    if "restored_keys" not in st.session_state:
        st.session_state.restored_keys = set()
    if "header_rows" not in st.session_state:
        st.session_state.header_rows = {}  # file_name -> int
    if "overrides" not in st.session_state:
        st.session_state.overrides = {}  # file_name -> dict


def render_upload_area(label: str, key: str) -> list:
    files = st.file_uploader(
        label,
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key=key,
    )
    return files or []


def render_preview_section(income_files: list, expense_files: list) -> None:
    st.subheader("📄 データ確認モード（先頭30行プレビュー）")
    st.caption("ヘッダー行がずれている場合は、ヘッダー行番号を変更して再読み込みできます。")

    all_files = [(f, "入金") for f in income_files] + [(f, "出金") for f in expense_files]
    if not all_files:
        st.info("ファイルをアップロードしてください。")
        return

    default_hdr = st.session_state.config["default_header_row"]

    for uf, kind in all_files:
        with st.expander(f"[{kind}] {uf.name}", expanded=False):
            current_hdr = st.session_state.header_rows.get(uf.name, default_hdr)
            new_hdr = st.number_input(
                f"ヘッダー行番号（1始まり） - {uf.name}",
                min_value=1,
                max_value=50,
                value=int(current_hdr),
                key=f"hdr_{uf.name}",
            )
            st.session_state.header_rows[uf.name] = int(new_hdr)

            try:
                preview_df = read_excel_preview(uf.getvalue(), n=30)
                preview_df.index = [i + 1 for i in preview_df.index]
                st.dataframe(preview_df, use_container_width=True, height=320)
            except Exception as e:
                st.error(f"プレビュー失敗: {e}")


def render_column_overrides(income_files: list, expense_files: list) -> None:
    cfg = st.session_state.config
    default_hdr = cfg["default_header_row"]

    st.subheader("🛠️ 列名の手動指定（任意）")
    st.caption("自動判定が間違った場合のみ指定してください。空欄なら自動判定が使われます。")

    all_files = [(f, "入金") for f in income_files] + [(f, "出金") for f in expense_files]
    if not all_files:
        return

    for uf, kind in all_files:
        with st.expander(f"[{kind}] {uf.name} の列指定", expanded=False):
            hdr = st.session_state.header_rows.get(uf.name, default_hdr)
            try:
                df = read_excel_with_header(uf.getvalue(), hdr)
                cols = [""] + list(df.columns)
            except Exception as e:
                st.error(f"列の取得に失敗: {e}")
                continue

            ov = st.session_state.overrides.get(uf.name, {})

            def _idx(value: str | None) -> int:
                if value and value in cols:
                    return cols.index(value)
                return 0

            month_col = st.selectbox(
                "月列", cols, index=_idx(ov.get("month_col")), key=f"m_{uf.name}"
            )
            amount_col = st.selectbox(
                "金額列", cols, index=_idx(ov.get("amount_col")), key=f"a_{uf.name}"
            )
            deal_col = st.selectbox(
                "商談名列", cols, index=_idx(ov.get("deal_col")), key=f"d_{uf.name}"
            )
            extra_kw_str = st.text_input(
                "追加除外キーワード（カンマ区切り）",
                value=",".join(ov.get("extra_exclude_keywords", [])),
                key=f"kw_{uf.name}",
            )
            extra_kw = [k.strip() for k in extra_kw_str.split(",") if k.strip()]

            st.session_state.overrides[uf.name] = {
                "month_col": month_col or None,
                "amount_col": amount_col or None,
                "deal_col": deal_col or None,
                "extra_exclude_keywords": extra_kw,
            }


def run_parsing(income_files: list, expense_files: list) -> list[FileResult]:
    cfg = st.session_state.config
    default_hdr = cfg["default_header_row"]
    results: list[FileResult] = []

    progress = st.progress(0.0, text="解析中...")
    total = len(income_files) + len(expense_files)
    done = 0

    for uf in income_files:
        hdr = st.session_state.header_rows.get(uf.name, default_hdr)
        ov = st.session_state.overrides.get(uf.name, {})
        try:
            r = parse_file(uf.name, uf.getvalue(), "入金", hdr, cfg, ov)
        except Exception as e:
            r = FileResult(
                file_name=uf.name, kind="入金", header_row=hdr,
                month_col=None, amount_col=None, deal_col=None,
                error=f"予期せぬ例外: {e}",
            )
            r.log.append(traceback.format_exc())
        results.append(r)
        done += 1
        progress.progress(done / max(total, 1), text=f"解析中... {uf.name}")

    for uf in expense_files:
        hdr = st.session_state.header_rows.get(uf.name, default_hdr)
        ov = st.session_state.overrides.get(uf.name, {})
        try:
            r = parse_file(uf.name, uf.getvalue(), "出金", hdr, cfg, ov)
        except Exception as e:
            r = FileResult(
                file_name=uf.name, kind="出金", header_row=hdr,
                month_col=None, amount_col=None, deal_col=None,
                error=f"予期せぬ例外: {e}",
            )
            r.log.append(traceback.format_exc())
        results.append(r)
        done += 1
        progress.progress(done / max(total, 1), text=f"解析中... {uf.name}")

    progress.empty()
    return results


def render_monthly_summary(monthly_df: pd.DataFrame) -> None:
    st.subheader("📊 月別集計")
    if monthly_df.empty:
        st.info("集計対象データがありません。")
        return

    display = monthly_df.copy()
    display["月"] = display["month"].apply(format_month_jp)
    display["入金合計"] = display["入金合計"].apply(format_yen)
    display["出金合計"] = display["出金合計"].apply(format_yen)
    display["差額表示"] = monthly_df["差額"].apply(
        lambda v: f"▲ {format_yen(abs(v))}" if v < 0 else format_yen(v)
    )
    display = display[["月", "入金合計", "出金合計", "差額表示"]].rename(columns={"差額表示": "差額"})
    st.dataframe(display, use_container_width=True, hide_index=True)

    total_in = monthly_df["入金合計"].sum()
    total_out = monthly_df["出金合計"].sum()
    diff = total_in - total_out
    c1, c2, c3 = st.columns(3)
    c1.metric("入金合計", format_yen(total_in))
    c2.metric("出金合計", format_yen(total_out))
    c3.metric("差額", format_yen(diff), delta=("マイナス" if diff < 0 else "プラス"))


def render_drill_down(monthly_df: pd.DataFrame, income_df: pd.DataFrame, expense_df: pd.DataFrame) -> None:
    st.subheader("🔍 月別詳細（ドリルダウン）")
    if monthly_df.empty:
        st.info("集計対象データがありません。")
        return

    months = monthly_df["month"].tolist()
    selected = st.selectbox("月を選択", months, format_func=format_month_jp)
    if not selected:
        return

    tab_in, tab_out = st.tabs(["入金明細", "出金明細"])

    with tab_in:
        sub = income_df[income_df["month"] == selected] if not income_df.empty else pd.DataFrame()
        _render_detail_table(sub, kind="入金")

    with tab_out:
        sub = expense_df[expense_df["month"] == selected] if not expense_df.empty else pd.DataFrame()
        _render_detail_table(sub, kind="出金")


def _render_detail_table(df: pd.DataFrame, kind: str) -> None:
    if df.empty:
        st.info(f"{kind}明細はありません。")
        return
    show = df[["file_name", "excel_row", "month", "deal_name", "amount"]].copy()
    show.columns = ["ファイル名", "行番号", "月", "商談名", "金額"]
    show["月"] = show["月"].apply(format_month_jp)
    show["金額"] = show["金額"].apply(format_yen)
    st.dataframe(show, use_container_width=True, hide_index=True)
    st.caption(f"件数: {len(df)} 件 / 合計: {format_yen(df['amount'].sum())}")


def render_file_summary(file_summary_df: pd.DataFrame) -> None:
    st.subheader("📁 ファイル別集計")
    if file_summary_df.empty:
        st.info("データがありません。")
        return
    disp = file_summary_df.copy()
    disp["合計金額"] = disp["合計金額"].apply(format_yen)
    st.dataframe(disp, use_container_width=True, hide_index=True)


def render_excluded_rows(results: list[FileResult]) -> None:
    st.subheader("🚫 除外行確認 & 手動復活")
    rows: list[dict[str, Any]] = []
    for r in results:
        for row in r.excluded_rows:
            rows.append(row)
    if not rows:
        st.info("除外行はありません。")
        return

    df = pd.DataFrame(rows)
    df["key"] = df.apply(lambda x: f"{x['file_name']}::{x['excel_row']}", axis=1)
    df["復活"] = df["key"].apply(lambda k: k in st.session_state.restored_keys)
    show_cols = ["復活", "file_name", "excel_row", "month", "amount", "exclude_reason", "deal_name"]
    show = df[show_cols].copy()
    show.columns = ["復活", "ファイル名", "行番号", "月", "金額", "除外理由", "商談名"]
    show["月"] = show["月"].apply(lambda v: format_month_jp(v) if v else "")
    show["金額"] = show["金額"].apply(format_yen)

    edited = st.data_editor(
        show,
        use_container_width=True,
        hide_index=True,
        disabled=["ファイル名", "行番号", "月", "金額", "除外理由", "商談名"],
        column_config={"復活": st.column_config.CheckboxColumn("復活", help="チェックすると集計対象に戻します")},
        key="excluded_editor",
    )

    # 反映
    new_keys: set[str] = set()
    for i, key in enumerate(df["key"].tolist()):
        if bool(edited.iloc[i]["復活"]):
            new_keys.add(key)
    if new_keys != st.session_state.restored_keys:
        st.session_state.restored_keys = new_keys
        st.rerun()


def render_processing_log(results: list[FileResult]) -> list[str]:
    st.subheader("📝 処理ログ")
    lines: list[str] = []
    for r in results:
        head = f"[{r.kind}] {r.file_name}（ヘッダー行: {r.header_row}）"
        lines.append(head)
        lines.append(f"  - 月列: {r.month_col} / 金額列: {r.amount_col} / 商談列: {r.deal_col}")
        for log in r.log:
            lines.append(f"  - {log}")
        if r.error:
            lines.append(f"  ❌ エラー: {r.error}")
    if not lines:
        st.info("ログはまだありません。")
    else:
        st.code("\n".join(lines), language="text")
    return lines


def render_config_section() -> None:
    st.subheader("⚙️ 設定の保存・読み込み")
    cfg = st.session_state.config

    with st.expander("キーワード設定（自動判定の調整）", expanded=False):
        month_kw = st.text_input(
            "月列キーワード (カンマ区切り)",
            value=",".join(cfg["month_column_keywords"]),
        )
        amount_kw = st.text_input(
            "金額列キーワード (カンマ区切り)",
            value=",".join(cfg["amount_column_keywords"]),
        )
        deal_kw = st.text_input(
            "商談名列キーワード (カンマ区切り)",
            value=",".join(cfg["deal_column_keywords"]),
        )
        excl_kw = st.text_input(
            "除外キーワード (カンマ区切り)",
            value=",".join(cfg["exclude_keywords"]),
        )
        default_hdr = st.number_input(
            "デフォルトのヘッダー行",
            min_value=1, max_value=50, value=int(cfg["default_header_row"]),
        )
        cfg["month_column_keywords"] = [k.strip() for k in month_kw.split(",") if k.strip()]
        cfg["amount_column_keywords"] = [k.strip() for k in amount_kw.split(",") if k.strip()]
        cfg["deal_column_keywords"] = [k.strip() for k in deal_kw.split(",") if k.strip()]
        cfg["exclude_keywords"] = [k.strip() for k in excl_kw.split(",") if k.strip()]
        cfg["default_header_row"] = int(default_hdr)

    c1, c2, c3 = st.columns(3)
    if c1.button("💾 設定を保存", use_container_width=True):
        # 列指定とヘッダー行もまとめて保存
        file_overrides = {}
        for fname, ov in st.session_state.overrides.items():
            file_overrides[fname] = {
                **ov,
                "header_row": st.session_state.header_rows.get(fname),
            }
        cfg["file_overrides"] = file_overrides
        save_config(cfg)
        st.success("config.json に保存しました。")

    if c2.button("🔄 設定を再読み込み", use_container_width=True):
        st.session_state.config = load_config()
        st.success("config.json から再読み込みしました。")
        st.rerun()

    cfg_bytes = json.dumps(cfg, ensure_ascii=False, indent=2).encode("utf-8")
    c3.download_button(
        "⬇ 設定JSONをダウンロード",
        data=cfg_bytes,
        file_name="config.json",
        mime="application/json",
        use_container_width=True,
    )


# =========================================================
# Main
# =========================================================
def main() -> None:
    st.set_page_config(page_title="月別 入金・出金 集計アプリ", layout="wide")
    init_session()

    st.title("月別 入金・出金 集計アプリ")
    st.caption(
        "Excelファイル（入金 / 出金）をアップロードして、月ごとの入金・出金・差額を集計します。"
    )

    with st.sidebar:
        st.header("📦 ファイルアップロード")
        st.markdown("**入金Excel（複数可）**")
        income_files = render_upload_area("入金ファイルをドラッグ＆ドロップ", "income_uploader")
        st.markdown("**出金Excel（複数可）**")
        expense_files = render_upload_area("出金ファイルをドラッグ＆ドロップ", "expense_uploader")

        st.divider()
        render_config_section()

    # データ確認モード
    render_preview_section(income_files, expense_files)
    st.divider()

    # 列名の手動指定
    render_column_overrides(income_files, expense_files)
    st.divider()

    # 実行ボタン
    col_run, col_clear = st.columns([1, 1])
    if col_run.button("🚀 集計を実行", type="primary", use_container_width=True):
        if not income_files and not expense_files:
            st.warning("入金または出金ファイルを少なくとも1つアップロードしてください。")
        else:
            with st.spinner("ファイル解析中..."):
                st.session_state.parsed_results = run_parsing(income_files, expense_files)
            st.success("解析が完了しました。")

    if col_clear.button("🗑️ 解析結果をクリア", use_container_width=True):
        st.session_state.parsed_results = []
        st.session_state.restored_keys = set()
        st.rerun()

    results: list[FileResult] = st.session_state.parsed_results
    if not results:
        st.info("ファイルをアップロードして「集計を実行」を押してください。")
        return

    # 期間フィルター
    st.divider()
    st.subheader("📅 期間フィルター")
    all_months: set[str] = set()
    for r in results:
        for row in r.detail_rows + r.excluded_rows:
            if row.get("month"):
                all_months.add(row["month"])
    sorted_months = sorted(all_months)
    if sorted_months:
        c1, c2 = st.columns(2)
        month_from = c1.selectbox(
            "開始月", sorted_months, index=0, format_func=format_month_jp, key="month_from"
        )
        month_to = c2.selectbox(
            "終了月", sorted_months, index=len(sorted_months) - 1, format_func=format_month_jp, key="month_to"
        )
    else:
        month_from = month_to = None
        st.info("月の判定ができたデータがありません。")

    # 集計
    monthly_df, income_df, expense_df, excluded_df, file_summary_df = build_dataframes(
        results, st.session_state.restored_keys, month_from, month_to
    )

    st.divider()
    render_monthly_summary(monthly_df)
    st.divider()
    render_drill_down(monthly_df, income_df, expense_df)
    st.divider()
    render_file_summary(file_summary_df)
    st.divider()
    render_excluded_rows(results)
    st.divider()
    log_lines = render_processing_log(results)
    st.divider()

    # Excel ダウンロード
    st.subheader("⬇ Excelダウンロード")
    try:
        excel_bytes = build_excel(
            monthly_df, income_df, expense_df, excluded_df, file_summary_df,
            log_lines, st.session_state.config,
        )
        fname = f"集計結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            "📥 Excelをダウンロード",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"Excel生成失敗: {e}")
        st.code(traceback.format_exc(), language="text")


if __name__ == "__main__":
    main()
