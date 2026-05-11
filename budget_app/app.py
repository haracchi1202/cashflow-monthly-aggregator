"""資金繰り予算ファイル 上書き更新アプリ (Streamlit)

入金・出金 Excel を月別集計し、予算ファイル (MBJ26年資金繰...xlsx) の
『確定入金 / 確定支払』×『予算列』だけを安全に上書きする。

[安全設計]
    - 元ファイルは触らず、メモリ上のコピーに対して上書き
    - 上書きするのは事前検出した予算列セルのみ
    - 実績列・他シート・他行・数式・書式は変更しない
    - 上書き前確認 → 実行 → 上書き後比較 を必ず通す
"""
from __future__ import annotations

import io
import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from excel_reader import (
    FileResult,
    format_month_jp,
    format_yen,
    parse_file,
    read_excel_preview,
    read_excel_with_header,
)
from aggregator import build_dataframes, monthly_summary_to_dict
from budget_updater import (
    DEFAULT_TARGET_SHEETS,
    analyze_budget_workbook,
    apply_overwrite,
    build_overwrite_plan,
    verify_overwrite,
)
from report_writer import build_report


CONFIG_PATH = Path(__file__).parent / "config.json"

DEFAULT_CONFIG: dict[str, Any] = {
    "default_header_row": 7,
    "month_column_keywords": ["入金日", "支払日", "支払い日", "入金月", "支払月", "月", "対象月", "年月", "月度", "期間"],
    "amount_column_keywords": ["税込", "合計", "原価総額", "入金額", "実績入金", "入金実績", "売上入金", "確定入金", "支払額", "出金額", "確定支払", "支払実績", "金額"],
    "deal_column_keywords": ["商談名", "案件名", "クライアント", "顧客名", "取引先", "件名", "摘要"],
    "exclude_keywords": ["合計", "小計", "総計", "Grand Total", "Total", "計"],
    "budget_target_sheets": list(DEFAULT_TARGET_SHEETS),
    "budget_fiscal_year": 2026,
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
# Session 初期化
# =========================================================
def init_session() -> None:
    ss = st.session_state
    if "config" not in ss:
        ss.config = load_config()
    if "parsed_results" not in ss:
        ss.parsed_results = []  # list[FileResult]
    if "restored_keys" not in ss:
        ss.restored_keys = set()
    if "header_rows" not in ss:
        ss.header_rows = {}
    if "overrides" not in ss:
        ss.overrides = {}
    if "budget_analysis" not in ss:
        ss.budget_analysis = None  # dict
    if "budget_bytes" not in ss:
        ss.budget_bytes = None
    if "budget_filename" not in ss:
        ss.budget_filename = None
    if "updated_bytes" not in ss:
        ss.updated_bytes = None
    if "written_log" not in ss:
        ss.written_log = []
    if "verify_log" not in ss:
        ss.verify_log = []
    if "log_lines" not in ss:
        ss.log_lines = []


# =========================================================
# UI: アップロード
# =========================================================
def render_upload_area(label: str, key: str, accept_multi: bool = True):
    files = st.file_uploader(
        label,
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=accept_multi,
        key=key,
    )
    return files or ([] if accept_multi else None)


# =========================================================
# UI: データ確認モード
# =========================================================
def render_preview_section(income_files: list, expense_files: list) -> None:
    st.subheader("📄 データ確認モード（先頭30行プレビュー）")
    st.caption("ヘッダー行がずれている場合は、ヘッダー行番号を変更して再読み込みできます。")
    all_files = [(f, "入金") for f in income_files] + [(f, "出金") for f in expense_files]
    if not all_files:
        st.info("入金 / 出金 ファイルをアップロードしてください。")
        return

    default_hdr = st.session_state.config["default_header_row"]
    for uf, kind in all_files:
        with st.expander(f"[{kind}] {uf.name}", expanded=False):
            current_hdr = st.session_state.header_rows.get(uf.name, default_hdr)
            new_hdr = st.number_input(
                f"ヘッダー行番号（1始まり） - {uf.name}",
                min_value=1, max_value=50, value=int(current_hdr),
                key=f"hdr_{uf.name}",
            )
            st.session_state.header_rows[uf.name] = int(new_hdr)
            try:
                preview_df = read_excel_preview(uf.getvalue(), n=30)
                preview_df.index = [i + 1 for i in preview_df.index]
                st.dataframe(preview_df, use_container_width=True, height=320)
            except Exception as e:
                st.error(f"プレビュー失敗: {e}")


# =========================================================
# UI: 列名 手動指定
# =========================================================
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

            month_col = st.selectbox("月列", cols, index=_idx(ov.get("month_col")), key=f"m_{uf.name}")
            amount_col = st.selectbox("金額列", cols, index=_idx(ov.get("amount_col")), key=f"a_{uf.name}")
            deal_col = st.selectbox("商談名列", cols, index=_idx(ov.get("deal_col")), key=f"d_{uf.name}")
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


# =========================================================
# 集計実行
# =========================================================
def run_parsing(income_files: list, expense_files: list) -> list[FileResult]:
    cfg = st.session_state.config
    default_hdr = cfg["default_header_row"]
    results: list[FileResult] = []
    progress = st.progress(0.0, text="解析中...")
    total = len(income_files) + len(expense_files)
    done = 0
    for files, kind in ((income_files, "入金"), (expense_files, "出金")):
        for uf in files:
            hdr = st.session_state.header_rows.get(uf.name, default_hdr)
            ov = st.session_state.overrides.get(uf.name, {})
            try:
                r = parse_file(uf.name, uf.getvalue(), kind, hdr, cfg, ov)
            except Exception as e:
                r = FileResult(
                    file_name=uf.name, kind=kind, header_row=hdr,
                    month_col=None, amount_col=None, deal_col=None,
                    error=f"予期せぬ例外: {e}",
                )
                r.log.append(traceback.format_exc())
            results.append(r)
            done += 1
            progress.progress(done / max(total, 1), text=f"解析中... {uf.name}")
    progress.empty()
    return results


# =========================================================
# UI: 月別集計表示
# =========================================================
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
    selected = st.selectbox("月を選択", months, format_func=format_month_jp, key="drill_month")
    if not selected:
        return
    tab_in, tab_out = st.tabs(["入金明細", "出金明細"])
    with tab_in:
        sub = income_df[income_df["month"] == selected] if not income_df.empty else pd.DataFrame()
        _render_detail_table(sub, "入金")
    with tab_out:
        sub = expense_df[expense_df["month"] == selected] if not expense_df.empty else pd.DataFrame()
        _render_detail_table(sub, "出金")


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
        use_container_width=True, hide_index=True,
        disabled=["ファイル名", "行番号", "月", "金額", "除外理由", "商談名"],
        column_config={"復活": st.column_config.CheckboxColumn("復活", help="チェックすると集計対象に戻します")},
        key="excluded_editor",
    )
    new_keys: set[str] = set()
    for i, key in enumerate(df["key"].tolist()):
        if bool(edited.iloc[i]["復活"]):
            new_keys.add(key)
    if new_keys != st.session_state.restored_keys:
        st.session_state.restored_keys = new_keys
        st.rerun()


def render_processing_log(results: list[FileResult], extra_lines: list[str] | None = None) -> list[str]:
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
    if extra_lines:
        lines.append("")
        lines.append("== 予算ファイル ==")
        lines.extend(extra_lines)
    if not lines:
        st.info("ログはまだありません。")
    else:
        st.code("\n".join(lines), language="text")
    return lines


# =========================================================
# UI: 予算ファイル解析
# =========================================================
def render_budget_analysis() -> None:
    st.subheader("📒 予算ファイル解析")
    if not st.session_state.budget_bytes:
        st.info("予算ファイル（資金繰り）をサイドバーからアップロードしてください。")
        return

    cfg = st.session_state.config
    c1, c2 = st.columns([2, 1])
    with c1:
        sheets_str = st.text_input(
            "対象シート（カンマ区切り）",
            value=",".join(cfg.get("budget_target_sheets", DEFAULT_TARGET_SHEETS)),
            key="budget_sheets_input",
        )
    with c2:
        fy = st.number_input(
            "会計年度（半期 / 月のみ表記の年判定に使用）",
            min_value=2000, max_value=2099,
            value=int(cfg.get("budget_fiscal_year", datetime.now().year)),
            key="budget_fy_input",
        )

    target_sheets = tuple(s.strip() for s in sheets_str.split(",") if s.strip())
    cfg["budget_target_sheets"] = list(target_sheets)
    cfg["budget_fiscal_year"] = int(fy)

    if st.button("🔎 予算ファイルを解析", use_container_width=True):
        with st.spinner("予算ファイル解析中..."):
            cell_maps, diagnostics, log_lines = analyze_budget_workbook(
                st.session_state.budget_bytes,
                target_sheets=target_sheets,
                fiscal_year_override=int(fy),
            )
        st.session_state.budget_analysis = {
            "cell_maps": cell_maps,
            "diagnostics": diagnostics,
            "log_lines": log_lines,
        }
        st.session_state.updated_bytes = None
        st.session_state.written_log = []
        st.session_state.verify_log = []

    analysis = st.session_state.budget_analysis
    if not analysis:
        return

    diags = analysis["diagnostics"]
    rows = []
    for d in diags:
        rows.append(
            {
                "シート": d.sheet,
                "状態": "✅ 検出" if d.found else "❌ 失敗",
                "月ヘッダ行": d.month_header_row,
                "サブヘッダ行": d.subheader_row,
                "確定入金行": d.income_row,
                "確定支払行": d.expense_row,
                "検出月数": len(d.months),
                "エラー": d.error or "",
            }
        )
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    cell_maps = analysis["cell_maps"]
    if cell_maps:
        with st.expander(f"検出セル一覧（{len(cell_maps)} 件）", expanded=False):
            tbl = [
                {
                    "シート": cm.sheet,
                    "月": format_month_jp(cm.month),
                    "区分": "確定" + cm.kind,
                    "セル": cm.cell_address,
                    "上書き前値": format_yen(cm.before_value),
                    "元データ": str(cm.before_raw) if cm.before_raw is not None else "",
                }
                for cm in cell_maps
            ]
            st.dataframe(pd.DataFrame(tbl), use_container_width=True, hide_index=True)


# =========================================================
# UI: 更新期間 + プレビュー + 上書き実行
# =========================================================
def render_overwrite_section(monthly_df: pd.DataFrame) -> None:
    st.subheader("✏️ 予算上書き")
    analysis = st.session_state.budget_analysis
    if not analysis or not analysis["cell_maps"]:
        st.warning("先に予算ファイルを解析してください。")
        return

    cell_maps = analysis["cell_maps"]
    budget_months = sorted({cm.month for cm in cell_maps})

    if not budget_months:
        st.warning("予算ファイル内に対象月が見つかりませんでした。")
        return

    c1, c2 = st.columns(2)
    month_from = c1.selectbox(
        "更新開始月", budget_months, index=0, format_func=format_month_jp, key="update_from"
    )
    month_to = c2.selectbox(
        "更新終了月", budget_months, index=len(budget_months) - 1,
        format_func=format_month_jp, key="update_to",
    )

    if month_from > month_to:
        st.error("更新開始月 が 更新終了月 より後になっています。")
        return

    monthly_summary = monthly_summary_to_dict(monthly_df)
    plan_rows, summary_only, budget_only = build_overwrite_plan(
        cell_maps, monthly_summary, month_from, month_to
    )

    target_count = sum(1 for p in plan_rows if p.update_flag)

    st.markdown("#### 上書き前確認（上書き予定一覧）")
    if not plan_rows:
        st.info("更新対象がありません。")
    else:
        cmp_records = [
            {
                "シート": p.sheet,
                "月": format_month_jp(p.month),
                "入金セル": p.income_cell,
                "上書き前 入金予算": format_yen(p.income_before),
                "上書き後 入金予算": format_yen(p.income_after),
                "入金差額": format_yen(p.income_diff),
                "出金セル": p.expense_cell,
                "上書き前 支払予算": format_yen(p.expense_before),
                "上書き後 支払予算": format_yen(p.expense_after),
                "支払差額": format_yen(p.expense_diff),
                "更新対象": "○" if p.update_flag else "×",
                "更新理由": p.reason,
            }
            for p in plan_rows
        ]
        st.dataframe(pd.DataFrame(cmp_records), use_container_width=True, hide_index=True)
        st.caption(
            f"更新対象セル: {target_count * 2} 件 "
            f"(確定入金 × {target_count} + 確定支払 × {target_count})"
        )

    if summary_only or budget_only:
        with st.expander("未照合月の確認", expanded=False):
            if summary_only:
                st.warning(
                    "集計データにあるが予算ファイルにない月: "
                    + ", ".join(format_month_jp(m) for m in summary_only)
                )
            if budget_only:
                st.info(
                    "予算ファイルにあるが集計データにない月（指定期間内）: "
                    + ", ".join(format_month_jp(m) for m in budget_only)
                )

    confirm = st.checkbox(
        "⚠️ 上記の内容で予算ファイルを上書きすることを確認しました。",
        key="overwrite_confirm",
    )
    if st.button("🚀 上書きを実行", type="primary", disabled=not confirm, use_container_width=True):
        if target_count == 0:
            st.warning("更新対象がありません。")
        else:
            with st.spinner("予算ファイル上書き中..."):
                updated_bytes, updated_count, skipped_count, written_log = apply_overwrite(
                    st.session_state.budget_bytes,
                    cell_maps,
                    monthly_summary,
                    month_from,
                    month_to,
                )
                verify_log = verify_overwrite(updated_bytes, cell_maps)
            st.session_state.updated_bytes = updated_bytes
            st.session_state.written_log = written_log
            st.session_state.verify_log = verify_log
            st.session_state.last_month_from = month_from
            st.session_state.last_month_to = month_to
            st.success(
                f"上書き完了: 更新 {updated_count} 件 / スキップ {skipped_count} 件"
            )

    if st.session_state.updated_bytes:
        st.markdown("#### 上書き後比較")
        verify_map = {(v["sheet"], v["month"], v["kind"]): v["after"] for v in st.session_state.verify_log}
        cmp_after = []
        for p in plan_rows:
            cmp_after.append(
                {
                    "シート": p.sheet,
                    "月": format_month_jp(p.month),
                    "入金セル": p.income_cell,
                    "上書き前 入金予算": format_yen(p.income_before),
                    "上書き後 入金予算（実測）": format_yen(verify_map.get((p.sheet, p.month, "入金"), p.income_before)),
                    "出金セル": p.expense_cell,
                    "上書き前 支払予算": format_yen(p.expense_before),
                    "上書き後 支払予算（実測）": format_yen(verify_map.get((p.sheet, p.month, "出金"), p.expense_before)),
                    "更新対象": "○" if p.update_flag else "×",
                    "更新理由": p.reason,
                }
            )
        st.dataframe(pd.DataFrame(cmp_after), use_container_width=True, hide_index=True)


# =========================================================
# UI: 設定セクション
# =========================================================
def render_config_section() -> None:
    st.subheader("⚙️ 設定の保存・読み込み")
    cfg = st.session_state.config
    with st.expander("キーワード設定（自動判定の調整）", expanded=False):
        month_kw = st.text_input("月列キーワード (カンマ区切り)", value=",".join(cfg["month_column_keywords"]))
        amount_kw = st.text_input("金額列キーワード (カンマ区切り)", value=",".join(cfg["amount_column_keywords"]))
        deal_kw = st.text_input("商談名列キーワード (カンマ区切り)", value=",".join(cfg["deal_column_keywords"]))
        excl_kw = st.text_input("除外キーワード (カンマ区切り)", value=",".join(cfg["exclude_keywords"]))
        default_hdr = st.number_input("デフォルトのヘッダー行", min_value=1, max_value=50, value=int(cfg["default_header_row"]))
        cfg["month_column_keywords"] = [k.strip() for k in month_kw.split(",") if k.strip()]
        cfg["amount_column_keywords"] = [k.strip() for k in amount_kw.split(",") if k.strip()]
        cfg["deal_column_keywords"] = [k.strip() for k in deal_kw.split(",") if k.strip()]
        cfg["exclude_keywords"] = [k.strip() for k in excl_kw.split(",") if k.strip()]
        cfg["default_header_row"] = int(default_hdr)
    c1, c2, c3 = st.columns(3)
    if c1.button("💾 設定を保存", use_container_width=True):
        save_config(cfg)
        st.success("config.json に保存しました。")
    if c2.button("🔄 設定を再読み込み", use_container_width=True):
        st.session_state.config = load_config()
        st.success("config.json から再読み込みしました。")
        st.rerun()
    cfg_bytes = json.dumps(cfg, ensure_ascii=False, indent=2).encode("utf-8")
    c3.download_button(
        "⬇ 設定JSONをダウンロード", data=cfg_bytes,
        file_name="config.json", mime="application/json",
        use_container_width=True,
    )


# =========================================================
# Main
# =========================================================
def main() -> None:
    st.set_page_config(page_title="月別 入金・出金 集計 → 予算上書きアプリ", layout="wide")
    init_session()
    st.title("月別 入金・出金 集計 → 資金繰り予算 上書きアプリ")
    st.caption(
        "入金 / 出金 Excel をアップロードして月別集計し、資金繰り予算ファイル（MBJ26年資金繰...）の "
        "『確定入金 / 確定支払』の予算列だけを安全に上書きします。"
    )

    with st.sidebar:
        st.header("📦 ファイルアップロード")
        st.markdown("**① 入金Excel（複数可）**")
        income_files = render_upload_area("入金ファイル", "income_uploader", True)
        st.markdown("**② 出金Excel（複数可）**")
        expense_files = render_upload_area("出金ファイル", "expense_uploader", True)
        st.markdown("**③ 資金繰り予算Excel**")
        budget_file = render_upload_area("予算ファイル（例: MBJ26年資金繰...xlsx）", "budget_uploader", False)
        if budget_file is not None:
            st.session_state.budget_bytes = budget_file.getvalue()
            st.session_state.budget_filename = budget_file.name
        st.divider()
        render_config_section()

    render_preview_section(income_files, expense_files)
    st.divider()
    render_column_overrides(income_files, expense_files)
    st.divider()

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
        st.session_state.budget_analysis = None
        st.session_state.updated_bytes = None
        st.rerun()

    results: list[FileResult] = st.session_state.parsed_results
    if not results:
        st.info("入金 / 出金 ファイルをアップロードして「集計を実行」を押してください。")
        return

    # 期間フィルター（集計表示用 — 予算更新期間とは別）
    st.divider()
    st.subheader("📅 期間フィルター（表示用）")
    all_months: set[str] = set()
    for r in results:
        for row in r.detail_rows + r.excluded_rows:
            if row.get("month"):
                all_months.add(row["month"])
    sorted_months = sorted(all_months)
    if sorted_months:
        c1, c2 = st.columns(2)
        month_from = c1.selectbox(
            "開始月", sorted_months, index=0, format_func=format_month_jp, key="filter_from"
        )
        month_to = c2.selectbox(
            "終了月", sorted_months, index=len(sorted_months) - 1,
            format_func=format_month_jp, key="filter_to",
        )
    else:
        month_from = month_to = None
        st.info("月の判定ができたデータがありません。")

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
    render_budget_analysis()
    st.divider()
    render_overwrite_section(monthly_df)
    st.divider()
    extra_log = (st.session_state.budget_analysis or {}).get("log_lines", []) if st.session_state.budget_analysis else []
    log_lines = render_processing_log(results, extra_log)
    st.divider()

    # ダウンロード
    st.subheader("⬇ ダウンロード")
    c1, c2 = st.columns(2)

    # 更新済み Excel
    if st.session_state.updated_bytes:
        src_name = st.session_state.budget_filename or "budget.xlsx"
        stem = src_name.rsplit(".", 1)[0]
        out_name = f"{stem}_updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        c1.download_button(
            "📥 更新済み予算Excelをダウンロード",
            data=st.session_state.updated_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    else:
        c1.info("予算ファイルの上書きを実行するとここからダウンロードできます。")

    # 比較レポート
    analysis = st.session_state.budget_analysis or {}
    cell_maps = analysis.get("cell_maps", []) if analysis else []
    if not monthly_df.empty:
        # 比較レポート用の plan_rows を最新の更新期間で生成
        m_from = st.session_state.get("update_from") or month_from
        m_to = st.session_state.get("update_to") or month_to
        monthly_summary = monthly_summary_to_dict(monthly_df)
        plan_rows, summary_only, budget_only = build_overwrite_plan(
            cell_maps, monthly_summary, m_from, m_to
        )
        try:
            report_bytes = build_report(
                monthly_df=monthly_df,
                income_df=income_df,
                expense_df=expense_df,
                excluded_df=excluded_df,
                file_summary_df=file_summary_df,
                cell_maps=cell_maps,
                plan_rows=plan_rows,
                summary_only_months=summary_only,
                budget_only_months=budget_only,
                written_log=st.session_state.written_log,
                log_lines=log_lines,
                config=st.session_state.config,
                month_from=m_from,
                month_to=m_to,
            )
            out_name = f"比較レポート_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            c2.download_button(
                "📊 比較レポートExcelをダウンロード",
                data=report_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            c2.error(f"比較レポート生成失敗: {e}")
            st.code(traceback.format_exc(), language="text")
    else:
        c2.info("集計実行後に比較レポートを生成できます。")


if __name__ == "__main__":
    main()
