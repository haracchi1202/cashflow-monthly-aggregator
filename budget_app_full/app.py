"""資金繰り予算更新アプリ（フル機能版）

確定入金 / 確定支払 / 予測入金 / 予測支払 の4種類の Excel をアップロードして
月別集計（確定 / 予測 / 合算）を行い、資金繰り予算ファイルの『確定入金 / 確定支払』
の予算列に対して、選択した更新モード（確定のみ / 予測のみ / 合算）で安全に上書きする。

[安全設計]
    - 原本は触らず、メモリ上のコピーに対して上書き
    - 上書きするのは事前検出した予算列セルのみ
    - 実績列・他シート・他行・数式・書式は変更しない
"""
from __future__ import annotations

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
    parse_transaction_file,
    read_excel_preview,
    read_excel_with_header,
)
from aggregator import (
    SOURCE_GROUPS,
    build_dataframes,
    confirmed_summary_dict,
    forecast_summary_dict,
    update_scope_label,
)
from budget_updater import (
    DEFAULT_TARGET_SHEETS,
    analyze_budget_workbook,
    apply_overwrite,
    build_overwrite_plan,
    verify_overwrite,
)
from report_writer import build_report
from snapshot_manager import (
    build_comparison_report,
    compare_snapshots,
    comparison_to_dataframe,
    comparison_to_summary,
    delete_snapshot,
    list_snapshots,
    load_snapshot,
    save_snapshot,
)
from zoho_fetcher import FetchedFile, ZohoFetcher, fetch_transaction_xlsx, load_zoho_config


CONFIG_PATH = Path(__file__).parent / "config.json"

DEFAULT_CONFIG: dict[str, Any] = {
    "default_header_row": 7,
    "month_column_keywords": [
        "入金日", "支払日", "支払い日", "入金月", "支払月", "月",
        "対象月", "年月", "月度", "期間",
        "初回入金月", "残金入金月",
        "初回支払日", "残金支払日", "予備支払日",
    ],
    "amount_column_keywords": [
        "税込", "合計", "原価総額",
        "入金額", "支払額", "出金額",
        "実績入金", "入金実績", "売上入金", "確定入金",
        "確定支払", "支払実績",
        "初回入金額", "残金入金額", "初回入金",
        "初回支払額", "残金支払額", "予備支払額", "初回支払",
        "金額",
    ],
    "deal_column_keywords": ["商談名", "案件名", "件名", "摘要"],
    "client_column_keywords": ["クライアント", "顧客名", "取引先", "クライアント名", "顧客"],
    "exclude_keywords": ["合計", "小計", "総計", "Grand Total", "Total", "計"],
    "month_keywords_by_group": {
        "予測入金": ["初回入金月", "残金入金月", "入金月", "入金日"],
        "予測支払": ["初回支払日", "残金支払日", "予備支払日", "支払日", "支払月"],
    },
    "amount_keywords_by_group": {
        "予測入金": ["初回入金額", "残金入金額", "入金額", "売上入金"],
        "予測支払": ["初回支払額", "残金支払額", "予備支払額", "支払額", "出金額"],
    },
    "budget_target_sheets": list(DEFAULT_TARGET_SHEETS),
    "budget_fiscal_year": 2026,
    "apply_tax_to_forecast": True,
    "forecast_tax_rate": 0.10,
    "file_overrides": {},
}


# =========================================================
# 設定
# =========================================================
def load_config() -> dict[str, Any]:
    if CONFIG_PATH.exists():
        try:
            with CONFIG_PATH.open("r", encoding="utf-8") as f:
                cfg = json.load(f)
            merged = json.loads(json.dumps(DEFAULT_CONFIG))  # deep copy
            merged.update(cfg)
            # ネスト辞書もマージ
            for k in ("month_keywords_by_group", "amount_keywords_by_group"):
                if k in cfg:
                    merged[k] = {**DEFAULT_CONFIG[k], **cfg[k]}
            return merged
        except Exception:
            return json.loads(json.dumps(DEFAULT_CONFIG))
    return json.loads(json.dumps(DEFAULT_CONFIG))


def save_config(cfg: dict[str, Any]) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# =========================================================
# Session
# =========================================================
def init_session() -> None:
    ss = st.session_state
    if "config" not in ss:
        ss.config = load_config()
    if "parsed_results" not in ss:
        ss.parsed_results = []
    if "restored_keys" not in ss:
        ss.restored_keys = set()
    if "header_rows" not in ss:
        ss.header_rows = {}
    if "overrides" not in ss:
        ss.overrides = {}
    if "budget_analysis" not in ss:
        ss.budget_analysis = None
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
    if "update_scope" not in ss:
        ss.update_scope = "both"
    if "comparison_rows" not in ss:
        ss.comparison_rows = None
    if "comparison_meta" not in ss:
        ss.comparison_meta = None
    if "zoho_fetched" not in ss:
        # Zoho から取得した bytes を保持（再集計時の手動アップロードと同等に扱う）
        ss.zoho_fetched: dict[str, FetchedFile] = {}
    if "zoho_view_names" not in ss:
        ss.zoho_view_names = {
            "confirmed": "confirmed_transactions",
            "forecast": "forecast_transactions",
        }


# =========================================================
# UI: アップロード
# =========================================================
def render_upload(label: str, key: str, multi: bool = True):
    files = st.file_uploader(
        label,
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=multi,
        key=key,
    )
    return files or ([] if multi else None)


# =========================================================
# UI: データ確認
# =========================================================
def render_preview(file_lists: dict[str, list]) -> None:
    st.subheader("📄 データ確認モード（先頭30行プレビュー）")
    st.caption("ヘッダー行がずれている場合はヘッダー行番号を変更して再読み込みできます。")
    any_file = any(len(v) for v in file_lists.values())
    if not any_file:
        st.info("ファイルをアップロードしてください。")
        return
    default_hdr = st.session_state.config["default_header_row"]
    for group, files in file_lists.items():
        for uf in files:
            with st.expander(f"[{group}] {uf.name}", expanded=False):
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
def render_column_overrides(file_lists: dict[str, list]) -> None:
    cfg = st.session_state.config
    default_hdr = cfg["default_header_row"]
    st.subheader("🛠️ 列名の手動指定（任意）")
    st.caption("自動判定が間違った場合のみ指定してください。空欄なら自動判定が使われます。")
    any_file = any(len(v) for v in file_lists.values())
    if not any_file:
        return
    for group, files in file_lists.items():
        for uf in files:
            with st.expander(f"[{group}] {uf.name} の列指定", expanded=False):
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
                client_col = st.selectbox("顧客名列", cols, index=_idx(ov.get("client_col")), key=f"c_{uf.name}")
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
                    "client_col": client_col or None,
                    "extra_exclude_keywords": extra_kw,
                }


# =========================================================
# 集計実行
# =========================================================
def run_parsing(file_lists: dict[str, list]) -> list[FileResult]:
    cfg = st.session_state.config
    default_hdr = cfg["default_header_row"]
    results: list[FileResult] = []
    total = sum(len(v) for v in file_lists.values())
    if total == 0:
        return results
    progress = st.progress(0.0, text="解析中...")
    done = 0
    for group, files in file_lists.items():
        for uf in files:
            hdr = st.session_state.header_rows.get(uf.name, default_hdr)
            ov = st.session_state.overrides.get(uf.name, {})
            try:
                r = parse_file(uf.name, uf.getvalue(), group, hdr, cfg, ov)
            except Exception as e:
                from excel_reader import GROUP_TO_STATUS_TYPE
                status, type_ = GROUP_TO_STATUS_TYPE.get(group, ("unknown", "unknown"))
                r = FileResult(
                    file_name=uf.name,
                    source_group=group,
                    transaction_status=status,
                    transaction_type=type_,
                    header_row=hdr,
                    month_col=None, amount_col=None, deal_col=None, client_col=None,
                    error=f"予期せぬ例外: {e}",
                )
                r.log.append(traceback.format_exc())
            results.append(r)
            done += 1
            progress.progress(done / total, text=f"解析中... {uf.name}")
    progress.empty()
    return results


# =========================================================
# UI: 月別集計
# =========================================================
def _yen_or_blank(v) -> str:
    return format_yen(v) if v is not None else ""


def _format_diff_jp(v: float) -> str:
    """マイナスは ▲ 強調表示で返す。"""
    if v is None:
        return ""
    if v < 0:
        return f"▲ {format_yen(abs(v))}"
    return format_yen(v)


def _render_summary_table(
    monthly_df: pd.DataFrame,
    in_col: str,
    out_col: str,
    diff_col: str,
    labels: tuple[str, str, str],
) -> None:
    """ある軸(確定 / 予測 / 合算)の表+メトリクスをレンダリング。"""
    display = monthly_df[["target_month", in_col, out_col, diff_col]].copy()
    display.insert(0, "月", display["target_month"].apply(format_month_jp))
    display = display.drop(columns=["target_month"])
    display.columns = ["月", labels[0], labels[1], labels[2]]
    display[labels[0]] = display[labels[0]].apply(format_yen)
    display[labels[1]] = display[labels[1]].apply(format_yen)
    display[labels[2]] = display[labels[2]].apply(_format_diff_jp)
    st.dataframe(display, use_container_width=True, hide_index=True)

    total_in = monthly_df[in_col].sum()
    total_out = monthly_df[out_col].sum()
    diff = total_in - total_out
    c1, c2, c3 = st.columns(3)
    c1.metric(labels[0] + "（合計）", format_yen(total_in))
    c2.metric(labels[1] + "（合計）", format_yen(total_out))
    c3.metric(
        labels[2] + "（合計）",
        format_yen(diff),
        delta=("マイナス" if diff < 0 else "プラス"),
    )


def render_monthly_summary(monthly_df: pd.DataFrame) -> None:
    st.subheader("📊 月別集計")
    cfg = st.session_state.config
    if cfg.get("apply_tax_to_forecast", True):
        rate = float(cfg.get("forecast_tax_rate", 0.10)) * 100
        st.caption(
            f"💴 予測入金・予測支払 には消費税 {rate:.1f}% を自動加算済み（確定は税込のまま）"
        )
    if monthly_df.empty:
        st.info("集計対象データがありません。")
        return

    tab_conf, tab_fc, tab_comb, tab_all = st.tabs(
        ["✅ 確定のみ", "🔮 予測のみ", "Σ 確定＋予測（合算）", "📋 全列まとめて"]
    )

    with tab_conf:
        st.caption("確定入金 / 確定支払 / 確定差額")
        _render_summary_table(
            monthly_df,
            in_col="確定入金", out_col="確定支払", diff_col="確定差額",
            labels=("確定入金", "確定支払", "確定差額"),
        )

    with tab_fc:
        st.caption("予測入金 / 予測支払 / 予測差額")
        _render_summary_table(
            monthly_df,
            in_col="予測入金", out_col="予測支払", diff_col="予測差額",
            labels=("予測入金", "予測支払", "予測差額"),
        )

    with tab_comb:
        st.caption("確定＋予測 入金合計 / 確定＋予測 支払合計 / 確定＋予測 差額")
        _render_summary_table(
            monthly_df,
            in_col="合算入金", out_col="合算支払", diff_col="合算差額",
            labels=("確定＋予測 入金合計", "確定＋予測 支払合計", "確定＋予測 差額"),
        )

    with tab_all:
        st.caption("確定 / 予測 / 合算 を1行に並べた一覧")
        display = monthly_df.copy()
        display.insert(0, "月", display["target_month"].apply(format_month_jp))
        display = display.drop(columns=["target_month"])
        for col in display.columns[1:]:
            if col.endswith("差額"):
                display[col] = display[col].apply(_format_diff_jp)
            else:
                display[col] = display[col].apply(format_yen)
        st.dataframe(display, use_container_width=True, hide_index=True)

    # 共通の総合メトリクス
    st.markdown("##### 全期間サマリー")
    total_conf_in = monthly_df["確定入金"].sum()
    total_conf_out = monthly_df["確定支払"].sum()
    total_fc_in = monthly_df["予測入金"].sum()
    total_fc_out = monthly_df["予測支払"].sum()
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("確定入金", format_yen(total_conf_in))
    c2.metric("確定支払", format_yen(total_conf_out))
    c3.metric("予測入金", format_yen(total_fc_in))
    c4.metric("予測支払", format_yen(total_fc_out))
    diff_combined = (total_conf_in + total_fc_in) - (total_conf_out + total_fc_out)
    c5.metric(
        "合算 差額",
        format_yen(diff_combined),
        delta=("マイナス" if diff_combined < 0 else "プラス"),
    )


def render_drilldown(monthly_df: pd.DataFrame, by_source_dfs: dict[str, pd.DataFrame]) -> None:
    st.subheader("🔍 月別詳細（ドリルダウン）")
    if monthly_df.empty:
        st.info("集計対象データがありません。")
        return
    months = monthly_df["target_month"].tolist()
    selected = st.selectbox("月を選択", months, format_func=format_month_jp, key="drill_month")
    if not selected:
        return
    tabs = st.tabs(SOURCE_GROUPS)
    for tab, grp in zip(tabs, SOURCE_GROUPS):
        with tab:
            sub = by_source_dfs.get(grp, pd.DataFrame())
            sub = sub[sub["target_month"] == selected] if not sub.empty else pd.DataFrame()
            _render_detail_table(sub, grp)


def _render_detail_table(df: pd.DataFrame, label: str) -> None:
    if df.empty:
        st.info(f"{label}明細はありません。")
        return
    keep = ["source_file", "raw_row_index", "target_month", "client_name", "deal_name", "amount"]
    cols = [c for c in keep if c in df.columns]
    show = df[cols].copy()
    show.columns = ["ファイル名", "行番号", "月", "顧客名", "商談名", "金額"][: len(cols)]
    if "月" in show.columns:
        show["月"] = show["月"].apply(format_month_jp)
    if "金額" in show.columns:
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


def render_excluded(results: list[FileResult]) -> None:
    st.subheader("🚫 除外行確認 & 手動復活")
    rows: list[dict[str, Any]] = []
    for r in results:
        for row in r.excluded_rows:
            rows.append(row)
    if not rows:
        st.info("除外行はありません。")
        return
    df = pd.DataFrame(rows)
    df["key"] = df.apply(lambda x: f"{x['source_file']}::{x['raw_row_index']}", axis=1)
    df["復活"] = df["key"].apply(lambda k: k in st.session_state.restored_keys)
    show_cols = [
        "復活", "source_file", "raw_row_index", "target_month",
        "amount", "exclude_reason", "deal_name", "source_group",
    ]
    show = df[show_cols].copy()
    show.columns = ["復活", "ファイル名", "行番号", "月", "金額", "除外理由", "商談名", "区分"]
    show["月"] = show["月"].apply(lambda v: format_month_jp(v) if v else "")
    show["金額"] = show["金額"].apply(format_yen)
    edited = st.data_editor(
        show,
        use_container_width=True, hide_index=True,
        disabled=["ファイル名", "行番号", "月", "金額", "除外理由", "商談名", "区分"],
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


def render_logs(results: list[FileResult], extra: list[str] | None = None) -> list[str]:
    st.subheader("📝 処理ログ")
    lines: list[str] = []
    for r in results:
        lines.append(f"[{r.source_group}] {r.file_name}（ヘッダー行: {r.header_row}）")
        lines.append(f"  - 月列: {r.month_col} / 金額列: {r.amount_col} / 商談列: {r.deal_col} / 顧客列: {r.client_col}")
        for log in r.log:
            lines.append(f"  - {log}")
        if r.error:
            lines.append(f"  ❌ エラー: {r.error}")
    if extra:
        lines.append("")
        lines.append("== 予算ファイル ==")
        lines.extend(extra)
    if not lines:
        st.info("ログはまだありません。")
    else:
        st.code("\n".join(lines), language="text")
    return lines


# =========================================================
# UI: 予算解析
# =========================================================
def render_budget_analysis() -> None:
    st.subheader("📒 予算ファイル解析")
    if not st.session_state.budget_bytes:
        st.info("予算ファイルをサイドバーからアップロードしてください。")
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
                "確定入金行": d.confirmed_income_row,
                "確定支払行": d.confirmed_expense_row,
                "予測入金行": d.forecast_income_row if d.forecast_income_row is not None else "—",
                "予測支払行": d.forecast_expense_row if d.forecast_expense_row is not None else "—",
                "検出月数": len(d.months),
                "備考": "; ".join(d.notes) if d.notes else "",
                "エラー": d.error or "",
            }
        )
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # 予測行が見つからなかったシートに対して警告
    missing_forecast = [
        d.sheet for d in diags
        if d.found and (d.forecast_income_row is None or d.forecast_expense_row is None)
    ]
    if missing_forecast:
        st.warning(
            "次のシートで予測入金/予測支払 行が見つかりませんでした: "
            + ", ".join(missing_forecast)
            + "\n→ 予測値の書き込みはスキップされます。"
        )

    cell_maps = analysis["cell_maps"]
    if cell_maps:
        with st.expander(f"検出セル一覧（{len(cell_maps)} 件）", expanded=False):
            tbl = [
                {
                    "シート": cm.sheet,
                    "月": format_month_jp(cm.month),
                    "ステータス": "確定" if cm.transaction_status == "confirmed" else "予測",
                    "区分": cm.kind,
                    "セル": cm.cell_address,
                    "上書き前値": format_yen(cm.before_value),
                    "元データ": str(cm.before_raw) if cm.before_raw is not None else "",
                }
                for cm in cell_maps
            ]
            st.dataframe(pd.DataFrame(tbl), use_container_width=True, hide_index=True)


# =========================================================
# UI: 上書きセクション
# =========================================================
def _plan_to_display_records(plans) -> list[dict[str, Any]]:
    """確定/予測 を 1 行にまとめた表示用レコード（月×シート 単位）。"""
    by_key: dict[tuple[str, str], dict[str, Any]] = {}
    for p in plans:
        key = (p.sheet, p.month)
        if key not in by_key:
            by_key[key] = {
                "シート": p.sheet,
                "月": format_month_jp(p.month),
                "_month": p.month,
                "確定入金セル": "", "確定入金 前": None, "確定入金 後": None, "確定入金 差額": None,
                "確定支払セル": "", "確定支払 前": None, "確定支払 後": None, "確定支払 差額": None,
                "予測入金セル": "", "予測入金 前": None, "予測入金 後": None, "予測入金 差額": None,
                "予測支払セル": "", "予測支払 前": None, "予測支払 後": None, "予測支払 差額": None,
                "更新対象": "—",
                "更新理由": "",
            }
        rec = by_key[key]
        col_prefix = ("確定" if p.transaction_status == "confirmed" else "予測") + p.kind
        rec[col_prefix + "セル"] = p.cell_address
        rec[col_prefix + " 前"] = p.before
        rec[col_prefix + " 後"] = p.after
        rec[col_prefix + " 差額"] = p.diff
        # 行全体の更新状態（どれかが更新対象なら○）
        if p.update_flag:
            rec["更新対象"] = "○"
        if not rec["更新理由"] or p.update_flag:
            rec["更新理由"] = p.reason

    out: list[dict[str, Any]] = []
    for rec in sorted(by_key.values(), key=lambda r: (r["シート"], r["_month"])):
        rec_display = {k: (format_yen(v) if k.endswith(" 前") or k.endswith(" 後") or k.endswith(" 差額") else v) for k, v in rec.items() if k != "_month"}
        out.append(rec_display)
    return out


def render_overwrite(monthly_df: pd.DataFrame) -> None:
    st.subheader("✏️ 予算上書き")
    analysis = st.session_state.budget_analysis
    if not analysis or not analysis["cell_maps"]:
        st.warning("先に予算ファイルを解析してください。")
        return
    if monthly_df.empty:
        st.info("月別集計データがありません。先に「集計を実行」してください。")
        return

    st.caption(
        "**確定値は確定行へ / 予測値は予測行へ書き込みます（合算しません）**"
    )
    update_scope = st.radio(
        "更新スコープ",
        options=["both", "confirmed_only", "forecast_only"],
        format_func=update_scope_label,
        index=["both", "confirmed_only", "forecast_only"].index(st.session_state.update_scope),
        horizontal=True,
        key="update_scope_radio",
    )
    st.session_state.update_scope = update_scope

    cell_maps = analysis["cell_maps"]
    budget_months = sorted({cm.month for cm in cell_maps})
    if not budget_months:
        st.warning("予算ファイル内に対象月が見つかりませんでした。")
        return

    # 予測行が無いシートを警告
    has_forecast_row = any(cm.transaction_status == "forecast" for cm in cell_maps)
    if update_scope != "confirmed_only" and not has_forecast_row:
        st.warning(
            "予算ファイルに予測入金/予測支払 行が見つかりません。"
            "予測値は書き込めません（自動でスキップされます）。"
        )

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

    confirmed_sum = confirmed_summary_dict(monthly_df)
    forecast_sum = forecast_summary_dict(monthly_df)
    plans, summary_only, budget_only = build_overwrite_plan(
        cell_maps,
        confirmed_summary=confirmed_sum,
        forecast_summary=forecast_sum,
        month_from=month_from,
        month_to=month_to,
        update_scope=update_scope,
    )
    target_count = sum(1 for p in plans if p.update_flag)
    confirmed_target = sum(1 for p in plans if p.update_flag and p.transaction_status == "confirmed")
    forecast_target = sum(1 for p in plans if p.update_flag and p.transaction_status == "forecast")

    st.markdown(f"#### 上書き前確認 — スコープ: **{update_scope_label(update_scope)}**")
    st.dataframe(
        pd.DataFrame(_plan_to_display_records(plans)),
        use_container_width=True, hide_index=True,
    )
    st.caption(
        f"更新対象セル合計: {target_count} 件 "
        f"（確定 {confirmed_target} 件 + 予測 {forecast_target} 件）"
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
        f"⚠️ 「{update_scope_label(update_scope)}」で予算ファイルを上書きすることを確認しました。",
        key="overwrite_confirm",
    )
    if st.button("🚀 上書きを実行", type="primary", disabled=not confirm, use_container_width=True):
        if target_count == 0:
            st.warning("更新対象がありません。")
        else:
            with st.spinner("予算ファイル上書き中..."):
                updated_bytes, updated_count, skipped_count, written_log = apply_overwrite(
                    st.session_state.budget_bytes,
                    plans,
                    cell_maps,
                )
                verify_log = verify_overwrite(updated_bytes, cell_maps)
            st.session_state.updated_bytes = updated_bytes
            st.session_state.written_log = written_log
            st.session_state.verify_log = verify_log
            st.success(f"上書き完了: 更新 {updated_count} 件 / スキップ {skipped_count} 件")

    if st.session_state.updated_bytes:
        st.markdown("#### 上書き後比較（実測値）")
        verify_map = {
            (v["sheet"], v["month"], v["transaction_status"], v["kind"]): v["after"]
            for v in st.session_state.verify_log
        }
        # plans からセル単位で実測値を結合
        rows = []
        for p in plans:
            actual = verify_map.get((p.sheet, p.month, p.transaction_status, p.kind), p.before)
            rows.append(
                {
                    "シート": p.sheet,
                    "月": format_month_jp(p.month),
                    "ステータス": "確定" if p.transaction_status == "confirmed" else "予測",
                    "区分": p.kind,
                    "セル": p.cell_address,
                    "上書き前": format_yen(p.before),
                    "上書き後（実測）": format_yen(actual),
                    "更新対象": "○" if p.update_flag else "×",
                    "更新理由": p.reason,
                }
            )
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# =========================================================
# UI: スナップショット保存 / 比較
# =========================================================
def render_snapshots(
    monthly_df: pd.DataFrame,
    results: list[FileResult],
) -> None:
    st.subheader("💾 スナップショット保存 & 前回との比較")
    st.caption(
        "現在の集計結果を保存しておき、次回 Excel をアップロードした時に "
        "「金額がいくらに変わったか」「入金日が動いたか」などを商談名ベースで比較できます。"
    )

    # ---- 保存 ----
    with st.expander("✏️ 現在の集計を保存", expanded=False):
        if monthly_df.empty:
            st.info("集計データがありません。先に「🚀 集計を実行」してください。")
        else:
            default_name = f"snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            snap_name = st.text_input("スナップショット名", value=default_name, key="snap_name_input")
            if st.button("💾 保存", use_container_width=True, key="snap_save_btn"):
                path = save_snapshot(
                    name=snap_name,
                    results=results,
                    monthly_df=monthly_df,
                    written_log=st.session_state.written_log,
                    update_scope=st.session_state.update_scope,
                    month_from=st.session_state.get("update_from"),
                    month_to=st.session_state.get("update_to"),
                )
                st.success(f"保存しました: {path.name}")

    # ---- 一覧 ----
    snaps = list_snapshots()
    if not snaps:
        st.info("まだスナップショットは保存されていません。")
        return

    with st.expander(f"📚 保存済みスナップショット一覧（{len(snaps)} 件）", expanded=False):
        snap_table = [
            {
                "ファイル名": s.path.name,
                "名前": s.name,
                "保存日時": s.saved_at,
                "明細件数": s.detail_count,
                "月別行数": s.monthly_count,
            }
            for s in snaps
        ]
        st.dataframe(pd.DataFrame(snap_table), use_container_width=True, hide_index=True)

        # 削除
        del_target = st.selectbox(
            "削除するスナップショット",
            options=[""] + [s.path.name for s in snaps],
            key="snap_delete_select",
        )
        if del_target and st.button("🗑️ 選択したスナップショットを削除", key="snap_delete_btn"):
            for s in snaps:
                if s.path.name == del_target:
                    delete_snapshot(s.path)
                    st.success(f"削除しました: {del_target}")
                    st.rerun()
                    break

    # ---- 比較 ----
    st.markdown("#### 🔄 前回スナップショットと比較")
    options = [(s.path.name, f"{s.saved_at}  /  {s.name}") for s in snaps]
    if len(options) < 1:
        return

    c1, c2 = st.columns(2)
    base_choice = c1.selectbox(
        "前回（比較元）",
        options=[o[0] for o in options],
        format_func=lambda v: dict(options)[v],
        index=min(1, len(options) - 1),
        key="snap_base_select",
    )
    new_choice = c2.selectbox(
        "今回（比較先 / 既定: 最新）",
        options=["__current__"] + [o[0] for o in options],
        format_func=lambda v: "現在の集計結果（未保存）" if v == "__current__" else dict(options)[v],
        index=0,
        key="snap_new_select",
    )

    if st.button("🔍 比較を実行", type="primary", use_container_width=True, key="snap_compare_btn"):
        try:
            base_data = load_snapshot(next(s.path for s in snaps if s.path.name == base_choice))
        except StopIteration:
            st.error("前回スナップショットの読み込みに失敗しました。")
            return

        if new_choice == "__current__":
            if monthly_df.empty:
                st.warning("現在の集計データがありません。集計を実行してください。")
                return
            # 現在の状態をその場で擬似的にスナップショット化
            details = []
            excluded = []
            for r in results:
                for row in r.detail_rows:
                    details.append({
                        k: row.get(k) for k in [
                            "source_file", "source_group", "transaction_status", "transaction_type",
                            "target_month", "transaction_date", "amount",
                            "client_name", "deal_name", "raw_row_index",
                        ]
                    })
            new_data = {
                "name": "現在（未保存）",
                "saved_at": datetime.now().isoformat(timespec="seconds"),
                "details": details,
            }
        else:
            new_data = load_snapshot(next(s.path for s in snaps if s.path.name == new_choice))

        rows = compare_snapshots(base_data, new_data)
        st.session_state.comparison_rows = rows
        st.session_state.comparison_meta = {"old": base_data, "new": new_data}

    rows = st.session_state.comparison_rows
    meta = st.session_state.comparison_meta
    if not rows or not meta:
        return

    # ---- 比較サマリー ----
    summary = comparison_to_summary(rows)
    st.markdown("##### 📊 比較サマリー")
    sum_table = [
        {
            "ファイル区分": sg,
            "追加": counts["added"],
            "削除": counts["removed"],
            "変更": counts["changed"],
            "変更なし": counts["unchanged"],
            "合計": sum(counts.values()),
        }
        for sg, counts in sorted(summary.items())
    ]
    st.dataframe(pd.DataFrame(sum_table), use_container_width=True, hide_index=True)

    # ---- フィルタ + 表 ----
    st.markdown("##### 🧾 変更明細")
    status_filter = st.multiselect(
        "表示する状態",
        options=["変更", "追加", "削除", "変更なし"],
        default=["変更", "追加", "削除"],
        key="snap_status_filter",
    )
    group_filter = st.multiselect(
        "ファイル区分",
        options=sorted({r.source_group for r in rows}),
        default=sorted({r.source_group for r in rows}),
        key="snap_group_filter",
    )
    df = comparison_to_dataframe(rows)
    if not df.empty:
        df = df[df["状態"].isin(status_filter)]
        df = df[df["ファイル区分"].isin(group_filter)]
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption(f"表示中: {len(df)} 行 / 全 {len(rows)} 行")

    # ---- レポート Excel ダウンロード ----
    try:
        report_bytes = build_comparison_report(rows, meta["old"], meta["new"])
        out_name = f"変更レポート_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            "📥 変更レポート Excel をダウンロード",
            data=report_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"変更レポート生成失敗: {e}")
        st.code(traceback.format_exc(), language="text")


# =========================================================
# UI: 設定
# =========================================================
def render_config() -> None:
    st.subheader("⚙️ 設定の保存・読み込み")
    cfg = st.session_state.config

    with st.expander("💴 予測ファイル 消費税設定", expanded=False):
        st.caption(
            "予測入金 / 予測支払 ファイルは税抜のため、読み込み時に自動で消費税を加算します。"
            "確定ファイルは税込のためそのまま使われます。"
        )
        apply_tax = st.checkbox(
            "予測ファイルに消費税を自動加算する",
            value=bool(cfg.get("apply_tax_to_forecast", True)),
            key="apply_tax_to_forecast_input",
        )
        tax_rate_pct = st.number_input(
            "税率 (%)",
            min_value=0.0, max_value=30.0, step=0.5,
            value=float(cfg.get("forecast_tax_rate", 0.10)) * 100,
            key="forecast_tax_rate_input",
        )
        cfg["apply_tax_to_forecast"] = bool(apply_tax)
        cfg["forecast_tax_rate"] = round(tax_rate_pct / 100, 4)

    with st.expander("キーワード設定（自動判定の調整）", expanded=False):
        month_kw = st.text_input("月列キーワード (カンマ区切り)", value=",".join(cfg["month_column_keywords"]))
        amount_kw = st.text_input("金額列キーワード (カンマ区切り)", value=",".join(cfg["amount_column_keywords"]))
        deal_kw = st.text_input("商談名列キーワード (カンマ区切り)", value=",".join(cfg["deal_column_keywords"]))
        client_kw = st.text_input("顧客名列キーワード (カンマ区切り)", value=",".join(cfg["client_column_keywords"]))
        excl_kw = st.text_input("除外キーワード (カンマ区切り)", value=",".join(cfg["exclude_keywords"]))
        default_hdr = st.number_input("デフォルトのヘッダー行", min_value=1, max_value=50, value=int(cfg["default_header_row"]))
        cfg["month_column_keywords"] = [k.strip() for k in month_kw.split(",") if k.strip()]
        cfg["amount_column_keywords"] = [k.strip() for k in amount_kw.split(",") if k.strip()]
        cfg["deal_column_keywords"] = [k.strip() for k in deal_kw.split(",") if k.strip()]
        cfg["client_column_keywords"] = [k.strip() for k in client_kw.split(",") if k.strip()]
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
    st.set_page_config(
        page_title="資金繰り予算更新アプリ（フル機能版）",
        layout="wide",
    )
    init_session()
    st.title("資金繰り予算更新アプリ（確定 + 予測 フル機能版）")
    st.caption(
        "確定入金 / 確定支払 / 予測入金 / 予測支払 の Excel をアップロードして月別に集計し、"
        "選択した更新モードで資金繰り予算ファイルの『確定入金 / 確定支払』の予算列だけを安全に上書きします。"
    )

    with st.sidebar:
        st.header("📦 データソース")

        # ---- Zoho から自動取得 ----
        with st.expander("⚡ Zoho Analytics から自動取得", expanded=True):
            zcfg, src = load_zoho_config()
            if zcfg is None:
                st.warning(f"⚠ Zoho 認証情報が未設定です: {src}")
                st.caption(
                    ".env または .streamlit/secrets.toml に "
                    "ZOHO_REGION / CLIENT_ID / CLIENT_SECRET / REFRESH_TOKEN / "
                    "ORG_ID / WORKSPACE_ID を設定してください。"
                )
            else:
                st.caption(f"✅ 認証情報を {src} から読み込みました")
                view_conf = st.text_input(
                    "確定 view 名",
                    value=st.session_state.zoho_view_names["confirmed"],
                    key="zoho_conf_view_input",
                )
                view_fcst = st.text_input(
                    "予測 view 名",
                    value=st.session_state.zoho_view_names["forecast"],
                    key="zoho_fcst_view_input",
                )
                st.session_state.zoho_view_names = {
                    "confirmed": view_conf.strip(),
                    "forecast": view_fcst.strip(),
                }
                if st.button("🔄 Zoho から取得", use_container_width=True, key="zoho_fetch_btn"):
                    try:
                        fetcher = ZohoFetcher(zcfg)
                        with st.spinner("確定 transactions を取得中..."):
                            ff_c = fetch_transaction_xlsx(fetcher, view_conf)
                        with st.spinner("予測 transactions を取得中..."):
                            ff_f = fetch_transaction_xlsx(fetcher, view_fcst)
                        st.session_state.zoho_fetched = {
                            "confirmed": ff_c,
                            "forecast": ff_f,
                        }
                        st.success(
                            f"取得完了: 確定 {ff_c.row_count} 行 / 予測 {ff_f.row_count} 行"
                        )
                    except Exception as e:
                        st.error(f"Zoho 取得失敗: {e}")
                        st.code(traceback.format_exc(), language="text")
            # 取得結果サマリ
            if st.session_state.zoho_fetched:
                for label, ff in st.session_state.zoho_fetched.items():
                    jp = "確定" if label == "confirmed" else "予測"
                    st.caption(f"  {jp}: {ff.row_count:,} 行 / viewId={ff.view_id}")

        st.divider()
        st.markdown("**手動アップロード（Zoho 取得を上書き）**")
        st.markdown("**① 確定 transactions（confirmed_transactions.xlsx）**")
        confirmed_file = render_upload(
            "確定 transactions ファイル", "uploader_confirmed", False
        )
        st.markdown("**② 予測 transactions（forecast_transactions.xlsx）**")
        forecast_file = render_upload(
            "予測 transactions ファイル", "uploader_forecast", False
        )
        st.markdown("**③ 資金繰り予算Excel**")
        budget_file = render_upload(
            "予算ファイル（例: MBJ26年資金繰...xlsx）", "uploader_budget", False
        )
        if budget_file is not None:
            st.session_state.budget_bytes = budget_file.getvalue()
            st.session_state.budget_filename = budget_file.name
        st.divider()
        render_config()

    # 旧 API との互換のため空 list の dict
    file_lists: dict[str, list] = {"確定入金": [], "確定支払": [], "予測入金": [], "予測支払": []}

    # 入力ソースを統一: (filename, bytes, label, source) のリスト
    transaction_sources: list[tuple[str, bytes, str, str]] = []

    # Zoho 取得結果（手動アップロードで上書きされる）
    zfetched = st.session_state.zoho_fetched
    if confirmed_file is not None:
        transaction_sources.append((confirmed_file.name, confirmed_file.getvalue(), "確定", "アップロード"))
    elif "confirmed" in zfetched:
        ff = zfetched["confirmed"]
        transaction_sources.append((ff.filename, ff.xlsx_bytes, "確定", f"Zoho ({ff.view_id})"))
    if forecast_file is not None:
        transaction_sources.append((forecast_file.name, forecast_file.getvalue(), "予測", "アップロード"))
    elif "forecast" in zfetched:
        ff = zfetched["forecast"]
        transaction_sources.append((ff.filename, ff.xlsx_bytes, "予測", f"Zoho ({ff.view_id})"))

    # 簡易プレビュー（任意）
    if transaction_sources:
        with st.expander("📄 データソース確認（先頭30行プレビュー）", expanded=False):
            for fname, bts, label, src in transaction_sources:
                st.markdown(f"**[{label}] {fname}** — ソース: {src}")
                try:
                    preview_df = read_excel_preview(bts, n=30)
                    preview_df.index = [i + 1 for i in preview_df.index]
                    st.dataframe(preview_df, use_container_width=True, height=260)
                except Exception as e:
                    st.error(f"プレビュー失敗: {e}")
    st.divider()

    col_run, col_clear = st.columns([1, 1])
    if col_run.button("🚀 集計を実行", type="primary", use_container_width=True):
        if not transaction_sources:
            st.warning(
                "確定または予測 transactions が必要です。"
                "サイドバーの「Zoho から取得」または「手動アップロード」を使ってください。"
            )
        else:
            with st.spinner("ファイル解析中..."):
                results: list[FileResult] = []
                progress = st.progress(0.0, text="解析中...")
                for i, (fname, bts, label, src) in enumerate(transaction_sources, 1):
                    try:
                        results.extend(
                            parse_transaction_file(fname, bts, st.session_state.config)
                        )
                    except Exception as e:
                        results.append(FileResult(
                            file_name=fname, source_group="(不明)",
                            transaction_status="unknown", transaction_type="unknown",
                            header_row=1,
                            month_col=None, amount_col=None, deal_col=None, client_col=None,
                            error=f"予期せぬ例外: {e}",
                        ))
                    progress.progress(i / len(transaction_sources), text=f"解析中... {fname}")
                progress.empty()
                st.session_state.parsed_results = results
            st.success("解析が完了しました。")
    if col_clear.button("🗑️ 解析結果をクリア", use_container_width=True):
        st.session_state.parsed_results = []
        st.session_state.restored_keys = set()
        st.session_state.budget_analysis = None
        st.session_state.updated_bytes = None
        st.rerun()

    results: list[FileResult] = st.session_state.parsed_results
    if not results:
        st.info("ファイルをアップロードして「集計を実行」を押してください。")
        return

    # 期間フィルター（表示用）
    st.divider()
    st.subheader("📅 期間フィルター（表示用）")
    all_months: set[str] = set()
    for r in results:
        for row in r.detail_rows + r.excluded_rows:
            if row.get("target_month"):
                all_months.add(row["target_month"])
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

    monthly_df, by_source_dfs, excluded_df, file_summary_df = build_dataframes(
        results, st.session_state.restored_keys, month_from, month_to
    )

    st.divider()
    render_monthly_summary(monthly_df)
    st.divider()
    render_drilldown(monthly_df, by_source_dfs)
    st.divider()
    render_file_summary(file_summary_df)
    st.divider()
    render_excluded(results)
    st.divider()
    render_budget_analysis()
    st.divider()
    render_overwrite(monthly_df)
    st.divider()
    render_snapshots(monthly_df, results)
    st.divider()
    extra_log = (st.session_state.budget_analysis or {}).get("log_lines", []) if st.session_state.budget_analysis else []
    log_lines = render_logs(results, extra_log)
    st.divider()

    # ダウンロード
    st.subheader("⬇ ダウンロード")
    c1, c2 = st.columns(2)
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

    analysis = st.session_state.budget_analysis or {}
    cell_maps = analysis.get("cell_maps", []) if analysis else []
    if not monthly_df.empty:
        m_from = st.session_state.get("update_from") or month_from
        m_to = st.session_state.get("update_to") or month_to
        update_scope = st.session_state.update_scope
        confirmed_sum = confirmed_summary_dict(monthly_df)
        forecast_sum = forecast_summary_dict(monthly_df)
        plans, summary_only, budget_only = build_overwrite_plan(
            cell_maps,
            confirmed_summary=confirmed_sum,
            forecast_summary=forecast_sum,
            month_from=m_from,
            month_to=m_to,
            update_scope=update_scope,
        )
        try:
            report_bytes = build_report(
                monthly_df=monthly_df,
                by_source_dfs=by_source_dfs,
                excluded_df=excluded_df,
                file_summary_df=file_summary_df,
                cell_maps=cell_maps,
                plans=plans,
                summary_only_months=summary_only,
                budget_only_months=budget_only,
                written_log=st.session_state.written_log,
                log_lines=log_lines,
                config=st.session_state.config,
                month_from=m_from,
                month_to=m_to,
                update_scope=update_scope,
            )
            out_name = f"比較レポート_{update_scope}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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
