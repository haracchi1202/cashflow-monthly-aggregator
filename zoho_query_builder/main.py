"""CLI エントリ。

使い方:

    # 1. .env を編集して認証情報を入れる
    # 2. ワークスペース内の views 一覧を確認
    python main.py list

    # 3. 元テーブルを指定して Query Table を作成（dry-run で SQL を確認できる）
    python main.py build --source-table "確定" \
                        --output-name "confirmed_transactions" \
                        --status confirmed \
                        --dry-run

    # 4. 確認できたら --no-dry-run で実行
    python main.py build --source-table "確定" \
                        --output-name "confirmed_transactions" \
                        --status confirmed

    # 5. 予測用も同様
    python main.py build --source-table "予測" \
                        --output-name "forecast_transactions" \
                        --status forecast
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path

from detector import detect_columns
from sql_generator import build_transaction_sql
from zoho_client import ZohoAnalyticsClient, ZohoConfig


THIS_DIR = Path(__file__).parent


def load_env() -> dict[str, str]:
    """.env を読み込む。"""
    env: dict[str, str] = dict(os.environ)
    env_path = THIS_DIR / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def make_client() -> ZohoAnalyticsClient:
    env = load_env()
    required = ["ZOHO_REGION", "ZOHO_CLIENT_ID", "ZOHO_CLIENT_SECRET",
                "ZOHO_REFRESH_TOKEN", "ZOHO_ORG_ID", "ZOHO_WORKSPACE_ID"]
    missing = [k for k in required if not env.get(k)]
    if missing:
        print(f"❌ .env で次の値が未設定です: {', '.join(missing)}", file=sys.stderr)
        sys.exit(1)
    cfg = ZohoConfig(
        region=env["ZOHO_REGION"],
        client_id=env["ZOHO_CLIENT_ID"],
        client_secret=env["ZOHO_CLIENT_SECRET"],
        refresh_token=env["ZOHO_REFRESH_TOKEN"],
        org_id=env["ZOHO_ORG_ID"],
        workspace_id=env["ZOHO_WORKSPACE_ID"],
    )
    return ZohoAnalyticsClient(cfg)


# ----- list -----
def cmd_list(args: argparse.Namespace) -> None:
    client = make_client()
    views = client.list_views()
    print(f"=== Workspace 内 view 一覧（{len(views)} 件）===")
    for v in views:
        name = v.get("viewName") or v.get("viewname") or ""
        vid = v.get("viewId") or v.get("viewid") or ""
        vtype = v.get("viewType") or v.get("viewtype") or ""
        if args.filter:
            if args.filter not in name:
                continue
        print(f"  [{vtype:12s}] {vid:>15s}  {name}")


# ----- inspect -----
def cmd_inspect(args: argparse.Namespace) -> None:
    """指定 view の列一覧と自動判定結果を表示。"""
    client = make_client()
    view_id = _resolve_view_id(client, args.source_table)
    cols = client.get_table_columns(view_id)
    print(f"=== {args.source_table} の列一覧（{len(cols)} 列）===")
    for c in cols:
        print(f"  - {c}")
    detection = detect_columns(cols)
    print()
    print("=== 自動判定結果 ===")
    print(f"  商談名 列      : {detection.deal_column}")
    print(f"  クライアント名 : {detection.client_column}")
    print(f"  入金 ペア      : {detection.income_count} 件")
    print(f"  支払 ペア      : {detection.payment_count} 件")
    print()
    for p in detection.pairs:
        print(f"  - {p.payment_round}: 日付={p.date_column} / 金額={p.amount_column}")
    if detection.warnings:
        print()
        print("⚠ 警告:")
        for w in detection.warnings:
            print(f"  - {w}")


# ----- build -----
def cmd_build(args: argparse.Namespace) -> None:
    client = make_client()
    view_id = _resolve_view_id(client, args.source_table)
    cols = client.get_table_columns(view_id)
    detection = detect_columns(cols)

    # --status に応じて pair をフィルタ
    # confirmed → forecast_* を除外、forecast → forecast_* のみ
    if args.status == "forecast":
        filtered = [p for p in detection.pairs if p.payment_round.startswith("forecast_")]
    else:  # confirmed
        filtered = [p for p in detection.pairs if not p.payment_round.startswith("forecast_")]
    detection.pairs = filtered
    detection.income_count = sum(1 for p in filtered if p.transaction_type == "income")
    detection.payment_count = sum(1 for p in filtered if p.transaction_type == "payment")

    if not detection.pairs:
        print(f"❌ '{args.status}' 対象の入金/支払 ペアが1つも検出できませんでした。", file=sys.stderr)
        for w in detection.warnings:
            print(f"  - {w}", file=sys.stderr)
        sys.exit(2)

    sql = build_transaction_sql(
        args.source_table, detection, args.status,
        stage_column=args.stage_column,
        stage_values=args.stage_values,
        qualify_column=args.qualify_column,
        qualify_date_from=args.qualify_from,
        qualify_date_to=args.qualify_to,
    )
    print("=== 生成 SQL ===")
    print(sql)
    print()
    print(f"=== サマリー ===")
    print(f"  入金 SELECT 数: {detection.income_count}")
    print(f"  支払 SELECT 数: {detection.payment_count}")
    print(f"  合計 UNION 数 : {detection.income_count + detection.payment_count}")
    print()
    if detection.warnings:
        print("⚠ 警告:")
        for w in detection.warnings:
            print(f"  - {w}")
        print()

    if args.dry_run:
        print("(--dry-run 指定のため、Query Table の作成はスキップしました)")
        return

    print(f"=== Query Table 作成中: {args.output_name} ===")

    # 同名 Query Table が既存なら、--replace 指定時のみ削除して作り直す
    if args.replace:
        try:
            existing = client.list_views()
            for v in existing:
                if (v.get("viewName") == args.output_name
                        and (v.get("viewType") or "").lower() in ("querytable", "query table")):
                    vid = str(v.get("viewId"))
                    print(f"  既存 Query Table を削除: viewId={vid}")
                    client.delete_view(vid)
                    break
        except Exception as e:
            print(f"  ⚠ 既存削除の試行に失敗: {e}")

    # description は API バージョンで対応していないことがあるためデフォルト空
    result = client.create_query_table(args.output_name, sql, description="")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\n✅ 完了: Query Table '{args.output_name}' が作成されました。")


def _resolve_view_id(client: ZohoAnalyticsClient, name: str) -> str:
    views = client.list_views()
    # name 完全一致 → 部分一致 の順
    matches = [v for v in views if (v.get("viewName") == name)]
    if not matches:
        matches = [v for v in views if name in (v.get("viewName") or "")]
    if not matches:
        raise SystemExit(
            f"❌ '{name}' に該当する view が見つかりません。`python main.py list` で一覧確認してください。"
        )
    if len(matches) > 1:
        print(f"⚠ '{name}' に複数の view がマッチしました:")
        for v in matches:
            print(f"  - viewId={v.get('viewId')}, viewName={v.get('viewName')}, viewType={v.get('viewType')}")
        # Table タイプを優先
        tables = [v for v in matches if (v.get("viewType") or "").upper() == "TABLE"]
        chosen = tables[0] if tables else matches[0]
    else:
        chosen = matches[0]
    return str(chosen.get("viewId"))


def _extract_columns(meta: dict) -> list[str]:
    """API 戻り値から列名のみを抜き出す。"""
    # Zoho の v2 API は views → columns の形を返す
    cols: list[str] = []
    candidates = [
        meta.get("columns"),
        meta.get("views", {}).get("columns") if isinstance(meta.get("views"), dict) else None,
    ]
    for c in candidates:
        if not c:
            continue
        for col in c:
            if isinstance(col, dict):
                name = col.get("columnName") or col.get("columnname") or col.get("name")
                if name:
                    cols.append(str(name))
            elif isinstance(col, str):
                cols.append(col)
        if cols:
            break
    return cols


# ----- workspaces -----
def cmd_workspaces(args: argparse.Namespace) -> None:
    """ワークスペース ID を調べる補助コマンド（org_id / workspace_id がわからない時）。"""
    client = make_client()
    workspaces = client.list_workspaces()
    print(f"=== 所有ワークスペース ({len(workspaces)}) ===")
    for w in workspaces:
        print(f"  workspaceId={w.get('workspaceId')} / orgId={w.get('orgId')} / name={w.get('workspaceName')}")


# ----- main -----
def main() -> None:
    parser = argparse.ArgumentParser(description="Zoho Analytics: 確定/予測 レポートから transaction 形式の Query Table を作成")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="ワークスペース内の view 一覧")
    p_list.add_argument("--filter", default="", help="名前部分一致フィルタ")
    p_list.set_defaults(func=cmd_list)

    p_ws = sub.add_parser("workspaces", help="所有ワークスペース一覧（orgId / workspaceId 探索用）")
    p_ws.set_defaults(func=cmd_workspaces)

    p_ins = sub.add_parser("inspect", help="指定 view の列一覧と自動判定結果")
    p_ins.add_argument("--source-table", required=True)
    p_ins.set_defaults(func=cmd_inspect)

    p_build = sub.add_parser("build", help="transaction 形式の Query Table を生成")
    p_build.add_argument("--source-table", required=True, help="元テーブル名（例: 商談）")
    p_build.add_argument("--output-name", required=True, help="作成する Query Table 名（例: confirmed_transactions）")
    p_build.add_argument("--status", required=True, choices=["confirmed", "forecast"],
                        help="transaction_status 列に入れる値")
    p_build.add_argument("--stage-column", default=None, help="ステージフィルタの列名（例: ステージ）")
    p_build.add_argument("--stage-values", nargs="*", default=None,
                        help="ステージフィルタに含める値（複数）")
    p_build.add_argument("--qualify-column", default=None,
                        help="商談単位の期間フィルタ列名（例: 国内仕入１　支払日）")
    p_build.add_argument("--qualify-from", default=None,
                        help="期間フィルタの開始日 YYYY-MM-DD")
    p_build.add_argument("--qualify-to", default=None,
                        help="期間フィルタの終了日 YYYY-MM-DD")
    p_build.add_argument("--replace", action="store_true",
                        help="同名 Query Table が既存ならば削除して作り直す")
    p_build.add_argument("--dry-run", action="store_true", help="SQL の表示のみ。Query Table は作成しない")
    p_build.set_defaults(func=cmd_build)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
