"""スナップショットの保存・一覧・比較・変更レポート生成。

集計結果（明細レコード）と月別サマリーを JSON ファイルに保存し、
次回の集計と差分比較できるようにする。

[同一性キー]
    商談名は基本的に変わらないという前提に基づき、次のキーで照合:
        key = (source_group, deal_name, client_name, sequence)

    同じ商談名・顧客名が同じファイル区分内に複数ある場合は raw_row_index 順に
    1, 2, 3... と sequence を割り当て、安定したマッチングを可能にする。

[検出する変更]
    - target_month が変わった (入金月 / 支払日 の月が動いた)
    - transaction_date が変わった (元の日付セルそのもの)
    - amount が変わった (金額が増減した)

[結果カテゴリ]
    - added     : 今回のみ存在
    - removed   : 前回のみ存在
    - changed   : 両方にあり値が変わった
    - unchanged : 両方にあり値も同じ
"""
from __future__ import annotations

import io
import json
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from excel_reader import FileResult, format_month_jp, format_yen


SNAPSHOT_DIR = Path(__file__).parent / "snapshots"


# =========================================================
# 保存
# =========================================================
def _safe_name(name: str) -> str:
    s = re.sub(r"[^\w\-]+", "_", name.strip())
    return s[:80] or "snapshot"


def ensure_snapshot_dir() -> Path:
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    return SNAPSHOT_DIR


def save_snapshot(
    name: str,
    results: list[FileResult],
    monthly_df: pd.DataFrame,
    written_log: list[dict[str, Any]] | None = None,
    update_scope: str | None = None,
    month_from: str | None = None,
    month_to: str | None = None,
) -> Path:
    """現在の集計結果をスナップショット JSON として保存する。"""
    ensure_snapshot_dir()
    ts = datetime.now()
    safe = _safe_name(name)
    fname = f"{ts.strftime('%Y%m%d_%H%M%S')}_{safe}.json"
    path = SNAPSHOT_DIR / fname

    details: list[dict[str, Any]] = []
    for r in results:
        for row in r.detail_rows:
            details.append(_serialize_record(row))
        # 除外行も別キーで保存しておく（参考）
    excluded: list[dict[str, Any]] = []
    for r in results:
        for row in r.excluded_rows:
            excluded.append(_serialize_record(row, is_excluded=True))

    monthly: list[dict[str, Any]] = []
    if not monthly_df.empty:
        for _, row in monthly_df.iterrows():
            monthly.append({k: (None if pd.isna(v) else v) for k, v in row.to_dict().items()})

    payload = {
        "name": name,
        "saved_at": ts.isoformat(timespec="seconds"),
        "update_scope": update_scope,
        "month_from": month_from,
        "month_to": month_to,
        "details": details,
        "excluded": excluded,
        "monthly_summary": monthly,
        "written_log": written_log or [],
    }
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=_json_default)
    return path


def _json_default(value: Any) -> Any:
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return str(value)


def _serialize_record(row: dict[str, Any], is_excluded: bool = False) -> dict[str, Any]:
    keep_keys = [
        "source_file", "source_group", "transaction_status", "transaction_type",
        "target_month", "transaction_date", "amount",
        "client_name", "deal_name", "raw_row_index",
        "exclude_reason",
    ]
    out: dict[str, Any] = {}
    for k in keep_keys:
        v = row.get(k)
        if isinstance(v, (pd.Timestamp, datetime)):
            v = v.isoformat()
        if isinstance(v, float) and pd.isna(v):
            v = None
        out[k] = v
    if is_excluded:
        out["is_excluded"] = True
    return out


# =========================================================
# 一覧 / 読み込み / 削除
# =========================================================
@dataclass
class SnapshotInfo:
    path: Path
    name: str
    saved_at: str
    detail_count: int
    monthly_count: int


def list_snapshots() -> list[SnapshotInfo]:
    ensure_snapshot_dir()
    out: list[SnapshotInfo] = []
    for p in sorted(SNAPSHOT_DIR.glob("*.json"), reverse=True):
        try:
            with p.open("r", encoding="utf-8") as f:
                data = json.load(f)
            out.append(
                SnapshotInfo(
                    path=p,
                    name=data.get("name", p.stem),
                    saved_at=data.get("saved_at", ""),
                    detail_count=len(data.get("details", [])),
                    monthly_count=len(data.get("monthly_summary", [])),
                )
            )
        except Exception:
            continue
    return out


def load_snapshot(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def delete_snapshot(path: Path) -> None:
    if path.exists():
        path.unlink()


# =========================================================
# 比較
# =========================================================
@dataclass
class ComparisonRow:
    source_group: str
    deal_name: str
    client_name: str
    sequence: int
    status: str  # "added" / "removed" / "changed" / "unchanged"
    changes: dict[str, tuple[Any, Any]] = field(default_factory=dict)
    record_old: dict[str, Any] | None = None
    record_new: dict[str, Any] | None = None


def _build_keyed_records(
    details: list[dict[str, Any]],
) -> dict[tuple[str, str, str, int], dict[str, Any]]:
    """(source_group, deal_name, client_name, sequence) → record の辞書を作る。

    同じ商談名+顧客名がファイル区分内に複数あれば raw_row_index 順に sequence を割当。
    """
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for r in details:
        if r.get("amount") is None:
            continue
        key3 = (
            str(r.get("source_group") or ""),
            str(r.get("deal_name") or ""),
            str(r.get("client_name") or ""),
        )
        grouped[key3].append(r)
    out: dict[tuple[str, str, str, int], dict[str, Any]] = {}
    for key3, rows in grouped.items():
        rows_sorted = sorted(rows, key=lambda r: int(r.get("raw_row_index") or 0))
        for i, r in enumerate(rows_sorted, 1):
            out[(*key3, i)] = r
    return out


def _values_equal(a: Any, b: Any) -> bool:
    if a is None and b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        try:
            return float(a) == float(b)
        except Exception:
            return False
    return str(a) == str(b)


COMPARED_FIELDS = ("target_month", "transaction_date", "amount")


def compare_snapshots(
    old_snapshot: dict[str, Any],
    new_snapshot: dict[str, Any],
) -> list[ComparisonRow]:
    """2 スナップショットを比較し、変更行リストを返す。"""
    old_keyed = _build_keyed_records(old_snapshot.get("details", []))
    new_keyed = _build_keyed_records(new_snapshot.get("details", []))

    rows: list[ComparisonRow] = []
    all_keys = set(old_keyed.keys()) | set(new_keyed.keys())
    for key in sorted(all_keys, key=lambda k: (k[0], k[1], k[2], k[3])):
        sg, dn, cn, seq = key
        ro = old_keyed.get(key)
        rn = new_keyed.get(key)
        if ro is None and rn is not None:
            rows.append(ComparisonRow(sg, dn, cn, seq, "added", record_new=rn))
            continue
        if rn is None and ro is not None:
            rows.append(ComparisonRow(sg, dn, cn, seq, "removed", record_old=ro))
            continue
        # both
        changes: dict[str, tuple[Any, Any]] = {}
        for f in COMPARED_FIELDS:
            ov = ro.get(f)
            nv = rn.get(f)
            if not _values_equal(ov, nv):
                changes[f] = (ov, nv)
        status = "changed" if changes else "unchanged"
        rows.append(ComparisonRow(sg, dn, cn, seq, status, changes, ro, rn))
    return rows


def comparison_to_summary(rows: list[ComparisonRow]) -> dict[str, dict[str, int]]:
    """ファイル区分ごとに added/removed/changed/unchanged をカウント。"""
    out: dict[str, dict[str, int]] = defaultdict(lambda: {"added": 0, "removed": 0, "changed": 0, "unchanged": 0})
    for r in rows:
        out[r.source_group][r.status] += 1
    return out


# =========================================================
# 変更レポート Excel
# =========================================================
_FIELD_JP = {
    "target_month": "月",
    "transaction_date": "日付セル",
    "amount": "金額",
}


def _fmt_field(field_name: str, v: Any) -> str:
    if v is None:
        return ""
    if field_name == "amount":
        return format_yen(v)
    if field_name == "target_month":
        return format_month_jp(str(v))
    return str(v)


def build_comparison_report(
    rows: list[ComparisonRow],
    old_snapshot: dict[str, Any],
    new_snapshot: dict[str, Any],
) -> bytes:
    """変更レポート Excel を生成して bytes を返す。

    シート構成:
        1. サマリー
        2. 変更あり (changed)
        3. 追加 (added)
        4. 削除 (removed)
        5. 全変更ログ（status 列付き）
        6. メタ情報
    """
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
        yen_fmt = wb.add_format({"num_format": '#,##0"円"'})
        neg_yen_fmt = wb.add_format({"num_format": '#,##0"円";[Red]-#,##0"円"'})

        # 1. サマリー
        summary = comparison_to_summary(rows)
        sum_rows = []
        for sg, counts in sorted(summary.items()):
            sum_rows.append({
                "ファイル区分": sg,
                "追加": counts["added"],
                "削除": counts["removed"],
                "変更": counts["changed"],
                "変更なし": counts["unchanged"],
                "合計": sum(counts.values()),
            })
        if not sum_rows:
            sum_rows = [{
                "ファイル区分": "(データなし)",
                "追加": 0, "削除": 0, "変更": 0, "変更なし": 0, "合計": 0,
            }]
        sum_df = pd.DataFrame(sum_rows)
        sum_df.to_excel(writer, sheet_name="サマリー", index=False)
        ws = writer.sheets["サマリー"]
        ws.set_column("A:A", 14)
        ws.set_column("B:F", 12)
        for c, v in enumerate(sum_df.columns):
            ws.write(0, c, v, header_fmt)

        # 2. 変更あり (changed)
        changed_rows = []
        for r in rows:
            if r.status != "changed":
                continue
            d = {
                "ファイル区分": r.source_group,
                "商談名": r.deal_name,
                "顧客名": r.client_name,
                "通番": r.sequence,
            }
            for f in COMPARED_FIELDS:
                if f in r.changes:
                    ov, nv = r.changes[f]
                    d[_FIELD_JP[f] + " 前"] = _fmt_field(f, ov)
                    d[_FIELD_JP[f] + " 後"] = _fmt_field(f, nv)
                    if f == "amount":
                        try:
                            diff = float(nv or 0) - float(ov or 0)
                            d["金額差額"] = format_yen(diff)
                        except Exception:
                            d["金額差額"] = ""
            d["元ファイル(前)"] = (r.record_old or {}).get("source_file", "")
            d["元ファイル(後)"] = (r.record_new or {}).get("source_file", "")
            changed_rows.append(d)
        if not changed_rows:
            changed_rows = [{"ファイル区分": "(変更なし)"}]
        chg_df = pd.DataFrame(changed_rows)
        chg_df.to_excel(writer, sheet_name="変更あり", index=False)
        ws = writer.sheets["変更あり"]
        ws.set_column("A:A", 12)
        ws.set_column("B:C", 28)
        ws.set_column("D:D", 6)
        ws.set_column("E:Z", 16)
        for c, v in enumerate(chg_df.columns):
            ws.write(0, c, v, header_fmt)

        # 3. 追加 (added)
        added_rows = []
        for r in rows:
            if r.status != "added":
                continue
            n = r.record_new or {}
            added_rows.append({
                "ファイル区分": r.source_group,
                "商談名": r.deal_name,
                "顧客名": r.client_name,
                "月": format_month_jp(str(n.get("target_month") or "")),
                "日付セル": n.get("transaction_date") or "",
                "金額": format_yen(n.get("amount")),
                "元ファイル": n.get("source_file") or "",
            })
        if not added_rows:
            added_rows = [{"ファイル区分": "(追加なし)"}]
        add_df = pd.DataFrame(added_rows)
        add_df.to_excel(writer, sheet_name="追加", index=False)
        ws = writer.sheets["追加"]
        ws.set_column("A:A", 12)
        ws.set_column("B:C", 28)
        ws.set_column("D:E", 14)
        ws.set_column("F:F", 16)
        ws.set_column("G:G", 36)
        for c, v in enumerate(add_df.columns):
            ws.write(0, c, v, header_fmt)

        # 4. 削除 (removed)
        removed_rows = []
        for r in rows:
            if r.status != "removed":
                continue
            o = r.record_old or {}
            removed_rows.append({
                "ファイル区分": r.source_group,
                "商談名": r.deal_name,
                "顧客名": r.client_name,
                "月": format_month_jp(str(o.get("target_month") or "")),
                "日付セル": o.get("transaction_date") or "",
                "金額": format_yen(o.get("amount")),
                "元ファイル": o.get("source_file") or "",
            })
        if not removed_rows:
            removed_rows = [{"ファイル区分": "(削除なし)"}]
        rem_df = pd.DataFrame(removed_rows)
        rem_df.to_excel(writer, sheet_name="削除", index=False)
        ws = writer.sheets["削除"]
        ws.set_column("A:A", 12)
        ws.set_column("B:C", 28)
        ws.set_column("D:E", 14)
        ws.set_column("F:F", 16)
        ws.set_column("G:G", 36)
        for c, v in enumerate(rem_df.columns):
            ws.write(0, c, v, header_fmt)

        # 5. 全変更ログ
        all_rows = []
        status_jp = {"added": "追加", "removed": "削除", "changed": "変更", "unchanged": "変更なし"}
        for r in rows:
            o = r.record_old or {}
            n = r.record_new or {}
            all_rows.append({
                "状態": status_jp.get(r.status, r.status),
                "ファイル区分": r.source_group,
                "商談名": r.deal_name,
                "顧客名": r.client_name,
                "通番": r.sequence,
                "月 前": format_month_jp(str(o.get("target_month") or "")) if o else "",
                "月 後": format_month_jp(str(n.get("target_month") or "")) if n else "",
                "日付セル 前": o.get("transaction_date") if o else "",
                "日付セル 後": n.get("transaction_date") if n else "",
                "金額 前": format_yen(o.get("amount")) if o else "",
                "金額 後": format_yen(n.get("amount")) if n else "",
            })
        if not all_rows:
            all_rows = [{"状態": "(差分なし)"}]
        all_df = pd.DataFrame(all_rows)
        all_df.to_excel(writer, sheet_name="全比較ログ", index=False)
        ws = writer.sheets["全比較ログ"]
        ws.set_column("A:A", 10)
        ws.set_column("B:B", 12)
        ws.set_column("C:D", 28)
        ws.set_column("E:E", 6)
        ws.set_column("F:K", 16)
        for c, v in enumerate(all_df.columns):
            ws.write(0, c, v, header_fmt)

        # 6. メタ情報
        meta = [
            ["前回スナップショット", old_snapshot.get("name", "")],
            ["前回保存日時", old_snapshot.get("saved_at", "")],
            ["今回スナップショット", new_snapshot.get("name", "")],
            ["今回保存日時", new_snapshot.get("saved_at", "")],
            ["前回 明細件数", len(old_snapshot.get("details", []))],
            ["今回 明細件数", len(new_snapshot.get("details", []))],
            ["比較生成日時", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ]
        meta_df = pd.DataFrame(meta, columns=["項目", "値"])
        meta_df.to_excel(writer, sheet_name="メタ情報", index=False)
        ws = writer.sheets["メタ情報"]
        ws.set_column("A:A", 20)
        ws.set_column("B:B", 60)
        for c, v in enumerate(meta_df.columns):
            ws.write(0, c, v, header_fmt)

    bio.seek(0)
    return bio.read()


# =========================================================
# DataFrame ビュー（画面表示用）
# =========================================================
def comparison_to_dataframe(rows: list[ComparisonRow]) -> pd.DataFrame:
    """画面表示用の DataFrame に変換。"""
    out: list[dict[str, Any]] = []
    status_jp = {"added": "追加", "removed": "削除", "changed": "変更", "unchanged": "変更なし"}
    for r in rows:
        o = r.record_old or {}
        n = r.record_new or {}
        rec = {
            "状態": status_jp.get(r.status, r.status),
            "ファイル区分": r.source_group,
            "商談名": r.deal_name,
            "顧客名": r.client_name,
            "通番": r.sequence,
            "月(前)": format_month_jp(str(o.get("target_month") or "")) if o else "",
            "月(後)": format_month_jp(str(n.get("target_month") or "")) if n else "",
            "日付セル(前)": o.get("transaction_date") if o else "",
            "日付セル(後)": n.get("transaction_date") if n else "",
            "金額(前)": format_yen(o.get("amount")) if o else "",
            "金額(後)": format_yen(n.get("amount")) if n else "",
        }
        out.append(rec)
    return pd.DataFrame(out)
