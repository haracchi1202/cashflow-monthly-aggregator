"""列名パターンマッチングによる入金/支払列の自動判定。

確定レポートの想定列名:
  入金日:   クライアント入金日１, クライアント入金日２
  入金額:   クライアント入金額１(税込), クライアント入金額２(税込)
  支払日:   国内仕入１ 支払日, 国内仕入２ 支払日, ...
  支払額:   国内仕入１ 原価総額(税込), 国内仕入２ 原価総額(税込), ...

  共通:     商談名, クライアント名

予測レポートでも同様の構造（命名は揺れる可能性あり）。

全角数字 / 半角数字 / 全角空白 / カッコ表記の有無を許容する。
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field


def _normalize_for_match(name: str) -> str:
    """全角→半角変換、空白・記号除去で比較しやすい形に。"""
    if name is None:
        return ""
    s = unicodedata.normalize("NFKC", str(name))
    s = s.replace(" ", "").replace("　", "")
    return s


@dataclass
class TransactionPair:
    """1つの transaction (日付列 + 金額列) のペア。"""
    transaction_type: str   # "income" / "payment"
    payment_round: str      # "income_1" / "payment_1" など
    date_column: str        # 元の列名（クォート前）
    amount_column: str      # 元の列名（クォート前）
    payee_column: str | None = None  # 支払先名 列（payment のみ。無ければ None）


@dataclass
class DetectionResult:
    deal_column: str | None
    client_column: str | None
    pairs: list[TransactionPair] = field(default_factory=list)
    income_count: int = 0
    payment_count: int = 0
    warnings: list[str] = field(default_factory=list)


# パターン: 確定 - クライアント入金日(1〜9), クライアント入金額(1〜9)(税込)
INCOME_DATE_RE = re.compile(r"^クライアント入金日(\d)$")
INCOME_AMOUNT_RE = re.compile(r"^クライアント入金額(\d)\(税込\)$")

# パターン: 確定 - 国内仕入(1〜9)支払日, 国内仕入(1〜9)原価総額(税込)
PAYMENT_DATE_RE = re.compile(r"^国内仕入(\d)支払日$")
PAYMENT_AMOUNT_RE = re.compile(r"^国内仕入(\d)原価総額\(税込\)$")
# 支払先 / 支払先名 どちらの命名にも対応
PAYMENT_PAYEE_RE = re.compile(r"^国内仕入(\d)支払先名?$")

# パターン: 予測 - 予測(初回|残金|予備)(入金|支払)日 / 予測(初回|残金|予備)(入金|支払)額
# 順序: 初回 → 残金 → 予備
_FORECAST_ROUND_ORDER = {"初回": 1, "残金": 2, "予備": 3}
FORECAST_DATE_RE = re.compile(r"^予測(初回|残金|予備)(入金|支払)日$")
FORECAST_AMOUNT_RE = re.compile(r"^予測(初回|残金|予備)(入金|支払)額$")
FORECAST_PAYEE_RE = re.compile(r"^予測(初回|残金|予備)支払先名?$")

# 商談名・クライアント名は完全一致を優先、部分マッチもフォールバック
DEAL_KEYWORDS = ["商談名", "案件名", "件名"]
CLIENT_KEYWORDS = ["クライアント名", "顧客名", "取引先", "クライアント"]


def detect_columns(columns: list[str]) -> DetectionResult:
    """列名リストを受け取り、income/payment ペアと共通列を判定。"""
    norm_map: dict[str, str] = {c: _normalize_for_match(c) for c in columns}

    deal_col: str | None = None
    client_col: str | None = None
    for kw in DEAL_KEYWORDS:
        for c, n in norm_map.items():
            if n == _normalize_for_match(kw):
                deal_col = c
                break
        if deal_col:
            break
    if not deal_col:
        for kw in DEAL_KEYWORDS:
            for c, n in norm_map.items():
                if _normalize_for_match(kw) in n:
                    deal_col = c
                    break
            if deal_col:
                break

    for kw in CLIENT_KEYWORDS:
        for c, n in norm_map.items():
            if n == _normalize_for_match(kw):
                client_col = c
                break
        if client_col:
            break
    if not client_col:
        for kw in CLIENT_KEYWORDS:
            for c, n in norm_map.items():
                if _normalize_for_match(kw) in n:
                    client_col = c
                    break
            if client_col:
                break

    # 確定パターン - 入金: 日付と金額をインデックスで突合
    income_dates: dict[int, str] = {}
    income_amounts: dict[int, str] = {}
    for c, n in norm_map.items():
        m = INCOME_DATE_RE.match(n)
        if m:
            income_dates[int(m.group(1))] = c
        m = INCOME_AMOUNT_RE.match(n)
        if m:
            income_amounts[int(m.group(1))] = c

    # 確定パターン - 支払: 同様
    payment_dates: dict[int, str] = {}
    payment_amounts: dict[int, str] = {}
    payment_payees: dict[int, str] = {}
    for c, n in norm_map.items():
        m = PAYMENT_DATE_RE.match(n)
        if m:
            payment_dates[int(m.group(1))] = c
        m = PAYMENT_AMOUNT_RE.match(n)
        if m:
            payment_amounts[int(m.group(1))] = c
        m = PAYMENT_PAYEE_RE.match(n)
        if m:
            payment_payees[int(m.group(1))] = c

    # 予測パターン - 初回/残金/予備 × 入金/支払
    forecast_income_dates: dict[int, str] = {}
    forecast_income_amounts: dict[int, str] = {}
    forecast_payment_dates: dict[int, str] = {}
    forecast_payment_amounts: dict[int, str] = {}
    forecast_payment_payees: dict[int, str] = {}
    for c, n in norm_map.items():
        m = FORECAST_DATE_RE.match(n)
        if m:
            round_name, kind = m.group(1), m.group(2)
            idx = _FORECAST_ROUND_ORDER.get(round_name, 99)
            if kind == "入金":
                forecast_income_dates[idx] = c
            else:
                forecast_payment_dates[idx] = c
        m = FORECAST_AMOUNT_RE.match(n)
        if m:
            round_name, kind = m.group(1), m.group(2)
            idx = _FORECAST_ROUND_ORDER.get(round_name, 99)
            if kind == "入金":
                forecast_income_amounts[idx] = c
            else:
                forecast_payment_amounts[idx] = c
        m = FORECAST_PAYEE_RE.match(n)
        if m:
            round_name = m.group(1)
            idx = _FORECAST_ROUND_ORDER.get(round_name, 99)
            forecast_payment_payees[idx] = c

    pairs: list[TransactionPair] = []
    warnings: list[str] = []

    for idx in sorted(set(income_dates.keys()) | set(income_amounts.keys())):
        d = income_dates.get(idx)
        a = income_amounts.get(idx)
        if d and a:
            pairs.append(TransactionPair("income", f"income_{idx}", d, a))
        else:
            warnings.append(f"入金{idx}: 日付列={d} / 金額列={a} のペアが揃わずスキップ")

    for idx in sorted(set(payment_dates.keys()) | set(payment_amounts.keys())):
        d = payment_dates.get(idx)
        a = payment_amounts.get(idx)
        p = payment_payees.get(idx)
        if d and a:
            pairs.append(TransactionPair("payment", f"payment_{idx}", d, a, p))
        else:
            warnings.append(f"支払{idx}: 日付列={d} / 金額列={a} のペアが揃わずスキップ")

    for idx in sorted(set(forecast_income_dates.keys()) | set(forecast_income_amounts.keys())):
        d = forecast_income_dates.get(idx)
        a = forecast_income_amounts.get(idx)
        if d and a:
            pairs.append(TransactionPair("income", f"forecast_income_{idx}", d, a))
        else:
            warnings.append(f"予測入金{idx}: 日付列={d} / 金額列={a} のペアが揃わずスキップ")

    for idx in sorted(set(forecast_payment_dates.keys()) | set(forecast_payment_amounts.keys())):
        d = forecast_payment_dates.get(idx)
        a = forecast_payment_amounts.get(idx)
        p = forecast_payment_payees.get(idx)
        if d and a:
            pairs.append(TransactionPair("payment", f"forecast_payment_{idx}", d, a, p))
        else:
            warnings.append(f"予測支払{idx}: 日付列={d} / 金額列={a} のペアが揃わずスキップ")

    return DetectionResult(
        deal_column=deal_col,
        client_column=client_col,
        pairs=pairs,
        income_count=sum(1 for p in pairs if p.transaction_type == "income"),
        payment_count=sum(1 for p in pairs if p.transaction_type == "payment"),
        warnings=warnings,
    )
