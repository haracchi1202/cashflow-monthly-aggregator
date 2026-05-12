"""旧 Excel ファイルと confirmed_transactions.xlsx を比較して差異を可視化する診断スクリプト。"""
import pandas as pd
import re

def normalize_month(v):
    if pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp,)):
        return f'{v.year:04d}-{v.month:02d}'
    s = str(v).strip()
    m = re.match(r'^\s*(20\d{2})[\s/\-.年]\s*(\d{1,2})', s)
    if m:
        return f'{int(m.group(1)):04d}-{int(m.group(2)):02d}'
    m = re.search(r'(\d{1,2})\s*月.*?(20\d{2})', s)
    if m:
        return f'{int(m.group(2)):04d}-{int(m.group(1)):02d}'
    try:
        ts = pd.to_datetime(s, errors='raise')
        return f'{ts.year:04d}-{ts.month:02d}'
    except Exception:
        return None


files_income = [
    (1, r'C:/Users/hara/Downloads/入金日１_2024年9月～2025年8月_複製済み (1).xlsx'),
    (2, r'C:/Users/hara/Downloads/入金日２_2023年9月～2024年8月_複製済み (1).xlsx'),
]
files_payment = [
    (1, r'C:/Users/hara/Downloads/国内支払１_2024年9月～2025年8月_複製済み (1).xlsx'),
    (2, r'C:/Users/hara/Downloads/国内支払２_2023年9月～2024年8月_複製済み (1).xlsx'),
    (3, r'C:/Users/hara/Downloads/国内支払３_2023年9月～2024年8月_複製済み (1).xlsx'),
    (4, r'C:/Users/hara/Downloads/国内支払４_2023年9月～2024年8月_複製済み (1).xlsx'),
    (5, r'C:/Users/hara/Downloads/国内支払５_2023年9月～2024年8月_複製済み (1).xlsx'),
]


def _read_old_file(fp, kind_label):
    """旧 Excel を読み込み、(月, 商談名, 金額) のレコードリストを返す。"""
    df = pd.read_excel(fp, header=6, engine='openpyxl', dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    date_col = amt_col = None
    for c in df.columns:
        if '入金日' in c or '支払日' in c:
            date_col = c
            break
    for c in df.columns:
        if ('入金額' in c or '原価総額' in c) and '合計' in c:
            amt_col = c
            break
    last_month = None
    rows = []
    for _, row in df.iterrows():
        raw_month = row.get(date_col) if date_col else None
        amount = row.get(amt_col) if amt_col else None
        deal = row.get('商談名')
        ym = normalize_month(raw_month)
        if ym:
            last_month = ym
        else:
            ym = last_month
        if amount is None or (isinstance(amount, float) and pd.isna(amount)):
            continue
        if not deal or (isinstance(deal, float) and pd.isna(deal)):
            continue
        if ym is None:
            continue
        try:
            a = float(amount)
        except Exception:
            continue
        if a == 0:
            continue
        rows.append({'month': ym, 'deal': str(deal).strip(), 'amount': a, 'kind': kind_label})
    return rows


old_income_records = []
for idx, fp in files_income:
    old_income_records.extend(_read_old_file(fp, f'入金{idx}'))

old_payment_records = []
for idx, fp in files_payment:
    old_payment_records.extend(_read_old_file(fp, f'支払{idx}'))

old_income_total = sum(r['amount'] for r in old_income_records)
old_payment_total = sum(r['amount'] for r in old_payment_records)

ct = pd.read_excel(r'C:/Users/hara/Downloads/confirmed_transactions.xlsx')
ct['月'] = ct['取引日'].apply(lambda d: f'{d.year:04d}-{d.month:02d}')
new_income = ct[ct['transaction_type'] == 'income'].copy()
new_payment = ct[ct['transaction_type'] == 'payment'].copy()

print('=== 旧ファイル ===')
print(f'  入金: {old_income_total:>15,.0f} 円 ({len(old_income_records)} 行)')
print(f'  支払: {old_payment_total:>15,.0f} 円 ({len(old_payment_records)} 行)')
print()
print('=== 新 confirmed_transactions ===')
new_in_sum = float(new_income['金額'].sum())
new_pay_sum = float(new_payment['金額'].sum())
print(f'  入金: {new_in_sum:>15,.0f} 円 ({len(new_income)} 行)')
print(f'  支払: {new_pay_sum:>15,.0f} 円 ({len(new_payment)} 行)')
print()
print(f'  入金差: {new_in_sum - old_income_total:>+15,.0f} 円')
print(f'  支払差: {new_pay_sum - old_payment_total:>+15,.0f} 円')

print()
print('=== 月別 入金 比較 ===')
old_income_by_month = {}
for r in old_income_records:
    old_income_by_month[r['month']] = old_income_by_month.get(r['month'], 0) + r['amount']
new_income_by_month = new_income.groupby('月')['金額'].sum().to_dict()

all_months = sorted(set(old_income_by_month) | set(new_income_by_month))
print(f'{"月":<10}{"旧":>16}{"新":>16}{"差額":>16}')
for m in all_months:
    o = old_income_by_month.get(m, 0)
    n = new_income_by_month.get(m, 0)
    d = n - o
    flag = '  *' if abs(d) > 1 else ''
    print(f'{m:<10}{int(o):>16,}{int(n):>16,}{int(d):>+16,}{flag}')

print()
print('=== 月別 支払 比較 ===')
old_payment_by_month = {}
for r in old_payment_records:
    old_payment_by_month[r['month']] = old_payment_by_month.get(r['month'], 0) + r['amount']
new_payment_by_month = new_payment.groupby('月')['金額'].sum().to_dict()
all_months = sorted(set(old_payment_by_month) | set(new_payment_by_month))
print(f'{"月":<10}{"旧":>16}{"新":>16}{"差額":>16}')
for m in all_months:
    o = old_payment_by_month.get(m, 0)
    n = new_payment_by_month.get(m, 0)
    d = n - o
    flag = '  *' if abs(d) > 1 else ''
    print(f'{m:<10}{int(o):>16,}{int(n):>16,}{int(d):>+16,}{flag}')

# 商談名ベースの差異も調べる（共通月のみ）
print()
print('=== 旧のみに含まれる商談（合計上位 10）===')
old_deals = {}
for r in old_income_records + old_payment_records:
    old_deals[r['deal']] = old_deals.get(r['deal'], 0) + r['amount']

new_deals = {}
for _, r in ct.iterrows():
    new_deals[r['商談名']] = new_deals.get(r['商談名'], 0) + r['金額']

old_only = {k: v for k, v in old_deals.items() if k not in new_deals}
for k, v in sorted(old_only.items(), key=lambda kv: -kv[1])[:10]:
    print(f'  {k[:50]:<52} {int(v):>12,} 円')

print()
print('=== 新のみに含まれる商談（合計上位 10）===')
new_only = {k: v for k, v in new_deals.items() if k not in old_deals}
for k, v in sorted(new_only.items(), key=lambda kv: -kv[1])[:10]:
    print(f'  {k[:50]:<52} {int(v):>12,} 円')

print()
print(f'共通商談数: {len(set(old_deals) & set(new_deals))}')
print(f'旧のみ商談数: {len(old_only)}')
print(f'新のみ商談数: {len(new_only)}')
