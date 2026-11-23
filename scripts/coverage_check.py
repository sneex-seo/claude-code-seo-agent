# -*- coding: utf-8 -*-
"""Проверка ПОКРЫТИЯ семантического ядра (Шаг 3 скилла semantic-clustering).

Сравнивает собранное ЯДРО с ЦЕЛЬЮ (url-keywords конкурентов и/или эталон):
сколько % ключей и объёма цели покрыто ядром + что именно пропущено (добери эти зоны).

  python scripts/coverage_check.py ядро.csv конкурент1.csv конкурент2.csv
  python scripts/coverage_check.py ядро.csv эталон.xlsx --sheet "Кластери" --top 25

Формат файлов: csv/tsv (колонки keyword[,volume]) или xlsx. Колонки ищутся по заголовку
(keyword/ключ/ключове слово; volume/частотность/...); нет заголовка ключа → берётся 1-й столбец.
xlsx без --sheet: авто - читаются только листы с колонкой-ключом.
"""
import argparse, csv, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')

KW_COLS  = ('keyword', 'ключ', 'ключевой запрос', 'ключове слово', 'фраза', 'query')
VOL_COLS = ('volume', 'vol', 'частотность', 'частотність', 'частота', 'region_queries_count')

def norm(s):
    return re.sub(r'\s+', ' ', str(s).lower().replace('ё', 'е').strip())

def _pick(header, options):
    for i, h in enumerate(header):
        if (h or '').strip().lower() in options:
            return i
    return None

def load_keys(path, sheet=None):
    ext = os.path.splitext(path)[1].lower()
    out = {}
    if ext in ('.csv', '.tsv'):
        rows = list(csv.reader(open(path, encoding='utf-8-sig'), delimiter='\t' if ext == '.tsv' else ','))
        if not rows:
            return out
        ki = _pick(rows[0], KW_COLS); vi = _pick(rows[0], VOL_COLS)
        if ki is None:
            ki = 0
        for r in rows[1:]:
            if len(r) <= ki or not str(r[ki]).strip():
                continue
            v = 0
            if vi is not None and len(r) > vi:
                try: v = int(float(r[vi]))
                except (ValueError, TypeError): v = 0
            k = norm(r[ki]); out[k] = max(out.get(k, 0), v)
    elif ext in ('.xlsx', '.xlsm'):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        for ws in ([wb[sheet]] if sheet else wb.worksheets):
            it = ws.iter_rows(values_only=True)
            header = next(it, None)
            if header is None:
                continue
            ki = _pick(list(header), KW_COLS); vi = _pick(list(header), VOL_COLS)
            if sheet is None and ki is None:
                continue  # авто-режим: пропускаем листы без колонки-ключа
            hdr_is_data = ki is None
            if ki is None:
                ki = 0
            def add(row):
                if len(row) <= ki or row[ki] is None or not str(row[ki]).strip():
                    return
                v = row[vi] if (vi is not None and len(row) > vi and isinstance(row[vi], (int, float))) else 0
                k = norm(row[ki]); out[k] = max(out.get(k, 0), int(v))
            if hdr_is_data:
                add(header)
            for row in it:
                add(row)
    else:
        raise SystemExit(f'формат не підтримується: {ext}')
    return out

def main():
    ap = argparse.ArgumentParser(description='Покриття ядра відносно цілі (конкуренти або еталон)')
    ap.add_argument('core', help='зібране ядро (csv/xlsx)')
    ap.add_argument('target', nargs='+', help='ціль: url-keywords конкурентів і/або еталон')
    ap.add_argument('--top', type=int, default=20, help='скільки пропущеного показати (за замовчуванням 20)')
    ap.add_argument('--sheet', help='конкретний аркуш xlsx')
    args = ap.parse_args()

    core_sheet = args.sheet if args.core.lower().endswith(('.xlsx', '.xlsm')) else None
    core = load_keys(args.core, core_sheet)
    target = {}
    for t in args.target:
        for k, v in load_keys(t, args.sheet).items():
            target[k] = max(target.get(k, 0), v)
    if not target:
        raise SystemExit('у цілі немає ключів - перевір файл або аркуш (--sheet)')

    ck, tk = set(core), set(target)
    covered = ck & tk; missing = tk - covered
    cov_k = 100 * len(covered) / len(tk)
    tvol = sum(target.values()) or 1
    mvol = sum(target[k] for k in missing)
    cov_v = 100 * (tvol - mvol) / tvol

    print(f'Ядро: {len(ck)} ключів | Ціль: {len(tk)} ключів (∑{tvol})')
    print(f'Покриття ключів: {cov_k:.0f}%  ({len(covered)}/{len(tk)})')
    print(f'Покриття обсягу: {cov_v:.0f}%  (пропущено ∑{mvol})')
    print(f'Вердикт: {"OK (>=85%)" if cov_k >= 85 else "МАЛО (<85%) - добери сіди або url за списком нижче"}')
    if missing:
        print(f'\n--- ТОП-{args.top} пропущеного (добери ці зони) ---')
        for k in sorted(missing, key=lambda x: -target[x])[:args.top]:
            print(f'  {str(target[k]).rjust(7)}  {k}')

if __name__ == '__main__':
    main()
