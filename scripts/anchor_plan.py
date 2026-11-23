# -*- coding: utf-8 -*-
"""Кандидаты анкоров для линкбилдинга из выгрузки позиций (+ существующие анкоры).

Логика перенесена из замороженного anchor_picker.py (без FastAPI):
- ключи страницы сортируются по позиции и частотке;
- анкоры, использованные >= порога раз, помечаются как переспам (избегать);
- приоритет: P1 (топ-3), P2 (топ-10), P3 (дальше).

Примеры:
  python scripts/anchor_plan.py positions.csv --domain site.ua --out anchors.xlsx
  python scripts/anchor_plan.py positions.csv --anchors existing.csv --domain site.ua --out anchors.xlsx
"""
import argparse
import csv
import io
import sys
from collections import Counter, defaultdict
from urllib.parse import urlparse

import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')


def read_rows(path: str) -> list:
    raw = open(path, 'rb').read()
    for enc in ('utf-8-sig', 'utf-16', 'cp1251'):
        try:
            text = raw.decode(enc)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    else:
        raise SystemExit(f'Не вдалося визначити кодування {path}')
    first = text.splitlines()[0] if text.splitlines() else ''
    delim = max(',;\t', key=first.count)
    return list(csv.DictReader(io.StringIO(text), delimiter=delim))


def col_of(rows: list, *candidates):
    if not rows:
        return None
    low = {k.lower().strip(): k for k in rows[0].keys()}
    for c in candidates:
        if c in low:
            return low[c]
    return None


def norm_url(url: str, domain: str = '') -> str:
    url = (url or '').strip()
    if not url:
        return ''
    if domain and url.startswith('/'):
        url = f'https://{domain}{url}'
    p = urlparse(url)
    host = p.netloc.lower()
    host = host[4:] if host.startswith('www.') else host
    path = p.path or '/'
    if path != '/' and path.endswith('/'):
        path = path[:-1]
    return f'{p.scheme or "https"}://{host}{path}'


def num(v, default=0.0):
    try:
        return float(str(v).replace(' ', '').replace(',', '.'))
    except (TypeError, ValueError):
        return default


def main():
    p = argparse.ArgumentParser(description='Анкор-план із вивантаження позицій')
    p.add_argument('positions_csv')
    p.add_argument('--anchors', help='CSV наявних анкорів (url, anchor)')
    p.add_argument('--domain', default='', help='домен для відносних URL')
    p.add_argument('--max-per-page', type=int, default=5)
    p.add_argument('--over-threshold', type=int, default=5,
                   help='анкор використано >= N разів -> уникати (за замовчуванням 5)')
    p.add_argument('--out', default='anchor_plan.xlsx')
    args = p.parse_args()

    rows = read_rows(args.positions_csv)
    url_c = col_of(rows, 'url', 'page', 'target url', 'current url')
    kw_c = col_of(rows, 'keyword', 'keywords', 'ключевое слово', 'ключова фраза')
    pos_c = col_of(rows, 'position', 'pos', 'current position', 'позиция', 'позиція')
    vol_c = col_of(rows, 'volume', 'search volume', 'global volume', 'частотность', 'частотність')
    if not url_c or not kw_c:
        raise SystemExit('У вивантаженні позицій не знайдено колонок URL і Keyword')

    by_url = defaultdict(list)
    for r in rows:
        u = norm_url(r.get(url_c, ''), args.domain)
        kw = (r.get(kw_c) or '').strip()
        if not u or not kw:
            continue
        by_url[u].append({'kw': kw, 'pos': num(r.get(pos_c)) if pos_c else 0.0,
                          'vol': num(r.get(vol_c)) if vol_c else 0.0})

    existing = defaultdict(Counter)
    if args.anchors:
        arows = read_rows(args.anchors)
        au_c = col_of(arows, 'url', 'target url', 'page')
        aa_c = col_of(arows, 'anchor', 'anchor text', 'анкор')
        if not au_c or not aa_c:
            raise SystemExit('У файлі анкорів не знайдено колонок url і anchor')
        for r in arows:
            u = norm_url(r.get(au_c, ''), args.domain)
            a = (r.get(aa_c) or '').strip()
            if u and a:
                existing[u][a] += 1

    rec_rows, avoid_rows = [], []
    for u, kws in by_url.items():
        kws.sort(key=lambda r: (r['pos'] if r['pos'] else 999, -r['vol']))
        seen = set()
        count = 0
        for r in kws:
            if count >= args.max_per_page:
                break
            kw = r['kw']
            if kw in seen:
                continue
            seen.add(kw)
            used = existing.get(u, Counter()).get(kw, 0)
            if used >= args.over_threshold:
                avoid_rows.append({'URL': u, 'Анкор': kw,
                                   'Причина': f'переспам: уже використано {used} разів'})
                continue
            prio = 'P1' if r['pos'] and r['pos'] <= 3 else ('P2' if r['pos'] and r['pos'] <= 10 else 'P3')
            rec_rows.append({'URL': u, 'Анкор': kw, 'Тип': 'точний (розбавити за правилами скілу)',
                             'Пріоритет': prio,
                             'Обґрунтування': f'позиція {r["pos"] or "-"}, частотність {int(r["vol"])}, '
                                            f'використано {used} разів'})
            count += 1

    with pd.ExcelWriter(args.out, engine='openpyxl') as xl:
        pd.DataFrame(rec_rows).to_excel(xl, sheet_name='Кандидати', index=False)
        pd.DataFrame(avoid_rows).to_excel(xl, sheet_name='Уникати', index=False)
        pd.DataFrame([
            {'Параметр': 'Сторінок', 'Значення': len(by_url)},
            {'Параметр': 'Кандидатів', 'Значення': len(rec_rows)},
            {'Параметр': 'Переспам (уникати)', 'Значення': len(avoid_rows)},
        ]).to_excel(xl, sheet_name='Зведення', index=False)
    print(f'кандидатів: {len(rec_rows)}, уникати: {len(avoid_rows)} -> {args.out}')


if __name__ == '__main__':
    main()
