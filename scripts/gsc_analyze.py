# -*- coding: utf-8 -*-
"""Аналитика выгрузки Google Search Console: quick wins, каннибализация, new/lost.

Логика - из замороженного analytics_pipeline.py (импортируется как модуль).
Вход: CSV с колонками (гибкие имена, RU/UA/EN): Page/URL, Query/Запрос, Clicks,
Impressions, CTR, Position. Каннибализация и new/lost требуют колонок URL и Query
одновременно (такой срез даёт GSC API / Looker Studio; UI-экспорт одной размерности
тоже принимается - тогда часть листов будет пустой).

Примеры:
  python scripts/gsc_analyze.py current.csv --out gsc.xlsx
  python scripts/gsc_analyze.py current.csv --previous prev.csv --out gsc.xlsx
"""
import argparse
import csv
import io
import os
import sys

import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import analytics_pipeline as ap  # noqa: E402  (замороженный модуль, только stdlib)

COLMAP = {
    'url': ['page', 'url', 'страница', 'страницы', 'сторінка', 'сторінки', 'landing page',
            'адрес страницы', 'top pages'],
    'query': ['query', 'queries', 'запрос', 'запросы', 'запит', 'запити', 'top queries',
              'search query', 'самые популярные запросы'],
    'clicks': ['clicks', 'клики', 'кліки'],
    'impressions': ['impressions', 'показы', 'покази'],
    'ctr': ['ctr', 'ctr страницы', 'ctr запроса'],
    'position': ['position', 'позиция', 'позиція', 'средняя позиция', 'середня позиція',
                 'avg. position', 'average position'],
}


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
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)

    def resolve(fieldnames):
        low = {f.lower().strip(): f for f in fieldnames if f}
        out = {}
        for target, variants in COLMAP.items():
            for v in variants:
                if v in low:
                    out[target] = low[v]
                    break
        return out

    cols = resolve(reader.fieldnames or [])

    def numval(v):
        s = str(v or '').replace('%', '').replace(' ', '').replace('\xa0', '').replace(',', '.')
        try:
            return float(s)
        except ValueError:
            return 0.0

    rows = []
    for r in reader:
        rows.append({
            'url': (r.get(cols['url']) or '').strip() if 'url' in cols else '',
            'query': (r.get(cols['query']) or '').strip() if 'query' in cols else '',
            'clicks': numval(r.get(cols['clicks'])) if 'clicks' in cols else 0.0,
            'impressions': numval(r.get(cols['impressions'])) if 'impressions' in cols else 0.0,
            'ctr': numval(r.get(cols['ctr'])) if 'ctr' in cols else None,
            'position': numval(r.get(cols['position'])) if 'position' in cols else 0.0,
        })
    print(f'{path}: рядків {len(rows)}, колонки: {sorted(cols.keys())}')
    return rows


def main():
    p = argparse.ArgumentParser(description='Аналіз вивантаження GSC')
    p.add_argument('current')
    p.add_argument('--previous', help='вивантаження минулого періоду (для нових і втрачених)')
    p.add_argument('--out', default='gsc_analysis.xlsx')
    p.add_argument('--min-impressions', type=float, default=10)
    p.add_argument('--pos-from', type=float, default=4, help='quick wins від позиції (за замовчуванням 4)')
    p.add_argument('--pos-to', type=float, default=20, help='quick wins до позиції (за замовчуванням 20)')
    args = p.parse_args()

    rows = read_rows(args.current)
    dash = ap.build_dashboard(rows)
    quick = ap.find_quick_wins(rows, min_position=args.pos_from, max_position=args.pos_to,
                               min_impressions=args.min_impressions)
    cann = ap.analyze_queries_cannibalization(rows)
    newlost = None
    if args.previous:
        prev = read_rows(args.previous)
        newlost = ap.analyze_queries_new_lost(prev, rows)

    with pd.ExcelWriter(args.out, engine='openpyxl') as xl:
        s = dash['summary']
        pd.DataFrame([
            {'Метрика': 'Кліки', 'Значення': s['total_clicks']},
            {'Метрика': 'Покази', 'Значення': s['total_impressions']},
            {'Метрика': 'CTR', 'Значення': s['avg_ctr']},
            {'Метрика': 'Середня позиція', 'Значення': s['avg_position']},
        ]).to_excel(xl, sheet_name='Зведення', index=False)
        qdf = pd.DataFrame(quick)
        if len(qdf):
            qdf['Дія'] = ''
        qdf.to_excel(xl, sheet_name='Quick wins', index=False)
        cdf = pd.DataFrame([{'query': c['query'], 'urls_count': c['urls_count'],
                             'urls': ' | '.join(c['urls']), 'Канонічний URL': ''} for c in cann])
        cdf.to_excel(xl, sheet_name='Канібалізація', index=False)
        if newlost:
            pd.DataFrame(newlost['new_keywords']).to_excel(xl, sheet_name='Нові', index=False)
            pd.DataFrame(newlost['lost_keywords']).to_excel(xl, sheet_name='Втрачені', index=False)
        pd.DataFrame(dash['top_pages']).to_excel(xl, sheet_name='Топ сторінок', index=False)
        pd.DataFrame(dash['top_queries']).to_excel(xl, sheet_name='Топ запитів', index=False)

    print(f'-> {args.out}')
    print(f"кліки {s['total_clicks']}, покази {s['total_impressions']}, "
          f"quick wins {len(quick)}, канібалізація {len(cann)}"
          + (f", нових {len(newlost['new_keywords'])} / втрачених {len(newlost['lost_keywords'])}"
             if newlost else ''))


if __name__ == '__main__':
    main()
