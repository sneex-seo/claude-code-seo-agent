# -*- coding: utf-8 -*-
"""Риск-скор ссылочного профиля из Ahrefs-выгрузки бэклинков.

Формула детерминированная (перенесена из agent_tasks/link_builder.yaml):
  DR < 10: +30 | DR 10-19: +15
  Domain traffic = 0: +25
  Referring domains < 40: +20
  Page traffic = 0 (у всех ссылок домена): +10
  Ahrefs "Is spam" = true (любая ссылка): +25
  все ссылки nofollow и DR > 30: -15 | все nofollow и DR < 10: +5
  DR > 30 и Domain traffic > 1000: -30
Категории: toxic >= --min-risk (60), suspicious 30..min-risk, good < 30.

Пример:
  python scripts/link_risk.py "clients/x/inbox/ahrefs.csv" --out links.xlsx --disavow disavow.txt
"""
import argparse
import csv
import io
import sys
from collections import defaultdict
from datetime import date
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
    delim = '\t' if text.splitlines()[0].count('\t') > text.splitlines()[0].count(',') else ','
    return list(csv.DictReader(io.StringIO(text), delimiter=delim))


def pick(row: dict, *names):
    low = {k.lower().strip(): k for k in row.keys()}
    for n in names:
        if n in low:
            return row.get(low[n])
    return None


def num(val, default=0.0) -> float:
    try:
        return float(str(val).replace(' ', '').replace(',', '.'))
    except (TypeError, ValueError):
        return default


def num_or_none(val):
    """Пустая ячейка выгрузки = «нет данных» (None), а не 0 - чтобы не штрафовать зря."""
    s = str(val).strip() if val is not None else ''
    if s == '':
        return None
    try:
        return float(s.replace(' ', '').replace(',', '.'))
    except ValueError:
        return None


def max_known(cur, new):
    if new is None:
        return cur
    return new if cur is None else max(cur, new)


def is_true(val) -> bool:
    return str(val).strip().lower() == 'true'


def domain_of(url: str) -> str:
    host = urlparse(str(url).strip()).netloc.lower()
    return host[4:] if host.startswith('www.') else host


def main():
    p = argparse.ArgumentParser(description='Ризик-скор беклінків (Ahrefs CSV)')
    p.add_argument('csv_path')
    p.add_argument('--min-risk', type=float, default=60, help='поріг toxic (за замовчуванням 60)')
    p.add_argument('--out', default='link_risk.xlsx')
    p.add_argument('--disavow', help='шлях до disavow.txt (лише toxic)')
    p.add_argument('--include-lost', action='store_true', help='не виключати lost-посилання')
    args = p.parse_args()

    rows = read_rows(args.csv_path)
    total_rows = len(rows)
    lost_excluded = 0
    per_domain = defaultdict(lambda: {
        'dr': 0.0, 'dtraf': None, 'rd': None, 'ptraf_max': None, 'links': 0,
        'nofollow': [], 'spam': False, 'titles': [], 'anchors': [],
    })

    for r in rows:
        if not args.include_lost and str(pick(r, 'lost status') or '').strip():
            lost_excluded += 1
            continue
        url = pick(r, 'referring page url') or ''
        dom = domain_of(url)
        if not dom:
            continue
        d = per_domain[dom]
        d['dr'] = max(d['dr'], num(pick(r, 'domain rating', 'dr')))
        d['dtraf'] = max_known(d['dtraf'], num_or_none(pick(r, 'domain traffic')))
        d['rd'] = max_known(d['rd'], num_or_none(pick(r, 'referring domains')))
        d['ptraf_max'] = max_known(d['ptraf_max'], num_or_none(pick(r, 'page traffic')))
        d['links'] += 1
        d['nofollow'].append(is_true(pick(r, 'nofollow')))
        d['spam'] = d['spam'] or is_true(pick(r, 'is spam'))
        title = str(pick(r, 'referring page title') or '').strip()
        anchor = str(pick(r, 'anchor') or '').strip()
        if title and title not in d['titles'] and len(d['titles']) < 3:
            d['titles'].append(title)
        if anchor and anchor not in d['anchors'] and len(d['anchors']) < 3:
            d['anchors'].append(anchor)

    out_rows = []
    for dom, d in per_domain.items():
        score = 0.0
        reasons = []
        if d['dr'] < 10:
            score += 30; reasons.append('DR<10')
        elif d['dr'] < 20:
            score += 15; reasons.append('DR<20')
        if d['dtraf'] is not None and d['dtraf'] == 0:
            score += 25; reasons.append('трафік домену 0')
        elif d['dtraf'] is None:
            reasons.append('немає даних: трафік домену')
        if d['rd'] is not None and d['rd'] < 40:
            score += 20; reasons.append('RD<40')
        elif d['rd'] is None:
            reasons.append('немає даних: RD')
        if d['ptraf_max'] is not None and d['ptraf_max'] == 0:
            score += 10; reasons.append('трафік сторінок 0')
        if d['spam']:
            score += 25; reasons.append('Ahrefs is_spam')
        all_nofollow = bool(d['nofollow']) and all(d['nofollow'])
        if all_nofollow and d['dr'] > 30:
            score -= 15; reasons.append('nofollow з доброго домену')
        if all_nofollow and d['dr'] < 10:
            score += 5; reasons.append('nofollow зі слабкого домену')
        if d['dr'] > 30 and (d['dtraf'] or 0) > 1000:
            score -= 30; reasons.append('DR>30 і живий трафік')
        score = max(0.0, min(100.0, score))
        cat = 'toxic' if score >= args.min_risk else ('suspicious' if score >= 30 else 'good')
        out_rows.append({
            'Домен': dom, 'DR': d['dr'], 'Трафік домену': d['dtraf'],
            'Referring domains': d['rd'], 'Посилань': d['links'],
            'Усі nofollow': all_nofollow, 'Is spam': d['spam'],
            'Ризик-скор': score, 'Категорія': cat,
            'Фактори': '; '.join(reasons),
            'Приклади title': ' || '.join(d['titles']),
            'Приклади анкорів': ' || '.join(d['anchors']),
            'Оцінка LLM': '', 'Коментар': '',
        })

    out_rows.sort(key=lambda x: -x['Ризик-скор'])
    df = pd.DataFrame(out_rows)
    counts = df['Категорія'].value_counts().to_dict() if len(df) else {}
    summary = pd.DataFrame([
        {'Параметр': 'Джерело', 'Значення': args.csv_path},
        {'Параметр': 'Дата аналізу', 'Значення': date.today().isoformat()},
        {'Параметр': 'Рядків на вході', 'Значення': total_rows},
        {'Параметр': 'Виключено lost', 'Значення': lost_excluded},
        {'Параметр': 'Унікальних доменів', 'Значення': len(out_rows)},
        {'Параметр': 'toxic (>= %.0f)' % args.min_risk, 'Значення': counts.get('toxic', 0)},
        {'Параметр': 'suspicious (30..)', 'Значення': counts.get('suspicious', 0)},
        {'Параметр': 'good (<30)', 'Значення': counts.get('good', 0)},
    ])
    with pd.ExcelWriter(args.out, engine='openpyxl') as xl:
        df.to_excel(xl, sheet_name='Домени', index=False)
        summary.to_excel(xl, sheet_name='Зведення', index=False)
    print(f'доменів: {len(out_rows)} -> {args.out}')
    print(f'на вході {total_rows} рядків, lost виключено {lost_excluded}, категорії: {counts}')

    if args.disavow:
        toxic = [r['Домен'] for r in out_rows if r['Категорія'] == 'toxic']
        with open(args.disavow, 'w', encoding='utf-8') as f:
            f.write(f'# Disavow {date.today().isoformat()} - чернетка, перевірте перед завантаженням\n')
            for dom in sorted(toxic):
                f.write(f'domain:{dom}\n')
        print(f'у disavow доменів: {len(toxic)} -> {args.disavow}')


if __name__ == '__main__':
    main()
