# -*- coding: utf-8 -*-
"""DataForSEO API - хелпер для скиллов SEO-агента.

Креды: DATAFORSEO_LOGIN / DATAFORSEO_PASSWORD в .env в корне проекта.

Примеры:
  python scripts/dataforseo_api.py balance
  python scripts/dataforseo_api.py volume "кава в зернах" "кофе в зернах" --location 2804 --lang uk --out vol.csv
  python scripts/dataforseo_api.py serp "купити каву" --location 2804 --lang uk --depth 10

Коды локаций: Украина 2804, США 2840, Канада 2124, Польша 2616, Германия 2276.
"""
import argparse
import csv
import json
import os
import sys

import requests

sys.stdout.reconfigure(encoding='utf-8')
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = 'https://api.dataforseo.com'


def env(key: str) -> str:
    path = os.path.join(ROOT, '.env')
    if os.path.exists(path):
        for line in open(path, encoding='utf-8'):
            if line.strip().startswith(key + '='):
                return line.split('=', 1)[1].strip()
    val = os.environ.get(key)
    if val:
        return val
    raise SystemExit(f'{key} не знайдено ні в .env, ні в змінних оточення')


def auth():
    return (env('DATAFORSEO_LOGIN'), env('DATAFORSEO_PASSWORD'))


def post(path: str, payload: list) -> list:
    resp = requests.post(BASE + path, auth=auth(), json=payload, timeout=120)
    data = resp.json()
    if data.get('status_code') != 20000:
        raise SystemExit(f'DataForSEO error: {data.get("status_code")} {data.get("status_message")}')
    task = data['tasks'][0]
    if task.get('status_code') != 20000:
        raise SystemExit(f'DataForSEO task error: {task.get("status_code")} {task.get("status_message")}')
    return task.get('result') or []


def write_csv(path: str, header: list, rows: list) -> None:
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f'рядків: {len(rows)} -> {path}')


def cmd_balance(args):
    resp = requests.get(BASE + '/v3/appendix/user_data', auth=auth(), timeout=60).json()
    result = resp['tasks'][0]['result'][0]
    print(json.dumps({'money': result.get('money'), 'rates': result.get('rates', {}).get('limits')},
                     ensure_ascii=False, indent=2))


def cmd_volume(args):
    payload = [{'keywords': args.keywords, 'location_code': args.location, 'language_code': args.lang}]
    result = post('/v3/keywords_data/google_ads/search_volume/live', payload)
    rows = [[i.get('keyword'), i.get('search_volume'), i.get('cpc'), i.get('competition')]
            for i in result]
    rows.sort(key=lambda r: -(r[1] or 0))
    if args.out:
        write_csv(args.out, ['keyword', 'search_volume', 'cpc', 'competition'], rows)
    else:
        for kw, vol, cpc, comp in rows:
            print(f'{vol!s:>8}  {kw}')


def cmd_serp(args):
    payload = [{'keyword': args.keyword, 'location_code': args.location,
                'language_code': args.lang, 'depth': args.depth}]
    result = post('/v3/serp/google/organic/live/regular', payload)
    items = (result[0].get('items') or []) if result else []
    rows = [[it.get('rank_absolute'), it.get('title'), it.get('url')]
            for it in items if it.get('type') == 'organic']
    if args.out:
        write_csv(args.out, ['rank', 'title', 'url'], rows)
    else:
        for rank, title, url in rows:
            print(f'{rank:>3}. {title}\n     {url}')


def main():
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument('--location', type=int, default=2804, help='location_code (за замовчуванням 2804, Україна)')
    common.add_argument('--lang', default='uk', help='language_code (за замовчуванням uk)')

    p = argparse.ArgumentParser(description='DataForSEO API helper')
    sub = p.add_subparsers(dest='cmd', required=True)

    sub.add_parser('balance', help='баланс акаунта (безкоштовно)')

    v = sub.add_parser('volume', parents=[common], help='частотності Google Ads за списком ключів')
    v.add_argument('keywords', nargs='+')
    v.add_argument('--out', help='шлях до csv')

    s = sub.add_parser('serp', parents=[common], help='органічна видача за ключем')
    s.add_argument('keyword')
    s.add_argument('--depth', type=int, default=10)
    s.add_argument('--out', help='шлях до csv')

    args = p.parse_args()
    {'balance': cmd_balance, 'volume': cmd_volume, 'serp': cmd_serp}[args.cmd](args)


if __name__ == '__main__':
    main()
