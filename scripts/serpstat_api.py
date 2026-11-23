# -*- coding: utf-8 -*-
"""Serpstat API v4 - хелпер для скиллов SEO-агента.

Токен: SERPSTAT_TOKEN в .env в корне проекта (или в переменных окружения).

Примеры:
  python scripts/serpstat_api.py limits
  python scripts/serpstat_api.py keywords "купити каву" --min-vol 10 --out kw.csv
  python scripts/serpstat_api.py domain example.com.ua
  python scripts/serpstat_api.py domain-keywords example.com.ua --max-pages 5 --out dk.csv

База по умолчанию: g_ua (Google Украина).
"""
import argparse
import csv
import json
import os
import sys
import time

import requests

sys.stdout.reconfigure(encoding='utf-8')
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
API_URL = 'https://api.serpstat.com/v4/'


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


def call(method: str, params: dict) -> dict:
    body = {'id': '1', 'method': method, 'params': params}
    resp = requests.post(API_URL, params={'token': env('SERPSTAT_TOKEN')}, json=body, timeout=120)
    data = resp.json()
    if 'error' in data:
        raise SystemExit(f'Serpstat error: {json.dumps(data["error"], ensure_ascii=False)}')
    return data.get('result', {})


def write_csv(path: str, header: list, rows: list) -> None:
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f'рядків: {len(rows)} -> {path}')


def load_minus(path: str) -> list:
    return [l.strip().lower() for l in open(path, encoding='utf-8')
            if l.strip() and not l.startswith('#')]


def is_minus(keyword: str, minus_words: list) -> bool:
    """Однословный минус - совпадение с токеном ключа; многословный - подстрока."""
    kw = keyword.lower()
    tokens = set(kw.split())
    for m in minus_words:
        if ' ' in m:
            if m in kw:
                return True
        elif m in tokens:
            return True
    return False


def apply_minus(rows_dict: dict, minus_file: str) -> dict:
    if not minus_file:
        return rows_dict
    minus = load_minus(minus_file)
    kept = {k: v for k, v in rows_dict.items() if not is_minus(k, minus)}
    print(f'мінус-фільтр: {len(rows_dict) - len(kept)} відсіяно, {len(kept)} залишилось')
    return kept


def cmd_limits(args):
    print(json.dumps(call('SerpstatLimitsProcedure.getStats', {}), ensure_ascii=False, indent=2))


def cmd_keywords(args):
    rows = {}
    page = 1
    while page <= args.max_pages:
        params = {
            'keyword': args.seed, 'se': args.se, 'page': page, 'size': 1000,
            'sort': {'region_queries_count': 'desc'},
            'filters': {'region_queries_count_from': args.min_vol},
        }
        data = call('SerpstatKeywordProcedure.getKeywords', params).get('data', [])
        if not data:
            break
        for r in data:
            rows[r['keyword'].lower()] = r.get('region_queries_count') or 0
        print(f'  стор. {page}: +{len(data)} (усього {len(rows)})', flush=True)
        if len(data) < 1000:
            break
        page += 1
        time.sleep(1.05)
    rows = apply_minus(rows, args.minus_file)
    out = args.out or 'serpstat_keywords.csv'
    write_csv(out, ['keyword', 'volume'], sorted(rows.items(), key=lambda x: -x[1]))


def cmd_url_keywords(args):
    """Ключи, по которым ранжируется конкретный URL (свой или конкурента)."""
    rows = {}
    page = 1
    while page <= args.max_pages:
        params = {'url': args.url, 'se': args.se, 'page': page, 'size': 1000,
                  'sort': {'region_queries_count': 'desc'}}
        data = call('SerpstatUrlProcedure.getUrlKeywords', params).get('data', [])
        if not data:
            break
        for r in data:
            kw = (r.get('keyword') or '').lower()
            rows[kw] = (r.get('region_queries_count') or 0, r.get('position'))
        print(f'  стор. {page}: +{len(data)} (усього {len(rows)})', flush=True)
        if len(data) < 1000:
            break
        page += 1
        time.sleep(1.05)
    if args.minus_file:
        minus = load_minus(args.minus_file)
        before = len(rows)
        rows = {k: v for k, v in rows.items() if not is_minus(k, minus)}
        print(f'мінус-фільтр: {before - len(rows)} відсіяно, {len(rows)} залишилось')
    out = args.out or 'serpstat_url_keywords.csv'
    write_csv(out, ['keyword', 'volume', 'position'],
              [[k, v[0], v[1]] for k, v in sorted(rows.items(), key=lambda x: -x[1][0])])


def cmd_domain(args):
    print(json.dumps(
        call('SerpstatDomainProcedure.getDomainsInfo', {'domains': [args.domain], 'se': args.se}),
        ensure_ascii=False, indent=2))


def cmd_top(args):
    result = call('SerpstatKeywordProcedure.getKeywordTop',
                  {'keyword': args.keyword, 'se': args.se, 'filters': {'top_size': args.size}})
    top = (result.get('data') or {}).get('top') or result.get('top') or []
    if not top:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    rows = [[t.get('position'), t.get('url'), t.get('title') or ''] for t in top]
    if args.out:
        write_csv(args.out, ['position', 'url', 'title'], rows)
    else:
        for pos, url, title in rows:
            print(f'{pos:>3}. {title}\n     {url}' if title else f'{pos:>3}. {url}')


def cmd_domain_keywords(args):
    rows = []
    page = 1
    while page <= args.max_pages:
        params = {'domain': args.domain, 'se': args.se, 'page': page, 'size': 1000,
                  'sort': {'region_queries_count': 'desc'}}
        data = call('SerpstatDomainProcedure.getDomainKeywords', params).get('data', [])
        if not data:
            break
        for r in data:
            rows.append([r.get('keyword'), r.get('region_queries_count') or 0,
                         r.get('position'), r.get('url')])
        print(f'  стор. {page}: +{len(data)} (усього {len(rows)})', flush=True)
        if len(data) < 1000:
            break
        page += 1
        time.sleep(1.05)
    out = args.out or 'serpstat_domain_keywords.csv'
    write_csv(out, ['keyword', 'volume', 'position', 'url'], rows)


def main():
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument('--se', default='g_ua', help='пошукова база Serpstat (за замовчуванням g_ua)')

    p = argparse.ArgumentParser(description='Serpstat API helper')
    sub = p.add_subparsers(dest='cmd', required=True)

    sub.add_parser('limits', parents=[common], help='ліміти акаунта (безкоштовний метод)')

    k = sub.add_parser('keywords', parents=[common], help='підбір ключів за сід-фразою')
    k.add_argument('seed')
    k.add_argument('--min-vol', type=int, default=10, help='мін. частотність (за замовчуванням 10)')
    k.add_argument('--max-pages', type=int, default=70)
    k.add_argument('--minus-file', help='файл мінус-слів (по слову в рядку)')
    k.add_argument('--out', help='шлях до csv (за замовчуванням serpstat_keywords.csv)')

    uk = sub.add_parser('url-keywords', parents=[common],
                        help='ключі, за якими ранжується URL (свій або конкурента)')
    uk.add_argument('url')
    uk.add_argument('--max-pages', type=int, default=10)
    uk.add_argument('--minus-file', help='файл мінус-слів (по слову в рядку)')
    uk.add_argument('--out', help='шлях до csv (за замовчуванням serpstat_url_keywords.csv)')

    d = sub.add_parser('domain', parents=[common], help='зведення по домену (видимість, ключі)')
    d.add_argument('domain')

    t = sub.add_parser('top', parents=[common], help='SERP-топ за ключем із бази Serpstat')
    t.add_argument('keyword')
    t.add_argument('--size', type=int, default=10, help='розмір топу (за замовчуванням 10)')
    t.add_argument('--out', help='шлях до csv')

    dk = sub.add_parser('domain-keywords', parents=[common], help='ключі домену з позиціями')
    dk.add_argument('domain')
    dk.add_argument('--max-pages', type=int, default=10)
    dk.add_argument('--out', help='шлях до csv (за замовчуванням serpstat_domain_keywords.csv)')

    args = p.parse_args()
    {'limits': cmd_limits, 'keywords': cmd_keywords, 'top': cmd_top,
     'url-keywords': cmd_url_keywords,
     'domain': cmd_domain, 'domain-keywords': cmd_domain_keywords}[args.cmd](args)


if __name__ == '__main__':
    main()
