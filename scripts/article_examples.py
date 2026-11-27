# -*- coding: utf-8 -*-
"""Що Google показує на теми статей і які статті там є, насамперед у конкурентів (скіл semantic-clustering).

Бере теми статей з <out>.pages.json (semantic_core.py групує інфо-запити в теми), для кожної - топ-10 за головним
ключем (той самий кеш і джерела, що в serp_check.py: DataForSEO або Serpstat), рахує, які сторінки у видачі
(статті, відео, форуми, категорії, товари), і бере до --examples статей: спершу з доменів конкурентів із конфігу,
потім будь-які інші статті. Заголовки статей скрипт читає зі сторінок (кеш cache/pages/).

  python scripts/article_examples.py clients/shop/semantics/shop_semantics_2026-11-27.pages.json --config .../config.json
  python scripts/semantic_core.py .../config.json --articles .../articles.json
"""
import argparse
import hashlib
import html
import json
import os
import re
import sys
import time
from datetime import date

import requests

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from page_types import classify, excluded_host, host_of, serp_profile  # noqa: E402
from serp_check import EST_USD, Serp, cache_path, lang_of  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(ROOT, 'cache', 'pages')
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
TEXT_TYPES = ('стаття', 'вікі', 'форум')


def page_title(url):
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, hashlib.sha1(url.encode()).hexdigest()[:20] + '.title')
    if os.path.exists(path):
        return open(path, encoding='utf-8').read()
    title = ''
    try:
        r = requests.get(url, headers={'User-Agent': UA, 'Accept-Language': 'uk,ru;q=0.8'}, timeout=20)
        m = re.search(r'<title[^>]*>(.*?)</title>', r.text[:200000], re.S | re.I) if r.status_code == 200 else None
        if m:
            title = re.sub(r'\s+', ' ', html.unescape(m.group(1))).strip()[:160]
    except requests.RequestException:
        pass
    open(path, 'w', encoding='utf-8').write(title)
    time.sleep(0.7)
    return title


def pick_examples(items, n):
    """До n текстових сторінок з топу: спершу конкуренти, далі решта за позицією. Російські сайти не беремо."""
    texts = [x for x in items if x['type'] in TEXT_TYPES and not excluded_host(x['host'])]
    picked = [x for x in texts if x['competitor']][:n]
    return picked + [x for x in texts if x not in picked][:n - len(picked)]


def main():
    ap = argparse.ArgumentParser(description='Що у видачі на теми статей і приклади статей конкурентів')
    ap.add_argument('pages', help='<out>.pages.json від semantic_core.py')
    ap.add_argument('--config', help='конфіг ніші: домени конкурентів (competitors) йдуть у приклади першими')
    ap.add_argument('--examples', type=int, default=3, help='скільки прикладів статей на тему (3)')
    ap.add_argument('--source', choices=['dataforseo', 'serpstat'], default='dataforseo')
    ap.add_argument('--location', type=int, default=2804)
    ap.add_argument('--max-usd', type=float, default=1.0, help='ліміт витрат DataForSEO')
    ap.add_argument('--out', help='json (за замовчуванням articles.json поруч)')
    ap.add_argument('--dry-run', action='store_true', help='лише порахувати запити, без API')
    args = ap.parse_args()

    data = json.load(open(args.pages, encoding='utf-8'))
    topics = data.get('articles', [])
    need = [t['head'] for t in topics if not os.path.exists(cache_path(args.source, args.location, lang_of(t['head']), t['head']))]
    print(f'тем статей: {len(topics)}, головних ключів не в кеші: {len(need)}')
    if args.source == 'dataforseo':
        print(f'оцінка витрат: ${len(need) * EST_USD:.2f} і більше, якщо доведеться брати запасні ключі (ліміт ${args.max_usd:.2f})')
    if args.dry_run:
        return
    comp = set()
    if args.config:
        cfg = json.load(open(args.config, encoding='utf-8'))
        comp = {host_of(c['url'] if isinstance(c, dict) else c) for c in cfg.get('competitors', [])}
    serp = Serp(args.source, args.location, args.max_usd)
    result = {}
    for t in sorted(topics, key=lambda t: -t['volume']):
        key, top = serp.first_top(t)
        if top is None:
            print('ліміт витрат вичерпано: решту тем не перевірено')
            break
        profile, rec = serp_profile(top) if top else ({}, '')
        items = []
        for pos, it in enumerate(top, 1):
            title = it.get('title') or ''
            kind = classify(it['url'], title)
            items.append({'pos': pos, 'url': it['url'], 'host': host_of(it['url']), 'type': kind, 'title': title,
                          'competitor': host_of(it['url']) in comp})
        picked = pick_examples(items, args.examples)
        for x in picked:
            if not x['title']:
                x['title'] = page_title(x['url'])
        result[t['page']] = {'head': key, 'profile': profile, 'recommendation': rec, 'examples': picked, 'top': items}
        print(f'  {t["page"]}: {", ".join(f"{k} {v}" for k, v in profile.items())}; прикладів {len(picked)}')
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.pages)), 'articles.json')
    json.dump({'source': args.source, 'date': date.today().isoformat(), 'live_requests': serp.live,
               'cost_usd': round(serp.spent, 4), 'pages': result}, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'-> {out}')
    print(f'тем: {len(result)}, живих запитів: {serp.live}, витрачено ${serp.spent:.3f}')


if __name__ == '__main__':
    main()
