# -*- coding: utf-8 -*-
"""Перевірка видачі: чи потрібна детальна сторінка, чи її ключі краще лишити на батьківській.

Для кожної сторінки з <out>.pages.json (результат semantic_core.py), крім бази й брендів, береться
топ-10 Google за ключем сторінки і за ключем батьківської сторінки (головний ключ; якщо топу
немає - наступний за частотністю, до трьох спроб). Рішення:
  - спільних адрес >= --same (4): намір той самий, окрема сторінка не потрібна -> merge;
  - інакше, якщо >= --spec (3) результатів саме про уточнення (у адресі чи заголовку є його
    ознака: «2-kvt», «ac180», «dlya-domu»...) -> keep, сторінка потрібна;
  - інакше видача уточнення не виділяє -> merge (слабкий сигнал).

Джерело видачі - DataForSEO (живий топ з заголовками, ~$0.002 за запит) або Serpstat (топ з бази
g_ua, лише адреси). Відповіді кешуються в cache/serp/: повторний прогін не платить вдруге.

  python scripts/serp_check.py clients/shop/semantics/shop_semantics_2026-11-23.pages.json --dry-run
  python scripts/serp_check.py clients/shop/semantics/shop_semantics_2026-11-23.pages.json --max-usd 2
  python scripts/semantic_core.py clients/shop/semantics/config.json --serp clients/shop/semantics/serp_check.json
"""
import argparse
import hashlib
import json
import os
import re
import sys
from datetime import date
from urllib.parse import unquote, urlparse

import requests

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from serpstat_api import env  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(ROOT, 'cache', 'serp')
EST_USD = 0.002  # оцінка ціни живого запиту DataForSEO, якщо відповідь не назвала ціну
TRIES = 3
STOP = {'для', 'під', 'под', 'купити', 'купить', 'куплю', 'ціна', 'цена', 'ціни', 'цены', 'недорого', 'відгуки', 'отзывы', 'україна', 'украина', 'києві', 'киеве'}
TRANSLIT = str.maketrans({'а': 'a', 'б': 'b', 'в': 'v', 'г': 'h', 'ґ': 'g', 'д': 'd', 'е': 'e', 'є': 'ie', 'ж': 'zh',
                          'з': 'z', 'и': 'y', 'і': 'i', 'ї': 'i', 'й': 'i', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n',
                          'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'ts',
                          'ч': 'ch', 'ш': 'sh', 'щ': 'shch', 'ь': '', 'ю': 'iu', 'я': 'ia', 'ы': 'y', 'э': 'e',
                          'ъ': '', 'ё': 'e', "'": ''})


def lang_of(k):
    if re.search(r"[іїєґ']", k):
        return 'uk'
    if re.search(r'[ыэъё]', k):
        return 'ru'
    return 'uk'


def norm_url(u):
    p = urlparse(u.strip())
    host = p.netloc.lower()
    host = host[4:] if host.startswith('www.') else host
    return host + (p.path.rstrip('/') or '/')


def flat(s):
    return re.sub(r'[^a-z0-9а-яіїєґ]+', '', unquote(s).lower())


def refinement_tokens(page, parent):
    """Ознаки уточнення: чим сторінка відрізняється від батьківської (для пошуку в адресах і заголовках)."""
    rest = page['url'][len(parent['url']):] if page['url'].startswith(parent['url']) else page['url']
    slug = rest.strip('/').split('/')[-1] if rest.strip('/') else ''
    toks = set()
    m = re.fullmatch(r'(\d+)(?:-(\d+))?-kvt', slug)
    if m:
        kw = float(m.group(1) + ('.' + m.group(2) if m.group(2) else ''))
        n = m.group(1) + (m.group(2) or '')
        toks |= {f'{round(kw * 1000)}', n + 'kvt', n + 'kw', n + 'квт'}
    elif re.fullmatch(r'\d+-vt', slug):
        n = slug.split('-')[0]
        toks |= {n + 'w', n + 'vt', n + 'вт', n + 'wt'}
    elif re.fullmatch(r'\d+-vt-hod', slug):
        n = slug.split('-')[0]
        toks |= {n + 'wh', n + 'vtgod', n + 'втгод', n + 'втч'}
    elif slug:
        toks.add(flat(slug))
        toks |= {p for p in slug.split('-') if len(p) >= 4}
    # слова головного ключа, яких немає в ключі батьківської сторінки: «кабель» -> «kabe», «дома» -> «dom»
    pstems = {w[:4] for w in re.findall(r'[a-zа-яіїєґ0-9]+', parent['head'].lower())}
    for w in re.findall(r'[a-zа-яіїєґ0-9]+', page['head'].lower()):
        if len(w) >= 3 and w[:4] not in pstems and w not in STOP:
            stem = w[:max(3, len(w) - 2)]
            toks |= {stem, stem.translate(TRANSLIT)}
    label = page['page'].split(' > ')[-1].lower()
    for w in re.findall(r'[a-zа-яіїєґ0-9]+', label):
        if len(w) >= 4 and w not in {'для', 'плюс'}:
            stem = w[:max(4, len(w) - 2)]
            toks |= {stem, stem.translate(TRANSLIT)}
    if re.search(r'\d', label):
        toks.add(flat(label))
    return {t for t in toks if len(t) >= 3}


def decide(top, parent_top, toks, same=4, spec_min=3):
    """(рішення, пояснення, спільних, про уточнення) для топу сторінки проти топу батьківської."""
    if not top or not parent_top:
        return 'keep', 'немає топу - не перевірено', 0, 0
    spec = sum(1 for it in top if any(t in flat(it['url']) or t in it.get('title', '').lower() for t in toks))
    overlap = len({norm_url(it['url']) for it in top} & {norm_url(it['url']) for it in parent_top})
    if overlap >= same:
        return 'merge', f'спільних з батьківською {overlap} з 10', overlap, spec
    if spec >= spec_min:
        return 'keep', f'про уточнення {spec} з 10, спільних {overlap}', overlap, spec
    return 'merge', f'видача не виділяє уточнення: про нього {spec} з 10, спільних {overlap}', overlap, spec


def cache_path(source, location, lang, k):
    h = hashlib.sha1(k.encode('utf-8')).hexdigest()[:16]
    return os.path.join(CACHE, f'{source}_{location}_{lang}_{h}.json')


def fetch_dataforseo(k, location, lang):
    body = [{'keyword': k, 'location_code': location, 'language_code': lang, 'depth': 10}]
    r = requests.post('https://api.dataforseo.com/v3/serp/google/organic/live/regular',
                      auth=(env('DATAFORSEO_LOGIN'), env('DATAFORSEO_PASSWORD')), json=body, timeout=120).json()
    if r.get('status_code') != 20000:
        raise SystemExit(f'DataForSEO: {r.get("status_code")} {r.get("status_message")}')
    task = r['tasks'][0]
    if task.get('status_code') != 20000:
        raise SystemExit(f'DataForSEO: {task.get("status_code")} {task.get("status_message")}')
    items = ((task.get('result') or [{}])[0].get('items')) or []
    top = [{'url': it['url'], 'title': it.get('title') or ''} for it in items if it.get('type') == 'organic' and it.get('url')]
    return top[:10], float(task.get('cost') or EST_USD)


def fetch_serpstat(k, location, lang):
    body = {'id': '1', 'method': 'SerpstatKeywordProcedure.getKeywordTop',
            'params': {'keyword': k, 'se': 'g_ua', 'filters': {'top_size': 10}}}
    r = requests.post('https://api.serpstat.com/v4/', params={'token': env('SERPSTAT_TOKEN')}, json=body, timeout=120).json()
    if 'error' in r:
        return [], 0.0  # ключа немає в базі Serpstat
    res = r.get('result', {})
    top = (res.get('data') or {}).get('top') or res.get('top') or []
    return [{'url': t['url'], 'title': t.get('title') or ''} for t in top if t.get('url')][:10], 0.0


class Serp:
    def __init__(self, source, location, max_usd):
        self.source, self.location, self.max_usd = source, location, max_usd
        self.fetch = fetch_dataforseo if source == 'dataforseo' else fetch_serpstat
        self.spent, self.live, self.stopped = 0.0, 0, False

    def top(self, k):
        path = cache_path(self.source, self.location, lang_of(k), k)
        if os.path.exists(path):
            d = json.load(open(path, encoding='utf-8'))
            return d.get('top') or [{'url': u, 'title': ''} for u in d.get('urls', [])]
        if self.stopped or (self.source == 'dataforseo' and self.spent + EST_USD > self.max_usd):
            self.stopped = True
            return None
        top, cost = self.fetch(k, self.location, lang_of(k))
        self.spent += cost
        self.live += 1
        json.dump({'keyword': k, 'top': top, 'source': self.source, 'date': date.today().isoformat()},
                  open(path, 'w', encoding='utf-8'), ensure_ascii=False)
        if self.live % 25 == 0:
            print(f'  запитів {self.live}, витрачено ${self.spent:.3f}', flush=True)
        return top

    def first_top(self, page):
        """Топ за головним ключем, а якщо його немає - за наступним за частотністю."""
        keys = [page['head']] + [k for k, _ in sorted(page['keywords'], key=lambda x: -x[1]) if k != page['head']]
        for k in keys[:TRIES]:
            t = self.top(k)
            if t is None:
                return None, None
            if t:
                return k, t
        return keys[0], []


def main():
    ap = argparse.ArgumentParser(description='Перевірка видачі: окрема сторінка чи злити з батьківською')
    ap.add_argument('pages', help='<out>.pages.json від semantic_core.py')
    ap.add_argument('--out', help='json з рішеннями (за замовчуванням serp_check.json поруч)')
    ap.add_argument('--same', type=int, default=4, help='спільних адрес у топ-10, з яких намір той самий (4)')
    ap.add_argument('--spec', type=int, default=3, help='результатів саме про уточнення, з яких сторінка потрібна (3)')
    ap.add_argument('--source', choices=['dataforseo', 'serpstat'], default='dataforseo')
    ap.add_argument('--location', type=int, default=2804, help='location_code DataForSEO (2804 - Україна)')
    ap.add_argument('--max-usd', type=float, default=2.0, help='ліміт витрат DataForSEO на прогін')
    ap.add_argument('--dry-run', action='store_true', help='лише порахувати запити, без API')
    args = ap.parse_args()

    data = json.load(open(args.pages, encoding='utf-8'))
    pages = {p['page']: p for p in data['pages']}
    # перевіряємо і поточні сторінки, і злиті раніше: їхній склад міг змінитися
    todo = [p for p in list(pages.values()) + data.get('merged', []) if p['kind'] not in ('base', 'brand') and p['parent'] in pages]
    os.makedirs(CACHE, exist_ok=True)
    heads = {p['head'] for p in todo} | {pages[p['parent']]['head'] for p in todo}
    need = [k for k in heads if not os.path.exists(cache_path(args.source, args.location, lang_of(k), k))]
    print(f'сторінок до перевірки: {len(todo)}, головних ключів: {len(heads)}, з них не в кеші: {len(need)}')
    if args.source == 'dataforseo':
        print(f'оцінка витрат: ${len(need) * EST_USD:.2f} і більше, якщо доведеться брати запасні ключі (ліміт ${args.max_usd:.2f})')
    if args.dry_run:
        return

    serp = Serp(args.source, args.location, args.max_usd)
    result = {}
    for p in sorted(todo, key=lambda p: -p['volume']):
        parent = pages[p['parent']]
        pk, a = serp.first_top(p)
        qk, b = serp.first_top(parent)
        if a is None or b is None:
            break
        decision, why, overlap, spec = decide(a, b, refinement_tokens(p, parent), args.same, args.spec)
        result[p['page']] = {'head': pk, 'parent': parent['page'], 'parent_head': qk, 'overlap': overlap,
                             'specific': spec, 'decision': decision, 'why': why,
                             'top': [it['url'] for it in a], 'parent_top': [it['url'] for it in b]}
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.pages)), 'serp_check.json')
    json.dump({'source': args.source, 'same': args.same, 'spec': args.spec, 'date': date.today().isoformat(),
               'live_requests': serp.live, 'cost_usd': round(serp.spent, 4), 'pages': result},
              open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    merged = [g for g, d in result.items() if d['decision'] == 'merge']
    unknown = [g for g, d in result.items() if d['why'].startswith('немає топу')]
    print(f'-> {out}')
    print(f'перевірено сторінок: {len(result)} з {len(todo)}; окремих: {len(result) - len(merged)} '
          f'(з них без топу: {len(unknown)}), злити з батьківською: {len(merged)}')
    print(f'живих запитів: {serp.live}, витрачено ${serp.spent:.3f}')
    if serp.stopped:
        print(f'ліміт ${args.max_usd:.2f} вичерпано: решту сторінок не перевірено')


if __name__ == '__main__':
    main()
