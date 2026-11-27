# -*- coding: utf-8 -*-
"""Структура конкурентів: які сторінки розділу в них є і які з них ранжуються (скіл semantic-clustering, Крок 1).

На вхід - адреси категорій конкурентів (у конфігу ніші `competitors` або в аргументах). Для кожного сайту:
  1. Serpstat: сторінки домену, в адресі яких є слаг категорії, і скільки ключів у кожної (getDomainUrls);
  2. HTML категорії: посилання всередині розділу з текстом анкора - фільтри й підкатегорії (--no-fetch вимикає);
  3. тип кожної сторінки (категорія, фільтр, товар, стаття) і назва - з анкора або з кінця адреси;
  4. ознаки, які повторюються в кількох конкурентів (бренди, потужності, тип батареї) - підказка для конфігу ніші;
  5. з --collect N: ключі з N найсильніших фільтрів і категорій (getUrlKeywords) у inbox/ - другий канал збору.

  python scripts/competitors.py clients/shop/semantics/config.json
  python scripts/competitors.py clients/shop/semantics/config.json --collect 20
  python scripts/competitors.py --urls https://site.ua/cat/ https://other.ua/category/ --out comp.xlsx
"""
import argparse
import csv
import hashlib
import html
import json
import os
import re
import sys
import time
from collections import Counter, defaultdict
from urllib.parse import unquote, urljoin, urlparse

import openpyxl
import requests
from openpyxl.styles import Font

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from page_types import classify, host_of  # noqa: E402
from serpstat_api import call  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(ROOT, 'cache', 'pages')
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
LANG = {'ua', 'uk', 'ru', 'en', 'pl'}
STOP = {'ua', 'uk', 'ru', 'catalog', 'category', 'kategoriya', 'c', 'www', 'com', 'html', 'php', 'page', 'filter', 'f'}


def section_slug(url):
    """Слаг розділу: останній змістовний сегмент шляху (не мова і не «c4674585»)."""
    segs = [s for s in urlparse(url).path.split('/') if s]
    segs = [s for s in segs if s.lower() not in LANG and not re.fullmatch(r'c\d+', s)]
    return segs[-1] if segs else ''


def slug_stem(slug):
    return re.sub(r'-?\d{3,}$', '', slug)


def fetch(url):
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, hashlib.sha1(url.encode()).hexdigest()[:20] + '.html')
    if os.path.exists(path):
        return open(path, encoding='utf-8', errors='ignore').read()
    try:
        r = requests.get(url, headers={'User-Agent': UA, 'Accept-Language': 'uk,ru;q=0.8'}, timeout=30)
    except requests.RequestException as e:
        print(f'  {url}: не завантажилось ({type(e).__name__})')
        return ''
    if r.status_code != 200 or 'Just a moment' in r.text[:3000]:
        print(f'  {url}: код {r.status_code}, HTML не розбираємо (захист від ботів або помилка)')
        return ''
    open(path, 'w', encoding='utf-8').write(r.text)
    time.sleep(1.0)
    return r.text


def text_of(fragment):
    t = html.unescape(re.sub(r'<[^>]+>', ' ', fragment))
    return re.sub(r'\s+', ' ', t).strip()


PAGER = re.compile(r'(page|p)[-=]?\d+', re.I)
LANG_LABEL = re.compile(r'^\W*(ru|ua|uk|en|рус|укр|ру|eng)\b[\W\w]{0,12}$', re.I)


def segments(url):
    """Сегменти шляху без мовних префіксів і службових «c4674585»: так /ru/cat/ і /ua/cat/ - одна сторінка."""
    return [s for s in unquote(urlparse(url).path).split('/') if s and s.lower() not in LANG and not re.fullmatch(r'c\d+', s)]


def page_key(url):
    q = urlparse(url).query
    return '/'.join(segments(url)) + ('?' + q if q else '')


def label_from_url(url, stem):
    segs = segments(url)
    tail = segs[-1] if segs else ''
    tail = tail.replace(stem, '').strip('-_')
    if not tail and urlparse(url).query:
        tail = urlparse(url).query
    tail = re.sub(r'(\w+)[=:]', r'\1: ', tail)
    return re.sub(r'[-_;&]+', ' ', tail).strip()


def section_type(url, slug, stem):
    """Місце сторінки відносно розділу: сама категорія (інша мова чи шлях), фільтр усередині або суміжна категорія."""
    segs, q = segments(url), urlparse(url).query
    if (segs and PAGER.fullmatch(segs[-1])) or re.search(r'(^|&)(page|p)=\d+', q):
        return 'пагінація'
    if slug in segs:
        return 'категорія' if segs[-1] == slug and not q else 'фільтр'
    last = segs[-1] if segs else ''
    if stem and last.startswith(stem + '-'):
        return 'фільтр'  # zaryadnye-stantsii-ecoflow - бренд як окрема адреса
    return 'суміжна категорія'  # poverbanki-i-zaryadnie-stantsii - ширша або сусідня категорія


def html_links(cat_url, page, stem):
    host = host_of(cat_url)
    base_path = urlparse(cat_url).path.rstrip('/')
    out = {}
    for href, inner in re.findall(r'<a\b[^>]*href="([^"#]+)"[^>]*>(.*?)</a>', page, re.S | re.I):
        u = urljoin(cat_url, html.unescape(href))
        p = urlparse(u)
        if host_of(u) != host or u.rstrip('/') == cat_url.rstrip('/'):
            continue
        inside = p.path.startswith(base_path + '/') or (stem and stem in p.path)
        if not inside:
            continue
        label = text_of(inner)[:80]
        if u not in out or (label and len(label) < len(out[u] or 'x' * 99)):
            out[u] = label
    return out


def domain_urls(host, slug, max_pages):
    rows, page = {}, 1
    while page <= max_pages:
        try:
            data = call('SerpstatDomainProcedure.getDomainUrls',
                        {'domain': host, 'se': 'g_ua', 'page': page, 'size': 1000, 'filters': {'url_contain': slug}}).get('data', [])
        except SystemExit as e:  # домену немає в базі Serpstat - лишаємо лише HTML
            print(f'  Serpstat: {e}')
            break
        for d in data:
            rows[d['url']] = d.get('keywords', 0)
        if len(data) < 1000:
            break
        page += 1
    return rows


def analyse(cat_url, fetch_html=True, max_pages=1):
    host = host_of(cat_url)
    slug = section_slug(cat_url)
    stem = slug_stem(slug)
    print(f'{host}: розділ «{slug}»')
    ranked = domain_urls(host, stem or slug, max_pages) if stem or slug else {}
    page = fetch(cat_url) if fetch_html else ''
    title = text_of(re.search(r'<title>(.*?)</title>', page, re.S).group(1)) if page and re.search(r'<title>', page) else ''
    h1m = re.search(r'<h1[^>]*>(.*?)</h1>', page, re.S | re.I) if page else None
    links = html_links(cat_url, page, stem) if page else {}
    # мовні версії однієї сторінки (/ru/..., /ua/...) зводимо в один рядок, ключі сумуємо
    groups = defaultdict(list)
    for u in set(ranked) | set(links):
        groups[page_key(u)].append(u)
    pages = []
    for key, urls in sorted(groups.items()):
        urls.sort(key=lambda u: (u not in links, -ranked.get(u, 0)))
        u = urls[0]
        labels = [re.sub(r'\s*\(\d+\)$', '', links[x]) for x in urls if links.get(x) and not LANG_LABEL.match(links[x])]
        kind = classify(u, category_url=cat_url)
        if kind in ('категорія', 'маркетплейс', 'фільтр'):
            kind = section_type(u, slug, stem)
        pages.append({'url': u, 'type': kind, 'label': labels[0] if labels else label_from_url(u, stem) or slug,
                      'keywords': sum(ranked.get(x, 0) for x in urls), 'versions': len(urls),
                      'source': '+'.join(s for s, ok in (('Serpstat', any(x in ranked for x in urls)),
                                                         ('HTML', any(x in links for x in urls))) if ok)})
    by = Counter(p['type'] for p in pages)
    print(f'  сторінок: {len(pages)} ({", ".join(f"{k} {v}" for k, v in by.most_common())}), з HTML: {len(links)}')
    return {'competitor': host, 'category_url': cat_url, 'slug': slug, 'title': title, 'h1': text_of(h1m.group(1)) if h1m else '',
            'pages': pages, 'types': dict(by)}


def norm_token(t):
    t = t.lower().replace('ё', 'е')
    return re.sub(r'[^a-zа-яіїєґ0-9 ]+', ' ', t).strip()


def facet_candidates(results):
    """Слова й фрази з назв фільтрів, які трапляються в кількох конкурентів."""
    seen = defaultdict(set)
    examples = defaultdict(list)
    kw = Counter()
    for r in results:
        for p in r['pages']:
            if p['type'] not in ('фільтр', 'категорія') or not p['label']:
                continue
            label = norm_token(p['label'])
            for t in set(label.split()):
                if len(t) < 3 or t in STOP or t.isdigit():
                    continue
                seen[t].add(r['competitor'])
                kw[t] += p['keywords']
                if len(examples[t]) < 3 and p['label'] not in examples[t]:
                    examples[t].append(p['label'])
    rows = [(t, len(c), kw[t], examples[t]) for t, c in seen.items() if len(c) >= 2]
    return sorted(rows, key=lambda x: (-x[1], -x[2]))


def collect_keywords(results, n, inbox):
    """url-keywords з N найсильніших фільтрів і категорій конкурентів - у inbox/ як канал збору."""
    cand = [(p['keywords'], r['competitor'], p['url']) for r in results for p in r['pages']
            if p['type'] in ('фільтр', 'категорія') and p['keywords'] >= 20]
    cand.sort(reverse=True)
    os.makedirs(inbox, exist_ok=True)
    done = 0
    for kwn, host, url in cand[:n]:
        name = f'url_{host.split(".")[0]}_{hashlib.sha1(url.encode()).hexdigest()[:6]}.csv'
        path = os.path.join(inbox, name)
        if os.path.exists(path):
            continue
        data = call('SerpstatUrlProcedure.getUrlKeywords', {'url': url, 'se': 'g_ua', 'page': 1, 'size': 1000,
                                                            'sort': {'region_queries_count': 'desc'}}).get('data', [])
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(['keyword', 'volume', 'position'])
            for d in data:
                w.writerow([(d.get('keyword') or '').lower(), d.get('region_queries_count') or 0, d.get('position')])
        done += 1
        print(f'  ключі {url}: {len(data)} -> {name}')
        time.sleep(1.05)
    return done


def write_xlsx(results, facets, out):
    bold = Font(bold=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Конкуренти'
    ws.append(['Конкурент', 'Категорія', 'Title', 'H1', 'Сторінок', 'Фільтрів і підкатегорій', 'Товарів', 'Статей', 'Найсильніші фільтри'])
    for r in results:
        top = sorted([p for p in r['pages'] if p['type'] == 'фільтр'], key=lambda p: -p['keywords'])[:5]
        ws.append([r['competitor'], r['category_url'], r['title'][:120], r['h1'][:80], len(r['pages']), r['types'].get('фільтр', 0),
                   r['types'].get('товар', 0), r['types'].get('стаття', 0), '; '.join(f'{p["label"]} ({p["keywords"]})' for p in top)])
    ws2 = wb.create_sheet('Сторінки конкурентів')
    ws2.append(['Конкурент', 'Тип', 'Назва', 'URL', 'Ключів у Serpstat', 'Джерело'])
    for r in results:
        for p in sorted(r['pages'], key=lambda p: (p['type'], -p['keywords'])):
            ws2.append([r['competitor'], p['type'], p['label'], p['url'], p['keywords'], p['source']])
    ws3 = wb.create_sheet('Ознаки конкурентів')
    ws3.append(['Ознака', 'У скількох конкурентів', 'Ключів на сторінках', 'Приклади назв'])
    for t, n, k, ex in facets:
        ws3.append([t, n, k, ' | '.join(ex)])
    for w in wb.worksheets:
        for c in w[1]:
            c.font = bold
        w.freeze_panes = 'A2'
        w.auto_filter.ref = w.dimensions
    wb.save(out)


def main():
    ap = argparse.ArgumentParser(description='Структура конкурентів: сторінки розділу, які ранжуються, і спільні ознаки')
    ap.add_argument('config', nargs='?', help='JSON-конфіг ніші з полем competitors')
    ap.add_argument('--urls', nargs='*', help='адреси категорій конкурентів (замість конфігу)')
    ap.add_argument('--collect', type=int, default=0, help='зібрати ключі з N найсильніших фільтрів і категорій у inbox/')
    ap.add_argument('--no-fetch', action='store_true', help='не відкривати сторінки конкурентів, лише Serpstat')
    ap.add_argument('--max-pages', type=int, default=1, help='сторінок відповіді getDomainUrls по 1000 адрес (1)')
    ap.add_argument('--out', help='xlsx (за замовчуванням competitors.xlsx поруч із конфігом)')
    args = ap.parse_args()
    if args.config:
        cfg = json.load(open(args.config, encoding='utf-8'))
        base_dir = os.path.dirname(os.path.abspath(args.config))
        urls = [c['url'] if isinstance(c, dict) else c for c in cfg.get('competitors', [])]
    else:
        base_dir, urls = os.getcwd(), []
    urls += args.urls or []
    if not urls:
        raise SystemExit('немає адрес конкурентів: додай competitors у конфіг або --urls')
    results = [analyse(u, not args.no_fetch, args.max_pages) for u in urls]
    facets = facet_candidates(results)
    out = args.out or os.path.join(base_dir, 'competitors.xlsx')
    write_xlsx(results, facets, out)
    json.dump({'competitors': results, 'facets': facets}, open(os.path.splitext(out)[0] + '.json', 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    print(f'-> {out}')
    print('Спільні ознаки (у 2+ конкурентів): ' + ', '.join(f'{t} ({n})' for t, n, _, _ in facets[:40]))
    if args.collect:
        n = collect_keywords(results, args.collect, os.path.join(base_dir, 'inbox'))
        print(f'зібрано файлів ключів: {n} (inbox/)')


if __name__ == '__main__':
    main()
