# -*- coding: utf-8 -*-
"""Тянет КАНДИДАТОВ в бренды/атрибуты из фильтров категории конкурента (данные, не память).
Движки rozetka/allo/foxtrot/epicentrk отдают полный HTML на браузерный UA (проверено).
  python scripts/brands_from_competitor.py "https://rozetka.com.ua/ua/tablet_covers/c305219/" [URL2 ...]
Выводит значения фильтров (бренды/типы) → вычитай в reference/brands/<ниша>.txt и в кластеризатор.
Заглушка/403 → фолбэк: открой фильтры в браузере вручную.
"""
import re, sys, urllib.request
sys.stdout.reconfigure(encoding='utf-8')
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36'
PATS = [r'brand__([a-z0-9_\-]+)', r'tip__([a-z0-9_\-]+)', r'prop__([a-z0-9_\-]+)',
        r'\btm[-_]([a-z0-9\-]{2,})', r'producer[=/_\-]([a-z0-9\-]{2,})',
        r'vendor[=_]([a-z0-9\-]{2,})', r'filter[/_]brand[=/_]([a-z0-9\-]+)', r'brands?[=/]([a-z0-9\-]{2,})']

def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept-Language': 'uk,ru;q=0.9'})
    return urllib.request.urlopen(req, timeout=30).read().decode('utf-8', 'ignore')

def main():
    urls = sys.argv[1:]
    if not urls:
        raise SystemExit('передай URL категорій конкурентів')
    from collections import Counter
    cand = Counter()
    for u in urls:
        try:
            html = fetch(u)
        except Exception as e:
            print(f'WARN {u}: не завантажилось ({e}) -> відкрий фільтри в браузері вручну', file=sys.stderr); continue
        low = html.lower()
        for p in PATS:
            for m in re.findall(p, low):
                v = m.strip('-_')
                if len(v) >= 2 and not v.isdigit():
                    cand[v] += 1
        print(f'{u}: HTML {len(html)} симв.', file=sys.stderr)
    print('=== КАНДИДАТИ в бренди та фільтри (зі структури конкурентів) ===')
    for v, n in cand.most_common(80):
        print(f'  {n:>2}  {v}')
    if not cand:
        print('порожньо: інший рушій або заглушка -> відкрий фільтри в браузері вручну')

if __name__ == '__main__':
    main()
