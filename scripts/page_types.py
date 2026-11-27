# -*- coding: utf-8 -*-
"""Тип сторінки за адресою (і заголовком, якщо є): категорія, фільтр, товар, стаття, маркетплейс, відео, форум.

Евристики, а не істина: для звітів про структуру конкурентів і про те, які сторінки Google показує у видачі.
Використовують competitors.py, serp_check.py і article_examples.py.
"""
import re
from urllib.parse import unquote, urlparse

MARKETPLACES = ('rozetka.com.ua', 'prom.ua', 'olx.ua', 'epicentrk.ua', 'allo.ua', 'hotline.ua', 'ek.ua', 'e-katalog',
                'price.ua', 'bigl.ua', 'shafa.ua', 'kasta.ua', 'aliexpress', 'amazon.')
VIDEO = ('youtube.com', 'youtu.be', 'tiktok.com', 'vimeo.com')
FORUMS = ('reddit.com', 'forum', 'pikabu', 'quora.com', 'dou.ua/forums')
WIKI = ('wikipedia.org', 'wiki.')

ARTICLE_PATH = re.compile(r'/(blog|blogs|news|novosti|novyny|stat[tiy]\w*|article|articles|info|journal|zhurnal|guide|'
                          r'poradi|porady|sovety|obzor\w*|ohlyad\w*|oglyad\w*|review|faq|knowledge|pidbir|podbor|media)(/|$)', re.I)
ARTICLE_SLUG = re.compile(r'(^|[/-])(yak|kak|shcho|chto|chomu|pochemu|skilky|skolko|navishcho|zachem|top-?\d+|reitynh|rejting|'
                          r'rating|how-to|what-is|vs|porivnyannya|sravnenie|vybraty|vybrat|obraty|vibor|vybir)([/-]|$)', re.I)
ARTICLE_TITLE = re.compile(r'як |как |що таке|что такое|чому|почему|скільки|сколько|топ[ -]?\d|рейтинг|огляд|обзор|порівнян|сравнен|'
                           r'гайд|поради|советы|інструкц|инструкц|\bvs\b|how to|what is|review', re.I)
PRODUCT = re.compile(r'(/p\d{5,}/|/product/|/products/|/tovar/|/item/|/goods/|[-/]p\d{5,}\b|\d{5,}\.html$|/dp/)', re.I)
FILTER_MARK = re.compile(r'(\w+=\w|;|:|/filter/|/f/|/brand[-_/]|/producer|/proizvoditel|/vyrobnyk|/seriya|/series|/tip[-_:]|'
                         r'/type[-_:]|/moshch|/potuzhn|/power)', re.I)


EXCLUDED_HOSTS = re.compile(r'(\.ru|\.su|\.рф|\.xn--p1ai|\.by)$|russia|yandex|(^|\.)dzen\.|(^|\.)vk\.com$|(^|\.)ok\.ru$|rutube', re.I)


def excluded_host(host):
    """Російські й білоруські сайти та РФ-платформи: у приклади й рекомендації не беремо (див. geo_exclude)."""
    return bool(EXCLUDED_HOSTS.search(host))


def host_of(url):
    h = urlparse(url).netloc.lower()
    return h[4:] if h.startswith('www.') else h


def classify(url, title='', category_url=None):
    """Тип сторінки. category_url - адреса категорії цього сайту: усе глибше за неї - фільтр або підкатегорія."""
    host = host_of(url)
    path = unquote(urlparse(url).path)
    full = url.lower()
    if any(v in host for v in VIDEO):
        return 'відео'
    if any(w in host for w in WIKI):
        return 'вікі'
    if any(f in full for f in FORUMS):
        return 'форум'
    if ARTICLE_PATH.search(path) or (ARTICLE_SLUG.search(path) and not PRODUCT.search(path)) or (title and ARTICLE_TITLE.search(title)):
        return 'стаття'
    if PRODUCT.search(path):
        return 'товар'
    if category_url:
        base = urlparse(category_url).path.rstrip('/')
        if path.rstrip('/') == base:
            return 'категорія'
        if path.startswith(base + '/') or (urlparse(url).query and path.rstrip('/') == base):
            return 'фільтр'
    if FILTER_MARK.search(path) or urlparse(url).query:
        return 'фільтр'
    if any(m in host for m in MARKETPLACES):
        return 'маркетплейс'
    # довгий слаг із моделлю й цифрами - найчастіше картка товару
    last = path.rstrip('/').split('/')[-1]
    if len(last) > 40 and re.search(r'\d{3,}', last):
        return 'товар'
    return 'категорія'


def serp_profile(top, category_hint=None):
    """Зведення по топу: скільки яких сторінок і який тип посадкової підходить під цей запит."""
    kinds = [classify(it['url'], it.get('title', '')) for it in top]
    n = len(kinds) or 1
    count = {k: kinds.count(k) for k in set(kinds)}
    listing = count.get('категорія', 0) + count.get('фільтр', 0) + count.get('маркетплейс', 0)
    articles = count.get('стаття', 0) + count.get('вікі', 0) + count.get('форум', 0) + count.get('відео', 0)
    products = count.get('товар', 0)
    if articles >= max(4, n * 0.5):
        rec = 'стаття'
    elif products >= max(4, n * 0.5):
        rec = 'картка товару або лендинг моделі'
    elif listing >= max(4, n * 0.4):
        rec = 'категорія або фільтр'
    else:
        rec = 'змішана видача'
    return count, rec
