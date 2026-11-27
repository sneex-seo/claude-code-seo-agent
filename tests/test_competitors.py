"""Конкуренти, типи сторінок і приклади статей. Мережа й ключі API не потрібні: Serpstat і HTML підмінено."""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))

import article_examples  # noqa: E402
import competitors as cp  # noqa: E402
from page_types import classify, serp_profile  # noqa: E402

CAT = 'https://shop.ua/ua/catalog/zaryadni-stantsii/'


def test_classify_page_types():
    assert classify('https://www.youtube.com/watch?v=1') == 'відео'
    assert classify('https://shop.ua/blog/yak-vybraty-zaryadnu-stantsiyu/') == 'стаття'
    assert classify('https://shop.ua/ecoflow-delta-2/p358123456/') == 'товар'
    assert classify('https://shop.ua/cat/brand-ecoflow/', category_url='https://shop.ua/cat/') == 'фільтр'
    assert classify('https://rozetka.com.ua/ua/zaryadni-stantsii/') == 'маркетплейс'


def test_serp_profile_recommends_page_type():
    arts = [{'url': f'https://b{i}.ua/blog/yak-vybraty/', 'title': ''} for i in range(6)]
    cats = [{'url': f'https://s{i}.ua/zaryadni-stantsii/', 'title': ''} for i in range(4)]
    assert serp_profile(arts + cats)[1] == 'стаття'
    assert serp_profile(cats * 2 + arts[:2])[1] == 'категорія або фільтр'


def test_language_versions_are_one_page():
    assert cp.section_slug('https://rozetka.com.ua/ua/zaryadnie-stantsii-4674585/c4674585/') == 'zaryadnie-stantsii-4674585'
    assert cp.page_key('https://shop.ua/ru/catalog/zaryadni-stantsii/ecoflow/') == cp.page_key(CAT + 'ecoflow/')
    for label in ('RU', 'РУС ( RU )', 'УКР', 'ua'):
        assert cp.LANG_LABEL.match(label), label
    assert not cp.LANG_LABEL.match('EcoFlow') and not cp.LANG_LABEL.match('Russell Hobbs')


def test_section_type():
    slug, stem = 'zaryadni-stantsii', 'zaryadni-stantsii'
    assert cp.section_type('https://shop.ua/ru/zaryadni-stantsii/', slug, stem) == 'категорія'
    assert cp.section_type(CAT + 'brand-ecoflow/', slug, stem) == 'фільтр'
    assert cp.section_type('https://shop.ua/catalog/zaryadni-stantsii-jackery/', slug, stem) == 'фільтр'
    assert cp.section_type('https://shop.ua/poverbanky-i-zaryadni-stantsii/', slug, stem) == 'суміжна категорія'
    assert cp.section_type(CAT + 'page-3/', slug, stem) == 'пагінація'


def test_analyse_merges_languages_and_skips_switch_label(monkeypatch):
    ranked = {CAT: 100, 'https://shop.ua/ru/catalog/zaryadni-stantsii/': 60,
              CAT + 'brand-ecoflow/': 10, 'https://shop.ua/ru/catalog/zaryadni-stantsii/brand-ecoflow/': 5}
    html = ('<title>Зарядні станції</title><h1>Зарядні станції</h1>'
            '<a href="/ru/catalog/zaryadni-stantsii/">RU</a>'
            '<a href="/ua/catalog/zaryadni-stantsii/brand-ecoflow/">EcoFlow (12)</a>'
            '<a href="/ua/catalog/zaryadni-stantsii/page-2/">2</a>')
    monkeypatch.setattr(cp, 'domain_urls', lambda host, slug, max_pages: ranked)
    monkeypatch.setattr(cp, 'fetch', lambda url: html)
    r = cp.analyse(CAT)
    by = {p['type']: p for p in r['pages']}
    assert by['категорія']['keywords'] == 160 and by['категорія']['versions'] == 2
    assert by['фільтр']['label'] == 'EcoFlow' and by['фільтр']['keywords'] == 15
    assert 'пагінація' in by and r['types'].get('фільтр') == 1


def test_facet_candidates_need_two_competitors():
    res = [{'competitor': 'a.ua', 'pages': [{'type': 'фільтр', 'label': 'EcoFlow', 'keywords': 10}]},
           {'competitor': 'b.ua', 'pages': [{'type': 'фільтр', 'label': 'Зарядні станції EcoFlow', 'keywords': 5},
                                            {'type': 'фільтр', 'label': 'Jackery', 'keywords': 3}]}]
    rows = {t: (n, k) for t, n, k, _ in cp.facet_candidates(res)}
    assert rows['ecoflow'] == (2, 15) and 'jackery' not in rows


def test_article_examples_prefer_competitors():
    items = [{'pos': 1, 'type': 'стаття', 'competitor': False, 'url': 'a'},
             {'pos': 2, 'type': 'категорія', 'competitor': True, 'url': 'b'},
             {'pos': 3, 'type': 'стаття', 'competitor': True, 'url': 'c'},
             {'pos': 4, 'type': 'відео', 'competitor': False, 'url': 'd'},
             {'pos': 5, 'type': 'форум', 'competitor': False, 'url': 'e'}]
    assert [x['url'] for x in article_examples.pick_examples(items, 2)] == ['c', 'a']
