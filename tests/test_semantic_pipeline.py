"""Конвеєр семантики на синтетичній ніші «кавомашини»: фасети, пороги сторінок, рішення перевірки видачі, пачки ШІ-ревю.
Ключі API не потрібні: видача і відповіді моделі тут не запитуються."""
import csv
import json
import os
import subprocess
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))

import semantic_core as sc  # noqa: E402
import serp_check  # noqa: E402
import ai_review  # noqa: E402

KEYS = [
    ('кавомашина', 9000), ('кавомашина купити', 3000), ('кавомашини', 9000),
    ('кавомашина delonghi', 2000), ('кавомашина delonghi magnifica s', 900), ('delonghi magnifica s', 700),
    ('кавомашина делонги магнифика s', 300), ('кавомашина delonghi dinamica', 400), ('delonghi dinamica plus', 350),
    ('кавомашина delonghi 1450 вт', 40),
    ('кавомашина philips', 1500), ('кавомашина philips 2200', 600), ('philips ep2220', 300), ('кавомашина philips ep2224', 200),
    ('кавомашина для дому', 1200), ('кавомашина для офісу', 500), ('кавомашина для офиса', 400), ('кавомашина для офісу купити', 200),
    ('кавомашина з капучинатором', 800), ('кавомашина з капучинатором для дому', 150), ('кавомашина з капучинатором недорого', 90),
    ('як вибрати кавомашину', 700), ('ремонт кавомашини', 500), ('кавомашина київ', 400),
    ('кавомашина для дому з капучинатором', 120),
    ('як вибрати кавомашину delonghi', 200), ('як чистити кавомашину', 250), ('як чистити кавомашину delonghi', 120),
    ('скільки коштує кавомашина', 300), ('кофемашина для дома купить', 60),
]


def make_niche(tmp_path):
    (tmp_path / 'inbox').mkdir()
    with open(tmp_path / 'inbox' / 'seed_1.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['keyword', 'volume'])
        w.writerows(KEYS)
    cfg = {
        'name': 'kavomashyny', 'category_name': 'кавомашини', 'base_name': 'Кавомашини', 'category_url': '/kavomashyny/',
        'inputs': ['inbox/*.csv'], 'minus_lists': ['reference/minus_geo_market.txt'],
        'thresholds': {'min_keys': 3, 'min_volume': 50},
        'aliases': {'делонги': 'delonghi', 'магнифика': 'magnifica'},
        'brands': [
            {'name': "De'Longhi", 'slug': 'delonghi', 'pattern': r'delonghi|делонги',
             'models': [r'\b(magnifica|dinamica)(\s+(s|plus|evo))?\b']},
            {'name': 'Philips', 'slug': 'philips', 'pattern': r'philips|філіпс',
             'models': [r'\bep\s*\d{4}\b', r'\b\d{4}\b']},
        ],
        'numeric': [{'name': 'Потужність', 'unit': 'W', 'patterns': [{'regex': r'(\d{3,4})\s*(?:вт|w)\b', 'mult': 1}]}],
        'attributes': [
            {'name': 'Призначення', 'values': [
                {'name': 'Для дому', 'slug': 'dlya-domu', 'pattern': r'для дому|для дома'},
                {'name': 'Для офісу', 'slug': 'dlya-ofisu', 'pattern': r'для офіс|для офис'}]},
            {'name': 'Тип', 'values': [{'name': 'З капучинатором', 'slug': 'z-kapuchynatorom', 'pattern': r'капучинатор'}]},
        ],
        'combos': [['Призначення', 'Тип']],
        'commercial_override': r'скільки коштує',
        'articles': {
            'topics': [{'name': 'Як вибрати кавомашину', 'pattern': r'вибрати|вибір'},
                       {'name': 'Догляд', 'slug': 'doglyad', 'pattern': r'чистити|чистка|декальцин'}],
            'split_by': ['Бренд'], 'thresholds': {'min_keys': 1, 'min_volume': 50}},
    }
    path = tmp_path / 'config.json'
    path.write_text(json.dumps(cfg, ensure_ascii=False), encoding='utf-8')
    return str(path)


def test_facets_models_and_power(tmp_path):
    cfg = sc.load_config(make_niche(tmp_path))
    f = sc.facets('кавомашина делонги магнифика s', cfg)
    assert f['brand'][0] == "De'Longhi"
    assert f['model'][0] == 'Magnifica S'
    f = sc.facets('кавомашина philips ep2224', cfg)
    assert f['model'] == ('EP2224', 'ep2224')
    f = sc.facets('кавомашина delonghi 1450 вт', cfg)
    assert f['num:Потужність'][1] == '1,45 кВт'


def test_model_label_keeps_words_apart():
    assert sc.model_label('delta pro 3')[0] == 'Delta Pro 3'
    assert sc.model_label('ac 180p')[0] == 'AC180P'
    assert sc.model_label('elite 100 v2')[1] == 'elite-100-v2'
    assert sc.model_label('ac200 max') == sc.model_label('ac200max') == ('AC200 Max', 'ac200-max')
    assert sc.model_label('p2001e plus')[0] == 'P2001E Plus'
    assert sc.model_label('ac200pl')[0] == 'AC200PL'
    assert sc.model_label('delta 2 max')[0] == 'Delta 2 Max'


def test_pages_respect_thresholds_and_balance(tmp_path):
    cfg = sc.load_config(make_niche(tmp_path))
    res = sc.build(cfg)
    pages = res['pages']
    assert res['page_of']['кавомашина делонги магнифика s'] == "De'Longhi > Magnifica S"
    # «Для дому > З капучинатором»: 2 ключі < порогу 3 -> ключі піднімаються вище
    assert 'Для дому > З капучинатором' not in pages
    for p in pages.values():
        if p['kind'] not in ('base', 'brand'):
            assert p['keys'] >= 3 and p['volume'] >= 50, p['page']
    assert 'як вибрати кавомашину' in res['infos']
    assert any(k == 'кавомашина київ' for k, _, _ in res['removed'])
    assert len(res['kept']) + len(res['infos']) + len(res['removed']) == len(KEYS)


def test_tree_levels_follow_real_parents(tmp_path):
    cfg = sc.load_config(make_niche(tmp_path))
    pages = sc.build(cfg)['pages']
    assert pages['Кавомашини (база)']['level'] == 1
    model = pages["De'Longhi > Magnifica S"]
    assert model['parent'] == "De'Longhi" and model['level'] == 3
    for p in pages.values():
        if p['parent'] in pages:
            assert p['level'] == pages[p['parent']]['level'] + 1, p['page']


def test_rel_label_drops_parent_and_merged_series():
    assert sc.rel_label({'page': 'EcoFlow > Delta 2', 'parent': 'EcoFlow'}) == 'Delta 2'
    assert sc.rel_label({'page': 'EcoFlow > Trail > Trail 300', 'parent': 'EcoFlow'}) == 'Trail 300'
    assert sc.rel_label({'page': 'Для дому > 5 кВт', 'parent': 'База'}) == 'Для дому > 5 кВт'


def test_default_tree_without_config_tree(tmp_path):
    cfg = sc.load_config(make_niche(tmp_path))
    assert cfg['_tree'][0] == ['Бренд', 'Серія?', 'Модель']
    assert ['Призначення', 'Тип'] in cfg['_tree'] and ['Потужність'] in cfg['_tree']


def test_articles_and_commercial_override(tmp_path):
    cfg = sc.load_config(make_niche(tmp_path))
    res = sc.build(cfg)
    assert 'скільки коштує кавомашина' in res['kept']          # «скільки» - інфо-маркер, але це ціна
    arts = res['articles']
    assert arts['Як вибрати кавомашину']['level'] == 2
    assert res['article_of']['як вибрати кавомашину delonghi'] == "Як вибрати кавомашину > De'Longhi"
    split = arts["Як вибрати кавомашину > De'Longhi"]
    assert split['level'] == 3 and split['parent'] == 'Як вибрати кавомашину' and split['url'].startswith('/blog/')
    assert res['article_of']['як чистити кавомашину'] == 'Догляд' and arts['Догляд']['url'] == '/blog/doglyad/'


def test_serp_merge_moves_keys_to_parent(tmp_path):
    cfg = sc.load_config(make_niche(tmp_path))
    serp = {'pages': {'Для офісу': {'decision': 'merge', 'overlap': 6, 'head': 'кавомашина для офісу', 'parent': 'Кавомашини (база)'}}}
    res = sc.build(cfg, serp)
    assert 'Для офісу' not in res['pages']
    assert res['page_of']['кавомашина для офісу'] == 'Кавомашини (база)'
    assert any(m['page'] == 'Для офісу' for m in res['merged_pages'])


def test_xlsx_and_pages_json(tmp_path):
    cfg_path = make_niche(tmp_path)
    comp = {'competitors': [{'competitor': 'shop.ua', 'category_url': 'https://shop.ua/kavomashyny/', 'types': {'фільтр': 1},
                             'pages': [{'url': 'https://shop.ua/kavomashyny/delonghi/', 'type': 'фільтр', 'label': 'DeLonghi',
                                        'keywords': 40}]}],
            'facets': [['delonghi', 2, 40, ['DeLonghi']]]}
    (tmp_path / 'competitors.json').write_text(json.dumps(comp, ensure_ascii=False), encoding='utf-8')
    out = tmp_path / 'core.xlsx'
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    r = subprocess.run([sys.executable, os.path.join(ROOT, 'scripts', 'semantic_core.py'), cfg_path, '--out', str(out)],
                       capture_output=True, text=True, encoding='utf-8', env=env, cwd=ROOT)
    assert r.returncode == 0, r.stderr
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ['Зведення', 'UA', 'RU', 'Структура', 'Блог', 'Статті', 'Видалено', 'Конкуренти']
    header = [c.value for c in wb['UA'][1]]
    assert header == ['Категорія', 'Keyword', 'Volume', 'Intent', 'Group 1', 'Group 2', 'Group 3', 'URL', 'Source', 'Коментар']
    rows = {r[1]: r for r in wb['UA'].iter_rows(min_row=2, values_only=True)}
    assert rows['кавомашина делонги магнифика s'][4:7] == ('Кавомашини', "Бренд - De'Longhi", 'Модель - Magnifica S')
    assert rows['кавомашина делонги магнифика s'][7] == '/kavomashyny/delonghi/magnifica-s/'
    assert rows['кавомашина'][4:7] == ('Кавомашини', None, None) and rows['кавомашина'][8] == 'сіди'
    assert [r[1] for r in wb['RU'].iter_rows(min_row=2, values_only=True)] == ['кофемашина для дома купить']
    blog = {r[1]: r for r in wb['Блог'].iter_rows(min_row=2, values_only=True)}
    assert blog['як вибрати кавомашину delonghi'][4:7] == ('Кавомашини', 'Як вибрати кавомашину', "Бренд - De'Longhi")
    assert 'DeLonghi (40)' in [c.value for c in wb['Конкуренти']['F']]
    import cluster_gaps  # аудит бази читає новий формат: база - рядки, де заповнено лише Group 1
    base_keys = {k for k, _ in cluster_gaps.load(str(out), r'\(баз', False)}
    assert 'кавомашина' in base_keys and 'кавомашина делонги магнифика s' not in base_keys
    data = json.loads((tmp_path / 'core.pages.json').read_text(encoding='utf-8'))
    assert {p['kind'] for p in data['pages']} >= {'base', 'brand', 'model'}
    assert any(a['page'] == 'Догляд' for a in data['articles'])


def test_serp_decision_rules():
    a = [{'url': f'https://shop{i}.ua/kavomashyny/dlya-ofisu/', 'title': ''} for i in range(5)] + \
        [{'url': f'https://blog{i}.ua/news/', 'title': ''} for i in range(5)]
    b = [{'url': f'https://shop{i}.ua/kavomashyny/', 'title': ''} for i in range(10)]
    toks = {'dlyaofisu', 'ofis'}
    assert serp_check.decide(a, b, toks)[0] == 'keep'          # 5 з 10 саме про уточнення
    assert serp_check.decide(b, b, toks)[0] == 'merge'         # та сама видача
    assert serp_check.decide(a[5:], b, toks)[0] == 'merge'     # видача не виділяє уточнення
    assert serp_check.decide([], b, toks)[0] == 'keep'         # немає топу - не зливаємо наосліп


def test_refinement_tokens_from_slug_and_head():
    page = {'page': 'Потужність > 2 кВт', 'url': '/cat/2-kvt/', 'head': 'зарядна станція 2 квт'}
    parent = {'page': 'База', 'url': '/cat/', 'head': 'зарядна станція'}
    toks = serp_check.refinement_tokens(page, parent)
    assert {'2000', '2kvt', '2kw'} <= toks
    page = {'page': 'Аксесуари', 'url': '/cat/aksesuary/', 'head': 'кабель для зарядної станції'}
    toks = serp_check.refinement_tokens(page, parent)
    assert 'kabe' in toks and not any(t.startswith('zar') for t in toks)


def test_ai_batches_keep_pages_together():
    pages = [{'page': 'A', 'keywords': [[f'a{i}', 1] for i in range(3)]},
             {'page': 'B', 'keywords': [[f'b{i}', 1] for i in range(5)]},
             {'page': 'C', 'keywords': [[f'c{i}', 1] for i in range(12)]}]
    bs = ai_review.batches(pages, 6)
    assert sum(len(kws) for b in bs for _, kws in b) == 20
    assert all(sum(len(kws) for _, kws in b) <= 6 for b in bs)
    assert '## A' in ai_review.batch_text(bs[0])


def test_one_sheet_with_language_column(tmp_path):
    cfg = sc.load_config(make_niche(tmp_path))
    cfg['output'] = {'split_lang': False, 'site': 'https://shop.ua'}
    out = tmp_path / 'one.xlsx'
    sc.write_outputs(cfg, sc.build(cfg), str(out))
    wb = openpyxl.load_workbook(out)
    assert 'Кластери' in wb.sheetnames and 'UA' not in wb.sheetnames
    header = [c.value for c in wb['Кластери'][1]]
    assert header[-2:] == ['Мова', 'Коментар']
    rows = {r[1]: r for r in wb['Кластери'].iter_rows(min_row=2, values_only=True)}
    assert rows['кофемашина для дома купить'][header.index('Мова')] == 'ru'
    assert rows['кавомашина'][header.index('URL')] == 'https://shop.ua/kavomashyny/'


def test_key_lang_and_step_label():
    assert sc.key_lang('зарядна станція') == 'uk' and sc.key_lang('як вибрати кавомашину') == 'uk'
    assert sc.key_lang('купить зарядную станцию') == 'ru' and sc.key_lang('зарядная станция') == 'ru'
    assert sc.key_lang('ecoflow delta 2') == ''
    assert sc.step_label([('Бренд', 'EcoFlow')]) == 'Бренд - EcoFlow'
    assert sc.step_label([('Серія', 'Trail'), ('Модель', 'Trail 300')]) == 'Модель - Trail 300'
    assert sc.step_label([('Призначення', 'Для дому'), ('Потужність', '5 кВт')]) == 'Призначення - Для дому, Потужність - 5 кВт'
    assert sc.step_label([('Тема', 'Як вибрати')]) == 'Як вибрати'
