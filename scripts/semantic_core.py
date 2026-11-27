# -*- coding: utf-8 -*-
"""Збір, чистка і багаторівнева структура семантики за конфігом ніші (скіл semantic-clustering).

Конфіг ніші (JSON) пише агент: звідки брати ключі, правила чистки, бренди, серії й моделі, числові ознаки
(потужність, ємність), атрибути (призначення, тип), дерево сторінок, теми статей і пороги. Скрипт детерміновано:
  1. зливає вхідні csv (keyword, volume) з максимумом частотності;
  2. відсіює гео, майданчики й сміття (мінус-листи), чужі теми (topic_rules); інфо-запити йдуть у статті;
  3. позначає дублі з однаковою частотністю (variant_dedupe) - вони в кластері, але не в ∑;
  4. для кожного комерційного ключа будує кандидатів у сторінки за ланцюжками дерева (`tree`):
     «Бренд > Серія? > Модель», «Призначення > Потужність > Тип»... Рівень 1 - база категорії,
     далі 2, 3, 4 і глибше - скільки дозволяють ланцюжки;
  5. кладе ключ на найглибшу сторінку, де >= min_keys ключів і >= min_volume попиту, крім злитих
     перевіркою видачі (serp_check.py); дрібні сторінки розформовуються, ключі піднімаються вище;
  6. інфо-запити групує в теми статей (`articles`), за потреби з розбивкою за брендом чи призначенням;
  7. пише xlsx (Кластери, Структура, Статті, Інфо, Видалено, Конкуренти, Зведення) і <out>.pages.json.

  python scripts/semantic_core.py clients/shop/semantics/config.json
  python scripts/semantic_core.py clients/shop/semantics/config.json --serp .../serp_check.json --ai .../ai_review.csv \
      --articles .../articles.json
"""
import argparse
import csv
import glob
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from serpstat_api import is_minus, load_minus  # noqa: E402
from variant_dedupe import mark  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LAT2CYR = str.maketrans("qwertyuiop[]asdfghjkl;'zxcvbnm,.", 'йцукенгшщзхїфівапролджєячсмитьбю')
BASE_LEVEL = 1
BRAND, SERIES, MODEL = 'Бренд', 'Серія', 'Модель'
TOPIC = 'Тема'
LANG_SHEETS = {'uk': 'UA', 'ru': 'RU'}
TRANSLIT = str.maketrans({'а': 'a', 'б': 'b', 'в': 'v', 'г': 'h', 'ґ': 'g', 'д': 'd', 'е': 'e', 'є': 'ie', 'ж': 'zh', 'з': 'z',
                          'и': 'y', 'і': 'i', 'ї': 'i', 'й': 'i', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p',
                          'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh',
                          'щ': 'shch', 'ь': '', 'ю': 'iu', 'я': 'ia', "'": '', ' ': '-', ',': '', ':': ''})


def norm(s):
    return re.sub(r'\s+', ' ', str(s).lower().replace('ё', 'е').strip())


def rx(p):
    return re.compile(p, re.I)


def slugify(text):
    return re.sub(r'-+', '-', text.lower().translate(TRANSLIT)).strip('-')


# ---------- конфіг ----------

def load_config(path):
    cfg = json.load(open(path, encoding='utf-8'))
    cfg['_dir'] = os.path.dirname(os.path.abspath(path))
    cfg.setdefault('thresholds', {}).setdefault('min_keys', 3)
    cfg['thresholds'].setdefault('min_volume', 50)
    cfg.setdefault('base_name', 'Категорія')
    cfg.setdefault('category_url', '/')
    for b in cfg.get('brands', []):
        b['_rx'] = rx(b['pattern'])
        b['_models'] = [rx(m) for m in b.get('models', [])]
        b['_series'] = [(s['name'], s['slug'], rx(s['pattern'])) for s in b.get('series', [])]
    for n in cfg.get('numeric', []):
        n['_rx'] = [(rx(p['regex']), p.get('mult', 1)) for p in n['patterns']]
    for a in cfg.get('attributes', []):
        a['_values'] = [(v['name'], v['slug'], rx(v['pattern']), v.get('override', False)) for v in a['values']]
    for r in cfg.get('topic_rules', []):
        r['_rx'] = rx(r['pattern'])
        r['_unless'] = rx(r['unless']) if r.get('unless') else None
    for c in cfg.get('comments', []):
        c['_rx'] = rx(c['pattern'])
    for r in cfg.get('intent_rules', []):
        r['_rx'] = rx(r['pattern'])
    cfg['_require'] = rx(cfg['require_pattern']) if cfg.get('require_pattern') else None
    # комерційні запити з інфо-словами («скільки коштує», «який купити») лишаються в кластерах, а не в статтях
    cfg['_commercial'] = rx(cfg['commercial_override']) if cfg.get('commercial_override') else None
    for r in cfg.get('ai_dismiss', []):
        r['_rx'] = rx(r['pattern'])
    art = cfg.setdefault('articles', {})
    for t in art.get('topics', []):
        t['_rx'] = rx(t['pattern'])
        t.setdefault('slug', slugify(t['name']))
    art.setdefault('thresholds', {'min_keys': 3, 'min_volume': 30})
    cfg['_tree'] = cfg.get('tree') or default_tree(cfg)
    return cfg


def default_tree(cfg):
    """Дерево для конфігу без `tree`: бренд > серія > модель, бренд > число, пари `combos`, числа й атрибути окремо."""
    tree = [[BRAND, SERIES + '?', MODEL]]
    tree += [[BRAND, n['name']] for n in cfg.get('numeric', []) if n.get('with_brand', True)]
    tree += [list(c) for c in cfg.get('combos', [])]
    tree += [[n['name']] for n in cfg.get('numeric', [])]
    tree += [[a['name']] for a in cfg.get('attributes', [])]
    return tree


# ---------- ознаки ----------

MODEL_WORDS = {'pro', 'max', 'air', 'mini', 'plus', 'ultra'}


def model_label(text):
    t = re.sub(r'\s+', ' ', text.strip())
    # «ac 180» -> «ac180»; «delta 2» і «pro 3» лишаються окремими словами
    t = re.sub(r'\b([a-z]{1,3})\s+(\d)', lambda m: m.group(0) if m.group(1) in MODEL_WORDS else m.group(1) + m.group(2), t)
    # «ac200max» = «ac200 max»: суфікс-слово після коду моделі відділяємо, «ac200pl» і «ac180p» не чіпаємо
    t = re.sub(r'\b([a-z]+\d+[a-z]?)(max|pro|plus)\b', r'\1 \2', t)
    words = t.split()
    label = ' '.join(w.upper() if re.search(r'\d', w) else w.capitalize() for w in words)
    slug = '-'.join(re.sub(r'[^a-z0-9а-яіїєґ]+', '', w) for w in words)
    return label, slug


def number_label(value, unit):
    if unit == 'W':
        if value >= 1000:
            v = f'{value / 1000:g}'.replace('.', ',')
            return f'{v} кВт', f'{v.replace(",", "-")}-kvt'
        return f'{value} Вт', f'{value}-vt'
    if unit == 'Wh':
        return f'{value} Вт·год', f'{value}-vt-hod'
    return f'{value} {unit}', f'{value}-{unit.lower()}'


def facets(k, cfg):
    f = {}
    aliases = cfg.get('aliases', {})
    for b in cfg.get('brands', []):
        if b['_rx'].search(k):
            f['brand'] = (b['name'], b['slug'])
            kk = k
            for a, t in aliases.items():
                kk = re.sub(a, t, kk)
            for name, slug, s in b['_series']:
                if s.search(kk):
                    f['series'] = (name, slug)
                    break
            for m in b['_models']:
                hit = m.search(kk)
                if hit:
                    f['model'] = model_label(hit.group(0))
                    break
            break
    for n in cfg.get('numeric', []):
        for r, mult in n['_rx']:
            hit = r.search(k)
            if hit:
                val = round(float(hit.group(1).replace(',', '.')) * mult)
                f['num:' + n['name']] = (val, *number_label(val, n.get('unit', '')))
                break
    for a in cfg.get('attributes', []):
        for name, slug, r, override in a['_values']:
            if r.search(k):
                f['attr:' + a['name']] = (name, slug)
                if override:
                    f['override'] = True  # аксесуар до станції бренду - сторінка аксесуара, а не бренду
                break
    return f


def facet_value(f, name):
    """(назва для сторінки, slug) значення ознаки або None."""
    if name == BRAND:
        return f.get('brand')
    if name == SERIES:
        return f.get('series')
    if name == MODEL:
        return f.get('model')
    if ('attr:' + name) in f:
        return f['attr:' + name]
    if ('num:' + name) in f:
        v = f['num:' + name]
        return v[1], v[2]
    return None


def candidates(f, cfg, steps=None):
    """(сторінка, батьківська, url, рівень, вид) за ланцюжками дерева, від найглибшої до бази.
    steps - словник, куди записати шлях сторінки ознаками: [(«Бренд», «EcoFlow»), («Серія», «Delta»)]."""
    cat = cfg['category_url']
    base = f"{cfg['base_name']} (база)"
    found = {}
    for ci, chain in enumerate(cfg['_tree']):
        if chain and chain[0].rstrip('?') == BRAND and f.get('override'):
            continue  # аксесуар не йде на сторінки бренду
        names, slugs, parent, depth, path = [], [], base, 0, []
        for step in chain:
            optional = step.endswith('?')
            facet = step.rstrip('?')
            if facet == BRAND and depth > 0 and not f.get('override'):
                break  # «Тип > Бренд» лише для аксесуарів
            v = facet_value(f, facet)
            if v is None:
                if optional:
                    continue
                break
            label = v[0]
            if ('num:' + facet) in f and depth == 0:
                label = f'{facet} {v[0]}'  # «Потужність 2 кВт»: на першому рівні число без назви ознаки незрозуміле
            names.append(label)
            slugs.append(v[1])
            path.append((facet, v[0]))
            depth += 1
            name = ' > '.join(names)
            if steps is not None:
                steps.setdefault(name, list(path))
            kind = 'brand' if (facet == BRAND and depth == 1) else ('model' if facet == MODEL else ('series' if facet == SERIES else 'facet'))
            if name not in found:
                found[name] = (name, parent, cat + '/'.join(slugs) + '/', depth + 1, kind, ci)
            parent = name
    out = sorted(found.values(), key=lambda c: (-c[3], c[5]))
    out = [c[:5] for c in out]
    out.append((base, None, cat, BASE_LEVEL, 'base'))
    return out


# ---------- чистка ----------

def removal_reason(k, f, cfg, minus):
    if is_minus(k, minus):
        hit = next(m for m in minus if (m in k if ' ' in m else m in k.split()))
        return f'гео, майданчик або сміття: «{hit}»'
    stems = cfg.get('layout_stems')
    if stems and re.fullmatch(r"[a-z\[\];',. ]+", k) and re.search(stems, k.translate(LAT2CYR)):
        return 'крива розкладка'
    for r in cfg.get('numeric_limits', []):
        v = f.get('num:' + r['facet'])
        if v and v[0] >= r['max']:
            return r['reason']
    for r in cfg.get('topic_rules', []):
        if r['_rx'].search(k) and not (r['_unless'] and r['_unless'].search(k)):
            return r['reason']
    if cfg.get('_require') and 'brand' not in f and not cfg['_require'].search(k):
        return 'без ознак категорії'
    return None


def read_inputs(cfg):
    core, src = {}, defaultdict(set)
    for pattern in cfg['inputs']:
        for path in sorted(glob.glob(os.path.join(cfg['_dir'], pattern))):
            fname = os.path.basename(path)
            m = re.match(r'url_([^_]+)_', fname)
            tag = 'сіди' if fname.startswith('seed') else (f'конкурент: {m.group(1)}' if m else os.path.splitext(fname)[0])
            for r in csv.DictReader(open(path, encoding='utf-8-sig')):
                k = norm(r.get('keyword') or r.get('Keyword') or '')
                if not k:
                    continue
                try:
                    v = int(float(r.get('volume') or r.get('Volume') or 0))
                except ValueError:
                    v = 0
                core[k] = max(core.get(k, 0), v)
                src[k].add(tag)
    return core, src


# ---------- розкладка ----------

def assign(keys, cands, weight, dup_of, th, exempt=('base', 'brand'), excluded=()):
    """Найглибша сторінка, що набирає поріг; дрібні виключаються ітеративно."""
    excluded = set(excluded)
    for _ in range(30):
        page_of, meta = {}, {}
        for k in keys:
            c = next(c for c in cands[k] if c[4] == 'base' or c[0] not in excluded)
            page_of[k] = c[0]
            meta[c[0]] = c
        n_k, n_v = Counter(), Counter()
        for k in keys:
            n_k[page_of[k]] += 1
            if k not in dup_of:
                n_v[page_of[k]] += weight[k]
        small = {g for g, c in meta.items() if c[4] not in exempt and (n_k[g] < th['min_keys'] or n_v[g] < th['min_volume'])}
        if not small:
            return page_of, meta
        excluded |= small
    return page_of, meta


def page_records(keys, page_of, meta, weight, dup_of, base, chains_of):
    pages = defaultdict(list)
    for k in keys:
        pages[page_of[k]].append(k)
    out = {}
    for g, (name, parent, url, level, kind) in meta.items():
        # батьківська - найближча наявна сторінка вгору по ланцюжку
        p = parent
        while p is not None and p not in pages and p != base:
            p = chains_of.get(p)
        if p is None or p not in pages:
            p = base if g != base else None
        head = next((k for k in pages[g] if k not in dup_of), pages[g][0])
        out[g] = {'page': g, 'parent': p, 'url': url, 'level': level, 'kind': kind, 'head': head, 'keys': len(pages[g]),
                  'volume': sum(weight[k] for k in pages[g] if k not in dup_of), 'keywords': [[k, weight[k]] for k in pages[g]]}

    # рівень - глибина в реальному дереві: якщо проміжну сторінку злито або вона замала, гілка піднімається
    def depth(g, seen=()):
        p = out[g]['parent']
        if p is None or p not in out or p in seen:
            return out[g]['level']
        return depth(p, seen + (g,)) + 1
    for g, lv in {g: depth(g) for g in out}.items():
        out[g]['level'] = lv
    return out


def rel_label(p):
    """Назва сторінки відносно батьківської: «EcoFlow > Delta 2» під «EcoFlow» -> «Delta 2»."""
    if p.get('parent') and p['page'].startswith(p['parent'] + ' > '):
        rest = p['page'][len(p['parent']) + 3:].split(' > ')
        # проміжну «Trail» злито: «Trail > Trail 300» -> «Trail 300», модель і так містить серію
        return rest[-1] if len(rest) == 2 and rest[1].startswith(rest[0]) else ' > '.join(rest)
    return p['page']


def article_candidates(k, f, cfg, steps=None):
    art = cfg['articles']
    topic = next((t for t in art.get('topics', []) if t['_rx'].search(k)), None)
    tname, tslug = (topic['name'], topic['slug']) if topic else ('Інше', 'inshe')
    out = []
    for facet in art.get('split_by', []):
        v = facet_value(f, facet)
        if v:
            out.append((f'{tname} > {v[0]}', tname, f'/blog/{tslug}-{v[1]}/', 3, 'article'))
            if steps is not None:
                steps.setdefault(out[-1][0], [(TOPIC, tname), (facet, v[0])])
    out.append((tname, None, f'/blog/{tslug}/', 2, 'base'))  # тема статті лишається навіть малою
    if steps is not None:
        steps.setdefault(tname, [(TOPIC, tname)])
    return out


def key_lang(k):
    """Мова ключа: 'uk', 'ru' або '' (латиниця, бренд, модель - мову не визначити)."""
    if re.search(r"[іїєґ']", k) or re.search(r'\b(як|що|де|або|ціна|ціни|купити|скільки|який|яка|яку)\b', k):
        return 'uk'
    if re.search(r'[ыэъё]', k) or re.search(r'\b(как|что|где|или|цена|цены|купить|сколько|какой|какая|какую|для дома)\b', k) \
            or re.search(r'[а-я](ая|ое|ые|ую|ть|ться|ция|ции|ций)\b', k):
        return 'ru'
    return ''


def step_label(own):
    """Підпис рівня для колонок Group: «Бренд - EcoFlow»; тема статті - без назви ознаки."""
    if len(own) == 2 and own[1][1].startswith(own[0][1]):
        own = own[1:]  # «Серія - Trail, Модель - Trail 300» -> «Модель - Trail 300»
    return ', '.join(v if f == TOPIC else f'{f} - {v}' for f, v in own)


def group_parts(p, pages, base_name):
    """Шлях сторінки в реальному дереві для колонок Group 1..N: категорія, далі кожен рівень своєю колонкою."""
    chain, cur, seen = [], p, set()
    while cur and cur['page'] not in seen:
        seen.add(cur['page'])
        if not (cur['kind'] == 'base' and cur['page'].endswith('(база)')):
            chain.insert(0, cur)
        cur = pages.get(cur['parent']) if cur.get('parent') else None
    parts = [base_name]
    for node in chain:
        mine = node.get('steps', [])
        par = pages.get(node.get('parent') or '', {}).get('steps', [])
        own = mine[len(par):] if par and mine[:len(par)] == par else mine
        parts.append(step_label(own) or node['page'])
    return parts


def build(cfg, serp=None):
    core, src = read_inputs(cfg)
    minus = []
    for m in cfg.get('minus_lists', []):
        minus += load_minus(os.path.join(ROOT, m))
    info = load_minus(os.path.join(ROOT, 'reference', 'info_markers.txt')) + cfg.get('info_extra', [])

    kept, removed, infos, feats = {}, [], {}, {}
    for k, v in core.items():
        f = facets(k, cfg)
        feats[k] = f
        why = removal_reason(k, f, cfg, minus)
        if why:
            removed.append((k, v, why))
        elif is_minus(k, info) and not (cfg['_commercial'] and cfg['_commercial'].search(k)):
            infos[k] = v
        else:
            kept[k] = v

    keys = sorted(kept, key=lambda k: -kept[k])
    dup = mark([[k, kept[k]] for k in keys], 0, 1, None)
    dup_of = {keys[i]: h for i, h in dup.items()}
    merged = {g for g, d in (serp or {}).get('pages', {}).items() if d.get('decision') == 'merge'}

    steps = {f"{cfg['base_name']} (база)": []}
    cands = {k: candidates(feats[k], cfg, steps) for k in keys}
    chains_of = {}
    for cl in cands.values():
        for c in cl:
            chains_of.setdefault(c[0], c[1])
    page_of, meta = assign(keys, cands, kept, dup_of, cfg['thresholds'], excluded=merged)
    base = f"{cfg['base_name']} (база)"
    info_page = page_records(keys, page_of, meta, kept, dup_of, base, chains_of)

    # злиті перевіркою видачі сторінки з поточними ключами: serp_check.py перевіряє їх щоразу заново
    merged_pages = {}
    for k in keys:
        for c in cands[k]:
            if c[0] in merged and c[0] != page_of[k]:
                merged_pages.setdefault(c[0], {'c': c, 'keys': []})['keys'].append(k)
                break
    merged_info = []
    for g, m in merged_pages.items():
        name, parent, url, level, kind = m['c']
        p = parent if parent in info_page else base
        head = next((k for k in m['keys'] if k not in dup_of), m['keys'][0])
        merged_info.append({'page': g, 'parent': p, 'url': url, 'level': level, 'kind': kind, 'head': head, 'keys': len(m['keys']),
                            'volume': sum(kept[k] for k in m['keys'] if k not in dup_of), 'keywords': [[k, kept[k]] for k in m['keys']]})

    # статті з інфо-запитів
    ikeys = sorted(infos, key=lambda k: -infos[k])
    idup = mark([[k, infos[k]] for k in ikeys], 0, 1, None) if ikeys else {}
    idup_of = {ikeys[i]: h for i, h in idup.items()}
    asteps = {}
    acands = {k: article_candidates(k, feats[k], cfg, asteps) for k in ikeys}
    aparents = {c[0]: c[1] for cl in acands.values() for c in cl}
    apage_of, ameta = assign(ikeys, acands, infos, idup_of, cfg['articles']['thresholds'], exempt=('base',)) if ikeys else ({}, {})
    articles = page_records(ikeys, apage_of, ameta, infos, idup_of, None, aparents) if ikeys else {}
    for recs, st in ((info_page, steps), (articles, asteps)):
        for g, p in recs.items():
            p['steps'] = st.get(g, [])
    for m in merged_info:
        m['steps'] = steps.get(m['page'], [])
    return {'core': core, 'src': src, 'kept': kept, 'removed': removed, 'infos': infos, 'feats': feats,
            'dup_of': dup_of, 'pages': info_page, 'page_of': page_of, 'merged': merged, 'merged_pages': merged_info,
            'articles': articles, 'article_of': apage_of, 'info_dup_of': idup_of}


def ai_status(cfg, res, ai_rows):
    """Знахідки ai_review.py проти поточного ядра: враховано правилами, відхилено агентом або лишилось тімліду."""
    brands = [b['_rx'] for b in cfg.get('brands', [])]
    out = {'applied': 0, 'dismissed': 0, 'flagged': {}}
    for r in ai_rows:
        k, prob = r['Keyword'], r['Проблема']
        # враховано: ключ видалено або в інфо; для «інша сторінка» - ключ переїхав. Інші переїзди (перейменування
        # сторінки, модель замість бренду) знахідку не закривають.
        if k not in res['kept'] or (prob == 'інша сторінка' and res['page_of'].get(k) != r['Сторінка']):
            out['applied'] += 1
            continue
        dismiss = ((prob == 'інша сторінка' and not r.get('Пропозиція'))
                   or (prob == 'не про товар' and any(b.search(k) for b in brands))
                   or any(d['problem'] == prob and d['_rx'].search(k) for d in cfg.get('ai_dismiss', [])))
        if dismiss:
            out['dismissed'] += 1
        else:
            out['flagged'][k] = f"ШІ: {prob}{' -> ' + r['Пропозиція'] if r.get('Пропозиція') else ''} - {r.get('Чому', '')}"
    return out


# ---------- вихід ----------

def tree_order(pages):
    """Сторінки в порядку дерева: база, під нею гілки від найбільших за попитом."""
    children = defaultdict(list)
    roots = []
    for p in pages.values():
        if p['parent'] in pages:
            children[p['parent']].append(p)
        else:
            roots.append(p)
    out = []

    def walk(p):
        out.append(p)
        for c in sorted(children[p['page']], key=lambda c: -c['volume']):
            walk(c)
    for r in sorted(roots, key=lambda p: (p['kind'] != 'base', -p['volume'])):
        walk(r)
    return out


def competitors_summary(cfg):
    path = os.path.join(cfg['_dir'], 'competitors.json')
    if not os.path.exists(path):
        return None
    return json.load(open(path, encoding='utf-8'))


def write_outputs(cfg, res, out, serp=None, ai=None, articles_ex=None):
    head_font, head_fill = Font(bold=True, color='FFFFFF', name='Arial'), PatternFill('solid', fgColor='2F5597')
    opts = cfg.get('output', {})
    split = opts.get('split_lang', True)      # окремі аркуші UA і RU, як у робочому документі тімліда
    neutral = opts.get('neutral_lang', 'uk')  # куди йдуть ключі без ознак мови: латиниця, бренди, моделі
    site = opts.get('site', '').rstrip('/')   # домен клієнта: URL у таблиці стане повною адресою
    category = cfg.get('name', '')
    base_name = cfg['base_name']

    tree = tree_order(res['pages'])
    atree = tree_order(res['articles'])
    intents = cfg.get('intent_rules', [])
    comments = cfg.get('comments', [])
    sp = (serp or {}).get('pages', {})
    groups = {p['page']: group_parts(p, res['pages'], base_name) for p in tree}
    agroups = {p['page']: group_parts(p, res['articles'], base_name) for p in atree}
    n_group = max(len(g) for g in list(groups.values()) + list(agroups.values()) + [[base_name]])
    group_cols = [f'Group {i}' for i in range(1, n_group + 1)]

    def note_of(k, dup_of):
        note = [c['comment'] for c in comments if c['_rx'].search(k)]
        if k in dup_of:
            note.append(f'дубль «{dup_of[k]}», в ∑ не йде')
        if ai and k in ai['flagged']:
            note.append(ai['flagged'][k])
        return '; '.join(note)

    def flat_row(k, v, intent, parts, url, dup_of, extra=()):
        return ([category, k, v, intent] + parts + [''] * (n_group - len(parts))
                + [site + url if url else '', '; '.join(sorted(res['src'].get(k, ())))] + list(extra) + [note_of(k, dup_of)])

    # Кластери плоскою таблицею: рядок = ключ, шлях сторінки - у колонках Group 1..N («Бренд - EcoFlow», «Серія - Delta»)
    rows = defaultdict(list)
    for p in tree:
        for k, v in p['keywords']:
            intent = next((r['intent'] for r in intents if r['_rx'].search(k)), 'комерційний')
            lang = key_lang(k) or neutral
            if split:
                rows[lang].append(flat_row(k, v, intent, groups[p['page']], p['url'], res['dup_of']))
            else:
                rows['all'].append(flat_row(k, v, intent, groups[p['page']], p['url'], res['dup_of'], (lang,)))
    header = ['Категорія', 'Keyword', 'Volume', 'Intent'] + group_cols + ['URL', 'Source']
    wb = openpyxl.Workbook()
    ws_z = wb.active
    ws_z.title = 'Зведення'
    flat_sheets = []
    for lang in (sorted(rows, key=lambda x: (x not in LANG_SHEETS, list(LANG_SHEETS).index(x) if x in LANG_SHEETS else 0, x))
                 if split else ['all']):
        name = LANG_SHEETS.get(lang, lang.upper()) if split else 'Кластери'
        ws = wb.create_sheet(name)
        ws.append(header + ([] if split else ['Мова']) + ['Коментар'])
        for r in rows[lang]:
            ws.append(r)
        flat_sheets.append(name)

    ws_s = wb.create_sheet('Структура')
    ws_s.append(['Дерево', 'Рівень', 'Сторінка', 'Батьківська', 'URL', 'Ключів', '∑', 'Головний ключ', 'Перевірка видачі', 'Тип видачі'])
    for p in tree:
        d = sp.get(p['page'])
        if p['kind'] == 'base':
            check = ''
        elif p['kind'] == 'brand':
            check = 'бренд: окрема сторінка, не перевіряється'
        elif d:
            check = ('окрема: ' if d['decision'] == 'keep' else 'злито: ') + d.get('why', f'спільних {d["overlap"]} з 10')
        else:
            check = 'не перевірено'
        rec = (d or {}).get('recommendation', '')
        indent = '    ' * (p['level'] - 1) + rel_label(p)
        ws_s.append([indent, p['level'], p['page'], p['parent'] or '', site + p['url'], p['keys'], p['volume'], p['head'], check, rec])
    for g, d in sp.items():
        if d['decision'] == 'merge':
            ws_s.append(['', '', g, d['parent'], '', 0, 0, d['head'], 'злито: ' + d.get('why', ''), d.get('recommendation', '')])

    # Блог: інфо-запити тією самою плоскою таблицею, Group 2 - тема статті, Group 3 - її поділ
    ws_i = wb.create_sheet('Блог')
    ws_i.append(header + ['Коментар'])
    for p in atree:
        for k, v in p['keywords']:
            ws_i.append(flat_row(k, v, 'інформаційний', agroups[p['page']], p['url'], res['info_dup_of']))
    flat_sheets.append('Блог')

    ws_a = wb.create_sheet('Статті')
    ws_a.append(['Тема статті', 'Рівень', 'Ключів', '∑', 'Головний ключ', 'Що у видачі', 'Приклади статей у видачі', 'Ключі'])
    ex = (articles_ex or {}).get('pages', {})
    for p in atree:
        e = ex.get(p['page'], {})
        examples = ' | '.join(f'{x["title"] or x["url"]} - {x["url"]}' for x in e.get('examples', []))
        what = ', '.join(f'{k} {v}' for k, v in sorted(e.get('profile', {}).items(), key=lambda x: -x[1]))
        ws_a.append([p['page'], p['level'], p['keys'], p['volume'], p['head'], what, examples,
                     ', '.join(k for k, _ in p['keywords'][:15])])

    ws_r = wb.create_sheet('Видалено')
    ws_r.append(['Keyword', 'Volume', 'Причина'])
    for k, v, why in sorted(res['removed'], key=lambda x: (x[2], -x[1])):
        ws_r.append([k, v, why])

    comp = competitors_summary(cfg)
    if comp:
        ws_c = wb.create_sheet('Конкуренти')
        ws_c.append(['Конкурент', 'Категорія', 'Сторінок розділу', 'Фільтрів і підкатегорій', 'Товарів', 'Найсильніші фільтри (ключів у Serpstat)'])
        for r in comp['competitors']:
            top = sorted([p for p in r['pages'] if p['type'] == 'фільтр'], key=lambda p: -p['keywords'])[:6]
            ws_c.append([r['competitor'], r['category_url'], len(r['pages']), r['types'].get('фільтр', 0), r['types'].get('товар', 0),
                         '; '.join(f'{p["label"]} ({p["keywords"]})' for p in top)])
        ws_c.append([])
        ws_c.append(['Ознака', 'У скількох конкурентів', 'Ключів на сторінках', 'Приклади назв'])
        for t, n, k, exl in comp.get('facets', [])[:40]:
            ws_c.append([t, n, k, ' | '.join(exl)])

    total_in = len(res['core'])
    n_kept = len(res['kept'])
    levels = Counter(p['level'] for p in res['pages'].values())
    rem_by = Counter(w.split(':')[0] if w.startswith('гео') else w for _, _, w in res['removed'])
    summary = [
        ('Дата', date.today().isoformat()),
        ('На вході', total_in),
        ('Кластери', f'{n_kept} ключів, {len(res["pages"])} сторінок: ' + ', '.join(f'рівень {lv} - {levels[lv]}' for lv in sorted(levels))),
        ('Попит у кластерах', f'{sum(p["volume"] for p in res["pages"].values())} без дублів'),
        ('Статті', f'{len(res["infos"])} інфо-запитів, {len(res["articles"])} тем статей'),
        ('Видалено', len(res['removed'])),
        ('Баланс', f'{n_kept} + {len(res["infos"])} + {len(res["removed"])} = {n_kept + len(res["infos"]) + len(res["removed"])} (вхід {total_in})'),
        ('Дублів з тією ж частотою', len(res['dup_of'])),
        ('Поріг сторінки', f'>= {cfg["thresholds"]["min_keys"]} ключі і >= {cfg["thresholds"]["min_volume"]} на місяць'),
        ('Перевірка видачі', (f'{len(sp)} сторінок, злито {sum(1 for d in sp.values() if d["decision"] == "merge")}' if sp else 'не запускалась')),
        ('Перевірка ШІ', (f'{ai["total"]} знахідок: враховано правилами {ai["applied"]}, відхилено агентом {ai["dismissed"]}, '
                          f'на рішення тімліда {len(ai["flagged"])} (у колонці «Коментар»)') if ai else 'не запускалась'),
        ('Конкуренти', ', '.join(r['competitor'] for r in comp['competitors']) if comp else 'не аналізувались'),
        ('Причини видалення', '; '.join(f'{k}: {v}' for k, v in rem_by.most_common())),
    ]
    if split:
        summary.insert(3, ('Мови', ', '.join(f'{LANG_SHEETS.get(lg, lg.upper())} - {len(r)} ключів' for lg, r in rows.items())
                           + f' (без ознак мови - на аркуші {LANG_SHEETS.get(neutral, neutral.upper())})'))
    ws_z.append(['Параметр', 'Значення'])
    for a, b in summary:
        ws_z.append([a, b])

    flat_w = (16, 46, 9, 13) + (30,) * n_group + (44, 26, 50)
    widths = {**{n: flat_w for n in flat_sheets}, 'Структура': (40, 8, 36, 30, 44, 8, 10, 40, 40, 24),
              'Статті': (40, 8, 8, 10, 40, 30, 90, 80), 'Видалено': (46, 10, 50),
              'Конкуренти': (22, 50, 12, 12, 10, 90), 'Зведення': (28, 90)}
    for w in wb.worksheets:
        for c in w[1]:
            c.font, c.fill = head_font, head_fill
        w.freeze_panes = 'A2'
        if w.title != 'Конкуренти':
            w.auto_filter.ref = w.dimensions
        for i, width in enumerate(widths.get(w.title, ()), 1):
            w.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    wb.save(out)
    pages_json = os.path.splitext(out)[0] + '.pages.json'
    json.dump({'base': f"{cfg['base_name']} (база)", 'category': cfg.get('category_name', cfg['base_name']),
               'pages': list(res['pages'].values()), 'merged': res['merged_pages'], 'articles': list(res['articles'].values())},
              open(pages_json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    assert n_kept + len(res['infos']) + len(res['removed']) == total_in, 'баланс не сходиться'
    return summary, pages_json


def main():
    ap = argparse.ArgumentParser(description='Чистка і багаторівнева структура семантики за конфігом ніші')
    ap.add_argument('config', help='JSON-конфіг ніші')
    ap.add_argument('--serp', help='результат serp_check.py: сторінки, які злити з батьківськими')
    ap.add_argument('--ai', help='знахідки ai_review.py (csv): позначити в «Коментар» те, що не враховано правилами')
    ap.add_argument('--articles', help='результат article_examples.py: що у видачі тем статей і приклади статей')
    ap.add_argument('--out', help='xlsx (за замовчуванням <назва>_semantics_<дата>.xlsx поруч із конфігом)')
    args = ap.parse_args()
    cfg = load_config(args.config)
    serp = json.load(open(args.serp, encoding='utf-8')) if args.serp else None
    res = build(cfg, serp)
    out = args.out or os.path.join(cfg['_dir'], f"{cfg.get('name', 'semantics')}_semantics_{date.today().isoformat()}.xlsx")
    ai = None
    if args.ai:
        rows = list(csv.DictReader(open(args.ai, encoding='utf-8-sig')))
        ai = ai_status(cfg, res, rows)
        ai['total'] = len(rows)
    articles_ex = json.load(open(args.articles, encoding='utf-8')) if args.articles else None
    summary, pages_json = write_outputs(cfg, res, out, serp, ai, articles_ex)
    print('->', out)
    print('->', pages_json)
    for a, b in summary:
        print(f'{a}: {b}')
    print('\nНайбільші сторінки:')
    for p in sorted(res['pages'].values(), key=lambda p: -p['volume'])[:12]:
        print(f'  {p["volume"]:>8}  {p["keys"]:>5}  рівень {p["level"]}  {p["page"]}')
    print('\nТеми статей:')
    for p in sorted(res['articles'].values(), key=lambda p: -p['volume'])[:10]:
        print(f'  {p["volume"]:>8}  {p["keys"]:>5}  {p["page"]}')


if __name__ == '__main__':
    main()
