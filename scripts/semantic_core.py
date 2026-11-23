# -*- coding: utf-8 -*-
"""Збір, чистка і фасетна кластеризація семантики за конфігом ніші (скіл semantic-clustering).

Конфіг ніші (JSON) пише агент: звідки брати ключі, правила чистки, бренди й моделі, числові ознаки
(потужність, ємність, розмір), атрибути (призначення, тип), пороги сторінки. Скрипт детерміновано:
  1. зливає вхідні csv (keyword, volume) з максимумом частотності;
  2. відсіює гео, майданчики й сміття (мінус-листи), чужі теми (topic_rules), інфо-запити;
  3. позначає дублі з однаковою частотністю (variant_dedupe) - вони в кластері, але не в ∑;
  4. для кожного ключа знаходить ознаки і кандидатів у сторінки від найдетальнішої до загальної:
     бренд > модель, бренд > серія, бренд > потужність, бренд; поєднання атрибутів і чисел
     («Для дому > 2 кВт»); число окремо; атрибут окремо; база категорії;
  5. кладе ключ на найдетальнішу сторінку, де >= min_keys ключів і >= min_volume попиту,
     крім сторінок, які перевірка видачі (serp_check.py) звелила злити з батьківською;
  6. пише xlsx (Кластери, Структура, Інфо, Видалено, Зведення) і <out>.pages.json для перевірок.

  python scripts/semantic_core.py clients/shop/semantics/config.json
  python scripts/semantic_core.py clients/shop/semantics/config.json --serp clients/shop/semantics/serp_check.json
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
BASE_LEVEL = 0


def norm(s):
    return re.sub(r'\s+', ' ', str(s).lower().replace('ё', 'е').strip())


def rx(p):
    return re.compile(p, re.I)


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
    for r in cfg.get('ai_dismiss', []):
        r['_rx'] = rx(r['pattern'])
    return cfg


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
            for m in b['_models']:
                hit = m.search(kk)
                if hit:
                    f['model'] = model_label(hit.group(0))
                    break
            if 'model' not in f:
                for name, slug, s in b['_series']:
                    if s.search(kk):
                        f['series'] = (name, slug)
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


def candidates(f, cfg):
    """(сторінка, батьківська, url, рівень) від найдетальнішої до бази."""
    cat = cfg['category_url']
    base = f"{cfg['base_name']} (база)"
    out = []
    if 'brand' in f and not f.get('override'):
        bn, bs = f['brand']
        if 'model' in f:
            out.append((f'{bn} > {f["model"][0]}', bn, f'{cat}{bs}/{f["model"][1]}/', 2, 'model'))
        if 'series' in f:
            out.append((f'{bn} > {f["series"][0]}', bn, f'{cat}{bs}/{f["series"][1]}/', 2, 'series'))
        for n in cfg.get('numeric', []):
            v = f.get('num:' + n['name'])
            if v and n.get('with_brand', True):
                out.append((f'{bn} > {v[1]}', bn, f'{cat}{bs}/{v[2]}/', 2, 'brand_num'))
        out.append((bn, base, f'{cat}{bs}/', 1, 'brand'))
        out.append((base, None, cat, BASE_LEVEL, 'base'))
        return out

    def val(facet):
        if ('attr:' + facet) in f:
            return f['attr:' + facet]
        if ('num:' + facet) in f:
            return f['num:' + facet][1:]
        return None

    def single(facet):
        if ('attr:' + facet) in f:
            return f['attr:' + facet][0]
        if ('num:' + facet) in f:
            return f'{facet} > {f["num:" + facet][1]}'
        return None

    for a, b in cfg.get('combos', []):
        va, vb = val(a), val(b)
        if va and vb:
            out.append((f'{va[0]} > {vb[0]}', single(a), f'{cat}{va[1]}/{vb[1]}/', 2, 'combo'))
    for n in cfg.get('numeric', []):
        v = f.get('num:' + n['name'])
        if v:
            out.append((f'{n["name"]} > {v[1]}', base, f'{cat}{v[2]}/', 1, 'num'))
    for a in cfg.get('attributes', []):
        v = f.get('attr:' + a['name'])
        if v:
            out.append((v[0], base, f'{cat}{v[1]}/', 1, 'attr'))
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
            tag = 'сіди' if os.path.basename(path).startswith('seed') else 'конкуренти'
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


# ---------- головне ----------

def build(cfg, serp=None):
    core, src = read_inputs(cfg)
    minus = []
    for m in cfg.get('minus_lists', []):
        minus += load_minus(os.path.join(ROOT, m))
    info = load_minus(os.path.join(ROOT, 'reference', 'info_markers.txt')) + cfg.get('info_extra', [])

    kept, removed, infos, feats = {}, [], {}, {}
    for k, v in core.items():
        f = facets(k, cfg)
        why = removal_reason(k, f, cfg, minus)
        if why:
            removed.append((k, v, why))
        elif is_minus(k, info):
            infos[k] = v
        else:
            kept[k] = v
            feats[k] = f

    keys = sorted(kept, key=lambda k: -kept[k])
    dup = mark([[k, kept[k]] for k in keys], 0, 1, None)
    dup_of = {keys[i]: h for i, h in dup.items()}

    merged = set()
    if serp:
        merged = {g for g, d in serp.get('pages', {}).items() if d.get('decision') == 'merge'}

    cands = {k: candidates(feats[k], cfg) for k in keys}
    th = cfg['thresholds']
    # Ключ іде на найдетальнішу сторінку, яку не виключено. Сторінки, де фактично лишилось менше порогу,
    # виключаються, і їхні ключі піднімаються на рівень вище; повторюємо, доки нічого не зміниться.
    # Сторінки брендів лишаються навіть малими (правило скілу), злиті перевіркою видачі - ні.
    excluded = set(merged)
    for _ in range(20):
        page_of, meta = {}, {}
        for k in keys:
            c = next(c for c in cands[k] if c[4] == 'base' or c[0] not in excluded)
            page_of[k] = c[0]
            meta[c[0]] = c
        n_k, n_v = Counter(), Counter()
        for k in keys:
            n_k[page_of[k]] += 1
            if k not in dup_of:
                n_v[page_of[k]] += kept[k]
        small = {g for g, c in meta.items() if c[4] not in ('base', 'brand')
                 and (n_k[g] < th['min_keys'] or n_v[g] < th['min_volume'])}
        if not small:
            break
        excluded |= small

    pages = defaultdict(list)
    for k in keys:
        pages[page_of[k]].append(k)
    # батьківська сторінка має існувати; якщо її немає, батько - база
    base = f"{cfg['base_name']} (база)"
    info_page = {}
    for g, (name, parent, url, level, kind) in meta.items():
        if parent is not None and parent not in pages:
            parent = base
        head = next((k for k in pages[g] if k not in dup_of), pages[g][0])
        info_page[g] = {'page': g, 'parent': parent, 'url': url, 'level': level, 'kind': kind, 'head': head,
                        'keys': len(pages[g]), 'volume': sum(kept[k] for k in pages[g] if k not in dup_of),
                        'keywords': [[k, kept[k]] for k in pages[g]]}
    # Злиті перевіркою видачі сторінки: з поточними ключами, щоб serp_check.py перевіряв їх щоразу заново
    merged_pages = {}
    for k in keys:
        for c in cands[k]:
            if c[0] in merged and c[0] != page_of[k]:
                merged_pages.setdefault(c[0], {'c': c, 'keys': []})['keys'].append(k)
                break
    merged_info = []
    for g, m in merged_pages.items():
        name, parent, url, level, kind = m['c']
        if parent is not None and parent not in pages:
            parent = base
        head = next((k for k in m['keys'] if k not in dup_of), m['keys'][0])
        merged_info.append({'page': g, 'parent': parent, 'url': url, 'level': level, 'kind': kind, 'head': head,
                            'keys': len(m['keys']), 'volume': sum(kept[k] for k in m['keys'] if k not in dup_of),
                            'keywords': [[k, kept[k]] for k in m['keys']]})
    return {'core': core, 'src': src, 'kept': kept, 'removed': removed, 'infos': infos, 'feats': feats,
            'dup_of': dup_of, 'pages': info_page, 'page_of': page_of, 'merged': merged, 'merged_pages': merged_info}


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
        dismiss = (prob == 'інша сторінка' and not r.get('Пропозиція')) or                   (prob == 'не про товар' and any(b.search(k) for b in brands)) or                   any(d['problem'] == prob and d['_rx'].search(k) for d in cfg.get('ai_dismiss', []))
        if dismiss:
            out['dismissed'] += 1
        else:
            out['flagged'][k] = f"ШІ: {prob}{' -> ' + r['Пропозиція'] if r.get('Пропозиція') else ''} - {r.get('Чому', '')}"
    return out


def write_outputs(cfg, res, out, serp=None, ai=None):
    bold, fill = Font(bold=True), PatternFill('solid', fgColor='E8F0E8')
    facet_cols = (['Бренд', 'Модель'] + [n['name'] for n in cfg.get('numeric', [])]
                  + [a['name'] for a in cfg.get('attributes', [])])

    def facet_vals(f):
        vals = [f.get('brand', ('',))[0], (f.get('model') or f.get('series') or ('',))[0]]
        vals += [f.get('num:' + n['name'], (None, ''))[1] for n in cfg.get('numeric', [])]
        vals += [f.get('attr:' + a['name'], ('',))[0] for a in cfg.get('attributes', [])]
        return vals

    intents = cfg.get('intent_rules', [])
    comments = cfg.get('comments', [])
    order = sorted(res['pages'].values(), key=lambda p: (p['level'] == BASE_LEVEL, p['parent'] or '', -p['volume']))
    # дерево: батьківські сторінки, за ними дочірні
    tree = []
    seen = set()
    for r in sorted([p for p in order if p['level'] == 1], key=lambda p: -p['volume']):
        tree.append(r); seen.add(r['page'])
        for c in sorted([p for p in order if p['parent'] == r['page'] and p['level'] == 2], key=lambda p: -p['volume']):
            tree.append(c); seen.add(c['page'])
    tree += [p for p in order if p['page'] not in seen]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Кластери'
    ws.append(['Keyword', 'Volume', 'Group', 'URL', 'Інтент', 'Коментар'] + facet_cols)
    for p in tree:
        ws.append([f'▸ {p["page"]}', p['volume'], p['page'], 'створити: ' + p['url'], '', f'ключів: {p["keys"]}'])
        for c in ws[ws.max_row]:
            c.font, c.fill = bold, fill
        for k, v in p['keywords']:
            note = []
            for c in comments:
                if c['_rx'].search(k):
                    note.append(c['comment'])
            if k in res['dup_of']:
                note.append(f'дубль «{res["dup_of"][k]}», в ∑ не йде')
            if ai and k in ai['flagged']:
                note.append(ai['flagged'][k])
            intent = next((r['intent'] for r in intents if r['_rx'].search(k)), 'комерційний')
            ws.append([k, v, p['page'], p['url'], intent, '; '.join(note)] + facet_vals(res['feats'][k]))

    ws_s = wb.create_sheet('Структура')
    ws_s.append(['Рівень', 'Сторінка', 'Батьківська', 'URL', 'Ключів', '∑', 'Головний ключ', 'Перевірка видачі'])
    sp = (serp or {}).get('pages', {})
    for p in tree:
        d = sp.get(p['page'])
        check = ''
        if p['level'] == BASE_LEVEL:
            check = ''
        elif p.get('kind') == 'brand':
            check = 'бренд: окрема сторінка, не перевіряється'
        elif d:
            check = f'окрема: спільних {d["overlap"]} з 10' if d['decision'] == 'keep' else f'злито: спільних {d["overlap"]} з 10'
        else:
            check = 'не перевірено'
        ws_s.append([p['level'], p['page'], p['parent'] or '', p['url'], p['keys'], p['volume'], p['head'], check])
    for g, d in sp.items():
        if d['decision'] == 'merge':
            ws_s.append(['', g, d['parent'], '', 0, 0, d['head'], f'злито з «{d["parent"]}»: спільних {d["overlap"]} з 10'])

    ws_i = wb.create_sheet('Інфо')
    ws_i.append(['Keyword', 'Volume', 'Group', 'URL', 'Інтент', 'Коментар'])
    for k, v in sorted(res['infos'].items(), key=lambda x: -x[1]):
        ws_i.append([k, v, 'Інфо', 'блог / FAQ', 'інформаційний', ''])
    ws_r = wb.create_sheet('Видалено')
    ws_r.append(['Keyword', 'Volume', 'Причина'])
    for k, v, why in sorted(res['removed'], key=lambda x: (x[2], -x[1])):
        ws_r.append([k, v, why])

    total_in = len(res['core'])
    n_kept = len(res['kept'])
    levels = Counter(p['level'] for p in res['pages'].values())
    rem_by = Counter(w.split(':')[0] if w.startswith('гео') else w for _, _, w in res['removed'])
    summary = [
        ('Дата', date.today().isoformat()),
        ('На вході', total_in),
        ('Кластери', f'{n_kept} ключів, {len(res["pages"])} сторінок: 1-го рівня {levels[1]}, 2-го рівня {levels[2]}, база 1'),
        ('Попит у кластерах', f'{sum(p["volume"] for p in res["pages"].values())} без дублів'),
        ('Інфо', len(res['infos'])),
        ('Видалено', len(res['removed'])),
        ('Баланс', f'{n_kept} + {len(res["infos"])} + {len(res["removed"])} = {n_kept + len(res["infos"]) + len(res["removed"])} (вхід {total_in})'),
        ('Дублів з тією ж частотою', len(res['dup_of'])),
        ('Поріг сторінки', f'>= {cfg["thresholds"]["min_keys"]} ключі і >= {cfg["thresholds"]["min_volume"]} на місяць'),
        ('Перевірка видачі', (f'{len(sp)} сторінок, злито {sum(1 for d in sp.values() if d["decision"] == "merge")}'
                              if sp else 'не запускалась')),
        ('Перевірка ШІ', (f'{ai["total"]} знахідок: враховано правилами {ai["applied"]}, відхилено агентом {ai["dismissed"]}, '
                          f'на рішення тімліда {len(ai["flagged"])} (у колонці «Коментар»)') if ai else 'не запускалась'),
        ('Причини видалення', '; '.join(f'{k}: {v}' for k, v in rem_by.most_common())),
    ]
    ws_z = wb.create_sheet('Зведення')
    ws_z.append(['Параметр', 'Значення'])
    for a, b in summary:
        ws_z.append([a, b])

    widths = {'Кластери': (46, 10, 34, 44, 14, 40), 'Структура': (8, 36, 30, 44, 8, 10, 40, 30),
              'Інфо': (46, 10, 10, 12, 14, 20), 'Видалено': (46, 10, 50), 'Зведення': (28, 90)}
    for w in wb.worksheets:
        for c in w[1]:
            c.font = bold
        w.freeze_panes = 'A2'
        w.auto_filter.ref = w.dimensions
        for i, width in enumerate(widths.get(w.title, ()), 1):
            w.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    wb.save(out)
    pages_json = os.path.splitext(out)[0] + '.pages.json'
    json.dump({'base': f"{cfg['base_name']} (база)", 'category': cfg.get('category_name', cfg['base_name']),
               'pages': list(res['pages'].values()), 'merged': res['merged_pages']},
              open(pages_json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    assert n_kept + len(res['infos']) + len(res['removed']) == total_in, 'баланс не сходиться'
    return summary, pages_json


def main():
    ap = argparse.ArgumentParser(description='Чистка і фасетна кластеризація семантики за конфігом ніші')
    ap.add_argument('config', help='JSON-конфіг ніші')
    ap.add_argument('--serp', help='результат serp_check.py: сторінки, які злити з батьківськими')
    ap.add_argument('--ai', help='знахідки ai_review.py (csv): позначити в «Коментар» те, що не враховано правилами')
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
    summary, pages_json = write_outputs(cfg, res, out, serp, ai)
    print('->', out)
    print('->', pages_json)
    for a, b in summary:
        print(f'{a}: {b}')
    top = sorted(res['pages'].values(), key=lambda p: -p['volume'])[:15]
    print('\nНайбільші сторінки:')
    for p in top:
        print(f'  {p["volume"]:>8}  {p["keys"]:>5}  {p["page"]}')


if __name__ == '__main__':
    main()
