# -*- coding: utf-8 -*-
"""Собирает боевые минус-листы из компонентов (единый источник правды).
Компоненты (reference/): geo/geo_exclude.txt, geo/geo_ua.txt, market_junk.txt
Генерит:
  minus_geo_market.txt      = geo_exclude + geo_ua + market_junk    (semantic-clustering, e-commerce)
  minus_services.txt        = geo_exclude + market_junk             (услуги: БЕЗ городов UA)
Правь компоненты, потом прогони: python scripts/build_minus.py
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
REF = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'reference')

def load(rel):
    return [l.strip().lower() for l in open(os.path.join(REF, rel), encoding='utf-8')
            if l.strip() and not l.strip().startswith('#')]

def write(rel, header, words):
    seen, uniq = set(), []
    for w in words:
        if w not in seen: seen.add(w); uniq.append(w)
    open(os.path.join(REF, rel), 'w', encoding='utf-8').write(header + '\n'.join(uniq) + '\n')
    return len(uniq)

geo_ex, geo_ua = load('geo/geo_exclude.txt'), load('geo/geo_ua.txt')
market = load('market_junk.txt')
H = '# АВТОГЕНЕРАЦІЯ scripts/build_minus.py з компонентів reference/. Руками НЕ правити - правте компоненти.\n'
n1 = write('minus_geo_market.txt', H + '# = geo_exclude + geo_ua + market_junk (semantic-clustering, e-commerce)\n', geo_ex + geo_ua + market)
n3 = write('minus_services.txt', H + '# = geo_exclude + market_junk (послуги: БЕЗ міст UA - місто цільове)\n', geo_ex + market)
print(f'компоненти: geo_exclude {len(set(geo_ex))} | geo_ua {len(set(geo_ua))} | market_junk {len(set(market))}')
print(f'-> minus_geo_market.txt: {n1} | minus_services.txt: {n3}')
E, U, M = set(geo_ex), set(geo_ua), set(market)
print('перевірка: москва в exclude', 'москва' in E, '| київ в ua', 'київ' in U, '| розетка в market', 'розетка' in M,
      '| москва НЕ в ua', 'москва' not in U, '| київ НЕ в базі послуг', 'київ' not in (E | M))
