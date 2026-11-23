# -*- coding: utf-8 -*-
"""Аудит кластеризации: ищет ПРОПУЩЕННЫЕ бренды/атрибуты в «базовом» (catch-all) кластере.

После первого прохода rule-based кластеризации прогони это на своём ядре - покажет частые
токены, осевшие в «Кепки (база)/загальна категорія», которых, скорее всего, не хватает в
списках брендов/фасонов/атрибутов. Домен-агностично: для одежды всплывут бренды/фасоны,
для электроники - samsung/xiaomi/ips/дюйми/пам'ять и т.п. Добавь их в списки и перекластеризуй.

  python scripts/cluster_gaps.py clustered.xlsx              # авто: строки Group ~ "база/загальн"
  python scripts/cluster_gaps.py clustered.xlsx --base "(база)"
  python scripts/cluster_gaps.py keys.csv --all              # весь файл как множество для скана
"""
import argparse, csv, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')

# служебные/коммерческие слова - не кандидаты в кластеры
STOP = set('''и в на с к по для от до за не или та й що чи the a an of and for + -
купить купити куплю цена ціна вартість недорого дешево дешеві магазин заказать замовити
заказ замовлення доставка отзывы відгуки интернет інтернет онлайн online shop买 україна украина'''.split())

def norm(s): return re.sub(r'\s+', ' ', str(s).lower().replace('ё', 'е').strip())

def load(path, base_pat, use_all):
    ext = os.path.splitext(path)[1].lower()
    out = []
    if ext in ('.xlsx', '.xlsm'):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb['Кластери'] if 'Кластери' in wb.sheetnames else (
             wb['Кластеры'] if 'Кластеры' in wb.sheetnames else wb.worksheets[0])
        it = ws.iter_rows(values_only=True)
        header = [str(h or '').strip().lower() for h in next(it)]
        ki = header.index('keyword') if 'keyword' in header else 0
        vi = header.index('volume') if 'volume' in header else 1
        gi = header.index('group') if 'group' in header else None
        rx = re.compile(base_pat, re.I)
        for r in it:
            if not r or len(r) <= ki or r[ki] is None or str(r[ki]).startswith('▸'):
                continue
            g = str(r[gi]) if (gi is not None and len(r) > gi and r[gi]) else ''
            if not use_all and gi is not None and not rx.search(g):
                continue
            v = r[vi] if (len(r) > vi and isinstance(r[vi], (int, float))) else 0
            out.append((norm(r[ki]), int(v)))
    else:
        for r in csv.DictReader(open(path, encoding='utf-8-sig')):
            kw = r.get('keyword') or r.get('Keyword') or next(iter(r.values()))
            try: v = int(float(r.get('volume') or r.get('Volume') or 0))
            except (ValueError, TypeError): v = 0
            out.append((norm(kw), v))
    return out

def main():
    ap = argparse.ArgumentParser(description='Аудит catch-all кластера: пропущені бренди й атрибути')
    ap.add_argument('path', help='кластеризований xlsx (аркуш «Кластери») або csv ключів')
    ap.add_argument('--base', default=r'\(баз|загальна категор|catch|general|\bбаза\b', help='regex catch-all кластера в Group (за замовчуванням ловить «... (база)» і «загальна категорія», але НЕ справжні фасети на кшталт «Загальна > Панами»)')
    ap.add_argument('--all', action='store_true', help='сканувати весь файл, а не лише базовий кластер')
    ap.add_argument('--top', type=int, default=30)
    ap.add_argument('--minlen', type=int, default=3)
    args = ap.parse_args()

    keys = load(args.path, args.base, args.all)
    if not keys:
        raise SystemExit('немає ключів (перевір шлях, --base або --all)')

    from collections import Counter
    tf = Counter(t for kw, _ in keys for t in set(kw.split()))
    ubiq = {t for t, c in tf.items() if c > 0.25 * len(keys)}  # вездесущие = слово категории/пол

    lat_c, lat_v, cyr_c, cyr_v = Counter(), Counter(), Counter(), Counter()
    for kw, v in keys:
        for t in set(kw.split()):
            if len(t) < args.minlen or t.isdigit() or t in STOP or t in ubiq:
                continue
            if re.fullmatch(r"[a-z][a-z0-9'’.\-]+", t):
                lat_c[t] += 1; lat_v[t] += v
            elif re.search(r'[а-яіїєґ]', t):
                cyr_c[t] += 1; cyr_v[t] += v

    print(f'Ключів у скані: {len(keys)} | фільтр: {"весь файл" if args.all else "база ~ "+args.base}')
    print(f'Виключено як усюдисущі (слово категорії або стать): {sorted(ubiq)[:12]}')
    print(f'\n=== ЛАТИНСЬКІ токени (кандидати в БРЕНДИ), топ-{args.top} [кількість | ∑vol] ===')
    for t, n in lat_c.most_common(args.top): print(f'  {n:>4}  ∑{lat_v[t]:>8}  {t}')
    print(f'\n=== КИРИЛИЧНІ токени (бренди, фасони, атрибути), топ-{args.top} ===')
    for t, n in cyr_c.most_common(args.top): print(f'  {n:>4}  ∑{cyr_v[t]:>8}  {t}')
    print('\n→ Часті токени в «базі» - імовірно, кластери, яких бракує. Додай їх у списки брендів, '
          'фасонів і атрибутів та перекластеризуй (межі слова обов\'язкові: jack≠panama jack).')

if __name__ == '__main__':
    main()
