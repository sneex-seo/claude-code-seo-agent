# -*- coding: utf-8 -*-
"""Розмітка неявних дублів (варіанти порядку слів і словоформи) за правилом частотності.

Правило (14.10.2026, скіл semantic-clustering):
- однаковий набір токенів (закінчення відкинуто, порядок і прийменники неважливі) І однакова
  частотність -> це один запит: Google Ads склеїв варіанти, і Serpstat ставить усім одне число
  («палатка / палатки / палаткой» = 22 200). У ∑ рахуємо один раз, решта - «дубль»;
- та сама частотність і відмінність в одному слові на 1-2 літери (одрук) -> теж дубль;
- різна частотність -> різні запити, обидва йдуть у ∑.
Порівняння - всередині групи (колонка Group), якщо вона є, інакше по всьому файлу.

  python scripts/variant_dedupe.py ключі.csv
  python scripts/variant_dedupe.py ядро.xlsx --sheet "Кластери" --out ядро_dedupe.csv

Вихід: той самий список + колонки «Дубль» (так/ні) і «Головна фраза»; у консоль - ∑ до і після по групах.
Формат входу: csv/tsv або xlsx, колонки шукаються за заголовком (keyword/ключ..., volume/частотність..., group/кластер...).
"""
import argparse, csv, os, re, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

KW_COLS = ('keyword', 'ключ', 'ключевой запрос', 'ключове слово', 'фраза', 'query')
VOL_COLS = ('volume', 'vol', 'частотность', 'частотність', 'частота', 'region_queries_count')
GRP_COLS = ('group', 'cluster', 'кластер', 'група', 'группа')
STOP = {'для', 'на', 'в', 'у', 'з', 'с', 'и', 'і', 'та', 'по', 'от', 'від', 'до', 'за', 'из', 'із'}
ENDINGS = sorted('ами ями ого ому ими ыми ему ової евої ий ій ый ой ая яя ое ее ые ие ої ою ею ам ям ах ях ов ев ей ів їв ом ем '
                 'ую юю у ю а я о е и ы і ї ь й'.split(), key=len, reverse=True)


def stem(w):
    """палатками / палатки / палаткой -> палатк; термокружки -> термокружк (не «термо»: не резать до N букв)."""
    if not re.match(r'[а-яіїєґё]', w):
        return w
    for e in ENDINGS:
        if w.endswith(e) and len(w) - len(e) >= 4:
            return w[:-len(e)]
    return w


def sig(kw):
    words = re.findall(r'[a-zа-яіїєґё0-9]+', kw.lower().replace('ё', 'е'))
    return ' '.join(sorted(stem(w) for w in words if w not in STOP))


def lev(a, b):
    if abs(len(a) - len(b)) > 2:
        return 9
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def typo_pair(s1, s2):
    a, b = s1.split(), s2.split()
    if len(a) != len(b):
        return False
    ra, rb = list(a), list(b)
    for w in a:
        if w in rb:
            rb.remove(w); ra.remove(w)
    return len(ra) == 1 and len(ra[0]) >= 4 and lev(ra[0], rb[0]) <= 2


def _pick(header, options):
    for i, h in enumerate(header):
        if str(h or '').strip().lower() in options:
            return i
    return None


def load(path, sheet=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.csv', '.tsv'):
        rows = list(csv.reader(open(path, encoding='utf-8-sig'), delimiter='\t' if ext == '.tsv' else ','))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    head = [str(h or '') for h in rows[0]]
    ki, vi, gi = _pick(head, KW_COLS), _pick(head, VOL_COLS), _pick(head, GRP_COLS)
    if ki is None or vi is None:
        sys.exit(f'У заголовку немає колонок ключа і частотності: {head}')
    return head, rows[1:], ki, vi, gi


def mark(rows, ki, vi, gi):
    """Возвращает {индекс строки: головная фраза} для строк-дублей."""
    groups = defaultdict(list)
    for i, r in enumerate(rows):
        if len(r) > ki and r[ki]:
            groups[r[gi] if gi is not None and len(r) > gi else ''].append(i)
    dup = {}
    for idxs in groups.values():
        def vol(i):
            try:
                return float(rows[i][vi] or 0)
            except (TypeError, ValueError):
                return 0.0
        heads = {}  # (sig, vol) -> индекс головной строки
        for i in sorted(idxs, key=lambda i: -vol(i)):
            key = (sig(str(rows[i][ki])), vol(i))
            h = heads.get(key)
            if h is None:
                h = next((hi for (s, v), hi in heads.items() if v == key[1] and typo_pair(s, key[0])), None)
            if h is None:
                heads[key] = i
            else:
                dup[i] = str(rows[h][ki])
    return dup


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('path')
    ap.add_argument('--sheet')
    ap.add_argument('--out')
    a = ap.parse_args()
    head, rows, ki, vi, gi = load(a.path, a.sheet)
    dup = mark(rows, ki, vi, gi)
    tot_all = tot_dedup = 0.0
    by = defaultdict(lambda: [0.0, 0.0])
    for i, r in enumerate(rows):
        try:
            v = float(r[vi] or 0)
        except (TypeError, ValueError):
            v = 0.0
        g = r[gi] if gi is not None and len(r) > gi else ''
        by[g][0] += v
        tot_all += v
        if i not in dup:
            by[g][1] += v
            tot_dedup += v
    out = a.out or os.path.splitext(a.path)[0] + '_dedupe.csv'
    with open(out, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(head + ['Дубль', 'Головна фраза'])
        for i, r in enumerate(rows):
            w.writerow(list(r) + (['так', dup[i]] if i in dup else ['ні', '']))
    print(f'Ключів: {len(rows)}, дублів з тією ж частотою: {len(dup)}')
    print(f'∑ усіх варіантів: {tot_all:,.0f}  ∑ за правилом: {tot_dedup:,.0f}  (завищення ×{tot_all / max(tot_dedup, 1):.2f})')
    if gi is not None:
        top = sorted(by.items(), key=lambda x: -(x[1][0] - x[1][1]))[:10]
        print('Групи з найбільшою різницею:')
        for g, (s_all, s_d) in top:
            print(f'  {g}: {s_all:,.0f} → {s_d:,.0f}')
    print('Файл:', out)


if __name__ == '__main__':
    main()
