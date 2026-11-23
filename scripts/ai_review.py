# -*- coding: utf-8 -*-
"""Перевірка ядра другою моделлю (ChatGPT через OpenAI API): кожен ключ читається в контексті своєї сторінки.

Правила й регулярки не бачать смислу, тому після semantic_core.py і serp_check.py ключі пачками
йдуть у модель з інструкцією reference/ai_review_prompt.md. Модель повертає лише проблемні ключі:
не про товар, краще на іншій сторінці, інформаційний, незрозумілий. Скрипт нічого не змінює в ядрі:
знахідки перевіряє агент і переносить підтверджені в правила конфігу ніші, потім перезбирає ядро.

Ключ: OPENAI_API_KEY (.env або змінні оточення), модель: OPENAI_MODEL або --model (за замовчуванням gpt-4.1).
Відповіді кешуються в cache/ai/: повторний прогін тієї самої пачки безкоштовний.

  python scripts/ai_review.py clients/shop/semantics/shop_semantics_2026-11-23.pages.json --dry-run
  python scripts/ai_review.py clients/shop/semantics/shop_semantics_2026-11-23.pages.json --max-usd 3
"""
import argparse
import csv
import hashlib
import json
import os
import sys
from collections import Counter

import requests

sys.stdout.reconfigure(encoding='utf-8')
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(ROOT, 'cache', 'ai')
PROMPT = os.path.join(ROOT, 'reference', 'ai_review_prompt.md')
PRICES = {'gpt-4.1': (2.0, 8.0), 'gpt-4.1-mini': (0.4, 1.6), 'gpt-4.1-nano': (0.1, 0.4),
          'gpt-4o': (2.5, 10.0), 'gpt-4o-mini': (0.15, 0.6)}  # $ за 1 млн токенів: вхід, вихід (оцінка)
LABELS = {'off_topic': 'не про товар', 'other_page': 'інша сторінка', 'info': 'інформаційний', 'unclear': 'незрозумілий'}


def env(key, default=None):
    path = os.path.join(ROOT, '.env')
    if os.path.exists(path):
        for line in open(path, encoding='utf-8'):
            if line.strip().startswith(key + '='):
                val = line.split('=', 1)[1].strip()
                if val:
                    return val
    return os.environ.get(key) or default


def plain(text):
    """Типографіка моделі -> звичайні символи: дефіс замість тире, «...» замість трикрапки."""
    return str(text).replace('—', '-').replace('–', '-').replace('…', '...').replace('“', '«').replace('”', '»')


def batches(pages, size):
    """Пачки до size ключів; сторінка не розривається, якщо влазить; велика ріжеться на шматки."""
    out, cur, n = [], [], 0
    for p in pages:
        kws = p['keywords']
        for i in range(0, len(kws), size):
            part = kws[i:i + size]
            if n + len(part) > size and cur:
                out.append(cur)
                cur, n = [], 0
            cur.append((p['page'], part))
            n += len(part)
    if cur:
        out.append(cur)
    return out


def batch_text(batch):
    lines = []
    for page, kws in batch:
        lines.append(f'## {page}')
        lines += [f'{k} | {v}' for k, v in kws]
    return '\n'.join(lines)


def cost_of(model, usage):
    p = PRICES.get(model)
    if not p or not usage:
        return 0.0
    return usage.get('prompt_tokens', 0) / 1e6 * p[0] + usage.get('completion_tokens', 0) / 1e6 * p[1]


def ask(model, system, user, key):
    r = requests.post('https://api.openai.com/v1/chat/completions', timeout=300,
                      headers={'Authorization': f'Bearer {key}'},
                      json={'model': model, 'temperature': 0, 'response_format': {'type': 'json_object'},
                            'messages': [{'role': 'system', 'content': system}, {'role': 'user', 'content': user}]})
    data = r.json()
    if 'error' in data:
        raise SystemExit(f'OpenAI: {data["error"].get("message")}')
    return data['choices'][0]['message']['content'], data.get('usage', {})


def main():
    ap = argparse.ArgumentParser(description='Перевірка ключів ядра моделлю OpenAI')
    ap.add_argument('pages', help='<out>.pages.json від semantic_core.py')
    ap.add_argument('--model', default=env('OPENAI_MODEL', 'gpt-4.1'))
    ap.add_argument('--batch-keys', type=int, default=250, help='ключів в одному запиті (250)')
    ap.add_argument('--max-usd', type=float, default=3.0, help='зупинитися, коли оцінка витрат дійде до ліміту')
    ap.add_argument('--out', help='csv зі знахідками (за замовчуванням ai_review.csv поруч)')
    ap.add_argument('--dry-run', action='store_true', help='лише порахувати пачки, без API')
    args = ap.parse_args()

    data = json.load(open(args.pages, encoding='utf-8'))
    pages = sorted(data['pages'], key=lambda p: (p['level'] == 0, p['page']))
    names = [p['page'] for p in pages]
    system = open(PROMPT, encoding='utf-8').read().replace('{category}', data.get('category', '')).replace('{pages}', '\n'.join(names))
    bs = batches(pages, args.batch_keys)
    os.makedirs(CACHE, exist_ok=True)
    keys_total = sum(len(p['keywords']) for p in pages)
    print(f'модель {args.model}: {keys_total} ключів, {len(bs)} запитів по <= {args.batch_keys} ключів')
    if args.dry_run:
        return
    key = env('OPENAI_API_KEY')
    if not key:
        raise SystemExit('OPENAI_API_KEY не знайдено ні в .env, ні в змінних оточення')

    vol = {k: v for p in pages for k, v in p['keywords']}
    page_of = {k: p['page'] for p in pages for k, _ in p['keywords']}
    findings, spent, live, usage_all = [], 0.0, 0, Counter()
    for i, b in enumerate(bs, 1):
        user = batch_text(b)
        h = hashlib.sha1((args.model + system + user).encode('utf-8')).hexdigest()[:20]
        path = os.path.join(CACHE, f'{h}.json')
        if os.path.exists(path):
            d = json.load(open(path, encoding='utf-8'))
        else:
            if spent >= args.max_usd:
                print(f'ліміт ${args.max_usd:.2f} вичерпано: перевірено {i - 1} з {len(bs)} пачок')
                break
            content, usage = ask(args.model, system, user, key)
            d = {'content': content, 'usage': usage}
            json.dump(d, open(path, 'w', encoding='utf-8'), ensure_ascii=False)
            spent += cost_of(args.model, usage)
            live += 1
            usage_all.update({k: v for k, v in usage.items() if isinstance(v, int)})
            print(f'  пачка {i}/{len(bs)}, витрачено ~${spent:.3f}', flush=True)
        try:
            items = json.loads(d['content']).get('findings', [])
        except json.JSONDecodeError:
            print(f'  пачка {i}: відповідь не JSON, пропущено')
            continue
        batch_keys = {k for _, kws in b for k, _ in kws}
        for it in items:
            k = str(it.get('keyword', '')).strip().lower()
            if k not in batch_keys:
                continue  # модель вигадала або змінила ключ
            sug = it.get('suggest') or ''
            findings.append({'Keyword': k, 'Volume': vol[k], 'Сторінка': page_of[k],
                             'Проблема': LABELS.get(it.get('problem'), it.get('problem')),
                             'Пропозиція': sug if sug in names else '', 'Чому': plain(it.get('why', ''))})

    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.pages)), 'ai_review.csv')
    with open(out, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=['Keyword', 'Volume', 'Сторінка', 'Проблема', 'Пропозиція', 'Чому'])
        w.writeheader()
        w.writerows(sorted(findings, key=lambda x: (x['Проблема'], -x['Volume'])))
    by = Counter(x['Проблема'] for x in findings)
    print(f'-> {out}')
    print(f'знахідок: {len(findings)} з {keys_total} ключів: ' + ', '.join(f'{k} {v}' for k, v in by.most_common()))
    print(f'нових запитів до API: {live}, токенів: вхід {usage_all["prompt_tokens"]}, вихід {usage_all["completion_tokens"]}, '
          f'оцінка ${spent:.3f}')
    print('Знахідки - підказка, а не рішення: перевірте їх і перенесіть підтверджені в правила конфігу.')


if __name__ == '__main__':
    main()
