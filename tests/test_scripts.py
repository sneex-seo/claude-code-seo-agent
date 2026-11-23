"""Перевірки скриптів на синтетичних прикладах з examples/. Ключі API не потрібні."""
import csv
import os
import subprocess
import sys

import openpyxl
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, 'scripts')
EXAMPLES = os.path.join(ROOT, 'examples')
sys.path.insert(0, SCRIPTS)


def run(script, *args):
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    res = subprocess.run([sys.executable, os.path.join(SCRIPTS, script), *args],
                         capture_output=True, text=True, encoding='utf-8', env=env, cwd=ROOT)
    assert res.returncode == 0, res.stderr
    return res.stdout


def write_csv(path, header, rows):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def test_is_minus_matches_whole_words():
    from serpstat_api import is_minus
    minus = ['київ', 'б/у', 'де купити']
    assert is_minus('кава київ', minus)
    assert is_minus('де купити каву', minus)
    assert not is_minus('кава кияни', minus)
    assert not is_minus('сукня на свадьбу', minus)


def test_variant_dedupe_same_volume_is_duplicate():
    from variant_dedupe import mark
    rows = [['палатка', 22200, 'Палатки'], ['палатки', 22200, 'Палатки'],
            ['палаткой', 22200, 'Палатки'], ['палатки туристичні', 1300, 'Палатки'],
            ['туристичні палатки', 880, 'Палатки']]
    dup = mark(rows, 0, 1, 2)
    assert set(dup) == {1, 2}
    assert dup[1] == 'палатка'


def test_coverage_check_counts_missing(tmp_path):
    core = tmp_path / 'core.csv'
    target = tmp_path / 'rival.csv'
    write_csv(core, ['keyword', 'volume'], [['кава в зернах', 1000], ['арабіка', 500]])
    write_csv(target, ['keyword', 'volume'], [['кава в зернах', 1000], ['арабіка', 500], ['робуста', 500]])
    out = run('coverage_check.py', str(core), str(target))
    assert 'Покриття ключів: 67%' in out
    assert 'Покриття обсягу: 75%' in out
    assert 'робуста' in out


def test_gsc_analyze_builds_all_sheets(tmp_path):
    out = tmp_path / 'gsc.xlsx'
    run('gsc_analyze.py', os.path.join(EXAMPLES, 'gsc_current.csv'),
        '--previous', os.path.join(EXAMPLES, 'gsc_previous.csv'), '--out', str(out))
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ['Зведення', 'Quick wins', 'Канібалізація', 'Нові', 'Втрачені',
                             'Топ сторінок', 'Топ запитів']
    cann = [r[0] for r in wb['Канібалізація'].iter_rows(min_row=2, values_only=True)]
    assert 'кава для турки' in cann
    lost = [r for r in wb['Втрачені'].iter_rows(min_row=2, values_only=True)]
    assert len(lost) == 2


def test_link_risk_flags_spam_and_skips_lost(tmp_path):
    out, disavow = tmp_path / 'links.xlsx', tmp_path / 'disavow.txt'
    stdout = run('link_risk.py', os.path.join(EXAMPLES, 'ahrefs_backlinks.csv'),
                 '--out', str(out), '--disavow', str(disavow))
    assert 'lost виключено 1' in stdout
    text = disavow.read_text(encoding='utf-8')
    assert 'domain:casino-bonus.example' in text
    assert 'domain:city-news.example' not in text
    ws = openpyxl.load_workbook(out)['Домени']
    header = [c.value for c in ws[1]]
    assert 'Ризик-скор' in header and 'Оцінка LLM' in header


def test_anchor_plan_avoids_overused_anchor(tmp_path):
    out = tmp_path / 'anchors.xlsx'
    run('anchor_plan.py', os.path.join(EXAMPLES, 'positions.csv'),
        '--anchors', os.path.join(EXAMPLES, 'existing_anchors.csv'),
        '--domain', 'example.com.ua', '--out', str(out))
    wb = openpyxl.load_workbook(out)
    avoid = [r[1] for r in wb['Уникати'].iter_rows(min_row=2, values_only=True)]
    assert avoid == ['купити каву в зернах']
    prios = {r[3] for r in wb['Кандидати'].iter_rows(min_row=2, values_only=True)}
    assert prios <= {'P1', 'P2', 'P3'}


def test_cluster_gaps_finds_brand_in_base(tmp_path):
    path = tmp_path / 'core.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Кластери'
    ws.append(['Keyword', 'Volume', 'Group', 'URL'])
    for kw, vol in [('кепка new era', 300), ('кепка new era чорна', 90), ('кепка жіноча', 500),
                    ('кепка чоловіча', 450), ('кепка new era 9forty', 70)]:
        ws.append([kw, vol, 'Кепки (база)', '/kepky/'])
    ws.append(['кепка nike', 800, 'Кепки > Nike', '/kepky/nike/'])
    wb.save(path)
    out = run('cluster_gaps.py', str(path))
    assert 'new' in out and 'era' in out
    assert 'nike' not in out.split('ЛАТИНСЬКІ')[1].split('КИРИЛИЧНІ')[0]


@pytest.mark.parametrize('script', ['serpstat_api.py', 'dataforseo_api.py', 'gsc_analyze.py',
                                    'link_risk.py', 'anchor_plan.py', 'coverage_check.py',
                                    'variant_dedupe.py', 'cluster_gaps.py'])
def test_help_works_without_keys(script):
    assert 'usage' in run(script, '--help')
