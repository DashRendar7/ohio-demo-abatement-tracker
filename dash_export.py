"""Writes the files the OVSSR dashboard reads from the OhioDemoAbatementTracker Hugging Face Space.

The dashboard reads Hugging Face files through a connector that returns at most 80,000 bytes per
read, so job rows are split into parts well under that size. dash/summary.json holds the run's
health, one row per Local, and the list of parts with their row counts, so the dashboard can
confirm it received every row. Run after build.py and verify.py; reads data/, writes dash/.

Job rows: every abatement or demolition phase marked Upcoming, and every abatement under way with
regulated asbestos (RACM) of 160 square feet or more, the federal notification threshold.
County and Local assignments are copied from the tracker, never re-derived here.
"""
import json, os, datetime, hashlib

HERE = os.path.dirname(os.path.abspath(__file__))
DATA, OUT = os.path.join(HERE, 'data'), os.path.join(HERE, 'dash')
LIMIT = 60000


def dumps(x):
    return json.dumps(x, ensure_ascii=False, separators=(',', ':'))


def parts(columns, rows, prefix):
    out, cur = [], []
    for r in rows:
        cur.append([r.get(c) for c in columns])
        if len(dumps({'columns': columns, 'rows': cur}).encode('utf-8')) > LIMIT:
            cur.pop()
            out.append(cur)
            cur = [[r.get(c) for c in columns]]
    if cur:
        out.append(cur)
    files = []
    for i, chunk in enumerate(out, 1):
        name = '%s_%02d.json' % (prefix, i)
        body = dumps({'part': i, 'of': len(out), 'columns': columns, 'rows': chunk})
        assert len(body.encode('utf-8')) < 78000, name
        with open(os.path.join(OUT, name), 'w', encoding='utf-8') as f:
            f.write(body)
        files.append({'file': name, 'rows': len(chunk)})
    return files


def title(s):
    return ' '.join(w.capitalize() for w in str(s or '').split())


def num(v):
    try:
        return float(str(v).replace(',', ''))
    except (TypeError, ValueError):
        return None


def county_label(s):
    if s['county']:
        return title(s['county'])
    if s.get('likely'):
        return 'Likely ' + title(s['likely'])
    c = s.get('candidates') or []
    return 'Not confirmed: ' + ' or '.join(title(x) for x in c) if c else 'Not confirmed'


def main():
    os.makedirs(OUT, exist_ok=True)
    for n in os.listdir(OUT):
        if n.endswith('.json'):
            os.remove(os.path.join(OUT, n))
    st = json.load(open(os.path.join(DATA, 'status.json'), encoding='utf-8'))
    sites = json.load(open(os.path.join(DATA, 'sites.json'), encoding='utf-8'))['sites']
    assert len(sites) == st['sites'], 'sites.json and status.json disagree'
    jobs, seen = [], {}
    for s in sites:
        base = s['id'] + '-' + hashlib.sha1((s['address'] + '|' + s['building']).encode('utf-8')).hexdigest()[:6]
        seen[base] = seen.get(base, 0) + 1
        base += '' if seen[base] == 1 else '-%d' % seen[base]
        for p, name in (('abate', 'Abatement'), ('demo', 'Demolition')):
            status = s[p + '_status']
            racm = num(s['racm_surface'])
            if status == 'Upcoming' or (p == 'abate' and status == 'Underway' and racm is not None and racm >= 160):
                jobs.append({
                    'key': base + '-' + p, 'notice': s['id'], 'phase': name, 'status': status,
                    'start': s[p + '_start'] or '', 'end': s[p + '_end'] or '',
                    'locals': ','.join(s['locals']), 'might_be': ','.join(s['maybe_locals']),
                    'county': county_label(s), 'county_status': 'Confirmed' if s['confirmed'] else ('Likely' if s.get('likely') else 'Not confirmed'),
                    'owner': s['owner'], 'building': s['building'] or '', 'address': s['address'] or '',
                    'contractor': s[p + '_contractor'] or '', 'racm_sq_ft': racm, 'racm_pipe_ft': num(s['racm_pipe']),
                    'activity': s['activity'],
                })
    jobs.sort(key=lambda j: (j['start'] or '9999', j['key']))
    keys = [j['key'] for j in jobs]
    assert len(set(keys)) == len(keys), 'duplicate job keys'
    columns = sorted({k for j in jobs for k in j})
    files = parts(columns, jobs, 'jobs')
    locals_rows = []
    for k in sorted(st['locals'], key=lambda x: int(x)):
        v = st['locals'][k]
        locals_rows.append({'local': k, 'label': st['local_labels'][k], 'counties': ', '.join(title(c) for c in st['local_counties'][k]),
                            'sites': v['sites'], 'abate_upcoming': v['abate_upcoming'], 'abate_underway': v['abate_underway'],
                            'demo_upcoming': v['demo_upcoming'], 'demo_underway': v['demo_underway'],
                            'not_confirmed': v['not_confirmed'], 'likely': v['likely'], 'maybe': v['maybe']})
    now = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)
    run = os.environ.get('GITHUB_RUN_ID')
    summary = {
        'kind': 'ohio-demo-abatement-tracker',
        'exported': now.isoformat(), 'built': st['built'], 'epa_as_of': st['epa_as_of'], 'epa_posted': st.get('epa_posted', ''),
        'run_url': ('%s/%s/actions/runs/%s' % (os.environ.get('GITHUB_SERVER_URL', 'https://github.com'), os.environ.get('GITHUB_REPOSITORY', ''), run)) if run else '',
        'sites': st['sites'], 'need_county_check': st['review'], 'confirmed': st['confirmed'], 'likely': st.get('likely', 0),
        'not_confirmed': st['not_confirmed'], 'basis': st['basis'],
        'job_rows': len(jobs), 'parts': files, 'locals': locals_rows,
    }
    body = dumps(summary)
    assert len(body.encode('utf-8')) < 78000, 'summary.json too large'
    with open(os.path.join(OUT, 'summary.json'), 'w', encoding='utf-8') as f:
        f.write(body)
    print('dash export:', len(jobs), 'job rows in', len(files), 'parts;', len(locals_rows), 'Locals')


if __name__ == '__main__':
    main()
