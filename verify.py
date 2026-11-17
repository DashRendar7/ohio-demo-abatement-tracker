"""Independent check of build.py's output before anything is saved.

Re-reads the same EPA file with a different spreadsheet reader (python-calamine instead of openpyxl)
and separately written date rules, then checks that:
- data/sites.json holds exactly the sites that are upcoming or underway today, with the same statuses;
- every county placement follows from the Census match or a single-county ZIP, and Locals follow from it;
- every Excel report contains exactly the sites it should, sheet by sheet;
- the counts in data/status.json match.
Exits with an error (so nothing is saved and the owner is emailed) on any difference."""
import datetime, json, os, re, sys
from collections import Counter
from python_calamine import CalamineWorkbook
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'data')
problems = []


def bad(msg):
    problems.append(msg)
    print('MISMATCH:', msg, flush=True)


status = json.load(open(os.path.join(DATA, 'status.json')))
sites = json.load(open(os.path.join(DATA, 'sites.json')))['sites']
today = datetime.date.fromisoformat(status['today'])
LOCALS = status['local_counties']
NO_END_DAYS = 60

# 1. Sites, read independently
book = CalamineWorkbook.from_path(os.environ.get('EPA_COPY', '/tmp/epa.xlsx'))
grid = book.get_sheet_by_index(0).to_python(skip_empty_area=False)
h = next(i for i, r in enumerate(grid) if 'Notification ID' in [str(c).strip() for c in r])
head = [str(c).strip() for c in grid[h]]
col = lambda name: head.index(name)
I = {k: col(n) for k, n in [('id', 'Notification ID'), ('addr', 'Building Address'), ('as', 'ABATEMENT DATE'), ('ae', 'COMPLETION DATE'),
                            ('ah', 'OFFSITE ON HOLD DATE'), ('ar', 'ONSITE OFF HOLD DATE'), ('ds', 'Demo Start Date'), ('de', 'Demo Completion Date')]}


def day(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
    return None


def state(s, e):
    if s is None and e is None:
        return ''
    if e is not None:
        if e < today:
            return ''
        return 'Upcoming' if (s is not None and s > today) else 'Underway'
    if s > today:
        return 'Upcoming'
    return 'Underway' if (today - s).days <= NO_END_DAYS else ''


expect = Counter()
seen = set()
for r in grid[h + 1:]:
    if all(c in (None, '') for c in r):
        continue
    t = tuple(repr(c) for c in r)
    if t in seen:
        continue
    seen.add(t)
    a = state(day(r[I['as']]), day(r[I['ae']]))
    d = state(day(r[I['ds']]), day(r[I['de']]))
    if a or d:
        idv = r[I['id']]
        idv = str(int(idv)) if isinstance(idv, float) and idv.is_integer() else str(idv).strip()
        expect[(idv, re.sub(r'\s+', ' ', str(r[I['addr']] or '')).strip(), a, d)] += 1
got = Counter((s['id'], s['address'], s['abate_status'], s['demo_status']) for s in sites)
for k in set(expect) | set(got):
    if expect[k] != got[k]:
        bad('site %s: expected %d, built %d' % (k, expect[k], got[k]))
print('sites expected', sum(expect.values()), 'built', len(sites))

# 2. County placements and Locals
zz = json.load(open(os.environ.get('ZIP_COPY', '/tmp/zipcounties.json')))
zc, names = zz['zc'], zz['names']
cache = json.load(open(os.path.join(DATA, 'geocache.json')))
for s in sites:
    a = s['address']
    m = re.search(r'(?<!\d)(\d{5})(?:-\d{4})?\s*$', a)
    z = m.group(1) if m else None
    c = cache.get(a, {})
    if s['confirmed']:
        if s['county_basis'].startswith('Census'):
            alt = c.get('alt') or {}
            primary = c.get('match') == 'Match' and names.get(c.get('county')) == s['county']
            second = (c.get('match') != 'Match' and names.get(alt.get('county')) == s['county'] and (z not in zc or alt.get('county') in zc[z]))
            if not (primary or second):
                bad('census placement not backed by the cache: %s -> %s' % (a, s['county']))
        elif s['county_basis'].startswith('ZIP'):
            alt_ok = str((c.get('alt') or {}).get('county', '')).startswith('39') and (z not in zc or c['alt']['county'] in zc[z])
            if not (z in zc and len(zc[z]) == 1 and names[next(iter(zc[z]))] == s['county']) or c.get('match') == 'Match' or alt_ok:
                bad('ZIP placement not valid: %s -> %s' % (a, s['county']))
        else:
            bad('unknown basis %s' % s['county_basis'])
        if s['candidates'] != [s['county']]:
            bad('confirmed site with several candidates: %s' % a)
    else:
        if s['county']:
            bad('unconfirmed site has a county: %s' % a)
        if c.get('match') == 'Match' and str(c.get('county', '')).startswith('39'):
            bad('site left unconfirmed although the Census matched it: %s' % a)
        alt = c.get('alt') or {}
        if str(alt.get('county', '')).startswith('39') and (z not in zc or alt['county'] in zc[z]):
            bad('site left unconfirmed although the second Census pass matched it: %s' % a)
        if z in zc and sorted(names[x] for x in zc[z]) != sorted(s['candidates']):
            bad('candidates differ from the ZIP counties: %s' % a)
    want = sorted((k for k, cs in LOCALS.items() if set(s['candidates']) & set(cs)), key=lambda k: (not k.isdigit(), int(k) if k.isdigit() else 0))
    if want != s['locals']:
        bad('locals for %s: expected %s, got %s' % (a, want, s['locals']))

# 3. Excel reports, sheet by sheet
SHEETS = {'Abatement upcoming': ('abate_status', 'Upcoming'), 'Abatement underway': ('abate_status', 'Underway'),
          'Demolition upcoming': ('demo_status', 'Upcoming'), 'Demolition underway': ('demo_status', 'Underway')}
files = {('Local_%s.xlsx' % k): [s for s in sites if k in s['locals']] for k in LOCALS}
files['All_Ohio.xlsx'] = sites
files['Needs_review.xlsx'] = [s for s in sites if s['review']]
for f, ss in files.items():
    wb = openpyxl.load_workbook(os.path.join(DATA, 'reports', f), read_only=True)
    for sh, (k, v) in SHEETS.items():
        rows = list(wb[sh].iter_rows(values_only=True))
        hd = list(rows[0])
        ia, ii = hd.index('Address'), hd.index('Notification ID')
        have = Counter((str(r[ii]), str(r[ia] or '')) for r in rows[1:])
        need = Counter((s['id'], s['address']) for s in ss if s[k] == v)
        if have != need:
            bad('%s / %s: %d rows, expected %d' % (f, sh, sum(have.values()), sum(need.values())))
    wb.close()

# 4. Status counts
for k in LOCALS:
    ss = [s for s in sites if k in s['locals']]
    if status['locals'][k]['sites'] != len(ss):
        bad('status count for %s' % k)
if status['sites'] != len(sites):
    bad('status site count')

if problems:
    print('%d problems; nothing will be saved.' % len(problems))
    sys.exit(1)
print('verify: all checks passed (%d sites, %d reports)' % (len(sites), len(files)))
