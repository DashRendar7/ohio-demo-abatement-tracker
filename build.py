"""Builds the Ohio demolition and abatement tracker from the Ohio EPA notification list.

Every run (GitHub Actions, .github/workflows/tracker.yml):
1. Downloads the Ohio EPA asbestos/demolition notification list (one row per site; a notification
   can cover several sites).
2. Keeps sites with an abatement or demolition phase that is upcoming or underway today (Ohio time).
3. Places each site in a county:
   - Census Bureau address match (exact county for the matched address), or
   - when the address can't be matched, the ZIP code, but only if that ZIP lies in a single county.
   Anything else is "county not confirmed": the site is listed for every county it could be in,
   marked as such, and on the statewide review list. No site is dropped.
4. Writes data/sites.json, data/status.json and one Excel report per Local in data/reports/.
verify.py then re-reads the EPA file independently and checks every number before anything is saved.
"""
import datetime, hashlib, io, json, os, re, sys, time, urllib.request
from collections import defaultdict
from zoneinfo import ZoneInfo

import openpyxl
import xlsxwriter

UA = 'ohio-demo-abatement-tracker/1.0 (public Ohio EPA notification data; github.com/DashRendar7/ohio-demo-abatement-tracker)'
EPA = 'https://dam.assets.ohio.gov/raw/upload/epa.ohio.gov/Public/DAPC/NotificationList.xlsx'
ZCTA = 'https://www2.census.gov/geo/docs/maps-data/data/rel2020/zcta520/tab20_zcta520_county20_natl.txt'
CENSUS = 'https://geocoding.geo.census.gov/geocoder/geographies/addressbatch'
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'data')
OHIO = ZoneInfo('America/New_York')
RECENT_NO_END_DAYS = 60   # a phase that started without a completion date counts as underway this long
RETRY_UNMATCHED_DAYS = 14  # addresses the Census could not match are tried again after this many days

LOCALS = {
    '83': ['ADAMS', 'ATHENS', 'GALLIA', 'HIGHLAND', 'JACKSON', 'LAWRENCE', 'MEIGS', 'PIKE', 'ROSS', 'SCIOTO', 'VINTON'],
    '125': ['MAHONING'],
    '134': ['COSHOCTON', 'HOLMES', 'TUSCARAWAS'],
    '245': ['ASHTABULA'],
    '265': ['BROWN', 'CLERMONT', 'CLINTON', 'HAMILTON'],
    '310': ['CUYAHOGA', 'GEAUGA', 'LAKE'],
    '329': ['ALLEN', 'AUGLAIZE', 'MERCER', 'PAULDING', 'PUTNAM', 'SHELBY', 'VAN WERT'],
    '423': ['FAIRFIELD', 'FAYETTE', 'FRANKLIN', 'HOCKING', 'LICKING', 'MADISON', 'PICKAWAY', 'UNION'],
    '480': ['ERIE', 'HURON', 'OTTAWA', 'SANDUSKY'],
    '500': ['DEFIANCE', 'FULTON', 'HENRY', 'LUCAS', 'WILLIAMS', 'WOOD'],
    '530': ['GUERNSEY', 'MUSKINGUM', 'NOBLE', 'PERRY'],
    '534': ['BUTLER', 'WARREN'],
    '574': ['DELAWARE', 'HANCOCK', 'HARDIN', 'MARION', 'SENECA', 'WYANDOT'],
    '639': ['MONROE', 'MORGAN', 'WASHINGTON'],
    '758': ['LORAIN'],
    '809': ['COLUMBIANA', 'HARRISON', 'JEFFERSON'],
    '860': ['CUYAHOGA', 'GEAUGA', 'LAKE'],
    '894': ['MEDINA', 'PORTAGE', 'SUMMIT'],
    '935': ['TRUMBULL'],
    '1015': ['CARROLL', 'STARK', 'WAYNE'],
    '1216': ['ASHLAND', 'CRAWFORD', 'KNOX', 'MORROW', 'RICHLAND'],
    '1410': ['CHAMPAIGN', 'CLARK', 'DARKE', 'GREENE', 'LOGAN', 'MIAMI', 'MONTGOMERY', 'PREBLE'],
    'BELMONT': ['BELMONT'],  # Belmont County is covered by a West Virginia local
}
LOCAL_LABEL = {k: ('Local ' + k) for k in LOCALS}
LOCAL_LABEL['BELMONT'] = 'Belmont County (WV local)'

# The EPA sheet has a group row above the header row; some headers repeat under different groups.
COLS = [  # (key, group contains, header contains)
    ('id', None, 'notification id'), ('owner', None, 'owner name'), ('owner_phone', None, 'owner phone'),
    ('building', None, 'building name'), ('address', None, 'building address'), ('activity', None, 'on site activity'),
    ('racm_pipe', 'regulated', 'racm pipe'), ('racm_surface', 'regulated', 'racm surface'), ('racm_volume', 'regulated', 'racm volume'),
    ('nf_rm_c1_pipe', 'to be removed', 'nf cat-1 pipe'), ('nf_rm_c1_surface', 'to be removed', 'nf cat-1 surface'),
    ('nf_rm_c2_pipe', 'to be removed', 'nf cat-2 pipe'), ('nf_rm_c2_surface', 'to be removed', 'nf cat-2 surface'),
    ('nf_stay_c1_pipe', 'not to be removed', 'nf cat-1 pipe'), ('nf_stay_c1_surface', 'not to be removed', 'nf cat-1 surface'),
    ('nf_stay_c2_pipe', 'not to be removed', 'nf cat-2 pipe'), ('nf_stay_c2_surface', 'not to be removed', 'nf cat-2 surface'),
    ('abate_contractor_id', None, 'abatement contractor id'), ('abate_contractor', None, 'abatement contractor name'),
    ('abate_phone', None, 'abatement phone'), ('specialist', None, 'on site specialist name'),
    ('setup', None, 'setup date'), ('abate_start', None, 'abatement date'), ('abate_end', None, 'completion date'),
    ('abate_hold', None, 'offsite on hold date'), ('abate_resume', None, 'onsite off hold date'),
    ('demo_contractor', None, 'demo contractor name'), ('demo_phone', None, 'demo contractor phone'),
    ('demo_start', None, 'demo start date'), ('demo_end', None, 'demo completion date'),
    ('demo_hold', None, 'demo on hold date'), ('demo_resume', None, 'demo off hold date'),
]
DATES = ['setup', 'abate_start', 'abate_end', 'abate_hold', 'abate_resume', 'demo_start', 'demo_end', 'demo_hold', 'demo_resume']


def log(*a):
    print('[tracker]', *a, flush=True)


def get(url, data=None, headers=None, timeout=300, tries=4):
    err = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, data=data, headers=dict({'User-Agent': UA}, **(headers or {})))
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read(), dict(r.headers)
        except Exception as e:
            err = e
            time.sleep(5 * (i + 1))
    raise RuntimeError('%s failed: %s' % (url, err))


def norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip().lower()


def read_epa(raw):
    """Returns (rows as dicts, 'data current as of' text). Raises if the layout is not recognised."""
    wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    hi = next(i for i, r in enumerate(rows[:30]) if r and any(norm(c) == 'notification id' for c in r) and any('building address' in norm(c) for c in r))
    head, groups = [norm(c) for c in rows[hi]], []
    g = ''
    for i in range(len(head)):
        v = rows[hi - 1][i] if hi > 0 and i < len(rows[hi - 1]) else None
        g = norm(v) if v not in (None, '') else g
        groups.append(g)
    idx = {}
    for key, grp, hdr in COLS:
        hits = [i for i, h in enumerate(head) if (h == hdr or h.startswith(hdr)) and (grp is None or (grp in groups[i] and not (grp == 'to be removed' and 'not to be removed' in groups[i])))]
        if len(hits) != 1:
            raise SystemExit('EPA file layout changed: column for %s found %d times. Stopping without changes.' % (key, len(hits)))
        idx[key] = hits[0]
    asof = ''
    for r in rows[:hi]:
        for c in r or []:
            m = re.search(r'current as of\s*:?\s*(.+)$', str(c or ''), re.I)
            if m:
                asof = m.group(1).strip()
    out, seen = [], set()
    for r in rows[hi + 1:]:
        if not r or all(c in (None, '') for c in r):
            continue
        key = tuple(repr(c) for c in r)
        if key in seen:  # exact duplicate row
            continue
        seen.add(key)
        rec = {}
        for k, i in idx.items():
            v = r[i] if i < len(r) else None
            if k in DATES:
                rec[k] = as_date(v)
            else:
                rec[k] = re.sub(r'\s+', ' ', str(v)).strip() if v not in (None, '') else ''
        out.append(rec)
    return out, asof


def as_date(v):
    """A cell as a date: a date/datetime cell, or a spreadsheet date number if the cell lost its date format."""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
    return None


def phase(start, end, hold, resume, today):
    """'Upcoming', 'Underway' or None (finished or no dates), plus a note."""
    if start is None and end is None:
        return None, ''
    if end is not None and end < today:
        return None, ''
    note = ''
    if end is None:
        if start > today:
            st, note = 'Upcoming', 'no completion date in the notice'
        elif (today - start).days <= RECENT_NO_END_DAYS:
            st, note = 'Underway', 'no completion date in the notice'
        else:
            return None, ''
    elif start is None:
        st, note = 'Underway', 'no start date in the notice'
    else:
        st = 'Upcoming' if start > today else 'Underway'
    if hold and hold <= today and not (resume and resume >= hold and resume <= today):
        note = ('on hold since %s' % hold.isoformat()) + ('; ' + note if note else '')
    return st, note


def zip_of(addr):
    """The ZIP at the end of the address ("... OH 43215" or "... 43215-1234"); None if there isn't one."""
    m = re.search(r'(?:\bOH(?:IO)?\b\.?,?\s*)?(?<!\d)(\d{5})(?:-\d{4})?\s*$', addr or '', re.I)
    return m.group(1) if m else None


def load_zip_counties():
    raw, _ = get(ZCTA)
    zc, names = defaultdict(dict), {}
    for line in raw.decode('utf-8-sig').split('\n')[1:]:
        f = line.split('|')
        if len(f) > 17 and f[9].startswith('39') and f[1]:
            zc[f[1]][f[9]] = zc[f[1]].get(f[9], 0) + int(f[16] or 0)
            names[f[9]] = f[10].replace(' County', '').upper()
    if len(names) != 88:
        raise SystemExit('Census ZIP-county file has %d Ohio counties, expected 88. Stopping without changes.' % len(names))
    return dict(zc), names


def geocode(addresses, cache, today):
    """Census batch geocoding for addresses not yet in the cache (and unmatched ones after a while)."""
    todo = []
    for a in addresses:
        c = cache.get(a)
        if c is None or (c['match'] != 'Match' and (today - datetime.date.fromisoformat(c['when'])).days >= RETRY_UNMATCHED_DAYS):
            todo.append(a)
    log('geocoding', len(todo), 'addresses')
    import csv
    for s in range(0, len(todo), 1000):
        chunk = todo[s:s + 1000]
        body_csv = '\n'.join('%d,"%s",,,' % (j, a.replace('"', '')) for j, a in enumerate(chunk))
        bd = '----trk' + hashlib.md5(body_csv.encode()).hexdigest()[:12]
        body = ('--%s\r\nContent-Disposition: form-data; name="benchmark"\r\n\r\nPublic_AR_Current\r\n'
                '--%s\r\nContent-Disposition: form-data; name="vintage"\r\n\r\nCurrent_Current\r\n'
                '--%s\r\nContent-Disposition: form-data; name="addressFile"; filename="a.csv"\r\nContent-Type: text/csv\r\n\r\n%s\r\n--%s--\r\n'
                % (bd, bd, bd, body_csv, bd)).encode()
        raw, _ = get(CENSUS, data=body, headers={'Content-Type': 'multipart/form-data; boundary=' + bd})
        got = set()
        for f in csv.reader(io.StringIO(raw.decode('utf-8', 'replace'))):
            if len(f) >= 3 and f[0].isdigit() and int(f[0]) < len(chunk):
                a = chunk[int(f[0])]
                got.add(a)
                cache[a] = {'match': f[2], 'type': f[3] if len(f) > 3 else '', 'matched': f[4] if len(f) > 4 else '',
                            'county': (f[8] + f[9]) if len(f) > 9 and f[8] and f[9] else '', 'when': today.isoformat()}
        missing = [a for a in chunk if a not in got]
        if missing:
            raise RuntimeError('Census geocoder returned no answer for %d addresses; stopping without changes' % len(missing))


def place(addr, cache, zc, names):
    """County for one site: (county name or '', basis, candidates, confirmed?)."""
    z = zip_of(addr)
    zcs = zc.get(z) if z else None
    c = cache.get(addr) or {}
    if c.get('match') == 'Match' and c.get('county', '').startswith('39'):
        cty = names[c['county']]
        if not zcs or c['county'] in zcs:
            return cty, 'Census address match', [cty], True
        zn = ', '.join(sorted(names[x] for x in zcs))
        return cty, 'Census address match (%s); the ZIP in the notice, %s, is in %s' % (c['matched'], z, zn.title()), [cty], True
    if zcs and len(zcs) == 1:
        cty = names[next(iter(zcs))]
        return cty, 'ZIP %s lies entirely in this county' % z, [cty], True
    if zcs:
        cands = [names[x] for x in sorted(zcs, key=lambda x: -zcs[x])]
        return '', 'County not confirmed: ZIP %s spans %s, and the address could not be matched' % (z, ', '.join(x.title() for x in cands)), cands, False
    return '', 'County not confirmed: no usable ZIP, and the address could not be matched', [], False


def alerts(s):
    a = []
    o = s['owner'].upper()
    if any(w in o for w in ('LAND BANK', 'LANDBANK', 'REUTILIZATION', 'LAND REUTIL')):
        a.append('Land bank: one notice can cover many houses')
    if any(w in s['building'].upper() for w in ('ANNUAL', 'MAINTENANCE', 'O&M', 'O & M')):
        a.append('Maintenance permit')
    for k1, k2 in (('abate_start', 'abate_end'), ('demo_start', 'demo_end')):
        if s[k1] and s[k2] and (s[k2] - s[k1]).days > 365:
            a.append('Runs more than a year')
            break
    return a


def build(today=None, epa_raw=None, min_rows=20000):
    today = today or datetime.datetime.now(OHIO).date()
    os.makedirs(os.path.join(DATA, 'reports'), exist_ok=True)
    if epa_raw is None:
        epa_raw, hdr = get(EPA)
    else:
        hdr = {}
    epa_sha = hashlib.sha256(epa_raw).hexdigest()
    open(os.environ.get('EPA_COPY', '/tmp/epa.xlsx'), 'wb').write(epa_raw)  # for verify.py; not saved in the project
    rows, asof = read_epa(epa_raw)
    log('EPA rows (after removing exact duplicates):', len(rows), 'as of', asof)
    if len(rows) < min_rows:  # the list holds about 73,000 rows; far fewer means a bad download
        raise SystemExit('EPA file has only %d rows; stopping without changes.' % len(rows))
    zc, names = load_zip_counties()
    json.dump({'zc': zc, 'names': names}, open(os.environ.get('ZIP_COPY', '/tmp/zipcounties.json'), 'w'))
    cache_path = os.path.join(DATA, 'geocache.json')
    cache = json.load(open(cache_path)) if os.path.exists(cache_path) else {}
    sites, no_details = [], 0
    for r in rows:
        ab, abn = phase(r['abate_start'], r['abate_end'], r['abate_hold'], r['abate_resume'], today)
        de, den = phase(r['demo_start'], r['demo_end'], r['demo_hold'], r['demo_resume'], today)
        if not ab and not de:
            continue
        if not r['address']:
            no_details += 1
        s = dict(r, abate_status=ab or '', abate_note=abn, demo_status=de or '', demo_note=den)
        sites.append(s)
    geocode(sorted({s['address'] for s in sites if s['address']}), cache, today)
    json.dump(cache, open(cache_path, 'w'), indent=0, sort_keys=True)
    for s in sites:
        if s['address']:
            s['county'], s['county_basis'], s['candidates'], s['confirmed'] = place(s['address'], cache, zc, names)
        else:
            s['county'], s['county_basis'], s['candidates'], s['confirmed'] = '', 'County not confirmed: the notice has no address ("Contact Ohio EPA for Details")', [], False
        s['alerts'] = alerts(s)
        s['locals'] = sorted((k for k, cs in LOCALS.items() if set(s['candidates']) & set(cs)), key=lambda k: (not k.isdigit(), int(k) if k.isdigit() else 0))
        s['review'] = (not s['confirmed']) or ('ZIP in the notice' in s['county_basis'])
    sites.sort(key=lambda s: (s['county'] or 'ZZZ', min([d for d in (s['abate_start'], s['demo_start']) if d] or [datetime.date.max]), s['id'], s['address']))
    # Outputs
    def j(s):
        return {k: (v.isoformat() if isinstance(v, datetime.date) else v) for k, v in s.items()}
    status = {'built': datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat(), 'today': today.isoformat(),
              'epa_as_of': asof, 'epa_posted': hdr.get('Last-Modified', ''), 'epa_sha256': epa_sha, 'epa_rows': len(rows),
              'sites': len(sites), 'notifications': len({s['id'] for s in sites}),
              'confirmed': sum(1 for s in sites if s['confirmed']), 'not_confirmed': sum(1 for s in sites if not s['confirmed']),
              'review': sum(1 for s in sites if s['review']), 'no_details': no_details,
              'basis': dict(sorted(_count(s['county_basis'].split(' (')[0].split(';')[0].split(':')[0] for s in sites).items())),
              'locals': {k: _local_counts([s for s in sites if k in s['locals']]) for k in LOCALS},
              'local_labels': LOCAL_LABEL, 'local_counties': LOCALS}
    json.dump({'sites': [j(s) for s in sites]}, open(os.path.join(DATA, 'sites.json'), 'w'), indent=0, sort_keys=True)
    for k in LOCALS:
        write_xlsx(os.path.join(DATA, 'reports', 'Local_%s.xlsx' % k), LOCAL_LABEL[k], [s for s in sites if k in s['locals']], status, LOCALS[k])
    write_xlsx(os.path.join(DATA, 'reports', 'All_Ohio.xlsx'), 'All Ohio', sites, status, None)
    write_xlsx(os.path.join(DATA, 'reports', 'Needs_review.xlsx'), 'Needs a county check', [s for s in sites if s['review']], status, None)
    json.dump(status, open(os.path.join(DATA, 'status.json'), 'w'), indent=1, sort_keys=True)
    log('sites', len(sites), 'confirmed', status['confirmed'], 'not confirmed', status['not_confirmed'])
    return sites, status


def _count(it):
    d = defaultdict(int)
    for x in it:
        d[x] += 1
    return d


def _local_counts(ss):
    return {'sites': len(ss), 'abate_upcoming': sum(1 for s in ss if s['abate_status'] == 'Upcoming'),
            'abate_underway': sum(1 for s in ss if s['abate_status'] == 'Underway'),
            'demo_upcoming': sum(1 for s in ss if s['demo_status'] == 'Upcoming'),
            'demo_underway': sum(1 for s in ss if s['demo_status'] == 'Underway'),
            'not_confirmed': sum(1 for s in ss if not s['confirmed'])}


SHEET_COLS = [  # (header, key or function, width)
    ('Notification ID', 'id', 12), ('County', lambda s: s['county'].title() if s['county'] else 'Not confirmed', 13),
    ('County check', lambda s: '' if (s['confirmed'] and not s['review']) else s['county_basis'], 40),
    ('Organizer alerts', lambda s: '; '.join(s['alerts']), 30),
    ('Owner', 'owner', 30), ('Owner phone', 'owner_phone', 14), ('Building', 'building', 30), ('Address', 'address', 40),
    ('Activity', 'activity', 22),
    ('Abatement', 'abate_status', 11), ('Abatement start', 'abate_start', 12), ('Abatement completion', 'abate_end', 12),
    ('Abatement note', 'abate_note', 24), ('Setup date', 'setup', 12),
    ('Abatement contractor', 'abate_contractor', 30), ('Abatement contractor phone', 'abate_phone', 14), ('On-site specialist', 'specialist', 20),
    ('Demolition', 'demo_status', 11), ('Demolition start', 'demo_start', 12), ('Demolition completion', 'demo_end', 12),
    ('Demolition note', 'demo_note', 24), ('Demolition contractor', 'demo_contractor', 30), ('Demolition contractor phone', 'demo_phone', 14),
    ('RACM pipe (ft)', 'racm_pipe', 9), ('RACM surface (sq ft)', 'racm_surface', 9), ('RACM volume (cu ft)', 'racm_volume', 9),
    ('Non-friable to remove: Cat 1 pipe (ft)', 'nf_rm_c1_pipe', 9), ('Non-friable to remove: Cat 1 surface (sq ft)', 'nf_rm_c1_surface', 9),
    ('Non-friable to remove: Cat 2 pipe (ft)', 'nf_rm_c2_pipe', 9), ('Non-friable to remove: Cat 2 surface (sq ft)', 'nf_rm_c2_surface', 9),
    ('Counted for Locals', lambda s: ', '.join(LOCAL_LABEL[k] for k in s['locals']), 18),
]
SHEETS = [('Abatement upcoming', 'abate_status', 'Upcoming', 'abate_start'), ('Abatement underway', 'abate_status', 'Underway', 'abate_end'),
          ('Demolition upcoming', 'demo_status', 'Upcoming', 'demo_start'), ('Demolition underway', 'demo_status', 'Underway', 'demo_end')]


def sheet_rows(sites, key, val, sortkey):
    ss = [s for s in sites if s[key] == val]
    return sorted(ss, key=lambda s: (s[sortkey] or datetime.date.max, s['id'], s['address']))


def write_xlsx(path, title, sites, status, counties):
    wb = xlsxwriter.Workbook(path, {'constant_memory': False})
    wb.set_properties({'title': title, 'created': datetime.datetime(2026, 1, 1)})
    bold = wb.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'bg_color': '#E8EEF6', 'border': 1})
    datef = wb.add_format({'num_format': 'yyyy-mm-dd'})
    warn = wb.add_format({'bg_color': '#FFF2CC'})
    ws = wb.add_worksheet('About')
    lines = [title, '',
             'Source: Ohio EPA asbestos and demolition notification list (epa.ohio.gov), data current as of %s.' % status['epa_as_of'],
             'Projects upcoming or underway as of %s (Ohio time).' % status['today'],
             'Counties covered: %s.' % (', '.join(c.title() for c in counties) if counties else 'all of Ohio'),
             '',
             'Each row is one site. One notification can cover several sites (land banks often list many houses).',
             'Upcoming: the phase starts after today. Underway: it started on or before today and its completion date is today or later.',
             'A phase with a start date but no completion date counts as underway for %d days after it starts.' % RECENT_NO_END_DAYS,
             'County: Census Bureau address match, or the ZIP code when the address could not be matched and the ZIP lies in one county.',
             'Rows highlighted yellow could not be placed in a county for certain. They are listed for every county they could be in,',
             'with the reason in "County check". Check the address before acting on them.']
    for i, t in enumerate(lines):
        ws.write(i, 0, t, bold if i == 0 else None)
    ws.set_column(0, 0, 120)
    for name, key, val, sortkey in SHEETS:
        rows = sheet_rows(sites, key, val, sortkey)
        ws = wb.add_worksheet(name)
        for c, (h, _, w) in enumerate(SHEET_COLS):
            ws.write(0, c, h, bold)
            ws.set_column(c, c, w)
        for r, s in enumerate(rows, start=1):
            fmt_row = warn if (not s['confirmed'] or s['review']) else None
            for c, (h, k, _) in enumerate(SHEET_COLS):
                v = k(s) if callable(k) else s[k]
                if isinstance(v, datetime.date):
                    ws.write_datetime(r, c, datetime.datetime(v.year, v.month, v.day), datef)
                elif v in (None, ''):
                    ws.write_blank(r, c, None, fmt_row)
                elif isinstance(v, str) and re.fullmatch(r'\d+(\.\d+)?', v) and k != 'id':
                    ws.write_number(r, c, float(v), fmt_row)
                else:
                    ws.write(r, c, v, fmt_row)
        ws.freeze_panes(1, 2)
        ws.autofilter(0, 0, max(1, len(rows)), len(SHEET_COLS) - 1)
    wb.close()


if __name__ == '__main__':
    build()
