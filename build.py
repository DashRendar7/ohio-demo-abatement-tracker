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
import csv, datetime, hashlib, io, json, os, re, sys, time, urllib.request
from collections import defaultdict
from zoneinfo import ZoneInfo

import openpyxl
import xlsxwriter

UA = 'ohio-demo-abatement-tracker/1.0 (public Ohio EPA notification data; github.com/DashRendar7/ohio-demo-abatement-tracker)'
EPA = 'https://dam.assets.ohio.gov/raw/upload/epa.ohio.gov/Public/DAPC/NotificationList.xlsx'
ZCTA = 'https://www2.census.gov/geo/docs/maps-data/data/rel2020/zcta520/tab20_zcta520_county20_natl.txt'
CENSUS = 'https://geocoding.geo.census.gov/geocoder/geographies/addressbatch'
BAF = 'https://www2.census.gov/geo/docs/maps-data/data/baf2020/BlockAssign_ST39_OH.zip'
GAZ = 'https://www2.census.gov/geo/docs/maps-data/data/gazetteer/2020_Gazetteer/2020_gaz_place_39.txt'
# ODOT's three-letter county codes, used in project names such as "HIG-73-21.04"
ODOT = {'ADA': 'ADAMS', 'ALL': 'ALLEN', 'ASD': 'ASHLAND', 'ATB': 'ASHTABULA', 'ATH': 'ATHENS', 'AUG': 'AUGLAIZE', 'BEL': 'BELMONT',
        'BRO': 'BROWN', 'BUT': 'BUTLER', 'CAR': 'CARROLL', 'CHP': 'CHAMPAIGN', 'CLA': 'CLARK', 'CLE': 'CLERMONT', 'CLI': 'CLINTON',
        'COL': 'COLUMBIANA', 'COS': 'COSHOCTON', 'CRA': 'CRAWFORD', 'CUY': 'CUYAHOGA', 'DAR': 'DARKE', 'DEF': 'DEFIANCE', 'DEL': 'DELAWARE',
        'ERI': 'ERIE', 'FAI': 'FAIRFIELD', 'FAY': 'FAYETTE', 'FRA': 'FRANKLIN', 'FUL': 'FULTON', 'GAL': 'GALLIA', 'GEA': 'GEAUGA',
        'GRE': 'GREENE', 'GUE': 'GUERNSEY', 'HAM': 'HAMILTON', 'HAN': 'HANCOCK', 'HAR': 'HARDIN', 'HAS': 'HARRISON', 'HEN': 'HENRY',
        'HIG': 'HIGHLAND', 'HOC': 'HOCKING', 'HOL': 'HOLMES', 'HUR': 'HURON', 'JAC': 'JACKSON', 'JEF': 'JEFFERSON', 'KNO': 'KNOX',
        'LAK': 'LAKE', 'LAW': 'LAWRENCE', 'LIC': 'LICKING', 'LOG': 'LOGAN', 'LOR': 'LORAIN', 'LUC': 'LUCAS', 'MAD': 'MADISON',
        'MAH': 'MAHONING', 'MAR': 'MARION', 'MED': 'MEDINA', 'MEG': 'MEIGS', 'MER': 'MERCER', 'MIA': 'MIAMI', 'MOE': 'MONROE',
        'MOT': 'MONTGOMERY', 'MRG': 'MORGAN', 'MRW': 'MORROW', 'MUS': 'MUSKINGUM', 'NOB': 'NOBLE', 'OTT': 'OTTAWA', 'PAU': 'PAULDING',
        'PER': 'PERRY', 'PIC': 'PICKAWAY', 'PIK': 'PIKE', 'POR': 'PORTAGE', 'PRE': 'PREBLE', 'PUT': 'PUTNAM', 'RIC': 'RICHLAND',
        'ROS': 'ROSS', 'SAN': 'SANDUSKY', 'SCI': 'SCIOTO', 'SEN': 'SENECA', 'SHE': 'SHELBY', 'STA': 'STARK', 'SUM': 'SUMMIT',
        'TRU': 'TRUMBULL', 'TUS': 'TUSCARAWAS', 'UNI': 'UNION', 'VAN': 'VAN WERT', 'VIN': 'VINTON', 'WAR': 'WARREN', 'WAS': 'WASHINGTON',
        'WAY': 'WAYNE', 'WIL': 'WILLIAMS', 'WOO': 'WOOD', 'WYA': 'WYANDOT'}
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
}
LOCAL_LABEL = {k: ('Local ' + k) for k in LOCALS}

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
    ('specialist_cert', None, 'specialist cert number'), ('cert_exp', None, 'cert expiration date'),
    ('setup', None, 'setup date'), ('abate_start', None, 'abatement date'), ('abate_end', None, 'completion date'),
    ('abate_hold', None, 'offsite on hold date'), ('abate_resume', None, 'onsite off hold date'),
    ('demo_contractor', None, 'demo contractor name'), ('demo_phone', None, 'demo contractor phone'),
    ('demo_start', None, 'demo start date'), ('demo_end', None, 'demo completion date'),
    ('demo_hold', None, 'demo on hold date'), ('demo_resume', None, 'demo off hold date'),
]
DATES = ['setup', 'abate_start', 'abate_end', 'abate_hold', 'abate_resume', 'demo_start', 'demo_end', 'demo_hold', 'demo_resume', 'cert_exp']
# Quantities: written to the spreadsheets as numbers when the EPA cell holds a plain number. Everything else stays text
# exactly as the EPA wrote it (so IDs, certificate numbers and phone numbers keep any leading zeros).
NUMERIC = ['racm_pipe', 'racm_surface', 'racm_volume', 'nf_rm_c1_pipe', 'nf_rm_c1_surface', 'nf_rm_c2_pipe', 'nf_rm_c2_surface',
           'nf_stay_c1_pipe', 'nf_stay_c1_surface', 'nf_stay_c2_pipe', 'nf_stay_c2_surface']


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


def load_places(names):
    """Ohio towns (incorporated places and CDPs) -> the counties they lie in, from the Census 2020 block
    assignment file (each block's ID carries its county) and the 2020 Gazetteer (place names)."""
    import zipfile
    raw, _ = get(BAF)
    z = zipfile.ZipFile(io.BytesIO(raw))
    f = next(n for n in z.namelist() if 'INCPLACE_CDP' in n.upper())
    by_fp = defaultdict(set)
    for line in z.read(f).decode('utf-8-sig').splitlines()[1:]:
        p = line.split('|')
        if len(p) >= 2 and p[1].strip():
            by_fp[p[1].strip()].add(p[0][:5])
    gaz, _ = get(GAZ)
    towns = defaultdict(set)
    for line in gaz.decode('utf-8-sig', 'replace').splitlines()[1:]:
        p = line.split('\t')
        if len(p) > 4 and p[1].startswith('39'):
            name = re.sub(r'\s+(city|village|CDP|town)$', '', p[3].strip(), flags=re.I).upper()
            towns[name] |= {names[c] for c in by_fp.get(p[1][2:], ()) if c in names}
    towns = {k: sorted(v) for k, v in towns.items() if v}
    if len(towns) < 1000:
        raise SystemExit('Census town list has only %d Ohio towns; stopping without changes.' % len(towns))
    return towns


def town_in(text, towns, before_oh=True):
    """The longest town name at the end of the street-and-town part of an address (or after "City of" /
    "Village of" in an owner name). Returns (town, counties) or (None, [])."""
    t = (text or '').upper()
    if before_oh:
        m = re.match(r'^(.*\S)\s+OH(?:IO)?\.?,?\s*(?:\d{4,5}\S*)?\s*$', t)
        if not m:
            return None, []
        words = re.sub(r'[^A-Z0-9 \'.-]', ' ', m.group(1)).split()
        for n in (4, 3, 2, 1):
            if len(words) > n:
                cand = ' '.join(words[-n:]).replace('ST.', 'ST')
                for k in (cand, cand.replace(' ST ', ' ST. ')):
                    if k in towns:
                        return k, towns[k]
        return None, []
    m = re.search(r'\b(?:CITY|VILLAGE) OF ([A-Z .\'-]+?)(?:\s+(?:WATER|SCHOOL|SCHOOLS|BOARD|DEPT|DEPARTMENT|PUBLIC|PARKS|HOUSING)\b|$|[,(])', t)
    if m and m.group(1).strip() in towns:
        return m.group(1).strip(), towns[m.group(1).strip()]
    return None, []


def geocode(addresses, cache, today):
    """Census batch geocoding for addresses not yet in the cache (and unmatched ones after a while)."""
    todo = []
    for a in addresses:
        c = cache.get(a)
        if c is None or (c['match'] != 'Match' and (today - datetime.date.fromisoformat(c['when'])).days >= RETRY_UNMATCHED_DAYS):
            todo.append(a)
    log('geocoding', len(todo), 'addresses')
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
    # Second pass for unmatched addresses: street and city in separate fields, trying the last 1-3 words
    # before "OH" as the city. Kept only if every split that matches lands in the same Ohio county.
    retry = [a for a in addresses if cache.get(a, {}).get('match') != 'Match' and
             ('alt_when' not in cache[a] or (today - datetime.date.fromisoformat(cache[a]['alt_when'])).days >= RETRY_UNMATCHED_DAYS)]
    log('second pass for', len(retry), 'addresses')
    rows = []
    for k, a in enumerate(retry):
        m = re.match(r'^(.*\S)\s+OH(?:IO)?\.?,?\s*(\d{5})?(?:-\d{4})?\s*$', a, re.I)
        if not m:
            continue
        words = m.group(1).replace(',', ' ').split()
        for n in (1, 2, 3):
            if len(words) > n + 1:
                rows.append(('%d_%d' % (k, n), ' '.join(words[:-n]), ' '.join(words[-n:]), m.group(2) or ''))
    found = defaultdict(set)
    for s in range(0, len(rows), 1000):
        chunk = rows[s:s + 1000]
        body_csv = '\n'.join('%s,"%s","%s",OH,%s' % (i, st.replace('"', ''), ci.replace('"', ''), z) for i, st, ci, z in chunk)
        bd = '----trk' + hashlib.md5(body_csv.encode()).hexdigest()[:12]
        body = ('--%s\r\nContent-Disposition: form-data; name="benchmark"\r\n\r\nPublic_AR_Current\r\n'
                '--%s\r\nContent-Disposition: form-data; name="vintage"\r\n\r\nCurrent_Current\r\n'
                '--%s\r\nContent-Disposition: form-data; name="addressFile"; filename="a.csv"\r\nContent-Type: text/csv\r\n\r\n%s\r\n--%s--\r\n'
                % (bd, bd, bd, body_csv, bd)).encode()
        raw, _ = get(CENSUS, data=body, headers={'Content-Type': 'multipart/form-data; boundary=' + bd})
        for f in csv.reader(io.StringIO(raw.decode('utf-8', 'replace'))):
            if len(f) > 9 and f[2] == 'Match' and f[8] and f[9] and re.match(r'^\d+_\d$', f[0]):
                found[int(f[0].split('_')[0])].add((f[8] + f[9], f[4]))
    for k, a in enumerate(retry):
        counties = {c for c, _ in found.get(k, ())}
        alt = None
        if len(counties) == 1 and next(iter(counties)).startswith('39'):
            c = next(iter(counties))
            alt = {'county': c, 'matched': sorted(m for cc, m in found[k] if cc == c)[0]}
        cache[a]['alt'] = alt
        cache[a]['alt_when'] = today.isoformat()


def place(addr, cache, zc, names, towns, building=''):
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
    alt = c.get('alt')
    if alt and alt['county'].startswith('39') and (not zcs or alt['county'] in zcs):
        cty = names[alt['county']]
        return cty, 'Census address match with street and city entered separately (%s)' % alt['matched'], [cty], True
    zn = {names[x] for x in zcs} if zcs else set()
    codes = sorted({ODOT[m] for m in re.findall(r'\b([A-Z]{3})-\d', (addr or '') + ' ' + (building or '')) if m in ODOT})
    if len(codes) == 1 and (not zn or codes[0] in zn):
        return codes[0], 'ODOT county code in the project name (%s)' % next(k for k, v in ODOT.items() if v == codes[0]), [codes[0]], True
    if zcs and len(zcs) == 1:
        cty = names[next(iter(zcs))]
        return cty, 'ZIP %s lies entirely in this county' % z, [cty], True
    if zcs:
        cands = [names[x] for x in sorted(zcs, key=lambda x: -zcs[x])]
        return '', 'County not confirmed: ZIP %s spans %s, and the address could not be matched' % (z, ', '.join(x.title() for x in cands)), cands, False
    town, tc = town_in(addr, towns)
    if tc:
        return '', 'County not confirmed: no usable ZIP and the address could not be matched; the town in the address, %s, lies in %s' % (town.title(), ', '.join(x.title() for x in tc)), list(tc), False
    return '', 'County not confirmed: no usable ZIP, and the address could not be matched', [], False


def likely(s, towns, siblings):
    """For a site whose county is not confirmed: a county that every available clue points to. Clues:
    the town in the address or in the owner's name ("City of ...") lies entirely in one county; the owner
    or building names a county; the other sites on the same notice were placed in one county. Returns
    (county, [evidence]) or (None, [])."""
    ev = []
    for text, before_oh, what in ((s['address'], True, 'the town in the address'), (s['owner'], False, "the town in the owner's name")):
        town, tc = town_in(text, towns, before_oh)
        if town:
            inside = [c for c in tc if not s['candidates'] or c in s['candidates']]
            if len(inside) == 1 and (len(tc) == 1 or s['candidates']):
                ev.append((inside[0], '%s, %s, %s' % (what, town.title(), 'lies in %s County' % inside[0].title() if len(tc) == 1 else 'lies partly in %s County, the only one the ZIP shares' % inside[0].title())))
    named = sorted({c for c in set(ODOT.values()) if re.search(r'\b' + re.escape(c) + r'\s+COUNTY\b', (s['owner'] + ' ' + s['building']).upper())})
    if len(named) == 1:
        ev.append((named[0], 'the owner or building name mentions %s County' % named[0].title()))
    sib = {o['county'] for o in siblings if o is not s and o['confirmed']}
    if len(sib) == 1:
        c = next(iter(sib))
        ev.append((c, 'the other sites on notice %s are in %s County' % (s['id'], c.title())))
    counties = {c for c, _ in ev}
    if len(counties) == 1:
        c = next(iter(counties))
        if not s['candidates'] or c in s['candidates']:
            return c, [e for _, e in ev]
    return None, []


def locals_of(counties):
    return {k for k, cs in LOCALS.items() if set(counties) & set(cs)}


def alerts(s):
    a = []
    o = s['owner'].upper()
    if any(w in o for w in ('LAND BANK', 'LANDBANK', 'REUTILIZATION', 'LAND REUTIL')):
        a.append('Land bank: one notice can cover many houses')
    if any(w in s['building'].upper() for w in ('ANNUAL', 'MAINTENANCE', 'O&M', 'O & M')) or 'ANNUAL' in s['activity'].upper():
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
    towns = load_places(names)
    json.dump(towns, open(os.environ.get('TOWN_COPY', '/tmp/towns.json'), 'w'))
    for s in sites:
        if s['address']:
            s['county'], s['county_basis'], s['candidates'], s['confirmed'] = place(s['address'], cache, zc, names, towns, s['building'])
        else:
            s['county'], s['county_basis'], s['candidates'], s['confirmed'] = '', 'County not confirmed: the notice has no address ("Contact Ohio EPA for Details")', [], False
        s['alerts'] = alerts(s)
    by_notice = defaultdict(list)
    for s in sites:
        by_notice[s['id']].append(s)
    order = lambda ks: sorted(ks, key=lambda k: int(k))
    for s in sites:
        s['likely'], s['likely_evidence'] = (None, [])
        if s['confirmed']:
            s['locals'], s['maybe_locals'] = order(locals_of([s['county']])), []
        else:
            if not s['candidates'] and s['owner']:
                town, tc = town_in(s['owner'], towns, False)
                if tc:
                    s['candidates'] = list(tc)
                    s['county_basis'] += "; the town in the owner's name, %s, lies in %s" % (town.title(), ', '.join(x.title() for x in tc))
            s['likely'], s['likely_evidence'] = likely(s, towns, by_notice[s['id']])
            every = [locals_of([c]) for c in s['candidates']]
            sure = set.intersection(*every) if every else set()  # Locals that cover every county the site could be in
            if s['likely']:
                sure |= locals_of([s['likely']])
            s['locals'] = order(sure)
            s['maybe_locals'] = order(set().union(*every) - sure) if every else []
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
              'likely': sum(1 for s in sites if s['likely']),
              'not_confirmed_no_local': sum(1 for s in sites if not s['confirmed'] and not s['locals']),
              'basis': dict(sorted(_count(('Census address match' if s['county_basis'].startswith('Census') else 'ZIP in one county' if s['county_basis'].startswith('ZIP')
                                            else 'ODOT county code' if s['county_basis'].startswith('ODOT') else 'Likely county' if s['likely']
                                            else 'County not confirmed') for s in sites).items())),
              'locals': {k: dict(_local_counts([s for s in sites if k in s['locals']]), maybe=sum(1 for s in sites if k in s['maybe_locals'])) for k in LOCALS},
              'local_labels': LOCAL_LABEL, 'local_counties': LOCALS}
    json.dump({'sites': [j(s) for s in sites]}, open(os.path.join(DATA, 'sites.json'), 'w'), indent=0, sort_keys=True)
    wanted = {'Local_%s.xlsx' % k for k in LOCALS} | {'All_Ohio.xlsx', 'Needs_review.xlsx'}
    for f in os.listdir(os.path.join(DATA, 'reports')):  # remove reports for Locals no longer listed
        if f.endswith('.xlsx') and f not in wanted:
            os.remove(os.path.join(DATA, 'reports', f))
    for k in LOCALS:
        write_xlsx(os.path.join(DATA, 'reports', 'Local_%s.xlsx' % k), LOCAL_LABEL[k], [s for s in sites if k in s['locals']], status, LOCALS[k],
                   maybe=[s for s in sites if k in s['maybe_locals']])
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
            'not_confirmed': sum(1 for s in ss if not s['confirmed']), 'likely': sum(1 for s in ss if s['likely'])}


def county_label(s):
    if s['county']:
        return s['county'].title()
    if s['likely']:
        return 'Likely ' + s['likely'].title()
    return ('Not confirmed: ' + ' or '.join(c.title() for c in s['candidates'])) if s['candidates'] else 'Not confirmed'


SHEET_COLS = [  # (header, key or function, width). Every column of the Ohio EPA list is here, plus the tracker's own.
    ('Notification ID', 'id', 12), ('County', county_label, 16),
    ('County check', lambda s: '' if (s['confirmed'] and not s['review']) else (('Likely %s County: %s. ' % (s['likely'].title(), '; '.join(s['likely_evidence']))) if s['likely'] else '') + s['county_basis'], 50),
    ('Counted for Locals', lambda s: ', '.join(LOCAL_LABEL[k] for k in s['locals']), 18),
    ('Might also be in', lambda s: ', '.join(LOCAL_LABEL[k] for k in s['maybe_locals']), 18),
    ('Organizer alerts', lambda s: '; '.join(s['alerts']), 30),
    ('Owner', 'owner', 30), ('Owner phone', 'owner_phone', 14), ('Building', 'building', 30), ('Address', 'address', 40),
    ('Activity', 'activity', 22),
    ('Abatement', 'abate_status', 11), ('Abatement start', 'abate_start', 12), ('Abatement completion', 'abate_end', 12),
    ('Abatement note', 'abate_note', 24), ('Setup date', 'setup', 12),
    ('Abatement on hold (EPA: off-site on hold)', 'abate_hold', 12), ('Abatement resumed (EPA: on-site off hold)', 'abate_resume', 12),
    ('Abatement contractor', 'abate_contractor', 30), ('Abatement contractor ID', 'abate_contractor_id', 12),
    ('Abatement contractor phone', 'abate_phone', 14), ('On-site specialist', 'specialist', 20),
    ('Specialist cert number', 'specialist_cert', 14), ('Specialist cert expires', 'cert_exp', 12),
    ('Demolition', 'demo_status', 11), ('Demolition start', 'demo_start', 12), ('Demolition completion', 'demo_end', 12),
    ('Demolition note', 'demo_note', 24), ('Demolition on hold', 'demo_hold', 12), ('Demolition resumed', 'demo_resume', 12),
    ('Demolition contractor', 'demo_contractor', 30), ('Demolition contractor phone', 'demo_phone', 14),
    ('RACM pipe (ft)', 'racm_pipe', 9), ('RACM surface (sq ft)', 'racm_surface', 9), ('RACM volume (cu ft)', 'racm_volume', 9),
    ('Non-friable to remove: Cat 1 pipe (ft)', 'nf_rm_c1_pipe', 9), ('Non-friable to remove: Cat 1 surface (sq ft)', 'nf_rm_c1_surface', 9),
    ('Non-friable to remove: Cat 2 pipe (ft)', 'nf_rm_c2_pipe', 9), ('Non-friable to remove: Cat 2 surface (sq ft)', 'nf_rm_c2_surface', 9),
    ('Non-friable left in place: Cat 1 pipe (ft)', 'nf_stay_c1_pipe', 9), ('Non-friable left in place: Cat 1 surface (sq ft)', 'nf_stay_c1_surface', 9),
    ('Non-friable left in place: Cat 2 pipe (ft)', 'nf_stay_c2_pipe', 9), ('Non-friable left in place: Cat 2 surface (sq ft)', 'nf_stay_c2_surface', 9),
]
SHEETS = [('Abatement upcoming', 'abate_status', 'Upcoming', 'abate_start'), ('Abatement underway', 'abate_status', 'Underway', 'abate_end'),
          ('Demolition upcoming', 'demo_status', 'Upcoming', 'demo_start'), ('Demolition underway', 'demo_status', 'Underway', 'demo_end')]


def sheet_rows(sites, key, val, sortkey):
    ss = [s for s in sites if s[key] == val]
    return sorted(ss, key=lambda s: (s[sortkey] or datetime.date.max, s['id'], s['address']))


def write_xlsx(path, title, sites, status, counties, maybe=None):
    wb = xlsxwriter.Workbook(path, {'constant_memory': False, 'strings_to_formulas': False, 'strings_to_urls': False, 'strings_to_numbers': False})
    wb.set_properties({'title': title, 'created': datetime.datetime(2026, 1, 1)})
    bold = wb.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'bg_color': '#E8EEF6', 'border': 1})
    datef = wb.add_format({'num_format': 'yyyy-mm-dd'})
    warn = wb.add_format({'bg_color': '#FFF2CC'})
    sheets = [(n, sheet_rows(sites, k, v, sk)) for n, k, v, sk in SHEETS]
    if maybe is not None:
        sheets.append(('Might be in your area', sorted(maybe, key=lambda s: (min([d for d in (s['abate_start'], s['demo_start']) if d] or [datetime.date.max]), s['id'], s['address']))))
    ws = wb.add_worksheet('About')
    lines = [title, '',
             'Asbestos removal (abatement) and demolition jobs in Ohio that are starting soon or going on now.',
             'From the notices contractors must file with the Ohio EPA. EPA list as of %s; jobs counted as of %s.' % (status['epa_as_of'], status['today']),
             'Counties: %s.' % (', '.join(c.title() for c in counties) if counties else 'all of Ohio'),
             '',
             'THE SHEETS',
             '"Upcoming": the work starts after today.',
             '"Underway": the work has started and is not past its completion date. With no completion date, it counts as underway for %d days.' % RECENT_NO_END_DAYS]
    if maybe is not None:
        lines.append('"Might be in your area": the address could be in this Local or a neighboring one. Check it before acting.')
    lines += ['Rows in this file: ' + '; '.join('%s %d' % (n, len(r)) for n, r in sheets) + '. A job with both abatement and demolition is on both sheets.',
              '',
              'TIPS',
              'Each row is one address. One EPA notice can cover many addresses (land banks often list many houses).',
              'Yellow rows: the county could not be confirmed from the address. The "County check" column says why.',
              '"Organizer alerts" marks land bank notices, maintenance or annual permits, and work set to run more than a year.',
              'Use the arrows in the top row to sort or filter. Dates and amounts sort correctly.',
              '',
              'WHERE THE DATA COMES FROM',
              'Every column of the Ohio EPA list is included, unchanged. The tracker adds the County, County check, Counted for Locals, Might also be in, Organizer alerts, and Abatement and Demolition status columns.',
              'County comes from the Census Bureau address lookup, or from the ZIP code when the whole ZIP is in one county. "Likely" means every other clue (town, owner, other addresses on the notice) points to that county.',
              'Before a file is published, every value in it is checked against a second, separate reading of the EPA list.']
    for i, t in enumerate(lines):
        ws.write(i, 0, t, bold if i == 0 or t in ('THE SHEETS', 'TIPS', 'WHERE THE DATA COMES FROM') else None)
    ws.set_column(0, 0, 120)
    for name, rows in sheets:
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
                elif k in NUMERIC and isinstance(v, str) and re.fullmatch(r'\d{1,3}(,\d{3})+(\.\d+)?|\d+(\.\d+)?', v):
                    ws.write_number(r, c, float(v.replace(',', '')), fmt_row)
                else:
                    ws.write(r, c, v, fmt_row)
        ws.freeze_panes(1, 2)
        ws.autofilter(0, 0, max(1, len(rows)), len(SHEET_COLS) - 1)
    wb.close()


if __name__ == '__main__':
    build()
