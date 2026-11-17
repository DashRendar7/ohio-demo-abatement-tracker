"""Second look at the EPA file: duplicate notification IDs, hold dates, annual permits, and a full
Census address-matching trial on every active or upcoming project. Writes probe/report2.json."""
import datetime, io, json, os, re, time, urllib.request
from collections import Counter, defaultdict
from zoneinfo import ZoneInfo
import openpyxl

UA = 'ohio-demo-abatement-tracker/1.0 (public Ohio EPA notification data)'
EPA = 'https://dam.assets.ohio.gov/raw/upload/epa.ohio.gov/Public/DAPC/NotificationList.xlsx'
ZCTA = 'https://www2.census.gov/geo/docs/maps-data/data/rel2020/zcta520/tab20_zcta520_county20_natl.txt'
out = {}


def get(url, data=None, headers=None, timeout=300):
    req = urllib.request.Request(url, data=data, headers=dict({'User-Agent': UA}, **(headers or {})))
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


wb = openpyxl.load_workbook(io.BytesIO(get(EPA)), read_only=True, data_only=True)
rows = [list(r) for r in wb[wb.sheetnames[0]].iter_rows(values_only=True)]
head = [str(c).strip() if c is not None else '' for c in rows[4]]
data = [r + [None] * (len(head) - len(r)) for r in rows[5:] if any(c not in (None, '') for c in r)]
C = {h: i for i, h in enumerate(head)}
ID = C['Notification ID']

# 1. Duplicate IDs: what differs between rows that share an ID, and does row order mean anything?
groups = defaultdict(list)
for k, r in enumerate(data):
    groups[r[ID]].append(k)
dup = {i: ks for i, ks in groups.items() if len(ks) > 1}
diffcols = Counter()
identical = 0
for i, ks in dup.items():
    rs = [data[k] for k in ks]
    cols = [h for h, j in C.items() if len({repr(r[j]) for r in rs}) > 1]
    if not cols:
        identical += 1
    diffcols.update(cols)
out['dups'] = {'ids_with_more_than_one_row': len(dup), 'max_rows_per_id': max((len(v) for v in dup.values()), default=1),
               'size_hist': Counter(len(v) for v in dup.values()).most_common(10), 'fully_identical_groups': identical,
               'columns_that_differ': diffcols.most_common(40)}
ex = []
for i, ks in list(dup.items())[:400]:
    rs = [data[k] for k in ks]
    cols = [h for h, j in C.items() if len({repr(r[j]) for r in rs}) > 1]
    if cols and len(ex) < 12:
        ex.append({'id': i, 'row_positions': ks, 'differs': {h: [str(r[C[h]])[:60] for r in rs] for h in cols},
                   'address': str(rs[0][C['Building Address']])[:80], 'activity': rs[0][C['On Site Activity']]})
out['dup_examples'] = ex
out['id_order'] = {'first_ids': [data[k][ID] for k in range(8)], 'descending_share': sum(1 for a, b in zip(data, data[1:]) if a[ID] >= b[ID]) / max(1, len(data) - 1)}

# 2. Dates, holds, annual permits, land banks (relative to today in Ohio)
today = datetime.datetime.now(ZoneInfo('America/New_York')).date()
out['today'] = str(today)
def d(v):
    return v.date() if isinstance(v, datetime.datetime) else None
names = ['SETUP DATE', 'ABATEMENT DATE', 'COMPLETION DATE', 'OFFSITE ON HOLD DATE', 'ONSITE OFF HOLD DATE', 'Demo Start Date', 'Demo Completion Date', 'Demo On Hold Date', 'Demo Off Hold Date']
times = Counter()
for r in data:
    for n in names:
        v = r[C[n]]
        if isinstance(v, datetime.datetime):
            times[v.strftime('%H:%M')] += 1
out['date_times_of_day'] = times.most_common(8)
bn = Counter()
for r in data:
    b = str(r[C['BUILDING NAME']] or '').upper()
    for w in ('ANNUAL', 'O&M', 'O & M', 'MAINTENANCE', 'VARIOUS', 'MULTIPLE', 'SCATTERED'):
        if w in b:
            bn[w] += 1
out['building_name_words'] = bn.most_common()
lb = Counter()
for r in data:
    o = str(r[C['Owner Name']] or '').upper()
    for w in ('LAND BANK', 'REUTILIZATION', 'LANDBANK', 'LAND REUTIL'):
        if w in o:
            lb[w] += 1
out['owner_words'] = lb.most_common()

def window(r, s, e):
    a, b = d(r[C[s]]), d(r[C[e]])
    return a, b
cur = []
stats = Counter()
for k, r in enumerate(data):
    act = r[C['On Site Activity']]
    aa, ab = window(r, 'ABATEMENT DATE', 'COMPLETION DATE')
    da, db = window(r, 'Demo Start Date', 'Demo Completion Date')
    keep = False
    for kind, a, b in (('abate', aa, ab), ('demo', da, db)):
        if a is None and b is None:
            continue
        if b is not None and b < today:
            continue
        if a is None:
            stats[kind + '_no_start_future_end'] += 1
        if b is None:
            stats[kind + '_no_end'] += 1
            if a is not None and a < today:
                stats[kind + '_started_no_end'] += 1
                continue
        if a and b and b < a:
            stats[kind + '_end_before_start'] += 1
        if a and b and (b - a).days > 300:
            stats[kind + '_longer_than_300_days'] += 1
        keep = True
    if keep:
        cur.append(k)
out['current'] = {'rows': len(cur), 'ids': len({data[k][ID] for k in cur}), 'stats': dict(stats),
                  'activity': Counter(data[k][C['On Site Activity']] for k in cur).most_common()}
long_ex = []
for k in cur:
    r = data[k]
    for s, e in (('ABATEMENT DATE', 'COMPLETION DATE'), ('Demo Start Date', 'Demo Completion Date')):
        a, b = d(r[C[s]]), d(r[C[e]])
        if a and b and (b - a).days > 300 and len(long_ex) < 10:
            long_ex.append([r[ID], str(r[C['Owner Name']])[:40], str(r[C['BUILDING NAME']])[:40], str(a), str(b)])
out['long_examples'] = long_ex

# 3. ZIP -> county from the Census relationship file
zc = defaultdict(dict)
for line in get(ZCTA).decode('utf-8-sig').split('\n')[1:]:
    f = line.split('|')
    if len(f) > 17 and f[9].startswith('39') and f[1]:
        zc[f[1]][f[9]] = zc[f[1]].get(f[9], 0) + int(f[16] or 0)
def zip_of(a):
    m = re.search(r'\bOH(?:IO)?\.?,?\s*(\d{5})(?:-\d{4})?\s*$', a, re.I)
    if m:
        return m.group(1)
    f = re.findall(r'(?<!\d)(\d{5})(?:-\d{4})?(?!\d)', a)
    return f[-1] if f else None
# 4. Census batch geocoding of the current addresses
addrs = sorted({str(data[k][C['Building Address']]).strip() for k in cur if data[k][C['Building Address']]})
res = {}
t0 = time.time()
for s in range(0, len(addrs), 1000):
    chunk = addrs[s:s + 1000]
    csv = '\n'.join('%d,"%s",,,' % (s + j, a.replace('"', '')) for j, a in enumerate(chunk))
    bd = '----trk'
    body = ('--%s\r\nContent-Disposition: form-data; name="benchmark"\r\n\r\nPublic_AR_Current\r\n--%s\r\nContent-Disposition: form-data; name="vintage"\r\n\r\nCurrent_Current\r\n--%s\r\nContent-Disposition: form-data; name="addressFile"; filename="a.csv"\r\nContent-Type: text/csv\r\n\r\n%s\r\n--%s--\r\n' % (bd, bd, bd, csv, bd)).encode()
    try:
        txt = get('https://geocoding.geo.census.gov/geocoder/geographies/addressbatch', data=body, headers={'Content-Type': 'multipart/form-data; boundary=' + bd}).decode('utf-8', 'replace')
    except Exception as e:
        out.setdefault('batch_errors', []).append(str(e)[:200])
        continue
    import csv as csvmod
    for f in csvmod.reader(io.StringIO(txt)):
        if len(f) >= 3:
            i = int(f[0])
            res[addrs[i]] = {'match': f[2], 'type': f[3] if len(f) > 3 else '', 'matched': f[4] if len(f) > 4 else '',
                             'county': (f[8] + f[9]) if len(f) > 9 and f[8] else ''}
out['census_seconds'] = round(time.time() - t0)
mt = Counter(v['match'] + '/' + v['type'] for v in res.values())
agree = Counter()
review = []
for a in addrs:
    z = zip_of(a)
    zcs = zc.get(z, {}) if z else {}
    v = res.get(a, {})
    cty = v.get('county', '')
    if v.get('match') == 'Match' and cty.startswith('39'):
        if not zcs:
            agree['census_only_zip_invalid'] += 1
        elif cty in zcs:
            agree['census_in_zip_counties'] += 1
        else:
            agree['census_not_in_zip_counties_' + v.get('type', '')] += 1
            if len(review) < 15:
                review.append([a, cty, v.get('matched'), sorted(zcs)])
    elif v.get('match') == 'Match':
        agree['census_outside_ohio'] += 1
    else:
        if not zcs:
            agree['no_match_no_valid_zip'] += 1
        else:
            tot = sum(zcs.values())
            top = max(zcs.values()) / tot if tot else 0
            agree['no_match_zip_single_county' if len(zcs) == 1 else ('no_match_zip_99pct' if top >= 0.99 else 'no_match_zip_split')] += 1
out['geocode'] = {'addresses': len(addrs), 'match_types': mt.most_common(), 'agreement': dict(agree), 'conflicts': review,
                  'unmatched_examples': [a for a in addrs if res.get(a, {}).get('match') != 'Match'][:25]}
zs = Counter(len(v) for v in zc.values())
out['zcta_ohio'] = {'zips': len(zc), 'counties_per_zip': sorted(zs.items())}
os.makedirs('probe', exist_ok=True)
json.dump(out, open('probe/report2.json', 'w'), indent=1, default=str)
print(json.dumps({k: out[k] for k in ('dups', 'current', 'geocode')}, indent=1, default=str)[:6000])
