"""One-off look at the Ohio EPA asbestos/demolition notification file, run on GitHub.
Writes probe/report.json and a readable summary. Changes nothing anywhere else."""
import datetime, io, json, os, re, random, urllib.request, urllib.parse
from collections import Counter

UA = 'ohio-demo-abatement-tracker/1.0 (public Ohio EPA notification data)'
EPA = 'https://dam.assets.ohio.gov/raw/upload/epa.ohio.gov/Public/DAPC/NotificationList.xlsx'
out = {}


def get(url, data=None, headers=None, timeout=120):
    req = urllib.request.Request(url, data=data, headers=dict({'User-Agent': UA}, **(headers or {})))
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read(), dict(r.headers)


raw, hdr = get(EPA)
out['epa'] = {'bytes': len(raw), 'last_modified': hdr.get('Last-Modified'), 'content_type': hdr.get('Content-Type')}
import openpyxl
wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
out['sheets'] = wb.sheetnames
ws = wb[wb.sheetnames[0]]
rows = [list(r) for r in ws.iter_rows(values_only=True)]
out['n_rows_raw'] = len(rows)
out['top'] = [[str(c)[:60] if c is not None else None for c in r[:30]] for r in rows[:8]]
hi = next(i for i, r in enumerate(rows) if r and any(isinstance(c, str) and 'address' in c.lower() for c in r))
head = [str(c).strip() if c is not None else '' for c in rows[hi]]
out['header_row'] = hi
out['columns'] = head
data = [r for r in rows[hi + 1:] if any(c not in (None, '') for c in r)]
out['n_data'] = len(data)
cols = {h: i for i, h in enumerate(head) if h}
out['types'] = {h: dict(Counter(type(r[i]).__name__ for r in data if i < len(r))) for h, i in cols.items()}
out['sample'] = [{h: (str(r[i])[:80] if i < len(r) and r[i] is not None else None) for h, i in cols.items()} for r in data[:6]]
random.seed(1)
out['random'] = [{h: (str(r[i])[:80] if i < len(r) and r[i] is not None else None) for h, i in cols.items()} for r in random.sample(data, min(12, len(data)))]
# Columns with few distinct values (notification type, activity, status...)
out['categorical'] = {}
for h, i in cols.items():
    vals = Counter(str(r[i]).strip() for r in data if i < len(r) and r[i] not in (None, ''))
    if 0 < len(vals) <= 40:
        out['categorical'][h] = vals.most_common(40)
    out.setdefault('distinct', {})[h] = len(vals)
# Address shapes
ai = next(i for h, i in cols.items() if 'address' in h.lower())
addrs = [str(r[ai]) for r in data if ai < len(r) and r[ai]]
fives = [re.findall(r'(?<!\d)(\d{5})(?:-\d{4})?(?!\d)', a) for a in addrs]
out['address'] = {
    'column': head[ai], 'n': len(addrs),
    'no_5digit': sum(1 for f in fives if not f),
    'first_not_last': sum(1 for f in fives if len(f) > 1 and f[0] != f[-1]),
    'first_not_ohio': sum(1 for f in fives if f and not ('43000' <= f[0] <= '45999')),
    'last_not_ohio': sum(1 for f in fives if f and not ('43000' <= f[-1] <= '45999')),
    'examples_first_not_last': [a for a, f in zip(addrs, fives) if len(f) > 1 and f[0] != f[-1]][:15],
    'examples_no_zip': [a for a, f in zip(addrs, fives) if not f][:15],
    'examples_last_not_ohio': [a for a, f in zip(addrs, fives) if f and not ('43000' <= f[-1] <= '45999')][:10],
}
# Census geocoder: one-line and batch
tests = addrs[:3]
geo = {}
for a in tests:
    try:
        q = urllib.parse.urlencode({'address': a, 'benchmark': 'Public_AR_Current', 'vintage': 'Current_Current', 'format': 'json', 'layers': 'Counties'})
        b, _ = get('https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress?' + q, timeout=60)
        j = json.loads(b)
        m = j['result']['addressMatches']
        geo[a] = [(x['matchedAddress'], x['geographies']['Counties'][0]['NAME'], x['geographies']['Counties'][0]['GEOID']) for x in m[:2]]
    except Exception as e:
        geo[a] = 'ERR ' + str(e)[:200]
out['census_oneline'] = geo
try:
    boundary = '----probe'
    lines = '\n'.join('%d,"%s",,,' % (k, a.replace('"', '')) for k, a in enumerate(addrs[:5]))
    body = ('--%s\r\nContent-Disposition: form-data; name="benchmark"\r\n\r\nPublic_AR_Current\r\n'
            '--%s\r\nContent-Disposition: form-data; name="vintage"\r\n\r\nCurrent_Current\r\n'
            '--%s\r\nContent-Disposition: form-data; name="addressFile"; filename="a.csv"\r\nContent-Type: text/csv\r\n\r\n%s\r\n--%s--\r\n'
            % (boundary, boundary, boundary, lines, boundary)).encode()
    b, _ = get('https://geocoding.geo.census.gov/geocoder/geographies/addressbatch', data=body,
               headers={'Content-Type': 'multipart/form-data; boundary=' + boundary}, timeout=120)
    out['census_batch'] = b.decode('utf-8', 'replace')[:2000]
except Exception as e:
    out['census_batch'] = 'ERR ' + str(e)[:300]
try:
    b, h = get('https://www2.census.gov/geo/docs/maps-data/data/rel2020/zcta520/tab20_zcta520_county20_natl.txt', timeout=120)
    lines = b.decode('utf-8', 'replace').split('\n')
    out['zcta_county'] = {'bytes': len(b), 'header': lines[0][:400], 'ohio_example': [l for l in lines if '|39049|' in l or '|39049' in l][:2]}
except Exception as e:
    out['zcta_county'] = 'ERR ' + str(e)[:300]
os.makedirs('probe', exist_ok=True)
json.dump(out, open('probe/report.json', 'w'), indent=1, default=str)
print(json.dumps({k: out[k] for k in ('epa', 'sheets', 'header_row', 'columns', 'n_data', 'address')}, indent=1, default=str))
