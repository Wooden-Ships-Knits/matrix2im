"""Raw MATRIX JSON -> FINAL IM MASTER JSON (for review before the PRD).

records      : one per colorway, keys = IM column names (header row 54)
               DESCRIPTION, COLOR1..COLORn (only slots that have a yarn; IM spells 'COLOR 4' with a space),
               WS TAG COLOR (null = no rule yet, will not be written), WHS (IM column CE)
               '_ref' = where it came from in MATRIX (never written to IM)
needs_review : everything the rules could not decide, linked by record index
"""
import json, re, sys, datetime, collections
from to_im_json import transform, IM_COLOR_KEYS

DESCRIPTOR = re.compile(r'\b(STRIPE|STRIPED|MARL|MULTI)$')


def main(src, dst):
    raw = json.load(open(src))
    recs = transform(raw)
    records, review = [], []
    seen = collections.defaultdict(list)
    for i, r in enumerate(recs):
        m = r['_meta']
        out = {'DESCRIPTION': r['DESCRIPTION']}
        for k in IM_COLOR_KEYS:
            if r[k]:
                out[k] = r[k]
        out['WS TAG COLOR'] = r['WS TAG COLOR']
        out['WHS'] = r['WHS']
        sec = f" ({m['section']})" if m['section'] else ''
        out['_ref'] = f"{m['sheet']}{sec} No.{m['no']} | {m['matrix_cell']} | {m['group']}"
        records.append(out)

        colorway = '/'.join(r[k] for k in IM_COLOR_KEYS if r[k])
        for p in m['problems']:
            review.append({'record': i, 'DESCRIPTION': r['DESCRIPTION'], 'colorway': colorway,
                           'issue': p, 'tag_instruction': m['tag_instruction'], '_ref': out['_ref']})
        if m['whs_capped_by_style']:
            review.append({'record': i, 'DESCRIPTION': r['DESCRIPTION'], 'colorway': colorway,
                           'issue': f"info: WHS warna {m['color_whs']} diturunkan ke {r['WHS']} (batas WHS style)",
                           '_ref': out['_ref']})
        desc = [r[k] for k in IM_COLOR_KEYS if r[k] and DESCRIPTOR.search(r[k].upper())]
        if desc:
            src_cell = next((c for s in raw['styles'] if s['style'] == r['DESCRIPTION']
                             for g in ('color_above_black_line', 'color_between_black_and_red_line', 'color_below_red_line')
                             for c in s[g] if c['cell'] == m['matrix_cell']), None)
            review.append({'record': i, 'DESCRIPTION': r['DESCRIPTION'], 'colorway': colorway,
                           'issue': f'COLOR berisi deskripsi, bukan nama benang: {desc}',
                           'comment_di_matrix': src_cell and src_cell['comment'], '_ref': out['_ref']})
        seen[(r['DESCRIPTION'], colorway)].append(i)
    for (d, cw), idx in seen.items():
        if len(idx) > 1:
            whs = {records[j]['WHS'] for j in idx}
            review.append({'record': idx, 'DESCRIPTION': d, 'colorway': cw,
                           'issue': 'style muncul di 2 sheet MATRIX, record dobel'
                                    + (f' dan WHS beda: {sorted(whs)}' if len(whs) > 1 else ''),
                           '_ref': [records[j]['_ref'] for j in idx]})
    for w in raw.get('warnings', []):
        review.append({'record': None, 'issue': f'warning ekstraksi: {w}'})

    review.sort(key=lambda x: (x['record'] if isinstance(x['record'], int) else
                               (x['record'][0] if isinstance(x['record'], list) else 10**6)))
    summary = {
        'total_records': len(records),
        'total_styles': len({r['DESCRIPTION'] for r in records}),
        'WHS': dict(collections.Counter(r['WHS'] for r in records)),
        'WS_TAG_COLOR_filled': sum(1 for r in records if r['WS TAG COLOR']),
        'WS_TAG_COLOR_null': sum(1 for r in records if not r['WS TAG COLOR']),
        'needs_review_items': len(review),
        'needs_review_by_issue': dict(collections.Counter(re.split(r'[:(]', x['issue'])[0].strip() for x in review)),
    }
    doc = {
        'target': 'F27 IM MASTER (sheet F27, header baris 54)',
        'source': 'F27 MATRIX MASTER',
        'generated': datetime.date.today().isoformat(),
        'rules': {
            'DESCRIPTION': 'nama style dari MATRIX (tanpa size); berlaku untuk semua size di IM',
            'COLOR1-COLOR18': '1 benang = 1 kolom, nilai apa adanya dari MATRIX; benang berulang (stripe) ditulis sekali; '
                              'slot tanpa benang tidak dicantumkan (tidak ditulis ke IM)',
            'WS TAG COLOR': 'instruksi TAG, C1/C2/... diganti nama warna benang (tanpa prefix YCA/YST dan kode); '
                            'STRIPED -> STRIPE; null = belum ada aturan, tidak ditulis',
            'WHS': 'kolom CE. Warna: di bawah garis merah = SY; di atas garis merah = 000 & SY, '
                   'kecuali font ungu/pink = 000. Style (background No): ungu 000, abu 000 & SY, merah SY. '
                   'WHS style = batas atas (irisan)',
            '_ref': 'asal di MATRIX: sheet (sub-koleksi) No | sel | posisi terhadap garis; tidak ditulis ke IM',
        },
        'summary': summary,
        'records': records,
        'needs_review': review,
    }
    json.dump(doc, open(dst, 'w'), indent=2, ensure_ascii=False)
    return doc


if __name__ == '__main__':
    d = main(sys.argv[1], sys.argv[2])
    print(json.dumps(d['summary'], indent=1, ensure_ascii=False))
