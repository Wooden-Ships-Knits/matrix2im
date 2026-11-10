"""MATRIX extract JSON -> IM MASTER JSON (keys = IM column names, header row 54).

One record per colorway. Keys written to IM:
  DESCRIPTION            <- style name (without size; the writer matches 'DESCRIPTION - X/S' etc.)
  COLOR1 .. COLOR18      <- 1 yarn = 1 colour, value as written in MATRIX (IM spells 'COLOR 4' with a space)
  WS TAG COLOR           <- TAG instruction with C1, C2 ... replaced by the yarn colour name
  WHS                    <- IM column CE (NOT column G). Values: "000" | "SY" | "000 & SY"

WHS rules (agreed with user):
  colour level : below red line -> SY
                 above red line (above or below black line) -> 000 & SY,
                 but font purple/pink (000 ONLY COLOR / 000 REPEAT COLOR) -> 000
  style level  : background of the No cell: purple -> 000 | grey -> 000 & SY | red -> SY
  final        : style WHS is the upper bound -> intersection of style and colour WHS
                 (e.g. style 000 + colour 000 & SY -> 000). Empty intersection -> null + review.
Keys starting with '_' are metadata and are never written to IM.
"""
import json, re, sys

IM_COLOR_KEYS = ['COLOR1', 'COLOR2', 'COLOR3', 'COLOR 4'] + [f'COLOR{i}' for i in range(5, 19)]
PREFIX = r'^(YCA|YST|YTX|YP|BJ|POODLE)\s+'
TAG_ALIAS = {'BLACK B': 'BLACK'}           # name in tag differs from yarn name
GROUPS = [('color_above_black_line', 'above_black'),
          ('color_between_black_and_red_line', 'between_black_red'),
          ('color_below_red_line', 'below_red')]
WHS_SETS = {'000': {'000'}, 'SY': {'SY'}, '000 & SY': {'000', 'SY'}}
FONT_000 = {'000 ONLY COLOR', '000 REPEAT COLOR'}
FONT_OK = FONT_000 | {'NEW COLOR', 'REPEAT COLOR'}


def color_name(yarn):
    """'YCA KHAKI - 749' -> 'KHAKI'  |  'YCA BLACK B - 1708 B' -> 'BLACK'  |  'DARK DUSK' -> 'DARK DUSK'"""
    n = re.sub(PREFIX, '', yarn.strip().upper())
    n = re.sub(r'\s*-\s*\d+\s*[A-Z]?$', '', n).strip()
    return TAG_ALIAS.get(n, n)


def unique(seq):
    seen, out = set(), []
    for s in seq:
        if s.upper() not in seen:
            seen.add(s.upper()); out.append(s)
    return out


def build_tag(tag, yarns):
    """Replace C1, C2 ... in the TAG instruction with yarn colour names. Returns (value, problem)."""
    if not tag:
        return None, 'TAG kosong'
    t = tag.upper()
    if '|' in t or 'FOR SOLID' in t or 'FOR HEATHER' in t or 'FOR MARL' in t or 'ABBREVIAT' in t:
        return None, 'TAG kondisional / singkatan: perlu aturan khusus'
    body = re.sub(r'^TAG\s*:?\s*', '', t).strip()
    missing = [m for m in re.findall(r'\bC(\d+)\b', body) if int(m) > len(yarns)]
    if missing:
        return None, f'TAG minta C{missing[0]} tapi colorway hanya punya {len(yarns)} benang'
    val = re.sub(r'\bC(\d+)\b', lambda m: color_name(yarns[int(m.group(1)) - 1]), body)
    val = re.sub(r'\bSTRIPED\b', 'STRIPE', val)      # IM note: STRIPE tanpa D
    val = re.sub(r'\b(STRIPE|MARL|MULTI|HEATHERED)( \1\b)+', r'\1', val)   # colour name already ends with suffix
    return re.sub(r'\s*/\s*', '/', val).strip(), None


def color_whs(group, font):
    """WHS from the colour's position (red line) and font colour. Returns (value, problem)."""
    if font not in FONT_OK:
        return None, f'warna font "{font}" belum punya aturan WHS'
    if group == 'below_red':
        return 'SY', None
    return ('000' if font in FONT_000 else '000 & SY'), None


def final_whs(style_whs, col_whs):
    """Style WHS is the upper bound: intersection of both sets."""
    if col_whs is None:
        return None, None
    if style_whs is None:
        return col_whs, 'WHS style tidak diketahui (background nomor kosong), dipakai WHS warna'
    inter = WHS_SETS[style_whs] & WHS_SETS[col_whs]
    if not inter:
        return None, f'WHS style {style_whs} bertentangan dengan WHS warna {col_whs}'
    return ('000 & SY' if inter == {'000', 'SY'} else inter.pop()), None


def transform(extract):
    records = []
    for s in extract['styles']:
        for group_key, group in GROUPS:
            for i, c in enumerate(s[group_key], 1):
                yarns = unique(c['yarns'])
                rec = {'DESCRIPTION': s['style']}
                for k, key in enumerate(IM_COLOR_KEYS):
                    rec[key] = yarns[k] if k < len(yarns) else None
                tag, p_tag = build_tag(s['tag'], yarns)
                cw, p_cw = color_whs(group, c['font_color'])
                whs, p_whs = final_whs(s.get('style_whs'), cw)
                rec['WS TAG COLOR'] = tag
                rec['WHS'] = whs
                rec['_meta'] = {
                    'no': s['no'], 'sheet': s['sheet'], 'section': s['section'], 'group': group, 'order': i,
                    'matrix_cell': c['cell'], 'tag_instruction': s['tag'], 'font_color': c['font_color'],
                    'style_whs': s.get('style_whs'), 'color_whs': cw, 'whs_capped_by_style': bool(cw and whs and cw != whs),
                    'yarns_raw': c['yarns'] if len(c['yarns']) != len(yarns) else None,
                    'too_many_yarns': len(yarns) > 18,
                    'problems': [p for p in (p_tag, p_cw, p_whs) if p],
                }
                if len(yarns) > 18:
                    rec['_meta']['problems'].append(f'{len(yarns)} benang unik, slot COLOR hanya 18')
                records.append(rec)
    return records


if __name__ == '__main__':
    recs = transform(json.load(open(sys.argv[1])))
    json.dump({'target': 'F27 IM MASTER', 'total_records': len(recs), 'records': recs},
              open(sys.argv[2], 'w'), indent=2, ensure_ascii=False)
    print('records', len(recs), '| with problems', sum(1 for r in recs if r['_meta']['problems']),
          '| WS TAG COLOR filled', sum(1 for r in recs if r['WS TAG COLOR']),
          '| WHS filled', sum(1 for r in recs if r['WHS']))
