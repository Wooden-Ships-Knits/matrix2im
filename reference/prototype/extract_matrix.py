"""MATRIX MASTER extractor v2: No, style name, TAG, green-box colors, blue-box colors.

Layout rules (anchor based, no fixed row numbers):
- page        = block starting at a cell in col A beginning with 'FALL 27'
- label col   = column holding '- LEBAR/WIDTH' etc. on that page; cards are columns left of it
- NEW row     = row whose cells are exactly NEW / EXACT REPEAT / REPEAT / UPDATE ...; name row = row above
- number row  = first row below the legend with numeric values (cached formula results)
- TAG row     = per card, first cell below the name starting with 'TAG'
- red line    = row on the page whose cells carry a red top border (or red bottom border on the row above)
- green box   = non-empty cells after TAG row and above red line
- blue box    = non-empty cells from red line down to the row above '- LEBAR/WIDTH'
"""
import openpyxl, json, re, sys, collections
from openpyxl.utils import get_column_letter as L

RED = {'FFFF0000', 'FFE40500', 'FFC00000'}
FONT_MEANING = {  # from the legend at the top of every page
    'FF00B050': 'NEW COLOR', 'FFFF00FF': '000 ONLY COLOR', 'FFFF99FF': '000 REPEAT COLOR',
    'FF9900FF': 'YR COLOR', 'FFFF0000': 'HOLD COLOR FOR SY', 'FF0066FF': 'DISCONTINUED COLOR FOR SY',
}
NO_FILL_MEANING = {'FFFF66FF': '000 ONLY STYLE', 'FFFF0000': 'SY ONLY'}
# style-level WHS = background of the number cell (confirmed by user):
# purple -> 000 | grey -> 000 & SY | red -> SY
STYLE_WHS = {'FFFF66FF': '000', 'theme0-0.15': '000 & SY', 'FFD9D9D9': '000 & SY', 'FFFF0000': 'SY'}
NUM_RE = re.compile(r'^(\d+(\.0)?[A-Z]?|S\d+( - [A-Z])?)$')
NR_RE = re.compile(r'^(NEW|EXACT REPEAT|REPEAT|REPEAT WITH UPDATE|UPDATE|SIMILAR)$')
YARN_RE = re.compile(r'\s*/\s*\n\s*|\s*\n\s*|\s*/\s*$')


def norm(v):
    return re.sub(r'[ \t\xa0]+', ' ', str(v)).strip()


def font_rgb(c):
    col = c.font.color if c.font else None
    if col is not None and col.type == 'rgb':
        return col.rgb
    return None


def fill_rgb(c):
    f = c.fill
    if f is None or f.fill_type is None:
        return None
    col = f.fgColor
    return col.rgb if col.type == 'rgb' else (f'theme{col.theme}' if col.type == 'theme' else None)


def fill_detail(c):
    """Fill colour incl. theme tint, e.g. 'FFFF66FF' or 'theme0-0.15'."""
    f = c.fill
    if f is None or f.fill_type is None:
        return None
    col = f.fgColor
    if col.type == 'rgb':
        return col.rgb
    if col.type == 'theme':
        return f'theme{col.theme}{col.tint:+.2f}'
    return f'{col.type}{col.indexed if col.type == "indexed" else ""}'


def is_red(side):
    return side is not None and side.style and side.color is not None and side.color.type == 'rgb' and side.color.rgb in RED


def is_black(side):
    """Thick line that is not red: Excel stores it as auto (indexed 64), theme 1 or plain black."""
    if side is None or side.style not in ('medium', 'thick', 'double'):
        return False
    c = side.color
    if c is None:
        return True
    if c.type == 'rgb':
        return c.rgb in ('FF000000', '00000000')
    if c.type == 'indexed':
        return c.indexed in (8, 64)
    if c.type == 'theme':
        return c.theme == 1
    return c.type == 'auto'


def line_above(ws, r, ci, test):
    """True when a line matching `test` sits between row r-1 and row r in column ci."""
    return test(ws.cell(r, ci).border.top) or (r > 1 and test(ws.cell(r - 1, ci).border.bottom))


def split_yarns(text):
    parts = [norm(p) for p in re.split(r'/\s*\n|\n|/(?=\s*$)', str(text))]
    out = []
    for p in parts:
        out += [norm(x) for x in p.split('/') if norm(x)] if '/' in p and len(p) < 60 else ([p] if p else [])
    return out


def extract(path):
    wb = openpyxl.load_workbook(path)
    wbv = openpyxl.load_workbook(path, data_only=True)
    cards, issues = [], []
    for ws in wb.worksheets:
        wv = wbv[ws.title]
        tops = [c.row for c in ws['A'] if isinstance(c.value, str) and c.value.startswith('FALL 27')]
        for pi, (top, bot) in enumerate(zip(tops, tops[1:] + [ws.max_row + 1]), 1):
            # label column + width row
            labcol = width_row = None
            for r in range(top, bot):
                for c in ws[r]:
                    if isinstance(c.value, str) and norm(c.value) == '- LEBAR/WIDTH':
                        labcol, width_row = c.column, r
                        break
                if labcol:
                    break
            if not labcol:
                issues.append(f'{ws.title} hal.{pi}: label "- LEBAR/WIDTH" tidak ditemukan')
                continue
            ncols = range(1, labcol)
            # NEW / REPEAT row (exact match, so 'NEWPORT ...' is not picked up)
            nr_row = next((r for r in range(top + 6, width_row)
                           if sum(1 for ci in ncols if isinstance(ws.cell(r, ci).value, str)
                                  and NR_RE.match(norm(ws.cell(r, ci).value).upper())) >= 1), None)
            if not nr_row:
                issues.append(f'{ws.title} hal.{pi}: baris NEW/REPEAT tidak ditemukan')
                continue
            name_row = nr_row - 1
            # number row: first row after legend with numeric cached values
            num_row = next((r for r in range(top + 5, name_row)
                            if sum(1 for ci in ncols if wv.cell(r, ci).value is not None
                                   and NUM_RE.match(norm(wv.cell(r, ci).value))) >= 1), None)
            section = None
            if num_row and isinstance(ws.cell(num_row - 1, 1).value, str):
                v = norm(ws.cell(num_row - 1, 1).value)
                if v and not any(k in v.upper() for k in ('COLOR', 'STYLE', 'ORDER')):
                    section = v
            # red line row on this page
            red_hits = collections.Counter()
            for r in range(nr_row, width_row + 1):
                for ci in ncols:
                    if is_red(ws.cell(r, ci).border.top) or is_red(ws.cell(r - 1, ci).border.bottom):
                        red_hits[r] += 1
            red_row = red_hits.most_common(1)[0][0] if red_hits else None
            if red_row is None:
                issues.append(f'{ws.title} hal.{pi}: garis merah tidak ditemukan, semua warna dianggap box hijau')
            for ci in ncols:
                name = ws.cell(name_row, ci).value
                if name in (None, ''):
                    continue
                cl = L(ci)
                tag_row = next((r for r in range(nr_row, width_row)
                                if isinstance(ws.cell(r, ci).value, str) and norm(ws.cell(r, ci).value).upper().startswith('TAG')), None)
                tag_rows = []
                if tag_row:
                    r = tag_row
                    while r < width_row and isinstance(ws.cell(r, ci).value, str) and norm(ws.cell(r, ci).value).upper().startswith('TAG'):
                        tag_rows.append(r); r += 1
                start = (tag_rows[-1] if tag_rows else nr_row) + 1
                extra_tags = []
                green, blue = [], []
                # lines are detected per card (column); page-level red line is the fallback
                red_c = next((r for r in range(start, width_row + 1) if line_above(ws, r, ci, is_red)), red_row)
                black_c = next((r for r in range(start + 1, (red_c or width_row))
                                if line_above(ws, r, ci, is_black)), None)
                above_black, between, below_red = [], [], []
                for r in range(start, width_row):
                    c = ws.cell(r, ci)
                    if c.value in (None, '') or not norm(c.value):
                        continue
                    if norm(c.value).upper().startswith('TAG'):
                        extra_tags.append(f'{c.coordinate}: {norm(c.value)}'); continue
                    if re.fullmatch(r'[\W_]+', norm(c.value)):
                        issues.append(f'{ws.title} {c.coordinate}: karakter "{norm(c.value)}" di area warna diabaikan'); continue
                    fr = font_rgb(c)
                    item = {
                        'cell': c.coordinate,
                        'text': str(c.value).strip(),
                        'yarns': split_yarns(c.value),
                        'font_color': fr,
                        'meaning': FONT_MEANING.get(fr, 'REPEAT COLOR' if fr in (None, 'FF000000') else f'? ({fr})'),
                        'comment': norm(c.comment.text) if c.comment else None,
                    }
                    (blue if red_c and r >= red_c else green).append(item)
                    if red_c and r >= red_c:
                        below_red.append(item)
                    elif black_c and r < black_c:
                        above_black.append(item)
                    else:
                        between.append(item)
                no_val = wv.cell(num_row, ci).value if num_row else None
                card = {
                    'sheet': ws.title, 'page': pi, 'column': cl,
                    'section': section,
                    'no': (int(no_val) if isinstance(no_val, (int, float))
                           else int(norm(no_val)) if no_val is not None and norm(no_val).isdigit()
                           else norm(no_val) if no_val is not None else None),
                    'no_status': NO_FILL_MEANING.get(fill_rgb(ws.cell(num_row, ci))) if num_row else None,
                    'no_cell': f'{cl}{num_row}' if num_row else None,
                    'no_fill': fill_detail(ws.cell(num_row, ci)) if num_row else None,
                    'style_whs': STYLE_WHS.get(fill_detail(ws.cell(num_row, ci))) if num_row else None,
                    'style_name': norm(name),
                    'new_repeat': norm(ws.cell(nr_row, ci).value or ''),
                    'tag': ' | '.join(norm(ws.cell(r, ci).value).replace('\n\n', ' | ').replace('\n', ' ') for r in tag_rows) or None,
                    'tag_cell': ','.join(f'{cl}{r}' for r in tag_rows) or None,
                    'tag_notes_in_color_area': extra_tags or None,
                    'green_box': green,
                    'blue_box': blue,
                    'black_line_row': black_c,
                    'red_line_row': red_c,
                    'above_black': above_black,
                    'between_black_red': between,
                    'below_red': below_red,
                }
                if card['style_whs'] is None:
                    issues.append(f'{ws.title} {card["no_cell"]} No.{card["no"]} {card["style_name"]}: '
                                  f'background nomor tidak dikenali ({card["no_fill"]}), WHS style tidak bisa ditentukan')
                if green and not black_c:
                    issues.append(f'{ws.title} {cl}{name_row} {card["style_name"]}: ada warna di atas garis merah tapi garis hitam tidak ditemukan')
                if not tag_row:
                    issues.append(f'{ws.title} {cl}{name_row} {card["style_name"]}: tidak ada baris TAG')
                if not green and not blue:
                    issues.append(f'{ws.title} {cl}{name_row} {card["style_name"]}: tidak ada warna')
                cards.append(card)
    return cards, issues


if __name__ == '__main__':
    cards, issues = extract(sys.argv[1])
    json.dump({'source': 'F27 MATRIX MASTER', 'cards': cards, 'issues': issues},
              open(sys.argv[2], 'w'), indent=2, ensure_ascii=False, default=str)
    print('cards', len(cards))
    print(collections.Counter(c['sheet'] for c in cards))
    print('issues', len(issues))
    for i in issues:
        print(' -', i)
