"""Export MATRIX MASTER to the JSON format agreed with the user (incl. style_whs):
NO, STYLE, TAG, COLOR above black line, COLOR between black & red line, COLOR below red line."""
import json, sys, collections
from extract_matrix import extract


def color(x):
    return {
        'colorway': '/'.join(x['yarns']),     # e.g. "PURE SNOW/ALMOND BUTTER"
        'yarns': x['yarns'],                   # C1, C2, ... in order as written
        'font_color': x['meaning'],            # legend meaning of the font colour
        'cell': x['cell'],
        'comment': x['comment'],
    }


def main(src, dst):
    cards, issues = extract(src)
    out = []
    for c in cards:
        out.append({
            'no': c['no'],
            'style': c['style_name'],
            'style_whs': c['style_whs'],          # background of the No cell: purple 000 | grey 000 & SY | red SY
            'tag': c['tag'],
            'color_above_black_line': [color(x) for x in c['above_black']],
            'color_between_black_and_red_line': [color(x) for x in c['between_black_red']],
            'color_below_red_line': [color(x) for x in c['below_red']],
            'sheet': c['sheet'],
            'section': c['section'],
        })
    json.dump({'source': 'F27 MATRIX MASTER', 'total_styles': len(out), 'styles': out, 'warnings': issues},
              open(dst, 'w'), indent=2, ensure_ascii=False)
    return out, issues


if __name__ == '__main__':
    out, issues = main(sys.argv[1], sys.argv[2])
    a = sum(len(s['color_above_black_line']) for s in out)
    b = sum(len(s['color_between_black_and_red_line']) for s in out)
    r = sum(len(s['color_below_red_line']) for s in out)
    print('styles', len(out), '| above black', a, '| between', b, '| below red', r, '| total', a + b + r)
    print('warnings', len(issues))
    for i in issues:
        print(' -', i)
