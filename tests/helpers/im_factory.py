"""IM MASTER sintetis untuk test penulisan: sheet F27 + elemen fileSharing (Password to modify)."""
from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import quoteattr

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from matrix_to_im.legend_config import IM_COLOR_KEYS, IM_HEADER_ROW, IM_SHEET

TEST_PASSWORD = "rahasia-uji"
TEST_SALT_B64 = "AAECAwQFBgcICQoLDA0ODw=="
TEST_HASH_SPIN1000 = "PyNhGFHLSUU4sQ1nbaw7z3yfxE0RsH4rbX3I46rm+XP4xYVebx88dqEbZ8rPlx/BMzftVlY7uScXWiCAscS58Q=="
TEST_HASH_SPIN100000 = "hzY8X1+Zw+dryKltyS4V+shqaVXxQ9h3IGLqGwSL3/G1Et9T0OIbktcJNfqEkIQzY2rr+r5nPM3mF7c9GkPQig=="

# Kolom A..Z; WHS di C adalah pengecoh (kolom G di IM asli), WHS sungguhan di Y.
_HEADERS = ["DESCRIPTION", "PCS", "WHS", "C1Y%", "COLOR1", "C2Y%", "COLOR2", "COLOR3", "COLOR 4",
            *[f"COLOR{i}" for i in range(5, 19)], "WS TAG COLOR", "WHS", "NOTE"]
_COLOR_COLS = [5, 7, 8, 9, *range(10, 24)]      # E, G, H, I, J..W untuk IM_COLOR_KEYS
_WS_TAG_COL, _WHS_COL, _NOTE_COL = 24, 25, 26
assert len(_COLOR_COLS) == len(IM_COLOR_KEYS)


@dataclass
class GroupSpec:
    description: str
    colorways: list[list[str]]
    subtotal: str | None = "BOTH WHS"
    subtotal_label: str | None = None      # label baris subtotal apa adanya (menggantikan "<desc> - <subtotal>")
    row_descriptions: list[str | None] | None = None   # DESCRIPTION per baris colorway (anotasi); None = description
    subtotal_func: str = "SUBTOTAL"        # nama fungsi di rumus subtotal (mis. huruf kecil)
    ws_tags: list[str | None] | None = None
    whs: list[str | None] | None = None
    ws_tag_fills: list[str | None] | None = None   # warna latar WS TAG COLOR (6 hex) per baris


def build_im_sheet(path: Path, groups: list[GroupSpec], *, extra_header: dict[str, str] | None = None) -> Path:
    """Tulis IM sintetis; extra_header memetakan huruf kolom ke teks header pengganti."""
    wb = Workbook()
    ws = wb.active
    ws.title = IM_SHEET
    for col, text in enumerate(_HEADERS, 1):
        ws.cell(IM_HEADER_ROW, col, text)
    for letter, text in (extra_header or {}).items():
        ws[f"{letter}{IM_HEADER_ROW}"] = text
    yellow = PatternFill("solid", fgColor="FFFFFF00")
    row = IM_HEADER_ROW + 1
    for g in groups:
        first = row
        for i, colors in enumerate(g.colorways):
            ws.cell(row, 1, (g.row_descriptions[i] if g.row_descriptions else None) or g.description)
            ws.cell(row, 2, 10)
            ws.cell(row, 4, "=1/2")
            ws.cell(row, 6, "=1/2")
            for col, value in zip(_COLOR_COLS, colors):
                ws.cell(row, col, value)
            ws.cell(row, _COLOR_COLS[0]).fill = yellow
            if g.ws_tags and g.ws_tags[i] is not None:
                ws.cell(row, _WS_TAG_COL, g.ws_tags[i])
            if g.whs and g.whs[i] is not None:
                ws.cell(row, _WHS_COL, g.whs[i])
            if g.ws_tag_fills and g.ws_tag_fills[i] is not None:
                ws.cell(row, _WS_TAG_COL).fill = PatternFill("solid", fgColor="FF" + g.ws_tag_fills[i])
            ws.cell(row, _NOTE_COL, "catatan")
            row += 1
        if g.subtotal is not None:
            ws.cell(row, 1, g.subtotal_label or f"{g.description} - {g.subtotal}")
            last = row - 1
            rng = f"B{first}" if first == last else f"B{first}:B{last}"
            ws.cell(row, 2, f"={g.subtotal_func}(9,{rng})")
            row += 1
    wb.save(path)
    return path


def add_file_sharing(path: Path, hash_b64: str = TEST_HASH_SPIN1000, salt_b64: str = TEST_SALT_B64,
                     spin: int = 1000, user: str = "Tester") -> Path:
    """Sisipkan fileSharing ke xl/workbook.xml (sesudah fileVersion, atau anak pertama workbook)."""
    tag = (f'<fileSharing readOnlyRecommended="1" userName={quoteattr(user)} algorithmName="SHA-512" '
           f'hashValue={quoteattr(hash_b64)} saltValue={quoteattr(salt_b64)} spinCount="{spin}"/>')
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/workbook.xml":
                xml = data.decode("utf-8")
                end = xml.find("/>", xml.find("<fileVersion")) + 2 if "<fileVersion" in xml else -1
                if end > 1:
                    xml = xml[:end] + tag + xml[end:]
                else:
                    start = xml.find("<workbook")
                    gt = xml.find(">", start) + 1
                    xml = xml[:gt] + tag + xml[gt:]
                data = xml.encode("utf-8")
            dst.writestr(item, data)
    tmp.replace(path)
    return path
