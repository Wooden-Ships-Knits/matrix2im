"""Aturan WHS (gudang): dari posisi warna terhadap garis merah + warna font, dibatasi WHS style."""
from .legend_config import FONT_WHS_000, FONT_WHS_000_SY, WHS_SETS
from .review import Code, Problem


def color_whs(group: str, font_color: str) -> tuple[str | None, Problem | None]:
    """WHS warna. Di bawah garis merah selalu SY; di atas, arti font memilih 000 atau 000 & SY.

    Font tanpa aturan (YR, HOLD, DISCONTINUED, tidak dikenal) tidak ditebak: None + review.
    """
    if font_color not in FONT_WHS_000 | FONT_WHS_000_SY:
        return None, Problem(Code.FONT_COLOR_NO_RULE, f'warna font "{font_color}" belum punya aturan WHS')
    if group == "below_red":
        return "SY", None
    return ("000" if font_color in FONT_WHS_000 else "000 & SY"), None


def final_whs(style_whs: str | None, col_whs: str | None) -> tuple[str | None, Problem | None]:
    """WHS style adalah batas atas: hasil = irisan himpunan gudang style dan warna.

    col_whs None berarti masalah font sudah dilaporkan oleh color_whs.
    """
    if col_whs is None:
        return None, None
    if style_whs is None:
        return col_whs, Problem(Code.WHS_STYLE_UNKNOWN,
                                "WHS style tidak diketahui (background nomor kosong), dipakai WHS warna")
    inter = WHS_SETS[style_whs] & WHS_SETS[col_whs]
    if not inter:
        return None, Problem(Code.WHS_CONFLICT, f"WHS style {style_whs} bertentangan dengan WHS warna {col_whs}")
    result = "000 & SY" if inter == WHS_SETS["000 & SY"] else next(iter(inter))
    if result != col_whs:
        return result, Problem(Code.WHS_CAPPED_BY_STYLE,
                               f"info: WHS warna {col_whs} diturunkan ke {result} (batas WHS style)")
    return result, None
