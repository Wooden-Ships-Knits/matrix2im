"""Satu-satunya tempat konstanta yang bisa berubah (warna legend, label, alias, regex)."""

PAGE_MARKER = "FALL 27"
WIDTH_LABEL = "- LEBAR/WIDTH"
NEW_REPEAT_LABELS = ("NEW", "EXACT REPEAT", "REPEAT", "REPEAT WITH UPDATE", "UPDATE", "SIMILAR")
NUMBER_RE = r"^(\d+(\.0)?[A-Z]?|S\d+( - [A-Z])?)$"
TAG_PREFIX = "TAG"
TAG_PREFIX_RE = r"^TAG\s*:?\s*"            # dibuang dari instruksi TAG sebelum substitusi Cn
TAG_CN_RE = r"\bC(\d+)\b"                  # Cn = benang ke-n colorway
TAG_REPLACEMENTS = ((r"\bSTRIPED\b", "STRIPE"),)   # catatan header IM: "STRIPE tanpa D"
YARN_SPLIT_RE = r"[/\n]"                   # pemisah benang di sel warna (D2)
PUNCTUATION_ONLY_RE = r"[\W_]+"            # sel tanpa huruf/angka bukan data (D5)
SECTION_EXCLUDE_WORDS = ("COLOR", "STYLE", "ORDER")
LEGEND_LABELS = ("000 ONLY STYLE", "SY ONLY", "NEW COLOR", "REPEAT COLOR", "000 ONLY COLOR",
                 "000 REPEAT COLOR", "YR COLOR", "HOLD COLOR FOR SY", "DISCONTINUED COLOR FOR SY")
# Kode rgb dibandingkan tanpa byte alpha (excel_reader.rgb6): openpyxl menulis "00", Excel "FF".
# Putih / tanpa fill = abu = 000 & SY (keputusan review 10 Nov 2026, kasus MERRY STRIPED).
STYLE_WHS_BY_FILL = {"FFFF66FF": "000", "theme0-0.15": "000 & SY", "FFD9D9D9": "000 & SY", "FFFF0000": "SY",
                     "FFFFFFFF": "000 & SY", "theme0+0.00": "000 & SY"}
STYLE_WHS_NO_FILL = "000 & SY"             # sel nomor berisi tapi tanpa fill (tampil putih)
FONT_MEANING = {"FF00B050": "NEW COLOR", "FFFF00FF": "000 ONLY COLOR", "FFFF99FF": "000 REPEAT COLOR",
                "FF9900FF": "YR COLOR", "FFFF0000": "HOLD COLOR FOR SY", "FF0066FF": "DISCONTINUED COLOR FOR SY"}
DEFAULT_FONT_MEANING = "REPEAT COLOR"      # tanpa warna / FF000000 / theme / indexed
BLACK_FONT_RGB = frozenset({"FF000000"})
THICK_BORDER_STYLES = frozenset({"medium", "thick", "double"})
RED_LINE_RGB = frozenset({"FFFF0000", "FFE40500", "FFC00000"})
BLACK_LINE_RGB = frozenset({"FF000000", "00000000"})
BLACK_LINE_INDEXED = frozenset({8, 64})
BLACK_LINE_THEME = frozenset({1})
YARN_PREFIX_RE = r"^(YCA|YST|YTX|YP|BJ|POODLE)\s+"
YARN_CODE_SUFFIX_RE = r"\s*-\s*\d+\s*[A-Z]?$"
TAG_ALIAS = {"BLACK B": "BLACK"}
TAG_CONDITIONAL_MARKERS = ("|", "FOR SOLID", "FOR HEATHER", "FOR MARL", "ABBREVIAT")
TAG_DOUBLE_SUFFIX_RE = r"\b(STRIPE|MARL|MULTI|HEATHERED)( \1\b)+"
# Instruksi TAG yang belum diterjemahkan: Cn tanpa nomor ("C(car)") atau kata COLOR ("GONDOLA color").
TAG_UNRESOLVED_RE = r"\bC\(|\bCOLOU?R\b"
# Teks WS TAG COLOR bila aturan tidak bisa memutuskan; sengaja ditulis ke IM agar terlihat tim.
NEED_REVIEW = "Need Review"
DESCRIPTOR_RE = r"\b(STRIPE|STRIPED|MARL|MULTI)$"
IM_COLOR_KEYS = ("COLOR1", "COLOR2", "COLOR3", "COLOR 4") + tuple(f"COLOR{i}" for i in range(5, 19))
WHS_VALUES = ("000", "SY", "000 & SY")
IM_SHEET = "F27"
IM_HEADER_ROW = 54
IM_DESCRIPTION_KEY = "DESCRIPTION"
IM_WS_TAG_KEY = "WS TAG COLOR"
IM_WHS_KEY = "WHS"                      # yang ditulis: kolom tepat di kanan WS TAG COLOR (PRD aturan 6)
IM_SUBTOTAL_RE = r" - (BOTH WHS|000 ONLY|SHOPIFY ONLY)\b"
SOURCE_NAME = "F27 MATRIX MASTER"
TARGET_NAME = "F27 IM MASTER (sheet F27, header baris 54)"
# Arti font -> WHS warna di atas garis merah; di bawah garis merah selalu "SY".
FONT_WHS_000 = frozenset({"000 ONLY COLOR", "000 REPEAT COLOR"})
FONT_WHS_000_SY = frozenset({"NEW COLOR", "REPEAT COLOR"})
WHS_SETS = {"000": frozenset({"000"}), "SY": frozenset({"SY"}), "000 & SY": frozenset({"000", "SY"})}
# Kunci IM MASTER (Password to modify, ECMA-376 fileSharing).
IM_LOCK_ALGORITHMS = {"SHA-512": "sha512", "SHA-384": "sha384", "SHA-256": "sha256", "SHA-1": "sha1"}
OWNER_FILE_PREFIX = "~$"                   # file kunci Excel di folder yang sama dengan workbook
OWNER_FILE_TRUNCATE = 2                    # Excel kadang memotong 2 huruf pertama nama pada file pemilik
OWNER_NAME_MAX_LEN = 53                    # byte pertama file pemilik = panjang nama (maks 53)
OWNER_NAME_FALLBACK = "seseorang"
