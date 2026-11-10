# matrix2im

Mengubah file Excel **F27 MATRIX MASTER** menjadi JSON untuk **IM MASTER**.

- Masukan: file MATRIX `.xlsx`.
- Keluaran: `raw.json` (hasil baca MATRIX apa adanya, per style) dan `final.json` (satu record per colorway, dengan key kolom IM MASTER).
- Menulis ke IM MASTER (Excel) adalah fase 2 dan **belum dikerjakan**. Tool ini hanya membaca; file input tidak pernah diubah.

## Instalasi

Butuh Python 3.11+.

```bash
python3.12 -m venv .venv
.venv/bin/pip install -e '.[dev]'
```

Pillow sengaja tidak dipasang. Tanpa Pillow, openpyxl melewati 220 foto di MATRIX sehingga pembacaan tetap cepat (`run` pada F27 sekitar 2 detik).

## Perintah

```bash
# Tahap 1: MATRIX -> raw JSON
.venv/bin/matrix2im extract "data/Copy of F27 MATRIX MASTER.xlsx" -o out/raw.json

# Tahap 2: raw JSON -> JSON final (--im-master opsional: cocokkan key dengan header IM baris 54)
.venv/bin/matrix2im transform out/raw.json -o out/final.json --im-master "data/Copy of F27 IM MASTER.xlsx"

# Keduanya sekaligus; menulis raw.json dan final.json ke folder, lalu mencetak ringkasan
.venv/bin/matrix2im run "data/Copy of F27 MATRIX MASTER.xlsx" --im-master "data/Copy of F27 IM MASTER.xlsx" -o out/
.venv/bin/matrix2im run "data/Copy of F27 MATRIX MASTER.xlsx" --im-master "data/Copy of F27 IM MASTER.xlsx" -o out/ --strict
```

- `run` mewajibkan `--im-master`. `--strict` hanya ada di `run`.
- `raw.json` ditulis segera setelah Tahap 1 berhasil. `final.json` hanya ditulis bila validasi lolos (jumlah colorway = jumlah record, key record ada di header IM, WHS hanya tiga nilai, tidak ada nilai tebakan). Bila validasi gagal, tiap error dicatat dan `final.json` tidak ditulis.
- Ringkasan (jumlah style, record, WHS per nilai, WS TAG COLOR terisi/null, needs_review per kode) dicetak ke stdout. Log (progres, jumlah warning, error) ke stderr. Folder output dibuat otomatis.
- JSON ditulis UTF-8, `ensure_ascii=False`, indent 2.

### Exit code

| Kode | Arti |
| --- | --- |
| 0 | Sukses (item `needs_review` tidak membuat gagal) |
| 1 | Error fatal: argumen salah/kurang (mis. `run` tanpa `--im-master`), file tidak terbaca, tidak ada halaman/kartu, raw JSON rusak/tidak lengkap, validasi gagal, error tak terduga. Tidak ada traceback; pesan di stderr |
| 2 | Hanya dengan `--strict` dan `needs_review` tidak kosong (file tetap ditulis) |

## Skema keluaran

**raw.json**: `source`, `total_styles` (jumlah kartu), `styles`, `warnings`. Tiap style: `no`, `style`, `style_whs`, `tag`, `sheet`, `section`, dan tiga daftar colorway menurut garis: `color_above_black_line`, `color_between_black_and_red_line`, `color_below_red_line`. Tiap colorway: `colorway`, `yarns`, `font_color`, `cell`, `comment`.

**final.json**: `target`, `source`, `generated`, `rules`, `summary`, `records`, `needs_review`.

- `records`: satu per colorway, key persis kolom IM: `DESCRIPTION`, `COLOR1`..`COLOR18` (di IM kolom ke-4 bernama `COLOR 4`), `WS TAG COLOR`, `WHS`. `_ref` berisi asal di MATRIX (sheet, No., sel, posisi terhadap garis); tidak ditulis ke IM. Nilai yang tidak bisa diputuskan aturan adalah `null`.
- `needs_review`: item `{record, code, DESCRIPTION, colorway, issue, _ref, ...}`; `record` berupa index record (atau list index, atau `null`).

Kode `needs_review`:

| Kode | Arti |
| --- | --- |
| `TAG_CONDITIONAL` | TAG bersyarat atau singkatan; WS TAG COLOR `null` |
| `TAG_CN_OUT_OF_RANGE` | TAG meminta Cn melebihi jumlah benang |
| `TAG_EMPTY` | TAG kosong |
| `WHS_CAPPED_BY_STYLE` | info: WHS colorway dibatasi WHS style |
| `WHS_STYLE_UNKNOWN` | WHS style tidak diketahui (background No kosong/tak dikenal) |
| `WHS_CONFLICT` | WHS colorway dan style tidak beririsan; WHS `null` |
| `FONT_COLOR_NO_RULE` | warna font (mis. YR / HOLD / DISCONTINUED) belum punya aturan WHS; WHS `null` |
| `COLOR_IS_DESCRIPTOR` | COLOR berupa deskripsi (STRIPE/MARL/MULTI); benang asli ada di comment sel |
| `DUPLICATE_STYLE_ACROSS_SHEETS` | style + colorway sama muncul di dua sheet (`record` = list index; `whs_values` bila WHS beda) |
| `TOO_MANY_YARNS` | benang unik lebih dari 18 slot COLOR |
| `EXTRACTION_WARNING` | warning dari Tahap 1 (`record` = `null`) |

## Konfigurasi

Konstanta yang bisa berubah (warna legend, label, alias nama warna, prefix benang, regex, nama sheet/baris header IM) ada di satu tempat: `matrix_to_im/legend_config.py`. Tidak ada nomor baris/kolom MATRIX yang di-hardcode.

## Test

```bash
.venv/bin/pytest
```

Test regression memakai file asli di `data/` (tidak di-commit) dan di-skip dengan pesan bila file tidak ada. File dicari lewat env var `MATRIX2IM_MATRIX` / `MATRIX2IM_IM`, lalu `data/F27_MATRIX_MASTER.xlsx` / `data/F27_IM_MASTER.xlsx`, lalu pola `data/*F27*MATRIX*MASTER*.xlsx` (nama asli: `Copy of F27 MATRIX MASTER.xlsx`):

```bash
MATRIX2IM_MATRIX=/path/MATRIX.xlsx MATRIX2IM_IM=/path/IM.xlsx .venv/bin/pytest -m regression
```

## Open issues (belum diputuskan)

Kasus berikut sengaja tetap `null` dan/atau masuk `needs_review`; tidak ada aturan yang ditebak:

- WS TAG COLOR untuk TAG bersyarat atau singkatan.
- TAG yang meminta Cn melebihi jumlah benang.
- Style yang muncul di dua sheet (HUTTON di ESSENTIALS GRASSY No. 20 dan LUXE No. 11) dengan WHS berbeda.
- COLOR berupa deskripsi stripe/marl (benang asli ada di comment sel).
- Background nomor kosong (MERRY STRIPED).
- Aturan WHS untuk font YR / HOLD / DISCONTINUED.
- Arti garis hitam (tidak memengaruhi WHS).
