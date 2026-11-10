# Bangun sistem `matrix2im`: MATRIX MASTER (Excel) → JSON untuk IM MASTER

## 1. Tujuan

Bangun tool Python siap produksi bernama `matrix2im` yang membaca file Excel **F27 MATRIX MASTER** dan menghasilkan dua file JSON:

1. **raw JSON**: isi setiap kartu style di MATRIX apa adanya (Tahap 1, Ekstraksi).
2. **final JSON**: satu record per colorway dengan key = nama kolom IM MASTER, ditambah daftar `needs_review` (Tahap 2, Transformasi).

Sistem ini dijalankan setelah automation lain (Image → MATRIX MASTER) selesai memperbarui file MATRIX. **Input: file Excel. Output: JSON.** Penulisan ke file IM MASTER adalah fase berikutnya dan **tidak** dikerjakan sekarang.

## 2. File referensi di repo

Baca semua file ini sebelum menulis kode:

| Path | Isi | Cara memakainya |
| --- | --- | --- |
| `docs/PRD_MATRIX_MASTER_to_IM_JSON.md` | PRD lengkap | **Sumber kebenaran** untuk semua aturan bisnis |
| `reference/prototype/extract_matrix.py`, `export_json.py` | prototipe Tahap 1 | logika deteksi yang sudah terbukti di data F27; boleh ditulis ulang, tapi perilakunya harus sama |
| `reference/prototype/to_im_json.py`, `final_json.py` | prototipe Tahap 2 | idem |
| `reference/golden/F27_MATRIX_MASTER_extract.json` | output raw JSON prototipe untuk F27 | acuan regression test Tahap 1 |
| `reference/golden/F27_IM_MASTER_final.json` | output final JSON prototipe untuk F27 | acuan regression test Tahap 2 |
| `data/F27_MATRIX_MASTER.xlsx` | input contoh (125 MB, berisi 220 foto) | input test; **jangan diubah, jangan di-commit** |
| `data/F27_IM_MASTER.xlsx` | IM MASTER contoh | hanya untuk validasi nama kolom (header baris 54); **jangan diubah, jangan di-commit** |

Kalau PRD dan prototipe berbeda, ikuti PRD dan laporkan perbedaannya ke saya.

## 3. Scope

Dikerjakan:

- CLI + library Python untuk Tahap 1 dan Tahap 2.
- Validasi otomatis dan antrian `needs_review`.
- Unit test dan regression test terhadap file golden.
- README cara instal dan cara pakai.

Tidak dikerjakan:

- Menulis ke file IM MASTER (fase 2). Tapi rancang modulnya supaya penulis IM bisa ditambahkan nanti tanpa mengubah format final JSON.
- Melengkapi prefix/kode benang (YCA/YST/…) di MATRIX. Itu tugas automation lain; nilai COLOR diambil apa adanya.
- Memutuskan open issues di bagian 11. Biarkan tetap ditandai di `needs_review`.
- Mengubah file input apa pun.

## 4. Teknologi dan struktur project

- Python 3.11+, `openpyxl`, `pytest`. Hindari dependensi lain kecuali benar-benar perlu (CLI cukup `argparse`).
- Kemas sebagai package dengan `pyproject.toml` dan entry point `matrix2im`.
- Semua konstanta yang bisa berubah (kode warna legend, teks label, alias tag, daftar prefix benang) disimpan di satu file konfigurasi (`matrix_to_im/legend_config.py` atau YAML), **bukan** tersebar di kode.

Struktur yang disarankan (boleh disesuaikan, jelaskan alasannya):

```
matrix_to_im/
  __init__.py
  legend_config.py     # warna, label, alias, prefix, regex
  excel_reader.py      # load workbook 2x, helper warna fill/font/border
  extract.py           # halaman → kartu → field → raw JSON
  tag_rules.py         # pembentukan WS TAG COLOR
  whs_rules.py         # WHS warna, WHS style, irisan
  transform.py         # raw JSON → records
  validate.py          # cek header IM, invariant, needs_review
  cli.py
tests/
  unit/
  regression/
README.md
pyproject.toml
```

## 5. CLI

```
matrix2im extract   <matrix.xlsx> -o raw.json
matrix2im transform <raw.json> -o final.json [--im-master IM.xlsx]
matrix2im run       <matrix.xlsx> --im-master IM.xlsx -o <folder_output> [--strict]
```

- `run` menjalankan kedua tahap dan menulis `raw.json` + `final.json` ke folder output, lalu mencetak ringkasan (jumlah style, record, WHS, WS TAG COLOR terisi/null, needs_review per jenis).
- `--im-master`: kalau diberikan, semua key record dicocokkan dengan header IM MASTER baris 54. Kalau ada key yang tidak ada di header, proses gagal.
- Exit code: `0` sukses; `1` error fatal (file tidak bisa dibaca, tidak ada halaman/kartu); `2` hanya dengan `--strict` dan `needs_review` tidak kosong. Tanpa `--strict`, item `needs_review` tidak membuat proses gagal.
- Log ke stderr; JSON ditulis UTF-8 dengan `ensure_ascii=False`, indent 2.

## 6. Tahap 1: Ekstraksi (aturan wajib)

Struktur MATRIX dijelaskan di PRD bagian "Input". **Tidak boleh ada nomor baris atau kolom yang di-hardcode**: tinggi halaman (63–71 baris), posisi kolom label (E, I, J, K, M), dan jumlah kartu per halaman (1–10) berbeda-beda.

1. **Halaman**: blok yang dimulai sel kolom A berawalan `FALL 27` sampai sebelum blok berikutnya.
2. **Kolom label**: kolom di halaman itu yang berisi `- LEBAR/WIDTH`. Kartu style = setiap kolom di kirinya yang punya nama style.
3. **Baris NEW/REPEAT**: baris yang selnya **cocok persis** (full match, setelah trim dan uppercase) dengan `NEW`, `EXACT REPEAT`, `REPEAT`, `REPEAT WITH UPDATE`, `UPDATE`, atau `SIMILAR`. Nama style = sel tepat di atasnya.
4. **Nomor (`no`)**: baris pertama di bawah legend yang punya nilai cocok `^(\d+(\.0)?[A-Z]?|S\d+( - [A-Z])?)$`. Baca dari workbook `data_only=True` karena nomor berupa formula (`=A9+1`). Nilai teks berisi angka saja dijadikan integer. Halaman dengan 1 kartu tetap harus terbaca.
5. **Sub-koleksi (`section`)**: teks kolom A tepat di atas baris nomor (mis. `WINTER BEACH`, `CHRISTMAS - GRASSY`, `SKI & SNOW`); abaikan kalau teks itu bagian legend (mengandung `COLOR`, `STYLE`, atau `ORDER`); null kalau tidak ada. Nomor dimulai ulang per sub-koleksi, jadi kunci unik kartu = sheet + section + no.
6. **WHS style (`style_whs`)** dari warna background sel nomor:
   - `FFFF66FF` (ungu) → `"000"`
   - abu → `"000 & SY"`. Abu tersimpan dalam dua bentuk: theme 0 dengan tint ≈ −0.15 (bandingkan setelah dibulatkan 2 desimal) dan `FFD9D9D9`.
   - `FFFF0000` (merah) → `"SY"`
   - lainnya atau tanpa fill → `null` + warning
7. **TAG**: per kartu, sel pertama di bawah baris NEW/REPEAT yang berawalan `TAG`. Kalau sel-sel di bawahnya juga berawalan `TAG`, gabungkan dengan ` | `. Sel berawalan `TAG` yang muncul terpisah di area warna tidak dihitung sebagai colorway dan tidak masuk raw JSON (di F27 tidak ada kasus ini).
8. **Garis merah**: border medium/thick/double berwarna `FFFF0000`, `FFE40500`, atau `FFC00000`, baik sebagai border atas sel maupun border bawah sel di atasnya. Deteksi **per kartu**; kalau kartu itu tidak punya garis merah, pakai baris garis merah mayoritas di halaman.
9. **Garis hitam**: border medium/thick/double yang warnanya `None`, auto, indexed 8 atau 64, theme 1, atau `FF000000`. Deteksi **per kartu**, hanya di atas garis merah. Di halaman EXACT REPEAT garis merah ada di baris warna pertama dan tidak ada garis hitam, sehingga semua warna masuk kelompok "di bawah garis merah".
10. **Colorway**: setiap sel tidak kosong setelah baris TAG terakhir sampai sebelum baris `- LEBAR/WIDTH`. Sel yang hanya berisi tanda baca (`.`, `` ` ``, `*`) diabaikan + warning. Kelompokkan:
    - `color_above_black_line`: baris < garis hitam
    - `color_between_black_and_red_line`: garis hitam ≤ baris < garis merah
    - `color_below_red_line`: baris ≥ garis merah
11. **Per colorway** simpan: `colorway` (yarns digabung `/`), `yarns` (teks dipecah dengan `/` dan baris baru, di-trim, urutan dipertahankan), `font_color` (arti dari legend, lihat tabel di bawah), `cell`, `comment` (isi comment sel; sering berisi override tag atau daftar benang stripe).

Arti warna font colorway (warna font berupa rgb; selain rgb dianggap hitam):

| Warna font | `font_color` |
| --- | --- |
| `FF00B050` | `NEW COLOR` |
| tanpa warna / `FF000000` / theme / indexed | `REPEAT COLOR` |
| `FFFF00FF` | `000 ONLY COLOR` |
| `FFFF99FF` | `000 REPEAT COLOR` |
| `FF9900FF` | `YR COLOR` |
| `FFFF0000` | `HOLD COLOR FOR SY` |
| `FF0066FF` | `DISCONTINUED COLOR FOR SY` |
| lainnya | `? (<kode>)` + warning (di F27 tidak ada) |

Workbook dibuka dua kali: sekali normal (formula, warna, border, comment) dan sekali `data_only=True` (nilai nomor). Buka read-only terhadap file sumber; jangan pernah memanggil `save()` pada workbook input.

**Skema raw JSON** (key dan urutan harus sama dengan golden):

```json
{
  "source": "F27 MATRIX MASTER",
  "total_styles": 215,
  "styles": [
    {
      "no": 20,
      "style": "HUTTON MOCK T NECK CHUNKY",
      "style_whs": "000 & SY",
      "tag": "TAG: C1 (For Solid) | TAG: SNOW MARL (For PURE SNOW/ALMOND BUTTER)",
      "color_above_black_line": [{"colorway": "NIGHT NAVY", "yarns": ["NIGHT NAVY"], "font_color": "NEW COLOR", "cell": "C163", "comment": null}],
      "color_between_black_and_red_line": [],
      "color_below_red_line": [],
      "sheet": "ESSENTIALS GRASSY",
      "section": null
    }
  ],
  "warnings": ["..."]
}
```

Urutan style: urutan sheet → halaman → kolom kartu dari kiri ke kanan. Urutan colorway: dari atas ke bawah.

## 7. Tahap 2: Transformasi (aturan wajib)

Setiap colorway menjadi satu record, dengan urutan: style sesuai raw JSON, lalu kelompok above_black → between → below_red, dari atas ke bawah.

**DESCRIPTION** = `style` apa adanya (tanpa size).

**COLOR1 … COLOR18** = 1 benang = 1 kolom, nilai persis seperti di `yarns`.
- Key harus sama persis dengan header IM: `COLOR1`, `COLOR2`, `COLOR3`, **`COLOR 4` (pakai spasi)**, `COLOR5` … `COLOR18`.
- Benang yang berulang (stripe) ditulis sekali, case-insensitive, sesuai urutan kemunculan pertama.
- Slot tanpa benang **tidak dicantumkan** di record.
- Lebih dari 18 benang unik → tulis 18 pertama + review.

**WS TAG COLOR**:
1. TAG mengandung ` | `, `FOR SOLID`, `FOR HEATHER`, `FOR MARL`, atau `ABBREVIAT` → `null` + review (`TAG_CONDITIONAL`).
2. Buang awalan `^TAG\s*:?\s*`.
3. Kalau ada `Cn` dengan n > jumlah benang unik → `null` + review (`TAG_CN_OUT_OF_RANGE`).
4. Ganti setiap `\bC(\d+)\b` dengan nama warna benang ke-n. Nama warna = uppercase, buang prefix `^(YCA|YST|YTX|YP|BJ|POODLE)\s+`, buang kode di akhir `\s*-\s*\d+\s*[A-Z]?$`, lalu terapkan alias (`BLACK B` → `BLACK`).
5. `STRIPED` → `STRIPE` (catatan di header IM: "STRIPE tanpa D").
6. Akhiran dobel dipadatkan: `(STRIPE|MARL|MULTI|HEATHERED)( \1)+` → `\1`. Contoh: `BLUSH CHAMPAGNE STRIPE STRIPE` → `BLUSH CHAMPAGNE STRIPE`.
7. Rapikan spasi di sekitar `/`.

Contoh: TAG `C1/C2 MARL` + benang `YCA KHAKI - 749`, `YCA BLUSH CHAMPAGNE - 645` → `KHAKI/BLUSH CHAMPAGNE MARL`.

**WHS** (key `WHS`, ditujukan ke kolom CE di IM; IM juga punya kolom G bernama "WHS", tapi itu bukan target):

| Posisi colorway | `font_color` | WHS warna |
| --- | --- | --- |
| below_red | apa pun yang dikenal | `SY` |
| above_black / between | `000 ONLY COLOR`, `000 REPEAT COLOR` | `000` |
| above_black / between | `NEW COLOR`, `REPEAT COLOR` | `000 & SY` |
| apa pun | `YR COLOR`, `HOLD COLOR FOR SY`, `DISCONTINUED COLOR FOR SY`, tidak dikenal | `null` + review |

WHS final = **irisan** himpunan WHS style dan WHS warna (`000` = {000}, `SY` = {SY}, `000 & SY` = {000, SY}). WHS style menjadi batas atas. Contoh: style `000` + warna `000 & SY` → `000`. Irisan kosong → `null` + review. `style_whs` null → pakai WHS warna + review. Nilai WHS hanya boleh `"000"`, `"SY"`, atau `"000 & SY"`.

**`_ref`** = `"<sheet>[ (<section>)] No.<no> | <cell> | <above_black|between_black_red|below_red>"`. Key berawalan `_` adalah metadata dan tidak akan ditulis ke IM. Nilai `null` berarti "jangan ditulis".

**Skema final JSON:**

```json
{
  "target": "F27 IM MASTER (sheet F27, header baris 54)",
  "source": "F27 MATRIX MASTER",
  "generated": "YYYY-MM-DD",
  "rules": {"DESCRIPTION": "...", "COLOR1-COLOR18": "...", "WS TAG COLOR": "...", "WHS": "...", "_ref": "..."},
  "summary": {"total_records": 724, "total_styles": 211, "WHS": {}, "WS_TAG_COLOR_filled": 646, "WS_TAG_COLOR_null": 78, "needs_review_items": 115, "needs_review_by_issue": {}},
  "records": [
    {"DESCRIPTION": "JUDE T NECK CHUNKY", "COLOR1": "SIENNA", "WS TAG COLOR": "SIENNA", "WHS": "000", "_ref": "ESSENTIALS GRASSY No.21 | D166 | between_black_red"}
  ],
  "needs_review": [
    {"record": 61, "code": "TAG_CONDITIONAL", "DESCRIPTION": "ARIA V CHUNKY COTTON", "colorway": "...", "issue": "teks penjelasan dalam Bahasa Indonesia", "tag_instruction": "...", "_ref": "..."}
  ]
}
```

Perbaikan dari prototipe: tambahkan field **`code`** yang bisa dibaca mesin di setiap item `needs_review`, di samping teks `issue`. Kode yang dipakai: `TAG_CONDITIONAL`, `TAG_CN_OUT_OF_RANGE`, `TAG_EMPTY`, `WHS_CAPPED_BY_STYLE` (info), `WHS_STYLE_UNKNOWN`, `WHS_CONFLICT`, `FONT_COLOR_NO_RULE`, `COLOR_IS_DESCRIPTOR` (benang berakhiran STRIPE/STRIPED/MARL/MULTI; sertakan comment sel MATRIX), `DUPLICATE_STYLE_ACROSS_SHEETS` (field `record` berisi list index; sebutkan kalau WHS berbeda), `TOO_MANY_YARNS`, `EXTRACTION_WARNING` (`record` = null).

## 8. Validasi otomatis (setiap kali dijalankan)

- Jumlah colorway di raw JSON = jumlah record di final JSON.
- Semua key record (selain yang berawalan `_`) ada di header IM MASTER baris 54. Header kolom WS TAG COLOR di IM terdiri dari dua baris (`WS TAG COLOR\n(Note - …)`), jadi cocokkan dengan baris pertamanya saja.
- WHS hanya berisi tiga nilai di atas.
- Tidak ada nilai tebakan: semua yang tidak bisa diputuskan aturan bernilai `null` dan punya item di `needs_review`.

## 9. Test dan kriteria penerimaan

**Regression (wajib lulus) pada `data/F27_MATRIX_MASTER.xlsx`:**

- raw JSON: array `styles` identik dengan `reference/golden/F27_MATRIX_MASTER_extract.json`. `warnings` boleh berbeda kalimatnya, tapi jumlahnya harus 4 dan menunjuk sel/kartu yang sama.
- final JSON: array `records` identik dengan `reference/golden/F27_IM_MASTER_final.json`. Untuk `needs_review`, setiap item golden harus punya padanan dengan `record` yang sama dan jenis masalah yang sama; teks `issue` dan `rules` boleh berbeda.
- Angka baseline:

| Metrik | Nilai F27 |
| --- | --- |
| Kartu style | 215 (211 nama unik) |
| Colorway | 724: 158 above_black, 390 between, 176 below_red |
| `style_whs` | 85 `000 & SY`, 70 `000`, 59 `SY`, 1 null |
| WHS final | 354 `000`, 194 `000 & SY`, 176 `SY` |
| WS TAG COLOR | 646 terisi, 78 null |
| `needs_review` | 115 item: 57 TAG_CONDITIONAL, 21 TAG_CN_OUT_OF_RANGE, 17 WHS_CAPPED_BY_STYLE, 8 DUPLICATE_STYLE_ACROSS_SHEETS (2 di antaranya WHS berbeda), 6 COLOR_IS_DESCRIPTOR, 2 WHS_STYLE_UNKNOWN, 4 EXTRACTION_WARNING |

Kalau file `data/` tidak ada, test regression di-skip dengan pesan yang jelas (jangan gagal).

**Kasus emas (unit/regression test terpisah):**

- ESSENTIALS GRASSY No. 20 HUTTON MOCK T NECK CHUNKY: above_black = NIGHT NAVY; between = PURE SNOW/ALMOND BUTTER, GLASGOW, FAWN, RED CHILE; below_red = BLACK, DARK ROAST, SCANDIA BLUE; `style_whs` = `000 & SY`.
- ESSENTIALS GRASSY No. 21 JUDE T NECK CHUNKY: above_black = DARK CEDAR; between = ALMOND BUTTER, SLEET, SIENNA, BOLD INK; below_red kosong; TAG `C1`.
- Halaman 1 ESSENTIALS COTTON (ada style bernama "NEWPORT …") harus menghasilkan 8 kartu, dimulai dari PEBBLE STITCH RAGLAN SWEATSHIRT COTTON.
- FALL halaman 5 (10 kartu): garis merah ada di baris warna pertama, jadi ke-12 colorway-nya masuk below_red dan tidak ada yang masuk above_black/between.
- LUXE halaman 7 hanya berisi 1 kartu (No. `S1`, CHRISTMAS TREE CREW COTTON) dan nomornya tetap terbaca. Nomor teks seperti `S2 - A` juga terbaca.
- LUXE, CHRISTMAS - GRASSY, No. `S3` FA LA LA FAIR ISLE CREW CHUNKY: TAG dua baris menjadi `TAG: BLACK/LATTE/ RED | TAG: RED PAPRIKA/SHADLE/ KALE`.
- ESSENTIALS COTTON No. 31 SLOANE TOP COTTON: WHS warna `000 & SY` diturunkan menjadi `000` oleh style `000`.
- ESSENTIALS COTTON, WINTER BEACH, No. 1 RIHANNA STRIPED CARDI CHUNKY COTTON: WS TAG COLOR `BLUSH CHAMPAGNE STRIPE` (tanpa akhiran dobel).
- LUXE, CHRISTMAS - COTTON, No. 4 MERRY STRIPED CREW COTTON: `style_whs` null + warning.

**Unit test** untuk fungsi murni: pemecahan yarns, nama warna dari kode benang, pembentukan WS TAG COLOR (semua cabang), WHS warna, irisan WHS, deteksi warna (rgb, theme+tint, indexed, auto), dan regex label NEW/REPEAT.

## 10. Kebutuhan non-fungsional

- Deterministik: input sama → output sama (selain `generated`).
- Target waktu ≤ 15 detik untuk file F27 (prototipe: sekitar 7 detik).
- File input tidak pernah dimodifikasi.
- Pesan error dan isi `issue` dalam Bahasa Indonesia; nama key dan kode dalam bentuk yang ditetapkan di atas.
- Kode diberi type hints dan docstring singkat yang menjelaskan *kenapa* sebuah aturan ada (banyak aturan berasal dari keanehan data MATRIX).

## 11. Open issues: jangan diputuskan sendiri

Semua kasus ini harus tetap menghasilkan `null` dan/atau item `needs_review`. Jangan menambah aturan baru untuk kasus-kasus ini tanpa persetujuan saya:

- WS TAG COLOR untuk TAG bersyarat atau singkatan.
- TAG yang meminta Cn melebihi jumlah benang.
- Style yang muncul di dua sheet (HUTTON di ESSENTIALS GRASSY No. 20 dan LUXE No. 11) dengan WHS berbeda.
- COLOR berupa deskripsi stripe/marl (benang aslinya ada di comment sel).
- Background nomor kosong (MERRY STRIPED).
- Aturan WHS untuk font YR / HOLD / DISCONTINUED.
- Arti garis hitam (tidak memengaruhi WHS).

## 12. Cara kerja

1. Baca PRD, prototipe, dan contoh golden. Lalu kirim ke saya rencana singkat (struktur modul, urutan pengerjaan, dan perbedaan apa pun yang kamu temukan antara PRD dan prototipe) **sebelum** mulai menulis kode.
2. Kerjakan bertahap: Tahap 1 + test-nya sampai cocok dengan golden, baru Tahap 2 + test-nya, lalu CLI dan README.
3. Jalankan seluruh test setelah setiap tahap. Jangan menandai pekerjaan selesai selama ada test yang gagal.
4. Kalau ada aturan yang ambigu atau data yang tidak sesuai PRD, berhenti dan tanyakan. Jangan menebak.
5. Jangan commit file di `data/` atau file output; tambahkan ke `.gitignore`.
6. Di akhir, laporkan: hasil test, tabel ringkasan F27 dibanding baseline, cara menjalankan, dan hal yang masih terbuka.
