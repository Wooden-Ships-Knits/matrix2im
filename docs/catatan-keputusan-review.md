# Catatan keputusan review `needs_review` (untuk update PRD berikutnya)

> **Status:** semua keputusan di bawah sudah dipindahkan ke PRD (`docs/PRD_MATRIX_MASTER_to_IM_JSON.md`, bagian Riwayat perubahan) dan PROMPT pada 10 November 2026.

Catatan sementara hasil review tab *Needs review* di UI. Setelah semua kasus diputuskan, isinya dipindahkan ke `docs/PRD_MATRIX_MASTER_to_IM_JSON.md`, lalu aturan, test, dan baseline diperbarui.

## 1. DUPLICATE_STYLE_ACROSS_SHEETS — diputuskan 2026-11-10

**Keputusan:** satu style memang bisa muncul di dua koleksi (sheet), mis. HUTTON MOCK T NECK CHUNKY di ESSENTIALS GRASSY No. 20 dan LUXE No. 11. Ini kondisi normal, bukan masalah. Record dari tiap sheet dicatat sebagai nilai terpisah.

**Dampak yang direncanakan (belum dikerjakan):**
- Records: tidak berubah. Record tiap sheet sudah terpisah, masing-masing dengan `_ref` dan WHS sendiri.
- `needs_review`: kode `DUPLICATE_STYLE_ACROSS_SHEETS` tidak dipakai lagi. Pada F27, 8 item hilang (115 → 107), termasuk 2 item dengan WHS berbeda (GLASGOW, FAWN).
- PRD: open issue "Style yang muncul di dua sheet" ditutup.
- Perlu dipastikan di fase 2 (penulis IM): bagaimana record dari dua sheet untuk DESCRIPTION yang sama ditulis ke baris IM.

## 2. WHS_CAPPED_BY_STYLE — diputuskan 2026-11-10

**Keputusan:** jelas dan bukan masalah. WHS mengikuti batas dari background nomor style (irisan WHS style dan WHS warna), jadi tidak perlu dicek tim.

**Dampak yang direncanakan (belum dikerjakan):**
- Records: tidak berubah. Nilai WHS sudah diturunkan, mis. SLOANE TOP COTTON `000 & SY` → `000`.
- `needs_review`: kode `WHS_CAPPED_BY_STYLE` tidak dipakai lagi. Pada F27, 17 item hilang: SLOANE TOP COTTON (5), PURL-KNIT GAME DAY CREW CHUNKY COTTON (8), JOSEF MELANGE CREW CHUNKY (2), TARA RAILS MARLED T NECK CHUNKY (2).
- PRD: baris "WHS warna diturunkan oleh batas WHS style — hanya informasi" di tabel needs_review dihapus.

## 3. TAG_CONDITIONAL — diputuskan 2026-11-10

**Keputusan:** tetap masuk `needs_review`, tapi `WS TAG COLOR` diisi teks **`Need Review`** (bukan `null`), supaya terlihat langsung di IM bahwa kolom itu perlu dicek.

**Dampak yang direncanakan (belum dikerjakan):**
- Records: 57 record berubah, `WS TAG COLOR` dari `null` menjadi `"Need Review"`. Karena record golden ikut berubah, golden perlu disetujui ulang.
- `needs_review`: tetap 57 item `TAG_CONDITIONAL`, tidak berubah.
- Teks `Need Review` disimpan sebagai konstanta di `legend_config.py`, persis seperti yang diminta (huruf besar-kecil sama).
- Validasi: aturan "nilai yang tidak bisa diputuskan harus punya item review" diperluas. Record berisi `Need Review` wajib punya item `TAG_CONDITIONAL`.
- Summary (disetujui): hitungan `Need Review` dipisah sebagai `WS_TAG_COLOR_need_review`, di samping `WS_TAG_COLOR_filled` dan `WS_TAG_COLOR_null`. Angka F27 setelah keputusan 3 dan 4: lihat bagian 4.
- Fase 2: teks `Need Review` akan ikut ditulis ke kolom WS TAG COLOR (CD) di IM.
- `TAG_CN_OUT_OF_RANGE`: lihat bagian 4. `TAG_EMPTY`: belum diputuskan (0 kasus di F27).

## 4. TAG_CN_OUT_OF_RANGE — diputuskan 2026-11-10

**Keputusan:** sama dengan TAG_CONDITIONAL. Tetap masuk `needs_review`, dan `WS TAG COLOR` diisi `Need Review`.

**Dampak yang direncanakan (belum dikerjakan):**
- Records: 21 record berubah, `WS TAG COLOR` dari `null` menjadi `"Need Review"`.
- `needs_review`: tetap 21 item `TAG_CN_OUT_OF_RANGE`.
- Validasi: record berisi `Need Review` wajib punya item `TAG_CONDITIONAL` atau `TAG_CN_OUT_OF_RANGE`.
- Summary WS TAG COLOR F27: 646 terisi / 78 Need Review (57 + 21) / 0 null. Sebelumnya 646 terisi / 78 null.

## 5. WHS_STYLE_UNKNOWN — diputuskan 2026-11-10

**Keputusan:** background nomor putih, atau tanpa warna, sama artinya dengan abu, yaitu WHS style **`000 & SY`**. Kasus di F27: LUXE (CHRISTMAS - COTTON) No. 4 MERRY STRIPED CREW COTTON, sel D324 tanpa fill.

**Dampak yang direncanakan (belum dikerjakan):**
- Aturan `style_whs`: tanpa fill, putih `FFFFFFFF`, atau theme putih (theme 0 tanpa tint) → `000 & SY`, ditambahkan di `STYLE_WHS_BY_FILL` di `legend_config.py`. Warna lain yang tidak dikenal tetap `null` + warning + `WHS_STYLE_UNKNOWN`.
- Batas aturan: hanya berlaku bila sel nomornya ada. Kartu yang nomornya tidak ditemukan sama sekali tetap `null` + warning D8, karena tidak ada sel yang bisa dibaca warnanya.
- Raw JSON: MERRY STRIPED `style_whs` dari `null` menjadi `"000 & SY"`. Warning D324 ("background nomor tidak dikenali") hilang, jadi warning 4 → 3. Raw golden ikut berubah dan perlu disetujui ulang.
- `style_whs` F27: 85 → 86 `000 & SY`, null 1 → 0.
- Records: tidak berubah. WHS final MERRY tetap `000 & SY` (record 668) dan `000` (record 669), karena irisan dengan style `000 & SY` menghasilkan WHS warna yang sama.
- `needs_review`: 2 item `WHS_STYLE_UNKNOWN` dan 1 item `EXTRACTION_WARNING` (D324) hilang.
- PRD: open issue "Background nomor kosong (MERRY STRIPED)" ditutup. Baris "abu → 000 & SY" di tabel legend diperluas menjadi "abu atau putih/tanpa warna → 000 & SY".

## 6. COLOR_IS_DESCRIPTOR — diputuskan 2026-11-10

**Keputusan:** sementara diisi `Need Review`, karena benang aslinya masih harus dipastikan. Tetap masuk `needs_review` (6 item di F27, comment sel MATRIX tetap disertakan).

**Kasus F27:** RIHANNA STRIPED CARDI (`BLUSH CHAMPAGNE STRIPE`), KEY WEST CREW (`INDIGO COAST HEATHER STRIPE`, `KHAKI STRIPE`), PARK BUTTON CARDI (`BREAKER WHITE MARL`, `ARCTIC BLUE MARL` di COLOR2), SYLVAN MARLED CARDI (`PURE SNOW MARL` di COLOR2).

**Kolom yang diisi (dijawab user):** hanya `WS TAG COLOR` yang menjadi `Need Review`. Slot COLOR tetap persis seperti di MATRIX.

**Dampak yang direncanakan (belum dikerjakan):**
- Records: 6 record berubah, `WS TAG COLOR` menjadi `Need Review`: 167 (dulu `BLUSH CHAMPAGNE STRIPE`), 188, 189, 206 (dulu `SILVERSTONE`), 207 (dulu `AZURE SLATE`), 379 (dulu `OYSTER/PURE SNOW MARL`). COLOR1–COLOR18 tidak berubah.
- `needs_review`: tetap 6 item `COLOR_IS_DESCRIPTOR`.
- Validasi: `COLOR_IS_DESCRIPTOR` ditambahkan ke daftar kode yang sah untuk `Need Review`.
- Golden case PROMPT §9 "RIHANNA STRIPED … WS TAG COLOR `BLUSH CHAMPAGNE STRIPE`" berubah menjadi `Need Review`. Aturan pemadatan akhiran dobel tetap ada dan tetap diuji lewat unit test.

## 7. TAG belum terselesaikan (kasus GONDOLA / C(CAR)) — diputuskan 2026-11-10

**Keputusan:** diberi `Need Review`.

**Kasus F27 (7 record):** TREE FARM CREW CHUNKY (LUXE, CHRISTMAS - GRASSY, No. 7), TAG `TAG: C1/C(car)`, record 588–589; TELLURIDE CREW CHUNKY (LUXE, SKI & SNOW, No. 11), TAG `TAG: C1/GONDOLA color`, record 706–710.

**Dampak yang direncanakan (belum dikerjakan):**
- Aturan baru WS TAG COLOR: bila setelah Cn diganti teks TAG masih berisi `C(` (Cn tanpa nomor) atau kata `COLOR`/`COLOUR` (instruksi warna yang belum diterjemahkan), maka `WS TAG COLOR` = `Need Review` dan item review baru dengan kode **`TAG_UNRESOLVED`**. Polanya disimpan di `legend_config.py`. Dicek pada F27: dari 22 TAG unik, hanya kedua TAG ini yang kena.
- Records: 7 record berubah, `WS TAG COLOR` dari `FAWN/C(CAR)`, `…/GONDOLA COLOR` menjadi `Need Review`.
- `needs_review`: +7 item `TAG_UNRESOLVED`.
- Summary WS TAG COLOR F27 setelah keputusan 3, 4 dan 7: 639 terisi / 85 Need Review / 0 null. Angka akhir ada di ringkasan.
- PRD: kasus ini dicatat sebagai aturan baru (sebelumnya tidak tertangkap dan lolos ke golden).

## 8. Karakter nyasar di area warna — diputuskan 2026-11-10

**Keputusan:** dihapus saja. Sel yang hanya berisi tanda baca (`.`, `` ` ``, `*`) diabaikan tanpa warning.

**Dampak yang direncanakan (belum dikerjakan):**
- Raw JSON: warning G169, E227 dan D345 hilang. Bersama keputusan 5, warning F27 4 → 0. Sel itu memang sudah tidak menjadi colorway, jadi `styles` tidak berubah karena keputusan ini.
- `needs_review`: 3 item `EXTRACTION_WARNING` hilang. Kode `EXTRACTION_WARNING` tetap dipakai untuk warning lain (mis. garis merah tidak ditemukan, TAG di area warna).
- PRD dan PROMPT §6.10 diperbarui ("diabaikan + warning" → "diabaikan").

## 9. Benang berulang dipertahankan — diputuskan 2026-11-13

**Keputusan (user):** benang yang berulang di satu colorway (stripe, mis. `BLACK, BLACK, BROWN BRICK`) tidak lagi digabung. Urutan benang di MATRIX adalah urutan warna (sesuai COLOR SEQUENCE di IM), jadi COLOR1..COLORn ditulis apa adanya.

**Turunan (keputusan perancang, dicatat ke user):**
- Cn di TAG = benang ke-n sesuai urutan MATRIX, benang berulang ikut dihitung.
- Solid dan syarat `(For <nama colorway>)` dinilai dari benang unik: `FRIGID GRAY, FRIGID GRAY` tetap solid.
- `TOO_MANY_YARNS` dihapus. Final tidak dibatasi 18; colorway yang melebihi kolom COLOR di IM diperingatkan saat menulis.
- Validasi header: key COLOR di atas jumlah kolom COLOR IM bukan error (kelebihan kolom), key COLOR yang salah nama tetap error.

**Dampak pada golden F27 (dikerjakan):**
- Records: 66 record berubah kolom COLOR (benang berulang kembali di posisinya). 8 record lebih dari 18 benang, maksimum 32.
- `WS TAG COLOR`: 21 record dari `Need Review` menjadi terisi, semuanya bekas `TAG_CN_OUT_OF_RANGE` (mis. PEBBLE STITCH `TAG: C1/C3 MARL` → `BLACK/BROWN BRICK MARL`). Satu di antaranya, MIXX V CHUNKY E97 (`TAG: C1/C2 MARL`, benang `TWILIGHT SKY, TWILIGHT SKY`), awalnya menjadi `TWILIGHT SKY/TWILIGHT SKY MARL`. Aturan tambahan dari user (13 November 2026): nama warna yang sama di deret Cn bergaris miring (C1/C2, C1/C2/C3) ditulis sekali di TAG, jadi `TWILIGHT SKY MARL`; kolom COLOR tetap terpisah (COLOR1 dan COLOR2 sama-sama TWILIGHT SKY). Di F27 hanya record ini yang terdampak.
- Summary: 699 terisi / 25 `Need Review` / 0 null (sebelumnya 678 / 46 / 0). `needs_review` 25: 12 TAG_CONDITIONAL, 7 TAG_UNRESOLVED, 6 COLOR_IS_DESCRIPTOR.

## Ringkasan dampak semua keputusan (F27, rencana)

| Metrik | Sekarang | Setelah keputusan 1–8 |
| --- | --- | --- |
| Kartu style / colorway | 215 / 724 | tidak berubah |
| `style_whs` | 85 `000 & SY`, 70 `000`, 59 `SY`, 1 null | 86 `000 & SY`, 70 `000`, 59 `SY`, 0 null |
| Warning ekstraksi | 4 | 0 |
| WHS final | 354 `000`, 194 `000 & SY`, 176 `SY` | tidak berubah |
| WS TAG COLOR | 646 terisi, 78 null | 633 terisi, 91 `Need Review`, 0 null |
| `needs_review` | 115 | 91: 57 TAG_CONDITIONAL, 21 TAG_CN_OUT_OF_RANGE, 7 TAG_UNRESOLVED (baru), 6 COLOR_IS_DESCRIPTOR |
| Record yang berubah dibanding golden | — | 91, semuanya hanya kolom `WS TAG COLOR` |

Kode yang tidak dipakai lagi: `DUPLICATE_STYLE_ACROSS_SHEETS`, `WHS_CAPPED_BY_STYLE`. Kode baru: `TAG_UNRESOLVED`. Masih terbuka: `TAG_EMPTY` (0 kasus di F27; usulan ikut `Need Review`).

## Lampiran: record yang berubah dibanding golden lama (F27)

Semua perubahan hanya di kolom `WS TAG COLOR`, dari nilai lama menjadi `Need Review`. Raw JSON: hanya MERRY STRIPED (`style_whs` null → `000 & SY`), dan warning 4 → 0.

### TAG_CONDITIONAL (57 record)

| Record | DESCRIPTION | _ref | WS TAG COLOR lama |
| --- | --- | --- | --- |
| 61 | ARIA V CHUNKY COTTON | ESSENTIALS COTTON No.15 | G98 | above_black | null |
| 62 | ARIA V CHUNKY COTTON | ESSENTIALS COTTON No.15 | G99 | between_black_red | null |
| 63 | ARIA V CHUNKY COTTON | ESSENTIALS COTTON No.15 | G100 | between_black_red | null |
| 64 | ARIA V CHUNKY COTTON | ESSENTIALS COTTON No.15 | G105 | below_red | null |
| 65 | ARIA V CHUNKY COTTON | ESSENTIALS COTTON No.15 | G106 | below_red | null |
| 294 | HUTTON MOCK T NECK CHUNKY | ESSENTIALS GRASSY No.20 | C163 | above_black | null |
| 295 | HUTTON MOCK T NECK CHUNKY | ESSENTIALS GRASSY No.20 | C164 | between_black_red | null |
| 296 | HUTTON MOCK T NECK CHUNKY | ESSENTIALS GRASSY No.20 | C165 | between_black_red | null |
| 297 | HUTTON MOCK T NECK CHUNKY | ESSENTIALS GRASSY No.20 | C166 | between_black_red | null |
| 298 | HUTTON MOCK T NECK CHUNKY | ESSENTIALS GRASSY No.20 | C167 | between_black_red | null |
| 299 | HUTTON MOCK T NECK CHUNKY | ESSENTIALS GRASSY No.20 | C171 | below_red | null |
| 300 | HUTTON MOCK T NECK CHUNKY | ESSENTIALS GRASSY No.20 | C172 | below_red | null |
| 301 | HUTTON MOCK T NECK CHUNKY | ESSENTIALS GRASSY No.20 | C173 | below_red | null |
| 320 | FRAYED EDGE JACKET CHUNKY | ESSENTIALS GRASSY No.25 | H163 | above_black | null |
| 321 | FRAYED EDGE JACKET CHUNKY | ESSENTIALS GRASSY No.25 | H164 | above_black | null |
| 322 | FRAYED EDGE JACKET CHUNKY | ESSENTIALS GRASSY No.25 | H165 | between_black_red | null |
| 323 | FRAYED EDGE JACKET CHUNKY | ESSENTIALS GRASSY No.25 | H166 | between_black_red | null |
| 324 | JANE BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.26 | I163 | above_black | null |
| 325 | JANE BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.26 | I164 | between_black_red | null |
| 326 | JANE BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.26 | I165 | between_black_red | null |
| 327 | JANE BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.26 | I166 | between_black_red | null |
| 328 | JANE BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.26 | I171 | below_red | null |
| 329 | JANE BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.26 | I172 | below_red | null |
| 330 | JANE BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.26 | I173 | below_red | null |
| 341 | MAYA V BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.30 | D231 | above_black | null |
| 342 | MAYA V BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.30 | D232 | between_black_red | null |
| 343 | MAYA V BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.30 | D233 | between_black_red | null |
| 344 | MAYA V BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.30 | D234 | between_black_red | null |
| 345 | MAYA V BUTTON CARDI CHUNKY | ESSENTIALS GRASSY No.30 | D235 | between_black_red | null |
| 349 | COCOON WRAP CARDI CHUNKY | ESSENTIALS GRASSY No.32 | F231 | above_black | null |
| 350 | COCOON WRAP CARDI CHUNKY | ESSENTIALS GRASSY No.32 | F232 | between_black_red | null |
| 351 | COCOON WRAP CARDI CHUNKY | ESSENTIALS GRASSY No.32 | F233 | between_black_red | null |
| 352 | COCOON WRAP CARDI CHUNKY | ESSENTIALS GRASSY No.32 | F240 | below_red | null |
| 353 | COCOON WRAP CARDI CHUNKY | ESSENTIALS GRASSY No.32 | F241 | below_red | null |
| 534 | SHEPHERD CARDI CHUNKY | FALL No.26 | E221 | above_black | null |
| 535 | SHEPHERD CARDI CHUNKY | FALL No.26 | E222 | between_black_red | null |
| 536 | SHEPHERD CARDI CHUNKY | FALL No.26 | E223 | between_black_red | null |
| 597 | HUTTON MOCK T NECK CHUNKY | LUXE No.11 | D94 | above_black | null |
| 598 | HUTTON MOCK T NECK CHUNKY | LUXE No.11 | D95 | between_black_red | null |
| 599 | HUTTON MOCK T NECK CHUNKY | LUXE No.11 | D96 | between_black_red | null |
| 600 | HUTTON MOCK T NECK CHUNKY | LUXE No.11 | D97 | between_black_red | null |
| 601 | HUTTON MOCK T NECK CHUNKY | LUXE No.11 | D98 | between_black_red | null |
| 602 | HUTTON MOCK T NECK CHUNKY | LUXE No.11 | D99 | below_red | null |
| 603 | HUTTON MOCK T NECK CHUNKY | LUXE No.11 | D100 | below_red | null |
| 604 | HUTTON MOCK T NECK CHUNKY | LUXE No.11 | D101 | below_red | null |
| 624 | CANDY CANES HEART CREW CHUNKY | LUXE No.18 | C162 | above_black | null |
| 625 | CANDY CANES HEART CREW CHUNKY | LUXE No.18 | C163 | between_black_red | null |
| 626 | CANDY CANES HEART CREW CHUNKY | LUXE No.18 | C166 | below_red | null |
| 627 | CHEERS COCKTAIL CREW CHUNKY | LUXE No.19 | D162 | above_black | null |
| 628 | CHEERS COCKTAIL CREW CHUNKY | LUXE No.19 | D163 | between_black_red | null |
| 629 | CHEERS COCKTAIL CREW CHUNKY | LUXE No.19 | D166 | below_red | null |
| 630 | MERRY COCKTAIL CREW CHUNKY | LUXE No.20 | E162 | above_black | null |
| 631 | MERRY COCKTAIL CREW CHUNKY | LUXE No.20 | E163 | between_black_red | null |
| 632 | MERRY COCKTAIL CREW CHUNKY | LUXE No.20 | E164 | between_black_red | null |
| 633 | MERRY COCKTAIL CREW CHUNKY | LUXE No.20 | E166 | below_red | null |
| 645 | FA LA LA FAIR ISLE CREW CHUNKY | LUXE (CHRISTMAS - GRASSY) No.S3 | C225 | below_red | null |
| 646 | FA LA LA FAIR ISLE CREW CHUNKY | LUXE (CHRISTMAS - GRASSY) No.S3 | C226 | below_red | null |

### TAG_CN_OUT_OF_RANGE (21 record)

| Record | DESCRIPTION | _ref | WS TAG COLOR lama |
| --- | --- | --- | --- |
| 276 | MIXX V CHUNKY | ESSENTIALS GRASSY No.14 | E97 | between_black_red | null |
| 452 | PEBBLE STITCH MARLED RAGLAN SWEATSHIRT CHUNKY | FALL No.1 | A29 | above_black | null |
| 453 | PEBBLE STITCH MARLED RAGLAN SWEATSHIRT CHUNKY | FALL No.1 | A30 | between_black_red | null |
| 454 | PEBBLE STITCH MARLED RAGLAN SWEATSHIRT CHUNKY | FALL No.1 | A31 | between_black_red | null |
| 455 | PEBBLE STITCH MARLED RAGLAN SWEATSHIRT CHUNKY | FALL No.1 | A32 | between_black_red | null |
| 491 | TATUM CREW CHUNKY | FALL No.12 | E93 | above_black | null |
| 492 | TATUM CREW CHUNKY | FALL No.12 | E94 | between_black_red | null |
| 493 | TATUM CREW CHUNKY | FALL No.12 | E95 | between_black_red | null |
| 498 | SAGE CREW CHUNKY | FALL No.14 | G93 | above_black | null |
| 499 | SAGE CREW CHUNKY | FALL No.14 | G94 | between_black_red | null |
| 500 | SAGE CREW CHUNKY | FALL No.14 | G95 | between_black_red | null |
| 501 | SAGE CREW CHUNKY | FALL No.14 | G96 | between_black_red | null |
| 640 | NOEL TREE CARDI CHUNKY | LUXE No.23 | H162 | above_black | null |
| 641 | NOEL TREE CARDI CHUNKY | LUXE No.23 | H163 | between_black_red | null |
| 665 | NOEL TREE CREW CHUNKY COTTON | LUXE (CHRISTMAS - COTTON) No.3 | C343 | above_black | null |
| 666 | NOEL TREE CREW CHUNKY COTTON | LUXE (CHRISTMAS - COTTON) No.3 | C344 | between_black_red | null |
| 667 | NOEL TREE CREW CHUNKY COTTON | LUXE (CHRISTMAS - COTTON) No.3 | C345 | between_black_red | null |
| 697 | SKYE SHRUNKEN CREW CHUNKY | LUXE (SKI & SNOW) No.7 | A526 | above_black | null |
| 698 | SKYE SHRUNKEN CREW CHUNKY | LUXE (SKI & SNOW) No.7 | A527 | between_black_red | null |
| 711 | SKYE CARDI CHUNKY | LUXE (SKI & SNOW) No.12 | F526 | above_black | null |
| 712 | SKYE CARDI CHUNKY | LUXE (SKI & SNOW) No.12 | F527 | between_black_red | null |

### TAG_UNRESOLVED (7 record)

| Record | DESCRIPTION | _ref | WS TAG COLOR lama |
| --- | --- | --- | --- |
| 588 | TREE FARM CREW CHUNKY | LUXE (CHRISTMAS - GRASSY) No.7 | H27 | above_black | `FAWN/C(CAR)` |
| 589 | TREE FARM CREW CHUNKY | LUXE (CHRISTMAS - GRASSY) No.7 | H28 | between_black_red | `ALMOND BUTTER/C(CAR)` |
| 706 | TELLURIDE CREW CHUNKY | LUXE (SKI & SNOW) No.11 | E526 | above_black | `SILVER MOON/GONDOLA COLOR` |
| 707 | TELLURIDE CREW CHUNKY | LUXE (SKI & SNOW) No.11 | E527 | between_black_red | `ALMOND BUTTER/GONDOLA COLOR` |
| 708 | TELLURIDE CREW CHUNKY | LUXE (SKI & SNOW) No.11 | E530 | below_red | `BLACK/GONDOLA COLOR` |
| 709 | TELLURIDE CREW CHUNKY | LUXE (SKI & SNOW) No.11 | E531 | below_red | `PURE SNOW/GONDOLA COLOR` |
| 710 | TELLURIDE CREW CHUNKY | LUXE (SKI & SNOW) No.11 | E532 | below_red | `CHINA BLUE/GONDOLA COLOR` |

### COLOR_IS_DESCRIPTOR (6 record)

| Record | DESCRIPTION | _ref | WS TAG COLOR lama |
| --- | --- | --- | --- |
| 167 | RIHANNA STRIPED CARDI CHUNKY COTTON | ESSENTIALS COTTON (WINTER BEACH) No.1 | A370 | above_black | `BLUSH CHAMPAGNE STRIPE` |
| 188 | KEY WEST CREW CHUNKY COTTON | ESSENTIALS COTTON No.S1 | A440 | below_red | `INDIGO COAST HEATHER STRIPE` |
| 189 | KEY WEST CREW CHUNKY COTTON | ESSENTIALS COTTON No.S1 | A441 | below_red | `KHAKI STRIPE` |
| 206 | PARK BUTTON CARDI CHUNKY COTTON | ESSENTIALS COTTON No.S5 | F435 | below_red | `SILVERSTONE` |
| 207 | PARK BUTTON CARDI CHUNKY COTTON | ESSENTIALS COTTON No.S5 | F436 | below_red | `AZURE SLATE` |
| 379 | SYLVAN MARLED CARDI CHUNKY | ESSENTIALS GRASSY No.S5 | E297 | below_red | `OYSTER/PURE SNOW MARL` |
