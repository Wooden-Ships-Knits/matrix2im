# Tulis ke IM MASTER (fase 2) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** TIM IM bisa menulis `final.json` ke sheet F27 IM MASTER bersama lewat UI di PC pusat, tanpa merusak isi file dan dengan pengaman password, cadangan, deteksi perubahan, verifikasi, dan log.

**Architecture:** Python murni menyusun rencana:
- `im_layout.py` membaca tata letak IM.
- `write_plan.py` menggabungkan kartu, lalu menyusun rencana sisip baris dan rencana nilai sel.

Penulisan dua langkah, keduanya pada salinan sementara:
1. Excel di Mac (`excel_mac.py`, AppleScript) hanya menyisipkan baris, menjalankan `fill down`, dan memperlebar rumus subtotal grup 1 baris.
2. `xlsx_cells.py` mengubah sel COLOR/WS TAG COLOR/WHS langsung di XML sheet.

`im_writer.py` mengatur urutan: password → cek file dibuka → sidik jari → cadangan → Excel → nilai → verifikasi → salin balik ke file asli di tempat yang sama → log. UI server mendapat sesi per browser, endpoint tulis, dan kunci satu penulisan.

**Tech Stack:** Python 3.11+ (venv `.venv`, `/opt/homebrew/bin/python3.12`), openpyxl 3.1.x (hanya baca untuk IM), pytest, argparse, `zipfile`/`hashlib`/`tomllib`/`subprocess` bawaan, `osascript` bawaan macOS, Microsoft Excel for Mac 16.x.

**Spec:** `docs/superpowers/specs/2026-11-10-tulis-im-master-design.md` (otoritas). Aturan penulisan 1–7: `docs/PRD_MATRIX_MASTER_to_IM_JSON.md` bagian "Fase berikutnya: penulisan ke IM MASTER". Hasil spike yang sudah terbukti ada di spec §4. Jangan diulang, tapi jangan pula dilanggar.

## Global Constraints

**Dependensi dan format**
- Dependensi runtime tetap hanya `openpyxl`, dev hanya `pytest`; sisanya library bawaan Python. Pillow tidak dipasang.
- Pesan UI, CLI, error, dan log dalam Bahasa Indonesia. Type hints di semua fungsi publik; docstring menjelaskan *kenapa*.
- Konstanta aturan hanya di `matrix_to_im/legend_config.py`. Tidak ada huruf kolom atau nomor baris IM yang di-hardcode: kolom dicari dari header baris `IM_HEADER_ROW` di sheet `IM_SHEET`.

**Keamanan file IM MASTER**
- openpyxl **tidak pernah** `save()` file IM MASTER atau salinannya. Penyimpanan ulang lewat openpyxl terbukti menghapus 22 gambar dan 7 aturan conditional formatting.
- Yang boleh mengubah salinan hanya Excel (struktur) dan `xlsx_cells.py` (nilai).
- File asli hanya ditulis sekali, di akhir, dengan menimpa isi file di tempat yang sama (`open(path, "r+b")`, lalu `write` + `truncate`). Tidak boleh lewat rename atau hapus-buat, supaya identitas file di Google Drive tetap.

**Password**
- Password tidak pernah ditulis ke disk, log, laporan, argumen command line, atau pesan error.
- Password dikirim ke `osascript` hanya lewat environment `MATRIX2IM_PW`.

**Test**
- Test tidak pernah menyentuh file di Google Drive atau file di `data/`.
- Test regression hanya membaca file `data/` dan menulis ke salinan di `tmp_path`.
- Test yang menjalankan Excel diberi marker `excel` dan hanya jalan bila `MATRIX2IM_EXCEL_TESTS=1`, macOS, dan Excel terpasang.

**Git**
- File data, cadangan, log, dan `matrix2im.local.toml` tidak di-commit.

**Hasil spike yang wajib diikuti (spec §4)**
- Sisip memakai **baris utuh** (`range "57:58"`), bukan range kolom.
- Salinan baris memakai `fill down` yang **dibatasi** `A{baris}:{kolom terakhir}{baris}`, bukan `copy range`, karena `copy range` memecah conditional formatting.
- Parameter `read only true` pada `open workbook` menghasilkan error -50, jadi jangan dipakai.

## Review Focus

Lima masukan yang paling mungkin merugikan pengguna, dengan test pengikatnya:

1. **Header IM bervariasi.** Contoh: header berbaris ganda (`"WS TAG COLOR\n(ket)"`), dua kolom "WHS", atau "WHS" tidak di kanan WS TAG COLOR. Harapan: kolom yang benar dipakai, atau error jelas tanpa menulis apa pun. Diikat di Task 2.
2. **Sel tujuan berisi rumus.** Harapan: tidak pernah ditimpa, dan muncul di peringatan. Diikat di Task 5 dan Task 6.
3. **Nilai dengan `&`, `<`, `"`, dan huruf non-ASCII** (`APRÈS`). Harapan: Excel membuka file tanpa repair dan nilainya sama persis. Diikat di Task 6.
4. **IM MASTER berubah atau dibuka orang lain antara pratinjau dan tulis.** Harapan: penulisan dibatalkan, file asli tidak disentuh, pesan menyebut alasannya. Diikat di Task 8.
5. **Dua pengguna bersamaan.** Harapan: upload dan hasil tidak tercampur antar-browser, dan penulisan kedua ditolak dengan nama penulis yang sedang berjalan. Diikat di Task 10.

---

## File Structure

| File | Tanggung jawab |
| --- | --- |
| `matrix_to_im/legend_config.py` (ubah) | Konstanta baru IM, aksesoris, WHS, Excel |
| `matrix_to_im/im_lock.py` (baru) | Password to modify (`fileSharing`), deteksi file pemilik `~$` |
| `matrix_to_im/im_layout.py` (baru) | Baca kolom, grup size, nilai sel tujuan, dan rumus subtotal dari IM |
| `matrix_to_im/write_plan.py` (baru) | Gabung kartu per style, rencana sisip, proyeksi tata letak, rencana nilai |
| `matrix_to_im/xlsx_cells.py` (baru) | Ubah sel di XML sheet; bagian lain file tetap byte-identik |
| `matrix_to_im/excel_mac.py` (baru) | Bangun dan jalankan AppleScript sisip baris di Excel |
| `matrix_to_im/settings.py` (baru) | Baca `matrix2im.local.toml` |
| `matrix_to_im/im_writer.py` (baru) | `prepare_write` / `commit_write`: urutan pengaman lengkap |
| `matrix_to_im/cli.py` (ubah) | Subcommand `write`; `ui --host --settings` |
| `matrix_to_im/ui_server.py` (ubah) | Sesi per browser, endpoint tulis, `WriteCoordinator` |
| `matrix_to_im/ui/index.html` (ubah) | Panel "Tulis ke IM MASTER" |
| `tests/helpers/im_factory.py` (baru) | IM sintetis + penyuntik `fileSharing` |
| `matrix2im.local.example.toml` (baru), `.gitignore` (ubah) | Contoh setting; abaikan setting lokal |
| `tests/test_im_lock.py`, `test_im_layout.py`, `test_write_plan.py`, `test_xlsx_cells.py`, `test_excel_mac.py`, `test_settings.py`, `test_im_writer.py` (baru); `test_cli.py`, `test_ui_server.py` (ubah) | Unit test |
| `tests/regression/test_write_f27.py` (baru) | Regression pada salinan IM AI TEST V2 |
| `README.md`, `docs/PRD_MATRIX_MASTER_to_IM_JSON.md`, `docs/uji-terima-tulis-im.md` (baru) | Dokumentasi, panduan PC pusat, uji terima manual |

Konstanta baru di `legend_config.py`, persis seperti ini:

```python
IM_DESCRIPTION_KEY = "DESCRIPTION"
IM_WS_TAG_KEY = "WS TAG COLOR"
IM_WHS_KEY = "WHS"                      # yang ditulis: kolom tepat di kanan WS TAG COLOR (PRD aturan 6)
IM_SUBTOTAL_RE = r" - (BOTH WHS|000 ONLY|SHOPIFY ONLY)\b"
IM_ACCESSORY_RE = r"^X[A-Z]+(\s|$)"     # mis. "XBP 729 2 CM-DARK ROAST", "XZIP"
WHS_PARTS = {"000": frozenset({"000"}), "SY": frozenset({"SY"}), "000 & SY": frozenset({"000", "SY"})}
EXCEL_APP = "/Applications/Microsoft Excel.app"
EXCEL_TIMEOUT_S = 900
```

---

### Task 1: Kunci IM MASTER (`im_lock.py`) dan IM sintetis

**Files:**
- Create: `matrix_to_im/im_lock.py`, `tests/helpers/im_factory.py`, `tests/test_im_lock.py`
- Modify: `matrix_to_im/legend_config.py` (semua konstanta baru di atas)

**Interfaces:**
- Produces:
  - `@dataclass(frozen=True) FileSharing(user_name: str | None, algorithm: str | None, hash_value: str | None, salt_value: str | None, spin_count: int | None, read_only_recommended: bool, legacy: bool)`
  - `read_file_sharing(path: Path) -> FileSharing | None`
  - `hash_modify_password(password: str, salt_b64: str, spin_count: int, algorithm: str = "SHA-512") -> str`
  - `check_modify_password(path: Path, password: str) -> None`
  - `class WrongPassword(Matrix2ImError)`, `class NotLocked(Matrix2ImError)`
  - `owner_file(path: Path) -> Path | None`, `owner_name(owner: Path) -> str`
- Produces (test helper `tests/helpers/im_factory.py`):
  - `TEST_PASSWORD = "rahasia-uji"`
  - `TEST_SALT_B64 = "AAECAwQFBgcICQoLDA0ODw=="`
  - `TEST_HASH_SPIN1000 = "PyNhGFHLSUU4sQ1nbaw7z3yfxE0RsH4rbX3I46rm+XP4xYVebx88dqEbZ8rPlx/BMzftVlY7uScXWiCAscS58Q=="`
  - `TEST_HASH_SPIN100000 = "hzY8X1+Zw+dryKltyS4V+shqaVXxQ9h3IGLqGwSL3/G1Et9T0OIbktcJNfqEkIQzY2rr+r5nPM3mF7c9GkPQig=="`
  - `GroupSpec(description: str, colorways: list[list[str]], subtotal: str | None = "BOTH WHS", ws_tags: list[str | None] | None = None, whs: list[str | None] | None = None)`
  - `build_im_sheet(path: Path, groups: list[GroupSpec], *, extra_header: dict[str, int] | None = None) -> Path`
  - `add_file_sharing(path: Path, hash_b64: str = TEST_HASH_SPIN1000, salt_b64: str = TEST_SALT_B64, spin: int = 1000, user: str = "Tester") -> Path`

Rincian:
- Algoritma hash (ECMA-376, terbukti cocok dengan file asli):
  `h = H(salt + password.encode("utf-16-le"))`, lalu ulangi `spin` kali `h = H(h + i.to_bytes(4, "little"))` untuk `i` dari 0, kemudian base64.
- Nama algoritma `SHA-512`/`SHA-384`/`SHA-256`/`SHA-1` dipetakan ke `hashlib`. Algoritma lain menghasilkan `Matrix2ImError("Algoritma kunci IM MASTER tidak didukung: …")`.
- Atribut format lama `reservationPassword` membuat `legacy=True`, dan `check_modify_password` melempar `Matrix2ImError("Kunci IM MASTER memakai format lama (reservationPassword) yang belum didukung")`.
- Tidak ada `fileSharing` atau tidak ada `hashValue` → `NotLocked("IM MASTER tidak dikunci Password to modify; penulisan dibatalkan. Periksa apakah file yang dipilih benar.")`.
- Password salah → `WrongPassword("Password IM MASTER salah")`.
- `owner_file`: cari di folder yang sama file `"~$" + name` atau `"~$" + name[2:]`, karena Excel kadang memotong 2 huruf pertama nama.
- `owner_name`: byte pertama = panjang n. Bila `0 < n <= 53`, ambil `bytes[1:1+n]` dan decode UTF-8 dengan `errors="replace"`, lalu strip. Selain itu kembalikan `"seseorang"`.

`build_im_sheet` (IM sintetis):
- Sheet `F27` dengan header di baris `IM_HEADER_ROW`, kolom berurutan:

  | Kolom | Header |
  | --- | --- |
  | A | DESCRIPTION |
  | B | PCS |
  | C | WHS (WHS pengecoh, seperti kolom G di IM asli) |
  | D | C1Y% (rumus `=1/2` di baris colorway) |
  | E | COLOR1 |
  | F | C2Y% (rumus) |
  | G | COLOR2 |
  | H | COLOR3 |
  | I | COLOR 4 |
  | J–W | COLOR5..COLOR18 |
  | X | WS TAG COLOR |
  | Y | WHS |
  | Z | NOTE |

  `extra_header` bisa menimpa nama header per kolom, untuk test variasi header.
- Baris pertama di `IM_HEADER_ROW + 1`.
- Tiap baris colorway diisi:
  - DESCRIPTION = `description`, PCS = 10, rumus Y%;
  - COLOR sesuai daftar, dengan fill kuning pada COLOR1 untuk menguji style;
  - WS TAG COLOR dan WHS dari `ws_tags`/`whs` bila ada, NOTE = `"catatan"`.
- Bila `subtotal` tidak None, setelah grup ada baris `f"{description} - {subtotal}"`. PCS di baris itu berisi `=SUBTOTAL(9,B{awal}:B{akhir})`, atau `=SUBTOTAL(9,B{awal})` untuk grup 1 baris.
- `add_file_sharing` menyisipkan `<fileSharing readOnlyRecommended="1" userName=… algorithmName="SHA-512" hashValue=… saltValue=… spinCount=…/>` ke `xl/workbook.xml`, sesudah `<fileVersion …/>` bila ada, atau sebagai anak pertama `<workbook>` (urutan skema). Entri zip lain disalin apa adanya.

- [ ] **Step 1: Tulis test** di `tests/test_im_lock.py`:
  - `test_hash_matches_known_vector`: `hash_modify_password(TEST_PASSWORD, TEST_SALT_B64, 100000) == TEST_HASH_SPIN100000`.
  - `test_check_password_ok_and_wrong`: IM sintetis + `add_file_sharing`. Password benar tidak melempar; `"salah"` melempar `WrongPassword`.
  - `test_unlocked_file_raises_not_locked`
  - `test_read_file_sharing_fields`: `user_name == "Tester"`, `spin_count == 1000`, `read_only_recommended is True`.
  - `test_legacy_reservation_password_rejected`
  - `test_owner_file_full_and_truncated_name`: `"~$IM.xlsx"` ditemukan; untuk `"F27 IM.xlsx"`, `"~$7 IM.xlsx"` ditemukan; `"~$lain.xlsx"` menghasilkan `None`.
  - `test_owner_name_length_prefixed`: `bytes([5]) + b"Putri" + b"\x00" * 48` menghasilkan `"Putri"`; `b"\xff" * 10` menghasilkan `"seseorang"`.
- [ ] **Step 2:** `.venv/bin/pytest tests/test_im_lock.py -q`. Hasil yang diharapkan: FAIL (modul belum ada).
- [ ] **Step 3:** Implementasikan `im_lock.py`, helper, dan konstanta.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus (217 lama + baru).
- [ ] **Step 5:** Commit `feat(im): cek Password to modify dan file pemilik Excel`.

### Task 2: Tata letak IM (`im_layout.py`)

**Files:** Create `matrix_to_im/im_layout.py`, `tests/test_im_layout.py`

**Interfaces:**
- Consumes: `build_im_sheet`, `GroupSpec` (Task 1).
- Produces:
  - `normalize(text: object) -> str`: whitespace dirapatkan, strip, upper.
  - `@dataclass(frozen=True) ImColumns(description: int, colors: tuple[int, ...], ws_tag: int, whs: int, last: int)`: indeks kolom berbasis 1; `colors` berisi 18 kolom dengan urutan `IM_COLOR_KEYS`; `last` = kolom terpakai terakhir.
  - `@dataclass(frozen=True) ImGroup(description: str, rows: tuple[int, ...], subtotal_row: int | None)`
  - `@dataclass ImLayout(columns: ImColumns, groups: list[ImGroup], values: dict[tuple[int, int], object], subtotal_formulas: dict[int, dict[int, str]])`
  - `read_layout(path: Path, sheet: str = IM_SHEET, header_row: int = IM_HEADER_ROW) -> ImLayout`

Rincian:
- Baca dengan openpyxl `read_only=True` tanpa `data_only`, supaya rumus terbaca sebagai teks `"=…"`.
- Header dicocokkan dengan baris pertama tiap sel (`str(v).split("\n")[0].strip()`).
- Header yang tidak ada (DESCRIPTION, 18 COLOR, WS TAG COLOR) → `Matrix2ImError` yang menyebut nama header.
- WHS = kolom `ws_tag + 1`, dan header-nya harus `IM_WHS_KEY`. Selain itu → `Matrix2ImError("Kolom WHS tidak ditemukan tepat di kanan WS TAG COLOR")`.
- `last` = `ws.max_column`.
- Baris di bawah header dengan DESCRIPTION berisi `" - "`:
  - Cocok `IM_SUBTOTAL_RE` (setelah normalize) → baris subtotal.
  - Selain itu → baris colorway.
- Grup = deret baris colorway berurutan dengan DESCRIPTION ternormalisasi yang sama. `subtotal_row` = baris tepat sesudahnya bila itu baris subtotal yang diawali DESCRIPTION grup; selain itu None. Grup M/L berakhiran FAF/FAH menjadi dua grup, masing-masing dengan subtotalnya.
- `values` hanya menyimpan sel kolom COLOR, WS TAG COLOR, dan WHS di baris colorway (termasuk `None`).
- `subtotal_formulas` menyimpan semua sel rumus di baris subtotal. openpyxl sudah menerjemahkan shared formula, terbukti di spike.

- [ ] **Step 1: Tulis test**:
  - `test_columns_by_header_name`: `colors[0] == 5` (E), `colors[3] == 9` (I, "COLOR 4"), `ws_tag == 24`, `whs == 25`. WHS pengecoh di C diabaikan.
  - `test_multiline_header_matched`: `extra_header={"X": "WS TAG COLOR\n(ket)"}`.
  - `test_missing_header_raises`: pesan mengandung `"WS TAG COLOR"`.
  - `test_whs_not_right_of_ws_tag_raises`
  - `test_groups_and_subtotals`:
    - grup 3 baris + subtotal;
    - grup 1 baris + subtotal;
    - grup 2 baris tanpa subtotal;
    - `"S - M/L"`×2 + `"S - M/L - BOTH WHS - FAF"` + `"S - M/L"`×2 + `"S - M/L - BOTH WHS - FAH"` → dua grup terpisah.

    Nomor baris dan `subtotal_row` dicek persis.
  - `test_normalize`: `normalize("Reegan  cable - x/s ") == "REEGAN CABLE - X/S"`.
  - `test_formula_in_target_kept_as_text`: COLOR2 berisi `"=E55"` → `values[(55, 7)] == "=E55"`.
  - `test_subtotal_formulas_captured`: `subtotal_formulas[row][2] == "=SUBTOTAL(9,B55:B57)"`.
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan `im_layout.py`.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `feat(im): baca tata letak grup size IM MASTER`.

### Task 3: Gabung kartu per style (`write_plan.py` bagian 1)

**Files:** Create `matrix_to_im/write_plan.py`, `tests/test_write_plan.py`

**Interfaces:**
- Consumes: `normalize` (Task 2); format record `final.json`: key `IM_COLOR_KEYS` yang ada, `"WS TAG COLOR"`, `"WHS"`, dan `"_ref"` berbentuk `f"{sheet}{section} No.{no} | {cell} | {group}"` dengan section = `" (SECTION)"` atau kosong.
- Produces:
  - `@dataclass(frozen=True) Colorway(colors: tuple[str, ...], ws_tag: str | None, whs: str | None)`
  - `@dataclass(frozen=True) Skip(style: str, reason: str)`
  - `parse_ref(ref: str) -> tuple[str, str]`: mengembalikan `(sheet, kunci kartu)`; kunci kartu = teks sebelum `" | "` pertama.
  - `style_colorways(final: dict) -> tuple[dict[str, list[Colorway]], list[Skip]]`: key = style ternormalisasi.

Aturan (spec §6):
- Record dikelompokkan per style ternormalisasi, lalu per sheet (urutan kemunculan). Di dalam satu sheet, kartu-kartunya disambung sesuai urutan record (kartu A lalu B).
- Satu sheet → daftar itu.
- Lebih dari satu sheet:
  - Himpunan tuple COLOR ternormalisasi harus sama dan panjangnya sama. Bila tidak → `Skip(style, "colorway berbeda antar-sheet: <sheet1> n colorway, <sheet2> m colorway")`.
  - Bila sama, urutan mengikuti sheet pertama. WS TAG COLOR per colorway harus sama di semua sheet; bila tidak → `Skip(style, "WS TAG COLOR berbeda antar-sheet untuk colorway …")`.
  - WHS adalah gabungan `WHS_PARTS`, dikembalikan ke nilai lewat pemetaan terbalik.
  - **[Keputusan perancang]** Bila salah satu WHS `None`, hasilnya `None` (tidak ditulis).
- `parse_ref` memakai regex `^(?P<sheet>.+?)(?: \((?P<section>[^)]*)\))? No\.(?P<no>.+?) \| ` . Bila tidak cocok → `Matrix2ImError`.

- [ ] **Step 1: Tulis test**:
  - `test_parse_ref_with_and_without_section`:
    - `parse_ref("ESSENTIALS GRASSY No.20 | B15 | above_black") == ("ESSENTIALS GRASSY", "ESSENTIALS GRASSY No.20")`
    - `parse_ref("LUXE (WINTER) No.S9 - A | C16 | below_red") == ("LUXE", "LUXE (WINTER) No.S9 - A")`
  - `test_split_cards_same_sheet_concatenated`: kartu `S2 - A` (2 colorway) lalu `S2 - B` (1) menghasilkan 3 colorway berurutan.
  - `test_cross_sheet_same_set_uses_first_order_and_unions_whs`:
    - sheet1 urutan [N, P, G] dengan G=`"000"`; sheet2 urutan [G, N, P] dengan G=`"000 & SY"` → urutan [N, P, G], G=`"000 & SY"`;
    - pasangan `"000"`+`"SY"` → `"000 & SY"`.
  - `test_cross_sheet_different_set_skipped`: alasan mengandung `"berbeda"`; style tidak ada di dict.
  - `test_cross_sheet_ws_tag_differs_skipped`
  - `test_whs_none_wins`
  - `test_color_case_insensitive_match`: `"Glasgow"` vs `"GLASGOW"` dianggap sama.
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `feat(plan): gabung kartu A/B dan style lintas sheet`.

### Task 4: Rencana sisip baris dan proyeksi (`write_plan.py` bagian 2)

**Files:** Modify `matrix_to_im/write_plan.py`, `tests/test_write_plan.py`

**Interfaces:**
- Consumes: `ImLayout`, `ImGroup` (Task 2); `Colorway` (Task 3).
- Produces:
  - `match_style(description: str, styles: Iterable[str]) -> str | None`: memilih style terpanjang yang membuat `description.startswith(style + " - ")`.
  - `@dataclass(frozen=True) InsertOp(description: str, at_row: int, count: int, fill_from: int, fill_to: int, subtotal_row_after: int | None, formula_fixes: dict[int, str])`
  - `plan_inserts(layout: ImLayout, colorways: dict[str, list[Colorway]]) -> list[InsertOp]`: diurutkan `at_row` menurun.
  - `project_layout(layout: ImLayout, ops: list[InsertOp]) -> ImLayout`

Aturan (spec §6, terbukti di spike):
- Grup dengan R baris dan N colorway, N > R, k = N − R:
  - **R ≥ 2:** `at_row = rows[-1]`, `fill_from = rows[-1] - 1`, `fill_to = rows[-1] + k - 1`, `subtotal_row_after = None`, `formula_fixes = {}`.
  - **R = 1:** `at_row = rows[0] + 1`, `fill_from = rows[0]`, `fill_to = rows[0] + k`.
    - Bila ada `subtotal_row`, `subtotal_row_after = subtotal_row + k`.
    - `formula_fixes` = rumus subtotal yang merujuk sel tunggal `{COL}{r}`, yaitu yang tidak didahului huruf, `$`, atau `:`, dan tidak diikuti digit atau `:`. Sel itu diganti `{COL}{r}:{COL}{r+k}` dengan `$` dipertahankan di kedua sisi.
    - Range `X{r}:X{r}` juga diperlebar menjadi `X{r}:X{r+k}`.
- Grup yang style-nya tidak cocok atau N ≤ R tidak mendapat op.
- `project_layout`:
  - Baris asli x pindah ke `x + Σ count` untuk op dengan `at_row <= x`.
  - Baris sisipan mendapat salinan `values` dari baris `fill_from`, yang juga sudah diproyeksikan.
  - `rows` grup diperpanjang, dan `subtotal_row` ikut bergeser.
  - `subtotal_formulas` hanya digeser key-nya.

- [ ] **Step 1: Tulis test**:
  - `test_no_insert_when_enough_rows`
  - `test_insert_inside_group_two_or_more_rows`: rows 55–57, N=5 → `InsertOp(at_row=57, count=2, fill_from=56, fill_to=58, subtotal_row_after=None, formula_fixes={})`.
  - `test_single_row_group_inserts_below_and_widens_subtotal`:
    - grup baris 60, subtotal 61 berisi `=SUBTOTAL(9,B60)`, N=3 → `at_row=61, count=2, fill_from=60, fill_to=62, subtotal_row_after=63, formula_fixes={2: "=SUBTOTAL(9,B60:B62)"}`;
    - rumus `=B600*2` di baris itu tidak masuk fixes;
    - `=SUBTOTAL(9,$B$60)` menjadi `=SUBTOTAL(9,$B$60:$B$62)`.
  - `test_single_row_group_without_subtotal`: `subtotal_row_after is None` dan fixes kosong.
  - `test_ops_sorted_bottom_up`
  - `test_match_style_longest_prefix`: style `{"AMELIA ROLLNECK", "AMELIA ROLLNECK CARDI"}`; `"AMELIA ROLLNECK CARDI - X/S"` → `"AMELIA ROLLNECK CARDI"`.
  - `test_project_layout_shifts_rows_and_copies_values`: dua grup; op di grup atas menggeser grup bawah; nilai baris sisipan = nilai baris `fill_from`.
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `feat(plan): rencana sisip baris dan proyeksi tata letak`.

### Task 5: Rencana nilai sel (`write_plan.py` bagian 3)

**Files:** Modify `matrix_to_im/write_plan.py`, `tests/test_write_plan.py`

**Interfaces:**
- Consumes: `ImLayout` (biasanya hasil `project_layout`), `Colorway`, `match_style`.
- Produces:
  - `@dataclass(frozen=True) CellChange(row: int, col: int, old: object, new: str | None)`: `new=None` berarti sel dikosongkan.
  - `@dataclass ValuePlan(changes: list[CellChange], changed_colorway_rows: list[int], short_groups: list[str], unmatched_im_styles: list[str], missing_styles: list[str], warnings: list[str])`
  - `plan_values(layout: ImLayout, colorways: dict[str, list[Colorway]]) -> ValuePlan`

Aturan (PRD 1–7 dan spec §6):
- **Urutan baris.** Baris ke-i grup (0-based) mendapat colorway `min(i, N-1)`. Bila R < N (mode tanpa sisip), hanya R colorway pertama yang ditulis, dan deskripsi grup masuk `short_groups` dengan teks `"<deskripsi>: kurang k baris"`.
- **Slot COLOR ke-j.**
  - j < jumlah benang → nilai benang (selalu ditimpa).
  - Selain itu:
    - sel kosong → tidak berubah;
    - cocok `IM_ACCESSORY_RE` → tidak berubah;
    - lainnya → dikosongkan.
  - Benang lebih dari 18 → hanya 18 pertama.
- **WS TAG COLOR.**
  - `None` → tidak ditulis.
  - `NEED_REVIEW` → tidak ditulis bila sel lama tidak kosong.
  - Selain itu ditulis.
- **WHS.** `None` → tidak ditulis.
- **Tanpa perubahan.** Nilai baru yang sama persis dengan nilai lama (string) tidak menghasilkan `CellChange`.
- **Sel rumus.** Sel tujuan yang nilai lamanya teks berawalan `"="` tidak pernah diubah dan menambah peringatan `"sel <kolom><baris> berisi rumus, tidak ditulis"`. Huruf kolom memakai `openpyxl.utils.get_column_letter`.
- **Baris berganti colorway.** `changed_colorway_rows` = baris yang tuple benang lamanya (tanpa aksesoris, ternormalisasi) berbeda dari benang baru.
- **Daftar laporan.**
  - `unmatched_im_styles` = bagian sebelum `" - "` dari grup yang tidak cocok style mana pun, unik dan terurut.
  - `missing_styles` = style di `colorways` yang tidak punya grup.

- [ ] **Step 1: Tulis test**:
  - `test_rows_get_colorways_in_order_extra_rows_get_last`: R=4, N=2 → baris 3–4 berisi colorway ke-2.
  - `test_color_slots_overwrite_clear_old_yarn_keep_accessory`:
    - lama `["YCA BLACK B - 1708 B", "YCA CHOCOLATE ALE - 609", "XBP 729 2 CM-DARK ROAST", "YCA OLD - 1"]`, baru `["YCA KHAKI - 749"]` → COLOR1 = KHAKI, COLOR2 dikosongkan, COLOR3 tidak berubah, COLOR4 dikosongkan;
    - baru 3 benang → COLOR3 (aksesoris) ditimpa.
  - `test_need_review_never_replaces_filled_ws_tag`: lama `"BLACK MARL"`, baru `"Need Review"` → tidak ada perubahan; lama `None` → berubah.
  - `test_null_values_not_written`
  - `test_unchanged_values_produce_no_change`
  - `test_formula_target_cell_untouched_with_warning`
  - `test_short_group_fills_existing_rows_and_reports`: `short_groups == ["S - X/S: kurang 2 baris"]`.
  - `test_unmatched_and_missing_styles_reported`
  - `test_changed_colorway_rows_listed`
  - `test_whs_000_is_string`: `new == "000"` bertipe `str`.
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `feat(plan): rencana nilai sel COLOR/WS TAG/WHS`.

### Task 6: Ubah sel langsung di XML (`xlsx_cells.py`)

**Files:** Create `matrix_to_im/xlsx_cells.py`, `tests/test_xlsx_cells.py`

**Interfaces:**
- Consumes: `CellChange` (Task 5); `add_file_sharing`, `build_im_sheet` (Task 1).
- Produces:
  - `sheet_xml_path(zf: zipfile.ZipFile, sheet_name: str) -> str`: lewat `xl/workbook.xml` dan `xl/_rels/workbook.xml.rels`.
  - `apply_cell_values(src: Path, dst: Path, sheet_name: str, changes: Iterable[CellChange]) -> None`

Rincian:
- **Bagian file lain.** Semua entri zip selain XML sheet tujuan dan `xl/workbook.xml` disalin byte-identik dengan `compress_type` yang sama dan urutan entri yang sama.
- **`fullCalcOnLoad`.** Di `xl/workbook.xml`, atribut `fullCalcOnLoad="1"` ditambahkan ke `<calcPr …>`; bila tidak ada `calcPr`, elemen itu dibuat sesuai urutan skema (sebelum `oleSize`/`customWorkbookViews`/`pivotCaches`/`extLst`, atau di akhir).
- **Efisiensi.** XML sheet (F27 sekitar 118 MB) diproses dengan mencari offset `<row r="N"` sekali, lalu potongan disambung. Jangan parse seluruh XML menjadi pohon.
- **Menulis sel.**
  - Sel yang ada diganti dengan `<c r="…" [s="…"] t="inlineStr"><is><t>…</t></is></c>`: atribut `s` dipertahankan, nilai di-escape XML, dan `xml:space="preserve"` dipakai bila ada spasi di awal atau akhir.
  - Mengosongkan sel → `<c r="…" [s="…"]/>`.
- **Sel yang belum ada** disisipkan sesuai urutan kolom. Baris self-closing `<row …/>` dibuka menjadi `<row …>…</row>`.
- **Error** (dst tidak dibuat):
  - sel tujuan berisi `<f` → `Matrix2ImError("Sel <ref> berisi rumus; penulisan dibatalkan")`;
  - baris tidak ada → `Matrix2ImError`.

- [ ] **Step 1: Tulis test**:
  - `test_values_written_and_readable`: dibaca openpyxl.
  - `test_other_parts_byte_identical`: CRC semua entri kecuali sheet XML dan workbook.xml sama.
  - `test_style_preserved`: atribut `s` sel COLOR1 sama sebelum dan sesudah.
  - `test_missing_cell_inserted_in_column_order`: baris dengan sel A dan Z, tulis M → urutan `r` di XML A, M, Z.
  - `test_clear_keeps_style_removes_value`
  - `test_formula_cell_refused_and_no_output`
  - `test_special_chars_and_unicode`: `'A & B <C> "D" APRÈS'` terbaca sama persis.
  - `test_000_stays_text`: openpyxl membaca `"000"` (str).
  - `test_full_calc_on_load_set`
  - `test_file_sharing_preserved_and_password_still_valid`: memakai `check_modify_password`.
  - `test_missing_row_raises`
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `feat(im): tulis nilai sel langsung di XML sheet`.

### Task 7: Penjalan Excel untuk sisip baris (`excel_mac.py`)

**Files:** Create `matrix_to_im/excel_mac.py`, `tests/test_excel_mac.py`; Modify `pyproject.toml` (marker `excel: menjalankan Microsoft Excel sungguhan (butuh MATRIX2IM_EXCEL_TESTS=1)`)

**Interfaces:**
- Consumes: `InsertOp` (Task 4).
- Produces:
  - `class ExcelUnavailable(Matrix2ImError)`
  - `excel_available() -> bool`: `sys.platform == "darwin"` dan `Path(EXCEL_APP).exists()`.
  - `build_insert_script(ops: list[InsertOp], sheet: str, last_col: str) -> str`
  - `run_excel_inserts(path: Path, password: str, ops: list[InsertOp], sheet: str, last_col: str, timeout_s: int = EXCEL_TIMEOUT_S) -> float`: mengembalikan detik.

Isi script, dengan urutan op sesuai daftar (sudah dari bawah ke atas):
1. `open workbook workbook file name f write reserved password pw update links do not update links ignore read only recommended true`
   - `f` = `(POSIX file (item 1 of argv)) as text`
   - `pw` = `system attribute "MATRIX2IM_PW"`
2. `set screen updating to false` dan `set calculation to calculation manual`.
3. Per op:
   - `insert into range (range "{at}:{at+count-1}" of ws) shift shift down`
   - `fill down range "A{fill_from}:{last_col}{fill_to}" of ws`
   - bila ada fixes: kolom berurutan dikelompokkan menjadi run, lalu `set formula of range "{c1}{row}:{c2}{row}" of ws to {{"…", "…"}}`, dengan `\` dan `"` di-escape.
4. `set calculation to calculation automatic`, `set screen updating to true`, `save`, lalu `close … saving no`.

Semua dibungkus `try … on error errMsg number errNum`. Saat error, setting dikembalikan, file ditutup tanpa disimpan, dan script mengembalikan `"ERROR " & errNum & ": " & errMsg`.

Runner:
- `subprocess.run(["osascript", "-", str(path)], input=script, env={**os.environ, "MATRIX2IM_PW": password}, capture_output=True, text=True, timeout=timeout_s)`.
- Excel tidak tersedia → `ExcelUnavailable`.
- Timeout → `Matrix2ImError(f"Excel tidak merespons dalam {timeout_s} detik. Periksa apakah ada dialog terbuka di Excel PC pusat.")`.
- Keluaran `ERROR` atau returncode ≠ 0 → `Matrix2ImError("Excel gagal: …")`. Bila nomor error -1743, tambahkan `" Izinkan Terminal mengontrol Microsoft Excel di System Settings → Privacy & Security → Automation."`.

- [ ] **Step 1: Tulis test**:
  - `test_script_ops_in_given_order_and_fill_limited_to_last_col`: memeriksa string `insert into range (range "57:58"` dan `fill down range "A56:IJ58"`.
  - `test_script_formula_fixes_grouped_in_runs_and_escaped`: rumus `=IF(A1="x",1,2)` muncul sebagai `"=IF(A1=\"x\",1,2)"`; kolom 2,3,5 menjadi dua run.
  - `test_password_only_in_env`: monkeypatch `subprocess.run`, tangkap args dan env. Password tidak ada di `args` maupun `input`, dan ada di `env["MATRIX2IM_PW"]`.
  - `test_timeout_message`
  - `test_error_output_message_and_automation_hint`: `"ERROR -1743: …"` menghasilkan petunjuk Automation.
  - `test_unavailable_raises`: monkeypatch `excel_available` → False.
  - `@pytest.mark.excel test_real_excel_inserts_and_widens_subtotals` (skip kecuali `MATRIX2IM_EXCEL_TESTS=1` dan `excel_available()`):
    - IM sintetis berisi grup 3 baris + subtotal dan grup 1 baris + subtotal, lalu `add_file_sharing`;
    - `plan_inserts` + `run_excel_inserts` dengan `TEST_PASSWORD`;
    - openpyxl memastikan subtotal menjadi `=SUBTOTAL(9,B55:B59)` dan `=SUBTOTAL(9,B60:B62)` (nomor sesuai fixture), baris sisipan berisi NOTE `"catatan"`, dan `check_modify_password` masih lulus.
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q` (test excel ter-skip), lalu `MATRIX2IM_EXCEL_TESTS=1 .venv/bin/pytest -m excel -q`. Hasil yang diharapkan: lulus. Bila muncul pop-up izin, user harus mengizinkan.
- [ ] **Step 5:** Commit `feat(im): penjalan Excel Mac untuk sisip baris`.

### Task 8: Setting dan orkestrasi penulisan (`settings.py`, `im_writer.py`)

**Files:** Create `matrix_to_im/settings.py`, `matrix_to_im/im_writer.py`, `matrix2im.local.example.toml`, `tests/test_settings.py`, `tests/test_im_writer.py`; Modify `.gitignore` (tambah `matrix2im.local.toml`)

**Interfaces:**
- Consumes: semua Task 1–7.
- Produces:
  - `@dataclass(frozen=True) Settings(im_master: Path | None = None, backup_dir: Path | None = None, host: str = "127.0.0.1", port: int = 8765, insert_engine: str = "excel")`
  - `DEFAULT_SETTINGS_FILE = "matrix2im.local.toml"`
  - `load_settings(path: Path | None = None) -> Settings`
  - `@dataclass(frozen=True) Fingerprint(size: int, mtime_ns: int, sha256: str)`
  - `fingerprint(path: Path) -> Fingerprint`
  - `@dataclass WritePreview(im_master: Path, fingerprint: Fingerprint, engine: str, inserts: list[InsertOp], projected: ImLayout, values: ValuePlan, skipped: list[Skip], summary: dict)`
  - `prepare_write(final: dict, settings: Settings, password: str) -> WritePreview`
  - `@dataclass WriteReport(ok: bool, message: str, summary: dict, backup: Path | None, log_file: Path, seconds: float, inserted_rows: list[int])`
  - `commit_write(preview, final, settings, password, user, matrix_name, progress=lambda s: None, inserter=run_excel_inserts) -> WriteReport`

**Format TOML:**

```toml
[ui]
host = "127.0.0.1"
port = 8765
[im]
master = "~/Library/CloudStorage/.../F27 IM MASTER (AI TEST V2).xlsx"
backup_dir = "~/matrix2im-cadangan"
insert_engine = "excel"
```

- Path di-`expanduser`.
- `insert_engine` di luar `excel`/`none` → error.
- File yang disebut eksplisit tapi tidak ada → error. File default yang tidak ada → `Settings()`.

**`prepare_write`:**
1. `im_master` dan `backup_dir` wajib diisi, dan file harus ada.
2. `check_modify_password`.
3. `owner_file` ada → `Matrix2ImError(f"IM MASTER sedang dibuka oleh {nama}. Minta yang bersangkutan menutup file, lalu coba lagi.")`.
4. `fingerprint`, `read_layout`, `style_colorways`.
5. Engine: `excel` bila diminta dan `excel_available()`. Selain itu `none`, dengan peringatan `"Excel tidak tersedia: hanya mengisi baris yang sudah ada"` bila yang diminta `excel`.
6. `plan_inserts` hanya untuk engine `excel`, lalu `project_layout`, lalu `plan_values`. Untuk setiap op grup 1 baris (`op.at_row == op.fill_from + 1`), tambahkan peringatan `"<deskripsi>: grup hanya 1 baris, baris disisipkan di bawahnya; rumus di luar baris subtotal yang merujuk baris itu tidak ikut melebar"` (spec §6).
7. `summary`:

   | Key | Isi |
   | --- | --- |
   | `styles` | jumlah style yang ditulis |
   | `groups` | jumlah grup yang cocok |
   | `cells_changed` | jumlah perubahan sel |
   | `rows_inserted` | Σ count |
   | `groups_inserted` | jumlah op |
   | `changed_colorway_rows` | jumlah baris yang berganti colorway |
   | `short_groups` | jumlah |
   | `skipped` | jumlah |
   | `unmatched_im_styles` | jumlah |
   | `missing_styles` | jumlah |
   | `warnings` | jumlah |
   | `engine` | engine yang dipakai |

**`commit_write`** (urutan wajib; setiap kegagalan sebelum langkah 9 tidak menyentuh file asli):
1. Password dicek ulang.
2. `fingerprint` sama dengan pratinjau, dan file pemilik tidak ada. Bila tidak → batal dengan `"IM MASTER berubah atau sedang dibuka sejak rencana dibuat; buat rencana lagi"`.
3. Cadangan: salin ke `backup_dir / f"{stem} {YYYY-MM-DD HHMMSS} sebelum matrix2im{suffix}"`.
4. Salin ke folder sementara lokal. Bila ada op, jalankan `inserter(tmp, password, ops, IM_SHEET, get_column_letter(columns.last))`.
5. `read_layout(tmp)`: deskripsi, `rows`, dan `subtotal_row` tiap grup harus sama dengan `preview.projected`. Bila tidak → batal.
6. `plan_values` pada layout nyata harus menghasilkan himpunan `(row, col, new)` yang sama dengan `preview.values.changes`. Nilai `old` tidak dibandingkan, karena `fill down` menyesuaikan rumus relatif. Bila tidak sama → batal.
7. `apply_cell_values(tmp → tmp2)`.
8. Verifikasi tmp2:
   - setiap perubahan terbaca sesuai rencana;
   - `check_modify_password(tmp2)` lulus;
   - CRC semua entri selain sheet XML dan workbook.xml sama dengan tmp.
9. `fingerprint` asli sama dengan pratinjau, dan file pemilik tidak ada. Lalu salin isi tmp2 ke file asli di tempat yang sama (`r+b`, `write`, `truncate`, `flush`, `os.fsync`), dan cocokkan sha256.
10. Log: tambahkan satu baris JSON ke `backup_dir / "matrix2im-log.jsonl"` dengan key `waktu, user, matrix, im_master, engine, summary, backup, hasil ("ok"|"gagal"), pesan`. Kegagalan juga dicatat.

Fase `progress`, berurutan: `"cek password"`, `"cadangan"`, `"Excel: sisip baris"`, `"tulis nilai"`, `"verifikasi"`, `"simpan ke IM MASTER"`, `"selesai"`.

- [ ] **Step 1: Tulis test**:

  Inserter palsu di test memakai openpyxl `insert_rows` + salin nilai baris `fill_from` pada file sintetis. Ini boleh di test karena file sintetis tidak punya gambar.

  - `test_load_settings_default_missing_file_gives_defaults`, `test_load_settings_reads_toml_and_expands_user`, `test_invalid_engine_raises`
  - `test_prepare_requires_settings`
  - `test_prepare_wrong_password`
  - `test_prepare_blocks_when_owner_file_present`: pesan mengandung nama.
  - `test_prepare_summary_counts`: satu grup kurang, satu lebih, satu tak cocok.
  - `test_excel_unavailable_falls_back_to_none`
  - `test_commit_engine_none_in_place_backup_log`:
    - `st_ino` file asli sama sebelum dan sesudah;
    - cadangan = byte lama;
    - nilai terbaca;
    - log satu baris tanpa password.
  - `test_commit_with_fake_inserter_extends_groups`
  - `test_commit_aborts_if_file_changed_after_prepare`: file asli tidak ditimpa tool.
  - `test_commit_aborts_if_layout_differs_from_projection`: inserter palsu menyisip jumlah salah, file asli tidak disentuh.
  - `test_commit_aborts_when_owner_file_appears`
  - `test_password_not_in_log_report_or_messages`
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `feat(im): orkestrasi tulis IM MASTER dengan cadangan, verifikasi, log`.

### Task 9: CLI `write` dan `ui --host`

**Files:** Modify `matrix_to_im/cli.py`, `tests/test_cli.py`

**Interfaces:**
- Consumes: `load_settings`, `prepare_write`, `commit_write`, `serve`.
- Di task ini tanda tangan `serve` diubah menjadi `serve(host: str = "127.0.0.1", port: int = 8765, open_browser: bool = True, settings: Settings | None = None)`. `host` langsung dipakai untuk bind (menggantikan `"127.0.0.1"` yang di-hardcode), sedangkan `settings` baru disimpan; Task 10 yang memakainya.
- Produces:
  - `matrix2im write FINAL_JSON [--settings PATH] [--im-master PATH] [--backup-dir PATH] [--engine excel|none] [--user NAME] [--dry-run] [--yes]`
  - `matrix2im ui [--host HOST] [--settings PATH]`

Rincian:
- Opsi CLI menimpa setting.
- Password: `getpass.getpass("Password IM MASTER: ")` bila stdin adalah tty; selain itu baris pertama stdin.
- `--dry-run`: cetak ringkasan, exit 0, file tidak disentuh.
- Tanpa `--yes`: tampilkan ringkasan lalu minta ketik `ya` (tty). Bukan tty dan tanpa `--yes` → error exit 1.
- `--user` default `getpass.getuser()`.
- Ringkasan dan laporan ke stdout, log ke stderr, tanpa traceback.

- [ ] **Step 1: Tulis test**:
  - `test_write_dry_run_prints_summary_and_leaves_file`: password lewat stdin, sha256 file tidak berubah.
  - `test_write_wrong_password_exit_1`
  - `test_write_yes_writes_engine_none`
  - `test_write_non_tty_without_yes_exit_1`
  - `test_ui_passes_host_and_settings` (monkeypatch `serve`)
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `feat(cli): subcommand write dan opsi ui --host`.

### Task 10: Server UI multi-pengguna dan endpoint tulis

**Files:** Modify `matrix_to_im/ui_server.py`, `tests/test_ui_server.py`

**Interfaces:**
- Consumes: `UiState`, `run_pipeline` (sudah ada); `Settings`, `prepare_write`, `commit_write`, `WritePreview` (Task 8).
- Produces:
  - `SessionStore(root: Path, ttl_s: int = 12 * 3600)` dengan method `get(cookie_header: str | None) -> tuple[str, UiState, bool]` dan `cleanup_expired() -> None`.
  - `WriteCoordinator` dengan method `status() -> dict` dan `start(user: str, job: Callable[[Callable[[str], None]], WriteReport]) -> bool`.
  - `make_server(host: str, port: int, sessions: SessionStore, settings: Settings) -> _UiServer`
  - `serve(host: str = "127.0.0.1", port: int = 8765, open_browser: bool = True, settings: Settings | None = None) -> None`

Rincian:
- **Sesi.**
  - Cookie `m2i_session` (`secrets.token_urlsafe(16)`, `HttpOnly; SameSite=Strict; Path=/`) dikirim saat sesi baru.
  - Semua endpoint lama (upload, run, download) memakai `UiState` sesi masing-masing.
  - Sesi yang lewat TTL dibersihkan saat ada request.
- **Validasi header.** Bila `settings.im_master` ada dan user tidak upload IM, `run_pipeline` memakai file setting.
- **Endpoint baru** (body JSON maksimal 64 KB):

  | Endpoint | Perilaku |
  | --- | --- |
  | `GET /api/write/config` | `{enabled, reason, im_name, engine}`. `enabled` false bila setting `im_master`/`backup_dir` kosong. |
  | `POST /api/write/prepare` `{name, password}` | Butuh hasil run di sesi (400 bila belum). Mengembalikan ringkasan, `changes` (maks. 500 pertama + `changes_total`), daftar sisip (deskripsi + count), `skipped`, `short_groups`, `unmatched_im_styles`, `warnings`, `engine`, `token`. `WrongPassword` → 401. File sedang dibuka atau penulisan berjalan → 409. Error lain → 422. |
  | `POST /api/write/commit` `{token, password}` | Token harus milik sesi itu. Password dicek ulang di `commit_write`. Mulai job di thread → 202. Ada job berjalan → 409 `"sedang ada penulisan oleh <user>"`. |
  | `GET /api/write/status` | `{state: "idle"\|"running"\|"done"\|"error", phase, user, started, report?, error?}`. Laporan berisi `backup`, `seconds`, `summary`, `inserted_rows`, `message`. |

- **Password** tidak pernah masuk log (pastikan `log_message` dan log error tidak memuat body).
- **Alamat.** Saat host bukan 127.0.0.1, log menampilkan alamat jaringan, mis. `http://<ip-lan>:8765/` (best effort lewat `socket`).

- [ ] **Step 1: Tulis test** (pola server port 0 yang sudah ada):
  - `test_cookie_set_httponly_samesite`
  - `test_sessions_isolated_uploads`: sesi B tidak melihat upload sesi A, dan `/api/run` sesi B mendapat 400.
  - `test_write_config_disabled_without_settings`
  - `test_prepare_without_result_400`
  - `test_prepare_wrong_password_401`
  - `test_prepare_commit_status_end_to_end_engine_none`: IM sintetis di `tmp_path`, `final` kecil disuntik ke state sesi; status akhir `done` dan IM berubah.
  - `test_second_commit_while_running_409`: job pertama ditahan dengan `threading.Event` lewat monkeypatch `commit_write`.
  - `test_password_not_logged` (caplog)
  - `test_make_server_binds_given_host`
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `feat(ui): sesi per browser dan endpoint tulis IM MASTER`.

### Task 11: Panel "Tulis ke IM MASTER" di UI

**Files:** Modify `matrix_to_im/ui/index.html`, `tests/test_ui_server.py`

**Interfaces:** Consumes endpoint Task 10. Gaya dan pola mengikuti `index.html` yang ada (vanilla JS, token warna, dark mode).

Rincian:
- **Kapan tampil.** Panel muncul setelah hasil proses ada dan `GET /api/write/config` mengembalikan `enabled`. Bila tidak, tampilkan alasannya.
- **Form.** Nama, password (`type=password`, `autocomplete="off"`), dan tombol **Buat rencana**. Password hanya disimpan di variabel JS selama halaman terbuka; tidak pernah di localStorage.
- **Pratinjau.**
  - Ubin ringkasan: style, sel berubah, baris disisipkan, grup disisipi, baris berganti colorway, grup kurang baris, style dilewati, style IM tak ada di MATRIX, peringatan, dan engine.
  - Tabel perubahan sel (DESCRIPTION, sel, lama → baru; 500 pertama + jumlah total).
  - Daftar sisip, daftar dilewati, dan daftar peringatan.
- **Konfirmasi.** Checkbox wajib "Saya sudah memastikan tidak ada yang sedang membuka IM MASTER", lalu tombol **Tulis semua ke IM**.
- **Saat menulis.** Polling `/api/write/status` tiap 2 detik dan tampilkan fase. Tombol dinonaktifkan selama berjalan.
- **Laporan akhir.** Tampilkan laporan dengan path cadangan, atau error dengan pesan dan lokasi cadangan.
- **Status awal.** Bila status `running` milik orang lain, tampilkan `"Sedang ada penulisan oleh <user>"`.

- [ ] **Step 1: Tulis test** `test_index_has_write_panel`: `GET /` memuat `"Tulis semua ke IM"`, `id="write-password"`, dan `/api/write/prepare`.
- [ ] **Step 2:** Jalankan test. Hasil yang diharapkan: FAIL.
- [ ] **Step 3:** Implementasikan.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: lulus.
- [ ] **Step 5:** Cek manual di browser memakai IM sintetis + setting sementara dengan `insert_engine = "none"`: buat rencana, tulis, lihat laporan. Lakukan di light dan dark mode.
- [ ] **Step 6:** Commit `feat(ui): panel tulis ke IM MASTER`.

### Task 12: Regression F27, dokumentasi, dan uji terima

**Files:** Create `tests/regression/test_write_f27.py`, `docs/uji-terima-tulis-im.md`; Modify `tests/conftest.py`, `tests/regression/conftest.py`, `README.md`, `docs/PRD_MATRIX_MASTER_to_IM_JSON.md`

**Interfaces:**
- `tests/conftest.py` mendapat kind `"im_write": ("MATRIX2IM_IM_WRITE", "F27_IM_MASTER_WRITE.xlsx", "*IM*MASTER*AI*TEST*.xlsx")`.
- Fixture `im_write_path` (skip dengan pesan bila tidak ada).
- Password untuk test Excel diambil dari env `MATRIX2IM_IM_PASSWORD` (tidak pernah ditulis di repo).

Test regression (salinan ke `tmp_path`, golden final):
- `test_style_colorways_f27`: 211 style, 0 skip; HUTTON 8 colorway dengan GLASGOW dan FAWN `"000 & SY"`; CROP BOYFRIEND CREW COTTON 5 colorway.
- `test_layout_and_insert_plan_f27`: perkiraan spike adalah 177 op, Σ count 286, dan 27 op untuk grup 1 baris. **Bila angkanya berbeda, cari grup penyebabnya dan jelaskan di laporan task sebelum mengganti angka di test.**
- `test_value_plan_f27_runs_fast`: `read_layout` + rencana < 60 detik; tidak ada `missing_styles`; `unmatched_im_styles` = 179 (perkiraan spike, dengan aturan yang sama seperti di atas).
- `test_apply_values_on_copy_keeps_parts`: tanpa sisip (engine none). Semua entri selain sheet XML dan workbook.xml byte-identik, kunci lulus `check_modify_password` (skip bila env password tidak ada), dan waktu < 30 detik.
- `@pytest.mark.excel @pytest.mark.regression test_full_write_on_copy` (butuh `MATRIX2IM_EXCEL_TESTS=1` dan `MATRIX2IM_IM_PASSWORD`). `prepare_write` + `commit_write` dijalankan pada salinan (`settings.im_master` = salinan, `backup_dir` = `tmp_path`). Yang dicek:
  - `report.ok`;
  - jumlah entri `xl/media/*` dan `xl/drawings/drawing*.xml` sama dengan file sumber;
  - jumlah `<conditionalFormatting` dan `<cfRule` sama;
  - grup BOYFRIEND RAGLAN CREW CHUNKY COTTON - X/L (grup 1 baris di baris 229 file sumber) punya subtotal ber-range;
  - total waktu < 5 menit.

Dokumentasi:
- **README**, bagian "Menulis ke IM MASTER":
  - alur di UI dan CLI `write`;
  - setting `matrix2im.local.toml`;
  - panduan PC pusat: izin Google Drive dan Automation di System Settings, Mac tidak sleep, firewall, `matrix2im ui --host 0.0.0.0`;
  - cara memulihkan dari cadangan;
  - batasan (celah file sudah dibuka sebelum tulis; rujukan dari file lain).
- **PRD**: entri "Riwayat perubahan" untuk fase 2, dan bagian "Fase berikutnya" diperbarui dengan rujukan ke spec.
- **`docs/uji-terima-tulis-im.md`**, daftar centang untuk user:
  1. Tulis ke **salinan uji** di folder Google Drive (bukan file asli). Cek link dan riwayat versi di Google Drive tetap sama (spec §10.3).
  2. Buka salinan uji di Excel di PC lain. Cek apakah tool di PC pusat menolak dengan nama orang itu (spec §10.4).
  3. Buka UI dari PC Windows lewat browser.
  4. Bandingkan hasil dengan file asli memakai cara user sendiri.
  5. Coba dua browser menulis bersamaan.

- [ ] **Step 1:** Tulis test regression.
- [ ] **Step 2:** `.venv/bin/pytest -m regression -q`. Hasil yang diharapkan: lulus (test excel ter-skip). Lalu `MATRIX2IM_EXCEL_TESTS=1 MATRIX2IM_IM_PASSWORD=… .venv/bin/pytest -m excel -q`, dijalankan oleh controller bersama user. Hasil yang diharapkan: lulus.
- [ ] **Step 3:** Tulis dokumentasi.
- [ ] **Step 4:** `.venv/bin/pytest -q`. Hasil yang diharapkan: semua lulus.
- [ ] **Step 5:** Commit `docs: panduan tulis IM MASTER, regression F27, uji terima`.
