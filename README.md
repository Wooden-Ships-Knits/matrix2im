# matrix2im

Mengubah file Excel **F27 MATRIX MASTER** menjadi JSON untuk **IM MASTER**.

- Masukan: file MATRIX `.xlsx`.
- Keluaran: `raw.json` (hasil baca MATRIX apa adanya, per style) dan `final.json` (satu record per colorway, dengan key kolom IM MASTER).
- Fase 2: menulis `final.json` ke IM MASTER (Excel) yang dibagikan, lewat UI atau `matrix2im write` (lihat bagian "Menulis ke IM MASTER"). File MATRIX tidak pernah diubah; IM MASTER hanya diubah saat Anda menekan tombol tulis (atau `write` tanpa `--dry-run`).

## Instalasi

Butuh Python 3.11+.

```bash
python3.12 -m venv .venv
.venv/bin/pip install -e '.[dev]'
```

Pillow sengaja tidak dipasang. Tanpa Pillow, openpyxl melewati 220 foto di MATRIX sehingga pembacaan tetap cepat (`run` pada F27 sekitar 2 detik).

## Mencoba lewat UI (browser)

```bash
.venv/bin/matrix2im ui
```

Perintah ini membuka halaman `http://127.0.0.1:8765/` di browser (pakai `--port` untuk port lain, `--no-browser` agar tidak membuka browser otomatis). Di halaman itu:

1. Pilih atau seret file **MATRIX MASTER** (wajib) dan, bila ingin mencocokkan nama kolom, **IM MASTER** (opsional).
2. Klik **Proses**. Tahap 1, Tahap 2, dan validasi dijalankan sama seperti `matrix2im run`.
3. Lihat hasilnya: ringkasan, status validasi, tab **Records** (bisa dicari dan difilter per WHS / perlu review), **Needs review** (per kode, dengan comment sel MATRIX), **Per style** (colorway dikelompokkan menurut garis hitam/merah, dengan warna font legend), dan **Warning**. Tombol **raw.json** / **final.json** mengunduh hasilnya.

Secara default server hanya bisa diakses dari komputer ini (127.0.0.1); untuk PC pusat lihat "Menulis ke IM MASTER". File yang diupload disimpan sebagai salinan di folder sementara dan dihapus saat server dihentikan (Ctrl+C); file asli tidak disentuh. Berbeda dengan `run`, UI tetap menampilkan hasil walau validasi gagal (errornya ditampilkan di atas), supaya datanya bisa diperiksa.

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
- File MATRIX harus disimpan oleh Excel, supaya nilai hasil formula ikut tersimpan (nomor style berupa formula, mis. `=A9+1`). File yang disimpan oleh tool yang tidak menghitung formula (mis. skrip openpyxl) tidak punya nilai itu; `extract`/`run` lalu berhenti dengan exit 1 dan menyebut contoh selnya. Buka dan simpan ulang file di Excel, lalu jalankan lagi.
- `raw.json` ditulis segera setelah Tahap 1 berhasil. `final.json` hanya ditulis bila validasi lolos (jumlah colorway = jumlah record, key record ada di header IM, WHS hanya tiga nilai, tidak ada nilai tebakan, dan tiap nilai record berupa `null` atau teks tidak kosong tanpa spasi di awal/akhir yang tidak diawali `=`, sehingga string kosong, spasi, atau teks mirip formula ditolak). Bila validasi (atau bagian lain Tahap 2) gagal, tiap error dicatat, `final.json` tidak ditulis, dan `final.json` lama di path tujuan dihapus (dicatat sebagai WARNING) supaya tidak tertinggal di samping `raw.json` baru.
- Ringkasan (jumlah style, record, WHS per nilai, WS TAG COLOR terisi/null, needs_review per kode) dicetak ke stdout. Log (progres, jumlah warning, error) ke stderr. Folder output dibuat otomatis.
- JSON ditulis UTF-8, `ensure_ascii=False`, indent 2.

### Exit code

| Kode | Arti |
| --- | --- |
| 0 | Sukses (item `needs_review` tidak membuat gagal) |
| 1 | Error fatal: argumen salah/kurang (mis. `run` tanpa `--im-master`), file tidak terbaca, tidak ada halaman/kartu, raw JSON rusak/tidak lengkap, validasi gagal, error tak terduga. Tidak ada traceback; pesan di stderr |
| 2 | Hanya dengan `--strict` dan `needs_review` tidak kosong (file tetap ditulis) |

## Menulis ke IM MASTER

Aturan penulisan lengkap ada di `docs/superpowers/specs/2026-11-10-tulis-im-master-design.md` (§6 aturan, §7 pengaman, §8 PC pusat). Daftar uji terima untuk TIM IM: `docs/uji-terima-tulis-im.md`.

Ringkasnya: tool mengisi COLOR1..n, WS TAG COLOR, dan WHS di baris colorway tiap grup size IM MASTER, menyisipkan baris bila colorway lebih banyak dari baris, dan tidak menyentuh kolom rumus. Warna latar WS TAG COLOR dan WHS mengikuti WHS (000 = FF99FF, SY = FFCC99, BOTH WHS = 969696). Bila baris grup lebih banyak dari colorway, di baris lebih itu WS TAG COLOR dan WHS dikosongkan dan berlatar kuning (FFFF00), dan COLOR dibiarkan. Sebelum menulis, tool mencocokkan password, membuat cadangan, dan memverifikasi hasilnya.

### Setting `matrix2im.local.toml`

Salin contoh di bawah ke file `matrix2im.local.toml` di folder tempat perintah dijalankan (atau tunjuk file lain dengan `--settings`). File ini berisi path lokal, tidak di-commit, dan **tidak boleh berisi password**.

```toml
[ui]
host = "127.0.0.1"      # "0.0.0.0" agar bisa dibuka dari PC lain di jaringan kantor
port = 8765

[im]
master = "/path/ke/F27 IM MASTER.xlsx"   # file IM MASTER yang ditulis
backup_dir = "/path/ke/folder cadangan"  # cadangan, folder kerja, dan log
insert_engine = "excel"                  # "excel" (sisip baris lewat Excel) atau "none" (hanya isi baris yang ada)
```

Dengan `insert_engine = "none"` (atau bila Excel tidak terpasang), tool hanya mengisi baris yang sudah ada; grup yang kekurangan baris didaftar sebagai peringatan untuk ditambah manual.

### Lewat UI

1. Jalankan `.venv/bin/matrix2im ui` (di PC pusat: `--host 0.0.0.0`, lihat di bawah), upload MATRIX, klik **Proses**, lalu periksa hasilnya. IM MASTER tidak perlu diupload: penulisan selalu memakai file `master` di `matrix2im.local.toml` (header file itu juga dipakai untuk validasi bila tidak ada upload IM). Kotak upload IM MASTER opsional dan hanya untuk validasi header (fase 1); file yang diupload di sana tidak pernah ditulis.
2. Di panel **Tulis ke IM MASTER**, isi **Nama Anda** (masuk log) dan **Password IM MASTER**, lalu klik **Buat rencana**. Belum ada yang ditulis pada tahap ini.
3. Periksa pratinjau: jumlah sel dan baris yang berubah, daftar sel (nilai lama dan baru), baris yang disisipkan, style yang dilewati, style MATRIX yang tidak ada di IM (tidak ditulis sama sekali), dan peringatan. Tabel di layar dibatasi 500 sel; klik **Unduh semua perubahan (Excel)** untuk daftar lengkap dalam file .xlsx (isi di "Laporan Excel" di bawah).
4. Centang pernyataan bahwa tidak ada yang sedang membuka IM MASTER, lalu klik **Tulis semua ke IM**. Tunggu laporan; satu penulisan penuh F27 memakan waktu beberapa menit.

Hanya satu orang yang bisa menulis pada satu waktu; sesi lain melihat pesan bahwa penulisan sedang berjalan.

### Lewat CLI

```bash
# Lihat rencana saja; IM MASTER tidak diubah
.venv/bin/matrix2im write out/final.json --dry-run

# Tulis (menanyakan password dan konfirmasi 'ya')
.venv/bin/matrix2im write out/final.json --user "Nama Anda"
```

- Password dibaca dari prompt (atau baris pertama stdin), tidak pernah dari argumen perintah. Tanpa terminal, `--yes` wajib.
- `--report rencana.xlsx` menulis laporan Excel semua perubahan rencana (juga dengan `--dry-run`), sebelum konfirmasi. Bila nama file tidak berakhiran `.xlsx`, akhiran itu ditambahkan; path akhirnya dicetak.
- Opsi `--settings`, `--im-master`, `--backup-dir`, dan `--engine` menimpa isi `matrix2im.local.toml`. Exit code 0 bila berhasil, 1 bila gagal atau dibatalkan.

### Menyiapkan PC pusat

Langkah ini dikerjakan sendiri oleh pengelola (menyangkut pengaturan sistem), bukan oleh tool:

- **Google Drive:** beri aplikasi yang menjalankan server (Terminal/Python) izin ke folder Google Drive di System Settings, Privacy & Security, Files and Folders (atau Full Disk Access). Tanpa izin ini, proses dan rencana tulis berhenti dengan pesan "Beri izin akses folder Google Drive ...".
- **Automation:** saat pop-up pertama "Terminal ingin mengontrol Microsoft Excel" muncul, pilih izinkan (System Settings, Privacy & Security, Automation).
- **Folder kerja:** Excel memakai salinan sementara di `<backup_dir>/kerja`. Pertama kali, macOS bisa bertanya apakah Excel boleh membuka file di sana; izinkan sekali saja.
- **Tidak sleep:** atur Mac supaya tidak tidur (System Settings, Energy / Lock Screen).
- **Firewall:** izinkan koneksi masuk untuk Python di System Settings, Network, Firewall, untuk port server (default 8765).
- **Jalankan server:** `.venv/bin/matrix2im ui --host 0.0.0.0 --no-browser`. Anggota TIM membuka `http://<alamat IP Mac>:8765/` dari browser (Mac atau Windows); alamat IP ada di System Settings, Network, atau lewat `ipconfig getifaddr en0`.
- **Saat menulis:** Excel di PC pusat membuka salinan kerja bernama unik (`im-kerja-<kode>.xlsx`) di folder `kerja`, lalu menyimpan dan menutupnya sendiri. Tool hanya menyimpan atau menutup workbook yang path lengkapnya (*full name*) sama dengan salinan kerja itu; workbook lain yang sedang terbuka di Excel tidak disentuh. Bila Excel sudah membuka workbook lain dengan nama yang sama, penulisan berhenti sebelum ada yang diubah ("Excel sudah membuka workbook bernama ...; tutup dulu"). Bila langkah Excel gagal atau macet, tool hanya menutup salinan kerjanya sendiri, tanpa menyimpan. Selama otomatisasi, peringatan (alert) Excel dimatikan dan dinyalakan lagi sesudahnya, supaya dialog tak terduga tidak membuat Excel dan server macet. Jangan bekerja di Excel di PC itu selama penulisan.
- **Pemeriksaan sesudah Excel:** sebelum nilai ditulis, tool membandingkan salinan kerja hasil Excel dengan cadangan. Bila atribut kunci `fileSharing` (`userName`, `readOnlyRecommended`, `algorithmName`, `spinCount`, `hashValue`, `saltValue`) atau daftar nama file gambar, drawing, comment, link eksternal, dan embedding di dalam xlsx berbeda, penulisan dibatalkan dan file asli tidak diubah. Jumlah blok conditional formatting yang berubah hanya menjadi peringatan di laporan. Tata letak grup dan rencana nilai juga diperiksa ulang pada file hasil Excel. Isi gambar/drawing dan rumus di luar baris subtotal tidak dibandingkan.

**Password di jaringan kantor.** UI memakai HTTP biasa, jadi password melewati jaringan kantor tanpa enkripsi. Pakai hanya di LAN kantor dan jangan pernah meneruskan (port forward) port ini ke internet. Tool tidak pernah menyimpan password: tidak di disk, log, laporan, maupun argumen perintah.

### Cadangan, log, dan pemulihan

- Tiap penulisan lebih dulu membuat cadangan `<nama> <tanggal jam> sebelum matrix2im.xlsx` di `backup_dir`. Folder `kerja` di sana hanya berisi salinan sementara dan dibersihkan sendiri.
- Log berupa file `matrix2im-log.jsonl` di `backup_dir`, satu baris JSON per penulisan: waktu, nama user, nama file MATRIX yang diupload (di UI) atau nama file `final.json` (di CLI), ringkasan rencana, hasil (ok/gagal), pesan, path cadangan, dan path laporan Excel (`laporan`). Password tidak pernah dicatat. Penulisan yang gagal pun dicatat.
- **Laporan Excel:** tiap penulisan (berhasil maupun gagal) menyimpan `<nama> <tanggal jam> laporan matrix2im.xlsx` di `backup_dir`, sepasang dengan cadangannya (`... sebelum matrix2im.xlsx`; jangan tertukar saat memulihkan); path-nya tampil di laporan UI dan CLI. File ini workbook baru dengan satu sheet `Laporan`, berkolom `DESCRIPTION, sel, baris, kolom, lama, baru, jenis`, dan terbuka benar di Excel Mac maupun Windows (termasuk setelan regional Indonesia). `jenis` berisi `ubah` (perubahan sel), `warna` (warna latar WS TAG COLOR/WHS lama dan baru), `sisip` (baris sisipan, nomor barisnya di file hasil), `baris lebih` (baris melebihi colorway MATRIX; WS TAG COLOR dan WHS dikosongkan, latar kuning), `baris colorway berubah` (benang lama dan baru), `dilewati`, `tidak ada di IM` (style MATRIX yang tidak ditulis), `IM tanpa MATRIX`, atau `peringatan`; untuk empat jenis terakhir kolom `baru` berisi keterangan. Bila penulisan gagal, alasannya ada di baris `peringatan`. Laporan yang sama bisa diunduh dari pratinjau UI atau ditulis lewat `--report`.
- Selama menulis, tool memegang kunci file `matrix2im.lock` di `backup_dir`, supaya dua proses di PC yang sama (mis. CLI dan server) tidak menulis bersamaan; proses kedua langsung berhenti dengan pesan "Ada penulisan lain yang sedang berjalan (proses lain di PC ini)" tanpa menyentuh apa pun. Bila penulisan dihentikan paksa (Ctrl+C) saat isi IM MASTER sedang ditimpa, tool memulihkan isinya dari cadangan, memeriksa sha256-nya, dan mencatatnya di log.
- **Memulihkan:** tutup IM MASTER di semua PC, lalu salin isi file cadangan menimpa file IM MASTER (isi ditimpa, bukan file diganti, supaya link Google Drive tetap). Di Google Drive, lebih baik pakai riwayat versi (klik kanan file, Kelola versi).

### Batasan

- **Celah file sudah dibuka:** pengecekan "file sedang dibuka" bergantung pada file pemilik `~$` buatan Excel. Apakah `~$` dari PC lain terlihat lewat Google Drive masih harus diuji (uji terima nomor 2, spec §10.4). Sampai terkonfirmasi, TIM IM tetap berkoordinasi manual. Bila seseorang sudah membuka IM MASTER sebelum tool menulis lalu menyimpan sesudahnya, simpanannya bisa menimpa hasil tool; cadangan dan log adalah jalan pemulihan.
- **Rujukan dari file lain:** file lain yang merujuk ke sel IM MASTER diperlakukan seperti saat TIM menyisipkan baris manual; tidak ada penanganan khusus. Laporan menyebut baris yang disisipkan.
- Grup yang hanya 1 baris dan disisipi baris: rumus di luar baris subtotal yang merujuk baris itu tidak ikut melebar (ada peringatan di pratinjau).
- Baris subtotal dikenali dari rumus `SUBTOTAL(` di barisnya, bukan dari tulisan labelnya (label di IM asli bervariasi). Baris tanpa rumus SUBTOTAL dianggap baris colorway. Kunci grup adalah `STYLE - SIZE` (dua segmen pertama DESCRIPTION); teks sesudahnya, mis. `- ALT 1`, dianggap catatan dan tidak diubah.
- Style IM yang tidak ada di MATRIX tidak disentuh. Label subtotal tidak diubah.

## Skema keluaran

**raw.json**: `source`, `total_styles` (jumlah kartu), `styles`, `warnings`. Tiap style: `no`, `style`, `style_whs`, `tag`, `sheet`, `section`, dan tiga daftar colorway menurut garis: `color_above_black_line`, `color_between_black_and_red_line`, `color_below_red_line`. Tiap colorway: `colorway`, `yarns`, `font_color`, `cell`, `comment`.

**final.json**: `target`, `source`, `generated`, `rules`, `summary`, `records`, `needs_review`.

- `records`: satu per colorway, key persis kolom IM: `DESCRIPTION`, `COLOR1`..`COLOR18` (di IM kolom ke-4 bernama `COLOR 4`), `WS TAG COLOR`, `WHS`. `_ref` berisi asal di MATRIX (sheet, No., sel, posisi terhadap garis); tidak ditulis ke IM. WS TAG COLOR yang tidak bisa diputuskan aturan diisi teks `Need Review` (ikut ditulis ke IM supaya terlihat tim); WHS yang tidak bisa diputuskan bernilai `null` (tidak ditulis). Style yang muncul di dua sheet dicatat sebagai record terpisah.
- `needs_review`: item `{record, code, DESCRIPTION, colorway, issue, _ref, ...}`; `record` berupa index record, atau `null` untuk warning ekstraksi.
- `summary`: `total_records`, `total_styles`, `WHS` per nilai, `WS_TAG_COLOR_filled` / `WS_TAG_COLOR_need_review` / `WS_TAG_COLOR_null`, `needs_review_items`, `needs_review_by_issue`.

Kode `needs_review`:

| Kode | Arti |
| --- | --- |
| `TAG_CONDITIONAL` | TAG bersyarat atau singkatan; WS TAG COLOR `Need Review` |
| `TAG_CN_OUT_OF_RANGE` | TAG meminta Cn melebihi jumlah benang; WS TAG COLOR `Need Review` |
| `TAG_EMPTY` | TAG kosong; WS TAG COLOR `Need Review` |
| `TAG_UNRESOLVED` | TAG masih berisi instruksi belum diterjemahkan (`C(car)`, `GONDOLA color`); WS TAG COLOR `Need Review` |
| `WHS_STYLE_UNKNOWN` | warna background nomor di luar legend; dipakai WHS warna (putih/tanpa warna = `BOTH WHS`, bukan masalah) |
| `WHS_CONFLICT` | WHS colorway dan style tidak beririsan; WHS `null` |
| `FONT_COLOR_NO_RULE` | warna font (mis. YR / HOLD / DISCONTINUED) belum punya aturan WHS; WHS `null` |
| `COLOR_IS_DESCRIPTOR` | COLOR berupa deskripsi (STRIPE/MARL/MULTI), ditulis apa adanya; WS TAG COLOR `Need Review`; benang asli ada di comment sel |
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

Test tulis (`tests/regression/test_write_f27.py`) memakai salinan uji IM MASTER lewat `MATRIX2IM_IM_WRITE` (pola `data/*IM*MASTER*AI*TEST*.xlsx`); file itu hanya dibaca dan hasil tulis selalu ke salinan di folder sementara. Test `test_full_write_on_copy` menjalankan Microsoft Excel sungguhan dan hanya jalan bila `MATRIX2IM_EXCEL_TESTS=1` dan `MATRIX2IM_IM_PASSWORD` (password ubah salinan uji, diberikan lewat environment, tidak pernah ditulis di repo) diisi: `.venv/bin/pytest -m excel`.

## Open issues (belum diputuskan)

Kasus berikut sengaja tetap `Need Review` / `null` dan masuk `needs_review`; tidak ada aturan yang ditebak:

- Aturan definitif WS TAG COLOR untuk TAG bersyarat atau singkatan.
- TAG yang meminta Cn melebihi jumlah benang.
- TAG yang masih berisi instruksi belum diterjemahkan (TREE FARM `C1/C(car)`, TELLURIDE `C1/GONDOLA color`).
- COLOR berupa deskripsi stripe/marl (benang asli ada di comment sel).
- Aturan WHS untuk font YR / HOLD / DISCONTINUED.
- Arti garis hitam (tidak memengaruhi WHS).

Sudah diputuskan pada review 10 November 2026 (lihat Riwayat perubahan di PRD): style di dua sheet dicatat terpisah; WHS yang dibatasi style bukan masalah; background nomor putih/kosong = `BOTH WHS`; karakter nyasar dibuang tanpa warning.
