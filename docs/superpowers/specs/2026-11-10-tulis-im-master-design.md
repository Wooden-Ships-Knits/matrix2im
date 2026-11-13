# Rancangan: menulis hasil MATRIX ke IM MASTER bersama (fase 2)

Tanggal: 10 November 2026. Status: draf, sudah melalui review builder dan jawaban user, menunggu persetujuan.

Dokumen ini melanjutkan `docs/PRD_MATRIX_MASTER_to_IM_JSON.md`. PRD mengatur cara membaca MATRIX dan membentuk `final.json`. Dokumen ini mengatur cara menulis `final.json` ke file IM MASTER yang dipakai bersama oleh TIM IM.

Penanda di dokumen ini:

- **[Keputusan user]**: diputuskan oleh user dalam percakapan 10 November 2026.
- **[Keputusan perancang]**: diputuskan oleh perancang (Claude). Builder boleh mengikutinya, dan user diminta memeriksanya.
- **[Terbuka]**: belum diputuskan. Tool tidak boleh menebak; bagian ini dibiarkan apa adanya sampai diputuskan.

## 1. Latar belakang dan tujuan

Saat ini `matrix2im` hanya menghasilkan JSON. TIM IM masih memindahkan colorway, WS TAG COLOR, dan WHS dari MATRIX ke IM MASTER dengan tangan.

IM MASTER adalah satu file Excel di Google Shared Drive (`PTIF SERVER/Collection/27 FALL/IM/F27 IM MASTER (AI TEST V2).xlsx`). Setiap PC membuka file itu lewat Google Drive for Desktop. Path-nya berbeda di tiap PC, dan PC TIM IM ada yang Mac dan ada yang Windows. File dikunci dengan **Password to modify**. Selama ini TIM IM hanya mengandalkan koordinasi supaya tidak ada dua orang yang mengedit bersamaan.

Tujuan: setiap anggota TIM IM, dari PC mana pun, bisa memproses MATRIX dan menulis hasilnya ke IM MASTER dengan satu tombol, tanpa merusak isi file dan tanpa diam-diam menimpa pekerjaan orang lain.

## 2. Cara pemakaian

- **[Keputusan user]** PC Mac ini untuk sementara menjadi PC pusat yang selalu menyala. Hanya PC pusat yang menjalankan tool dan menulis ke IM MASTER.
- **[Keputusan user]** Anggota TIM IM membuka UI lewat browser ke alamat PC pusat di jaringan kantor. PC mereka, baik Mac maupun Windows, tidak perlu memasang apa pun.
- **[Keputusan user]** Hanya TIM IM yang mengedit IM MASTER. Password to modify berfungsi sebagai izin menulis.
- **[Keputusan user]** Satu tombol **"Tulis semua ke IM"** menulis semua style hasil proses sekaligus. Tidak ada pemilihan per style.

Alurnya:

1. User membuka UI, upload MATRIX, lalu klik Proses. Ini sudah ada sekarang.
2. User memeriksa hasilnya (Records, Needs review, Per style). Ini juga sudah ada.
3. User klik **Tulis semua ke IM**, lalu mengisi nama dan password IM MASTER.
4. Tool menampilkan ringkasan rencana perubahan dan peringatan, lalu meminta konfirmasi.
5. Tool menulis, memverifikasi, dan menampilkan laporan.

## 3. Sasaran

**Goals**

- Menulis `final.json` ke sheet F27 IM MASTER sesuai aturan penulisan di PRD, ditambah aturan di §6.
- Isi file selain sel tujuan tetap utuh: gambar, drawing, comment, conditional formatting, link ke file lain, kunci password, dan Read-only recommended.
- Bisa dipakai dari PC Mac maupun Windows lewat browser.
- Sebelum menulis selalu ada cadangan, dan setiap penulisan tercatat di log.

**Non-goals**

- Menjalankan penulis di Windows. Yang dibangun sekarang hanya penjalan Excel untuk Mac; rancangannya disiapkan supaya penjalan Windows bisa ditambah nanti.
- Menghapus baris IM untuk style yang tidak ada di MATRIX.
- Menulis ke sheet selain F27.
- Menyelesaikan open issue aturan WS TAG COLOR. Record seperti itu tetap ditulis `Need Review`.
- Login per user. Jaringan kantor dianggap tepercaya; penulisan dilindungi password IM MASTER.

**Anti-goals.** Kalau salah satu terjadi, fiturnya dianggap gagal walaupun "berjalan":

- IM MASTER rusak atau kehilangan isi. Sudah terbukti menyimpan ulang lewat openpyxl menghapus 22 gambar dan 7 aturan conditional formatting (§4).
- Kunci Password to modify hilang, atau tool menulis tanpa password yang benar.
- Tool menulis nilai tebakan.
- Pekerjaan orang lain tertimpa diam-diam tanpa cadangan dan tanpa peringatan.
- File di Google Drive berganti identitas sehingga link dan riwayat versinya putus.

## 4. Fakta dari uji coba (10 November 2026, di salinan file)

Semua uji dilakukan di salinan. File di Google Drive tidak disentuh.

| Uji | Hasil |
| --- | --- |
| Jenis kunci | `<fileSharing readOnlyRecommended="1" userName="Dev-Putri" algorithmName="SHA-512" … spinCount="100000"/>`, artinya Password to modify. Tidak dienkripsi, tidak ada proteksi sheet atau workbook. |
| Cek password | Hash SHA-512 standar ECMA-376 (salt + password UTF-16LE, lalu 100.000 putaran hash + nomor putaran 4 byte little-endian) cocok dengan password dari user. |
| Simpan ulang via openpyxl | Gambar 22→0, drawing 5→0, conditional formatting 249→242, Read-only recommended hilang, printer settings hilang. **Tidak boleh dipakai.** |
| Ubah sel langsung di XML (cara B) | Sekitar 2 detik, hanya `sheet1.xml` yang berubah, kunci tetap ada, Excel membukanya normal. Tidak bisa menyisipkan baris dengan aman. |
| Lewat Excel + AppleScript (cara A) | Buka sekitar 9 detik, tulis + simpan 3 sel sekitar 3 detik. Semua bagian file utuh, kunci tetap ada. Butuh izin macOS sekali (Terminal mengontrol Excel, akses file). |
| Akses Google Drive dari Terminal | Ditolak macOS ("Operation not permitted") sampai user memberi izin di System Settings. |
| Pencocokan nama | Ke-211 style MATRIX F27 punya baris di IM bila spasi berlebih diabaikan (IM memakai mis. `REEGAN CABLE CREW POODLE  - X/S`). Ada 179 style di IM yang tidak ada di MATRIX. |
| Sisip baris di dalam grup (Excel) | Sisip di atas baris colorway terakhir, lalu `fill down` dari baris di atasnya. SUBTOTAL di baris subtotal ikut melebar (`SUBTOTAL(1,CG55:CG57)` → `CG55:CG58`), dan rumus baris baru ikut menyesuaikan (HM58 → `HJ58+HL58`). |
| Kecepatan sisip | Sisip baris utuh sekitar 0,25 detik. `fill down` yang dibatasi ke kolom terpakai (A:IJ) sekitar 0,3 detik per grup, termasuk sisip (50 grup dalam 16 detik). `copy range` satu baris sekitar 2 detik dan memecah conditional formatting (249 → 1.153 blok setelah 50 kali); `fill down` terbatas tidak mengubah CF (248 blok, 328 aturan). Sisip pada range kolom (bukan baris utuh) sekitar 3,4 detik. **Pakai: sisip baris utuh + `fill down` terbatas.** |
| Grup 1 baris (27 grup kurang baris di F27) | Sisip di atas baris terakhir tidak bisa, jadi baris disisipkan di bawah baris colorway lalu `fill down`. Rumus subtotal yang merujuk baris itu (mis. `=SUBTOTAL(1,CG229)`, 12 rumus) diperlebar lewat Excel menjadi `CG229:CG231`. Terbukti di salinan; bagian file lain tidak berubah. |
| Kebutuhan sisip F27 | 1.108 subgrup size milik style MATRIX: 176 pas, 755 kelebihan baris, 177 kurang (total 286 baris). Perkiraan sisip penuh sekitar 1 menit. |
| Menulis nilai lewat Excel | Sel tunggal sekitar 50 ms per sel (300 sel = 15 detik). Blok vertikal 300 sel sekitar 1 detik, tapi kolom COLOR tidak bersebelahan (diselingi kolom ENDS/Y% berisi rumus), jadi penulisan penuh lewat Excel diperkirakan 15–20 menit. |
| Rumus yang bergantung pada kolom tujuan | Dari 146.156 rumus, hanya 1 yang merujuk kolom COLOR/WS TAG COLOR/WHS (`3.1/AJ8749`). |

Perintah AppleScript yang terbukti jalan:

```applescript
tell application "Microsoft Excel"
    open workbook workbook file name <path HFS> write reserved password <pw> ¬
        update links do not update links ignore read only recommended true
    set value of range "CD57" of worksheet "F27" of active workbook to "…"
    save active workbook
    close active workbook saving no
end tell
```

Parameter `read only true` pada `open workbook` menghasilkan "Parameter error (-50)", jadi jangan dipakai.

**[Diperbarui 11 November 2026, saat membangun]** Contoh di atas hanya catatan spike. Penjalan yang dibangun tidak memakai `active workbook`: setelah membuka, skrip mencari workbook yang `full name`-nya sama dengan path file kerja dan hanya menyentuh workbook itu. Selama otomatisasi, alert Excel dimatikan (`display alerts false`), lalu dinyalakan lagi. Alasannya, saat menyimpan setelah ratusan sisip baris, Excel menampilkan dialog yang memblokir PC pusat. Hasil Excel lalu diverifikasi terhadap cadangan.

## 5. Arsitektur

**[Keputusan perancang]** Penulisan dibagi menjadi dua lapis, supaya logika bisa dites tanpa Excel dan penjalannya bisa diganti nanti:

1. **Perencana (Python murni, bisa dites unit).**
   - Masukan: `final.json` dan isi IM MASTER yang dibaca dengan openpyxl `read_only` (hanya membaca, tidak pernah `save()`).
   - Keluaran: *rencana perubahan* berbentuk data biasa:
     - daftar sisip baris per grup: di baris mana, berapa, dan menyalin baris mana;
     - daftar nilai sel yang ditulis setelah penyisipan;
     - style yang dilewati beserta alasannya;
     - peringatan.
   - Rencana bisa ditampilkan di UI sebelum ditulis, dan disimpan ke log.
2. **Penjalan (executor), gabungan dua langkah** (diputuskan dari hasil spike §4):
   1. *Struktur lewat Excel di Mac* (AppleScript via `osascript`, tanpa dependensi Python baru). Excel hanya mengerjakan sisip baris, `fill down`, dan pelebaran rumus subtotal untuk grup 1 baris, lalu menyimpan. Hanya Excel yang menggeser rumus, SUBTOTAL, conditional formatting, comment, dan drawing dengan benar.
   2. *Nilai lewat XML sheet (cara B)*. Setelah Excel menyimpan, Python membaca ulang tata letak IM (nomor baris sudah bergeser), menyusun rencana nilai, lalu mengubah hanya sel COLOR/WS TAG COLOR/WHS di XML sheet. Bagian file lain tidak disentuh. Supaya Excel menghitung ulang saat file dibuka, workbook diberi `fullCalcOnLoad="1"`.
   - Bila tidak ada grup yang kurang baris, langkah Excel dilewati.
   - Bila Excel tidak tersedia atau gagal, penulisan memakai rencana cadangan di §10.
   - Antarmuka langkah struktur cukup umum supaya penjalan Windows (Excel COM) bisa ditambahkan nanti tanpa mengubah perencana.

**Server (UI yang sudah ada, diperluas):**

- **[Keputusan perancang]** Server bisa diakses dari jaringan kantor. Host dan port diatur lewat setting, dengan default tetap `127.0.0.1` supaya pemakaian lokal tidak berubah.
- **[Keputusan perancang]** Data upload dan hasil proses dipisah per sesi browser (cookie sesi acak). Ini menutup kekurangan yang sudah tercatat: upload saat ini tercampur antar-tab dan antar-pengguna. Sesi lama dibersihkan otomatis.
- **[Keputusan perancang]** Hanya satu penulisan boleh berjalan pada satu waktu (kunci di server). Pengguna kedua mendapat pesan "sedang ada penulisan oleh <nama>, coba lagi nanti".
- **[Keputusan perancang]** Path IM MASTER, folder cadangan, host, dan port disimpan di file setting lokal di PC pusat (mis. `matrix2im.local.toml`, dibaca dengan `tomllib` bawaan Python). File ini di-ignore git. User tidak lagi upload IM MASTER untuk penulisan; untuk validasi header, server memakai file dari setting.
- **[Keputusan perancang]** MATRIX tetap di-upload lewat browser seperti sekarang.

## 6. Aturan penulisan

Aturan 1–7 di PRD bagian "Fase berikutnya: penulisan ke IM MASTER" berlaku penuh:

- hanya baris colorway yang diisi;
- semua grup size mendapat daftar colorway yang sama dengan urutan MATRIX;
- baris berlebih: WS TAG COLOR dan WHS dikosongkan dan berlatar kuning (lihat keputusan 11 November 2026 di bawah; menggantikan "diisi colorway terakhir");
- baris yang kurang disisipkan di dalam grup dengan menyalin baris di atasnya, dan formula/SUBTOTAL ikut menyesuaikan;
- WS TAG COLOR ke kolom CD dan WHS ke kolom CE, keduanya dicocokkan lewat nama header;
- nilai null tidak ditulis;
- kolom formula tidak pernah disentuh.

Tambahan:

- **[Keputusan perancang]** Pencocokan DESCRIPTION IM dengan style mengabaikan spasi berlebih dan huruf besar/kecil.
- **[Keputusan perancang, dari data F27, 11 November 2026]** Baris subtotal dikenali dari rumus `SUBTOTAL` di barisnya, bukan dari labelnya.
  - Di F27 ada 10 baris subtotal berlabel tidak baku, misalnya "-BOTH WHS", "M/L- SHOPIFY ONLY" dan "SY ONLY".
  - Ada juga 4 baris colorway yang namanya memuat "BOTH WHS".
- **[Keputusan user, 11 November 2026]** Kunci grup = DESCRIPTION dipotong pada dua segmen pertama ("STYLE - SIZE"). Teks sesudahnya di baris colorway adalah catatan.
  - Contoh: WHITNEY POCKET TUNIC CHUNKY "… - S/M - BOTH WHS - ALT 1 - FAF" dan "… - M/L - BOTH WHS - FAF - ALT 1/2/3 (INSET)".
  - Baris itu masuk grup size-nya, dan kolom COLOR-nya ditimpa seperti baris lain.
  - DESCRIPTION-nya tidak diubah.
- **[Keputusan user, 11 November 2026]** Style yang bloknya muncul lebih dari sekali di sheet ikut ditulis di semua blok. Contoh: AUTUMN LEAF CREW CHUNKY, yang muncul lagi di section "ADDITIONAL FULL PRICE AND SALE O4 (DELETE THIS WHEN SENDING IM TO PAOLA)". Pratinjau memberi peringatan per kunci yang muncul di bagian sheet terpisah.
- **[Keputusan perancang]** Kolom dicari lewat nama header di baris header IM (`IM_SHEET`, `IM_HEADER_ROW` di `legend_config.py`), bukan huruf kolom yang di-hardcode.
- **[Keputusan perancang]** Header "WHS" muncul dua kali (kolom G dan CE). Sesuai aturan PRD no. 6, yang ditulis adalah WHS yang tepat di kanan "WS TAG COLOR" (CE). Kolom G tidak disentuh. Bila posisi itu tidak ditemukan, penulisan berhenti dengan error.
- **[Keputusan perancang]** Nilai seperti `000` ditulis sebagai teks, bukan angka, supaya tidak berubah menjadi `0`. Verifikasi (§7 no. 6) memeriksanya.
- **[Keputusan user, sebelumnya]** WS TAG COLOR bernilai `Need Review` tetap ditulis ke IM supaya terlihat tim.
- **[Keputusan perancang]** Style di IM yang tidak ada di MATRIX tidak disentuh, dan didaftar di laporan.
- **[Keputusan user]** Style yang muncul di lebih dari satu sheet MATRIX (contoh F27: HUTTON MOCK T NECK CHUNKY):
  - Bila **himpunan colorway-nya sama**, style ditulis sekali dengan urutan sheet yang pertama muncul di MATRIX.
  - WHS yang berbeda antar-sheet digabung: `000` + `BOTH WHS` → `BOTH WHS`, `000` + `SY` → `BOTH WHS`.
  - Bila **colorway-nya berbeda**, style dilewati dan masuk laporan. Tidak ada tebakan.
  - Catatan: `final.json` tetap menyimpan record terpisah per sheet sesuai keputusan review. Penggabungan ini hanya terjadi di perencana.
- **[Keputusan user]** Isi lama di baris IM ditimpa. Di salinan AI TEST V2, 844 dari 851 grup size berisi colorway yang berbeda dari MATRIX F27. User menyimpan file asli dan akan membandingkannya sendiri setelah penulisan. Kolom lain di baris itu (kancing/aksesoris di slot COLOR lain, C1 ENDS, COLOR SEQUENCE, REPS COLOR, dst.) tidak diubah tool. Karena itu laporan mencantumkan setiap baris yang colorway-nya berubah, supaya mudah dicek.
- **[Keputusan perancang, tafsiran dari "timpa saja", minta dicek user]** Slot COLOR1..COLORn (n = jumlah benang colorway baru) selalu ditimpa, walaupun sebelumnya berisi aksesoris. Untuk slot di atas COLORn dipakai usulan PRD: slot berisi benang lama dikosongkan, sedangkan slot berisi aksesoris (awalan "X", mis. XBP/XZIP) dibiarkan.
- **[Keputusan user]** Menjalankan ulang:
  - Tool boleh menimpa nilai yang sudah ada, kecuali satu hal: sel yang sudah berisi tidak pernah diganti dengan `Need Review`. Ini supaya perbaikan tangan TIM IM tidak hilang.
  - Semua sel yang nilainya akan berubah ditampilkan di ringkasan sebelum konfirmasi.
  - Cara memenuhinya: UI menampilkan 500 perubahan pertama. Daftar lengkapnya bisa diunduh sebagai laporan Excel. Setiap penulisan juga menyimpan laporan Excel di samping cadangan.
- **[Keputusan user]** Baris sisipan menyalin seluruh isi baris di atasnya, termasuk kolom per colorway. Baris sisipan ditandai di laporan supaya dicek.
- **[Keputusan perancang, dari spike]** Cara menyisipkan k baris pada grup dengan R baris:
  - R ≥ 2: sisip k baris utuh tepat di atas baris colorway terakhir, lalu `fill down` kolom A sampai kolom terpakai terakhir, dari baris di atas titik sisip sampai baris sisipan terakhir. Semua rujukan ke grup, termasuk SUBTOTAL, ikut melebar.
  - R = 1: sisip k baris utuh di bawah baris colorway, lalu `fill down` dari baris itu. Rumus di baris subtotal grup yang merujuk baris colorway (mis. `CG229`) diperlebar menjadi range (`CG229:CG231`). Rujukan ke baris itu di luar baris subtotal tidak ikut melebar, dan laporan mencantumkan peringatannya.
  - Grup diproses dari bawah ke atas supaya nomor baris grup di atasnya tidak bergeser selama proses.
- **[Keputusan user]** Satu style dengan beberapa kartu di sheet yang sama (contoh F27: CROP BOYFRIEND CREW COTTON S2 - A/B, CHRISTMAS LABRADOR CREW CHUNKY 1A/1B, CHIN-CHIN CREW CHUNKY S9 - A/B) digabung menjadi satu daftar: colorway kartu A, lalu kartu B, sesuai urutan di MATRIX. Aturan lintas sheet di atas (HUTTON) tetap berlaku untuk style yang muncul di sheet berbeda.
- **[Keputusan perancang, menunggu konfirmasi user]** File lain yang merujuk ke sel IM MASTER diperlakukan sama seperti saat TIM IM menyisipkan baris secara manual: tidak ada penanganan khusus. Laporan mencantumkan baris mana saja yang disisipkan.
- **[Keputusan user, 11 November 2026]** Label WHS `000 & SY` diganti `BOTH WHS` (konstanta `WHS_BOTH`). `final.json` dengan label WHS di luar `000`/`SY`/`BOTH WHS` ditolak saat rencana dibuat, dengan pesan untuk membuat ulang `final.json`.
- **[Keputusan user, 11 November 2026]** Warna latar sel WS TAG COLOR **dan WHS** mengikuti WHS colorway: `000` = FF99FF, `SY` = FFCC99, `BOTH WHS` = 969696 (`FILL_BY_WHS`, sama dengan kebiasaan di IM asli). Warna diterapkan di semua baris colorway yang dikelola tool, juga bila teksnya tidak berubah; WHS null tidak mengubah warna.
  - **[Keputusan perancang]** Warna ditulis tanpa openpyxl: `styles.xml` hanya ditambah (warna dan format baru di akhir `fills`/`cellXfs`; format lama tidak diubah; format yang sama dipakai ulang, jadi menjalankan ulang tidak menumpuk format). Di sheet, hanya atribut `s` sel WS TAG COLOR dan WHS yang diganti, juga di sel rumus.
  - Baris sisipan mewarisi warna baris sumbernya (Excel `fill down` menyalin format); rencana warna dihitung ulang sesudah sisip dan harus sama dengan pratinjau.
  - Laporan mencatat jenis `warna` (warna lama dan baru, "tanpa warna" bila kosong). Ringkasan menghitung `fills_changed`.
- **[Keputusan user, 13 November 2026]** Sel COLOR1..COLORn yang ditulis dari MATRIX diberi latar abu 7F7F7F (`COLOR_FIXED_FILL`) sebagai tanda benang di sel itu sudah final.
  - Berlaku juga bila isinya sudah sama dengan MATRIX, di baris sisipan, dan untuk COLOR berisi deskripsi stripe/marl (Needs review).
  - Warna dibiarkan seperti di IM untuk slot di atas COLORn yang dikosongkan, aksesoris (XBP/XZIP), baris lebih, dan style yang tidak ditulis.
  - **[Keputusan user, 13 November 2026, lanjutan]** Sel yang sudah abu gelap tidak dicat ulang; isinya saja yang ditimpa. Abu gelap = 7F7F7F, 808080 (yang dipakai IM tim), atau indexed 23 (= 808080 di palet bawaan; `COLOR_DARK_FILLS`). Kondisi bawaan sebelum COLOR diisi adalah abu muda C0C0C0, dan sel itu menjadi 7F7F7F setelah ditulis.
  - **[Keputusan perancang]** Sel COLOR berisi rumus tidak ditulis, jadi juga tidak diberi abu.
  - Cara menulis, laporan (`warna`), dan verifikasi sama dengan warna WS TAG COLOR/WHS; `read_layout` kini membaca warna COLOR1..18 juga.
- **[Keputusan user, 11 November 2026]** Baris grup IM yang melebihi jumlah colorway MATRIX ("baris lebih", 3.044 baris di 758 grup pada salinan AI TEST V2): WS TAG COLOR dan WHS dikosongkan dan kedua selnya diberi latar kuning FFFF00 (`EXTRA_ROW_FILL`; revisi user setelah sempat "tanpa warna"), dan COLOR dibiarkan apa adanya. Jumlah baris tidak diubah. Laporan mencatat jenis `baris lebih`; ringkasan menghitung `extra_rows`. Sel rumus tetap tidak ditulis (peringatan).
- **[Terbuka]** Label subtotal otomatis (usulan PRD: semua 000 → "000 ONLY", semua SY → "SHOPIFY ONLY", campuran → "BOTH WHS"). Sampai diputuskan, label subtotal **tidak diubah**.

## 7. Pengaman

**[Keputusan perancang]**, disetujui arahnya oleh user ("hanya mengandalkan koordinasi"; tool memberi pengaman tanpa menuntut kebiasaan baru):

1. **Password.**
   - User mengetik password di UI saat menulis.
   - Server mencocokkan password dengan hash `fileSharing` di file sebelum melakukan apa pun.
   - Password tidak disimpan di disk, tidak masuk log, dan tidak muncul di argumen command line. Password dikirim ke `osascript` lewat environment atau stdin.
2. **Ringkasan dan konfirmasi.**
   - UI menampilkan jumlah style, baris yang diisi, baris yang disisipkan, baris yang colorway-nya berubah, style yang dilewati, dan peringatan.
   - UI juga menampilkan daftar sel yang nilainya akan berubah (nilai lama → baru).
   - User harus menekan konfirmasi.
3. **Cek file sedang dibuka.** Bila ada file pemilik Excel `~$<nama file>` di folder IM MASTER, penulisan dihentikan dan nama di dalam file itu ditampilkan.
4. **Cadangan.** Sebelum menulis, IM MASTER disalin ke folder cadangan dengan nama bertanggal-jam. Google Drive juga menyimpan versi lama sebagai lapis kedua.
5. **Deteksi perubahan.**
   - Tool mencatat ukuran, waktu ubah, dan hash file saat rencana dibuat.
   - Tepat sebelum Excel membuka file, tool mengecek ulang. Bila berbeda, penulisan dibatalkan dan user diminta mengulang.
6. **Verifikasi setelah menulis.** File dibaca ulang dengan openpyxl `read_only`. Nilai setiap sel dalam rencana harus sama, dan kunci `fileSharing` harus masih ada. Warna latar WS TAG COLOR dan WHS di semua baris grup harus sama dengan sebelum ditulis ditambah rencana warna, dan `styles.xml` hanya boleh bertambah di akhir daftar warna/format (tanpa rencana warna, `styles.xml` harus identik). Bila gagal, laporan menyebut sel yang berbeda dan lokasi cadangan.
7. **Log.** Setiap penulisan dicatat: waktu, nama user, nama file MATRIX, ringkasan rencana, hasil verifikasi, dan path cadangan. Formatnya satu baris JSON per penulisan, di file log di folder cadangan.

Celah yang diterima: bila seseorang sudah membuka IM MASTER di Excel di PC lain sebelum tool menulis lalu menyimpan sesudahnya, penyimpanannya bisa menimpa hasil tool. Cadangan dan log menjadi jalan pemulihan.

## 8. Operasional di PC pusat

Ini dilakukan user sendiri, karena menyangkut pengaturan sistem dan keamanan:

- Memberi izin akses folder Google Drive ke aplikasi yang menjalankan server (System Settings → Privacy & Security).
- Mengizinkan Terminal/Python mengontrol Microsoft Excel saat pop-up pertama muncul.
- Mengatur Mac supaya tidak sleep, dan mengizinkan koneksi masuk di firewall untuk port server.
- **[Keputusan perancang]** Server tidak dipasang sebagai layanan otomatis saat login pada tahap ini. Bila diinginkan nanti, itu langkah terpisah dengan persetujuan user.
- Saat ada penulisan, Excel di PC pusat membuka dan menutup file secara otomatis. Hindari bekerja di Excel di PC pusat pada saat itu.

## 9. Batasan

- Python 3.11+, openpyxl, pytest, argparse, dan library bawaan Python saja. Excel dikendalikan lewat `osascript` (bawaan macOS).
- File input tidak pernah di-`save()` oleh openpyxl. Satu-satunya yang mengubah IM MASTER adalah penjalan Excel, dan itu hanya saat user menekan Tulis.
- Test tidak pernah memakai file IM MASTER di Google Drive. Unit test perencana memakai IM sintetis dari `tests/helpers/matrix_factory.py`. Test penjalan memakai salinan sementara, diberi marker tersendiri, dan di-skip bila bukan macOS atau Excel tidak ada.
- Pesan UI, CLI, dan log berbahasa Indonesia. Konstanta aturan tetap di `legend_config.py`.
- File data, cadangan, log, dan setting lokal tidak di-commit.

## 10. Risiko yang harus diuji dulu (spike) sebelum membangun penuh

Status 10 November 2026:
- No. 1 dan no. 2 **sudah terjawab** di salinan (§4), dan hasilnya membentuk §5 dan §6.
- No. 5 ditangani dengan batas waktu di penjalan Excel.
- No. 3 dan no. 4 butuh Google Drive dan dua PC, jadi menjadi uji terima manual oleh user.

1. **Sisip baris lewat Excel di dalam range SUBTOTAL.** Bila baris disisipkan tepat di bawah baris terakhir range (sebelum baris subtotal), Excel tidak memperluas range `SUBTOTAL`. Baris harus disisipkan *di dalam* range, mis. di atas baris colorway terakhir, lalu diisi salinan. Ini harus dibuktikan di salinan IM, termasuk tangga harga WHLS LINE PRICE (+1 per baris) dan grup M/L yang punya subgrup FAF/FAH.
2. **Kecepatan menulis ribuan sel dan ratusan sisip baris lewat AppleScript.** Targetnya satu penulisan penuh F27 selesai dalam hitungan menit. Bila terlalu lambat, cari cara penulisan per blok yang lebih besar.
3. **Identitas file di Google Drive setelah Excel menyimpan.** Pastikan file tetap file yang sama (link dan riwayat versi tidak putus). Bila Excel menyimpan sebagai file baru, perlu strategi lain.
4. **Apakah file `~$` dari PC lain terlihat di PC pusat** lewat Google Drive. Bila tidak, pengaman no. 3 hanya berlaku untuk Excel di PC pusat, dan UI harus menyebut keterbatasan itu.
5. **Dialog tak terduga di Excel** (update link, repair, file terkunci) tidak boleh membuat server macet. Butuh batas waktu dan pesan error yang jelas.

**[Keputusan user]** Rencana cadangan bila spike no. 1 atau no. 3 gagal:
- Tool tetap dipakai dengan hanya mengisi baris yang sudah ada, lewat cara B: ubah sel langsung di XML sheet. Cara ini sekitar 2 detik dan tidak menyentuh bagian lain file.
- Grup yang kekurangan baris didaftar di laporan untuk ditambah manual.
- Karena itu, perencana, verifikasi, dan UI dirancang supaya bisa bekerja dengan kedua penjalan. Hasil spike menentukan penjalan mana yang dipakai sebelum pembangunan penuh.

## 11. Diserahkan ke builder

- Bentuk data rencana perubahan, nama modul, dan pembagian file.
- Tata letak UI untuk tombol Tulis, form nama/password, ringkasan, dan laporan.
- Format nama file cadangan dan isi baris log.
- Cara membersihkan sesi lama dan batas ukuran upload per sesi.
- Detail perintah AppleScript, selama memenuhi §5, §7, dan hasil spike §10.
