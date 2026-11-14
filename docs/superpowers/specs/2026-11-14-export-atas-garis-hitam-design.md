# Export Excel "Atas garis hitam": desain singkat

Disetujui user 14 November 2026 ("langsung kerjakan").

## Tujuan

Tim mendapat file Excel yang hanya berisi colorway **di atas garis hitam** di kartu MATRIX. File ini **untuk dicek tim**, bukan untuk diolah tool lain. Karena itu formatnya Excel, bukan JSON.

## Yang dibangun

1. **Tombol** **⬇ Atas garis hitam (Excel)** di UI, di sebelah ⬇ raw.json dan ⬇ final.json.
   - Tombol ini muncul setelah Proses, sama seperti tombol unduh lainnya.
   - Isinya diambil dari hasil Proses terakhir di sesi itu.
2. **Endpoint** `GET /api/download/atas-garis-hitam.xlsx`.
   - Mengembalikan 404 bila sesi belum menjalankan Proses.
   - Nama file unduhan: `matrix2im atas garis hitam <nama file MATRIX tanpa .xlsx>.xlsx`.
3. **Sheet "Atas garis hitam":** 1 baris = 1 colorway di atas garis hitam, urut seperti di MATRIX (style, lalu baris kartu).
   - Kolom: `COLLECTION | SUB | No | DESCRIPTION | COLOR1 … COLORn | WS TAG COLOR | WHS | SEL MATRIX | NEEDS REVIEW`.
   - Jumlah COLOR = jumlah benang terbanyak di antara colorway yang diekspor. Sel tanpa benang dibiarkan kosong.
   - Nilai `WS TAG COLOR` dan `WHS` sama dengan final.json, termasuk teks `Need Review`.
   - `NEEDS REVIEW` = semua `issue` needs_review untuk record itu, digabung dengan `; `.
   - Header ditebalkan, baris header dibekukan, dan semua kolom bisa difilter.
4. **Sheet "Tanpa warna atas garis hitam":** kartu yang punya warna di atas garis merah, tetapi tidak punya satu pun warna di atas garis hitam. Warna kartu-kartu ini tidak muncul di sheet pertama.
   - Kolom: `COLLECTION | SUB | No | DESCRIPTION | JUMLAH WARNA ANTARA HITAM & MERAH | KETERANGAN`.
   - `KETERANGAN` berisi `garis hitam tidak ditemukan` bila Tahap 1 memberi warning itu untuk kartu tersebut. Selain itu berisi `semua warna di bawah garis hitam`.
   - Bila tidak ada kartu seperti itu, sheet berisi header dan satu baris `(tidak ada)`.
   - Kartu yang semua warnanya di bawah garis merah (mis. halaman S1–S20 di F27) **tidak** dicantumkan, karena memang tidak punya warna di atas garis merah.

## Pendekatan

- **Dipilih:** server membuat file .xlsx dengan openpyxl write-only, dari `raw` + `final` yang sudah ada di sesi. Ini sama seperti laporan rencana tulis (`/api/write/preview.xlsx`).
- **Tidak dipilih:**
  - Excel dibuat di browser: butuh library JS tambahan, dan UI ini sengaja tanpa dependensi luar.
  - Filter di tabel Records: tidak menghasilkan file yang bisa dibagikan ke tim.

## Batasan

- final.json, raw.json, CLI, dan proses tulis ke IM tidak berubah. Format raw.json juga tidak berubah, karena dibandingkan persis dengan golden F27.
- Record dipasangkan ke kartu raw lewat urutan pembuatan di `transform` (style → grup garis → colorway). Bila `_ref` record tidak cocok dengan pasangannya (`above_black` dan sel yang sama), unduhan gagal dengan pesan jelas. Tidak boleh ada data yang tertukar diam-diam.
- Sel teks ditulis sebagai teks: `000` tetap `000`, dan teks berawalan `=` tidak menjadi rumus. Karakter kontrol dibuang. No berupa angka tetap angka.
