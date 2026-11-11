# Uji terima: menulis ke IM MASTER

Daftar centang untuk TIM IM. Semua uji dilakukan pada **salinan uji**, bukan file IM MASTER yang asli. Centang kotak setelah tiap langkah berhasil, dan catat bila ada yang aneh. Latar belakang uji 1 sampai 5 ada di spec (`docs/superpowers/specs/2026-11-10-tulis-im-master-design.md`, §10).

Jangan menulis password di dokumen ini, di chat, atau di file setting. Password hanya diketik di kotak password di UI atau di prompt perintah.

## 0. Menyiapkan PC pusat (dikerjakan oleh Anda, bukan oleh tool)

Langkah ini mengubah pengaturan sistem dan keamanan Mac, jadi harus Anda sendiri yang mengerjakannya.

- [ ] **Akses folder Google Drive.** Buka System Settings, Privacy & Security, Files and Folders (atau Full Disk Access), lalu beri Terminal (atau aplikasi yang menjalankan Python) izin ke folder Google Drive.
- [ ] **Izin Automation untuk Excel.** Saat pop-up "Terminal ingin mengontrol Microsoft Excel" muncul pertama kali, klik **Izinkan**. Bila terlewat: System Settings, Privacy & Security, Automation, aktifkan Excel di bawah Terminal.
- [ ] **Folder `kerja`.** Excel memakai salinan sementara di folder `kerja` di dalam folder cadangan. Bila macOS bertanya apakah Excel boleh membuka file di folder itu, klik **Izinkan** (cukup sekali).
- [ ] **Mac tidak tidur.** System Settings, Energy atau Lock Screen: matikan tidur otomatis selama PC pusat dipakai.
- [ ] **Firewall.** System Settings, Network, Firewall: izinkan koneksi masuk untuk Python (port 8765) bila firewall menyala.
- [ ] **Menjalankan server.** Di Terminal, di folder proyek: `.venv/bin/matrix2im ui --host 0.0.0.0 --no-browser`. Biarkan jendela Terminal tetap terbuka. Catatan: selama menulis, peringatan (alert) Excel dimatikan oleh tool dan dinyalakan lagi sesudahnya.
- [ ] **Mencari alamat IP Mac.** Jalankan `ipconfig getifaddr en0` di Terminal (atau lihat System Settings, Network). Alamat UI untuk PC lain: `http://<alamat IP>:8765/`.
- [ ] **File setting.** Buat `matrix2im.local.toml` (contoh di README, bagian "Setting `matrix2im.local.toml`"). Untuk uji ini, `master` menunjuk ke **salinan uji** (lihat uji 1), dan `backup_dir` ke folder cadangan di PC pusat.

## 1. Menulis ke salinan uji di Google Drive (spec §10.3)

Tujuan: memastikan link dan riwayat versi file di Google Drive tidak putus setelah tool menulis.

1. [ ] Di folder Google Drive, buat **salinan** IM MASTER (klik kanan, Buat salinan), beri nama jelas, misalnya `F27 IM MASTER (UJI TERIMA)`. Jangan memakai file asli.
2. [ ] Catat dua hal tentang salinan itu: link-nya (Bagikan, Salin link) dan jumlah versi di riwayat (klik kanan, Kelola versi).
3. [ ] Arahkan `master` di `matrix2im.local.toml` ke salinan uji tadi, lalu mulai ulang server.
4. [ ] Buka UI, upload MATRIX saja (IM MASTER tidak perlu diupload: tool menulis ke file `master` di setting; kotak upload IM hanya untuk validasi header), klik **Proses**, lalu di panel **Tulis ke IM MASTER** isi nama dan password, klik **Buat rencana**, periksa pratinjau, centang pernyataan, dan klik **Tulis semua ke IM**.
5. [ ] Tunggu laporan. Yang harus terlihat: laporan berhasil, lokasi cadangan, lokasi laporan Excel (`... laporan matrix2im.xlsx`), dan ringkasan jumlah sel serta baris.
6. [ ] Buka link yang dicatat di langkah 2. **Yang harus sama:** link masih bekerja dan file yang terbuka adalah file yang sama (bukan file baru bernama lain).
7. [ ] Buka Kelola versi. **Yang harus sama:** versi lama masih ada dan satu versi baru bertambah di atasnya.

Hasil: [ ] lulus  [ ] gagal. Catatan: ____________________

## 2. File `~$` dari PC lain (spec §10.4)

Tujuan: mengetahui apakah tool di PC pusat menolak menulis ketika orang lain membuka salinan uji di Excel.

1. [ ] Di PC lain (bukan PC pusat), buka **salinan uji** di Excel dan biarkan terbuka. Jangan menyimpan.
2. [ ] Tunggu sekitar satu menit supaya Google Drive menyinkronkan.
3. [ ] Di PC pusat, di UI, klik **Buat rencana** (atau jalankan `.venv/bin/matrix2im write final.json --dry-run`).
4. [ ] **Yang harus terlihat:** tool menolak dengan pesan "IM MASTER sedang dibuka oleh <nama orang itu>...". Bila pesan itu muncul, centang "terlihat".
5. [ ] Bila tool **tidak** menolak, centang "tidak terlihat". Artinya pengecekan hanya melindungi dari Excel di PC pusat, dan TIM IM harus tetap berkoordinasi manual sebelum menulis.
6. [ ] Tutup file di PC lain.

Hasil: [ ] terlihat (tool menolak)  [ ] tidak terlihat (koordinasi manual). Catatan: ____________________

## 3. Membuka UI dari PC Windows

1. [ ] Pastikan server di PC pusat berjalan dengan `--host 0.0.0.0` (uji 0).
2. [ ] Di PC Windows (satu jaringan kantor), buka browser dan ketik `http://<alamat IP PC pusat>:8765/`.
3. [ ] **Yang harus terlihat:** halaman matrix2im terbuka. Upload MATRIX, klik **Proses**, hasilnya tampil.
4. [ ] Di panel **Tulis ke IM MASTER**, klik **Buat rencana** dengan nama dan password; pratinjau muncul. (Penulisan penuh tidak wajib di langkah ini.)
5. [ ] Bila halaman tidak terbuka: periksa alamat IP, firewall PC pusat (uji 0), dan bahwa kedua PC satu jaringan.

Hasil: [ ] lulus  [ ] gagal. Catatan: ____________________

## 4. Membandingkan hasil dengan file asli

Tujuan: memastikan isi yang ditulis tool sesuai harapan menurut cara pemeriksaan TIM IM sendiri.

1. [ ] Buka dua file berdampingan: salinan uji setelah ditulis (uji 1) dan file IM MASTER yang asli (yang tidak disentuh).
2. [ ] Periksa dengan cara yang biasa Anda pakai (filter, rumus, atau perbandingan manual): COLOR1 dan seterusnya, WS TAG COLOR, dan WHS pada beberapa style, termasuk satu style yang bertambah barisnya.
3. [ ] **Yang harus terlihat:** kolom yang diubah tool hanya COLOR, WS TAG COLOR, dan WHS (plus warna latar WS TAG COLOR); kolom rumus dan kolom lain tidak berubah. Warna latar WS TAG COLOR: pink FF99FF untuk WHS 000, oranye muda FFCC99 untuk SY, abu 969696 untuk BOTH WHS. Di baris yang melebihi jumlah colorway MATRIX, WS TAG COLOR dan WHS kosong tanpa warna, sedangkan COLOR tetap seperti sebelumnya. Tangga harga WHLS LINE PRICE tetap utuh, dan SUBTOTAL di grup yang disisipi baris mencakup baris-baris baru. Bandingkan dengan file asli: sebagian range SUBTOTAL di file asli memang sudah tidak mencakup seluruh blok, jadi itu bukan kesalahan tool.
4. [ ] Gambar dan warna kondisional di salinan uji masih ada seperti di file asli.
5. [ ] Periksa log `matrix2im-log.jsonl` di folder cadangan: ada satu baris untuk penulisan tadi (hasil "ok"), dan tidak ada password di dalamnya.
6. [ ] Buka laporan Excel di folder cadangan (atau unduhan **Unduh semua perubahan (Excel)** dari pratinjau) di PC Windows: kolom terpisah dengan benar, dan baris `sisip`, `baris lebih`, `warna`, serta `baris colorway berubah` cocok dengan yang terlihat di file.

Hasil: [ ] lulus  [ ] ada perbedaan. Catatan: ____________________

## 5. Dua browser menulis bersamaan

1. [ ] Buka UI di dua browser (atau dua PC) dan proses MATRIX di keduanya sampai panel **Tulis ke IM MASTER**.
2. [ ] Di browser pertama, klik **Buat rencana** lalu **Tulis semua ke IM**. Sementara masih berjalan, di browser kedua klik **Buat rencana** atau **Tulis semua ke IM**.
3. [ ] **Yang harus terlihat:** browser kedua mendapat pesan bahwa penulisan sedang berjalan (ditolak dengan jelas), bukan error aneh atau data rusak.
4. [ ] Setelah penulisan pertama selesai, browser kedua bisa mencoba lagi. Bila file sudah berubah sejak rencana dibuat, tool meminta rencana dibuat ulang.
5. [ ] Periksa bahwa log hanya mencatat penulisan yang benar-benar berjalan.

Hasil: [ ] lulus  [ ] gagal. Catatan: ____________________

## Setelah semua lulus

- [ ] Kembalikan `master` di `matrix2im.local.toml` ke file yang dimaksud hanya setelah Anda yakin dan setelah menyimpan cadangan sendiri.
- [ ] Pulihkan bila perlu: tutup file di semua PC, lalu salin isi cadangan di `backup_dir` menimpa file (atau pakai Kelola versi di Google Drive).
