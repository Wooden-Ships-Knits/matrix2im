# Export Excel "Atas garis hitam": rencana implementasi

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tombol di UI yang mengunduh Excel berisi colorway di atas garis hitam saja, ditambah daftar kartu yang tidak punya warna di atas garis hitam.

**Architecture:** Modul baru `matrix_to_im/black_line_export.py` membangun baris (fungsi murni) dan bytes .xlsx (openpyxl write-only) dari `raw` + `final`. `ui_server` menambah satu route GET yang memakai hasil Proses terakhir di sesi. `index.html` menambah satu tautan unduh.

**Tech Stack:** Python 3.12, openpyxl 3.1.5, http.server, pytest.

**Spec:** `docs/superpowers/specs/2026-11-14-export-atas-garis-hitam-design.md`

## Global Constraints

- **Nama sheet:**
  - `Atas garis hitam`
  - `Tanpa warna atas garis hitam`
- **Kolom sheet 1:** `COLLECTION, SUB, No, DESCRIPTION, COLOR1..COLORn, WS TAG COLOR, WHS, SEL MATRIX, NEEDS REVIEW`.
- **Kolom sheet 2:** `COLLECTION, SUB, No, DESCRIPTION, JUMLAH WARNA ANTARA HITAM & MERAH, KETERANGAN`.
- **KETERANGAN:** `garis hitam tidak ditemukan` / `semua warna di bawah garis hitam`. Sheet kosong diisi `(tidak ada)`.
- **Route:** `GET /api/download/atas-garis-hitam.xlsx`.
- **Nama file:** `matrix2im atas garis hitam <stem MATRIX>.xlsx`.
- **Tidak berubah:** format raw.json/final.json.

## Review Focus

1. **`_ref` final tidak cocok dengan urutan raw** (mis. final dari file lain): yang diharapkan Matrix2ImError, bukan data tertukar. Diuji di Task 1.
2. **Teks `000` / `=X` / karakter kontrol:** harus tetap teks dan file tetap terbuka. Diuji di Task 1.
3. **Nama file MATRIX non-ASCII atau berisi tanda kutip:** header Content-Disposition harus valid. Diuji di Task 2.
4. **Unduh sebelum Proses atau dari sesi lain:** 404, sama seperti raw.json. Diuji di Task 2.
5. **Tidak ada colorway di atas garis hitam:** sheet 1 tetap ada, berisi header saja. Diuji di Task 1.

---

### Task 1: Modul export

**Files:**
- Create: `matrix_to_im/black_line_export.py`
- Modify: `matrix_to_im/extract.py`: pesan warning garis hitam menjadi konstanta `NO_BLACK_LINE`.
- Test: `tests/unit/test_black_line_export.py`

**Interfaces:**
- Produces:
  - `above_black_rows(raw: dict, final: dict) -> tuple[list[str], list[list]]` → (header, baris).
  - `missing_above_rows(raw: dict) -> list[list]`.
  - `above_black_xlsx_bytes(raw: dict, final: dict) -> bytes`.
  - Konstanta `ABOVE_SHEET`, `MISSING_SHEET`.

- [ ] Tes dari MATRIX sintetis (`build_matrix` → `extract` → `transform`):
  - Sheet 1 berisi tepat 1 baris: `TEST, None, 1, NEWPORT CREW, YCA KHAKI - 749, YCA ALMOND BUTTER - 872, …, A15`.
  - Sheet 2 berisi JUDE T NECK dan STRIPE CARDI dengan KETERANGAN `garis hitam tidak ditemukan`.
  - Kasus `_ref` diubah → Matrix2ImError.
  - Kasus tanpa colorway atas hitam → hanya header dan `(tidak ada)` di sheet 2.
  - Nilai `000` / `=X` tetap teks.
- [ ] Jalankan tes: gagal (modul belum ada).
- [ ] Implementasi. Pasangkan record lewat indeks berjalan atas `raw["styles"]` × `transform.GROUPS`.
- [ ] Jalankan seluruh tes unit: lolos.

### Task 2: Route + tombol UI + README

**Files:**
- Modify: `matrix_to_im/ui_server.py` (do_GET, handler unduh xlsx).
- Modify: `matrix_to_im/ui/index.html` (tautan di `.downloads`).
- Modify: `README.md` (baris tombol unduh).
- Test: `tests/unit/test_ui_server.py`

- [ ] Tes:
  - 404 sebelum Proses.
  - Sesudah Proses → 200, Content-Type xlsx, `Content-Disposition` memuat `filename*=UTF-8''matrix2im%20atas%20garis%20hitam%20m.xlsx`, dan kedua sheet ada.
  - Nama file berisi kutip atau non-ASCII → header valid.
- [ ] Jalankan tes: gagal.
- [ ] Implementasi route dan tautan.
- [ ] Jalankan seluruh test suite: lolos.
- [ ] Restart UI, unduh dari MATRIX asli, dan cek isi file.
