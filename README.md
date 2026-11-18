# 📊 SIPD RKA Excel Cleaner & Compactor

Script automasi Python untuk merapikan dan memadatkan file Excel RKA (Rencana Kerja dan Anggaran) Belanja dari sistem **SIPD-RI** hasil konversi alat seperti **PDFgear**.

---

## 🎯 Masalah yang Diselesaikan
Konversi dokumen PDF RKA SIPD ke format Spreadsheet (.xls/.xlsx) sering kali menghasilkan layout yang sangat berantakan:
1. **Pecah Kolom Liar:** Membentuk hingga 28 kolom acak dengan lebar hanya 1–2 pt karena mendeteksi koordinat visual PDF.
2. **Baris Menggelembung:** Setiap 1 pos rekening belanja dipecah menjadi 3–4 baris (baris nama, baris `[ # ] Gaji`, baris tahun, baris koefisien). Akibatnya dokumen membengkak hingga ratusan baris.
3. **Pecahan Header Berulang:** Header halaman dan footer cetak SIPD terselip di tengah-tengah tabel data.
4. **Kolom Nomor Rekening Rusak:** Digit nomor rekening terpotong atau wrap text tidak beraturan.

---

## ✨ Fitur & Solusi
* **Presisi Berbasis Digit Terpanjang:** Menyesuaikan lebar kolom nomor rekening tepat pada digit rekening terpanjang (21 digit, misal: `5.1.01.01.001.00001` - Belanja Gaji Pokok PNS) dengan padding pas sebesar `24.0` karakter.
* **Super Compact (Hemat Space):** Memadatkan tabel dari **ratusan baris x 28 kolom** menjadi format bersih (1 baris per rekening).
* **Hirarki Visual Bertingkat:** Memberikan indentasi spasi otomatis serta pewarnaan lembut berdasarkan tingkatan akun (Level 1 s/d Level 6).
* **Formula Excel Dinamis:** Kolom **Selisih** (`=E-D`) dan **% Perubahan** (`=IF(D=0,0,(E-D)/D)`) menggunakan formula Excel asli, dengan angka minus berwarna **merah**.
* **Filter & Freeze Panes:** Baris header tabel otomatis terkunci saat di-scroll dan auto-filter langsung aktif.
* **Multi-Sheet:** Menghasilkan sheet ringkas (`RKA_Kompak`), sheet standar SIPD lengkap (`RKA_Rincian_Standar`), dan sheet metadata kegiatan (`Info_Sub_Kegiatan`).

---

## 📁 Struktur Berkas
| File | Keterangan |
| :--- | :--- |
| `rapihkan_rka.py` | **Script All-in-One (Windows Native):** Menggunakan COM Excel via PowerShell, berjalan langsung tanpa perlu `pip install`. |
| `pure_cleaner.py` | **Versi 100% Python Murni:** Menggunakan `openpyxl` & `xlrd`, multiplatform (Windows, Linux, macOS). |
| `jalankan.bat` | Pintasan 1-klik untuk Windows Explorer (double click langsung jalan). |
| `requirements.txt` | Dependensi jika menggunakan `pure_cleaner.py`. |

---

## 🚀 Cara Penggunaan

### Cara 1: Menggunakan Script All-in-One (Windows Native)
1. Letakkan file `.xls` hasil konversi PDFgear di direktori yang sesuai (misal: di folder `Documents\PDFgear` atau folder yang sama dengan script).
2. Buka `rapihkan_rka.py` dan sesuaikan nama file sumber Anda:
   ```python
   source_path = os.path.join(base_dir, "contoh_rka_cetak_sipd_xxx.xls")
   ```
3. Jalankan:
   ```bash
   python rapihkan_rka.py
   ```
   Atau double-click file **`jalankan.bat`**.

### Cara 2: Menggunakan Versi Pure Python (Multiplatform)
1. Pasang dependensi:
   ```bash
   pip install -r requirements.txt
   ```
2. Jalankan script:
   ```bash
   python pure_cleaner.py
   ```

---

## 🔒 Perlindungan Data Sensitif

* Seluruh nama pemerintah daerah, OPD/SKPD, sub kegiatan, target, dan nominal anggaran di dalam kode script telah disanitasi menggunakan placeholder umum (`XXX`).
* File data Excel (`*.xlsx`, `*.xls`, `*.csv`) telah dimasukkan ke dalam `.gitignore` untuk mencegah kebocoran dokumen anggaran riil daerah ke publik.

---

## 📄 Lisensi
MIT License - Bebas digunakan dan dimodifikasi untuk kebutuhan administrasi pemerintahan dan penyusunan anggaran daerah.
