# SNAPRINT DIGITAL PRINTING — SISTEM ANALISIS KINERJA PENJUALAN MULTI-TAHUN
### Konsolidasi Laporan Pembukuan ACCURATE Accounting System (2024, 2025, 2026 s/d Agustus)

Repositori ini memuat sistem komprehensif pelaporan, visualisasi dashboard interaktif, ekspor data, dan basis data relasional SQL untuk penjualan **Snaprint Digital Printing**.

---

## 🌟 Ringkasan Performa Finansial Multi-Tahun

| Tahun Buku | Periode Pencatatan | Total Omzet Penjualan | Volume Order | Rata-rata Bulanan | Bulan Puncak (Peak) |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **2024** | 01 Jan 2024 – 31 Des 2024 (12 Bulan) | **Rp 930.847.260,01** | 275.498 unit | Rp 77,57 Jt/bln | Agustus (Rp 124,07 Jt) |
| **2025** | 01 Jan 2025 – 31 Des 2025 (12 Bulan) | **Rp 1.378.951.928,70** | 387.284 unit | Rp 114,91 Jt/bln | Agustus (Rp 140,09 Jt) |
| **2026** | 01 Jan 2026 – 31 Agu 2026 (8 Bulan) | **Rp 1.180.308.204,00** | 226.793 unit | Rp 147,54 Jt/bln | **Juli (Rp 173,73 Jt) ★** |

> **Highlight Pertumbuhan:**  
> • **2025 vs 2024 (Tahunan):** Omzet melonjak **+48,14% (+Rp 448,10 Juta)**.  
> • **2026 vs 2025 (8 Bulan Apel-ke-Apel):** Omzet bertumbuh **+37,29% (+Rp 320,57 Juta)**.  
> • **2026 vs 2024 (8 Bulan Apel-ke-Apel):** Omzet melesat **+132,22% (+Rp 672,04 Juta)**.

---

## 🚀 Fitur Sistem

1. **Dashboard Interaktif Multi-Tahun (`index.html`)**:
   - Tab pemilihan tahun instan (`2024`, `2025`, `2026`, dan `Komparasi Multi-Tahun`).
   - Grafik tren penjualan SVG beresolusi tinggi dengan *headroom* lapang dan label angka penjualan yang terbaca jelas tanpa terpotong.
   - Filter dinamis berdasarkan kuartal (Q1–Q4), semester (H1/H2), bulan, dan kategori produk.
   - Analisis Pareto 80/20 untuk 129 produk individu di tahun 2024.
2. **Basis Data Relasional SQL (`snaprint_sales.db` & `snaprint_sales.sql`)**:
   - File database SQLite mandiri (`.db`) dan skrip DDL/DML (`.sql`).
   - Dilengkapi tabel master tahun, kategori, penjualan bulanan, matriks kategori produk, serta *SQL Views* siap pakai.
   - Alat bantu query terminal via `query_db.py`.
3. **Ekspor Multi-Sheet Excel (`.xlsx`)**:
   - Menghasilkan 4 lembar kerja terintegrasi: *Komparasi 3 Tahun (8M)*, *Matriks 2026*, *Matriks 2025*, dan *129 Produk 2024*.
4. **Laporan Dokumen Resmi (PDF & Markdown)**:
   - Laporan PDF publication-grade 2024 (4 halaman), 2025 (2 halaman), dan 2026 (2 halaman).
   - Laporan analisis naratif eksekutif dalam format Markdown pada masing-masing folder tahun.

---

## 📁 Struktur Direktori Repositori

```text
SNAPRINT_GW2024/
├── index.html                                        # Aplikasi Dashboard Web Terpadu (2024-2026)
├── start_dashboard.bat                               # Windows launcher 1-klik untuk membuka dashboard
├── snaprint_sales.db                                 # Database relasional SQLite (Binary)
├── snaprint_sales.sql                                # Dump skrip SQL lengkap (DDL & DML)
├── query_db.py                                       # CLI tool untuk eksekusi query SQL cepat
├── DATABASE_README.md                                # Dokumentasi skema tabel & view database SQL
├── sales_data_master.json                            # Dataset induk konsolidasi JSON
├── data_multiyear.js                                 # Bundle JavaScript mandiri untuk offline viewer
├── Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf # Laporan Resmi PDF 2024 (4 Halaman)
├── Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf # Laporan Resmi PDF 2025 (2 Halaman)
├── Laporan_Penjualan_2026_Snaprint_Digital.pdf       # Laporan Resmi PDF 2026 (2 Halaman)
├── 2025/                                             # Folder Khusus Tahun 2025
│   ├── index.html                                    # Dashboard terintegrasi folder 2025
│   ├── LAPORAN_PENJUALAN_TAHUNAN_2025_SNAPRINT.md    # Laporan naratif Markdown 2025
│   ├── Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf
│   └── sales_data_2025.json                          # JSON database tahun 2025
└── 2026/                                             # Folder Khusus Tahun 2026 (s/d Agustus)
    ├── index.html                                    # Dashboard terintegrasi folder 2026
    ├── LAPORAN_PENJUALAN_TAHUNAN_2026_SNAPRINT.md    # Laporan naratif Markdown 2026
    ├── Laporan_Penjualan_2026_Snaprint_Digital.pdf
    └── sales_data_2026.json                          # JSON database tahun 2026
```

---

## 💻 Cara Menggunakan

### 1. Menjalankan Dashboard Penjualan
Cukup buka berkas `index.html` langsung di browser peramban web apapun (Google Chrome, Edge, Safari, Firefox), atau klik dua kali berkas `start_dashboard.bat`. Tidak memerlukan web server eksternal!

### 2. Menjalankan Query Database SQL
Jalankan perintah berikut di PowerShell atau Command Prompt:
```bash
# Menampilkan ringkasan tahunan dan komparasi bulanan
python query_db.py

# Menjalankan query kustom
python query_db.py "SELECT * FROM v_comparison_3_years_8m;"
python query_db.py "SELECT * FROM v_pareto_core80_2024 LIMIT 10;"
```

---

*Hak Cipta © 2024–2026 Snaprint Digital Printing. Seluruh data dikompilasi secara akurat dari ACCURATE Accounting System.*
