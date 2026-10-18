import json
import os

with open("sales_data_2024.json", "r", encoding="utf-8") as f:
    sales_data = json.load(f)

json_str = json.dumps(sales_data, ensure_ascii=False)

html_template = f"""<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SNAPRINT DIGITAL — Dashboard Kinerja Terintegrasi 2024</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    /* Custom Scrollbar */
    ::-webkit-scrollbar {{ width: 6px; height: 6px; }}
    ::-webkit-scrollbar-track {{ background: transparent; }}
    ::-webkit-scrollbar-thumb {{ background: rgba(148, 163, 184, 0.4); border-radius: 3px; }}
    ::-webkit-scrollbar-thumb:hover {{ background: rgba(148, 163, 184, 0.7); }}
    
    .chart-bar-group {{
      transition: transform 0.15s ease, filter 0.15s ease;
    }}
    .chart-bar-group:hover {{
      filter: brightness(1.08);
    }}
  </style>
</head>
<body class="bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] antialiased min-h-screen p-3 sm:p-5 text-xs sm:text-sm">

  <div class="max-w-7xl mx-auto space-y-4 sm:space-y-6">

    <!-- HEADER & STATUS BAR -->
    <header class="bg-[var(--card,#ffffff)] text-[var(--foreground,#0f172a)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
      <div class="flex items-center space-x-3 sm:space-x-4">
        <div class="w-11 h-11 rounded-xl bg-gradient-to-tr from-sky-600 to-indigo-600 flex items-center justify-center text-white font-extrabold text-xl shadow-md shadow-sky-600/20 shrink-0">
          S
        </div>
        <div>
          <div class="flex items-center space-x-2 flex-wrap">
            <h1 class="text-base sm:text-lg font-extrabold tracking-tight text-[var(--foreground,#0f172a)]">SNAPRINT DIGITAL PRINTING</h1>
            <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-sky-500/10 text-sky-600 border border-sky-500/20">TAHUN BUKU 2024</span>
          </div>
          <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
            Aplikasi Dashboard Kinerja Penjualan Dinamis &amp; Terintegrasi (12 Bulan Buku ACCURATE)
          </p>
        </div>
      </div>

      <div class="flex items-center space-x-2 sm:space-x-3 self-end md:self-center">
        <button onclick="resetAllFilters()" class="px-3 py-1.5 text-xs font-semibold rounded-lg bg-[var(--card,#ffffff)] hover:bg-[var(--border,#e2e8f0)] text-[var(--foreground,#0f172a)] border border-[var(--border,#cbd5e1)] transition-all flex items-center space-x-1.5 shadow-sm">
          <span>🔄</span>
          <span>Reset Semua Filter</span>
        </button>
        <a href="Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf" target="_blank" class="px-3.5 py-1.5 text-xs font-semibold rounded-lg bg-sky-600 hover:bg-sky-500 text-white transition-all shadow-sm flex items-center space-x-1.5">
          <span>📄</span>
          <span>Unduh PDF Resmi</span>
        </a>
      </div>
    </header>

    <!-- DYNAMIC FILTER CONTROL PANEL -->
    <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
      <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-2 border-b border-[var(--border,#f1f5f9)] gap-2">
        <div class="flex items-center space-x-2">
          <span class="text-sm font-bold text-[var(--foreground,#0f172a)] flex items-center space-x-1.5">
            <span>⚙️</span>
            <span>Panel Filter Dinamis &amp; Terintegrasi</span>
          </span>
          <span class="text-xs text-[var(--muted-foreground,#64748b)] hidden sm:inline">(Pilih filter atau klik batang grafik langsung untuk update instan)</span>
        </div>
        <div id="filter-active-summary" class="text-xs font-semibold text-sky-600 bg-sky-50 px-2.5 py-1 rounded-md border border-sky-200">
          Filter Aktif: Semua (1 Tahun Penuh)
        </div>
      </div>

      <!-- Filter Controls Grid -->
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        
        <!-- Filter 1: Periode Waktu -->
        <div>
          <label class="block text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase mb-1">Periode Waktu / Kuartal</label>
          <select id="filter-period" onchange="setPeriod(this.value)" class="w-full text-xs p-2.5 rounded-xl border border-[var(--border,#cbd5e1)] bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
            <option value="ALL">📅 Semua (1 Tahun Penuh - 12 Bulan)</option>
            <option value="H1">🌱 Semester 1 (Jan - Jun 2024)</option>
            <option value="H2">🚀 Semester 2 (Jul - Des 2024)</option>
            <option value="Q1">1️⃣ Kuartal 1 (Januari - Maret)</option>
            <option value="Q2">2️⃣ Kuartal 2 (April - Juni)</option>
            <option value="Q3">3️⃣ Kuartal 3 (Juli - September)</option>
            <option value="Q4">4️⃣ Kuartal 4 (Oktober - Desember)</option>
            <option disabled>────────── Bulan Tertentu ──────────</option>
            <option value="Jan">Januari 2024</option>
            <option value="Feb">Februari 2024</option>
            <option value="Mar">Maret 2024</option>
            <option value="Apr">April 2024</option>
            <option value="Mei">Mei 2024</option>
            <option value="Jun">Juni 2024</option>
            <option value="Jul">Juli 2024</option>
            <option value="Agu">Agustus 2024 (Rekor Tertinggi)</option>
            <option value="Sep">September 2024</option>
            <option value="Okt">Oktober 2024 (Puncak Kedua)</option>
            <option value="Nov">November 2024</option>
            <option value="Des">Desember 2024</option>
          </select>
        </div>

        <!-- Filter 2: Kategori Produk -->
        <div>
          <label class="block text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase mb-1">Kategori Produk / Layanan</label>
          <select id="filter-category" onchange="setCategory(this.value)" class="w-full text-xs p-2.5 rounded-xl border border-[var(--border,#cbd5e1)] bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
            <option value="ALL">🏷️ Semua Kategori (6 Lini Bisnis)</option>
            <option value="Large Format & Outdoor Banner">Large Format &amp; Outdoor Banner (39,6%)</option>
            <option value="Digital Print & Fotocopy Dokumen">Digital Print &amp; Fotocopy Dokumen (29,3%)</option>
            <option value="Kertas, Sticker Sheet & Bahan Cetak">Kertas, Sticker Sheet &amp; Bahan (10,7%)</option>
            <option value="Jasa Desain, Percetakan Lain & Perlengkapan">Jasa Desain &amp; Cetak Lain (9,6%)</option>
            <option value="Finishing, Jilid & Laminasi">Finishing, Jilid &amp; Laminasi (5,8%)</option>
            <option value="Merchandise & Produk Custom">Merchandise &amp; Produk Custom (5,1%)</option>
          </select>
        </div>

        <!-- Filter 3: Pareto Tier -->
        <div>
          <label class="block text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase mb-1">Klasifikasi Pareto 80/20</label>
          <select id="filter-pareto" onchange="setPareto(this.value)" class="w-full text-xs p-2.5 rounded-xl border border-[var(--border,#cbd5e1)] bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
            <option value="ALL">⭐ Semua Produk (129 Item)</option>
            <option value="CORE80">🏆 Core 80% Omzet (23 Produk Inti)</option>
            <option value="LONGTAIL">📦 Produk Pelengkap / Long-Tail (106 Item)</option>
          </select>
        </div>

        <!-- Filter 4: Real-time Search -->
        <div>
          <label class="block text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase mb-1">Cari Nama Barang / Jasa</label>
          <div class="relative">
            <input type="text" id="filter-search" oninput="setSearch(this.value)" placeholder="Ketik e.g. Flexy, Print A3, Cutting..." class="w-full text-xs p-2.5 pl-8 rounded-xl border border-[var(--border,#cbd5e1)] bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
            <span class="absolute left-2.5 top-2.5 text-slate-400">🔍</span>
          </div>
        </div>

      </div>
    </div>

    <!-- DYNAMIC KPI METRICS CARDS -->
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
      
      <!-- Card 1: Omzet Terpilih -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm hover:shadow transition-shadow">
        <div class="flex items-center justify-between">
          <span class="text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase tracking-wider">Omzet Penjualan Terpilih</span>
          <span class="w-7 h-7 rounded-lg bg-sky-50 text-sky-600 flex items-center justify-center font-bold text-sm">💰</span>
        </div>
        <div class="mt-2.5">
          <h2 class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-sales">Rp 930.847.260</h2>
          <p class="text-xs text-emerald-600 font-semibold mt-1 flex items-center space-x-1" id="kpi-sales-sub">
            <span>●</span>
            <span>100.00% dari Total Tahun 2024</span>
          </p>
        </div>
      </div>

      <!-- Card 2: Volume Terpilih -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm hover:shadow transition-shadow">
        <div class="flex items-center justify-between">
          <span class="text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase tracking-wider">Total Volume Terjual</span>
          <span class="w-7 h-7 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-sm">📦</span>
        </div>
        <div class="mt-2.5">
          <h2 class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-qty">275.498,79</h2>
          <p class="text-xs text-indigo-600 font-semibold mt-1" id="kpi-qty-sub">
            Unit / Meter Lari / Lembar Cetak
          </p>
        </div>
      </div>

      <!-- Card 3: Jumlah Varian Aktif -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm hover:shadow transition-shadow">
        <div class="flex items-center justify-between">
          <span class="text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase tracking-wider">Varian Produk Aktif</span>
          <span class="w-7 h-7 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center font-bold text-sm">🏷️</span>
        </div>
        <div class="mt-2.5">
          <h2 class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-items">129 Item</h2>
          <p class="text-xs text-amber-600 font-semibold mt-1" id="kpi-items-sub">
            Rata-rata: Rp 7,22 Jt / Varian Item
          </p>
        </div>
      </div>

      <!-- Card 4: Benchmark Rata-Rata Bulanan -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm hover:shadow transition-shadow">
        <div class="flex items-center justify-between">
          <span class="text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase tracking-wider">Rata-Rata Bulanan (Ref)</span>
          <span class="w-7 h-7 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold text-sm">📈</span>
        </div>
        <div class="mt-2.5">
          <h2 class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-benchmark">Rp 77.570.605</h2>
          <p class="text-xs text-slate-500 font-semibold mt-1" id="kpi-benchmark-sub">
            Puncak: Agustus (Rp 124,07 Jt)
          </p>
        </div>
      </div>

    </div>

    <!-- CHARTS & VISUALIZATIONS SECTION -->
    <div class="grid grid-cols-1 lg:grid-cols-3 gap-4 sm:gap-6">
      
      <!-- Chart 1: Interactive Monthly Trend (2 Cols) -->
      <div class="lg:col-span-2 bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm flex flex-col justify-between">
        
        <div class="flex items-start sm:items-center justify-between mb-2">
          <div>
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>📊</span>
              <span>Tren Omzet Penjualan Bulanan (Januari - Desember 2024)</span>
            </h3>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Setiap batang menunjukkan angka penjualan bersih (Rupiah) &amp; porsi persentase. Klik batang untuk filter!
            </p>
          </div>
          <span class="text-xs px-2.5 py-1 rounded-full bg-sky-50 text-sky-700 font-bold border border-sky-200 shrink-0">
            12 Bulan Buku
          </span>
        </div>

        <!-- Dynamic Live Hover Info Box -->
        <div id="chart-hover-info" class="bg-slate-50 border border-slate-200/80 rounded-xl px-3.5 py-2 text-xs text-slate-700 flex flex-wrap items-center justify-between gap-2 min-h-[38px] my-1">
          <span class="font-medium text-slate-500">💡 Sentuh/arahkan kursor pada batang bulan untuk melihat rincian omzet, volume, &amp; performa:</span>
          <span class="text-sky-600 font-semibold cursor-pointer hover:underline" onclick="setPeriod('ALL')">Tampilkan Semua Bulan</span>
        </div>

        <!-- High-Headroom SVG Bar Chart -->
        <div class="w-full overflow-x-auto py-1">
          <svg id="svg-monthly-chart" viewBox="0 0 860 300" class="w-full h-64 sm:h-72 select-none"></svg>
        </div>

        <!-- Legend Bar -->
        <div class="mt-2 pt-2 border-t border-[var(--border,#f1f5f9)] flex flex-wrap items-center justify-between gap-2 text-xs text-[var(--muted-foreground,#64748b)]">
          <div class="flex flex-wrap items-center gap-3">
            <span class="flex items-center space-x-1.5">
              <span class="w-3.5 h-3.5 rounded-md bg-gradient-to-t from-sky-700 to-sky-500 inline-block shadow-sm"></span>
              <span class="font-medium text-slate-700">Omzet Bulan Terpilih</span>
            </span>
            <span class="flex items-center space-x-1.5">
              <span class="w-3.5 h-3.5 rounded-md bg-amber-400 inline-block shadow-sm"></span>
              <span class="font-medium text-amber-800">Bulan Puncak (Agustus)</span>
            </span>
            <span class="flex items-center space-x-1.5">
              <span class="w-3.5 h-3.5 rounded-md bg-slate-300 inline-block"></span>
              <span>Bulan Lain</span>
            </span>
            <span class="flex items-center space-x-1.5">
              <span class="w-4 h-0.5 bg-rose-500 inline-block"></span>
              <span class="text-rose-600 font-semibold">Garis Rata-rata (Rp 77,57 Jt)</span>
            </span>
          </div>
          <button onclick="setPeriod('ALL')" class="text-xs text-sky-600 hover:text-sky-700 font-bold">
            Reset ke 1 Tahun
          </button>
        </div>

      </div>

      <!-- Chart 2: Dynamic Category Breakdown (1 Col) -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm flex flex-col justify-between">
        <div>
          <div class="flex items-center justify-between mb-2">
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>🥧</span>
              <span>Porsi Kategori Produk</span>
            </h3>
            <span class="text-xs text-[var(--muted-foreground,#64748b)]">6 Kategori</span>
          </div>
          <p class="text-xs text-[var(--muted-foreground,#64748b)] mb-3">
            Klik baris kategori di bawah untuk memfilter katalog produk:
          </p>

          <!-- Category Progress Bars -->
          <div class="space-y-2.5" id="category-progress-container">
            <!-- Populated via JS -->
          </div>
        </div>

        <div class="mt-4 pt-3 border-t border-[var(--border,#f1f5f9)] text-xs text-[var(--muted-foreground,#64748b)] flex items-center justify-between">
          <span>Dominasi Terbesar:</span>
          <span class="font-bold text-sky-700">Large Format (39,6%)</span>
        </div>
      </div>

    </div>

    <!-- DYNAMIC PRODUCT RANKINGS TABLE -->
    <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
      
      <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-3 border-b border-[var(--border,#f1f5f9)] gap-2">
        <div>
          <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
            <span>🏆</span>
            <span>Peringkat &amp; Rincian Produk Sesuai Filter</span>
          </h3>
          <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5" id="table-subtitle">
            Menampilkan produk aktif berdasarkan filter waktu &amp; kategori yang dipilih.
          </p>
        </div>
        <div class="flex items-center space-x-2 text-xs">
          <span class="text-[var(--muted-foreground,#64748b)]">💡 Klik baris produk untuk riwayat bulanan</span>
        </div>
      </div>

      <!-- Table Container -->
      <div class="overflow-x-auto max-h-[500px] border border-[var(--border,#e2e8f0)] rounded-xl">
        <table class="w-full text-left text-xs">
          <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider sticky top-0 z-10 shadow-sm">
            <tr>
              <th class="py-2.5 px-3 text-center w-12 cursor-pointer" onclick="changeSort('rank')">No ⇅</th>
              <th class="py-2.5 px-3 cursor-pointer" onclick="changeSort('name')">Nama Barang / Produk Layanan ⇅</th>
              <th class="py-2.5 px-3 cursor-pointer" onclick="changeSort('category')">Kategori ⇅</th>
              <th class="py-2.5 px-3 text-center w-16">Satuan</th>
              <th class="py-2.5 px-3 text-right cursor-pointer" onclick="changeSort('qty')">Volume Terjual ⇅</th>
              <th class="py-2.5 px-3 text-right cursor-pointer" onclick="changeSort('sales')">Total Omzet (Rp) ⇅</th>
              <th class="py-2.5 px-3 text-right cursor-pointer" onclick="changeSort('pct')">Porsi (%) ⇅</th>
              <th class="py-2.5 px-3 text-center">Klasifikasi</th>
            </tr>
          </thead>
          <tbody id="dynamic-products-tbody" class="divide-y divide-[var(--border,#f1f5f9)] text-[var(--foreground,#0f172a)] font-medium">
            <!-- Populated via JS -->
          </tbody>
          <tfoot class="bg-slate-100 font-bold text-slate-900 sticky bottom-0 z-10 border-t-2 border-slate-300">
            <tr>
              <td class="py-2.5 px-3 text-center" colspan="4">SUBTOTAL SESUAI FILTER AKTIF</td>
              <td class="py-2.5 px-3 text-right font-mono" id="foot-filtered-qty">275.498,79</td>
              <td class="py-2.5 px-3 text-right font-mono" id="foot-filtered-sales">Rp 930.847.260,01</td>
              <td class="py-2.5 px-3 text-right font-mono" id="foot-filtered-pct">100.00%</td>
              <td class="py-2.5 px-3 text-center" id="foot-filtered-count">129 Item</td>
            </tr>
          </tfoot>
        </table>
      </div>

    </div>

    <!-- EXECUTIVE STRATEGY INSIGHTS ACCORDION -->
    <div class="bg-gradient-to-r from-slate-900 to-indigo-950 text-white rounded-2xl p-5 shadow-sm border border-slate-800">
      <div class="flex items-center space-x-2.5 mb-3">
        <span class="text-lg">💡</span>
        <h3 class="font-bold text-sm sm:text-base text-sky-300">Kesimpulan Eksekutif Kinerja 2024 &amp; Rekomendasi 2025</h3>
      </div>
      <div class="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs text-slate-300">
        <div class="p-3 bg-white/5 rounded-xl border border-white/10">
          <span class="font-bold text-white block mb-1">1. Momentum Musiman Kuat:</span>
          Penjualan semester 2 (Rp 645,26 Jt) melompat <b>+125,9%</b> dibanding semester 1 (Rp 285,59 Jt). Puncak penjualan berada di bulan Agustus (Rp 124 Jt) dan Oktober (Rp 120 Jt).
        </div>
        <div class="p-3 bg-white/5 rounded-xl border border-white/10">
          <span class="font-bold text-white block mb-1">2. Kekuatan Core 80% Pareto:</span>
          Hanya dengan <b>23 produk utama (17,8%)</b>, Snaprint berhasil membukukan 80% omzet. Prioritas pengadaan bahan baku seperti Flexy 280 dan Print A3 wajib diamankan secara kontrak.
        </div>
        <div class="p-3 bg-white/5 rounded-xl border border-white/10">
          <span class="font-bold text-white block mb-1">3. Sinergi Finishing &amp; Paket:</span>
          Layanan cutting A3, laminasi, dan jilid memiliki margin tinggi. Paket bundling "Cetak + Cutting + Laminasi" terbukti meningkatkan nilai transaksi rata-rata per pelanggan.
        </div>
      </div>
    </div>

  </div>

  <!-- PRODUCT DETAIL POPUP MODAL -->
  <div id="product-detail-modal" class="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-sm hidden flex items-center justify-center p-3 sm:p-4">
    <div class="bg-[var(--card,#ffffff)] text-[var(--foreground,#0f172a)] w-full max-w-2xl rounded-2xl shadow-2xl border border-[var(--border,#e2e8f0)] overflow-hidden flex flex-col max-h-[90vh]">
      
      <!-- Modal Header -->
      <div class="p-4 sm:p-5 bg-slate-900 text-white flex items-start justify-between">
        <div>
          <div class="flex items-center space-x-2">
            <span class="px-2.5 py-0.5 rounded text-[11px] font-bold bg-sky-500/20 text-sky-300 border border-sky-400/30" id="modal-rank">RANK #1</span>
            <span class="text-xs text-slate-300 font-medium" id="modal-cat">Kategori</span>
          </div>
          <h2 class="text-base sm:text-lg font-bold text-white mt-1.5" id="modal-name">Nama Produk</h2>
        </div>
        <button onclick="closeModal()" class="text-slate-400 hover:text-white p-1 rounded-lg text-lg">✕</button>
      </div>

      <!-- Modal Body -->
      <div class="p-4 sm:p-5 overflow-y-auto space-y-4">
        
        <!-- Quick KPI Grid -->
        <div class="grid grid-cols-3 gap-3 text-center">
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
            <span class="text-[11px] text-slate-500 font-semibold uppercase">Total Omzet 2024</span>
            <p class="text-sm sm:text-base font-bold text-slate-900 font-mono mt-0.5" id="modal-sales">-</p>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
            <span class="text-[11px] text-slate-500 font-semibold uppercase">Total Volume</span>
            <p class="text-sm sm:text-base font-bold text-slate-900 font-mono mt-0.5" id="modal-qty">-</p>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
            <span class="text-[11px] text-slate-500 font-semibold uppercase">Porsi terhadap Omzet</span>
            <p class="text-sm sm:text-base font-bold text-sky-600 font-mono mt-0.5" id="modal-pct">-</p>
          </div>
        </div>

        <!-- Monthly SVG Spark Chart -->
        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
          <span class="text-xs font-bold text-slate-700 block mb-1">Riwayat Penjualan 12 Bulan (Januari - Desember 2024)</span>
          <svg id="modal-spark-chart" viewBox="0 0 540 100" class="w-full h-28"></svg>
        </div>

        <!-- Monthly Table -->
        <div class="border border-slate-200 rounded-xl overflow-hidden">
          <table class="w-full text-xs text-left">
            <thead class="bg-slate-100 font-semibold text-slate-700 uppercase text-[10px]">
              <tr>
                <th class="py-1.5 px-3">Bulan</th>
                <th class="py-1.5 px-3 text-right">Volume</th>
                <th class="py-1.5 px-3 text-right">Penjualan (Rp)</th>
              </tr>
            </thead>
            <tbody id="modal-month-tbody" class="divide-y divide-slate-100 font-mono text-slate-800">
              <!-- Populated via JS -->
            </tbody>
          </table>
        </div>

      </div>

      <!-- Modal Footer -->
      <div class="p-3.5 bg-slate-50 border-t border-slate-200 flex justify-end">
        <button onclick="closeModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold rounded-lg transition-colors">
          Tutup
        </button>
      </div>

    </div>
  </div>

  <!-- EMBEDDED COMPLETE DATASET -->
  <script>
    const RAW_DATA = {json_str};

    // Global State for Dynamic Cross-Filtering
    const state = {{
      period: 'ALL',      // 'ALL', 'H1', 'H2', 'Q1'..'Q4', or month 'Jan'..'Des'
      category: 'ALL',    // 'ALL' or specific category name
      pareto: 'ALL',      // 'ALL', 'CORE80', 'LONGTAIL'
      search: '',
      sortCol: 'sales',
      sortAsc: false
    }};

    const MONTH_KEYS = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    const MONTH_NAMES = {{
      'Jan': 'Januari', 'Feb': 'Februari', 'Mar': 'Maret', 'Apr': 'April',
      'Mei': 'Mei', 'Jun': 'Juni', 'Jul': 'Juli', 'Agu': 'Agustus',
      'Sep': 'September', 'Okt': 'Oktober', 'Nov': 'November', 'Des': 'Desember'
    }};

    function fmtIDR(num) {{
      return 'Rp ' + Math.round(num).toLocaleString('id-ID');
    }}

    function fmtNum(num) {{
      return Number(num).toLocaleString('id-ID', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
    }}

    // Get Active Month Keys based on state.period
    function getActiveMonths() {{
      const p = state.period;
      if (p === 'ALL') return MONTH_KEYS;
      if (p === 'H1') return ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun'];
      if (p === 'H2') return ['Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
      if (p === 'Q1') return ['Jan', 'Feb', 'Mar'];
      if (p === 'Q2') return ['Apr', 'Mei', 'Jun'];
      if (p === 'Q3') return ['Jul', 'Agu', 'Sep'];
      if (p === 'Q4') return ['Okt', 'Nov', 'Des'];
      if (MONTH_KEYS.includes(p)) return [p];
      return MONTH_KEYS;
    }}

    // Filter Change Handlers
    function setPeriod(val) {{
      state.period = val;
      document.getElementById('filter-period').value = val;
      updateDashboard();
    }}

    function setCategory(val) {{
      state.category = val;
      document.getElementById('filter-category').value = val;
      updateDashboard();
    }}

    function setPareto(val) {{
      state.pareto = val;
      document.getElementById('filter-pareto').value = val;
      updateDashboard();
    }}

    function setSearch(val) {{
      state.search = val.toLowerCase().trim();
      updateDashboard();
    }}

    function resetAllFilters() {{
      state.period = 'ALL';
      state.category = 'ALL';
      state.pareto = 'ALL';
      state.search = '';
      document.getElementById('filter-period').value = 'ALL';
      document.getElementById('filter-category').value = 'ALL';
      document.getElementById('filter-pareto').value = 'ALL';
      document.getElementById('filter-search').value = '';
      updateDashboard();
    }}

    function changeSort(col) {{
      if (state.sortCol === col) {{
        state.sortAsc = !state.sortAsc;
      }} else {{
        state.sortCol = col;
        state.sortAsc = (col === 'name' || col === 'category') ? true : false;
      }}
      renderProductsTable();
    }}

    // Main Update Function
    function updateDashboard() {{
      updateFilterSummaryLabel();
      computeAndRenderMetrics();
      renderMonthlyChart();
      renderCategoryBars();
      renderProductsTable();
    }}

    function updateFilterSummaryLabel() {{
      let parts = [];
      if (state.period !== 'ALL') {{
        parts.push(`Periode: ${{state.period}}`);
      }}
      if (state.category !== 'ALL') {{
        parts.push(`Kategori: ${{state.category}}`);
      }}
      if (state.pareto !== 'ALL') {{
        parts.push(state.pareto === 'CORE80' ? 'Klasifikasi: Core 80%' : 'Klasifikasi: Long-tail');
      }}
      if (state.search) {{
        parts.push(`Cari: "${{state.search}}"`);
      }}
      const el = document.getElementById('filter-active-summary');
      if (parts.length === 0) {{
        el.innerText = 'Filter Aktif: Semua (1 Tahun Penuh - 12 Bulan)';
        el.className = 'text-xs font-semibold text-sky-600 bg-sky-50 px-2.5 py-1 rounded-md border border-sky-200';
      }} else {{
        el.innerText = 'Filter Aktif: ' + parts.join(' | ');
        el.className = 'text-xs font-semibold text-amber-700 bg-amber-50 px-2.5 py-1 rounded-md border border-amber-300';
      }}
    }}

    // Compute Derived Data
    let currentFilteredProducts = [];
    let currentTotalSales = 0;
    let currentTotalQty = 0;

    function computeAndRenderMetrics() {{
      const activeMonths = getActiveMonths();
      const annualTotalSales = RAW_DATA.kpi.total_sales;

      currentFilteredProducts = [];
      currentTotalSales = 0;
      currentTotalQty = 0;

      RAW_DATA.products.forEach(p => {{
        // Check category filter
        if (state.category !== 'ALL' && p.category !== state.category) return;

        // Check pareto filter
        if (state.pareto === 'CORE80' && !p.is_pareto_80) return;
        if (state.pareto === 'LONGTAIL' && p.is_pareto_80) return;

        // Check search filter
        if (state.search && !p.name.toLowerCase().includes(state.search) && !p.category.toLowerCase().includes(state.search)) return;

        // Calculate sales & qty in active months
        let pSales = 0;
        let pQty = 0;
        activeMonths.forEach(m => {{
          pSales += (p.monthly_sales[m] || 0);
          pQty += (p.monthly_qty[m] || 0);
        }});

        if (pSales > 0 || pQty > 0 || activeMonths.length === 12) {{
          currentFilteredProducts.push({{
            ...p,
            period_sales: pSales,
            period_qty: pQty
          }});
          currentTotalSales += pSales;
          currentTotalQty += pQty;
        }}
      }});

      // Update KPI Cards
      document.getElementById('kpi-sales').innerText = fmtIDR(currentTotalSales);
      const pctOfAnnual = annualTotalSales > 0 ? (currentTotalSales / annualTotalSales) * 100 : 0;
      document.getElementById('kpi-sales-sub').innerHTML = `<span>●</span><span>${{pctOfAnnual.toFixed(2)}}% dari Total Tahunan (Rp 930,85 Jt)</span>`;

      document.getElementById('kpi-qty').innerText = fmtNum(currentTotalQty);
      document.getElementById('kpi-items').innerText = `${{currentFilteredProducts.length}} Item`;
      const avgPerItem = currentFilteredProducts.length > 0 ? currentTotalSales / currentFilteredProducts.length : 0;
      document.getElementById('kpi-items-sub').innerText = `Rata-rata: ${{fmtIDR(avgPerItem)}} / Item`;
    }}

    // ========================================================
    // HIGH-HEADROOM, PERFECTLY READABLE SVG MONTHLY CHART
    // ========================================================
    function renderMonthlyChart() {{
      const svg = document.getElementById('svg-monthly-chart');
      svg.innerHTML = '';

      const activeMonths = getActiveMonths();
      const monthlyData = RAW_DATA.monthly;
      const annualAvgSales = RAW_DATA.kpi.avg_monthly_sales; // ~77.570.605

      // Geometry & Dimensions
      const viewBoxWidth = 860;
      const viewBoxHeight = 300;
      
      const chartLeft = 70;
      const chartRight = 835;
      const chartWidth = chartRight - chartLeft; // 765px
      
      const chartTop = 50;       // 50px top headroom for badges (never clipped!)
      const chartBottom = 240;    // 60px bottom margin for labels
      const chartHeight = chartBottom - chartTop; // 190px
      
      const maxScale = 145000000; // Rp 145 Juta ceiling (Agustus 124 Jt fits at 85% height with 45px free headroom)

      const numBars = 12;
      const barWidth = 46;
      const gap = (chartWidth - (numBars * barWidth)) / (numBars + 1);

      // SVG Defs (Gradients & Shadows)
      const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
      defs.innerHTML = `
        <linearGradient id="grad-bar-active" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#0284c7" />
          <stop offset="100%" stop-color="#0369a1" />
        </linearGradient>
        <linearGradient id="grad-bar-peak" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#0284c7" />
          <stop offset="100%" stop-color="#1e3a8a" />
        </linearGradient>
        <linearGradient id="grad-bar-muted" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#cbd5e1" />
          <stop offset="100%" stop-color="#94a3b8" />
        </linearGradient>
        <filter id="badge-shadow" x="-10%" y="-10%" width="120%" height="130%">
          <feDropShadow dx="0" dy="1.5" stdDeviation="1.5" flood-color="#0f172a" flood-opacity="0.08"/>
        </filter>
      `;
      svg.appendChild(defs);

      // 1. Draw Horizontal Gridlines & Y-Axis Scale
      const gridLevels = [
        {{ val: 125000000, label: 'Rp 125 Jt' }},
        {{ val: 100000000, label: 'Rp 100 Jt' }},
        {{ val: 75000000,  label: 'Rp 75 Jt' }},
        {{ val: 50000000,  label: 'Rp 50 Jt' }},
        {{ val: 25000000,  label: 'Rp 25 Jt' }}
      ];

      gridLevels.forEach(lvl => {{
        const yPos = chartBottom - (lvl.val / maxScale) * chartHeight;
        
        // Gridline
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', chartLeft);
        line.setAttribute('y1', yPos);
        line.setAttribute('x2', chartRight);
        line.setAttribute('y2', yPos);
        line.setAttribute('stroke', '#e2e8f0');
        line.setAttribute('stroke-dasharray', '3,3');
        line.setAttribute('stroke-width', '1');
        svg.appendChild(line);

        // Y-Axis label
        const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt.setAttribute('x', chartLeft - 8);
        txt.setAttribute('y', yPos + 3.5);
        txt.setAttribute('text-anchor', 'end');
        txt.setAttribute('font-size', '9.5');
        txt.setAttribute('font-weight', '600');
        txt.setAttribute('fill', '#94a3b8');
        txt.textContent = lvl.label;
        svg.appendChild(txt);
      }});

      // Baseline (0 line)
      const baseLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      baseLine.setAttribute('x1', chartLeft);
      baseLine.setAttribute('y1', chartBottom);
      baseLine.setAttribute('x2', chartRight);
      baseLine.setAttribute('y2', chartBottom);
      baseLine.setAttribute('stroke', '#94a3b8');
      baseLine.setAttribute('stroke-width', '1.5');
      svg.appendChild(baseLine);

      // Zero Label
      const zeroTxt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      zeroTxt.setAttribute('x', chartLeft - 8);
      zeroTxt.setAttribute('y', chartBottom + 3.5);
      zeroTxt.setAttribute('text-anchor', 'end');
      zeroTxt.setAttribute('font-size', '9.5');
      zeroTxt.setAttribute('font-weight', '600');
      zeroTxt.setAttribute('fill', '#94a3b8');
      zeroTxt.textContent = 'Rp 0';
      svg.appendChild(zeroTxt);

      // 2. Draw Average Benchmark Line (Rp 77,57 Jt)
      const avgY = chartBottom - (annualAvgSales / maxScale) * chartHeight;
      const avgLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      avgLine.setAttribute('x1', chartLeft);
      avgLine.setAttribute('y1', avgY);
      avgLine.setAttribute('x2', chartRight - 130);
      avgLine.setAttribute('y2', avgY);
      avgLine.setAttribute('stroke', '#f43f5e');
      avgLine.setAttribute('stroke-dasharray', '5,3');
      avgLine.setAttribute('stroke-width', '1.5');
      svg.appendChild(avgLine);

      // Average Line Tag on far right
      const avgTagG = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      avgTagG.innerHTML = `
        <rect x="${{chartRight - 128}}" y="${{avgY - 10}}" width="128" height="20" rx="5" fill="#fff1f2" stroke="#fda4af" stroke-width="1" />
        <text x="${{chartRight - 64}}" y="${{avgY + 4}}" font-size="9.5" font-weight="700" fill="#e11d48" text-anchor="middle">Rata-rata: Rp 77,57 Jt</text>
      `;
      svg.appendChild(avgTagG);

      // 3. Draw Bars & Crisp Value Badges
      monthlyData.forEach((m, i) => {{
        const x = chartLeft + gap + i * (barWidth + gap);
        const barH = (m.sales / maxScale) * chartHeight;
        const barTop = chartBottom - barH;
        const centerX = x + barWidth / 2;

        const isSelected = activeMonths.includes(m.code);
        const isPeak = (m.code === 'Agu'); // Peak month
        const isSecondPeak = (m.code === 'Okt');

        const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        g.className.baseVal = 'chart-bar-group cursor-pointer';

        // Hover Handlers
        g.onmouseenter = () => {{
          const diffPct = ((m.sales - annualAvgSales) / annualAvgSales) * 100;
          const statusStr = diffPct >= 0 
            ? `<span class="text-emerald-700 font-bold">+${{diffPct.toFixed(1)}}% di atas rata-rata</span>` 
            : `<span class="text-rose-700 font-bold">${{diffPct.toFixed(1)}}% di bawah rata-rata</span>`;
          
          document.getElementById('chart-hover-info').innerHTML = `
            <div class="flex items-center space-x-2">
              <span class="font-extrabold text-slate-900 text-xs">📅 ${{m.name}} 2024:</span>
              <span class="font-bold text-sky-700 font-mono text-xs">${{fmtIDR(m.sales)}}</span>
              <span class="text-slate-400">|</span>
              <span class="text-slate-600 text-xs">Volume: <b>${{fmtNum(m.qty)}} unit</b></span>
              <span class="text-slate-400">|</span>
              <span class="text-slate-600 text-xs">Porsi: <b>${{m.pct.toFixed(2)}}%</b></span>
            </div>
            <div>${{statusStr}}</div>
          `;
        }};

        g.onmouseleave = () => {{
          document.getElementById('chart-hover-info').innerHTML = `
            <span class="font-medium text-slate-500">💡 Sentuh/arahkan kursor pada batang bulan untuk melihat rincian omzet, volume, &amp; performa:</span>
            <span class="text-sky-600 font-semibold cursor-pointer hover:underline" onclick="setPeriod('ALL')">Tampilkan Semua Bulan</span>
          `;
        }};

        g.onclick = () => setPeriod(m.code);

        // Bar Fill
        const barFill = isPeak && isSelected 
          ? 'url(#grad-bar-peak)' 
          : isSelected 
            ? 'url(#grad-bar-active)' 
            : 'url(#grad-bar-muted)';

        // Bar Rectangle
        const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect.setAttribute('x', x);
        rect.setAttribute('y', barTop);
        rect.setAttribute('width', barWidth);
        rect.setAttribute('height', barH);
        rect.setAttribute('rx', '6');
        rect.setAttribute('fill', barFill);
        if (isSelected) {{
          rect.setAttribute('stroke', isPeak ? '#f59e0b' : '#0284c7');
          rect.setAttribute('stroke-width', '1.5');
        }}
        g.appendChild(rect);

        // ========================================================
        // HIGH-CONTRAST VALUE PILL CARD ABOVE BAR (NEVER CLIPPED!)
        // ========================================================
        const badgeWidth = 52;
        const badgeHeight = 27;
        const badgeX = centerX - badgeWidth / 2;
        const badgeY = barTop - badgeHeight - 5; // 5px gap above bar

        const badgeG = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        badgeG.setAttribute('filter', 'url(#badge-shadow)');

        const badgeBg = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        badgeBg.setAttribute('x', badgeX);
        badgeBg.setAttribute('y', badgeY);
        badgeBg.setAttribute('width', badgeWidth);
        badgeBg.setAttribute('height', badgeHeight);
        badgeBg.setAttribute('rx', '6');

        if (isPeak) {{
          badgeBg.setAttribute('fill', '#fef3c7'); // Amber 100
          badgeBg.setAttribute('stroke', '#f59e0b');
          badgeBg.setAttribute('stroke-width', '1.5');
        }} else if (isSecondPeak) {{
          badgeBg.setAttribute('fill', '#f0f9ff'); // Sky 50
          badgeBg.setAttribute('stroke', '#0284c7');
          badgeBg.setAttribute('stroke-width', '1.2');
        }} else if (isSelected) {{
          badgeBg.setAttribute('fill', '#ffffff');
          badgeBg.setAttribute('stroke', '#cbd5e1');
          badgeBg.setAttribute('stroke-width', '1');
        }} else {{
          badgeBg.setAttribute('fill', '#ffffff');
          badgeBg.setAttribute('stroke', '#e2e8f0');
          badgeBg.setAttribute('stroke-width', '1');
        }}
        badgeG.appendChild(badgeBg);

        // Line 1: Omzet Value (e.g. 124,1 Jt)
        const textVal = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        textVal.setAttribute('x', centerX);
        textVal.setAttribute('y', badgeY + 12);
        textVal.setAttribute('text-anchor', 'middle');
        textVal.setAttribute('font-size', '10');
        textVal.setAttribute('font-weight', '800');
        textVal.setAttribute('fill', isPeak ? '#92400e' : isSelected ? '#0f172a' : '#64748b');
        textVal.textContent = (m.sales / 1e6).toFixed(1) + ' Jt';
        badgeG.appendChild(textVal);

        // Line 2: Share % or Peak badge
        const textPct = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        textPct.setAttribute('x', centerX);
        textPct.setAttribute('y', badgeY + 22);
        textPct.setAttribute('text-anchor', 'middle');
        textPct.setAttribute('font-size', '8.5');
        textPct.setAttribute('font-weight', '700');
        textPct.setAttribute('fill', isPeak ? '#b45309' : isSelected ? '#0284c7' : '#94a3b8');
        textPct.textContent = isPeak ? '👑 Puncak' : `${{m.pct.toFixed(1)}}%`;
        badgeG.appendChild(textPct);

        g.appendChild(badgeG);

        // ========================================================
        // X-AXIS MONTH LABEL & VOLUME BELOW BAR
        // ========================================================
        // Month Name Badge
        const textMonth = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        textMonth.setAttribute('x', centerX);
        textMonth.setAttribute('y', chartBottom + 18);
        textMonth.setAttribute('text-anchor', 'middle');
        textMonth.setAttribute('font-size', '11.5');
        textMonth.setAttribute('font-weight', isSelected ? '800' : '600');
        textMonth.setAttribute('fill', isSelected ? '#0f172a' : '#64748b');
        textMonth.textContent = m.code;
        g.appendChild(textMonth);

        // Volume text below month
        const textQty = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        textQty.setAttribute('x', centerX);
        textQty.setAttribute('y', chartBottom + 32);
        textQty.setAttribute('text-anchor', 'middle');
        textQty.setAttribute('font-size', '9');
        textQty.setAttribute('font-weight', '500');
        textQty.setAttribute('fill', '#94a3b8');
        textQty.textContent = (m.qty / 1000).toFixed(1) + 'k u';
        g.appendChild(textQty);

        svg.appendChild(g);
      }});
    }}

    // Render Category Breakdown Bars
    function renderCategoryBars() {{
      const container = document.getElementById('category-progress-container');
      container.innerHTML = '';

      const activeMonths = getActiveMonths();
      const catTotals = {{}};
      RAW_DATA.categories.forEach(c => {{ catTotals[c.name] = 0; }});

      RAW_DATA.products.forEach(p => {{
        activeMonths.forEach(m => {{
          catTotals[p.category] = (catTotals[p.category] || 0) + (p.monthly_sales[m] || 0);
        }});
      }});

      const totalPeriodSales = Object.values(catTotals).reduce((a, b) => a + b, 0);

      RAW_DATA.categories.forEach(cat => {{
        const cSales = catTotals[cat.name] || 0;
        const cPct = totalPeriodSales > 0 ? (cSales / totalPeriodSales) * 100 : 0;
        const isSelected = state.category === cat.name;

        const div = document.createElement('div');
        div.className = `p-2 rounded-xl transition-all cursor-pointer border ${{
          isSelected 
            ? 'bg-sky-50 border-sky-300 ring-2 ring-sky-400' 
            : 'bg-slate-50/70 border-slate-200/70 hover:bg-slate-100'
        }}`;
        div.onclick = () => setCategory(isSelected ? 'ALL' : cat.name);

        div.innerHTML = `
          <div class="flex items-center justify-between text-xs">
            <span class="font-bold text-slate-800 truncate max-w-[190px]">${{cat.name}}</span>
            <span class="font-mono font-bold text-slate-900">${{fmtIDR(cSales)}} (${{cPct.toFixed(1)}}%)</span>
          </div>
          <div class="w-full bg-slate-200 h-2 rounded-full overflow-hidden mt-1.5">
            <div class="bg-sky-600 h-full rounded-full transition-all duration-300" style="width: ${{Math.min(cPct, 100)}}%"></div>
          </div>
        `;
        container.appendChild(div);
      }});
    }}

    // Render Products Table
    function renderProductsTable() {{
      const tbody = document.getElementById('dynamic-products-tbody');
      tbody.innerHTML = '';

      // Sorting
      currentFilteredProducts.sort((a, b) => {{
        let vA = a.period_sales;
        let vB = b.period_sales;

        if (state.sortCol === 'rank') {{ vA = a.rank; vB = b.rank; }}
        if (state.sortCol === 'name') {{ vA = a.name.toLowerCase(); vB = b.name.toLowerCase(); }}
        if (state.sortCol === 'category') {{ vA = a.category.toLowerCase(); vB = b.category.toLowerCase(); }}
        if (state.sortCol === 'qty') {{ vA = a.period_qty; vB = b.period_qty; }}
        if (state.sortCol === 'sales') {{ vA = a.period_sales; vB = b.period_sales; }}
        if (state.sortCol === 'pct') {{ vA = a.period_sales; vB = b.period_sales; }}

        if (vA < vB) return state.sortAsc ? -1 : 1;
        if (vA > vB) return state.sortAsc ? 1 : -1;
        return 0;
      }});

      currentFilteredProducts.forEach((p, idx) => {{
        const itemPct = currentTotalSales > 0 ? (p.period_sales / currentTotalSales) * 100 : 0;
        const paretoBadge = p.is_pareto_80
          ? `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-300">⭐ Core 80%</span>`
          : `<span class="px-2 py-0.5 rounded-full text-[10px] font-medium bg-slate-100 text-slate-600">Pelengkap</span>`;

        const tr = document.createElement('tr');
        tr.className = 'hover:bg-sky-50/60 cursor-pointer transition-colors';
        tr.onclick = () => openProductModal(p);

        tr.innerHTML = `
          <td class="py-2 px-3 text-center font-mono text-slate-500">${{idx + 1}}</td>
          <td class="py-2 px-3 font-semibold text-slate-900 hover:text-sky-600 transition-colors">${{p.name}}</td>
          <td class="py-2 px-3 text-slate-500 text-[11px] truncate max-w-[170px]">${{p.category}}</td>
          <td class="py-2 px-3 text-center text-slate-500 text-[11px] font-mono">${{p.unit}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtNum(p.period_qty)}}</td>
          <td class="py-2 px-3 text-right font-mono font-bold text-slate-900">${{fmtIDR(p.period_sales)}}</td>
          <td class="py-2 px-3 text-right font-mono text-slate-600">${{itemPct.toFixed(2)}}%</td>
          <td class="py-2 px-3 text-center">${{paretoBadge}}</td>
        `;
        tbody.appendChild(tr);
      }});

      // Update Footers
      document.getElementById('foot-filtered-qty').innerText = fmtNum(currentTotalQty);
      document.getElementById('foot-filtered-sales').innerText = fmtIDR(currentTotalSales);
      document.getElementById('foot-filtered-count').innerText = `${{currentFilteredProducts.length}} Item`;
    }}

    // Open Product Modal
    function openProductModal(p) {{
      document.getElementById('modal-rank').innerText = `PERINGKAT TAHUNAN #${{p.rank}}`;
      document.getElementById('modal-cat').innerText = p.category;
      document.getElementById('modal-name').innerText = p.name;
      document.getElementById('modal-sales').innerText = fmtIDR(p.total_sales);
      document.getElementById('modal-qty').innerText = `${{fmtNum(p.total_qty)}} ${{p.unit}}`;
      document.getElementById('modal-pct').innerText = `${{p.pct.toFixed(2)}}% dari Omzet 2024`;

      // Spark chart
      const svg = document.getElementById('modal-spark-chart');
      svg.innerHTML = '';
      const monthlySales = MONTH_KEYS.map(m => p.monthly_sales[m] || 0);
      const maxM = Math.max(...monthlySales, 1);
      const w = 540;
      const h = 75;
      const bW = 28;
      const gap = (w - (12 * bW)) / 13;

      MONTH_KEYS.forEach((m, idx) => {{
        const x = gap + idx * (bW + gap);
        const s = p.monthly_sales[m] || 0;
        const bH = (s / maxM) * h;
        const y = h - bH;

        const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect.setAttribute('x', x);
        rect.setAttribute('y', y);
        rect.setAttribute('width', bW);
        rect.setAttribute('height', bH);
        rect.setAttribute('rx', '3');
        rect.setAttribute('fill', s > 0 ? '#0284c7' : '#e2e8f0');
        svg.appendChild(rect);

        const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt.setAttribute('x', x + bW / 2);
        txt.setAttribute('y', h + 14);
        txt.setAttribute('text-anchor', 'middle');
        txt.setAttribute('font-size', '9');
        txt.setAttribute('fill', '#64748b');
        txt.textContent = m;
        svg.appendChild(txt);
      }});

      // Mini Table
      const tbody = document.getElementById('modal-month-tbody');
      tbody.innerHTML = '';
      MONTH_KEYS.forEach(m => {{
        const s = p.monthly_sales[m] || 0;
        const q = p.monthly_qty[m] || 0;
        if (s > 0 || q > 0) {{
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td class="py-1 px-3 font-semibold text-slate-800">${{MONTH_NAMES[m]}} 2024</td>
            <td class="py-1 px-3 text-right font-mono">${{fmtNum(q)}}</td>
            <td class="py-1 px-3 text-right font-mono font-bold text-slate-900">${{fmtIDR(s)}}</td>
          `;
          tbody.appendChild(tr);
        }}
      }});

      document.getElementById('product-detail-modal').classList.remove('hidden');
    }}

    function closeModal() {{
      document.getElementById('product-detail-modal').classList.add('hidden');
    }}

    // Initial Load
    window.addEventListener('DOMContentLoaded', () => {{
      updateDashboard();
    }});
  </script>

</body>
</html>
"""

# 1. Write to Artifact Directory
artifact_path = r"C:\Users\Swanto\.gemini\antigravity\brain\186463fa-99d6-481c-8f3f-6ce410e3a005\dashboard_snaprint_2024.html"
with open(artifact_path, "w", encoding="utf-8") as f:
    f.write(html_template)
print(f"Generated artifact: {artifact_path}")

# 2. Write to Workspace index.html
workspace_path = r"c:\Users\Swanto\OneDrive\Desktop\SNAPRINT_GW2024\index.html"
with open(workspace_path, "w", encoding="utf-8") as f:
    f.write(html_template)
print(f"Generated workspace index.html: {workspace_path}")
