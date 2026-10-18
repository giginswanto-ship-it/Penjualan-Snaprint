import json
import os

with open("sales_data_master.json", "r", encoding="utf-8") as f:
    master_data = json.load(f)

json_master_str = json.dumps(master_data, ensure_ascii=False)

unified_html = f"""<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SNAPRINT DIGITAL — Sistem Dashboard Penjualan Terintegrasi (2024 - 2025)</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <!-- SheetJS untuk Export Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
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

    <!-- MASTER HEADER & MULTI-YEAR SELECTOR -->
    <header class="bg-[var(--card,#ffffff)] text-[var(--foreground,#0f172a)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-4">
      <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
        
        <!-- Brand & Title -->
        <div class="flex items-center space-x-3 sm:space-x-4">
          <div class="w-12 h-12 rounded-xl bg-gradient-to-tr from-sky-600 via-indigo-600 to-purple-600 flex items-center justify-center text-white font-extrabold text-2xl shadow-md shadow-sky-600/20 shrink-0">
            S
          </div>
          <div>
            <div class="flex items-center space-x-2 flex-wrap">
              <h1 class="text-base sm:text-lg font-extrabold tracking-tight text-[var(--foreground,#0f172a)]">SNAPRINT DIGITAL PRINTING</h1>
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-indigo-500/10 text-indigo-700 border border-indigo-500/20">SISTEM TERINTEGRASI MULTI-TAHUN</span>
            </div>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Platform Analisis Kinerja Penjualan Dinamis — Konsolidasi Laporan ACCURATE Accounting
            </p>
          </div>
        </div>

        <!-- Quick Export & Reset Actions -->
        <div class="flex items-center space-x-2 sm:space-x-2.5 self-end md:self-center flex-wrap">
          <button onclick="exportUnifiedExcel()" class="px-3.5 py-2 text-xs font-semibold rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white transition-all shadow-sm flex items-center space-x-1.5">
            <span>📊</span>
            <span>Export Excel (.xlsx)</span>
          </button>
          <a id="btn-pdf-2024" href="Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf" target="_blank" class="px-3.5 py-2 text-xs font-semibold rounded-xl bg-slate-900 hover:bg-slate-800 text-white transition-all shadow-sm flex items-center space-x-1.5">
            <span>📄</span>
            <span>PDF 2024</span>
          </a>
          <a id="btn-pdf-2025" href="Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf" target="_blank" class="px-3.5 py-2 text-xs font-semibold rounded-xl bg-indigo-700 hover:bg-indigo-600 text-white transition-all shadow-sm flex items-center space-x-1.5">
            <span>📄</span>
            <span>PDF 2025</span>
          </a>
        </div>

      </div>

      <!-- YEAR SELECTOR TABS BAR -->
      <div class="pt-3 border-t border-[var(--border,#f1f5f9)] flex flex-wrap items-center justify-between gap-3">
        <div class="flex items-center space-x-2 bg-slate-100/80 p-1 rounded-xl border border-slate-200">
          <button onclick="selectYear('2024')" id="year-btn-2024" class="px-4 py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            📅 Tahun Buku 2024 (Rp 930,8 Jt)
          </button>
          <button onclick="selectYear('2025')" id="year-btn-2025" class="px-4 py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            🚀 Tahun Buku 2025 (Rp 1.378,9 Jt)
          </button>
          <button onclick="selectYear('YOY')" id="year-btn-YOY" class="px-4 py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            📈 Komparasi YoY (2024 vs 2025)
          </button>
        </div>

        <div class="flex items-center space-x-2 text-xs">
          <span class="text-[var(--muted-foreground,#64748b)]">Mode Tampilan:</span>
          <span id="current-mode-badge" class="font-bold text-sky-700 bg-sky-50 px-2.5 py-1 rounded-lg border border-sky-200">
            Tahun Buku 2024
          </span>
        </div>
      </div>

    </header>

    <!-- ======================================================== -->
    <!-- VIEW CONTAINER 1 & 2: SINGLE YEAR DASHBOARD (2024 / 2025) -->
    <!-- ======================================================== -->
    <div id="view-single-year" class="space-y-4 sm:space-y-6">

      <!-- DYNAMIC FILTER CONTROL PANEL -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-2 border-b border-[var(--border,#f1f5f9)] gap-2">
          <div class="flex items-center space-x-2">
            <span class="text-sm font-bold text-[var(--foreground,#0f172a)] flex items-center space-x-1.5">
              <span>⚙️</span>
              <span>Filter Penjualan Dinamis</span>
            </span>
            <span class="text-xs text-[var(--muted-foreground,#64748b)]">(Pilih filter atau klik batang grafik untuk update instan)</span>
          </div>
          <button onclick="resetFilters()" class="text-xs text-sky-600 hover:text-sky-700 font-semibold flex items-center space-x-1">
            <span>🔄</span>
            <span>Reset Filter</span>
          </button>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          
          <!-- Filter 1: Periode -->
          <div>
            <label class="block text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase mb-1">Periode Waktu</label>
            <select id="filter-period" onchange="setPeriod(this.value)" class="w-full text-xs p-2.5 rounded-xl border border-[var(--border,#cbd5e1)] bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
              <option value="ALL">📅 Semua (1 Tahun Penuh - 12 Bulan)</option>
              <option value="H1">🌱 Semester 1 (Jan - Jun)</option>
              <option value="H2">🚀 Semester 2 (Jul - Des)</option>
              <option value="Q1">1️⃣ Kuartal 1 (Januari - Maret)</option>
              <option value="Q2">2️⃣ Kuartal 2 (April - Juni)</option>
              <option value="Q3">3️⃣ Kuartal 3 (Juli - September)</option>
              <option value="Q4">4️⃣ Kuartal 4 (Oktober - Desember)</option>
              <option disabled>──────── Bulan Tertentu ────────</option>
              <option value="Jan">Januari</option>
              <option value="Feb">Februari</option>
              <option value="Mar">Maret</option>
              <option value="Apr">April</option>
              <option value="Mei">Mei</option>
              <option value="Jun">Juni</option>
              <option value="Jul">Juli</option>
              <option value="Agu">Agustus (Puncak Rekor)</option>
              <option value="Sep">September</option>
              <option value="Okt">Oktober (Puncak Kedua)</option>
              <option value="Nov">November</option>
              <option value="Des">Desember</option>
            </select>
          </div>

          <!-- Filter 2: Kategori -->
          <div>
            <label class="block text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase mb-1">Kategori Produk</label>
            <select id="filter-category" onchange="setCategory(this.value)" class="w-full text-xs p-2.5 rounded-xl border border-[var(--border,#cbd5e1)] bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
              <option value="ALL">🏷️ Semua Kategori</option>
              <option value="Large Format">Large Format &amp; Outdoor Banner</option>
              <option value="Digital Print">Digital Print A3+ &amp; Lembaran</option>
              <option value="Doc">Dokumen &amp; Fotocopy</option>
              <option value="Desain">Jasa Desain &amp; Cetak Lain</option>
              <option value="Finishing">Finishing, Jilid &amp; Laminasi</option>
              <option value="Merchandise">Merchandise &amp; Produk Custom</option>
            </select>
          </div>

          <!-- Filter 3: Pareto (2024 item filter) -->
          <div id="filter-pareto-box">
            <label class="block text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase mb-1">Klasifikasi Pareto</label>
            <select id="filter-pareto" onchange="setPareto(this.value)" class="w-full text-xs p-2.5 rounded-xl border border-[var(--border,#cbd5e1)] bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
              <option value="ALL">⭐ Semua Produk</option>
              <option value="CORE80">🏆 Core 80% Omzet Inti</option>
              <option value="LONGTAIL">📦 Produk Pelengkap / Long-Tail</option>
            </select>
          </div>

          <!-- Filter 4: Search -->
          <div>
            <label class="block text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase mb-1">Pencarian Barang / Layanan</label>
            <div class="relative">
              <input type="text" id="filter-search" oninput="setSearch(this.value)" placeholder="Ketik e.g. Flexy, Print A3, Cutting..." class="w-full text-xs p-2.5 pl-8 rounded-xl border border-[var(--border,#cbd5e1)] bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
              <span class="absolute left-2.5 top-2.5 text-slate-400">🔍</span>
            </div>
          </div>

        </div>
      </div>

      <!-- DYNAMIC KPI METRICS CARDS -->
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        
        <!-- Card 1: Omzet -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm hover:shadow transition-shadow">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase tracking-wider">Omzet Penjualan</span>
            <span class="w-7 h-7 rounded-lg bg-sky-50 text-sky-600 flex items-center justify-center font-bold text-sm">💰</span>
          </div>
          <div class="mt-2.5">
            <h2 class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-sales">-</h2>
            <p class="text-xs text-emerald-600 font-semibold mt-1 flex items-center space-x-1" id="kpi-sales-sub">
              <span>●</span>
              <span>100.00% dari Total Tahunan</span>
            </p>
          </div>
        </div>

        <!-- Card 2: Volume -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm hover:shadow transition-shadow">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase tracking-wider">Total Volume Order</span>
            <span class="w-7 h-7 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-sm">📦</span>
          </div>
          <div class="mt-2.5">
            <h2 class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-qty">-</h2>
            <p class="text-xs text-indigo-600 font-semibold mt-1" id="kpi-qty-sub">
              Unit / Meter Lari / Lembar Cetak
            </p>
          </div>
        </div>

        <!-- Card 3: Rata-Rata Bulanan -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm hover:shadow transition-shadow">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase tracking-wider">Rata-Rata Bulanan</span>
            <span class="w-7 h-7 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center font-bold text-sm">📊</span>
          </div>
          <div class="mt-2.5">
            <h2 class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-avg">-</h2>
            <p class="text-xs text-amber-600 font-semibold mt-1" id="kpi-avg-sub">
              Benchmark Kinerja Bulanan
            </p>
          </div>
        </div>

        <!-- Card 4: Bulan Puncak -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm hover:shadow transition-shadow">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-[var(--muted-foreground,#64748b)] uppercase tracking-wider">Bulan Rekor Puncak</span>
            <span class="w-7 h-7 rounded-lg bg-rose-50 text-rose-600 flex items-center justify-center font-bold text-sm">👑</span>
          </div>
          <div class="mt-2.5">
            <h2 class="text-xl sm:text-2xl font-extrabold text-rose-700 font-mono tracking-tight" id="kpi-peak">-</h2>
            <p class="text-xs text-rose-600 font-semibold mt-1" id="kpi-peak-sub">
              Agustus (HUT RI &amp; Kampanye)
            </p>
          </div>
        </div>

      </div>

      <!-- CHARTS SECTION -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-4 sm:gap-6">
        
        <!-- Chart 1: Interactive Monthly Trend (2 Cols) -->
        <div class="lg:col-span-2 bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm flex flex-col justify-between">
          <div class="flex items-start sm:items-center justify-between mb-2">
            <div>
              <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
                <span>📊</span>
                <span id="chart-title">Tren Omzet Penjualan Bulanan (Januari - Desember)</span>
              </h3>
              <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
                Setiap batang menunjukkan angka penjualan bersih (Rupiah) &amp; porsi persentase. Klik batang untuk filter!
              </p>
            </div>
            <span class="text-xs px-2.5 py-1 rounded-full bg-sky-50 text-sky-700 font-bold border border-sky-200 shrink-0" id="chart-year-badge">
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
                <span class="text-rose-600 font-semibold" id="legend-avg-text">Garis Rata-rata</span>
              </span>
            </div>
            <button onclick="setPeriod('ALL')" class="text-xs text-sky-600 hover:text-sky-700 font-bold">
              Reset ke 1 Tahun
            </button>
          </div>
        </div>

        <!-- Chart 2: Category Breakdown (1 Col) -->
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
              Klik baris kategori di bawah untuk memfilter data:
            </p>

            <!-- Category Progress Bars -->
            <div class="space-y-2.5" id="category-progress-container"></div>
          </div>

          <div class="mt-4 pt-3 border-t border-[var(--border,#f1f5f9)] text-xs text-[var(--muted-foreground,#64748b)] flex items-center justify-between">
            <span>Dua Kategori Inti:</span>
            <span class="font-bold text-sky-700">Menyumbang &gt;70% Omzet</span>
          </div>
        </div>

      </div>

      <!-- TABLE SECTION (ADAPTS ACCORDING TO 2024 VS 2025) -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-3 border-b border-[var(--border,#f1f5f9)] gap-2">
          <div>
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>📋</span>
              <span id="table-title">Daftar Rincian Penjualan</span>
            </h3>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5" id="table-subtitle">
              Menampilkan data sesuai filter yang aktif.
            </p>
          </div>
          <div class="text-xs text-[var(--muted-foreground,#64748b)]" id="table-help-text">
            💡 Klik baris produk untuk riwayat bulanan
          </div>
        </div>

        <!-- Dynamic Container for Table (Products Table for 2024, Monthly Matrix Table for 2025) -->
        <div id="dynamic-table-container"></div>

      </div>

    </div>

    <!-- ======================================================== -->
    <!-- VIEW CONTAINER 3: YEAR-OVER-YEAR (YoY) COMPARISON VIEW   -->
    <!-- ======================================================== -->
    <div id="view-yoy" class="hidden space-y-4 sm:space-y-6">

      <!-- YoY Headline Growth Banner -->
      <div class="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-2xl p-5 sm:p-6 shadow-sm border border-slate-800 space-y-4">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-white/10">
          <div>
            <span class="px-2.5 py-1 rounded-full text-xs font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
              LOMPATAN REKOR BISNIS (YoY ANALYSIS)
            </span>
            <h2 class="text-xl sm:text-2xl font-black mt-2">Performa Penjualan 2024 vs 2025</h2>
            <p class="text-xs text-slate-300 mt-1 max-w-2xl leading-relaxed">
              Snaprint Digital Printing berhasil mencatatkan ekspansi pendapatan masif sebesar <b>+48,14%</b> (tumbuh <b>+Rp 448,10 Juta</b>), melonjak dari Rp 930,85 Jt pada 2024 menjadi <b>Rp 1.378,95 Jt (Rp 1,378 Miliar)</b> pada 2025.
            </p>
          </div>
          <div class="text-right shrink-0 bg-white/5 p-4 rounded-xl border border-white/10">
            <span class="text-xs text-slate-400 block font-medium">Pertumbuhan Omzet Tahunan:</span>
            <span class="text-3xl font-black text-emerald-400 font-mono">+48,14%</span>
            <span class="text-xs text-slate-300 block font-mono mt-0.5">+Rp 448.104.669</span>
          </div>
        </div>

        <!-- 4 YoY Stat Cards -->
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
          <div class="p-3 bg-white/5 rounded-xl border border-white/10">
            <span class="text-slate-400 block">Total Omzet:</span>
            <p class="text-sm font-bold font-mono text-white mt-1">Rp 930,8 Jt ➔ <b class="text-emerald-400">Rp 1.378,9 Jt</b></p>
            <span class="text-emerald-400 font-semibold font-mono">+48,14% YoY</span>
          </div>
          <div class="p-3 bg-white/5 rounded-xl border border-white/10">
            <span class="text-slate-400 block">Volume Pesanan:</span>
            <p class="text-sm font-bold font-mono text-white mt-1">275.498 ➔ <b class="text-sky-300">387.284 unit</b></p>
            <span class="text-sky-300 font-semibold font-mono">+40,58% (+111k unit)</span>
          </div>
          <div class="p-3 bg-white/5 rounded-xl border border-white/10">
            <span class="text-slate-400 block">Rata-Rata Bulanan:</span>
            <p class="text-sm font-bold font-mono text-white mt-1">Rp 77,6 Jt ➔ <b class="text-amber-300">Rp 114,9 Jt/bln</b></p>
            <span class="text-amber-300 font-semibold font-mono">+Rp 37,3 Jt / bulan</span>
          </div>
          <div class="p-3 bg-white/5 rounded-xl border border-white/10">
            <span class="text-slate-400 block">Lompatan Kuartal 1:</span>
            <p class="text-sm font-bold font-mono text-white mt-1">Rp 84,7 Jt ➔ <b class="text-purple-300">Rp 296,5 Jt</b></p>
            <span class="text-purple-300 font-semibold font-mono">+250,18% (Q1 Boom)</span>
          </div>
        </div>
      </div>

      <!-- Side-by-Side Monthly Comparison Chart (SVG) -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
        <div class="flex items-center justify-between">
          <div>
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>📊</span>
              <span>Perbandingan Omzet Bulanan: 2024 (Abu-Abu/Biru) vs 2025 (Emerald)</span>
            </h3>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Grafik perbandingan langsung setiap bulan beserta persentase kenaikan tahun-ke-tahun (YoY)
            </p>
          </div>
          <div class="flex items-center space-x-3 text-xs">
            <span class="flex items-center space-x-1.5">
              <span class="w-3.5 h-3.5 rounded bg-sky-600 inline-block"></span>
              <span class="font-semibold text-slate-700">Tahun 2024</span>
            </span>
            <span class="flex items-center space-x-1.5">
              <span class="w-3.5 h-3.5 rounded bg-emerald-500 inline-block"></span>
              <span class="font-semibold text-emerald-800">Tahun 2025</span>
            </span>
          </div>
        </div>

        <div class="w-full overflow-x-auto py-2">
          <svg id="svg-yoy-chart" viewBox="0 0 860 300" class="w-full h-64 sm:h-72 select-none"></svg>
        </div>
      </div>

      <!-- YoY Monthly Table Detailed Breakdown -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
        <div class="flex items-center justify-between pb-2 border-b border-[var(--border,#f1f5f9)]">
          <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
            <span>📑</span>
            <span>Tabel Pertumbuhan Penjualan Bulanan (YoY Matrix 2024 vs 2025)</span>
          </h3>
          <span class="text-xs font-semibold text-slate-500">12 Bulan Konsolidasi</span>
        </div>

        <div class="overflow-x-auto border border-[var(--border,#e2e8f0)] rounded-xl">
          <table class="w-full text-left text-xs font-medium">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
              <tr>
                <th class="py-2.5 px-3">Bulan</th>
                <th class="py-2.5 px-3 text-right">Omzet 2024 (Rp)</th>
                <th class="py-2.5 px-3 text-right">Omzet 2025 (Rp)</th>
                <th class="py-2.5 px-3 text-right">Selisih Pertumbuhan (Rp)</th>
                <th class="py-2.5 px-3 text-right">Pertumbuhan (%)</th>
                <th class="py-2.5 px-3 text-right">Volume 2024</th>
                <th class="py-2.5 px-3 text-right">Volume 2025</th>
              </tr>
            </thead>
            <tbody id="yoy-table-tbody" class="divide-y divide-slate-100 font-mono text-slate-800">
              <!-- Populated via JS -->
            </tbody>
            <tfoot class="bg-slate-100 font-bold text-slate-900 border-t-2 border-slate-300 font-mono">
              <tr>
                <td class="py-2.5 px-3 font-sans">TOTAL TAHUNAN</td>
                <td class="py-2.5 px-3 text-right">Rp 930.847.260</td>
                <td class="py-2.5 px-3 text-right text-emerald-700">Rp 1.378.951.929</td>
                <td class="py-2.5 px-3 text-right text-emerald-700">+Rp 448.104.669</td>
                <td class="py-2.5 px-3 text-right text-emerald-700">+48,14%</td>
                <td class="py-2.5 px-3 text-right">275.498,79</td>
                <td class="py-2.5 px-3 text-right text-emerald-700">387.284,00</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

    </div>

  </div>

  <!-- PRODUCT DETAIL POPUP MODAL -->
  <div id="product-detail-modal" class="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-sm hidden flex items-center justify-center p-3 sm:p-4">
    <div class="bg-[var(--card,#ffffff)] text-[var(--foreground,#0f172a)] w-full max-w-2xl rounded-2xl shadow-2xl border border-[var(--border,#e2e8f0)] overflow-hidden flex flex-col max-h-[90vh]">
      <div class="p-4 sm:p-5 bg-slate-900 text-white flex items-start justify-between">
        <div>
          <div class="flex items-center space-x-2">
            <span class="px-2.5 py-0.5 rounded text-[11px] font-bold bg-sky-500/20 text-sky-300 border border-sky-400/30" id="modal-rank">RANK</span>
            <span class="text-xs text-slate-300 font-medium" id="modal-cat">Kategori</span>
          </div>
          <h2 class="text-base sm:text-lg font-bold text-white mt-1.5" id="modal-name">Nama Produk</h2>
        </div>
        <button onclick="closeModal()" class="text-slate-400 hover:text-white p-1 rounded-lg text-lg">✕</button>
      </div>

      <div class="p-4 sm:p-5 overflow-y-auto space-y-4">
        <div class="grid grid-cols-3 gap-3 text-center">
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
            <span class="text-[11px] text-slate-500 font-semibold uppercase">Total Omzet</span>
            <p class="text-sm sm:text-base font-bold text-slate-900 font-mono mt-0.5" id="modal-sales">-</p>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
            <span class="text-[11px] text-slate-500 font-semibold uppercase">Total Volume</span>
            <p class="text-sm sm:text-base font-bold text-slate-900 font-mono mt-0.5" id="modal-qty">-</p>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
            <span class="text-[11px] text-slate-500 font-semibold uppercase">Porsi Kontribusi</span>
            <p class="text-sm sm:text-base font-bold text-sky-600 font-mono mt-0.5" id="modal-pct">-</p>
          </div>
        </div>

        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
          <span class="text-xs font-bold text-slate-700 block mb-1">Riwayat Penjualan 12 Bulan (Januari - Desember)</span>
          <svg id="modal-spark-chart" viewBox="0 0 540 100" class="w-full h-28"></svg>
        </div>

        <div class="border border-slate-200 rounded-xl overflow-hidden">
          <table class="w-full text-xs text-left">
            <thead class="bg-slate-100 font-semibold text-slate-700 uppercase text-[10px]">
              <tr>
                <th class="py-1.5 px-3">Bulan</th>
                <th class="py-1.5 px-3 text-right">Volume</th>
                <th class="py-1.5 px-3 text-right">Penjualan (Rp)</th>
              </tr>
            </thead>
            <tbody id="modal-month-tbody" class="divide-y divide-slate-100 font-mono text-slate-800"></tbody>
          </table>
        </div>
      </div>

      <div class="p-3.5 bg-slate-50 border-t border-slate-200 flex justify-end">
        <button onclick="closeModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold rounded-lg transition-colors">
          Tutup
        </button>
      </div>
    </div>
  </div>

  <!-- JAVASCRIPT APP LOGIC -->
  <script>
    const MASTER_DATA = {json_master_str};

    // State
    const appState = {{
      activeYear: '2024', // '2024', '2025', or 'YOY'
      period: 'ALL',      // 'ALL', 'H1', 'H2', 'Q1'..'Q4', or 'Jan'..'Des'
      category: 'ALL',
      pareto: 'ALL',
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

    // Year Switching
    function selectYear(yr) {{
      appState.activeYear = yr;
      
      // Update Buttons
      ['2024', '2025', 'YOY'].forEach(y => {{
        const btn = document.getElementById('year-btn-' + y);
        if (y === yr) {{
          btn.className = 'px-4 py-2 rounded-lg text-xs font-bold transition-all bg-sky-600 text-white shadow-sm';
        }} else {{
          btn.className = 'px-4 py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900';
        }}
      }});

      const modeBadge = document.getElementById('current-mode-badge');
      if (yr === '2024') {{
        modeBadge.innerText = 'Tahun Buku 2024 (Rp 930,85 Jt)';
        modeBadge.className = 'font-bold text-sky-700 bg-sky-50 px-2.5 py-1 rounded-lg border border-sky-200';
      }} else if (yr === '2025') {{
        modeBadge.innerText = 'Tahun Buku 2025 (Rp 1.378,95 Jt)';
        modeBadge.className = 'font-bold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200';
      }} else {{
        modeBadge.innerText = 'Komparasi YoY (2024 vs 2025: +48,14%)';
        modeBadge.className = 'font-bold text-purple-700 bg-purple-50 px-2.5 py-1 rounded-lg border border-purple-200';
      }}

      // Switch views
      if (yr === 'YOY') {{
        document.getElementById('view-single-year').classList.add('hidden');
        document.getElementById('view-yoy').classList.remove('hidden');
        renderYoYView();
      }} else {{
        document.getElementById('view-single-year').classList.remove('hidden');
        document.getElementById('view-yoy').classList.add('hidden');
        
        // Hide/show pareto for 2025 (2025 has category matrix)
        document.getElementById('filter-pareto-box').style.display = (yr === '2024') ? 'block' : 'none';
        
        renderSingleYearDashboard();
      }}
    }}

    function getActiveMonths() {{
      const p = appState.period;
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

    function setPeriod(p) {{
      appState.period = p;
      document.getElementById('filter-period').value = p;
      renderSingleYearDashboard();
    }}

    function setCategory(c) {{
      appState.category = c;
      document.getElementById('filter-category').value = c;
      renderSingleYearDashboard();
    }}

    function setPareto(val) {{
      appState.pareto = val;
      renderSingleYearDashboard();
    }}

    function setSearch(q) {{
      appState.search = q.toLowerCase().trim();
      renderSingleYearDashboard();
    }}

    function resetFilters() {{
      appState.period = 'ALL';
      appState.category = 'ALL';
      appState.pareto = 'ALL';
      appState.search = '';
      document.getElementById('filter-period').value = 'ALL';
      document.getElementById('filter-category').value = 'ALL';
      document.getElementById('filter-pareto').value = 'ALL';
      document.getElementById('filter-search').value = '';
      renderSingleYearDashboard();
    }}

    // ========================================================
    // RENDER SINGLE YEAR DASHBOARD (2024 or 2025)
    // ========================================================
    function renderSingleYearDashboard() {{
      const yr = appState.activeYear;
      const yrData = (yr === '2025') ? MASTER_DATA.data_2025 : MASTER_DATA.data_2024;
      const activeMonths = getActiveMonths();

      // Update KPI
      let periodSales = 0;
      let periodQty = 0;
      yrData.monthly.forEach(m => {{
        if (activeMonths.includes(m.code)) {{
          periodSales += m.sales;
          periodQty += m.qty;
        }}
      }});

      const totalAnnualSales = yrData.kpi.total_sales;
      const pctAnnual = totalAnnualSales > 0 ? (periodSales / totalAnnualSales) * 100 : 0;

      document.getElementById('kpi-sales').innerText = fmtIDR(periodSales);
      document.getElementById('kpi-sales-sub').innerHTML = `<span>●</span><span>${{pctAnnual.toFixed(2)}}% dari Total Tahun ${{yr}} (${{fmtIDR(totalAnnualSales)}})</span>`;
      
      document.getElementById('kpi-qty').innerText = fmtNum(periodQty);
      document.getElementById('kpi-avg').innerText = fmtIDR(yrData.kpi.avg_monthly_sales);
      document.getElementById('kpi-peak').innerText = `${{yrData.kpi.peak_month}} (${{fmtIDR(yrData.kpi.peak_month_sales)}})`;

      document.getElementById('chart-title').innerText = `Tren Omzet Penjualan Bulanan (Tahun Buku ${{yr}})`;
      document.getElementById('chart-year-badge').innerText = `Tahun Buku ${{yr}}`;
      document.getElementById('legend-avg-text').innerText = `Garis Rata-rata (${{fmtIDR(yrData.kpi.avg_monthly_sales)}})`;

      // Render Monthly SVG Chart
      renderMonthlySvgChart(yrData);

      // Render Category Progress Bars
      renderCategoryProgressBars(yrData);

      // Render Table
      if (yr === '2024') {{
        document.getElementById('table-title').innerText = 'Peringkat & Rincian 129 Produk Tahun 2024';
        document.getElementById('table-help-text').style.display = 'block';
        render2024ProductsTable(yrData);
      }} else {{
        document.getElementById('table-title').innerText = 'Matriks Penjualan Bulanan per Kategori Produk Tahun 2025';
        document.getElementById('table-help-text').style.display = 'none';
        render2025MatrixTable(yrData);
      }}
    }}

    // Monthly SVG Chart
    function renderMonthlySvgChart(yrData) {{
      const svg = document.getElementById('svg-monthly-chart');
      svg.innerHTML = '';

      const activeMonths = getActiveMonths();
      const monthlyData = yrData.monthly;
      const annualAvgSales = yrData.kpi.avg_monthly_sales;

      const chartLeft = 70;
      const chartRight = 835;
      const chartWidth = chartRight - chartLeft;
      const chartTop = 50;
      const chartBottom = 240;
      const chartHeight = chartBottom - chartTop;

      const maxScale = (appState.activeYear === '2025') ? 160000000 : 145000000;

      const numBars = 12;
      const barWidth = 46;
      const gap = (chartWidth - (numBars * barWidth)) / (numBars + 1);

      // Gradients
      const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
      defs.innerHTML = `
        <linearGradient id="grad-active" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="${{appState.activeYear === '2025' ? '#059669' : '#0284c7'}}" />
          <stop offset="100%" stop-color="${{appState.activeYear === '2025' ? '#047857' : '#0369a1'}}" />
        </linearGradient>
        <filter id="shadow-pill" x="-10%" y="-10%" width="120%" height="130%">
          <feDropShadow dx="0" dy="1.5" stdDeviation="1.5" flood-color="#0f172a" flood-opacity="0.08"/>
        </filter>
      `;
      svg.appendChild(defs);

      // Gridlines
      const step = (appState.activeYear === '2025') ? 35000000 : 25000000;
      for (let s = step; s < maxScale; s += step) {{
        const yPos = chartBottom - (s / maxScale) * chartHeight;
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', chartLeft);
        line.setAttribute('y1', yPos);
        line.setAttribute('x2', chartRight);
        line.setAttribute('y2', yPos);
        line.setAttribute('stroke', '#e2e8f0');
        line.setAttribute('stroke-dasharray', '3,3');
        svg.appendChild(line);

        const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt.setAttribute('x', chartLeft - 8);
        txt.setAttribute('y', yPos + 3.5);
        txt.setAttribute('text-anchor', 'end');
        txt.setAttribute('font-size', '9.5');
        txt.setAttribute('font-weight', '600');
        txt.setAttribute('fill', '#94a3b8');
        txt.textContent = 'Rp ' + (s / 1e6).toFixed(0) + ' Jt';
        svg.appendChild(txt);
      }}

      // Base line
      const baseLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      baseLine.setAttribute('x1', chartLeft);
      baseLine.setAttribute('y1', chartBottom);
      baseLine.setAttribute('x2', chartRight);
      baseLine.setAttribute('y2', chartBottom);
      baseLine.setAttribute('stroke', '#94a3b8');
      baseLine.setAttribute('stroke-width', '1.5');
      svg.appendChild(baseLine);

      // Average Line
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

      const avgTagG = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      avgTagG.innerHTML = `
        <rect x="${{chartRight - 128}}" y="${{avgY - 10}}" width="128" height="20" rx="5" fill="#fff1f2" stroke="#fda4af" stroke-width="1" />
        <text x="${{chartRight - 64}}" y="${{avgY + 4}}" font-size="9.5" font-weight="700" fill="#e11d48" text-anchor="middle">Rata-rata: ${{fmtIDR(annualAvgSales)}}</text>
      `;
      svg.appendChild(avgTagG);

      // Bars
      monthlyData.forEach((m, i) => {{
        const x = chartLeft + gap + i * (barWidth + gap);
        const barH = (m.sales / maxScale) * chartHeight;
        const barTop = chartBottom - barH;
        const centerX = x + barWidth / 2;
        const isSelected = activeMonths.includes(m.code);
        const isPeak = (m.code === 'Agu');

        const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        g.className.baseVal = 'chart-bar-group cursor-pointer';

        g.onmouseenter = () => {{
          const diffPct = ((m.sales - annualAvgSales) / annualAvgSales) * 100;
          const statusStr = diffPct >= 0 
            ? `<span class="text-emerald-700 font-bold">+${{diffPct.toFixed(1)}}% di atas rata-rata</span>` 
            : `<span class="text-rose-700 font-bold">${{diffPct.toFixed(1)}}% di bawah rata-rata</span>`;
          
          document.getElementById('chart-hover-info').innerHTML = `
            <div class="flex items-center space-x-2">
              <span class="font-extrabold text-slate-900 text-xs">📅 ${{m.name}} ${{appState.activeYear}}:</span>
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

        // Rect
        const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect.setAttribute('x', x);
        rect.setAttribute('y', barTop);
        rect.setAttribute('width', barWidth);
        rect.setAttribute('height', barH);
        rect.setAttribute('rx', '6');
        rect.setAttribute('fill', isSelected ? 'url(#grad-active)' : '#cbd5e1');
        if (isSelected && isPeak) {{
          rect.setAttribute('stroke', '#f59e0b');
          rect.setAttribute('stroke-width', '1.5');
        }}
        g.appendChild(rect);

        // Badge pill
        const bW = 54;
        const bH = 27;
        const bX = centerX - bW / 2;
        const bY = barTop - bH - 5;

        const badgeG = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        badgeG.setAttribute('filter', 'url(#shadow-pill)');

        const bBg = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        bBg.setAttribute('x', bX);
        bBg.setAttribute('y', bY);
        bBg.setAttribute('width', bW);
        bBg.setAttribute('height', bH);
        bBg.setAttribute('rx', '6');
        bBg.setAttribute('fill', isPeak ? '#fef3c7' : '#ffffff');
        bBg.setAttribute('stroke', isPeak ? '#f59e0b' : '#e2e8f0');
        bBg.setAttribute('stroke-width', isPeak ? '1.5' : '1');
        badgeG.appendChild(bBg);

        const txtVal = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txtVal.setAttribute('x', centerX);
        txtVal.setAttribute('y', bY + 12);
        txtVal.setAttribute('text-anchor', 'middle');
        txtVal.setAttribute('font-size', '10');
        txtVal.setAttribute('font-weight', '800');
        txtVal.setAttribute('fill', isPeak ? '#92400e' : '#0f172a');
        txtVal.textContent = (m.sales / 1e6).toFixed(1) + ' Jt';
        badgeG.appendChild(txtVal);

        const txtSub = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txtSub.setAttribute('x', centerX);
        txtSub.setAttribute('y', bY + 22);
        txtSub.setAttribute('text-anchor', 'middle');
        txtSub.setAttribute('font-size', '8.5');
        txtSub.setAttribute('font-weight', '700');
        txtSub.setAttribute('fill', isPeak ? '#b45309' : '#0284c7');
        txtSub.textContent = isPeak ? '👑 Puncak' : `${{m.pct.toFixed(1)}}%`;
        badgeG.appendChild(txtSub);

        g.appendChild(badgeG);

        // Month Text
        const txtMonth = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txtMonth.setAttribute('x', centerX);
        txtMonth.setAttribute('y', chartBottom + 18);
        txtMonth.setAttribute('text-anchor', 'middle');
        txtMonth.setAttribute('font-size', '11.5');
        txtMonth.setAttribute('font-weight', isSelected ? '800' : '600');
        txtMonth.setAttribute('fill', isSelected ? '#0f172a' : '#64748b');
        txtMonth.textContent = m.code;
        g.appendChild(txtMonth);

        // Qty Text
        const txtQty = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txtQty.setAttribute('x', centerX);
        txtQty.setAttribute('y', chartBottom + 32);
        txtQty.setAttribute('text-anchor', 'middle');
        txtQty.setAttribute('font-size', '9');
        txtQty.setAttribute('font-weight', '500');
        txtQty.setAttribute('fill', '#94a3b8');
        txtQty.textContent = (m.qty / 1000).toFixed(1) + 'k u';
        g.appendChild(txtQty);

        svg.appendChild(g);
      }});
    }}

    // Category Progress Bars
    function renderCategoryProgressBars(yrData) {{
      const container = document.getElementById('category-progress-container');
      container.innerHTML = '';

      yrData.categories.forEach(cat => {{
        const div = document.createElement('div');
        div.className = 'p-2 rounded-xl transition-all cursor-pointer border bg-slate-50/70 border-slate-200/70 hover:bg-slate-100';
        div.innerHTML = `
          <div class="flex items-center justify-between text-xs">
            <span class="font-bold text-slate-800 truncate max-w-[200px]">${{cat.name}}</span>
            <span class="font-mono font-bold text-slate-900">${{fmtIDR(cat.total_sales)}} (${{cat.pct.toFixed(1)}}%)</span>
          </div>
          <div class="w-full bg-slate-200 h-2 rounded-full overflow-hidden mt-1.5">
            <div class="bg-sky-600 h-full rounded-full transition-all duration-300" style="width: ${{cat.pct}}%"></div>
          </div>
        `;
        container.appendChild(div);
      }});
    }}

    // 2024 Products Table
    function render2024ProductsTable(yrData) {{
      const container = document.getElementById('dynamic-table-container');
      const activeMonths = getActiveMonths();

      let filtered = yrData.products.filter(p => {{
        if (appState.category !== 'ALL' && !p.category.toLowerCase().includes(appState.category.toLowerCase())) return false;
        if (appState.pareto === 'CORE80' && !p.is_pareto_80) return false;
        if (appState.pareto === 'LONGTAIL' && p.is_pareto_80) return false;
        if (appState.search && !p.name.toLowerCase().includes(appState.search)) return false;
        return true;
      }});

      let totS = 0;
      let totQ = 0;
      filtered = filtered.map(p => {{
        let s = 0, q = 0;
        activeMonths.forEach(m => {{
          s += (p.monthly_sales[m] || 0);
          q += (p.monthly_qty[m] || 0);
        }});
        totS += s;
        totQ += q;
        return {{ ...p, period_sales: s, period_qty: q }};
      }}).filter(p => p.period_sales > 0 || activeMonths.length === 12);

      filtered.sort((a, b) => b.period_sales - a.period_sales);

      let rowsHtml = filtered.map((p, idx) => `
        <tr class="hover:bg-sky-50/60 cursor-pointer transition-colors" onclick='openModal(${{JSON.stringify(p)}})'>
          <td class="py-2 px-3 text-center font-mono text-slate-500">${{idx + 1}}</td>
          <td class="py-2 px-3 font-semibold text-slate-900 hover:text-sky-600">${{p.name}}</td>
          <td class="py-2 px-3 text-slate-500 text-[11px] truncate max-w-[170px]">${{p.category}}</td>
          <td class="py-2 px-3 text-center text-slate-500 text-[11px] font-mono">${{p.unit}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtNum(p.period_qty)}}</td>
          <td class="py-2 px-3 text-right font-mono font-bold text-slate-900">${{fmtIDR(p.period_sales)}}</td>
          <td class="py-2 px-3 text-right font-mono text-slate-600">${{totS > 0 ? ((p.period_sales / totS) * 100).toFixed(2) : 0}}%</td>
          <td class="py-2 px-3 text-center">
            ${{p.is_pareto_80 ? '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800">⭐ Core 80%</span>' : '<span class="px-2 py-0.5 rounded-full text-[10px] font-medium bg-slate-100 text-slate-600">Pelengkap</span>'}}
          </td>
        </tr>
      `).join('');

      container.innerHTML = `
        <div class="overflow-x-auto max-h-[500px] border border-[var(--border,#e2e8f0)] rounded-xl">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider sticky top-0 z-10 shadow-sm">
              <tr>
                <th class="py-2.5 px-3 text-center w-12">No</th>
                <th class="py-2.5 px-3">Nama Barang / Produk Layanan</th>
                <th class="py-2.5 px-3">Kategori</th>
                <th class="py-2.5 px-3 text-center w-16">Satuan</th>
                <th class="py-2.5 px-3 text-right">Volume</th>
                <th class="py-2.5 px-3 text-right">Total Omzet (Rp)</th>
                <th class="py-2.5 px-3 text-right">Porsi (%)</th>
                <th class="py-2.5 px-3 text-center">Klasifikasi</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-[var(--border,#f1f5f9)] font-medium text-slate-800">${{rowsHtml}}</tbody>
            <tfoot class="bg-slate-100 font-bold text-slate-900 sticky bottom-0 z-10 border-t-2 border-slate-300">
              <tr>
                <td class="py-2.5 px-3 text-center" colspan="4">SUBTOTAL SESUAI FILTER AKTIF</td>
                <td class="py-2.5 px-3 text-right font-mono">${{fmtNum(totQ)}}</td>
                <td class="py-2.5 px-3 text-right font-mono text-emerald-700">${{fmtIDR(totS)}}</td>
                <td class="py-2.5 px-3 text-right font-mono">100.00%</td>
                <td class="py-2.5 px-3 text-center">${{filtered.length}} Item</td>
              </tr>
            </tfoot>
          </table>
        </div>
      `;
    }}

    // 2025 Matrix Table
    function render2025MatrixTable(yrData) {{
      const container = document.getElementById('dynamic-table-container');
      const activeMonths = getActiveMonths();

      const rows = yrData.monthly.filter(m => activeMonths.includes(m.code)).map((m, idx) => `
        <tr class="hover:bg-slate-50 transition-colors">
          <td class="py-2.5 px-3 font-semibold text-slate-900">${{m.name}} 2025</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.large_format)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.print_a3)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.doc_fc)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.finishing)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.merchandise)}}</td>
          <td class="py-2.5 px-3 text-right font-mono font-bold text-emerald-700 bg-emerald-50/50">${{fmtIDR(m.sales)}}</td>
        </tr>
      `).join('');

      container.innerHTML = `
        <div class="overflow-x-auto max-h-[500px] border border-[var(--border,#e2e8f0)] rounded-xl">
          <table class="w-full text-left text-xs font-medium">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider sticky top-0 z-10 shadow-sm">
              <tr>
                <th class="py-2.5 px-3">Bulan</th>
                <th class="py-2.5 px-3 text-right">Large Format</th>
                <th class="py-2.5 px-3 text-right">Print A3+ &amp; Sheet</th>
                <th class="py-2.5 px-3 text-right">Doc &amp; FC</th>
                <th class="py-2.5 px-3 text-right">Finishing</th>
                <th class="py-2.5 px-3 text-right">Merchandise</th>
                <th class="py-2.5 px-3 text-right">Total Omzet Bulanan</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100 font-mono text-slate-800">${{rows}}</tbody>
            <tfoot class="bg-slate-100 font-bold text-slate-900 sticky bottom-0 z-10 border-t-2 border-slate-300 font-mono">
              <tr>
                <td class="py-2.5 px-3 font-sans">TOTAL 2025</td>
                <td class="py-2.5 px-3 text-right">Rp 509,0 Jt</td>
                <td class="py-2.5 px-3 text-right">Rp 488,6 Jt</td>
                <td class="py-2.5 px-3 text-right">Rp 118,9 Jt</td>
                <td class="py-2.5 px-3 text-right">Rp 64,8 Jt</td>
                <td class="py-2.5 px-3 text-right">Rp 157,6 Jt</td>
                <td class="py-2.5 px-3 text-right text-emerald-700">Rp 1.378,95 Jt</td>
              </tr>
            </tfoot>
          </table>
        </div>
      `;
    }}

    // ========================================================
    // RENDER YoY VIEW (2024 vs 2025 COMPARISON)
    // ========================================================
    function renderYoYView() {{
      const yoyData = MASTER_DATA.yoy;
      
      // Render Table
      const tbody = document.getElementById('yoy-table-tbody');
      tbody.innerHTML = '';

      yoyData.monthly_comparison.forEach(m => {{
        const isPos = m.growth_rp >= 0;
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition-colors';
        tr.innerHTML = `
          <td class="py-2.5 px-3 font-sans font-semibold text-slate-900">${{m.name}}</td>
          <td class="py-2.5 px-3 text-right">${{fmtIDR(m.sales_2024)}}</td>
          <td class="py-2.5 px-3 text-right font-bold text-slate-900">${{fmtIDR(m.sales_2025)}}</td>
          <td class="py-2.5 px-3 text-right ${{isPos ? 'text-emerald-700' : 'text-rose-600'}} font-bold">${{isPos ? '+' : ''}}${{fmtIDR(m.growth_rp)}}</td>
          <td class="py-2.5 px-3 text-right ${{isPos ? 'text-emerald-700' : 'text-rose-600'}} font-bold">${{isPos ? '+' : ''}}${{m.growth_pct.toFixed(1)}}%</td>
          <td class="py-2.5 px-3 text-right text-slate-500">${{fmtNum(m.qty_2024)}}</td>
          <td class="py-2.5 px-3 text-right text-slate-900 font-bold">${{fmtNum(m.qty_2025)}}</td>
        `;
        tbody.appendChild(tr);
      }});

      // Render Side-by-Side SVG Chart
      renderYoYSvgChart();
    }}

    function renderYoYSvgChart() {{
      const svg = document.getElementById('svg-yoy-chart');
      svg.innerHTML = '';

      const yoy = MASTER_DATA.yoy.monthly_comparison;
      const maxVal = 150000000;

      const chartLeft = 70;
      const chartRight = 835;
      const chartWidth = chartRight - chartLeft;
      const chartTop = 50;
      const chartBottom = 240;
      const chartHeight = chartBottom - chartTop;

      const numMonths = 12;
      const groupWidth = chartWidth / numMonths;
      const barWidth = 20;

      // Base line
      const baseLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      baseLine.setAttribute('x1', chartLeft);
      baseLine.setAttribute('y1', chartBottom);
      baseLine.setAttribute('x2', chartRight);
      baseLine.setAttribute('y2', chartBottom);
      baseLine.setAttribute('stroke', '#94a3b8');
      baseLine.setAttribute('stroke-width', '1.5');
      svg.appendChild(baseLine);

      // Gridlines
      [150000000, 100000000, 50000000].forEach(val => {{
        const yPos = chartBottom - (val / maxVal) * chartHeight;
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', chartLeft);
        line.setAttribute('y1', yPos);
        line.setAttribute('x2', chartRight);
        line.setAttribute('y2', yPos);
        line.setAttribute('stroke', '#e2e8f0');
        line.setAttribute('stroke-dasharray', '3,3');
        svg.appendChild(line);

        const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt.setAttribute('x', chartLeft - 8);
        txt.setAttribute('y', yPos + 3.5);
        txt.setAttribute('text-anchor', 'end');
        txt.setAttribute('font-size', '9.5');
        txt.setAttribute('font-weight', '600');
        txt.setAttribute('fill', '#94a3b8');
        txt.textContent = 'Rp ' + (val / 1e6) + ' Jt';
        svg.appendChild(txt);
      }});

      // Draw Pairs of Bars
      yoy.forEach((m, i) => {{
        const groupX = chartLeft + i * groupWidth;
        const x24 = groupX + (groupWidth - (barWidth * 2 + 6)) / 2;
        const x25 = x24 + barWidth + 4;

        const h24 = (m.sales_2024 / maxVal) * chartHeight;
        const y24 = chartBottom - h24;

        const h25 = (m.sales_2025 / maxVal) * chartHeight;
        const y25 = chartBottom - h25;

        // Bar 2024 (Sky)
        const rect24 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect24.setAttribute('x', x24);
        rect24.setAttribute('y', y24);
        rect24.setAttribute('width', barWidth);
        rect24.setAttribute('height', h24);
        rect24.setAttribute('rx', '4');
        rect24.setAttribute('fill', '#0284c7');
        svg.appendChild(rect24);

        // Bar 2025 (Emerald)
        const rect25 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect25.setAttribute('x', x25);
        rect25.setAttribute('y', y25);
        rect25.setAttribute('width', barWidth);
        rect25.setAttribute('height', h25);
        rect25.setAttribute('rx', '4');
        rect25.setAttribute('fill', '#10b981');
        svg.appendChild(rect25);

        // Growth Badge on Top
        const badgeTop = Math.min(y24, y25) - 18;
        const txtGrowth = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txtGrowth.setAttribute('x', (x24 + x25 + barWidth) / 2);
        txtGrowth.setAttribute('y', badgeTop);
        txtGrowth.setAttribute('text-anchor', 'middle');
        txtGrowth.setAttribute('font-size', '9');
        txtGrowth.setAttribute('font-weight', '800');
        txtGrowth.setAttribute('fill', m.growth_pct >= 0 ? '#059669' : '#e11d48');
        txtGrowth.textContent = `+${{m.growth_pct.toFixed(0)}}%`;
        svg.appendChild(txtGrowth);

        // Month Label
        const txtMonth = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txtMonth.setAttribute('x', (x24 + x25 + barWidth) / 2);
        txtMonth.setAttribute('y', chartBottom + 18);
        txtMonth.setAttribute('text-anchor', 'middle');
        txtMonth.setAttribute('font-size', '11');
        txtMonth.setAttribute('font-weight', '700');
        txtMonth.setAttribute('fill', '#0f172a');
        txtMonth.textContent = m.code;
        svg.appendChild(txtMonth);
      }});
    }}

    // Modal Details for 2024 Products
    function openModal(p) {{
      document.getElementById('modal-rank').innerText = `RANK #${{p.rank}}`;
      document.getElementById('modal-cat').innerText = p.category;
      document.getElementById('modal-name').innerText = p.name;
      document.getElementById('modal-sales').innerText = fmtIDR(p.total_sales);
      document.getElementById('modal-qty').innerText = `${{fmtNum(p.total_qty)}} ${{p.unit}}`;
      document.getElementById('modal-pct').innerText = `${{p.pct.toFixed(2)}}% dari Omzet`;

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
            <td class="py-1 px-3 font-semibold text-slate-800">${{MONTH_NAMES[m]}}</td>
            <td class="py-1 px-3 text-right">${{fmtNum(q)}}</td>
            <td class="py-1 px-3 text-right font-bold text-slate-900">${{fmtIDR(s)}}</td>
          `;
          tbody.appendChild(tr);
        }}
      }});

      document.getElementById('product-detail-modal').classList.remove('hidden');
    }}

    function closeModal() {{
      document.getElementById('product-detail-modal').classList.add('hidden');
    }}

    // Export Excel Multi-sheet
    function exportUnifiedExcel() {{
      const wb = XLSX.utils.book_new();

      // Sheet 1: YoY Summary
      const yoyRows = [
        ['Bulan', 'Omzet 2024 (Rp)', 'Omzet 2025 (Rp)', 'Pertumbuhan Nominal (Rp)', 'Pertumbuhan (%)', 'Volume 2024', 'Volume 2025']
      ];
      MASTER_DATA.yoy.monthly_comparison.forEach(m => {{
        yoyRows.push([m.name, m.sales_2024, m.sales_2025, m.growth_rp, m.growth_pct / 100, m.qty_2024, m.qty_2025]);
      }});
      yoyRows.push(['TOTAL', MASTER_DATA.data_2024.kpi.total_sales, MASTER_DATA.data_2025.kpi.total_sales, MASTER_DATA.yoy.growth_sales, MASTER_DATA.yoy.growth_sales_pct / 100, MASTER_DATA.data_2024.kpi.total_qty, MASTER_DATA.data_2025.kpi.total_qty]);
      const wsYoY = XLSX.utils.aoa_to_sheet(yoyRows);
      XLSX.utils.book_append_sheet(wb, wsYoY, 'Komparasi YoY (24 vs 25)');

      // Sheet 2: 2025 Monthly Matrix
      const m25Rows = [
        ['Bulan', 'Large Format', 'Print A3+ & Sheet', 'Doc & FC', 'Finishing', 'Merchandise', 'Total Omzet Bulanan']
      ];
      MASTER_DATA.data_2025.monthly.forEach(m => {{
        m25Rows.push([m.name + ' 2025', m.matrix.large_format, m.matrix.print_a3, m.matrix.doc_fc, m.matrix.finishing, m.matrix.merchandise, m.sales]);
      }});
      const ws25 = XLSX.utils.aoa_to_sheet(m25Rows);
      XLSX.utils.book_append_sheet(wb, ws25, 'Matriks Penjualan 2025');

      // Sheet 3: 2024 Products
      const p24Rows = [
        ['Rank', 'Nama Barang', 'Kategori', 'Satuan', 'Total Volume', 'Total Omzet (Rp)', 'Porsi (%)']
      ];
      MASTER_DATA.data_2024.products.forEach(p => {{
        p24Rows.push([p.rank, p.name, p.category, p.unit, p.total_qty, p.total_sales, p.pct / 100]);
      }});
      const ws24 = XLSX.utils.aoa_to_sheet(p24Rows);
      XLSX.utils.book_append_sheet(wb, ws24, '129 Produk 2024');

      XLSX.writeFile(wb, 'Laporan_Penjualan_MultiTahun_Snaprint.xlsx');
    }}

    // Initial Load
    window.addEventListener('DOMContentLoaded', () => {{
      selectYear('2025'); // Default to 2025 as requested
    }});
  </script>
</body>
</html>
"""

# Write workspace index.html
with open("index.html", "w", encoding="utf-8") as f:
    f.write(unified_html)
print("Updated index.html to unified multi-year dynamic dashboard!")

# Write 2025/index.html as well for direct access in 2025 folder
html_2025 = unified_html.replace('href="Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"', 'href="../Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"')
with open("2025/index.html", "w", encoding="utf-8") as f:
    f.write(html_2025)
print("Created 2025/index.html!")

# Write artifact
artifact_path = r"C:\Users\Swanto\.gemini\antigravity\brain\186463fa-99d6-481c-8f3f-6ce410e3a005\dashboard_snaprint_unified.html"
with open(artifact_path, "w", encoding="utf-8") as f:
    f.write(unified_html)
print(f"Generated artifact: {artifact_path}")
