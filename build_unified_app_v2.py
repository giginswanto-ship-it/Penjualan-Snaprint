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
  <title>SNAPRINT DIGITAL — Sistem Dashboard Penjualan Terintegrasi Multi-Tahun (2024, 2025, 2026)</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <!-- SheetJS untuk Export Excel Multi-Sheet -->
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
          <div class="w-12 h-12 rounded-xl bg-gradient-to-tr from-sky-600 via-indigo-600 to-emerald-600 flex items-center justify-center text-white font-extrabold text-2xl shadow-md shadow-sky-600/20 shrink-0">
            S
          </div>
          <div>
            <div class="flex items-center space-x-2 flex-wrap">
              <h1 class="text-base sm:text-lg font-extrabold tracking-tight text-[var(--foreground,#0f172a)]">SNAPRINT DIGITAL PRINTING</h1>
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-indigo-500/10 text-indigo-700 border border-indigo-500/20">SISTEM DINAMIS TERPADU (2024 - 2026)</span>
            </div>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Platform Analisis Kinerja Penjualan Multi-Tahun — Konsolidasi Laporan ACCURATE Accounting System
            </p>
          </div>
        </div>

        <!-- Quick Export & PDF Actions -->
        <div class="flex items-center space-x-2 sm:space-x-2.5 self-end md:self-center flex-wrap gap-y-2">
          <button onclick="exportUnifiedExcel()" class="px-3 py-1.5 sm:px-3.5 sm:py-2 text-xs font-semibold rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white transition-all shadow-sm flex items-center space-x-1.5">
            <span>📊</span>
            <span>Export Excel</span>
          </button>
          <a href="snaprint_sales.sql" download="snaprint_sales.sql" class="px-2.5 py-1.5 sm:px-3 sm:py-2 text-xs font-semibold rounded-xl bg-purple-700 hover:bg-purple-600 text-white transition-all shadow-sm flex items-center space-x-1">
            <span>💾</span>
            <span>Database SQL</span>
          </a>
          <a id="btn-pdf-2024" href="Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf" target="_blank" class="px-2.5 py-1.5 sm:px-3 sm:py-2 text-xs font-semibold rounded-xl bg-slate-900 hover:bg-slate-800 text-white transition-all shadow-sm flex items-center space-x-1">
            <span>📄</span>
            <span>PDF 2024</span>
          </a>
          <a id="btn-pdf-2025" href="Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf" target="_blank" class="px-2.5 py-1.5 sm:px-3 sm:py-2 text-xs font-semibold rounded-xl bg-indigo-700 hover:bg-indigo-600 text-white transition-all shadow-sm flex items-center space-x-1">
            <span>📄</span>
            <span>PDF 2025</span>
          </a>
          <a id="btn-pdf-2026" href="Laporan_Penjualan_2026_Snaprint_Digital.pdf" target="_blank" class="px-2.5 py-1.5 sm:px-3 sm:py-2 text-xs font-semibold rounded-xl bg-amber-600 hover:bg-amber-500 text-white transition-all shadow-sm flex items-center space-x-1">
            <span>📄</span>
            <span>PDF 2026</span>
          </a>
        </div>

      </div>

      <!-- YEAR SELECTOR TABS BAR -->
      <div class="pt-3 border-t border-[var(--border,#f1f5f9)] flex flex-wrap items-center justify-between gap-3">
        <div class="flex flex-wrap items-center gap-1.5 bg-slate-100/90 p-1 rounded-xl border border-slate-200">
          <button onclick="selectYear('2024')" id="year-btn-2024" class="px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            📅 2024 (Rp 930,8 Jt)
          </button>
          <button onclick="selectYear('2025')" id="year-btn-2025" class="px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            🚀 2025 (Rp 1.378,9 Jt)
          </button>
          <button onclick="selectYear('2026')" id="year-btn-2026" class="px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            🔥 2026 (Rp 1.180,3 Jt - 8 Bln)
          </button>
          <button onclick="selectYear('MULTI')" id="year-btn-MULTI" class="px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            📈 Komparasi Multi-Tahun
          </button>
        </div>

        <div class="flex items-center space-x-2 text-xs">
          <span class="text-[var(--muted-foreground,#64748b)]">Mode Aktif:</span>
          <span id="current-mode-badge" class="font-bold text-amber-700 bg-amber-50 px-3 py-1 rounded-lg border border-amber-200">
            Tahun Buku 2026 (Jan–Agu)
          </span>
        </div>
      </div>

    </header>

    <!-- ======================================================== -->
    <!-- VIEW CONTAINER 1: SINGLE YEAR DASHBOARD (2024/2025/2026) -->
    <!-- ======================================================== -->
    <div id="view-single-year" class="space-y-4 sm:space-y-6">

      <!-- DYNAMIC FILTER CONTROL PANEL -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-3 pb-3 border-b border-[var(--border,#f1f5f9)]">
          <div class="flex items-center space-x-2">
            <span class="text-base">🎛️</span>
            <div>
              <h2 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)]" id="filter-header-title">
                Filter Interaktif &amp; Analisis Kinerja
              </h2>
              <p class="text-xs text-[var(--muted-foreground,#64748b)]">
                Pilih periode bulan, semester, atau filter kuartal untuk memutakhirkan grafik &amp; tabel seketika.
              </p>
            </div>
          </div>
          
          <div class="flex items-center space-x-2 self-start md:self-auto">
            <span class="text-xs text-[var(--muted-foreground,#64748b)] font-medium">Bulan Aktif:</span>
            <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-sky-100 text-sky-800" id="active-period-label">
              Semua Periode
            </span>
            <button onclick="resetFilters()" class="text-xs text-rose-600 hover:text-rose-700 font-medium ml-2 underline">
              Reset Filter
            </button>
          </div>
        </div>

        <!-- Filter Buttons Grid -->
        <div class="grid grid-cols-1 md:grid-cols-12 gap-3.5 items-center">
          
          <!-- Month/Period Selector (8 Cols) -->
          <div class="md:col-span-8 space-y-1.5">
            <label class="text-xs font-semibold text-[var(--muted-foreground,#64748b)] block">PILIH PERIODE WAKTU:</label>
            <div class="flex flex-wrap gap-1.5" id="month-buttons-container"></div>
          </div>

          <!-- Category Selector (4 Cols) -->
          <div class="md:col-span-4 space-y-1.5">
            <label class="text-xs font-semibold text-[var(--muted-foreground,#64748b)] block">KATEGORI PRODUK:</label>
            <select id="category-filter-select" onchange="onCategorySelect(this.value)" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs text-slate-800 font-medium focus:ring-2 focus:ring-sky-500 focus:outline-none">
              <option value="ALL">Semua Kategori Produk</option>
            </select>
          </div>

        </div>

        <!-- Secondary Filters Row for 2024 Product View -->
        <div class="pt-2 border-t border-[var(--border,#f1f5f9)] flex flex-wrap items-center justify-between gap-3" id="filter-pareto-box">
          <div class="flex items-center space-x-2">
            <span class="text-xs text-[var(--muted-foreground,#64748b)] font-semibold">Analisis Pareto:</span>
            <div class="inline-flex rounded-lg border border-slate-200 bg-slate-100 p-0.5 text-xs">
              <button onclick="setParetoFilter('ALL')" id="pareto-btn-ALL" class="px-2.5 py-1 font-bold rounded-md bg-white text-slate-800 shadow-xs">Semua (129)</button>
              <button onclick="setParetoFilter('CORE80')" id="pareto-btn-CORE80" class="px-2.5 py-1 font-medium rounded-md text-slate-600 hover:text-slate-900">⭐ Core 80% (24)</button>
              <button onclick="setParetoFilter('LONGTAIL')" id="pareto-btn-LONGTAIL" class="px-2.5 py-1 font-medium rounded-md text-slate-600 hover:text-slate-900">Pelengkap (105)</button>
            </div>
          </div>

          <!-- Search input for individual products -->
          <div class="relative w-full sm:w-64">
            <span class="absolute inset-y-0 left-0 flex items-center pl-2.5 pointer-events-none text-slate-400">🔍</span>
            <input type="text" id="product-search-input" oninput="onSearchInput(this.value)" placeholder="Cari nama barang..." class="w-full pl-8 pr-3 py-1.5 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-800 focus:ring-2 focus:ring-sky-500 focus:outline-none">
          </div>
        </div>

      </div>

      <!-- 4 EXECUTIVE KPI SUMMARY CARDS -->
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5">
        
        <!-- KPI 1: Total Omzet -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 shadow-sm relative overflow-hidden flex flex-col justify-between">
          <div class="absolute -right-2 -bottom-2 w-16 h-16 bg-sky-500/5 rounded-full pointer-events-none"></div>
          <div>
            <div class="flex items-center justify-between text-xs text-[var(--muted-foreground,#64748b)] mb-1">
              <span class="font-semibold uppercase tracking-wider" id="kpi-sales-title">TOTAL OMZET PENJUALAN</span>
              <span class="text-base">💰</span>
            </div>
            <div class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-total-sales">
              Rp 0
            </div>
          </div>
          <div class="mt-3 pt-2.5 border-t border-[var(--border,#f1f5f9)] text-xs flex items-center justify-between">
            <span class="text-[var(--muted-foreground,#64748b)]" id="kpi-sales-sub-label">Pertumbuhan vs Rerata:</span>
            <span class="font-bold text-sky-600" id="kpi-sales-sub-val">100%</span>
          </div>
        </div>

        <!-- KPI 2: Total Volume -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 shadow-sm relative overflow-hidden flex flex-col justify-between">
          <div class="absolute -right-2 -bottom-2 w-16 h-16 bg-indigo-500/5 rounded-full pointer-events-none"></div>
          <div>
            <div class="flex items-center justify-between text-xs text-[var(--muted-foreground,#64748b)] mb-1">
              <span class="font-semibold uppercase tracking-wider">TOTAL KUANTITAS PRODUKSI</span>
              <span class="text-base">📦</span>
            </div>
            <div class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-total-qty">
              0
            </div>
          </div>
          <div class="mt-3 pt-2.5 border-t border-[var(--border,#f1f5f9)] text-xs flex items-center justify-between">
            <span class="text-[var(--muted-foreground,#64748b)]">Satuan Produksi:</span>
            <span class="font-bold text-indigo-600">unit / m² / lbr</span>
          </div>
        </div>

        <!-- KPI 3: Rata-Rata Bulanan -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 shadow-sm relative overflow-hidden flex flex-col justify-between">
          <div class="absolute -right-2 -bottom-2 w-16 h-16 bg-emerald-500/5 rounded-full pointer-events-none"></div>
          <div>
            <div class="flex items-center justify-between text-xs text-[var(--muted-foreground,#64748b)] mb-1">
              <span class="font-semibold uppercase tracking-wider">RATA-RATA BULANAN</span>
              <span class="text-base">📅</span>
            </div>
            <div class="text-xl sm:text-2xl font-extrabold text-[var(--foreground,#0f172a)] font-mono tracking-tight" id="kpi-avg-sales">
              Rp 0
            </div>
          </div>
          <div class="mt-3 pt-2.5 border-t border-[var(--border,#f1f5f9)] text-xs flex items-center justify-between">
            <span class="text-[var(--muted-foreground,#64748b)]" id="kpi-avg-sub-label">Stabilitas Omzet:</span>
            <span class="font-bold text-emerald-600" id="kpi-avg-sub-val">Konsisten</span>
          </div>
        </div>

        <!-- KPI 4: Bulan Puncak -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 shadow-sm relative overflow-hidden flex flex-col justify-between">
          <div class="absolute -right-2 -bottom-2 w-16 h-16 bg-amber-500/5 rounded-full pointer-events-none"></div>
          <div>
            <div class="flex items-center justify-between text-xs text-[var(--muted-foreground,#64748b)] mb-1">
              <span class="font-semibold uppercase tracking-wider">BULAN TERTINGGI (PEAK)</span>
              <span class="text-base">👑</span>
            </div>
            <div class="text-xl sm:text-2xl font-extrabold text-amber-700 font-mono tracking-tight" id="kpi-peak-month">
              -
            </div>
          </div>
          <div class="mt-3 pt-2.5 border-t border-[var(--border,#f1f5f9)] text-xs flex items-center justify-between">
            <span class="text-[var(--muted-foreground,#64748b)]">Omzet Puncak:</span>
            <span class="font-bold text-amber-800 font-mono" id="kpi-peak-val">Rp 0</span>
          </div>
        </div>

      </div>

      <!-- CHARTS SECTION: MONTHLY SALES TREND (2 Cols) & CATEGORY BREAKDOWN (1 Col) -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-3.5">
        
        <!-- Chart 1: Monthly Trend Bar Chart (2 Cols) -->
        <div class="lg:col-span-2 bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div>
              <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
                <span>📊</span>
                <span id="chart-title">Tren Omzet Penjualan Bulanan</span>
              </h3>
              <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
                Setiap batang menunjukkan angka penjualan bersih (Rupiah) &amp; porsi persentase. Klik batang untuk filter!
              </p>
            </div>
            <span class="text-xs px-2.5 py-1 rounded-full bg-sky-50 text-sky-700 font-bold border border-sky-200 shrink-0" id="chart-year-badge">
              Bulan Buku Aktif
            </span>
          </div>

          <!-- Dynamic Live Hover Info Box -->
          <div id="chart-hover-info" class="bg-slate-50 border border-slate-200/80 rounded-xl px-3.5 py-2 text-xs text-slate-700 flex flex-wrap items-center justify-between gap-2 min-h-[38px] my-1">
            <span class="font-medium text-slate-500">💡 Sentuh/arahkan kursor pada batang bulan untuk melihat rincian omzet, volume, &amp; performa:</span>
            <span class="text-sky-600 font-semibold cursor-pointer hover:underline" onclick="setPeriod('ALL')">Tampilkan Semua Periode</span>
          </div>

          <!-- High-Headroom SVG Bar Chart -->
          <div class="w-full overflow-x-auto py-1">
            <svg id="svg-monthly-chart" viewBox="0 0 860 300" class="w-full h-64 sm:h-72 select-none"></svg>
          </div>

          <!-- Legend Bar -->
          <div class="mt-2 pt-2 border-t border-[var(--border,#f1f5f9)] flex flex-wrap items-center justify-between gap-2 text-xs text-[var(--muted-foreground,#64748b)]">
            <div class="flex flex-wrap items-center gap-3">
              <span class="flex items-center space-x-1.5">
                <span class="w-3.5 h-3.5 rounded-md bg-gradient-to-t from-sky-700 to-sky-500 inline-block shadow-sm" id="legend-active-color"></span>
                <span class="font-medium text-slate-700">Omzet Bulan Terpilih</span>
              </span>
              <span class="flex items-center space-x-1.5">
                <span class="w-3.5 h-3.5 rounded-md bg-amber-400 inline-block shadow-sm"></span>
                <span class="font-medium text-amber-800" id="legend-peak-text">Bulan Puncak</span>
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
              Reset ke Periode Penuh
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
              <span class="text-xs text-[var(--muted-foreground,#64748b)]" id="cat-count-badge">6 Kategori</span>
            </div>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mb-3">
              Distribusi kontribusi kategori terhadap omzet:
            </p>

            <!-- Category Progress Bars -->
            <div class="space-y-2.5" id="category-progress-container"></div>
          </div>

          <div class="mt-4 pt-3 border-t border-[var(--border,#f1f5f9)] text-xs text-[var(--muted-foreground,#64748b)] flex items-center justify-between">
            <span>Dua Kategori Inti:</span>
            <span class="font-bold text-sky-700" id="cat-core-text">Menyumbang &gt;70% Omzet</span>
          </div>
        </div>

      </div>

      <!-- TABLE SECTION (ADAPTS ACCORDING TO 2024, 2025, OR 2026) -->
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
            💡 Rincian Matriks Bulanan &amp; Klasifikasi Produk
          </div>
        </div>

        <!-- Dynamic Container for Table -->
        <div id="dynamic-table-container"></div>

      </div>

    </div>

    <!-- ======================================================== -->
    <!-- VIEW CONTAINER 2: MULTI-YEAR COMPARISON (2024, 2025, 2026) -->
    <!-- ======================================================== -->
    <div id="view-multi-year" class="hidden space-y-4 sm:space-y-6">

      <!-- MULTI-YEAR EXECUTIVE OVERVIEW -->
      <div class="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-2xl p-5 sm:p-6 shadow-md space-y-4">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div class="flex items-center space-x-2">
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">LOMPATAN BISNIS EKSPONENSIAL</span>
              <span class="text-xs text-slate-300">Periode Komparasi: 01 Jan s/d 31 Agu (8 Bulan Apel-ke-Apel)</span>
            </div>
            <h2 class="text-lg sm:text-2xl font-black tracking-tight mt-1">
              Analisis Komparasi Kinerja Multi-Tahun (2024 vs 2025 vs 2026)
            </h2>
            <p class="text-xs sm:text-sm text-slate-300 max-w-3xl mt-1">
              Evaluasi komparatif perkembangan bisnis Snaprint Digital Printing secara konsisten menunjukkan percepatan omzet: dari Rp 508,3 Jt di 2024 melonjak ke Rp 859,7 Jt di 2025 (+69,1%), dan melesat ke <b>Rp 1.180,3 Jt di 2026 (+37,3% vs 2025 atau +132,2% vs 2024)</b>!
            </p>
          </div>

          <div class="shrink-0 bg-white/10 backdrop-blur-md p-3.5 rounded-xl border border-white/20 text-center">
            <span class="text-[11px] uppercase tracking-wider text-emerald-300 font-semibold block">Pertumbuhan 2 Tahun (8M)</span>
            <span class="text-2xl sm:text-3xl font-extrabold text-white font-mono block mt-0.5">+132,22%</span>
            <span class="text-[10px] text-slate-300 block">+Rp 672,04 Juta Omzet</span>
          </div>
        </div>

        <!-- 3-Year Metric Mini-Cards -->
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-2 border-t border-white/10">
          <div class="bg-white/5 rounded-xl p-3 border border-white/10">
            <span class="text-xs text-slate-400 block font-medium">8 Bulan Pertama 2024:</span>
            <span class="text-base sm:text-lg font-bold font-mono text-white block mt-0.5">Rp 508.264.736</span>
            <span class="text-[11px] text-slate-400 block">Rata-rata: Rp 63,5 Jt / bln</span>
          </div>
          <div class="bg-white/5 rounded-xl p-3 border border-white/10">
            <span class="text-xs text-slate-400 block font-medium">8 Bulan Pertama 2025:</span>
            <span class="text-base sm:text-lg font-bold font-mono text-white block mt-0.5">Rp 859.742.115</span>
            <span class="text-[11px] text-emerald-400 block font-semibold">+69,15% YoY (+Rp 351,5 Jt)</span>
          </div>
          <div class="bg-emerald-500/20 rounded-xl p-3 border border-emerald-500/40">
            <span class="text-xs text-emerald-200 block font-medium">8 Bulan Pertama 2026:</span>
            <span class="text-base sm:text-lg font-black font-mono text-emerald-300 block mt-0.5">Rp 1.180.308.204</span>
            <span class="text-[11px] text-emerald-200 block font-semibold">+37,29% vs 2025 (+Rp 320,6 Jt)</span>
          </div>
        </div>
      </div>

      <!-- MULTI-YEAR COMPARISON BAR CHART -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
          <div>
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>📊</span>
              <span>Grafik Perbandingan Omzet Bulanan (Januari – Agustus: 2024 vs 2025 vs 2026)</span>
            </h3>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Kelompok 3 batang per bulan menunjukkan tren pertumbuhan yang konsisten naik di setiap bulannya.
            </p>
          </div>
          <div class="flex items-center space-x-3 text-xs">
            <span class="flex items-center space-x-1.5">
              <span class="w-3 h-3 rounded-xs bg-slate-400 inline-block"></span>
              <span class="font-semibold text-slate-700">2024</span>
            </span>
            <span class="flex items-center space-x-1.5">
              <span class="w-3 h-3 rounded-xs bg-sky-500 inline-block"></span>
              <span class="font-semibold text-sky-800">2025</span>
            </span>
            <span class="flex items-center space-x-1.5">
              <span class="w-3 h-3 rounded-xs bg-emerald-600 inline-block"></span>
              <span class="font-semibold text-emerald-800">2026</span>
            </span>
          </div>
        </div>

        <!-- SVG Grouped Bar Chart -->
        <div class="w-full overflow-x-auto py-2">
          <svg id="svg-multiyear-chart" viewBox="0 0 860 300" class="w-full h-64 sm:h-72 select-none"></svg>
        </div>
      </div>

      <!-- DETAILED 3-YEAR COMPARISON TABLE -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-3 border-b border-[var(--border,#f1f5f9)] gap-2">
          <div>
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>📋</span>
              <span>Tabel Rincian Komparasi Bulanan (Januari – Agustus)</span>
            </h3>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Data penjualan bersih, pertumbuhan nominal, dan persentase akselerasi bulanan.
            </p>
          </div>
          <span class="text-xs font-bold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200">
            8 Bulan Buku Apel-ke-Apel
          </span>
        </div>

        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
              <tr>
                <th class="py-2.5 px-3">Bulan</th>
                <th class="py-2.5 px-3 text-right">Omzet 2024 (Rp)</th>
                <th class="py-2.5 px-3 text-right">Omzet 2025 (Rp)</th>
                <th class="py-2.5 px-3 text-right">Omzet 2026 (Rp)</th>
                <th class="py-2.5 px-3 text-right">Pertumbuhan 2026 vs 2025</th>
                <th class="py-2.5 px-3 text-right">Pertumbuhan 2026 vs 2024</th>
              </tr>
            </thead>
            <tbody id="multiyear-tbody" class="divide-y divide-slate-100 font-mono text-slate-800"></tbody>
          </table>
        </div>
      </div>

    </div>

  </div>

  <!-- MODAL DETAIL PRODUK (KHUSUS 2024) -->
  <div id="product-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-3 hidden">
    <div class="bg-white rounded-2xl max-w-xl w-full border border-slate-200 shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
      <div class="p-4 bg-slate-900 text-white flex items-center justify-between">
        <div>
          <span class="text-[10px] font-bold uppercase tracking-wider text-sky-400" id="modal-category">Kategori</span>
          <h3 class="text-sm sm:text-base font-bold text-white mt-0.5" id="modal-title">Nama Produk</h3>
        </div>
        <button onclick="closeModal()" class="w-7 h-7 rounded-lg bg-white/10 hover:bg-white/20 flex items-center justify-center text-white text-base">✕</button>
      </div>

      <div class="p-4 space-y-3.5">
        <div class="grid grid-cols-3 gap-2 text-center">
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
            <span class="text-[11px] text-slate-500 font-semibold uppercase">Total Omzet</span>
            <p class="text-sm sm:text-base font-bold text-slate-900 font-mono mt-0.5" id="modal-sales">Rp 0</p>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
            <span class="text-[11px] text-slate-500 font-semibold uppercase">Total Volume</span>
            <p class="text-sm sm:text-base font-bold text-slate-900 font-mono mt-0.5" id="modal-qty">0</p>
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
      activeYear: '2026', // '2024', '2025', '2026', or 'MULTI'
      period: 'ALL',      // 'ALL', 'H1', 'H2', 'Q1'..'Q4', or month code
      category: 'ALL',
      pareto: 'ALL',
      search: '',
      sortCol: 'sales',
      sortAsc: false
    }};

    const MONTH_KEYS_12 = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    const MONTH_KEYS_8 = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu'];
    const MONTH_NAMES = {{
      'Jan': 'Januari', 'Feb': 'Februari', 'Mar': 'Maret', 'Apr': 'April',
      'Mei': 'Mei', 'Jun': 'Juni', 'Jul': 'Juli', 'Agu': 'Agustus',
      'Sep': 'September', 'Okt': 'Oktober', 'Nov': 'November', 'Des': 'Desember'
    }};

    function fmtIDR(num) {{
      return 'Rp ' + Math.round(num).toLocaleString('id-ID');
    }}

    function fmtNum(num) {{
      return Number(num).toLocaleString('id-ID', {{ minimumFractionDigits: 0, maximumFractionDigits: 2 }});
    }}

    // Year Switching
    function selectYear(yr) {{
      appState.activeYear = yr;
      appState.period = 'ALL';
      appState.category = 'ALL';
      appState.search = '';
      appState.pareto = 'ALL';

      // Update Tab Buttons
      ['2024', '2025', '2026', 'MULTI'].forEach(y => {{
        const btn = document.getElementById('year-btn-' + y);
        if (y === yr) {{
          if (y === '2026') {{
            btn.className = 'px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all bg-amber-600 text-white shadow-sm';
          }} else if (y === '2025') {{
            btn.className = 'px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all bg-emerald-600 text-white shadow-sm';
          }} else if (y === '2024') {{
            btn.className = 'px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all bg-sky-600 text-white shadow-sm';
          }} else {{
            btn.className = 'px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-sm';
          }}
        }} else {{
          btn.className = 'px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900';
        }}
      }});

      const modeBadge = document.getElementById('current-mode-badge');
      if (yr === '2024') {{
        modeBadge.innerText = 'Tahun Buku 2024 (Rp 930,85 Jt)';
        modeBadge.className = 'font-bold text-sky-700 bg-sky-50 px-3 py-1 rounded-lg border border-sky-200';
      }} else if (yr === '2025') {{
        modeBadge.innerText = 'Tahun Buku 2025 (Rp 1.378,95 Jt)';
        modeBadge.className = 'font-bold text-emerald-700 bg-emerald-50 px-3 py-1 rounded-lg border border-emerald-200';
      }} else if (yr === '2026') {{
        modeBadge.innerText = 'Tahun Buku 2026 (Rp 1.180,31 Jt | Jan–Agu)';
        modeBadge.className = 'font-bold text-amber-700 bg-amber-50 px-3 py-1 rounded-lg border border-amber-200';
      }} else {{
        modeBadge.innerText = 'Komparasi Multi-Tahun (2024 vs 2025 vs 2026)';
        modeBadge.className = 'font-bold text-indigo-700 bg-indigo-50 px-3 py-1 rounded-lg border border-indigo-200';
      }}

      // Switch views
      if (yr === 'MULTI') {{
        document.getElementById('view-single-year').classList.add('hidden');
        document.getElementById('view-multi-year').classList.remove('hidden');
        renderMultiYearView();
      }} else {{
        document.getElementById('view-single-year').classList.remove('hidden');
        document.getElementById('view-multi-year').classList.add('hidden');
        
        // Hide/show pareto & search (only for 2024 product breakdown)
        document.getElementById('filter-pareto-box').style.display = (yr === '2024') ? 'flex' : 'none';

        initPeriodButtons();
        initCategoryFilterOptions();
        updateDashboard();
      }}
    }}

    // Get Active Dataset
    function getActiveYearData() {{
      if (appState.activeYear === '2024') return MASTER_DATA.data_2024;
      if (appState.activeYear === '2025') return MASTER_DATA.data_2025;
      if (appState.activeYear === '2026') return MASTER_DATA.data_2026;
      return MASTER_DATA.data_2026;
    }}

    // Month Period Buttons
    function initPeriodButtons() {{
      const container = document.getElementById('month-buttons-container');
      container.innerHTML = '';

      const is2026 = (appState.activeYear === '2026');
      const monthKeys = is2026 ? MONTH_KEYS_8 : MONTH_KEYS_12;

      // "SEMUA" button
      const allBtn = document.createElement('button');
      allBtn.id = 'period-btn-ALL';
      allBtn.className = 'px-3 py-1 text-xs font-bold rounded-lg border transition-all';
      allBtn.innerText = is2026 ? 'Jan – Agu (8 Bln)' : 'Semua (12 Bln)';
      allBtn.onclick = () => setPeriod('ALL');
      container.appendChild(allBtn);

      // Semester buttons
      const h1Btn = document.createElement('button');
      h1Btn.id = 'period-btn-H1';
      h1Btn.className = 'px-2.5 py-1 text-xs font-medium rounded-lg border transition-all';
      h1Btn.innerText = 'H1 (Jan-Jun)';
      h1Btn.onclick = () => setPeriod('H1');
      container.appendChild(h1Btn);

      if (!is2026) {{
        const h2Btn = document.createElement('button');
        h2Btn.id = 'period-btn-H2';
        h2Btn.className = 'px-2.5 py-1 text-xs font-medium rounded-lg border transition-all';
        h2Btn.innerText = 'H2 (Jul-Des)';
        h2Btn.onclick = () => setPeriod('H2');
        container.appendChild(h2Btn);
      }}

      // Quarter buttons
      const q1Btn = document.createElement('button');
      q1Btn.id = 'period-btn-Q1';
      q1Btn.className = 'px-2 py-1 text-xs font-medium rounded-lg border transition-all';
      q1Btn.innerText = 'Q1';
      q1Btn.onclick = () => setPeriod('Q1');
      container.appendChild(q1Btn);

      const q2Btn = document.createElement('button');
      q2Btn.id = 'period-btn-Q2';
      q2Btn.className = 'px-2 py-1 text-xs font-medium rounded-lg border transition-all';
      q2Btn.innerText = 'Q2';
      q2Btn.onclick = () => setPeriod('Q2');
      container.appendChild(q2Btn);

      const q3Btn = document.createElement('button');
      q3Btn.id = 'period-btn-Q3';
      q3Btn.className = 'px-2 py-1 text-xs font-medium rounded-lg border transition-all';
      q3Btn.innerText = is2026 ? 'Q3 (Jul-Agu)' : 'Q3';
      q3Btn.onclick = () => setPeriod('Q3');
      container.appendChild(q3Btn);

      if (!is2026) {{
        const q4Btn = document.createElement('button');
        q4Btn.id = 'period-btn-Q4';
        q4Btn.className = 'px-2 py-1 text-xs font-medium rounded-lg border transition-all';
        q4Btn.innerText = 'Q4';
        q4Btn.onclick = () => setPeriod('Q4');
        container.appendChild(q4Btn);
      }}

      // Divider
      const sep = document.createElement('span');
      sep.className = 'text-slate-300 self-center mx-1';
      sep.innerText = '|';
      container.appendChild(sep);

      // Month keys
      monthKeys.forEach(m => {{
        const b = document.createElement('button');
        b.id = 'period-btn-' + m;
        b.className = 'px-2 py-1 text-xs font-medium rounded-lg border transition-all';
        b.innerText = m;
        b.onclick = () => setPeriod(m);
        container.appendChild(b);
      }});

      updatePeriodButtonStyles();
    }}

    function updatePeriodButtonStyles() {{
      const allPeriodIds = ['ALL', 'H1', 'H2', 'Q1', 'Q2', 'Q3', 'Q4', ...MONTH_KEYS_12];
      allPeriodIds.forEach(id => {{
        const btn = document.getElementById('period-btn-' + id);
        if (!btn) return;
        if (id === appState.period) {{
          btn.className = 'px-2.5 py-1 text-xs font-bold rounded-lg border bg-sky-600 text-white border-sky-600 shadow-xs';
        }} else {{
          btn.className = 'px-2 py-1 text-xs font-medium rounded-lg border bg-white border-slate-200 text-slate-700 hover:bg-slate-50';
        }}
      }});

      const activeLabel = document.getElementById('active-period-label');
      if (appState.period === 'ALL') {{
        activeLabel.innerText = (appState.activeYear === '2026') ? 'Januari – Agustus 2026 (8 Bulan)' : 'Semua Bulan (1 Tahun)';
      }} else {{
        activeLabel.innerText = 'Filter: ' + appState.period;
      }}
    }}

    function setPeriod(p) {{
      appState.period = p;
      updatePeriodButtonStyles();
      updateDashboard();
    }}

    function initCategoryFilterOptions() {{
      const yrData = getActiveYearData();
      const select = document.getElementById('category-filter-select');
      select.innerHTML = '<option value="ALL">Semua Kategori Produk</option>';

      yrData.categories.forEach(cat => {{
        const opt = document.createElement('option');
        opt.value = cat.name;
        opt.innerText = cat.name + ' (' + cat.pct.toFixed(1) + '%)';
        select.appendChild(opt);
      }});
      select.value = appState.category;
    }}

    function onCategorySelect(val) {{
      appState.category = val;
      updateDashboard();
    }}

    function setParetoFilter(val) {{
      appState.pareto = val;
      ['ALL', 'CORE80', 'LONGTAIL'].forEach(p => {{
        const btn = document.getElementById('pareto-btn-' + p);
        if (!btn) return;
        if (p === val) {{
          btn.className = 'px-2.5 py-1 font-bold rounded-md bg-white text-slate-900 shadow-xs';
        }} else {{
          btn.className = 'px-2.5 py-1 font-medium rounded-md text-slate-600 hover:text-slate-900';
        }}
      }});
      updateDashboard();
    }}

    function onSearchInput(val) {{
      appState.search = val.trim().toLowerCase();
      updateDashboard();
    }}

    function resetFilters() {{
      appState.period = 'ALL';
      appState.category = 'ALL';
      appState.pareto = 'ALL';
      appState.search = '';
      const sInput = document.getElementById('product-search-input');
      if (sInput) sInput.value = '';
      const catSelect = document.getElementById('category-filter-select');
      if (catSelect) catSelect.value = 'ALL';
      updatePeriodButtonStyles();
      setParetoFilter('ALL');
      updateDashboard();
    }}

    // Get Active Months according to Period
    function getActiveMonths() {{
      const is2026 = (appState.activeYear === '2026');
      const maxMonths = is2026 ? MONTH_KEYS_8 : MONTH_KEYS_12;

      if (appState.period === 'ALL') return maxMonths;
      if (appState.period === 'H1') return ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun'];
      if (appState.period === 'H2') return is2026 ? ['Jul', 'Agu'] : ['Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
      if (appState.period === 'Q1') return ['Jan', 'Feb', 'Mar'];
      if (appState.period === 'Q2') return ['Apr', 'Mei', 'Jun'];
      if (appState.period === 'Q3') return is2026 ? ['Jul', 'Agu'] : ['Jul', 'Agu', 'Sep'];
      if (appState.period === 'Q4') return ['Okt', 'Nov', 'Des'];
      return [appState.period];
    }}

    // Update Dashboard Views
    function updateDashboard() {{
      const yrData = getActiveYearData();
      renderKPICards(yrData);
      renderMonthlySvgChart(yrData);
      renderCategoryProgressBars(yrData);
      renderDynamicTable(yrData);
    }}

    // Render KPI Cards
    function renderKPICards(yrData) {{
      const activeMonths = getActiveMonths();
      const isFull = (activeMonths.length === (appState.activeYear === '2026' ? 8 : 12));

      let currentSales = 0;
      let currentQty = 0;

      yrData.monthly.forEach(m => {{
        if (activeMonths.includes(m.code)) {{
          currentSales += m.sales;
          currentQty += m.qty;
        }}
      }});

      document.getElementById('kpi-total-sales').innerText = fmtIDR(currentSales);
      document.getElementById('kpi-total-qty').innerText = fmtNum(currentQty);

      const numMonths = activeMonths.length || 1;
      const avgSales = currentSales / numMonths;
      document.getElementById('kpi-avg-sales').innerText = fmtIDR(avgSales);

      // Peak month in active selection
      let peakM = null;
      let maxS = -1;
      yrData.monthly.forEach(m => {{
        if (activeMonths.includes(m.code) && m.sales > maxS) {{
          maxS = m.sales;
          peakM = m;
        }}
      }});

      if (peakM) {{
        document.getElementById('kpi-peak-month').innerText = peakM.name + ' ' + appState.activeYear;
        document.getElementById('kpi-peak-val').innerText = fmtIDR(peakM.sales);
      }}

      // Sub-labels
      const pctOfYear = (currentSales / yrData.kpi.total_sales) * 100;
      document.getElementById('kpi-sales-sub-val').innerText = pctOfYear.toFixed(1) + '% dari Total';

      if (appState.activeYear === '2026') {{
        document.getElementById('kpi-sales-title').innerText = 'TOTAL OMZET (8 BULAN)';
        document.getElementById('kpi-avg-sub-val').innerText = '+28,4% vs 2025';
      }} else if (appState.activeYear === '2025') {{
        document.getElementById('kpi-sales-title').innerText = 'TOTAL OMZET PENJUALAN';
        document.getElementById('kpi-avg-sub-val').innerText = '+48,1% vs 2024';
      }} else {{
        document.getElementById('kpi-sales-title').innerText = 'TOTAL OMZET PENJUALAN';
        document.getElementById('kpi-avg-sub-val').innerText = 'Baseline 2024';
      }}
    }}

    // SVG Monthly Bar Chart
    function renderMonthlySvgChart(yrData) {{
      const svg = document.getElementById('svg-monthly-chart');
      svg.innerHTML = '';

      const is2026 = (appState.activeYear === '2026');
      const activeMonths = getActiveMonths();
      const monthlyData = yrData.monthly;
      const annualAvgSales = yrData.kpi.avg_monthly_sales;

      const chartLeft = 70;
      const chartRight = 835;
      const chartWidth = chartRight - chartLeft;
      const chartTop = 50;
      const chartBottom = 240;
      const chartHeight = chartBottom - chartTop;

      let maxScale = 145000000;
      let gridStep = 25000000;
      if (appState.activeYear === '2025') {{
        maxScale = 160000000;
        gridStep = 35000000;
      }} else if (appState.activeYear === '2026') {{
        maxScale = 200000000; // 200 Jt for 2026 (Peak Juli 173.7 Jt)
        gridStep = 40000000;
      }}

      const numBars = monthlyData.length; // 8 for 2026, 12 for 2024/2025
      const barWidth = is2026 ? 64 : 46;
      const gap = (chartWidth - (numBars * barWidth)) / (numBars + 1);

      // Gradients
      const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
      let gradStop0 = '#0284c7';
      let gradStop1 = '#0369a1';
      if (appState.activeYear === '2025') {{
        gradStop0 = '#059669';
        gradStop1 = '#047857';
      }} else if (appState.activeYear === '2026') {{
        gradStop0 = '#d97706';
        gradStop1 = '#b45309';
      }}

      defs.innerHTML = `
        <linearGradient id="grad-active" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="${{gradStop0}}" />
          <stop offset="100%" stop-color="${{gradStop1}}" />
        </linearGradient>
        <filter id="shadow-pill" x="-10%" y="-10%" width="120%" height="130%">
          <feDropShadow dx="0" dy="1.5" stdDeviation="1.5" flood-color="#0f172a" flood-opacity="0.1"/>
        </filter>
      `;
      svg.appendChild(defs);

      // Gridlines
      for (let s = gridStep; s < maxScale; s += gridStep) {{
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
      avgLine.setAttribute('x2', chartRight - 135);
      avgLine.setAttribute('y2', avgY);
      avgLine.setAttribute('stroke', '#f43f5e');
      avgLine.setAttribute('stroke-dasharray', '5,3');
      avgLine.setAttribute('stroke-width', '1.5');
      svg.appendChild(avgLine);

      const avgTagG = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      avgTagG.innerHTML = `
        <rect x="${{chartRight - 132}}" y="${{avgY - 10}}" width="132" height="20" rx="5" fill="#fff1f2" stroke="#fda4af" stroke-width="1" />
        <text x="${{chartRight - 66}}" y="${{avgY + 4}}" font-size="9.5" font-weight="700" fill="#e11d48" text-anchor="middle">Rata-rata: ${{fmtIDR(annualAvgSales)}}</text>
      `;
      svg.appendChild(avgTagG);

      // Peak Month identifier
      const peakMonthName = yrData.kpi.peak_month;

      // Bars
      monthlyData.forEach((m, i) => {{
        const x = chartLeft + gap + i * (barWidth + gap);
        const barH = (m.sales / maxScale) * chartHeight;
        const barTop = chartBottom - barH;
        const centerX = x + barWidth / 2;
        const isSelected = activeMonths.includes(m.code);
        const isPeak = (m.name.includes(peakMonthName));

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
            <span class="text-sky-600 font-semibold cursor-pointer hover:underline" onclick="setPeriod('ALL')">Tampilkan Semua Periode</span>
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
          rect.setAttribute('stroke-width', '2');
        }}
        g.appendChild(rect);

        // Badge pill
        const bW = is2026 ? 60 : 54;
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
        txtMonth.setAttribute('font-size', is2026 ? '12' : '11.5');
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

      // Update Legend Text
      document.getElementById('legend-peak-text').innerText = `Bulan Puncak (${{peakMonthName}})`;
      document.getElementById('chart-title').innerText = `Tren Omzet Bulanan Snaprint (${{appState.activeYear}})`;
      document.getElementById('chart-year-badge').innerText = is2026 ? '8 Bulan (Jan–Agu)' : '12 Bulan Buku';
    }}

    // Category Progress Bars
    function renderCategoryProgressBars(yrData) {{
      const container = document.getElementById('category-progress-container');
      container.innerHTML = '';

      yrData.categories.forEach(cat => {{
        const div = document.createElement('div');
        div.className = 'p-2.5 rounded-xl transition-all cursor-pointer border bg-slate-50/70 border-slate-200/70 hover:bg-slate-100';
        div.onclick = () => {{
          document.getElementById('category-filter-select').value = cat.name;
          onCategorySelect(cat.name);
        }};
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

    // Dynamic Table Renderer
    function renderDynamicTable(yrData) {{
      if (appState.activeYear === '2024') {{
        render2024ProductsTable(yrData);
      }} else if (appState.activeYear === '2025') {{
        render2025MatrixTable(yrData);
      }} else {{
        render2026MatrixTable(yrData);
      }}
    }}

    // 2024 Table
    function render2024ProductsTable(yrData) {{
      const container = document.getElementById('dynamic-table-container');
      document.getElementById('table-title').innerText = 'Daftar 129 Produk Penjualan Snaprint (2024)';
      document.getElementById('table-subtitle').innerText = 'Klik nama produk untuk melihat modal grafik riwayat 12 bulan & rincian kuantitas.';
      document.getElementById('table-help-text').innerText = '💡 Klik baris produk untuk riwayat transaksi';

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
                <th class="py-2.5 px-3 text-center">Satuan</th>
                <th class="py-2.5 px-3 text-right">Volume</th>
                <th class="py-2.5 px-3 text-right">Total Penjualan (Rp)</th>
                <th class="py-2.5 px-3 text-right">Porsi</th>
                <th class="py-2.5 px-3 text-center">Status</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">${{rowsHtml}}</tbody>
          </table>
        </div>
      `;
    }}

    // 2025 Table
    function render2025MatrixTable(yrData) {{
      const container = document.getElementById('dynamic-table-container');
      document.getElementById('table-title').innerText = 'Matriks Penjualan Bulanan per Lini Kategori (Tahun Buku 2025)';
      document.getElementById('table-subtitle').innerText = 'Konsolidasi 12 bulan buku ACCURATE berdasarkan 5 lini produk utama.';
      document.getElementById('table-help-text').innerText = '💡 Matriks Operasional 2025';

      let rows = yrData.monthly.map(m => `
        <tr class="hover:bg-slate-50 transition-colors">
          <td class="py-2.5 px-3 font-semibold text-slate-900">${{m.name}} 2025 ${{m.name === yrData.kpi.peak_month ? '👑' : ''}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.large_format)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.print_a3)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.doc_fc)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.finishing)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.merchandise)}}</td>
          <td class="py-2.5 px-3 text-right font-mono font-bold text-slate-900 bg-slate-50/50">${{fmtIDR(m.sales)}}</td>
        </tr>
      `).join('');

      container.innerHTML = `
        <div class="overflow-x-auto border border-[var(--border,#e2e8f0)] rounded-xl">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
              <tr>
                <th class="py-2.5 px-3">Bulan</th>
                <th class="py-2.5 px-3 text-right">Large Format</th>
                <th class="py-2.5 px-3 text-right">Print A3+ &amp; Sheet</th>
                <th class="py-2.5 px-3 text-right">Doc &amp; FC</th>
                <th class="py-2.5 px-3 text-right">Finishing</th>
                <th class="py-2.5 px-3 text-right">Merchandise</th>
                <th class="py-2.5 px-3 text-right bg-slate-800">Total Omzet</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">${{rows}}</tbody>
            <tfoot class="bg-slate-100 font-bold border-t-2 border-slate-300">
              <tr>
                <td class="py-3 px-3">TOTAL 2025</td>
                <td class="py-3 px-3 text-right font-mono">Rp 509,0 Jt</td>
                <td class="py-3 px-3 text-right font-mono">Rp 488,6 Jt</td>
                <td class="py-3 px-3 text-right font-mono">Rp 118,9 Jt</td>
                <td class="py-3 px-3 text-right font-mono">Rp 64,8 Jt</td>
                <td class="py-3 px-3 text-right font-mono">Rp 157,6 Jt</td>
                <td class="py-3 px-3 text-right font-mono text-emerald-800 text-sm">Rp 1.378.951.928</td>
              </tr>
            </tfoot>
          </table>
        </div>
      `;
    }}

    // 2026 Table (Exact ACCURATE 6-Category Matrix & Items Classification)
    function render2026MatrixTable(yrData) {{
      const container = document.getElementById('dynamic-table-container');
      document.getElementById('table-title').innerText = '1. Matriks Penjualan Bulanan per Kategori Produk (Tahun Buku 2026 - s/d Agustus)';
      document.getElementById('table-subtitle').innerText = 'Rincian penerimaan 8 bulan buku 2026 berdasarkan 6 kategori produk (ACCURATE Accounting System).';
      document.getElementById('table-help-text').innerText = '💡 Matriks Resmi 2026';

      let rows = yrData.monthly.map(m => `
        <tr class="hover:bg-amber-50/40 transition-colors">
          <td class="py-2.5 px-3 font-semibold text-slate-900">${{m.name}} 2026 ${{m.name === 'Juli' ? '👑 ★' : ''}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.large_format)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.print_a3)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.doc_fc)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.desain_khusus)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.merchandise)}}</td>
          <td class="py-2.5 px-3 text-right font-mono">${{fmtIDR(m.matrix.finishing)}}</td>
          <td class="py-2.5 px-3 text-right font-mono font-bold text-slate-900 bg-amber-50/60">${{fmtIDR(m.sales)}}</td>
        </tr>
      `).join('');

      // Classification cards HTML
      let classCards = yrData.categories.map((c, i) => `
        <div class="p-3 rounded-xl border border-slate-200 bg-slate-50/70">
          <div class="flex items-center justify-between text-xs mb-1">
            <span class="font-bold text-slate-900">${{String.fromCharCode(65 + i)}}. ${{c.name}}</span>
            <span class="font-bold text-amber-700 bg-amber-100 px-2 py-0.5 rounded-md">${{c.pct}}% Omzet</span>
          </div>
          <p class="text-[11px] text-slate-600 line-clamp-3"><b>Item:</b> ${{c.items}}</p>
        </div>
      `).join('');

      container.innerHTML = `
        <div class="space-y-4">
          <div class="overflow-x-auto border border-[var(--border,#e2e8f0)] rounded-xl">
            <table class="w-full text-left text-xs">
              <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
                <tr>
                  <th class="py-2.5 px-3">Periode</th>
                  <th class="py-2.5 px-3 text-right">Large Format</th>
                  <th class="py-2.5 px-3 text-right">Print A3+ &amp; Sheet</th>
                  <th class="py-2.5 px-3 text-right">Doc &amp; Cetak IR</th>
                  <th class="py-2.5 px-3 text-right">Desain &amp; Khusus</th>
                  <th class="py-2.5 px-3 text-right">Merchandise</th>
                  <th class="py-2.5 px-3 text-right">Finishing &amp; Jilid</th>
                  <th class="py-2.5 px-3 text-right bg-slate-800">Total Omzet</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-100">${{rows}}</tbody>
              <tfoot class="bg-slate-100 font-bold border-t-2 border-slate-300">
                <tr>
                  <td class="py-3 px-3">TOTAL 2026 (8 Bln)</td>
                  <td class="py-3 px-3 text-right font-mono">Rp 449,33 Jt</td>
                  <td class="py-3 px-3 text-right font-mono">Rp 389,12 Jt</td>
                  <td class="py-3 px-3 text-right font-mono">Rp 41,29 Jt</td>
                  <td class="py-3 px-3 text-right font-mono">Rp 62,29 Jt</td>
                  <td class="py-3 px-3 text-right font-mono">Rp 75,64 Jt</td>
                  <td class="py-3 px-3 text-right font-mono">Rp 92,64 Jt</td>
                  <td class="py-3 px-3 text-right font-mono text-amber-800 text-sm">Rp 1.180.308.204</td>
                </tr>
              </tfoot>
            </table>
          </div>

          <!-- Product Item Classifications 2026 -->
          <div class="pt-3 border-t border-slate-100">
            <h4 class="font-bold text-xs uppercase tracking-wider text-slate-700 mb-2.5">
              📦 Rincian &amp; Klasifikasi Item Produk per Kategori Bisnis 2026:
            </h4>
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2.5">
              ${{classCards}}
            </div>
          </div>
        </div>
      `;
    }}

    // Multi-Year View Renderer (3 Years Comparison)
    function renderMultiYearView() {{
      const comp = MASTER_DATA.comparison_multiyear;
      const tbody = document.getElementById('multiyear-tbody');
      tbody.innerHTML = '';

      let tot24 = 0, tot25 = 0, tot26 = 0;

      comp.monthly_comparison_8m.forEach(m => {{
        tot24 += m.sales_2024;
        tot25 += m.sales_2025;
        tot26 += m.sales_2026;

        const g26_25 = ((m.sales_2026 - m.sales_2025) / m.sales_2025) * 100;
        const g26_24 = ((m.sales_2026 - m.sales_2024) / m.sales_2024) * 100;

        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition-colors';
        tr.innerHTML = `
          <td class="py-2.5 px-3 font-semibold text-slate-900">${{m.name}}</td>
          <td class="py-2.5 px-3 text-right font-mono text-slate-600">${{fmtIDR(m.sales_2024)}}</td>
          <td class="py-2.5 px-3 text-right font-mono text-sky-800">${{fmtIDR(m.sales_2025)}}</td>
          <td class="py-2.5 px-3 text-right font-mono font-bold text-emerald-700">${{fmtIDR(m.sales_2026)}}</td>
          <td class="py-2.5 px-3 text-right font-mono font-bold ${{g26_25 >= 0 ? 'text-emerald-600' : 'text-rose-600'}}">
            ${{g26_25 >= 0 ? '+' : ''}}${{g26_25.toFixed(1)}}%
          </td>
          <td class="py-2.5 px-3 text-right font-mono font-bold text-indigo-700">
            +${{g26_24.toFixed(1)}}%
          </td>
        `;
        tbody.appendChild(tr);
      }});

      // Summary row
      const gTot26_25 = ((tot26 - tot25) / tot25) * 100;
      const gTot26_24 = ((tot26 - tot24) / tot24) * 100;

      const trFoot = document.createElement('tr');
      trFoot.className = 'bg-slate-100 font-bold border-t-2 border-slate-300 text-slate-900';
      trFoot.innerHTML = `
        <td class="py-3 px-3">TOTAL 8 BULAN (JAN - AGU)</td>
        <td class="py-3 px-3 text-right font-mono">${{fmtIDR(tot24)}}</td>
        <td class="py-3 px-3 text-right font-mono text-sky-900">${{fmtIDR(tot25)}}</td>
        <td class="py-3 px-3 text-right font-mono text-emerald-800 text-sm">${{fmtIDR(tot26)}}</td>
        <td class="py-3 px-3 text-right font-mono text-emerald-700">+${{gTot26_25.toFixed(2)}}%</td>
        <td class="py-3 px-3 text-right font-mono text-indigo-700">+${{gTot26_24.toFixed(2)}}%</td>
      `;
      tbody.appendChild(trFoot);

      // Render Multi-Year SVG Grouped Bar Chart
      renderMultiYearSvgChart(comp.monthly_comparison_8m);
    }}

    // Multi-Year SVG Grouped Bar Chart
    function renderMultiYearSvgChart(monthlyComp) {{
      const svg = document.getElementById('svg-multiyear-chart');
      svg.innerHTML = '';

      const chartLeft = 70;
      const chartRight = 835;
      const chartWidth = chartRight - chartLeft;
      const chartTop = 45;
      const chartBottom = 240;
      const chartHeight = chartBottom - chartTop;

      const maxScale = 190000000; // Rp 190 Jt

      const numGroups = 8;
      const groupWidth = chartWidth / numGroups;
      const singleBarW = 16;
      const groupGap = (groupWidth - (3 * singleBarW)) / 2;

      // Gridlines
      for (let s = 40000000; s < maxScale; s += 40000000) {{
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

      // Bars Grouped per month
      monthlyComp.forEach((m, idx) => {{
        const gX = chartLeft + idx * groupWidth + groupGap;

        // Bar 1: 2024
        const h24 = (m.sales_2024 / maxScale) * chartHeight;
        const r24 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        r24.setAttribute('x', gX);
        r24.setAttribute('y', chartBottom - h24);
        r24.setAttribute('width', singleBarW);
        r24.setAttribute('height', h24);
        r24.setAttribute('rx', '3');
        r24.setAttribute('fill', '#94a3b8');
        svg.appendChild(r24);

        // Bar 2: 2025
        const h25 = (m.sales_2025 / maxScale) * chartHeight;
        const r25 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        r25.setAttribute('x', gX + singleBarW + 2);
        r25.setAttribute('y', chartBottom - h25);
        r25.setAttribute('width', singleBarW);
        r25.setAttribute('height', h25);
        r25.setAttribute('rx', '3');
        r25.setAttribute('fill', '#0284c7');
        svg.appendChild(r25);

        // Bar 3: 2026
        const h26 = (m.sales_2026 / maxScale) * chartHeight;
        const r26 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        r26.setAttribute('x', gX + (singleBarW * 2) + 4);
        r26.setAttribute('y', chartBottom - h26);
        r26.setAttribute('width', singleBarW);
        r26.setAttribute('height', h26);
        r26.setAttribute('rx', '3');
        r26.setAttribute('fill', '#059669');
        svg.appendChild(r26);

        // Value text for 2026
        const txt26 = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt26.setAttribute('x', gX + (singleBarW * 2) + 4 + singleBarW / 2);
        txt26.setAttribute('y', chartBottom - h26 - 6);
        txt26.setAttribute('text-anchor', 'middle');
        txt26.setAttribute('font-size', '8.5');
        txt26.setAttribute('font-weight', '800');
        txt26.setAttribute('fill', '#065f46');
        txt26.textContent = (m.sales_2026 / 1e6).toFixed(0);
        svg.appendChild(txt26);

        // Month Label
        const txtM = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txtM.setAttribute('x', gX + singleBarW * 1.5 + 2);
        txtM.setAttribute('y', chartBottom + 18);
        txtM.setAttribute('text-anchor', 'middle');
        txtM.setAttribute('font-size', '11');
        txtM.setAttribute('font-weight', '700');
        txtM.setAttribute('fill', '#334155');
        txtM.textContent = m.code;
        svg.appendChild(txtM);
      }});
    }}

    // Product Modal for 2024
    function openModal(prod) {{
      document.getElementById('modal-title').innerText = prod.name;
      document.getElementById('modal-category').innerText = prod.category;
      document.getElementById('modal-sales').innerText = fmtIDR(prod.period_sales || prod.total_sales);
      document.getElementById('modal-qty').innerText = fmtNum(prod.period_qty || prod.total_qty) + ' ' + prod.unit;
      document.getElementById('modal-pct').innerText = prod.pct.toFixed(2) + '%';

      const tbody = document.getElementById('modal-month-tbody');
      tbody.innerHTML = '';
      MONTH_KEYS_12.forEach(m => {{
        const s = prod.monthly_sales[m] || 0;
        const q = prod.monthly_qty[m] || 0;
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td class="py-1 px-3">${{MONTH_NAMES[m]}}</td>
          <td class="py-1 px-3 text-right">${{fmtNum(q)}}</td>
          <td class="py-1 px-3 text-right font-bold text-slate-900">${{fmtIDR(s)}}</td>
        `;
        tbody.appendChild(tr);
      }});

      // Spark chart
      const svg = document.getElementById('modal-spark-chart');
      svg.innerHTML = '';
      const vals = MONTH_KEYS_12.map(m => prod.monthly_sales[m] || 0);
      const maxVal = Math.max(...vals, 1);
      const w = 540;
      const h = 100;
      const stepX = w / 11;

      let d = '';
      vals.forEach((v, i) => {{
        const x = i * stepX;
        const y = h - (v / maxVal) * 80 - 10;
        d += (i === 0 ? `M ${{x}} ${{y}}` : ` L ${{x}} ${{y}}`);

        const c = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        c.setAttribute('cx', x);
        c.setAttribute('cy', y);
        c.setAttribute('r', '3');
        c.setAttribute('fill', '#0284c7');
        svg.appendChild(c);
      }});

      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      path.setAttribute('d', d);
      path.setAttribute('fill', 'none');
      path.setAttribute('stroke', '#0284c7');
      path.setAttribute('stroke-width', '2');
      svg.appendChild(path);

      document.getElementById('product-modal').classList.remove('hidden');
    }}

    function closeModal() {{
      document.getElementById('product-modal').classList.add('hidden');
    }}

    // Export Excel Multi-Sheet
    function exportUnifiedExcel() {{
      if (typeof XLSX === 'undefined') {{
        alert('Library Excel belum selesai dimuat. Harap periksa koneksi internet Anda.');
        return;
      }}

      const wb = XLSX.utils.book_new();

      // Sheet 1: Komparasi 3 Tahun (8 Bulan)
      const compRows = [
        ['Bulan', 'Omzet 2024 (Rp)', 'Omzet 2025 (Rp)', 'Omzet 2026 (Rp)', 'Pertumbuhan 2026 vs 2025 (%)', 'Pertumbuhan 2026 vs 2024 (%)']
      ];
      MASTER_DATA.comparison_multiyear.monthly_comparison_8m.forEach(m => {{
        const g25 = ((m.sales_2026 - m.sales_2025) / m.sales_2025) * 100;
        const g24 = ((m.sales_2026 - m.sales_2024) / m.sales_2024) * 100;
        compRows.push([m.name, m.sales_2024, m.sales_2025, m.sales_2026, g25 / 100, g24 / 100]);
      }});
      compRows.push([
        'TOTAL (8 BULAN)',
        MASTER_DATA.comparison_multiyear.sales_8m_2024,
        MASTER_DATA.comparison_multiyear.sales_8m_2025,
        MASTER_DATA.comparison_multiyear.sales_8m_2026,
        MASTER_DATA.comparison_multiyear.growth_26_vs_25_8m_pct / 100,
        MASTER_DATA.comparison_multiyear.growth_26_vs_24_8m_pct / 100
      ]);
      const wsComp = XLSX.utils.aoa_to_sheet(compRows);
      XLSX.utils.book_append_sheet(wb, wsComp, 'Komparasi 3 Tahun (8M)');

      // Sheet 2: 2026 Matrix
      const mat26Rows = [
        ['Bulan', 'Large Format', 'Print A3+ & Sheet', 'Doc & Cetak IR', 'Desain & Cetak Khusus', 'Merchandise & Promosi', 'Finishing & Jilid', 'Total Omzet (Rp)']
      ];
      MASTER_DATA.data_2026.monthly.forEach(m => {{
        mat26Rows.push([
          m.name,
          m.matrix.large_format,
          m.matrix.print_a3,
          m.matrix.doc_fc,
          m.matrix.desain_khusus,
          m.matrix.merchandise,
          m.matrix.finishing,
          m.sales
        ]);
      }});
      mat26Rows.push(['TOTAL 2026 (8 BLN)', 449325263, 389124966, 41288500, 62290500, 75637432, 92641542, 1180308204]);
      const ws26 = XLSX.utils.aoa_to_sheet(mat26Rows);
      XLSX.utils.book_append_sheet(wb, ws26, 'Matriks Penjualan 2026');

      // Sheet 3: 2025 Matrix
      const mat25Rows = [
        ['Bulan', 'Large Format', 'Print A3+ & Sheet', 'Doc & Fotocopy', 'Finishing', 'Merchandise', 'Total Omzet (Rp)']
      ];
      MASTER_DATA.data_2025.monthly.forEach(m => {{
        mat25Rows.push([
          m.name,
          m.matrix.large_format,
          m.matrix.print_a3,
          m.matrix.doc_fc,
          m.matrix.finishing,
          m.matrix.merchandise,
          m.sales
        ]);
      }});
      mat25Rows.push(['TOTAL 2025', 509000000, 488600000, 118900000, 64801929, 157600000, 1378951929]);
      const ws25 = XLSX.utils.aoa_to_sheet(mat25Rows);
      XLSX.utils.book_append_sheet(wb, ws25, 'Matriks Penjualan 2025');

      // Sheet 4: 2024 Products
      const p24Rows = [
        ['Rank', 'Nama Barang', 'Kategori', 'Satuan', 'Total Volume', 'Total Omzet (Rp)', 'Porsi (%)']
      ];
      MASTER_DATA.data_2024.products.forEach(p => {{
        p24Rows.push([p.rank, p.name, p.category, p.unit, p.total_qty, p.total_sales, p.pct / 100]);
      }});
      const ws24 = XLSX.utils.aoa_to_sheet(p24Rows);
      XLSX.utils.book_append_sheet(wb, ws24, '129 Produk 2024');

      XLSX.writeFile(wb, 'Laporan_Penjualan_MultiTahun_Snaprint_2024_2026.xlsx');
    }}

    // Initial Load - Default to 2026 as user requested
    window.addEventListener('DOMContentLoaded', () => {{
      selectYear('2026');
    }});
  </script>
</body>
</html>
"""

# Write workspace index.html
with open("index.html", "w", encoding="utf-8") as f:
    f.write(unified_html)
print("Updated index.html to 3-Year Unified Multi-Year Dynamic Dashboard!")

# Write 2025/index.html (with relative PDF and SQL paths)
html_2025 = unified_html.replace('href="Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"', 'href="../Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"')
html_2025 = html_2025.replace('href="Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"', 'href="Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"')
html_2025 = html_2025.replace('href="Laporan_Penjualan_2026_Snaprint_Digital.pdf"', 'href="../2026/Laporan_Penjualan_2026_Snaprint_Digital.pdf"')
html_2025 = html_2025.replace('href="snaprint_sales.sql"', 'href="../snaprint_sales.sql"')
with open("2025/index.html", "w", encoding="utf-8") as f:
    f.write(html_2025)
print("Created 2025/index.html!")

# Write 2026/index.html (with relative PDF and SQL paths)
html_2026 = unified_html.replace('href="Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"', 'href="../Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"')
html_2026 = html_2026.replace('href="Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"', 'href="../2025/Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"')
html_2026 = html_2026.replace('href="Laporan_Penjualan_2026_Snaprint_Digital.pdf"', 'href="Laporan_Penjualan_2026_Snaprint_Digital.pdf"')
html_2026 = html_2026.replace('href="snaprint_sales.sql"', 'href="../snaprint_sales.sql"')
with open("2026/index.html", "w", encoding="utf-8") as f:
    f.write(html_2026)
print("Created 2026/index.html!")

# Write artifact
artifact_path = r"C:\Users\Swanto\.gemini\antigravity\brain\186463fa-99d6-481c-8f3f-6ce410e3a005\dashboard_snaprint_unified.html"
with open(artifact_path, "w", encoding="utf-8") as f:
    f.write(unified_html)
print(f"Generated artifact: {artifact_path}")
