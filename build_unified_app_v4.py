import json
import os
import re

# Load Master Data
with open("sales_data_master.json", "r", encoding="utf-8") as f:
    master_data = json.load(f)

json_master_str = json.dumps(master_data, ensure_ascii=False)

# Build unified HTML with complete Data Change & Save features
unified_html = f"""<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SNAPRINT DIGITAL — Sistem Dashboard Penjualan &amp; Laba Rugi Terintegrasi (2024 - 2026)</title>
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

    @keyframes pulse-ring {{
      0% {{ transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }}
      70% {{ transform: scale(1); box-shadow: 0 0 0 10px rgba(16, 185, 129, 0); }}
      100% {{ transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }}
    }}

    .btn-save-dirty {{
      animation: pulse-ring 2s infinite;
      background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
      color: white !important;
      cursor: pointer !important;
      border: 2px solid #34d399 !important;
    }}
  </style>
</head>
<body class="bg-[var(--background,#f8fafc)] text-[var(--foreground,#0f172a)] antialiased min-h-screen p-3 sm:p-5 text-xs sm:text-sm">

  <!-- TOAST NOTIFICATION -->
  <div id="toast-notification" class="fixed top-5 right-5 z-50 bg-emerald-600 text-white px-4 py-3 rounded-xl shadow-2xl flex items-center space-x-2.5 text-xs font-bold hidden transition-all duration-300 transform translate-y-[-20px] opacity-0">
    <span class="text-base">✅</span>
    <span id="toast-message">Perubahan data berhasil disimpan!</span>
  </div>

  <!-- FLOATING UNSAVED CHANGES ACTION BAR -->
  <div id="floating-unsaved-bar" class="fixed bottom-5 left-1/2 -translate-x-1/2 z-40 bg-slate-900/95 backdrop-blur-md text-white px-5 py-3 rounded-2xl shadow-2xl border-2 border-amber-400 flex items-center justify-between gap-4 max-w-xl w-[92%] sm:w-auto hidden transition-all duration-300">
    <div class="flex items-center space-x-2.5 shrink-0">
      <span class="text-xl animate-bounce">⚠️</span>
      <div>
        <p class="font-extrabold text-xs text-amber-300">Perubahan Data Terdeteksi!</p>
        <p class="text-[10px] text-slate-300">Ada modifikasi data pada tahun <span id="unsaved-year-label" class="font-bold text-white">2026</span> yang belum disimpan.</p>
      </div>
    </div>
    <div class="flex items-center space-x-2 shrink-0">
      <button onclick="saveDataChanges()" class="px-3.5 py-2 text-xs font-black rounded-xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 shadow-md flex items-center space-x-1.5 transition-transform hover:scale-105">
        <span>💾</span>
        <span>Simpan Sekarang</span>
      </button>
      <button onclick="discardChanges()" class="px-2.5 py-2 text-xs font-bold rounded-xl text-slate-400 hover:text-white">
        Batal
      </button>
    </div>
  </div>

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
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-indigo-500/10 text-indigo-700 border border-indigo-500/20">SISTEM PENJUALAN &amp; LABA BERSIH TERPADU</span>
            </div>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Konsolidasi Laporan Penjualan &amp; Laporan Laba/Rugi (P&amp;L) ACCURATE Accounting System (2024, 2025, 2026)
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
            📅 2024 (Laba: Rp 322,8 Jt)
          </button>
          <button onclick="selectYear('2025')" id="year-btn-2025" class="px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            🚀 2025 (Laba: Rp 572,3 Jt)
          </button>
          <button onclick="selectYear('2026')" id="year-btn-2026" class="px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            🔥 2026 (Laba: Rp 383,3 Jt - 8 Bln)
          </button>
          <button onclick="selectYear('MULTI')" id="year-btn-MULTI" class="px-3 py-1.5 sm:px-4 sm:py-2 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
            📈 Komparasi Multi-Tahun &amp; Laba Rugi
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

      <!-- DATA MANAGEMENT & SIMPAN PERUBAHAN BAR (IN EVERY FOLDER) -->
      <div id="data-management-bar" class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-3.5 sm:p-4 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div class="flex items-center space-x-3">
          <div id="sync-status-indicator" class="w-3.5 h-3.5 rounded-full bg-emerald-500 shadow-sm shrink-0"></div>
          <div>
            <div class="flex items-center space-x-2 flex-wrap">
              <span id="sync-status-text" class="text-xs font-bold text-slate-800">Status Data: Standar ACCURATE Tersimpan</span>
              <span id="badge-unsaved-changes" class="hidden text-[10px] font-black uppercase px-2 py-0.5 rounded-full bg-amber-500 text-white animate-pulse shadow-xs">
                ⚠️ Ada Perubahan Belum Disimpan
              </span>
              <span id="badge-custom-saved" class="hidden text-[10px] font-bold px-2 py-0.5 rounded-full bg-sky-100 text-sky-800 border border-sky-200">
                💾 Data Kustom Aktif
              </span>
            </div>
            <p class="text-[11px] text-slate-500 mt-0.5" id="sync-status-desc">
              Data sinkron dengan pembukuan resmi ACCURATE. Klik "Sesuaikan Data" jika ada perubahan data omzet atau beban.
            </p>
          </div>
        </div>

        <!-- Action Buttons: Sesuaikan, Simpan Perubahan, Reset, Unduh JSON -->
        <div class="flex items-center space-x-2 flex-wrap gap-y-2">
          
          <!-- Tombol Buka Editor -->
          <button onclick="openDataEditorModal()" class="px-3 py-1.5 sm:px-3.5 sm:py-2 text-xs font-bold rounded-xl bg-sky-50 text-sky-700 hover:bg-sky-100 border border-sky-200 transition-all flex items-center space-x-1.5 shadow-2xs">
            <span>✏️</span>
            <span id="btn-edit-text">Sesuaikan Data (2026)</span>
          </button>
          
          <!-- TOMBOL SIMPAN PERUBAHAN UTAMA -->
          <button id="btn-save-master" onclick="saveDataChanges()" class="px-3.5 py-1.5 sm:px-4 sm:py-2 text-xs font-bold rounded-xl bg-slate-200 text-slate-400 transition-all shadow-xs flex items-center space-x-1.5 cursor-not-allowed" disabled>
            <span>💾</span>
            <span id="btn-save-label">Simpan Perubahan</span>
          </button>

          <!-- Tombol Reset ke Default -->
          <button onclick="resetDataChanges()" id="btn-reset-master" class="px-2.5 py-1.5 sm:px-3 sm:py-2 text-xs font-medium rounded-xl text-slate-500 hover:text-rose-600 hover:bg-rose-50 border border-transparent hover:border-rose-200 transition-all flex items-center space-x-1">
            <span>🔄</span>
            <span>Reset Data</span>
          </button>

          <!-- Tombol Download JSON Folder -->
          <button onclick="downloadFolderJson()" class="px-2.5 py-1.5 sm:px-3 sm:py-2 text-xs font-semibold rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 transition-all flex items-center space-x-1">
            <span>📥</span>
            <span id="btn-dl-label">Unduh JSON</span>
          </button>

        </div>
      </div>

      <!-- FINANCIAL PROFITABILITY HIGHLIGHT RIBBON (LABA BERSIH) -->
      <div class="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 rounded-2xl p-4 sm:p-5 text-white shadow-md space-y-3">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-2.5 border-b border-white/10">
          <div class="flex items-center space-x-2">
            <span class="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
              RINGKASAN LABA / RUGI RESMI (ACCURATE)
            </span>
            <span class="text-xs text-slate-300" id="pl-ribbon-period">Tahun Buku</span>
          </div>
          <div class="text-xs text-slate-300 flex items-center space-x-2">
            <span>Standar Pembukuan: <b>Akuntansi Akrual ACCURATE</b></span>
          </div>
        </div>

        <!-- 4 Key Financial P&L Metrics -->
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          
          <!-- P&L 1: Pendapatan Bersih -->
          <div class="bg-white/5 rounded-xl p-3 border border-white/10">
            <span class="text-[11px] uppercase tracking-wider text-slate-400 font-semibold block">Pendapatan Bersih</span>
            <span class="text-lg sm:text-xl font-bold font-mono text-white block mt-0.5" id="pl-net-revenue">Rp 0</span>
            <span class="text-[10px] text-slate-400 block mt-0.5">Penjualan setelah retur &amp; diskon</span>
          </div>

          <!-- P&L 2: Beban Pokok (HPP) -->
          <div class="bg-white/5 rounded-xl p-3 border border-white/10">
            <span class="text-[11px] uppercase tracking-wider text-slate-400 font-semibold block">Beban Pokok Penjualan (HPP)</span>
            <span class="text-lg sm:text-xl font-bold font-mono text-rose-300 block mt-0.5" id="pl-cogs">Rp 0</span>
            <span class="text-[10px] text-slate-400 block mt-0.5" id="pl-cogs-pct">0% dari Pendapatan</span>
          </div>

          <!-- P&L 3: Beban Operasional -->
          <div class="bg-white/5 rounded-xl p-3 border border-white/10">
            <span class="text-[11px] uppercase tracking-wider text-slate-400 font-semibold block">Beban Operasional (Opex)</span>
            <span class="text-lg sm:text-xl font-bold font-mono text-amber-300 block mt-0.5" id="pl-opex">Rp 0</span>
            <span class="text-[10px] text-slate-400 block mt-0.5" id="pl-opex-pct">0% dari Pendapatan</span>
          </div>

          <!-- P&L 4: LABA BERSIH (NET PROFIT) -->
          <div class="bg-emerald-500/20 rounded-xl p-3 border border-emerald-500/40 relative overflow-hidden">
            <div class="absolute right-2 top-2 text-xl opacity-30">👑</div>
            <span class="text-[11px] uppercase tracking-wider text-emerald-200 font-extrabold block">LABA BERSIH (NET PROFIT)</span>
            <span class="text-xl sm:text-2xl font-black font-mono text-emerald-300 block mt-0.5" id="pl-net-profit">Rp 0</span>
            <div class="flex items-center justify-between text-[11px] text-emerald-200 font-bold mt-0.5">
              <span id="pl-net-margin">Net Margin: 0%</span>
              <span id="pl-avg-monthly-profit">Rerata: Rp 0/bln</span>
            </div>
          </div>

        </div>
      </div>

      <!-- DYNAMIC FILTER CONTROL PANEL -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-3 pb-3 border-b border-[var(--border,#f1f5f9)]">
          <div class="flex items-center space-x-2">
            <span class="text-base">🎛️</span>
            <div>
              <h2 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)]" id="filter-header-title">
                Filter Waktu &amp; Navigasi Kinerja
              </h2>
              <p class="text-xs text-[var(--muted-foreground,#64748b)]">
                Pilih periode bulan, semester, atau filter kuartal untuk memutakhirkan visualisasi seketika.
              </p>
            </div>
          </div>
          
          <div class="flex items-center space-x-2 self-start md:self-auto">
            <span class="text-xs text-[var(--muted-foreground,#64748b)] font-medium">Periode Aktif:</span>
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

      <!-- CHARTS SECTION: MONTHLY SALES TREND (2 Cols) & CATEGORY BREAKDOWN (1 Col) -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-3.5">
        
        <!-- Chart 1: Monthly Trend Bar Chart (2 Cols) -->
        <div class="lg:col-span-2 bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div>
              <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
                <span>📊</span>
                <span id="chart-title">Tren Penjualan Bulanan</span>
              </h3>
              <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
                Setiap batang menunjukkan angka penjualan bersih (Rupiah) &amp; porsi persentase. Klik batang untuk filter!
              </p>
            </div>
            <span class="text-xs px-2.5 py-1 rounded-full bg-sky-50 text-sky-700 font-bold border border-sky-200 shrink-0" id="chart-year-badge">
              Bulan Buku Aktif
            </span>
          </div>

          <!-- SVG Dynamic Bar Chart -->
          <div class="w-full overflow-x-auto py-2">
            <svg id="svg-monthly-chart" viewBox="0 0 680 260" class="w-full h-56 sm:h-64 select-none"></svg>
          </div>
        </div>

        <!-- Chart 2: Category Contribution Ranking (1 Col) -->
        <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm flex flex-col justify-between">
          <div class="space-y-3">
            <div class="flex items-center justify-between">
              <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
                <span>🎯</span>
                <span>Pangsa Kategori Produk</span>
              </h3>
              <span class="text-[11px] text-[var(--muted-foreground,#64748b)]">Kontribusi</span>
            </div>
            <p class="text-xs text-[var(--muted-foreground,#64748b)]">
              Distribusi pendapatan menurut kelompok kategori cetak utama.
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

      <!-- VIEW TOGGLE: DETAIL PENJUALAN VS LAPORAN LABA / RUGI -->
      <div class="flex items-center space-x-2 border-b border-slate-200 pb-2">
        <button onclick="switchSubTab('SALES')" id="subtab-btn-SALES" class="px-4 py-2 font-bold text-xs rounded-xl bg-slate-900 text-white shadow-sm flex items-center space-x-1.5">
          <span>📋</span>
          <span>Rincian Penjualan &amp; Produk</span>
        </button>
        <button onclick="switchSubTab('PL')" id="subtab-btn-PL" class="px-4 py-2 font-bold text-xs rounded-xl text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-all flex items-center space-x-1.5">
          <span>📑</span>
          <span>Laporan Laba / Rugi (ACCURATE P&amp;L)</span>
        </button>
      </div>

      <!-- SUB-PANEL 1: SALES & PRODUCTS TABLE -->
      <div id="subpanel-sales" class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
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

      <!-- SUB-PANEL 2: PROFIT & LOSS STATEMENT (P&L) -->
      <div id="subpanel-pl" class="hidden bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-4">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-3 border-b border-[var(--border,#f1f5f9)] gap-2">
          <div>
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>📑</span>
              <span id="pl-table-title">Laporan Laba / Rugi (Standar)</span>
            </h3>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Format standar akuntansi ACCURATE Accounting System.
            </p>
          </div>
          <div class="flex items-center space-x-2">
            <button onclick="openDataEditorModal()" class="px-2.5 py-1 text-xs font-bold rounded-lg bg-emerald-50 text-emerald-700 border border-emerald-300 hover:bg-emerald-100 flex items-center space-x-1">
              <span>✏️</span>
              <span>Edit Angka P&amp;L</span>
            </button>
            <span class="text-xs font-bold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200">
              Audit Ready
            </span>
          </div>
        </div>

        <!-- P&L Waterfall Table Container -->
        <div id="pl-statement-container"></div>
      </div>

    </div>

    <!-- ======================================================== -->
    <!-- VIEW CONTAINER 2: MULTI-YEAR COMPARISON & P&L TREND -->
    <!-- ======================================================== -->
    <div id="view-multi-year" class="hidden space-y-4 sm:space-y-6">

      <!-- MULTI-YEAR EXECUTIVE OVERVIEW -->
      <div class="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-2xl p-5 sm:p-6 shadow-md space-y-4">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div class="flex items-center space-x-2">
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">LOMPATAN BISNIS &amp; LABA BERSIH</span>
              <span class="text-xs text-slate-300">Konsolidasi Finansial 3 Tahun (2024, 2025, 2026)</span>
            </div>
            <h2 class="text-lg sm:text-2xl font-black tracking-tight mt-1">
              Analisis Komparasi Kinerja Penjualan &amp; Laba Bersih
            </h2>
            <p class="text-xs sm:text-sm text-slate-300 max-w-3xl mt-1">
              Evaluasi komparatif perkembangan bisnis Snaprint Digital Printing secara konsisten menunjukkan percepatan omzet dan profitabilitas yang kokoh: Laba bersih melompat dari Rp 322,8 Jt (2024) ke Rp 572,3 Jt (2025: +77,3%), dan dalam 8 bulan 2026 telah menembus <b>Rp 383,3 Jt (melampaui laba setahun penuh 2024)</b>!
            </p>
          </div>

          <div class="shrink-0 bg-white/10 backdrop-blur-md p-3.5 rounded-xl border border-white/20 text-center">
            <span class="text-[11px] uppercase tracking-wider text-emerald-300 font-semibold block">Pertumbuhan Omzet 2 Tahun (8M)</span>
            <span class="text-2xl sm:text-3xl font-extrabold text-white font-mono block mt-0.5">+132,22%</span>
            <span class="text-[10px] text-slate-300 block">+Rp 672,04 Juta Omzet</span>
          </div>
        </div>

        <!-- 3-Year Metric Mini-Cards (Revenue & Net Profit) -->
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-2 border-t border-white/10">
          <div class="bg-white/5 rounded-xl p-3 border border-white/10">
            <span class="text-xs text-slate-400 block font-medium">Tahun Buku 2024 (12 Bulan):</span>
            <span class="text-base sm:text-lg font-bold font-mono text-white block mt-0.5">Omzet: Rp 930.847.260</span>
            <span class="text-[11px] text-emerald-400 block font-semibold">Laba Bersih: Rp 322.842.690 (34,68%)</span>
          </div>
          <div class="bg-white/5 rounded-xl p-3 border border-white/10">
            <span class="text-xs text-slate-400 block font-medium">Tahun Buku 2025 (12 Bulan):</span>
            <span class="text-base sm:text-lg font-bold font-mono text-white block mt-0.5">Omzet: Rp 1.378.951.929</span>
            <span class="text-[11px] text-emerald-400 block font-semibold">Laba Bersih: Rp 572.331.685 (41,50%) 🚀</span>
          </div>
          <div class="bg-emerald-500/20 rounded-xl p-3 border border-emerald-500/40">
            <span class="text-xs text-emerald-200 block font-medium">Tahun Buku 2026 (8 Bulan Berjalan):</span>
            <span class="text-base sm:text-lg font-black font-mono text-emerald-300 block mt-0.5">Omzet: Rp 1.180.308.204</span>
            <span class="text-[11px] text-emerald-200 block font-semibold">Laba Bersih: Rp 383.265.304 (33,32%) 🔥</span>
          </div>
        </div>
      </div>

      <!-- MULTI-YEAR COMPREHENSIVE PROFIT & LOSS COMPARISON TABLE -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-3 border-b border-[var(--border,#f1f5f9)] gap-2">
          <div>
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>📑</span>
              <span>Laporan Laba / Rugi Komparatif 3 Tahun (ACCURATE Financial Statement)</span>
            </h3>
            <p class="text-xs text-[var(--muted-foreground,#64748b)] mt-0.5">
              Perbandingan pendapatan bersih, HPP, beban operasional, dan pertumbuhan laba bersih tahun ke tahun.
            </p>
          </div>
          <span class="text-xs font-bold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200">
            ACCURATE Standard P&amp;L
          </span>
        </div>

        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
              <tr>
                <th class="py-2.5 px-3">Komponen Laba / Rugi</th>
                <th class="py-2.5 px-3 text-right">Tahun 2024 (12 Bln)</th>
                <th class="py-2.5 px-3 text-right">Tahun 2025 (12 Bln)</th>
                <th class="py-2.5 px-3 text-right">Tahun 2026 (8 Bln s/d Agu)</th>
                <th class="py-2.5 px-3 text-right">Pertumbuhan 2025 vs 2024</th>
                <th class="py-2.5 px-3 text-right">Evaluasi Tren</th>
              </tr>
            </thead>
            <tbody id="multiyear-pl-tbody" class="divide-y divide-slate-100 font-mono text-slate-800"></tbody>
          </table>
        </div>
      </div>

      <!-- MULTI-YEAR COMPARISON BAR CHART (SALES) -->
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

      <!-- DETAILED 3-YEAR SALES COMPARISON TABLE -->
      <div class="bg-[var(--card,#ffffff)] border border-[var(--border,#e2e8f0)] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-3 border-b border-[var(--border,#f1f5f9)] gap-2">
          <div>
            <h3 class="font-bold text-sm sm:text-base text-[var(--foreground,#0f172a)] flex items-center space-x-2">
              <span>📋</span>
              <span>Tabel Rincian Omzet Bulanan Apel-ke-Apel (Januari – Agustus)</span>
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

  <!-- ======================================================== -->
  <!-- MODAL 1: DATA EDITOR & PERUBAHAN DATA (P&L & OMZET)      -->
  <!-- ======================================================== -->
  <div id="data-editor-modal" class="fixed inset-0 bg-slate-950/70 backdrop-blur-xs z-50 flex items-center justify-center p-3 hidden">
    <div class="bg-white rounded-2xl max-w-3xl w-full border border-slate-200 shadow-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in fade-in zoom-in-95 duration-200">
      
      <!-- Modal Header -->
      <div class="p-4 bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white flex items-center justify-between shrink-0">
        <div class="flex items-center space-x-2.5">
          <div class="w-8 h-8 rounded-lg bg-emerald-500/20 border border-emerald-400/30 flex items-center justify-center text-emerald-300 font-bold text-sm">
            ✏️
          </div>
          <div>
            <div class="flex items-center space-x-2">
              <h3 class="text-sm sm:text-base font-bold text-white">Editor Data Keuangan &amp; Laba Rugi</h3>
              <span id="modal-year-tag" class="px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-amber-500 text-slate-950">2026</span>
            </div>
            <p class="text-[11px] text-slate-300 mt-0.5">Ubah nilai angka penjualan atau beban operasional. Perhitungan laba bersih akan disesuaikan seketika.</p>
          </div>
        </div>
        <button onclick="closeDataEditorModal()" class="w-7 h-7 rounded-lg bg-white/10 hover:bg-white/20 flex items-center justify-center text-white text-sm">✕</button>
      </div>

      <!-- Modal Tabs (Laba Rugi vs Penjualan Bulanan) -->
      <div class="flex items-center space-x-1 bg-slate-100 px-4 py-2 border-b border-slate-200 shrink-0 text-xs">
        <button onclick="switchEditorTab('PL')" id="tab-edit-pl" class="px-3 py-1.5 font-bold rounded-lg bg-white text-slate-900 shadow-xs">
          📑 Formulir Laba / Rugi (P&amp;L)
        </button>
        <button onclick="switchEditorTab('SALES')" id="tab-edit-sales" class="px-3 py-1.5 font-medium rounded-lg text-slate-600 hover:text-slate-900">
          📊 Formulir Penjualan Bulanan
        </button>
      </div>

      <!-- Modal Body (Scrollable) -->
      <div class="p-4 sm:p-5 overflow-y-auto space-y-4 flex-1">
        
        <!-- Tab 1: P&L Accounts Form -->
        <div id="editor-panel-pl" class="space-y-4">
          
          <!-- Live Preview Banner -->
          <div class="bg-gradient-to-r from-emerald-950 via-slate-900 to-emerald-950 p-3.5 rounded-xl border border-emerald-500/30 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <span class="text-[10px] uppercase font-bold text-emerald-400 tracking-wider">Hasil Simulasi Laba Bersih</span>
              <p class="text-lg sm:text-xl font-black font-mono text-emerald-300 mt-0.5" id="sim-net-profit">Rp 0</p>
            </div>
            <div class="flex items-center space-x-3 text-xs">
              <div class="bg-white/5 px-2.5 py-1 rounded-lg border border-white/10">
                <span class="text-slate-400 block text-[10px]">Net Margin</span>
                <span class="font-bold text-emerald-300 font-mono" id="sim-net-margin">0%</span>
              </div>
              <div class="bg-white/5 px-2.5 py-1 rounded-lg border border-white/10">
                <span class="text-slate-400 block text-[10px]">Rerata / Bulan</span>
                <span class="font-bold text-emerald-300 font-mono" id="sim-avg-profit">Rp 0</span>
              </div>
            </div>
          </div>

          <!-- Section 1: Pendapatan -->
          <div class="border border-slate-200 rounded-xl p-3.5 space-y-3 bg-slate-50/50">
            <h4 class="font-bold text-xs text-slate-800 flex items-center justify-between">
              <span>1. Pendapatan Penjualan (Revenue)</span>
              <span class="font-mono text-emerald-700" id="sub-net-revenue">Net: Rp 0</span>
            </h4>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label class="block text-[11px] font-semibold text-slate-600 mb-1">Penjualan Kotor (Gross):</label>
                <input type="number" step="1000" id="inp-gross-sales" oninput="onEditorInputChange()" class="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-mono font-bold text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none">
              </div>
              <div>
                <label class="block text-[11px] font-semibold text-slate-600 mb-1">Diskon Penjualan (Rp):</label>
                <input type="number" step="1000" id="inp-sales-discount" oninput="onEditorInputChange()" class="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-mono text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none">
              </div>
              <div>
                <label class="block text-[11px] font-semibold text-slate-600 mb-1">Retur Penjualan (Rp):</label>
                <input type="number" step="1000" id="inp-sales-return" oninput="onEditorInputChange()" class="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-mono text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none">
              </div>
            </div>
          </div>

          <!-- Section 2: HPP & Laba Kotor -->
          <div class="border border-slate-200 rounded-xl p-3.5 space-y-3 bg-slate-50/50">
            <h4 class="font-bold text-xs text-slate-800 flex items-center justify-between">
              <span>2. Beban Pokok Penjualan (HPP / COGS)</span>
              <span class="font-mono text-slate-700" id="sub-gross-profit">Laba Kotor: Rp 0</span>
            </h4>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block text-[11px] font-semibold text-slate-600 mb-1">Total Beban Pokok (HPP Bahan &amp; Tinta):</label>
                <input type="number" step="1000" id="inp-cogs" oninput="onEditorInputChange()" class="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-mono font-bold text-rose-700 focus:ring-2 focus:ring-emerald-500 focus:outline-none">
              </div>
              <div class="flex items-center text-xs text-slate-600 pt-5">
                <span>Rasio HPP terhadap omzet: <b id="sub-cogs-pct" class="font-mono text-slate-900">0%</b></span>
              </div>
            </div>
          </div>

          <!-- Section 3: Beban Operasional Breakdown -->
          <div class="border border-slate-200 rounded-xl p-3.5 space-y-3 bg-slate-50/50">
            <h4 class="font-bold text-xs text-slate-800 flex items-center justify-between">
              <span>3. Beban Operasional (Operating Expenses)</span>
              <span class="font-mono text-amber-700" id="sub-total-opex">Total Opex: Rp 0</span>
            </h4>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3" id="editor-opex-inputs-container">
              <!-- Dynamically populated opex items -->
            </div>
          </div>

          <!-- Section 4: Non-Operasional -->
          <div class="border border-slate-200 rounded-xl p-3.5 space-y-3 bg-slate-50/50">
            <h4 class="font-bold text-xs text-slate-800">4. Pendapatan / (Beban) Non-Operasional</h4>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block text-[11px] font-semibold text-slate-600 mb-1">Pendapatan / (Beban) Lain-lain Bersih (Rp):</label>
                <input type="number" step="100" id="inp-other-income" oninput="onEditorInputChange()" class="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-mono text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none">
              </div>
              <div class="flex items-center text-xs text-slate-500 pt-5">
                <span>Contoh: bunga bank, biaya administrasi, selisih pembulatan.</span>
              </div>
            </div>
          </div>

        </div>

        <!-- Tab 2: Monthly Sales Form -->
        <div id="editor-panel-sales" class="hidden space-y-3">
          <p class="text-xs text-slate-600">Sesuaikan omzet dan kuantitas per bulan buku:</p>
          <div class="overflow-x-auto border border-slate-200 rounded-xl">
            <table class="w-full text-xs text-left">
              <thead class="bg-slate-100 font-semibold text-slate-700 uppercase text-[10px]">
                <tr>
                  <th class="py-2 px-3">Bulan</th>
                  <th class="py-2 px-3 text-right">Omzet Bersih (Rp)</th>
                  <th class="py-2 px-3 text-right">Volume / Order (Qty)</th>
                </tr>
              </thead>
              <tbody id="editor-monthly-tbody" class="divide-y divide-slate-100 font-mono">
                <!-- Dynamically generated rows -->
              </tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- Modal Footer -->
      <div class="p-4 bg-slate-50 border-t border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shrink-0">
        <div class="flex items-center space-x-2">
          <span class="w-2.5 h-2.5 rounded-full bg-slate-400" id="modal-status-dot"></span>
          <span class="text-xs font-medium text-slate-600" id="modal-change-status">Belum ada perubahan nilai.</span>
        </div>
        <div class="flex items-center space-x-2 justify-end">
          <button onclick="closeDataEditorModal()" class="px-3.5 py-2 text-xs font-semibold rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-700 transition-colors">
            Tutup
          </button>
          <button onclick="restoreInitialModalValues()" class="px-3 py-2 text-xs font-medium rounded-xl text-rose-600 hover:bg-rose-50 border border-transparent hover:border-rose-200 transition-colors">
            Kembalikan Nilai
          </button>
          
          <!-- TOMBOL SIMPAN DI MODAL -->
          <button id="modal-save-btn" onclick="saveDataChanges()" class="px-4 py-2 text-xs font-bold rounded-xl bg-slate-300 text-slate-500 cursor-not-allowed flex items-center space-x-1.5 transition-all shadow-sm" disabled>
            <span>💾</span>
            <span>Simpan Perubahan</span>
          </button>
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

  <!-- JAVASCRIPT APP LOGIC & DATA PERSISTENCE -->
  <script>
    // CERTIFIED ACCURATE BASE DATA (DEEP COPY BACKUP)
    const ORIGINAL_MASTER_DATA = {json_master_str};
    let MASTER_DATA = JSON.parse(JSON.stringify(ORIGINAL_MASTER_DATA));

    // LOCAL STORAGE PERSISTENCE CHECK
    const STORAGE_KEY = 'SNAPRINT_CUSTOM_FINANCIAL_DATA_V1';
    let savedCustomData = null;
    try {{
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {{
        savedCustomData = JSON.parse(stored);
        // Merge into active MASTER_DATA
        if (savedCustomData.data_2024) MASTER_DATA.data_2024 = savedCustomData.data_2024;
        if (savedCustomData.data_2025) MASTER_DATA.data_2025 = savedCustomData.data_2025;
        if (savedCustomData.data_2026) MASTER_DATA.data_2026 = savedCustomData.data_2026;
        if (savedCustomData.profit_loss_comparison) MASTER_DATA.profit_loss_comparison = savedCustomData.profit_loss_comparison;
      }}
    }} catch(e) {{
      console.warn('LocalStorage not available', e);
    }}

    // State
    const appState = {{
      activeYear: '2026', // '2024', '2025', '2026', or 'MULTI'
      activeSubTab: 'SALES', // 'SALES' or 'PL'
      activeEditorTab: 'PL', // 'PL' or 'SALES'
      period: 'ALL',      // 'ALL', 'H1', 'H2', 'Q1'..'Q4', or month code
      category: 'ALL',
      pareto: 'ALL',
      search: '',
      sortCol: 'sales',
      sortAsc: false,
      hasUnsavedChanges: false,
      initialEditorValues: {{}}
    }};

    const MONTH_KEYS_12 = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    const MONTH_KEYS_8 = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu'];
    const MONTH_NAMES = {{
      'Jan': 'Januari', 'Feb': 'Februari', 'Mar': 'Maret', 'Apr': 'April',
      'Mei': 'Mei', 'Jun': 'Juni', 'Jul': 'Juli', 'Agu': 'Agustus',
      'Sep': 'September', 'Okt': 'Oktober', 'Nov': 'November', 'Des': 'Desember'
    }};

    function fmtIDR(num) {{
      return 'Rp ' + Math.round(num || 0).toLocaleString('id-ID');
    }}

    function fmtNum(num) {{
      return Number(num || 0).toLocaleString('id-ID', {{ minimumFractionDigits: 0, maximumFractionDigits: 2 }});
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
        modeBadge.innerText = 'Tahun Buku 2024 (Laba: ' + fmtIDR(MASTER_DATA.data_2024.profit_loss.net_profit.amount) + ')';
        modeBadge.className = 'font-bold text-sky-700 bg-sky-50 px-3 py-1 rounded-lg border border-sky-200';
      }} else if (yr === '2025') {{
        modeBadge.innerText = 'Tahun Buku 2025 (Laba: ' + fmtIDR(MASTER_DATA.data_2025.profit_loss.net_profit.amount) + ')';
        modeBadge.className = 'font-bold text-emerald-700 bg-emerald-50 px-3 py-1 rounded-lg border border-emerald-200';
      }} else if (yr === '2026') {{
        modeBadge.innerText = 'Tahun Buku 2026 (Laba: ' + fmtIDR(MASTER_DATA.data_2026.profit_loss.net_profit.amount) + ' - 8 Bln)';
        modeBadge.className = 'font-bold text-amber-700 bg-amber-50 px-3 py-1 rounded-lg border border-amber-200';
      }} else {{
        modeBadge.innerText = 'Komparasi Multi-Tahun (2024 vs 2025 vs 2026)';
        modeBadge.className = 'font-bold text-indigo-700 bg-indigo-50 px-3 py-1 rounded-lg border border-indigo-200';
      }}

      // Update Action Bar Labels
      document.getElementById('btn-edit-text').innerText = yr !== 'MULTI' ? `Sesuaikan Data (${{yr}})` : 'Pilih Tahun';
      document.getElementById('btn-dl-label').innerText = yr !== 'MULTI' ? `Unduh JSON ${{yr}}` : 'Unduh JSON Master';
      document.getElementById('unsaved-year-label').innerText = yr;

      updateDataManagementBarStatus();

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
        renderProfitLossRibbon();
        renderProfitLossStatement();
      }}
    }}

    // Sub-tab switching between Sales and P&L
    function switchSubTab(tab) {{
      appState.activeSubTab = tab;
      const btnSales = document.getElementById('subtab-btn-SALES');
      const btnPL = document.getElementById('subtab-btn-PL');
      const panelSales = document.getElementById('subpanel-sales');
      const panelPL = document.getElementById('subpanel-pl');

      if (tab === 'SALES') {{
        btnSales.className = 'px-4 py-2 font-bold text-xs rounded-xl bg-slate-900 text-white shadow-sm flex items-center space-x-1.5';
        btnPL.className = 'px-4 py-2 font-bold text-xs rounded-xl text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-all flex items-center space-x-1.5';
        panelSales.classList.remove('hidden');
        panelPL.classList.add('hidden');
      }} else {{
        btnPL.className = 'px-4 py-2 font-bold text-xs rounded-xl bg-slate-900 text-white shadow-sm flex items-center space-x-1.5';
        btnSales.className = 'px-4 py-2 font-bold text-xs rounded-xl text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-all flex items-center space-x-1.5';
        panelPL.classList.remove('hidden');
        panelSales.classList.add('hidden');
        renderProfitLossStatement();
      }}
    }}

    // Get Active Dataset
    function getActiveYearData() {{
      if (appState.activeYear === '2024') return MASTER_DATA.data_2024;
      if (appState.activeYear === '2025') return MASTER_DATA.data_2025;
      if (appState.activeYear === '2026') return MASTER_DATA.data_2026;
      return MASTER_DATA.data_2026;
    }}

    // Render Profit & Loss Highlight Ribbon
    function renderProfitLossRibbon() {{
      const yrData = getActiveYearData();
      const pl = yrData.profit_loss;
      if (!pl) return;

      document.getElementById('pl-ribbon-period').innerText = pl.period_label;
      document.getElementById('pl-net-revenue').innerText = fmtIDR(pl.revenue.net_revenue);
      document.getElementById('pl-cogs').innerText = fmtIDR(pl.cogs.total_cogs);
      document.getElementById('pl-cogs-pct').innerText = (pl.cogs.cogs_pct || 0).toFixed(2) + '% dari Pendapatan';
      
      document.getElementById('pl-opex').innerText = fmtIDR(pl.operating_expenses.total);
      document.getElementById('pl-opex-pct').innerText = (pl.operating_expenses.pct_of_revenue || 0).toFixed(2) + '% dari Pendapatan';

      document.getElementById('pl-net-profit').innerText = fmtIDR(pl.net_profit.amount);
      document.getElementById('pl-net-margin').innerText = 'Net Margin: ' + (pl.net_profit.margin_pct || 0).toFixed(2) + '%';
      document.getElementById('pl-avg-monthly-profit').innerText = 'Rerata: ' + fmtIDR(pl.net_profit.avg_monthly) + '/bln';
    }}

    // Render Detailed Waterfall P&L Statement
    function renderProfitLossStatement() {{
      const yrData = getActiveYearData();
      const pl = yrData.profit_loss;
      if (!pl) return;

      document.getElementById('pl-table-title').innerText = `Laporan Laba / Rugi Standar (Tahun Buku ${{pl.year}})`;

      const container = document.getElementById('pl-statement-container');
      
      let opexItemsHtml = pl.operating_expenses.items.map(item => `
        <tr class="hover:bg-slate-50 text-[11px] text-slate-600">
          <td class="py-1.5 pl-8 pr-3">• ${{item.name}}</td>
          <td class="py-1.5 px-3 text-right font-mono">${{fmtIDR(item.amount)}}</td>
          <td class="py-1.5 px-3 text-right font-mono text-slate-400">${{((item.amount / pl.revenue.net_revenue) * 100).toFixed(2)}}%</td>
        </tr>
      `).join('');

      container.innerHTML = `
        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
              <tr>
                <th class="py-2.5 px-3">Deskripsi Akun Keuangan</th>
                <th class="py-2.5 px-3 text-right w-48">Nilai Nominal (Rp)</th>
                <th class="py-2.5 px-3 text-right w-36">% vs Pendapatan</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">
              <!-- Pendapatan -->
              <tr class="bg-slate-50/80 font-bold text-slate-900">
                <td class="py-2 px-3">PENDAPATAN PENJUALAN</td>
                <td class="py-2 px-3 text-right font-mono"></td>
                <td class="py-2 px-3 text-right font-mono"></td>
              </tr>
              <tr class="hover:bg-slate-50 text-slate-700">
                <td class="py-1.5 pl-6 pr-3">Penjualan Bersih (Barang Cetak &amp; Jasa)</td>
                <td class="py-1.5 px-3 text-right font-mono">${{fmtIDR(pl.revenue.gross_sales)}}</td>
                <td class="py-1.5 px-3 text-right font-mono text-slate-400">100.00%</td>
              </tr>
              ${{pl.revenue.sales_discounts ? `
              <tr class="hover:bg-slate-50 text-slate-600">
                <td class="py-1.5 pl-6 pr-3">Potongan Penjualan / Diskon</td>
                <td class="py-1.5 px-3 text-right font-mono text-rose-600">-${{fmtIDR(Math.abs(pl.revenue.sales_discounts_idr || pl.revenue.sales_discounts))}}</td>
                <td class="py-1.5 px-3 text-right font-mono text-slate-400">-${{((Math.abs(pl.revenue.sales_discounts_idr || pl.revenue.sales_discounts)/pl.revenue.gross_sales)*100).toFixed(2)}}%</td>
              </tr>` : ''}}
              ${{pl.revenue.sales_returns ? `
              <tr class="hover:bg-slate-50 text-slate-600">
                <td class="py-1.5 pl-6 pr-3">Retur Penjualan</td>
                <td class="py-1.5 px-3 text-right font-mono text-rose-600">-${{fmtIDR(Math.abs(pl.revenue.sales_returns))}}</td>
                <td class="py-1.5 px-3 text-right font-mono text-slate-400">-${{((Math.abs(pl.revenue.sales_returns)/pl.revenue.gross_sales)*100).toFixed(2)}}%</td>
              </tr>` : ''}}
              <tr class="bg-sky-50/60 font-bold text-sky-950 border-t border-sky-100">
                <td class="py-2 pl-4 pr-3">TOTAL PENDAPATAN BERSIH</td>
                <td class="py-2 px-3 text-right font-mono text-sky-800">${{fmtIDR(pl.revenue.net_revenue)}}</td>
                <td class="py-2 px-3 text-right font-mono text-sky-700">100.00%</td>
              </tr>

              <!-- HPP -->
              <tr class="bg-slate-50/80 font-bold text-slate-900">
                <td class="py-2 px-3">BEBAN POKOK PENJUALAN (HPP)</td>
                <td class="py-2 px-3 text-right font-mono"></td>
                <td class="py-2 px-3 text-right font-mono"></td>
              </tr>
              <tr class="hover:bg-slate-50 text-slate-700">
                <td class="py-1.5 pl-6 pr-3">Beban Pokok Penjualan (Bahan Baku, Tinta, Material)</td>
                <td class="py-1.5 px-3 text-right font-mono text-rose-600">${{fmtIDR(pl.cogs.total_cogs)}}</td>
                <td class="py-1.5 px-3 text-right font-mono text-rose-600">${{pl.cogs.cogs_pct.toFixed(2)}}%</td>
              </tr>
              <tr class="bg-slate-100/70 font-bold text-slate-900 border-t border-slate-200">
                <td class="py-2 pl-4 pr-3">LABA KOTOR (GROSS PROFIT)</td>
                <td class="py-2 px-3 text-right font-mono text-slate-900">${{fmtIDR(pl.gross_profit.amount)}}</td>
                <td class="py-2 px-3 text-right font-mono font-bold text-emerald-700">${{pl.gross_profit.margin_pct.toFixed(2)}}%</td>
              </tr>

              <!-- Beban Operasional -->
              <tr class="bg-slate-50/80 font-bold text-slate-900">
                <td class="py-2 px-3">BEBAN OPERASIONAL (OPEX)</td>
                <td class="py-2 px-3 text-right font-mono"></td>
                <td class="py-2 px-3 text-right font-mono"></td>
              </tr>
              ${{opexItemsHtml}}
              <tr class="bg-amber-50/60 font-bold text-amber-950 border-t border-amber-200">
                <td class="py-2 pl-4 pr-3">TOTAL BEBAN OPERASIONAL</td>
                <td class="py-2 px-3 text-right font-mono text-amber-800">${{fmtIDR(pl.operating_expenses.total)}}</td>
                <td class="py-2 px-3 text-right font-mono text-amber-700">${{pl.operating_expenses.pct_of_revenue.toFixed(2)}}%</td>
              </tr>

              <!-- Laba Operasional -->
              <tr class="bg-indigo-50/70 font-bold text-indigo-950 border-t border-indigo-200">
                <td class="py-2.5 px-3 text-indigo-900">PENDAPATAN OPERASIONAL (OPERATING INCOME)</td>
                <td class="py-2.5 px-3 text-right font-mono text-indigo-900 font-extrabold">${{fmtIDR(pl.operating_income.amount)}}</td>
                <td class="py-2.5 px-3 text-right font-mono text-indigo-800">${{pl.operating_income.pct_of_revenue.toFixed(2)}}%</td>
              </tr>

              <!-- Pendapatan / Beban Lain-lain -->
              <tr class="hover:bg-slate-50 text-slate-600">
                <td class="py-1.5 pl-6 pr-3">Pendapatan / (Beban) Lain-lain Bersih</td>
                <td class="py-1.5 px-3 text-right font-mono">${{fmtIDR(pl.other_income_expense.total)}}</td>
                <td class="py-1.5 px-3 text-right font-mono text-slate-400">${{((pl.other_income_expense.total / pl.revenue.net_revenue) * 100).toFixed(4)}}%</td>
              </tr>

              <!-- LABA BERSIH -->
              <tr class="bg-emerald-600 text-white font-extrabold text-sm border-t-2 border-emerald-700">
                <td class="py-3 px-3 tracking-wide flex items-center space-x-1.5">
                  <span>👑</span>
                  <span>LABA BERSIH TAHUN BERJALAN (NET PROFIT)</span>
                </td>
                <td class="py-3 px-3 text-right font-mono font-black text-emerald-100 text-base">${{fmtIDR(pl.net_profit.amount)}}</td>
                <td class="py-3 px-3 text-right font-mono text-emerald-100 font-black">${{pl.net_profit.margin_pct.toFixed(2)}}%</td>
              </tr>
            </tbody>
          </table>
        </div>
      `;
    }}

    // ========================================================
    // DATA CHANGE & PERSISTENCE MANAGEMENT LOGIC
    // ========================================================

    function updateDataManagementBarStatus() {{
      const yr = appState.activeYear;
      const statusIndicator = document.getElementById('sync-status-indicator');
      const statusText = document.getElementById('sync-status-text');
      const badgeUnsaved = document.getElementById('badge-unsaved-changes');
      const badgeCustom = document.getElementById('badge-custom-saved');
      const btnSaveMaster = document.getElementById('btn-save-master');
      const btnSaveLabel = document.getElementById('btn-save-label');
      const floatingBar = document.getElementById('floating-unsaved-bar');

      if (appState.hasUnsavedChanges) {{
        // Unsaved changes active
        statusIndicator.className = 'w-3.5 h-3.5 rounded-full bg-amber-500 animate-ping shadow-sm shrink-0';
        statusText.innerText = `Status Data: ⚠️ Ada Perubahan Belum Disimpan (Tahun ${{yr}})`;
        statusText.className = 'text-xs font-bold text-amber-700';
        badgeUnsaved.classList.remove('hidden');
        badgeCustom.classList.add('hidden');

        // Enable Main Save Button
        btnSaveMaster.disabled = false;
        btnSaveMaster.className = 'btn-save-dirty px-4 py-1.5 sm:px-4 sm:py-2 text-xs font-black rounded-xl shadow-lg flex items-center space-x-1.5 transition-all';
        btnSaveLabel.innerText = `Simpan Perubahan (${{yr}})`;

        floatingBar.classList.remove('hidden');
      }} else {{
        // All changes saved or pristine
        floatingBar.classList.add('hidden');
        badgeUnsaved.classList.add('hidden');

        const isCustom = savedCustomData && savedCustomData['data_' + yr];
        if (isCustom) {{
          statusIndicator.className = 'w-3.5 h-3.5 rounded-full bg-sky-500 shadow-sm shrink-0';
          statusText.innerText = `Status Data: 💾 Data Kustom Tersimpan (${{yr}})`;
          statusText.className = 'text-xs font-bold text-sky-800';
          badgeCustom.classList.remove('hidden');
        }} else {{
          statusIndicator.className = 'w-3.5 h-3.5 rounded-full bg-emerald-500 shadow-sm shrink-0';
          statusText.innerText = `Status Data: Standar ACCURATE Tersimpan (${{yr}})`;
          statusText.className = 'text-xs font-bold text-slate-800';
          badgeCustom.classList.add('hidden');
        }}

        btnSaveMaster.disabled = true;
        btnSaveMaster.className = 'px-3.5 py-1.5 sm:px-4 sm:py-2 text-xs font-bold rounded-xl bg-slate-200 text-slate-400 transition-all shadow-xs flex items-center space-x-1.5 cursor-not-allowed';
        btnSaveLabel.innerText = 'Data Tersimpan';
      }}
    }}

    // Open Data Editor Modal
    function openDataEditorModal() {{
      const yr = appState.activeYear;
      if (yr === 'MULTI') {{
        alert('Silakan pilih salah satu tahun buku (2024, 2025, atau 2026) untuk menyesuaikan data.');
        return;
      }}

      document.getElementById('modal-year-tag').innerText = yr;
      const yrData = getActiveYearData();
      const pl = yrData.profit_loss;

      // Populate P&L inputs
      document.getElementById('inp-gross-sales').value = pl.revenue.gross_sales || 0;
      document.getElementById('inp-sales-discount').value = Math.abs(pl.revenue.sales_discounts_idr || pl.revenue.sales_discounts || 0);
      document.getElementById('inp-sales-return').value = Math.abs(pl.revenue.sales_returns || 0);
      document.getElementById('inp-cogs').value = pl.cogs.total_cogs || 0;
      document.getElementById('inp-other-income').value = pl.other_income_expense.total || 0;

      // Populate Opex inputs
      const opexContainer = document.getElementById('editor-opex-inputs-container');
      opexContainer.innerHTML = '';
      pl.operating_expenses.items.forEach((item, idx) => {{
        const div = document.createElement('div');
        div.innerHTML = `
          <label class="block text-[11px] font-semibold text-slate-600 mb-1 truncate" title="${{item.name}}">${{item.name}}:</label>
          <input type="number" step="1000" id="inp-opex-${{idx}}" data-idx="${{idx}}" data-name="${{item.name}}" value="${{item.amount}}" oninput="onEditorInputChange()" class="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-mono text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none">
        `;
        opexContainer.appendChild(div);
      }});

      // Populate Monthly Sales inputs
      const monthlyTbody = document.getElementById('editor-monthly-tbody');
      monthlyTbody.innerHTML = '';
      yrData.monthly.forEach((m, idx) => {{
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50';
        tr.innerHTML = `
          <td class="py-1.5 px-3 font-semibold text-slate-700">${{m.name}}</td>
          <td class="py-1.5 px-3 text-right">
            <input type="number" step="1000" id="inp-month-sales-${{idx}}" data-idx="${{idx}}" value="${{m.sales}}" oninput="onEditorInputChange()" class="w-40 px-2.5 py-1 text-right border border-slate-300 rounded-lg text-xs font-mono font-bold text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none">
          </td>
          <td class="py-1.5 px-3 text-right">
            <input type="number" step="1" id="inp-month-qty-${{idx}}" data-idx="${{idx}}" value="${{m.qty}}" oninput="onEditorInputChange()" class="w-28 px-2.5 py-1 text-right border border-slate-300 rounded-lg text-xs font-mono text-slate-700 focus:ring-2 focus:ring-emerald-500 focus:outline-none">
          </td>
        `;
        monthlyTbody.appendChild(tr);
      }});

      // Snapshot initial values
      snapshotEditorInitialValues();
      recalculateSimulationPreview();

      document.getElementById('data-editor-modal').classList.remove('hidden');
    }}

    function closeDataEditorModal() {{
      document.getElementById('data-editor-modal').classList.add('hidden');
    }}

    function switchEditorTab(tab) {{
      appState.activeEditorTab = tab;
      const tabPL = document.getElementById('tab-edit-pl');
      const tabSales = document.getElementById('tab-edit-sales');
      const panelPL = document.getElementById('editor-panel-pl');
      const panelSales = document.getElementById('editor-panel-sales');

      if (tab === 'PL') {{
        tabPL.className = 'px-3 py-1.5 font-bold rounded-lg bg-white text-slate-900 shadow-xs';
        tabSales.className = 'px-3 py-1.5 font-medium rounded-lg text-slate-600 hover:text-slate-900';
        panelPL.classList.remove('hidden');
        panelSales.classList.add('hidden');
      }} else {{
        tabSales.className = 'px-3 py-1.5 font-bold rounded-lg bg-white text-slate-900 shadow-xs';
        tabPL.className = 'px-3 py-1.5 font-medium rounded-lg text-slate-600 hover:text-slate-900';
        panelSales.classList.remove('hidden');
        panelPL.classList.add('hidden');
      }}
    }}

    function snapshotEditorInitialValues() {{
      const vals = {{}};
      vals.gross_sales = parseFloat(document.getElementById('inp-gross-sales').value) || 0;
      vals.sales_discount = parseFloat(document.getElementById('inp-sales-discount').value) || 0;
      vals.sales_return = parseFloat(document.getElementById('inp-sales-return').value) || 0;
      vals.cogs = parseFloat(document.getElementById('inp-cogs').value) || 0;
      vals.other_income = parseFloat(document.getElementById('inp-other-income').value) || 0;

      const yrData = getActiveYearData();
      vals.opex = [];
      yrData.profit_loss.operating_expenses.items.forEach((_, idx) => {{
        const inp = document.getElementById('inp-opex-' + idx);
        vals.opex.push(parseFloat(inp ? inp.value : 0) || 0);
      }});

      vals.monthly = [];
      yrData.monthly.forEach((_, idx) => {{
        const inpS = document.getElementById('inp-month-sales-' + idx);
        const inpQ = document.getElementById('inp-month-qty-' + idx);
        vals.monthly.push({{
          sales: parseFloat(inpS ? inpS.value : 0) || 0,
          qty: parseFloat(inpQ ? inpQ.value : 0) || 0
        }});
      }});

      appState.initialEditorValues = vals;
    }}

    function restoreInitialModalValues() {{
      const vals = appState.initialEditorValues;
      if (!vals || !vals.gross_sales) return;

      document.getElementById('inp-gross-sales').value = vals.gross_sales;
      document.getElementById('inp-sales-discount').value = vals.sales_discount;
      document.getElementById('inp-sales-return').value = vals.sales_return;
      document.getElementById('inp-cogs').value = vals.cogs;
      document.getElementById('inp-other-income').value = vals.other_income;

      vals.opex.forEach((v, idx) => {{
        const inp = document.getElementById('inp-opex-' + idx);
        if (inp) inp.value = v;
      }});

      vals.monthly.forEach((m, idx) => {{
        const inpS = document.getElementById('inp-month-sales-' + idx);
        const inpQ = document.getElementById('inp-month-qty-' + idx);
        if (inpS) inpS.value = m.sales;
        if (inpQ) inpQ.value = m.qty;
      }});

      onEditorInputChange();
    }}

    // When inputs are edited
    function onEditorInputChange() {{
      recalculateSimulationPreview();

      // Check if dirty
      let isDifferent = false;
      const init = appState.initialEditorValues;
      if (init) {{
        const g = parseFloat(document.getElementById('inp-gross-sales').value) || 0;
        const d = parseFloat(document.getElementById('inp-sales-discount').value) || 0;
        const r = parseFloat(document.getElementById('inp-sales-return').value) || 0;
        const c = parseFloat(document.getElementById('inp-cogs').value) || 0;
        const o = parseFloat(document.getElementById('inp-other-income').value) || 0;

        if (g !== init.gross_sales || d !== init.sales_discount || r !== init.sales_return || c !== init.cogs || o !== init.other_income) {{
          isDifferent = true;
        }}

        if (!isDifferent && init.opex) {{
          for (let i = 0; i < init.opex.length; i++) {{
            const inp = document.getElementById('inp-opex-' + i);
            if (inp && (parseFloat(inp.value) || 0) !== init.opex[i]) {{
              isDifferent = true;
              break;
            }}
          }}
        }}

        if (!isDifferent && init.monthly) {{
          for (let i = 0; i < init.monthly.length; i++) {{
            const inpS = document.getElementById('inp-month-sales-' + i);
            const inpQ = document.getElementById('inp-month-qty-' + i);
            if (inpS && (parseFloat(inpS.value) || 0) !== init.monthly[i].sales) {{ isDifferent = true; break; }}
            if (inpQ && (parseFloat(inpQ.value) || 0) !== init.monthly[i].qty) {{ isDifferent = true; break; }}
          }}
        }}
      }}

      appState.hasUnsavedChanges = isDifferent;

      // Update Modal Controls
      const modalSaveBtn = document.getElementById('modal-save-btn');
      const modalDot = document.getElementById('modal-status-dot');
      const modalStatus = document.getElementById('modal-change-status');

      if (isDifferent) {{
        modalSaveBtn.disabled = false;
        modalSaveBtn.className = 'btn-save-dirty px-4 py-2 text-xs font-black rounded-xl shadow-lg flex items-center space-x-1.5 transition-all';
        modalDot.className = 'w-2.5 h-2.5 rounded-full bg-amber-500 animate-ping';
        modalStatus.innerText = '⚠️ Terdeteksi perubahan nilai data (belum disimpan).';
        modalStatus.className = 'text-xs font-bold text-amber-600';
      }} else {{
        modalSaveBtn.disabled = true;
        modalSaveBtn.className = 'px-4 py-2 text-xs font-bold rounded-xl bg-slate-300 text-slate-500 cursor-not-allowed flex items-center space-x-1.5 transition-all shadow-xs';
        modalDot.className = 'w-2.5 h-2.5 rounded-full bg-slate-400';
        modalStatus.innerText = 'Belum ada perubahan nilai.';
        modalStatus.className = 'text-xs font-medium text-slate-600';
      }}

      updateDataManagementBarStatus();
    }}

    function recalculateSimulationPreview() {{
      const gross = parseFloat(document.getElementById('inp-gross-sales').value) || 0;
      const disc = parseFloat(document.getElementById('inp-sales-discount').value) || 0;
      const ret = parseFloat(document.getElementById('inp-sales-return').value) || 0;
      const cogs = parseFloat(document.getElementById('inp-cogs').value) || 0;
      const other = parseFloat(document.getElementById('inp-other-income').value) || 0;

      const netRev = Math.max(0, gross - disc - ret);
      const grossProfit = netRev - cogs;
      const cogsPct = netRev > 0 ? (cogs / netRev) * 100 : 0;

      let totalOpex = 0;
      const yrData = getActiveYearData();
      yrData.profit_loss.operating_expenses.items.forEach((_, idx) => {{
        const inp = document.getElementById('inp-opex-' + idx);
        if (inp) totalOpex += parseFloat(inp.value) || 0;
      }});

      const opIncome = grossProfit - totalOpex;
      const netProfit = opIncome + other;
      const netMargin = netRev > 0 ? (netProfit / netRev) * 100 : 0;
      const mCount = yrData.profit_loss.months_count || (appState.activeYear === '2026' ? 8 : 12);
      const avgMonthly = netProfit / mCount;

      // Update Preview Elements
      document.getElementById('sub-net-revenue').innerText = 'Net: ' + fmtIDR(netRev);
      document.getElementById('sub-gross-profit').innerText = 'Laba Kotor: ' + fmtIDR(grossProfit);
      document.getElementById('sub-cogs-pct').innerText = cogsPct.toFixed(2) + '%';
      document.getElementById('sub-total-opex').innerText = 'Total Opex: ' + fmtIDR(totalOpex);

      document.getElementById('sim-net-profit').innerText = fmtIDR(netProfit);
      document.getElementById('sim-net-margin').innerText = netMargin.toFixed(2) + '%';
      document.getElementById('sim-avg-profit').innerText = fmtIDR(avgMonthly) + '/bln';
    }}

    function discardChanges() {{
      restoreInitialModalValues();
      appState.hasUnsavedChanges = false;
      updateDataManagementBarStatus();
      closeDataEditorModal();
      showToast('Perubahan data dibatalkan.');
    }}

    // ========================================================
    // SAVE DATA CHANGES TO SYSTEM & LOCAL STORAGE
    // ========================================================
    async function saveDataChanges() {{
      const yr = appState.activeYear;
      if (yr === 'MULTI') return;

      const yrData = getActiveYearData();
      const pl = yrData.profit_loss;

      // Extract values from editor inputs (or current if modal wasn't opened)
      let gross = document.getElementById('inp-gross-sales') ? (parseFloat(document.getElementById('inp-gross-sales').value) || pl.revenue.gross_sales) : pl.revenue.gross_sales;
      let disc = document.getElementById('inp-sales-discount') ? (parseFloat(document.getElementById('inp-sales-discount').value) || 0) : Math.abs(pl.revenue.sales_discounts_idr || 0);
      let ret = document.getElementById('inp-sales-return') ? (parseFloat(document.getElementById('inp-sales-return').value) || 0) : Math.abs(pl.revenue.sales_returns || 0);
      let cogs = document.getElementById('inp-cogs') ? (parseFloat(document.getElementById('inp-cogs').value) || pl.cogs.total_cogs) : pl.cogs.total_cogs;
      let other = document.getElementById('inp-other-income') ? (parseFloat(document.getElementById('inp-other-income').value) || pl.other_income_expense.total) : pl.other_income_expense.total;

      const netRev = Math.max(0, gross - disc - ret);
      const grossProfit = netRev - cogs;
      const cogsPct = netRev > 0 ? (cogs / netRev) * 100 : 0;
      const grossMarginPct = netRev > 0 ? (grossProfit / netRev) * 100 : 0;

      // Opex items
      let totalOpex = 0;
      pl.operating_expenses.items.forEach((item, idx) => {{
        const inp = document.getElementById('inp-opex-' + idx);
        if (inp) {{
          item.amount = parseFloat(inp.value) || 0;
        }}
        totalOpex += item.amount;
      }});

      const opIncome = grossProfit - totalOpex;
      const opIncomePct = netRev > 0 ? (opIncome / netRev) * 100 : 0;
      const netProfit = opIncome + other;
      const netMargin = netRev > 0 ? (netProfit / netRev) * 100 : 0;
      const mCount = pl.months_count || (yr === '2026' ? 8 : 12);
      const avgMonthly = netProfit / mCount;

      // Apply updates to in-memory P&L structure
      pl.revenue.gross_sales = gross;
      pl.revenue.sales_discounts = -disc;
      pl.revenue.sales_discounts_idr = -disc;
      pl.revenue.sales_returns = -ret;
      pl.revenue.net_revenue = netRev;

      pl.cogs.total_cogs = cogs;
      pl.cogs.cogs_pct = cogsPct;

      pl.gross_profit.amount = grossProfit;
      pl.gross_profit.margin_pct = grossMarginPct;

      pl.operating_expenses.total = totalOpex;
      pl.operating_expenses.pct_of_revenue = netRev > 0 ? (totalOpex / netRev) * 100 : 0;

      pl.operating_income.amount = opIncome;
      pl.operating_income.pct_of_revenue = opIncomePct;

      pl.other_income_expense.total = other;

      pl.net_profit.amount = netProfit;
      pl.net_profit.margin_pct = netMargin;
      pl.net_profit.avg_monthly = avgMonthly;

      // Update monthly sales if inputs present
      let totalAnnualSales = 0;
      let totalAnnualQty = 0;
      yrData.monthly.forEach((m, idx) => {{
        const inpS = document.getElementById('inp-month-sales-' + idx);
        const inpQ = document.getElementById('inp-month-qty-' + idx);
        if (inpS) m.sales = parseFloat(inpS.value) || 0;
        if (inpQ) m.qty = parseFloat(inpQ.value) || 0;
        totalAnnualSales += m.sales;
        totalAnnualQty += m.qty;
      }});

      yrData.kpi.total_sales = totalAnnualSales;
      yrData.kpi.total_qty = totalAnnualQty;
      yrData.kpi.avg_monthly_sales = totalAnnualSales / mCount;

      // Update multiyear comparison summary
      const sumIndex = (yr === '2024') ? 0 : (yr === '2025' ? 1 : 2);
      if (MASTER_DATA.profit_loss_comparison && MASTER_DATA.profit_loss_comparison.summary[sumIndex]) {{
        const s = MASTER_DATA.profit_loss_comparison.summary[sumIndex];
        s.revenue = netRev;
        s.cogs = cogs;
        s.cogs_pct = cogsPct;
        s.gross_profit = grossProfit;
        s.gross_margin_pct = grossMarginPct;
        s.opex = totalOpex;
        s.opex_pct = pl.operating_expenses.pct_of_revenue;
        s.net_profit = netProfit;
        s.net_margin_pct = netMargin;
        s.avg_monthly_net_profit = avgMonthly;
      }}

      // Persist to localStorage
      try {{
        if (!savedCustomData) savedCustomData = {{}};
        savedCustomData['data_' + yr] = yrData;
        savedCustomData.profit_loss_comparison = MASTER_DATA.profit_loss_comparison;
        localStorage.setItem(STORAGE_KEY, JSON.stringify(savedCustomData));
      }} catch(err) {{
        console.warn('LocalStorage save error:', err);
      }}

      // Attempt to save to Python Server API (if active)
      let backendSaved = false;
      try {{
        const resp = await fetch('/api/save', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{
            year: parseInt(yr),
            profit_loss: pl,
            monthly: yrData.monthly
          }})
        }});
        if (resp.ok) {{
          const resJson = await resp.json();
          if (resJson.success) backendSaved = true;
        }}
      }} catch(netErr) {{
        // Expected when opening locally via file:///
      }}

      // Reset dirty state
      appState.hasUnsavedChanges = false;
      snapshotEditorInitialValues();

      // Refresh Dashboard UI
      renderProfitLossRibbon();
      renderProfitLossStatement();
      updateDashboard();
      updateDataManagementBarStatus();
      closeDataEditorModal();

      if (backendSaved) {{
        showToast(`✅ Perubahan data tahun ${{yr}} berhasil disimpan ke file folder dan database!`);
      }} else {{
        showToast(`✅ Perubahan data tahun ${{yr}} berhasil disimpan ke sistem browser!`);
      }}
    }}

    // Reset Data Changes back to ACCURATE Standard
    function resetDataChanges() {{
      const yr = appState.activeYear;
      if (yr === 'MULTI') return;

      if (!confirm(`Kembalikan data Tahun ${{yr}} ke standar resmi ACCURATE? Seluruh penyesuaian kustom akan dihapus.`)) {{
        return;
      }}

      // Restore from original
      MASTER_DATA['data_' + yr] = JSON.parse(JSON.stringify(ORIGINAL_MASTER_DATA['data_' + yr]));
      
      const sumIndex = (yr === '2024') ? 0 : (yr === '2025' ? 1 : 2);
      MASTER_DATA.profit_loss_comparison.summary[sumIndex] = JSON.parse(JSON.stringify(ORIGINAL_MASTER_DATA.profit_loss_comparison.summary[sumIndex]));

      // Clear from localStorage
      if (savedCustomData) {{
        delete savedCustomData['data_' + yr];
        try {{
          localStorage.setItem(STORAGE_KEY, JSON.stringify(savedCustomData));
        }} catch(e) {{}}
      }}

      appState.hasUnsavedChanges = false;
      selectYear(yr);
      showToast(`🔄 Data tahun ${{yr}} telah direset kembali ke Standar ACCURATE.`);
    }}

    // Download Folder JSON
    function downloadFolderJson() {{
      const yr = appState.activeYear;
      let filename = `laba_rugi_${{yr}}.json`;
      let content = '';

      if (yr === 'MULTI') {{
        filename = 'sales_data_master.json';
        content = JSON.stringify(MASTER_DATA, null, 2);
      }} else {{
        const yrData = getActiveYearData();
        content = JSON.stringify(yrData.profit_loss, null, 2);
      }}

      const blob = new Blob([content], {{ type: 'application/json' }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      showToast(`📥 File ${{filename}} berhasil diunduh ke folder komputer Anda!`);
    }}

    // Toast Helper
    function showToast(msg) {{
      const toast = document.getElementById('toast-notification');
      const toastMsg = document.getElementById('toast-message');
      toastMsg.innerText = msg;

      toast.classList.remove('hidden', 'translate-y-[-20px]', 'opacity-0');
      toast.classList.add('translate-y-0', 'opacity-100');

      setTimeout(() => {{
        toast.classList.remove('translate-y-0', 'opacity-100');
        toast.classList.add('translate-y-[-20px]', 'opacity-0');
        setTimeout(() => toast.classList.add('hidden'), 300);
      }}, 3500);
    }}

    // ========================================================
    // PERIOD & CATEGORY FILTERS
    // ========================================================

    function initPeriodButtons() {{
      const container = document.getElementById('month-buttons-container');
      container.innerHTML = '';
      const yr = appState.activeYear;
      const monthKeys = (yr === '2026') ? MONTH_KEYS_8 : MONTH_KEYS_12;

      // All button
      const btnAll = document.createElement('button');
      btnAll.id = 'filter-btn-ALL';
      btnAll.className = 'px-2.5 py-1 rounded-lg text-xs font-bold transition-all bg-sky-600 text-white shadow-xs';
      btnAll.innerText = 'Semua';
      btnAll.onclick = () => setPeriodFilter('ALL');
      container.appendChild(btnAll);

      // Half-year buttons (H1, H2)
      const btnH1 = document.createElement('button');
      btnH1.id = 'filter-btn-H1';
      btnH1.className = 'px-2 py-1 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-200 transition-all';
      btnH1.innerText = 'Semester 1';
      btnH1.onclick = () => setPeriodFilter('H1');
      container.appendChild(btnH1);

      if (yr !== '2026') {{
        const btnH2 = document.createElement('button');
        btnH2.id = 'filter-btn-H2';
        btnH2.className = 'px-2 py-1 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-200 transition-all';
        btnH2.innerText = 'Semester 2';
        btnH2.onclick = () => setPeriodFilter('H2');
        container.appendChild(btnH2);
      }}

      // Months
      monthKeys.forEach(m => {{
        const btn = document.createElement('button');
        btn.id = 'filter-btn-' + m;
        btn.className = 'px-2 py-1 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-200 transition-all';
        btn.innerText = m;
        btn.onclick = () => setPeriodFilter(m);
        container.appendChild(btn);
      }});
    }}

    function setPeriodFilter(p) {{
      appState.period = p;
      const yr = appState.activeYear;
      const monthKeys = (yr === '2026') ? MONTH_KEYS_8 : MONTH_KEYS_12;
      const allButtons = ['ALL', 'H1', 'H2', ...monthKeys];

      allButtons.forEach(b => {{
        const el = document.getElementById('filter-btn-' + b);
        if (!el) return;
        if (b === p) {{
          el.className = 'px-2.5 py-1 rounded-lg text-xs font-bold transition-all bg-sky-600 text-white shadow-xs';
        }} else {{
          el.className = 'px-2 py-1 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-200 transition-all';
        }}
      }});

      const lbl = document.getElementById('active-period-label');
      if (p === 'ALL') lbl.innerText = 'Semua Periode (' + (yr === '2026' ? '8 Bulan' : '12 Bulan') + ')';
      else if (p === 'H1') lbl.innerText = 'Semester 1 (Jan - Jun)';
      else if (p === 'H2') lbl.innerText = 'Semester 2 (Jul - Des)';
      else lbl.innerText = 'Bulan ' + (MONTH_NAMES[p] || p);

      updateDashboard();
    }}

    function initCategoryFilterOptions() {{
      const select = document.getElementById('category-filter-select');
      select.innerHTML = '<option value="ALL">Semua Kategori Produk</option>';
      const yrData = getActiveYearData();

      if (yrData.categories) {{
        yrData.categories.forEach(c => {{
          const opt = document.createElement('option');
          opt.value = c.name;
          opt.innerText = c.name + ' (' + c.pct.toFixed(1) + '%)';
          select.appendChild(opt);
        }});
      }}
    }}

    function onCategorySelect(val) {{
      appState.category = val;
      updateDashboard();
    }}

    function setParetoFilter(val) {{
      appState.pareto = val;
      ['ALL', 'CORE80', 'LONGTAIL'].forEach(p => {{
        const btn = document.getElementById('pareto-btn-' + p);
        if (p === val) {{
          btn.className = 'px-2.5 py-1 font-bold rounded-md bg-white text-slate-800 shadow-xs';
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
      setPeriodFilter('ALL');
      appState.category = 'ALL';
      document.getElementById('category-filter-select').value = 'ALL';
      setParetoFilter('ALL');
      appState.search = '';
      const sInp = document.getElementById('product-search-input');
      if (sInp) sInp.value = '';
      updateDashboard();
    }}

    // Main Single Year View Update
    function updateDashboard() {{
      const yr = appState.activeYear;
      const yrData = getActiveYearData();

      document.getElementById('chart-title').innerText = `Tren Penjualan Bulanan (Tahun Buku ${{yr}})`;
      document.getElementById('chart-year-badge').innerText = `Tahun ${{yr}}`;
      document.getElementById('filter-header-title').innerText = `Navigasi &amp; Filter Kinerja Tahun ${{yr}}`;

      renderMonthlyChart();
      renderCategoryBreakdown();
      renderDynamicTable();
    }}

    // Monthly Chart SVG
    function renderMonthlyChart() {{
      const svg = document.getElementById('svg-monthly-chart');
      svg.innerHTML = '';
      const yr = appState.activeYear;
      const yrData = getActiveYearData();
      const monthData = yrData.monthly;

      const maxSales = Math.max(...monthData.map(m => m.sales), 1);
      const width = 680;
      const height = 240;
      const padLeft = 70;
      const padBottom = 35;
      const padTop = 20;
      const chartW = width - padLeft - 20;
      const chartH = height - padBottom - padTop;
      const count = monthData.length;
      const barW = Math.min(38, (chartW / count) * 0.65);
      const step = chartW / count;

      // Grid lines
      for (let i = 0; i <= 4; i++) {{
        const y = padTop + (chartH / 4) * i;
        const val = maxSales * (1 - i / 4);

        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', padLeft);
        line.setAttribute('x2', width - 20);
        line.setAttribute('y1', y);
        line.setAttribute('y2', y);
        line.setAttribute('stroke', '#e2e8f0');
        line.setAttribute('stroke-dasharray', '3,3');
        svg.appendChild(line);

        const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt.setAttribute('x', padLeft - 8);
        txt.setAttribute('y', y + 3);
        txt.setAttribute('text-anchor', 'end');
        txt.setAttribute('class', 'text-[9px] fill-slate-400 font-mono');
        txt.textContent = 'Rp ' + Math.round(val / 1000000) + ' Jt';
        svg.appendChild(txt);
      }}

      // Bars
      monthData.forEach((m, idx) => {{
        const bH = (m.sales / maxSales) * chartH;
        const x = padLeft + idx * step + (step - barW) / 2;
        const y = padTop + chartH - bH;

        const isSelected = (appState.period === m.code);
        const isPeak = (yrData.kpi.peak_month && m.name.startsWith(yrData.kpi.peak_month));

        let fillColor = '#0284c7'; // sky-600
        if (yr === '2026') fillColor = '#059669'; // emerald-600
        if (yr === '2025') fillColor = '#4f46e5'; // indigo-600
        if (isPeak) fillColor = '#f59e0b';       // amber-500
        if (isSelected) fillColor = '#e11d48';   // rose-600

        const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        g.setAttribute('class', 'chart-bar-group cursor-pointer');
        g.onclick = () => setPeriodFilter(m.code);

        const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect.setAttribute('x', x);
        rect.setAttribute('y', y);
        rect.setAttribute('width', barW);
        rect.setAttribute('height', bH);
        rect.setAttribute('rx', '4');
        rect.setAttribute('fill', fillColor);
        g.appendChild(rect);

        // Sales label on top of bar
        const valTxt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        valTxt.setAttribute('x', x + barW / 2);
        valTxt.setAttribute('y', y - 4);
        valTxt.setAttribute('text-anchor', 'middle');
        valTxt.setAttribute('class', 'text-[9px] font-bold fill-slate-700 font-mono');
        valTxt.textContent = (m.sales / 1000000).toFixed(1) + 'M';
        g.appendChild(valTxt);

        // Month code bottom label
        const lbl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        lbl.setAttribute('x', x + barW / 2);
        lbl.setAttribute('y', height - 12);
        lbl.setAttribute('text-anchor', 'middle');
        lbl.setAttribute('class', `text-[10px] ${{isSelected ? 'font-black fill-rose-600' : 'font-semibold fill-slate-600'}}`);
        lbl.textContent = m.code;
        g.appendChild(lbl);

        svg.appendChild(g);
      }});
    }}

    // Category Contribution Progress Bars
    function renderCategoryBreakdown() {{
      const container = document.getElementById('category-progress-container');
      container.innerHTML = '';
      const yrData = getActiveYearData();
      if (!yrData.categories) return;

      const colors = ['bg-sky-500', 'bg-indigo-500', 'bg-emerald-500', 'bg-amber-500', 'bg-purple-500', 'bg-rose-500'];

      yrData.categories.forEach((c, idx) => {{
        const color = colors[idx % colors.length];
        const isSelected = (appState.category === c.name);

        const div = document.createElement('div');
        div.className = `p-2 rounded-xl transition-all cursor-pointer ${{isSelected ? 'bg-sky-50 border border-sky-300 shadow-2xs' : 'hover:bg-slate-50'}}`;
        div.onclick = () => {{
          appState.category = (appState.category === c.name) ? 'ALL' : c.name;
          document.getElementById('category-filter-select').value = appState.category;
          updateDashboard();
        }};

        div.innerHTML = `
          <div class="flex items-center justify-between text-xs mb-1">
            <span class="font-bold text-slate-800 truncate mr-2">${{c.name}}</span>
            <span class="font-mono font-bold text-slate-700">${{fmtIDR(c.sales)}} <span class="text-[10px] text-slate-500 font-normal">(${{c.pct.toFixed(1)}}%)</span></span>
          </div>
          <div class="w-full bg-slate-100 rounded-full h-2 overflow-hidden">
            <div class="${{color}} h-2 rounded-full transition-all duration-500" style="width: ${{c.pct}}%"></div>
          </div>
        `;
        container.appendChild(div);
      }});

      if (yrData.categories.length >= 2) {{
        const corePct = yrData.categories[0].pct + yrData.categories[1].pct;
        document.getElementById('cat-core-text').innerText = `${{yrData.categories[0].name.split(' ')[0]}} & ${{yrData.categories[1].name.split(' ')[0]}} (${{corePct.toFixed(1)}}%)`;
      }}
    }}

    // Dynamic Table Rendering (Product table for 2024, Matrix table for 2025/2026)
    function renderDynamicTable() {{
      const container = document.getElementById('dynamic-table-container');
      const yr = appState.activeYear;

      if (yr === '2024') {{
        render2024ProductsTable(container);
      }} else if (yr === '2025') {{
        render2025MatrixTable(container);
      }} else if (yr === '2026') {{
        render2026MatrixTable(container);
      }}
    }}

    function render2024ProductsTable(container) {{
      document.getElementById('table-title').innerText = 'Daftar 129 Produk & Kategori (Tahun 2024)';
      document.getElementById('table-subtitle').innerText = 'Klik baris mana pun untuk melihat grafik penjualan 12 bulan & rincian kuantitas produk.';

      let prods = MASTER_DATA.data_2024.products.slice();

      // Filters
      if (appState.category !== 'ALL') {{
        prods = prods.filter(p => p.category === appState.category);
      }}
      if (appState.pareto === 'CORE80') {{
        prods = prods.filter(p => p.rank <= 24);
      }} else if (appState.pareto === 'LONGTAIL') {{
        prods = prods.filter(p => p.rank > 24);
      }}
      if (appState.search) {{
        prods = prods.filter(p => p.name.toLowerCase().includes(appState.search));
      }}

      // Sorting
      prods.sort((a, b) => {{
        let vA = (appState.period === 'ALL') ? a.total_sales : (a.monthly_sales[appState.period] || 0);
        let vB = (appState.period === 'ALL') ? b.total_sales : (b.monthly_sales[appState.period] || 0);
        return appState.sortAsc ? (vA - vB) : (vB - vA);
      }});

      let rowsHtml = prods.map(p => {{
        const sales = (appState.period === 'ALL') ? p.total_sales : (p.monthly_sales[appState.period] || 0);
        const qty = (appState.period === 'ALL') ? p.total_qty : (p.monthly_qty[appState.period] || 0);
        const isCore = p.rank <= 24;

        return `
          <tr class="hover:bg-sky-50/60 cursor-pointer transition-colors" onclick="openProductModal(${{p.rank}})">
            <td class="py-2 px-3 text-center font-bold ${{isCore ? 'text-amber-600 font-mono' : 'text-slate-400'}}">${{p.rank}}</td>
            <td class="py-2 px-3 font-semibold text-slate-800">
              <div class="flex items-center space-x-1.5">
                <span>${{p.name}}</span>
                ${{isCore ? '<span class="text-[10px] text-amber-500 font-bold" title="Core 80% Pareto">⭐</span>' : ''}}
              </div>
            </td>
            <td class="py-2 px-3 text-slate-500">${{p.category}}</td>
            <td class="py-2 px-3 text-right font-mono text-slate-700">${{fmtNum(qty)}} <span class="text-[10px] text-slate-400 font-sans">${{p.unit}}</span></td>
            <td class="py-2 px-3 text-right font-mono font-bold text-slate-900">${{fmtIDR(sales)}}</td>
            <td class="py-2 px-3 text-right font-mono text-sky-700 font-semibold">${{p.pct.toFixed(2)}}%</td>
          </tr>
        `;
      }}).join('');

      container.innerHTML = `
        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
              <tr>
                <th class="py-2.5 px-3 text-center w-14">Rank</th>
                <th class="py-2.5 px-3">Nama Produk / Barang Cetak</th>
                <th class="py-2.5 px-3">Kategori</th>
                <th class="py-2.5 px-3 text-right">Volume</th>
                <th class="py-2.5 px-3 text-right cursor-pointer select-none" onclick="toggleSort()">
                  Penjualan (Rp) ${{appState.sortAsc ? '▲' : '▼'}}
                </th>
                <th class="py-2.5 px-3 text-right w-20">Porsi</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">${{rowsHtml}}</tbody>
          </table>
        </div>
      `;
    }}

    function render2025MatrixTable(container) {{
      document.getElementById('table-title').innerText = 'Matriks Penjualan Bulanan per Kategori (Tahun 2025)';
      document.getElementById('table-subtitle').innerText = 'Data konsolidasi 12 bulan penuh dari pembukuan ACCURATE Accounting System.';

      const mData = MASTER_DATA.data_2025.monthly;
      let rowsHtml = mData.map(m => `
        <tr class="hover:bg-indigo-50/60 transition-colors ${{appState.period === m.code ? 'bg-indigo-50 font-bold' : ''}}">
          <td class="py-2 px-3 font-semibold text-slate-900">${{m.name}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.large_format)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.print_a3)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.doc_fc)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.finishing)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.merchandise)}}</td>
          <td class="py-2 px-3 text-right font-mono font-bold text-indigo-900 bg-indigo-50/50">${{fmtIDR(m.sales)}}</td>
        </tr>
      `).join('');

      container.innerHTML = `
        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
              <tr>
                <th class="py-2.5 px-3">Bulan</th>
                <th class="py-2.5 px-3 text-right">Large Format</th>
                <th class="py-2.5 px-3 text-right">Print A3+ &amp; Sheet</th>
                <th class="py-2.5 px-3 text-right">Doc &amp; Fotocopy</th>
                <th class="py-2.5 px-3 text-right">Finishing</th>
                <th class="py-2.5 px-3 text-right">Merchandise</th>
                <th class="py-2.5 px-3 text-right bg-slate-800">Total Omzet</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">${{rowsHtml}}</tbody>
            <tfoot class="bg-slate-100 font-mono font-bold text-slate-900 border-t-2 border-slate-300">
              <tr>
                <td class="py-2.5 px-3 font-sans">TOTAL 2025</td>
                <td class="py-2.5 px-3 text-right">Rp 509.000.000</td>
                <td class="py-2.5 px-3 text-right">Rp 488.600.000</td>
                <td class="py-2.5 px-3 text-right">Rp 118.900.000</td>
                <td class="py-2.5 px-3 text-right">Rp 64.801.929</td>
                <td class="py-2.5 px-3 text-right">Rp 157.600.000</td>
                <td class="py-2.5 px-3 text-right text-indigo-900 font-extrabold bg-indigo-100/60">${{fmtIDR(MASTER_DATA.data_2025.kpi.total_sales)}}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      `;
    }}

    function render2026MatrixTable(container) {{
      document.getElementById('table-title').innerText = 'Matriks Penjualan Bulanan per Kategori (Tahun 2026 - 8 Bulan)';
      document.getElementById('table-subtitle').innerText = 'Realisasi pembukuan ACCURATE periode 01 Januari s/d 31 Agustus 2026.';

      const mData = MASTER_DATA.data_2026.monthly;
      let rowsHtml = mData.map(m => `
        <tr class="hover:bg-amber-50/60 transition-colors ${{appState.period === m.code ? 'bg-amber-50 font-bold' : ''}}">
          <td class="py-2 px-3 font-semibold text-slate-900">${{m.name}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.large_format)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.print_a3)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.doc_fc)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.desain_khusus)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.merchandise)}}</td>
          <td class="py-2 px-3 text-right font-mono">${{fmtIDR(m.matrix.finishing)}}</td>
          <td class="py-2 px-3 text-right font-mono font-bold text-amber-900 bg-amber-50/50">${{fmtIDR(m.sales)}}</td>
        </tr>
      `).join('');

      container.innerHTML = `
        <div class="overflow-x-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-900 text-white uppercase text-[11px] font-semibold tracking-wider">
              <tr>
                <th class="py-2.5 px-3">Bulan</th>
                <th class="py-2.5 px-3 text-right">Large Format</th>
                <th class="py-2.5 px-3 text-right">Print A3+ &amp; Sheet</th>
                <th class="py-2.5 px-3 text-right">Doc &amp; IR</th>
                <th class="py-2.5 px-3 text-right">Desain &amp; Nota</th>
                <th class="py-2.5 px-3 text-right">Merchandise</th>
                <th class="py-2.5 px-3 text-right">Finishing</th>
                <th class="py-2.5 px-3 text-right bg-slate-800">Total Omzet</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">${{rowsHtml}}</tbody>
            <tfoot class="bg-slate-100 font-mono font-bold text-slate-900 border-t-2 border-slate-300">
              <tr>
                <td class="py-2.5 px-3 font-sans">TOTAL 2026 (8 BLN)</td>
                <td class="py-2.5 px-3 text-right">Rp 449.325.263</td>
                <td class="py-2.5 px-3 text-right">Rp 389.124.966</td>
                <td class="py-2.5 px-3 text-right">Rp 41.288.500</td>
                <td class="py-2.5 px-3 text-right">Rp 62.290.500</td>
                <td class="py-2.5 px-3 text-right">Rp 75.637.432</td>
                <td class="py-2.5 px-3 text-right">Rp 92.641.542</td>
                <td class="py-2.5 px-3 text-right text-amber-900 font-extrabold bg-amber-100/60">${{fmtIDR(MASTER_DATA.data_2026.kpi.total_sales)}}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      `;
    }}

    function toggleSort() {{
      appState.sortAsc = !appState.sortAsc;
      updateDashboard();
    }}

    // ========================================================
    // MULTI-YEAR COMPARISON VIEW
    // ========================================================

    function renderMultiYearView() {{
      renderMultiYearPLTable();
      renderMultiYearChart();
      renderMultiYearTable();
    }}

    function renderMultiYearPLTable() {{
      const tbody = document.getElementById('multiyear-pl-tbody');
      tbody.innerHTML = '';
      const summary = MASTER_DATA.profit_loss_comparison.summary;
      const p24 = summary[0];
      const p25 = summary[1];
      const p26 = summary[2];

      const rows = [
        {{ label: 'Pendapatan Bersih (Net Revenue)', v24: p24.revenue, v25: p25.revenue, v26: p26.revenue, format: 'CURR', isBold: true }},
        {{ label: 'Beban Pokok Penjualan (HPP)', v24: p24.cogs, v25: p25.cogs, v26: p26.cogs, format: 'CURR' }},
        {{ label: 'Rasio HPP (%)', v24: p24.cogs_pct, v25: p25.cogs_pct, v26: p26.cogs_pct, format: 'PCT' }},
        {{ label: 'Laba Kotor (Gross Profit)', v24: p24.gross_profit, v25: p25.gross_profit, v26: p26.gross_profit, format: 'CURR', isBold: true }},
        {{ label: 'Gross Profit Margin (%)', v24: p24.gross_margin_pct, v25: p25.gross_margin_pct, v26: p26.gross_margin_pct, format: 'PCT' }},
        {{ label: 'Beban Operasional (Opex)', v24: p24.opex, v25: p25.opex, v26: p26.opex, format: 'CURR' }},
        {{ label: 'Rasio Beban Operasional (%)', v24: p24.opex_pct, v25: p25.opex_pct, v26: p26.opex_pct, format: 'PCT' }},
        {{ label: 'LABA BERSIH (NET PROFIT)', v24: p24.net_profit, v25: p25.net_profit, v26: p26.net_profit, format: 'CURR', isHighlight: true }},
        {{ label: 'Net Profit Margin (NPM %)', v24: p24.net_margin_pct, v25: p25.net_margin_pct, v26: p26.net_margin_pct, format: 'PCT', isBold: true }},
        {{ label: 'Rata-rata Laba Bersih per Bulan', v24: p24.avg_monthly_net_profit, v25: p25.avg_monthly_net_profit, v26: p26.avg_monthly_net_profit, format: 'CURR' }}
      ];

      rows.forEach(r => {{
        const tr = document.createElement('tr');
        if (r.isHighlight) {{
          tr.className = 'bg-emerald-600 text-white font-extrabold text-xs';
        }} else if (r.isBold) {{
          tr.className = 'bg-slate-50 font-bold text-slate-900';
        }} else {{
          tr.className = 'hover:bg-slate-50 text-slate-700';
        }}

        const f24 = r.format === 'CURR' ? fmtIDR(r.v24) : r.v24.toFixed(2) + '%';
        const f25 = r.format === 'CURR' ? fmtIDR(r.v25) : r.v25.toFixed(2) + '%';
        const f26 = r.format === 'CURR' ? fmtIDR(r.v26) : r.v26.toFixed(2) + '%';

        let g25vs24 = '';
        if (r.format === 'CURR') {{
          const diffPct = ((r.v25 - r.v24) / r.v24) * 100;
          g25vs24 = (diffPct >= 0 ? '+' : '') + diffPct.toFixed(2) + '%';
        }} else {{
          const diff = r.v25 - r.v24;
          g25vs24 = (diff >= 0 ? '+' : '') + diff.toFixed(2) + ' pp';
        }}

        let trendEval = '📈 Meningkat';
        if (r.label.includes('LABA BERSIH')) trendEval = '🚀 Melonjak +77,3%';
        if (r.label.includes('HPP')) trendEval = 'Efisiensi Tinggi';
        if (r.label.includes('Rata-rata')) trendEval = '🔥 +78% Rerata';

        tr.innerHTML = `
          <td class="py-2 px-3 font-sans ${{r.isHighlight ? 'text-white' : ''}}">${{r.label}}</td>
          <td class="py-2 px-3 text-right font-mono">${{f24}}</td>
          <td class="py-2 px-3 text-right font-mono">${{f25}}</td>
          <td class="py-2 px-3 text-right font-mono">${{f26}}</td>
          <td class="py-2 px-3 text-right font-mono ${{r.isHighlight ? 'text-emerald-100 font-black' : 'font-bold text-emerald-700'}}">${{g25vs24}}</td>
          <td class="py-2 px-3 text-right font-sans text-[11px]">${{trendEval}}</td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    function renderMultiYearChart() {{
      const svg = document.getElementById('svg-multiyear-chart');
      svg.innerHTML = '';
      const comp = MASTER_DATA.comparison_multiyear.monthly_comparison_8m;

      const allVals = [];
      comp.forEach(c => {{
        allVals.push(c.sales_2024, c.sales_2025, c.sales_2026);
      }});
      const maxVal = Math.max(...allVals, 1);

      const width = 860;
      const height = 280;
      const padLeft = 80;
      const padBottom = 35;
      const padTop = 20;
      const chartW = width - padLeft - 20;
      const chartH = height - padBottom - padTop;
      const count = comp.length;
      const groupW = chartW / count;
      const barW = Math.min(20, (groupW - 14) / 3);

      // Grid lines
      for (let i = 0; i <= 4; i++) {{
        const y = padTop + (chartH / 4) * i;
        const val = maxVal * (1 - i / 4);

        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', padLeft);
        line.setAttribute('x2', width - 20);
        line.setAttribute('y1', y);
        line.setAttribute('y2', y);
        line.setAttribute('stroke', '#e2e8f0');
        line.setAttribute('stroke-dasharray', '3,3');
        svg.appendChild(line);

        const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt.setAttribute('x', padLeft - 8);
        txt.setAttribute('y', y + 3);
        txt.setAttribute('text-anchor', 'end');
        txt.setAttribute('class', 'text-[9px] fill-slate-400 font-mono');
        txt.textContent = 'Rp ' + Math.round(val / 1000000) + ' Jt';
        svg.appendChild(txt);
      }}

      // Grouped Bars
      comp.forEach((m, idx) => {{
        const gX = padLeft + idx * groupW + (groupW - barW * 3 - 6) / 2;

        // 2024 bar
        const h24 = (m.sales_2024 / maxVal) * chartH;
        const r24 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        r24.setAttribute('x', gX);
        r24.setAttribute('y', padTop + chartH - h24);
        r24.setAttribute('width', barW);
        r24.setAttribute('height', h24);
        r24.setAttribute('rx', '3');
        r24.setAttribute('fill', '#94a3b8');
        svg.appendChild(r24);

        // 2025 bar
        const h25 = (m.sales_2025 / maxVal) * chartH;
        const r25 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        r25.setAttribute('x', gX + barW + 3);
        r25.setAttribute('y', padTop + chartH - h25);
        r25.setAttribute('width', barW);
        r25.setAttribute('height', h25);
        r25.setAttribute('rx', '3');
        r25.setAttribute('fill', '#0284c7');
        svg.appendChild(r25);

        // 2026 bar
        const h26 = (m.sales_2026 / maxVal) * chartH;
        const r26 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        r26.setAttribute('x', gX + (barW + 3) * 2);
        r26.setAttribute('y', padTop + chartH - h26);
        r26.setAttribute('width', barW);
        r26.setAttribute('height', h26);
        r26.setAttribute('rx', '3');
        r26.setAttribute('fill', '#059669');
        svg.appendChild(r26);

        // Top value for 2026
        const vTxt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        vTxt.setAttribute('x', gX + (barW + 3) * 2 + barW / 2);
        vTxt.setAttribute('y', padTop + chartH - h26 - 4);
        vTxt.setAttribute('text-anchor', 'middle');
        vTxt.setAttribute('class', 'text-[8px] font-bold fill-emerald-800 font-mono');
        vTxt.textContent = (m.sales_2026 / 1000000).toFixed(0);
        svg.appendChild(vTxt);

        // Month code
        const lbl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        lbl.setAttribute('x', gX + barW * 1.5 + 3);
        lbl.setAttribute('y', height - 12);
        lbl.setAttribute('text-anchor', 'middle');
        lbl.setAttribute('class', 'text-[10px] font-bold fill-slate-700');
        lbl.textContent = m.name;
        svg.appendChild(lbl);
      }});
    }}

    function renderMultiYearTable() {{
      const tbody = document.getElementById('multiyear-tbody');
      tbody.innerHTML = '';
      const comp = MASTER_DATA.comparison_multiyear.monthly_comparison_8m;

      comp.forEach(m => {{
        const g25 = ((m.sales_2026 - m.sales_2025) / m.sales_2025) * 100;
        const g24 = ((m.sales_2026 - m.sales_2024) / m.sales_2024) * 100;

        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition-colors';
        tr.innerHTML = `
          <td class="py-2 px-3 font-semibold text-slate-900 font-sans">${{m.name}}</td>
          <td class="py-2 px-3 text-right font-mono text-slate-500">${{fmtIDR(m.sales_2024)}}</td>
          <td class="py-2 px-3 text-right font-mono text-sky-700">${{fmtIDR(m.sales_2025)}}</td>
          <td class="py-2 px-3 text-right font-mono font-bold text-emerald-700 bg-emerald-50/40">${{fmtIDR(m.sales_2026)}}</td>
          <td class="py-2 px-3 text-right font-mono font-bold text-emerald-600">+${{g25.toFixed(2)}}%</td>
          <td class="py-2 px-3 text-right font-mono font-bold text-sky-600">+${{g24.toFixed(2)}}%</td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    // Product Modal (2024 Detail)
    function openProductModal(rank) {{
      const prod = MASTER_DATA.data_2024.products.find(p => p.rank === rank);
      if (!prod) return;

      document.getElementById('modal-category').innerText = prod.category;
      document.getElementById('modal-title').innerText = prod.name;
      document.getElementById('modal-sales').innerText = fmtIDR(prod.total_sales);
      document.getElementById('modal-qty').innerText = fmtNum(prod.total_qty) + ' ' + prod.unit;
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

    // Export Excel Multi-Sheet (Sales + P&L)
    function exportUnifiedExcel() {{
      if (typeof XLSX === 'undefined') {{
        alert('Library Excel belum selesai dimuat. Harap periksa koneksi internet Anda.');
        return;
      }}

      const wb = XLSX.utils.book_new();

      // Sheet 1: Komparasi Laba Rugi 3 Tahun
      const plCompRows = [
        ['Komponen Laba / Rugi', 'Tahun 2024 (12 Bln)', 'Tahun 2025 (12 Bln)', 'Tahun 2026 (8 Bln)', 'Pertumbuhan 2025 vs 2024 (%)']
      ];
      const p24 = MASTER_DATA.profit_loss_comparison.summary[0];
      const p25 = MASTER_DATA.profit_loss_comparison.summary[1];
      const p26 = MASTER_DATA.profit_loss_comparison.summary[2];

      plCompRows.push(['Pendapatan Bersih (Net Revenue)', p24.revenue, p25.revenue, p26.revenue, ((p25.revenue - p24.revenue)/p24.revenue)]);
      plCompRows.push(['Beban Pokok Penjualan (HPP)', p24.cogs, p25.cogs, p26.cogs, ((p25.cogs - p24.cogs)/p24.cogs)]);
      plCompRows.push(['Laba Kotor (Gross Profit)', p24.gross_profit, p25.gross_profit, p26.gross_profit, ((p25.gross_profit - p24.gross_profit)/p24.gross_profit)]);
      plCompRows.push(['Gross Profit Margin (%)', p24.gross_margin_pct/100, p25.gross_margin_pct/100, p26.gross_margin_pct/100, '']);
      plCompRows.push(['Beban Operasional (Opex)', p24.opex, p25.opex, p26.opex, ((p25.opex - p24.opex)/p24.opex)]);
      plCompRows.push(['LABA BERSIH (NET PROFIT)', p24.net_profit, p25.net_profit, p26.net_profit, ((p25.net_profit - p24.net_profit)/p24.net_profit)]);
      plCompRows.push(['Net Profit Margin (NPM %)', p24.net_margin_pct/100, p25.net_margin_pct/100, p26.net_margin_pct/100, '']);
      plCompRows.push(['Rata-rata Laba Bersih per Bulan', p24.avg_monthly_net_profit, p25.avg_monthly_net_profit, p26.avg_monthly_net_profit, '']);

      const wsPL = XLSX.utils.aoa_to_sheet(plCompRows);
      XLSX.utils.book_append_sheet(wb, wsPL, 'Laba Rugi 3 Tahun');

      // Sheet 2: Komparasi Omzet 3 Tahun (8 Bulan)
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
      XLSX.utils.book_append_sheet(wb, wsComp, 'Komparasi Omzet (8M)');

      // Sheet 3: 2026 Matrix
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

      // Sheet 4: 2025 Matrix
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

      // Sheet 5: 2024 Products
      const p24Rows = [
        ['Rank', 'Nama Barang', 'Kategori', 'Satuan', 'Total Volume', 'Total Omzet (Rp)', 'Porsi (%)']
      ];
      MASTER_DATA.data_2024.products.forEach(p => {{
        p24Rows.push([p.rank, p.name, p.category, p.unit, p.total_qty, p.total_sales, p.pct / 100]);
      }});
      const ws24 = XLSX.utils.aoa_to_sheet(p24Rows);
      XLSX.utils.book_append_sheet(wb, ws24, '129 Produk 2024');

      XLSX.writeFile(wb, 'Laporan_Keuangan_LabaRugi_Snaprint_2024_2026.xlsx');
    }}

    // Initial Load - Default to 2026
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
print("Updated index.html to 3-Year Unified Sales & Profit/Loss Dashboard with Data Change & Save Features!")

# Write 2025/index.html (with relative PDF & SQL paths and default to 2025)
html_2025 = unified_html.replace("selectYear('2026');", "selectYear('2025');")
html_2025 = html_2025.replace('href="Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"', 'href="../Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"')
html_2025 = html_2025.replace('href="Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"', 'href="Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"')
html_2025 = html_2025.replace('href="Laporan_Penjualan_2026_Snaprint_Digital.pdf"', 'href="../2026/Laporan_Penjualan_2026_Snaprint_Digital.pdf"')
html_2025 = html_2025.replace('href="snaprint_sales.sql"', 'href="../snaprint_sales.sql"')
with open("2025/index.html", "w", encoding="utf-8") as f:
    f.write(html_2025)
print("Created 2025/index.html with specialized 2025 defaults!")

# Write 2026/index.html (with relative PDF & SQL paths and default to 2026)
html_2026 = unified_html.replace("selectYear('2026');", "selectYear('2026');")
html_2026 = html_2026.replace('href="Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"', 'href="../Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"')
html_2026 = html_2026.replace('href="Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"', 'href="../2025/Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"')
html_2026 = html_2026.replace('href="Laporan_Penjualan_2026_Snaprint_Digital.pdf"', 'href="Laporan_Penjualan_2026_Snaprint_Digital.pdf"')
html_2026 = html_2026.replace('href="snaprint_sales.sql"', 'href="../snaprint_sales.sql"')
with open("2026/index.html", "w", encoding="utf-8") as f:
    f.write(html_2026)
print("Created 2026/index.html with specialized 2026 defaults!")

# Write artifact
artifact_path = r"C:\Users\Swanto\.gemini\antigravity\brain\186463fa-99d6-481c-8f3f-6ce410e3a005\dashboard_snaprint_unified.html"
with open(artifact_path, "w", encoding="utf-8") as f:
    f.write(unified_html)
print(f"Generated artifact: {artifact_path}")
