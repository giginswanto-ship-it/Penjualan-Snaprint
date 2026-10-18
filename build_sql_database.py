import json
import sqlite3
import os

def create_sql_database():
    with open("sales_data_master.json", "r", encoding="utf-8") as f:
        master = json.load(f)

    db_path = "snaprint_sales.db"
    sql_dump_path = "snaprint_sales.sql"

    # Remove existing db if any to ensure clean fresh structure
    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # Enable foreign keys
    cur.execute("PRAGMA foreign_keys = ON;")

    # 1. CREATE TABLES
    cur.executescript("""
    -- ========================================================
    -- DATABASE SKEMA: SNAPRINT DIGITAL PRINTING SALES & PROFIT/LOSS SYSTEM
    -- Terintegrasi Multi-Tahun: 2024, 2025, 2026 (s/d Agustus)
    -- Sumber Data: ACCURATE Accounting System
    -- ========================================================

    DROP TABLE IF EXISTS operating_expenses_breakdown;
    DROP TABLE IF EXISTS profit_loss_summary;
    DROP TABLE IF EXISTS product_monthly_2024;
    DROP TABLE IF EXISTS products_2024;
    DROP TABLE IF EXISTS yoy_comparison_8m;
    DROP TABLE IF EXISTS annual_category_sales;
    DROP TABLE IF EXISTS monthly_category_matrix;
    DROP TABLE IF EXISTS monthly_sales;
    DROP TABLE IF EXISTS master_categories;
    DROP TABLE IF EXISTS master_years;
    DROP TABLE IF EXISTS company_profile;

    -- Profil Usaha
    CREATE TABLE company_profile (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_name TEXT NOT NULL,
        branch TEXT NOT NULL,
        system_source TEXT NOT NULL,
        report_currency TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    -- Master Tahun Buku
    CREATE TABLE master_years (
        year INTEGER PRIMARY KEY,
        period_label TEXT NOT NULL,
        months_count INTEGER NOT NULL,
        total_sales REAL NOT NULL,
        total_qty REAL NOT NULL,
        avg_monthly_sales REAL NOT NULL,
        peak_month TEXT NOT NULL,
        peak_month_sales REAL NOT NULL,
        lowest_month TEXT NOT NULL,
        lowest_month_sales REAL NOT NULL,
        q1_sales REAL,
        q2_sales REAL,
        q3_sales REAL,
        q4_sales REAL,
        h1_sales REAL,
        h2_sales REAL
    );

    -- Master Kategori Produk & Layanan
    CREATE TABLE master_categories (
        category_id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_name TEXT UNIQUE NOT NULL,
        short_name TEXT NOT NULL,
        description_items TEXT
    );

    -- Rekap Penjualan Bulanan (Semua Tahun)
    CREATE TABLE monthly_sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER NOT NULL,
        month_num INTEGER NOT NULL,
        month_code TEXT NOT NULL,
        month_name TEXT NOT NULL,
        sales_amount REAL NOT NULL,
        quantity REAL NOT NULL,
        pct_of_annual REAL NOT NULL,
        mom_growth_pct REAL NOT NULL,
        vs_avg_pct REAL NOT NULL,
        is_peak INTEGER DEFAULT 0,
        FOREIGN KEY (year) REFERENCES master_years(year)
    );

    -- Matriks Penjualan Bulanan per Lini Kategori Produk
    CREATE TABLE monthly_category_matrix (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER NOT NULL,
        month_code TEXT NOT NULL,
        month_name TEXT NOT NULL,
        large_format REAL DEFAULT 0,
        digital_print_a3 REAL DEFAULT 0,
        dokumen_fotocopy REAL DEFAULT 0,
        desain_khusus REAL DEFAULT 0,
        merchandise_promosi REAL DEFAULT 0,
        finishing_jilid REAL DEFAULT 0,
        total_sales REAL NOT NULL,
        FOREIGN KEY (year) REFERENCES master_years(year)
    );

    -- Rekapitulasi Tahunan per Kategori Bisnis
    CREATE TABLE annual_category_sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER NOT NULL,
        category_name TEXT NOT NULL,
        total_sales REAL NOT NULL,
        pct_of_year REAL NOT NULL,
        items_summary TEXT,
        FOREIGN KEY (year) REFERENCES master_years(year)
    );

    -- Rincian 129 Produk Spesifik Tahun 2024
    CREATE TABLE products_2024 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rank_no INTEGER NOT NULL,
        name TEXT NOT NULL,
        category TEXT NOT NULL,
        unit TEXT,
        total_qty REAL NOT NULL,
        total_sales REAL NOT NULL,
        pct_of_2024 REAL NOT NULL,
        is_pareto_80 INTEGER NOT NULL DEFAULT 0
    );

    -- Riwayat 12 Bulan per Produk Tahun 2024
    CREATE TABLE product_monthly_2024 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER NOT NULL,
        product_name TEXT NOT NULL,
        month_code TEXT NOT NULL,
        month_num INTEGER NOT NULL,
        qty REAL NOT NULL,
        sales REAL NOT NULL,
        FOREIGN KEY (product_id) REFERENCES products_2024(id)
    );

    -- Komparasi Apel-ke-Apel Multi-Tahun Periode 8 Bulan Pertama (Jan - Agu)
    CREATE TABLE yoy_comparison_8m (
        month_num INTEGER PRIMARY KEY,
        month_code TEXT NOT NULL,
        month_name TEXT NOT NULL,
        sales_2024 REAL NOT NULL,
        sales_2025 REAL NOT NULL,
        sales_2026 REAL NOT NULL,
        qty_2024 REAL NOT NULL,
        qty_2025 REAL NOT NULL,
        qty_2026 REAL NOT NULL,
        growth_26_vs_25_rp REAL NOT NULL,
        growth_26_vs_25_pct REAL NOT NULL,
        growth_26_vs_24_rp REAL NOT NULL,
        growth_26_vs_24_pct REAL NOT NULL
    );

    -- TABEL LABA / RUGI TAHUNAN (ACCURATE INCOME STATEMENT)
    CREATE TABLE profit_loss_summary (
        year INTEGER PRIMARY KEY,
        period_label TEXT NOT NULL,
        months_count INTEGER NOT NULL,
        gross_sales REAL NOT NULL,
        discounts REAL NOT NULL,
        net_revenue REAL NOT NULL,
        cogs REAL NOT NULL,
        cogs_pct REAL NOT NULL,
        gross_profit REAL NOT NULL,
        gross_margin_pct REAL NOT NULL,
        operating_expenses REAL NOT NULL,
        opex_pct REAL NOT NULL,
        operating_income REAL NOT NULL,
        operating_margin_pct REAL NOT NULL,
        non_operating_net REAL NOT NULL,
        net_profit REAL NOT NULL,
        net_profit_margin_pct REAL NOT NULL,
        avg_monthly_net_profit REAL NOT NULL,
        FOREIGN KEY (year) REFERENCES master_years(year)
    );

    -- TABEL RINCIAN BEBAN OPERASIONAL
    CREATE TABLE operating_expenses_breakdown (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER NOT NULL,
        expense_name TEXT NOT NULL,
        amount REAL NOT NULL,
        pct_of_revenue REAL NOT NULL,
        FOREIGN KEY (year) REFERENCES master_years(year)
    );
    """)

    # 2. INSERT DATA

    # Company Profile
    cur.execute("""
    INSERT INTO company_profile (company_name, branch, system_source, report_currency)
    VALUES ('SNAPRINT DIGITAL PRINTING', 'Semua Cabang', 'ACCURATE Accounting System', 'IDR (Rupiah)');
    """)

    # Master Categories
    categories = [
        ("Large Format (Banner, Flexy, Stiker Outdoor)", "Large Format & Outdoor", "Flexy 280 (110, 220, 320), Flexy 340 (110, 210, 320), Flexy 440 (110, 220, 320), Bahan Albatros, Sticker Ritrama Glossy (125), Standing Roll Banner (60x160, 85x200), Paket X-Banner Outdoor & Indoor, X-Banner Fiber."),
        ("Digital Print A3+ & Kertas Lembaran", "Digital Print A3+ & Sheet", "Print Color A3 (1 Muka & Bb), Print Color Kn A3 (4 Klik), Brosur Digital, Kertas AC 260, AP 150 (32.5x48 cm / A3+), Linen 210, Kertas BW A3, Stiker Cromo, Stiker Quantac (Glossy, Doff, Transparan), Sticker Vinyl & Oneway."),
        ("Finishing, Jilid & Pasca Cetak", "Finishing & Jilid", "Cutting A3, Ongkos Potong K/N, Laminating Tebal (A3, Folio/A4), Laminating Roll (Doff & Glossy), Jilid Lakban (100, 200, >300 lbr), Ring Kawat (No. 5, 6, 8, 10, 12, 14, 16), Mika Bening/Buram, Mata Ayam, Jilid Soft Cover."),
        ("Merchandise, Kalender, Custom & Promosi", "Merchandise & Promosi", "Mug Custom, Pin Peniti 5.8 cm, Custom Kaos Dewasa, Print DTF (Meteran, A3, A4), Paket ID Card & Print ID Card, Cetak Lanyard, Stempel Flash & Tinta Stempel, Kalender Meja A5 & Dudukan Kalender, Notebook Custom, Jam Dinding Custom."),
        ("Jasa Desain, Nota & Percetakan Khusus", "Desain & Cetak Khusus", "Paket Desain Grafis, Setting File Desain (min. 15), Cetak Nota Karbon/NCR, Cetak Dokumen Khusus/Lain-lain, Box Kartu Nama Polos, Stop Map Folio, Biaya Press, Offset Percetakan, Ongkos Kirim."),
        ("Dokumen Kantor, Fotocopy & Cetak IR", "Dokumen & Cetak IR", "Print Color A4 (1 Muka & Bb), Print BW A4 (Mesin IR 1 Muka & Bb), Print BW A3 (Mesin IR 1 Muka & Bb), Fotocopy A4/Q/Folio, Fotocopy A3, Kertas HVS (A4, Folio, A3 80 gr, 100 gr), Kertas Pindo Folio Warna, Amplop F4/Jaya, Scan A4.")
    ]
    cur.executemany("INSERT INTO master_categories (category_name, short_name, description_items) VALUES (?, ?, ?);", categories)

    # Master Years
    d24 = master["data_2024"]
    d25 = master["data_2025"]
    d26 = master["data_2026"]

    years_data = [
        (
            2024,
            "01 Januari 2024 s/d 31 Desember 2024",
            12,
            d24["kpi"]["total_sales"],
            d24["kpi"]["total_qty"],
            d24["kpi"]["avg_monthly_sales"],
            d24["kpi"]["peak_month"],
            d24["kpi"]["peak_month_sales"],
            d24["kpi"]["lowest_month"],
            d24["kpi"]["lowest_month_sales"],
            d24["kpi"]["q1_sales"],
            d24["kpi"]["q2_sales"],
            d24["kpi"]["q3_sales"],
            d24["kpi"]["q4_sales"],
            d24["kpi"]["h1_sales"],
            d24["kpi"]["h2_sales"]
        ),
        (
            2025,
            "01 Januari 2025 s/d 31 Desember 2025",
            12,
            d25["kpi"]["total_sales"],
            d25["kpi"]["total_qty"],
            d25["kpi"]["avg_monthly_sales"],
            d25["kpi"]["peak_month"],
            d25["kpi"]["peak_month_sales"],
            d25["kpi"]["lowest_month"],
            d25["kpi"]["lowest_month_sales"],
            d25["kpi"]["q1_sales"],
            d25["kpi"]["q2_sales"],
            d25["kpi"]["q3_sales"],
            d25["kpi"]["q4_sales"],
            d25["kpi"]["h1_sales"],
            d25["kpi"]["h2_sales"]
        ),
        (
            2026,
            "01 Januari 2026 s/d 31 Agustus 2026 (8 Bulan)",
            8,
            d26["kpi"]["total_sales"],
            d26["kpi"]["total_qty"],
            d26["kpi"]["avg_monthly_sales"],
            d26["kpi"]["peak_month"],
            d26["kpi"]["peak_month_sales"],
            d26["kpi"]["lowest_month"],
            d26["kpi"]["lowest_month_sales"],
            d26["kpi"]["q1_sales"],
            d26["kpi"]["q2_sales"],
            d26["kpi"]["q3_partial_sales"],
            0.0,
            d26["kpi"]["h1_sales"],
            0.0
        )
    ]
    cur.executemany("""
    INSERT INTO master_years (
        year, period_label, months_count, total_sales, total_qty, avg_monthly_sales,
        peak_month, peak_month_sales, lowest_month, lowest_month_sales,
        q1_sales, q2_sales, q3_sales, q4_sales, h1_sales, h2_sales
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    """, years_data)

    # Monthly Sales for all 3 years
    month_map = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'Mei': 5, 'Jun': 6, 'Jul': 7, 'Agu': 8, 'Sep': 9, 'Okt': 10, 'Nov': 11, 'Des': 12}
    
    for yr_key, yr_data in [(2024, d24), (2025, d25), (2026, d26)]:
        for m in yr_data["monthly"]:
            m_num = month_map[m["code"]]
            is_peak = 1 if (m["name"].startswith(yr_data["kpi"]["peak_month"])) else 0
            cur.execute("""
            INSERT INTO monthly_sales (
                year, month_num, month_code, month_name, sales_amount, quantity,
                pct_of_annual, mom_growth_pct, vs_avg_pct, is_peak
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (
                yr_key, m_num, m["code"], m["name"], m["sales"], m["qty"],
                m["pct"], m["mom_growth"], m["vs_avg"], is_peak
            ))

    # Monthly Category Matrix for 2025 & 2026
    for m in d25["monthly"]:
        mat = m.get("matrix", {})
        cur.execute("""
        INSERT INTO monthly_category_matrix (
            year, month_code, month_name, large_format, digital_print_a3,
            dokumen_fotocopy, desain_khusus, merchandise_promosi, finishing_jilid, total_sales
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            2025, m["code"], m["name"],
            mat.get("large_format", 0), mat.get("print_a3", 0), mat.get("doc_fc", 0),
            0.0, mat.get("merchandise", 0), mat.get("finishing", 0), m["sales"]
        ))

    for m in d26["monthly"]:
        mat = m.get("matrix", {})
        cur.execute("""
        INSERT INTO monthly_category_matrix (
            year, month_code, month_name, large_format, digital_print_a3,
            dokumen_fotocopy, desain_khusus, merchandise_promosi, finishing_jilid, total_sales
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            2026, m["code"], m["name"],
            mat.get("large_format", 0), mat.get("print_a3", 0), mat.get("doc_fc", 0),
            mat.get("desain_khusus", 0), mat.get("merchandise", 0), mat.get("finishing", 0), m["sales"]
        ))

    # Annual Category Summary (2024, 2025, 2026)
    for yr_key, yr_data in [(2024, d24), (2025, d25), (2026, d26)]:
        for cat in yr_data["categories"]:
            cur.execute("""
            INSERT INTO annual_category_sales (year, category_name, total_sales, pct_of_year, items_summary)
            VALUES (?, ?, ?, ?, ?);
            """, (yr_key, cat["name"], cat["total_sales"], cat["pct"], cat.get("items", "")))

    # Products 2024 (129 products)
    for p in d24["products"]:
        cur.execute("""
        INSERT INTO products_2024 (rank_no, name, category, unit, total_qty, total_sales, pct_of_2024, is_pareto_80)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            p["rank"], p["name"], p["category"], p["unit"],
            p["total_qty"], p["total_sales"], p["pct"], 1 if p["is_pareto_80"] else 0
        ))
        p_id = cur.lastrowid

        # Insert 12 months for this product
        for m_code in ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']:
            s_val = p["monthly_sales"].get(m_code, 0.0)
            q_val = p["monthly_qty"].get(m_code, 0.0)
            cur.execute("""
            INSERT INTO product_monthly_2024 (product_id, product_name, month_code, month_num, qty, sales)
            VALUES (?, ?, ?, ?, ?, ?);
            """, (p_id, p["name"], m_code, month_map[m_code], q_val, s_val))

    # YoY 8-Month Comparison
    comp_8m = master["comparison_multiyear"]["monthly_comparison_8m"]
    for row in comp_8m:
        cur.execute("""
        INSERT INTO yoy_comparison_8m (
            month_num, month_code, month_name, sales_2024, sales_2025, sales_2026,
            qty_2024, qty_2025, qty_2026, growth_26_vs_25_rp, growth_26_vs_25_pct,
            growth_26_vs_24_rp, growth_26_vs_24_pct
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            month_map[row["code"]], row["code"], row["name"],
            row["sales_2024"], row["sales_2025"], row["sales_2026"],
            row["qty_2024"], row["qty_2025"], row["qty_2026"],
            row["growth_26_vs_25_rp"], row["growth_26_vs_25_pct"],
            row["sales_2026"] - row["sales_2024"],
            ((row["sales_2026"] - row["sales_2024"]) / row["sales_2024"]) * 100
        ))

    # INSERT PROFIT & LOSS DATA (2024, 2025, 2026)
    with open("laba_rugi_2024.json", "r", encoding="utf-8") as f:
        pl24 = json.load(f)
    with open("2025/laba_rugi_2025.json", "r", encoding="utf-8") as f:
        pl25 = json.load(f)
    with open("2026/laba_rugi_2026.json", "r", encoding="utf-8") as f:
        pl26 = json.load(f)

    for pl in [pl24, pl25, pl26]:
        cur.execute("""
        INSERT INTO profit_loss_summary (
            year, period_label, months_count, gross_sales, discounts, net_revenue,
            cogs, cogs_pct, gross_profit, gross_margin_pct, operating_expenses, opex_pct,
            operating_income, operating_margin_pct, non_operating_net,
            net_profit, net_profit_margin_pct, avg_monthly_net_profit
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            pl["year"], pl["period_label"], pl["months_count"],
            pl["revenue"]["gross_sales"],
            pl["revenue"].get("sales_discounts", 0) + pl["revenue"].get("sales_discounts_idr", 0) + pl["revenue"].get("returns", 0),
            pl["revenue"]["net_revenue"],
            pl["cogs"]["total_cogs"], pl["cogs"]["cogs_pct"],
            pl["gross_profit"]["amount"], pl["gross_profit"]["margin_pct"],
            pl["operating_expenses"]["total"], pl["operating_expenses"]["pct_of_revenue"],
            pl["operating_income"]["amount"], pl["operating_income"]["margin_pct"],
            pl["non_operating"]["net"],
            pl["net_profit"]["amount"], pl["net_profit"]["margin_pct"],
            pl["net_profit"]["avg_monthly"]
        ))

        # Insert Operating Expenses Breakdown
        for item in pl["operating_expenses"]["items"]:
            cur.execute("""
            INSERT INTO operating_expenses_breakdown (year, expense_name, amount, pct_of_revenue)
            VALUES (?, ?, ?, ?);
            """, (pl["year"], item["name"], item["amount"], (item["amount"] / pl["revenue"]["net_revenue"]) * 100))

    # 3. CREATE CONVENIENT SQL VIEWS
    cur.executescript("""
    -- View Ringkasan Kinerja Penjualan
    CREATE VIEW v_annual_performance AS
    SELECT 
        year AS Tahun,
        period_label AS Periode,
        months_count AS Jumlah_Bulan,
        'Rp ' || printf('%,d', CAST(total_sales AS INT)) AS Total_Omzet_Rp,
        printf('%,d', CAST(total_qty AS INT)) AS Total_Volume_Order,
        'Rp ' || printf('%,d', CAST(avg_monthly_sales AS INT)) AS Rerata_Bulanan_Rp,
        peak_month || ' (Rp ' || printf('%,d', CAST(peak_month_sales AS INT)) || ')' AS Bulan_Puncak
    FROM master_years
    ORDER BY year;

    -- VIEW RINGKASAN LABA / RUGI RESMI
    CREATE VIEW v_profit_and_loss_summary AS
    SELECT 
        year AS Tahun,
        period_label AS Periode,
        'Rp ' || printf('%,d', CAST(net_revenue AS INT)) AS Pendapatan_Bersih_Rp,
        'Rp ' || printf('%,d', CAST(cogs AS INT)) AS Beban_Pokok_HPP_Rp,
        printf('%.2f%%', cogs_pct) AS HPP_Pct,
        'Rp ' || printf('%,d', CAST(gross_profit AS INT)) AS Laba_Kotor_Rp,
        printf('%.2f%%', gross_margin_pct) AS Gross_Margin,
        'Rp ' || printf('%,d', CAST(operating_expenses AS INT)) AS Beban_Operasional_Rp,
        'Rp ' || printf('%,d', CAST(net_profit AS INT)) AS LABA_BERSIH_Rp,
        printf('%.2f%%', net_profit_margin_pct) AS Net_Profit_Margin,
        'Rp ' || printf('%,d', CAST(avg_monthly_net_profit AS INT)) AS Laba_Bersih_Bulanan_Rp
    FROM profit_loss_summary
    ORDER BY year;

    -- View Komparasi Bulanan 3 Tahun (Jan - Agu)
    CREATE VIEW v_comparison_3_years_8m AS
    SELECT 
        month_num AS No_Bulan,
        month_name AS Bulan,
        'Rp ' || printf('%,d', CAST(sales_2024 AS INT)) AS Omzet_2024,
        'Rp ' || printf('%,d', CAST(sales_2025 AS INT)) AS Omzet_2025,
        'Rp ' || printf('%,d', CAST(sales_2026 AS INT)) AS Omzet_2026,
        printf('%+.2f%%', growth_26_vs_25_pct) AS Pertumbuhan_2026_vs_2025,
        printf('%+.2f%%', growth_26_vs_24_pct) AS Pertumbuhan_2026_vs_2024
    FROM yoy_comparison_8m
    ORDER BY month_num;

    -- View Kategori Penjualan Per Tahun
    CREATE VIEW v_category_summary AS
    SELECT 
        year AS Tahun,
        category_name AS Kategori,
        'Rp ' || printf('%,d', CAST(total_sales AS INT)) AS Total_Omzet_Rp,
        printf('%.2f%%', pct_of_year) AS Porsi
    FROM annual_category_sales
    ORDER BY year DESC, total_sales DESC;

    -- View Ranking Beban Operasional Terbesar Per Tahun
    CREATE VIEW v_operating_expenses_ranking AS
    SELECT 
        year AS Tahun,
        expense_name AS Nama_Akun_Beban,
        'Rp ' || printf('%,d', CAST(amount AS INT)) AS Nilai_Beban_Rp,
        printf('%.2f%%', pct_of_revenue) AS Porsi_vs_Omzet
    FROM operating_expenses_breakdown
    ORDER BY year DESC, amount DESC;

    -- View 24 Produk Core 80% Tahun 2024 (Pareto)
    CREATE VIEW v_pareto_core80_2024 AS
    SELECT 
        rank_no AS Peringkat,
        name AS Nama_Barang,
        category AS Kategori,
        unit AS Satuan,
        printf('%,d', CAST(total_qty AS INT)) AS Volume_Terjual,
        'Rp ' || printf('%,d', CAST(total_sales AS INT)) AS Total_Omzet_Rp,
        printf('%.2f%%', pct_of_2024) AS Porsi_Kontribusi
    FROM products_2024
    WHERE is_pareto_80 = 1
    ORDER BY rank_no;
    """)

    conn.commit()

    # 4. EXPORT TO FULL SQL SCRIPT (.sql)
    with open(sql_dump_path, "w", encoding="utf-8") as f:
        for line in conn.iterdump():
            f.write(f"{line}\n")

    conn.close()

    print(f"SQL Database successfully created: {db_path} ({os.path.getsize(db_path)} bytes)")
    print(f"SQL Dump file successfully created: {sql_dump_path} ({os.path.getsize(sql_dump_path)} bytes)")

if __name__ == "__main__":
    create_sql_database()
