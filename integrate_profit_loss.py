import json
import os

# 1. Structured Data Laba Rugi 2024
pl_2024 = {
    "year": 2024,
    "period_label": "01 Jan 2024 s/d 31 Des 2024 (12 Bulan Penuh)",
    "months_count": 12,
    "revenue": {
        "gross_sales": 933337771.0,
        "sales_discounts": -340998.0,
        "sales_discounts_idr": -2149513.0,
        "net_revenue": 930847260.0
    },
    "cogs": {
        "total_cogs": 326930387.45,
        "cogs_pct": (326930387.45 / 930847260.0) * 100
    },
    "gross_profit": {
        "amount": 603916872.55,
        "margin_pct": (603916872.55 / 930847260.0) * 100
    },
    "operating_expenses": {
        "total": 280927183.0,
        "pct_of_revenue": (280927183.0 / 930847260.0) * 100,
        "items": [
            {"name": "Beban Gaji, Upah & Honorer", "amount": 222980105.0},
            {"name": "Beban Listrik", "amount": 26529000.0},
            {"name": "Beban Operasional Lainnya", "amount": 12896078.0},
            {"name": "Beban Perlengkapan Kantor", "amount": 5073500.0},
            {"name": "Beban Telekomunikasi", "amount": 4194500.0},
            {"name": "Beban THR", "amount": 2500000.0},
            {"name": "Service / Perbaikan Mesin", "amount": 2089500.0},
            {"name": "Beban Retribusi & Sumbangan", "amount": 1225000.0},
            {"name": "Beban Tunjangan Kesehatan", "amount": 800000.0},
            {"name": "Beban Penyusutan Gedung", "amount": 749000.0},
            {"name": "Beban Penyusutan Inventaris Kantor", "amount": 500000.0},
            {"name": "Beban Perjalanan Dinas", "amount": 495000.0},
            {"name": "Beban Gojek", "amount": 495000.0},
            {"name": "Beban Iklan", "amount": 347500.0},
            {"name": "Beban Bensin, Parkir, Tol Kendaraan", "amount": 238000.0},
            {"name": "Beban Komisi", "amount": 150000.0},
            {"name": "Beban Ekspedisi, Pos & Materai", "amount": 92000.0},
            {"name": "Beban Transportasi Karyawan", "amount": 40000.0},
            {"name": "Beban Katering & Makan Karyawan", "amount": 28000.0}
        ]
    },
    "operating_income": {
        "amount": 322989689.55,
        "margin_pct": (322989689.55 / 930847260.0) * 100
    },
    "non_operating": {
        "income": 0.0,
        "expenses": 147000.0,
        "net": -147000.0
    },
    "net_profit": {
        "amount": 322842689.55,
        "margin_pct": (322842689.55 / 930847260.0) * 100,
        "avg_monthly": 322842689.55 / 12.0
    }
}

# 2. Structured Data Laba Rugi 2025
pl_2025 = {
    "year": 2025,
    "period_label": "01 Jan 2025 s/d 31 Des 2025 (12 Bulan Penuh)",
    "months_count": 12,
    "revenue": {
        "gross_sales": 1381263949.7,
        "returns": -904500.0,
        "sales_discounts": -619521.0,
        "sales_discounts_idr": -788000.0,
        "net_revenue": 1378951928.7
    },
    "cogs": {
        "total_cogs": 478379575.12,
        "cogs_pct": (478379575.12 / 1378951928.7) * 100
    },
    "gross_profit": {
        "amount": 900572353.58,
        "margin_pct": (900572353.58 / 1378951928.7) * 100
    },
    "operating_expenses": {
        "total": 328210670.0,
        "pct_of_revenue": (328210670.0 / 1378951928.7) * 100,
        "items": [
            {"name": "Beban Gaji, Upah & Honorer", "amount": 248899970.0},
            {"name": "Beban Listrik", "amount": 33659000.0},
            {"name": "Beban THR", "amount": 10800000.0},
            {"name": "Beban Operasional Lainnya", "amount": 8921000.0},
            {"name": "Beban Perlengkapan Kantor", "amount": 7425500.0},
            {"name": "Service / Perbaikan Mesin", "amount": 7134000.0},
            {"name": "Beban Telekomunikasi", "amount": 5050700.0},
            {"name": "Beban Perjalanan Dinas", "amount": 4420500.0},
            {"name": "Beban Gojek", "amount": 4420500.0},
            {"name": "Beban Retribusi & Sumbangan", "amount": 1900000.0}
        ]
    },
    "operating_income": {
        "amount": 572361683.58,
        "margin_pct": (572361683.58 / 1378951928.7) * 100
    },
    "non_operating": {
        "income": 1.3,
        "expenses": 30000.0,
        "net": -29998.7
    },
    "net_profit": {
        "amount": 572331684.88,
        "margin_pct": (572331684.88 / 1378951928.7) * 100,
        "avg_monthly": 572331684.88 / 12.0,
        "growth_vs_prev_year_rp": 572331684.88 - 322842689.55,
        "growth_vs_prev_year_pct": ((572331684.88 - 322842689.55) / 322842689.55) * 100
    }
}

# 3. Structured Data Laba Rugi 2026 (Jan - Agu)
pl_2026 = {
    "year": 2026,
    "period_label": "01 Jan 2026 s/d 31 Agu 2026 (8 Bulan Berjalan)",
    "months_count": 8,
    "revenue": {
        "gross_sales": 1166526630.0,
        "sales_discounts": -510213.0,
        "sales_discounts_idr": -15712509.78,
        "net_revenue": 1150303907.22
    },
    "cogs": {
        "total_cogs": 486574898.02,
        "cogs_pct": (486574898.02 / 1150303907.22) * 100
    },
    "gross_profit": {
        "amount": 663729009.20,
        "margin_pct": (663729009.20 / 1150303907.22) * 100
    },
    "operating_expenses": {
        "total": 280463706.0,
        "pct_of_revenue": (280463706.0 / 1150303907.22) * 100,
        "items": [
            {"name": "Beban Gaji, Upah & Honorer", "amount": 195692575.0},
            {"name": "Service / Perbaikan Mesin", "amount": 33302000.0},
            {"name": "Beban Listrik", "amount": 22635000.0},
            {"name": "Beban THR", "amount": 8783331.0},
            {"name": "Beban Perlengkapan Kantor", "amount": 8295300.0},
            {"name": "Beban Perjalanan Dinas", "amount": 5010000.0},
            {"name": "Beban Gojek", "amount": 5010000.0},
            {"name": "Beban Telekomunikasi", "amount": 3298000.0},
            {"name": "Beban Operasional Lainnya", "amount": 2497500.0},
            {"name": "Beban Retribusi & Sumbangan", "amount": 950000.0}
        ]
    },
    "operating_income": {
        "amount": 383265303.20,
        "margin_pct": (383265303.20 / 1150303907.22) * 100
    },
    "non_operating": {
        "income": 0.78,
        "expenses": 0.0,
        "net": 0.78
    },
    "net_profit": {
        "amount": 383265303.98,
        "margin_pct": (383265303.98 / 1150303907.22) * 100,
        "avg_monthly": 383265303.98 / 8.0,
        "projected_full_year": (383265303.98 / 8.0) * 12.0 # ~Rp 574,89 Juta
    }
}

# Save JSON to respective folders
with open("laba_rugi_2024.json", "w", encoding="utf-8") as f:
    json.dump(pl_2024, f, ensure_ascii=False, indent=2)

os.makedirs("2025", exist_ok=True)
with open("2025/laba_rugi_2025.json", "w", encoding="utf-8") as f:
    json.dump(pl_2025, f, ensure_ascii=False, indent=2)

os.makedirs("2026", exist_ok=True)
with open("2026/laba_rugi_2026.json", "w", encoding="utf-8") as f:
    json.dump(pl_2026, f, ensure_ascii=False, indent=2)

# Generate Markdown reports for each year
def fmt_idr(v):
    return f"Rp {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def generate_md(pl, filename):
    rev = pl["revenue"]
    cogs = pl["cogs"]
    gp = pl["gross_profit"]
    opex = pl["operating_expenses"]
    opinc = pl["operating_income"]
    nonop = pl["non_operating"]
    np = pl["net_profit"]

    opex_rows = "\n".join([f"| {item['name']} | {fmt_idr(item['amount'])} | {(item['amount']/rev['net_revenue'])*100:.2f}% |" for item in opex["items"]])

    md = f"""# LAPORAN LABA / RUGI (STANDAR) — ACCURATE ACCOUNTING
## SNAPRINT DIGITAL PRINTING — TAHUN BUKU {pl['year']}
**Periode:** {pl['period_label']}  
**Cabang:** Semua Cabang | **Mata Uang:** Indonesian Rupiah (IDR)  

---

### 1. RINGKASAN STRUKTUR LABA / RUGI {pl['year']}

| Komponen Finansial | Nilai Nominal (Rp) | % Terhadap Pendapatan | Keterangan / Analisis |
| :--- | :---: | :---: | :--- |
| **Jumlah Pendapatan Bersih (Net Revenue)** | **{fmt_idr(rev['net_revenue'])}** | **100,00%** | Penjualan bersih setelah diskon |
| **Beban Pokok Penjualan (HPP / BPP)** | **{fmt_idr(cogs['total_cogs'])}** | **{cogs['cogs_pct']:.2f}%** | Biaya bahan baku & produksi langsung |
| **LABA KOTOR (GROSS PROFIT)** | **{fmt_idr(gp['amount'])}** | **{gp['margin_pct']:.2f}%** | Margin kotor sangat sehat (>57%) |
| **Beban Operasional (Opex)** | **{fmt_idr(opex['total'])}** | **{opex['pct_of_revenue']:.2f}%** | Beban gaji, listrik, mesin, & kantor |
| **PENDAPATAN OPERASIONAL (EBIT)** | **{fmt_idr(opinc['amount'])}** | **{opinc['margin_pct']:.2f}%** | Keuntungan operasional inti bisnis |
| **Pendapatan & Beban Non Operasional** | **{fmt_idr(nonop['net'])}** | **0,00%** | Pendapatan/beban di luar usaha |
| **LABA BERSIH (NET PROFIT)** | 👑 **{fmt_idr(np['amount'])}** | 👑 **{np['margin_pct']:.2f}%** | **Laba bersih riil akhir pembukuan** |

---

### 2. RINCIAN BEBAN OPERASIONAL (OPERATING EXPENSES) {pl['year']}

| Nama Akun Beban Operasional | Nilai Beban (Rp) | % Beban vs Pendapatan |
| :--- | :---: | :---: |
{opex_rows}
| **TOTAL BEBAN OPERASIONAL** | **{fmt_idr(opex['total'])}** | **{opex['pct_of_revenue']:.2f}%** |

---

### 3. METRIK PROFITABILITAS UTAMA
- **Laba Bersih:** **{fmt_idr(np['amount'])}**
- **Net Profit Margin (NPM):** **{np['margin_pct']:.2f}%**
- **Rata-rata Laba Bersih per Bulan:** **{fmt_idr(np['avg_monthly'])} / bulan**
"""
    with open(filename, "w", encoding="utf-8") as f:
        f.write(md)

generate_md(pl_2024, "LAPORAN_LABA_RUGI_2024.md")
generate_md(pl_2025, "2025/LAPORAN_LABA_RUGI_2025.md")
generate_md(pl_2026, "2026/LAPORAN_LABA_RUGI_2026.md")

# Update master dataset
with open("sales_data_master.json", "r", encoding="utf-8") as f:
    master = json.load(f)

master["data_2024"]["profit_loss"] = pl_2024
master["data_2025"]["profit_loss"] = pl_2025
master["data_2026"]["profit_loss"] = pl_2026

# Add multi-year P&L comparison
master["profit_loss_comparison"] = {
    "summary": [
        {
            "year": 2024,
            "period": "12 Bulan",
            "revenue": pl_2024["revenue"]["net_revenue"],
            "cogs": pl_2024["cogs"]["total_cogs"],
            "gross_profit": pl_2024["gross_profit"]["amount"],
            "gross_margin_pct": pl_2024["gross_profit"]["margin_pct"],
            "opex": pl_2024["operating_expenses"]["total"],
            "net_profit": pl_2024["net_profit"]["amount"],
            "net_margin_pct": pl_2024["net_profit"]["margin_pct"],
            "avg_monthly_net_profit": pl_2024["net_profit"]["avg_monthly"]
        },
        {
            "year": 2025,
            "period": "12 Bulan",
            "revenue": pl_2025["revenue"]["net_revenue"],
            "cogs": pl_2025["cogs"]["total_cogs"],
            "gross_profit": pl_2025["gross_profit"]["amount"],
            "gross_margin_pct": pl_2025["gross_profit"]["margin_pct"],
            "opex": pl_2025["operating_expenses"]["total"],
            "net_profit": pl_2025["net_profit"]["amount"],
            "net_margin_pct": pl_2025["net_profit"]["margin_pct"],
            "avg_monthly_net_profit": pl_2025["net_profit"]["avg_monthly"],
            "growth_net_profit_pct": pl_2025["net_profit"]["growth_vs_prev_year_pct"]
        },
        {
            "year": 2026,
            "period": "8 Bulan (Jan-Agu)",
            "revenue": pl_2026["revenue"]["net_revenue"],
            "cogs": pl_2026["cogs"]["total_cogs"],
            "gross_profit": pl_2026["gross_profit"]["amount"],
            "gross_margin_pct": pl_2026["gross_profit"]["margin_pct"],
            "opex": pl_2026["operating_expenses"]["total"],
            "net_profit": pl_2026["net_profit"]["amount"],
            "net_margin_pct": pl_2026["net_profit"]["margin_pct"],
            "avg_monthly_net_profit": pl_2026["net_profit"]["avg_monthly"],
            "projected_full_year_net_profit": pl_2026["net_profit"]["projected_full_year"]
        }
    ]
}

with open("sales_data_master.json", "w", encoding="utf-8") as f:
    json.dump(master, f, ensure_ascii=False, indent=2)

with open("data_multiyear.js", "w", encoding="utf-8") as f:
    f.write("const SNAPRINT_MASTER_DATA = " + json.dumps(master, ensure_ascii=False) + ";\n")

print("Created laba_rugi JSON & Markdown reports in root, 2025/, and 2026/ folders!")
print(f"2024 Laba Bersih: Rp {pl_2024['net_profit']['amount']:,.2f} (Margin: {pl_2024['net_profit']['margin_pct']:.2f}%)")
print(f"2025 Laba Bersih: Rp {pl_2025['net_profit']['amount']:,.2f} (Margin: {pl_2025['net_profit']['margin_pct']:.2f}%) -> +{pl_2025['net_profit']['growth_vs_prev_year_pct']:.2f}%")
print(f"2026 Laba Bersih (8 Bln): Rp {pl_2026['net_profit']['amount']:,.2f} (Margin: {pl_2026['net_profit']['margin_pct']:.2f}%) -> Rerata: Rp {pl_2026['net_profit']['avg_monthly']:,.2f}/bln")
