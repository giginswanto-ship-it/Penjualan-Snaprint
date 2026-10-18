import json
import os
import pandas as pd

# 1. Prepare 2025 Data
monthly_2025_raw = [
    {"code": "Jan", "name": "Januari", "large_format": 42.7, "print_a3": 30.3, "doc_fc": 4.2, "finishing": 6.3, "merchandise": 15.8, "sales": 99.33, "qty": 27800},
    {"code": "Feb", "name": "Februari", "large_format": 38.9, "print_a3": 36.4, "doc_fc": 7.6, "finishing": 5.6, "merchandise": 8.8, "sales": 97.32, "qty": 28400},
    {"code": "Mar", "name": "Maret", "large_format": 41.5, "print_a3": 34.7, "doc_fc": 7.6, "finishing": 6.6, "merchandise": 9.4, "sales": 99.84, "qty": 29100},
    {"code": "Apr", "name": "April", "large_format": 34.8, "print_a3": 26.4, "doc_fc": 7.6, "finishing": 4.6, "merchandise": 10.9, "sales": 84.33, "qty": 23500},
    {"code": "Mei", "name": "Mei", "large_format": 46.2, "print_a3": 33.1, "doc_fc": 11.3, "finishing": 4.7, "merchandise": 14.7, "sales": 109.96, "qty": 31200},
    {"code": "Jun", "name": "Juni", "large_format": 44.4, "print_a3": 41.2, "doc_fc": 8.1, "finishing": 5.5, "merchandise": 10.9, "sales": 110.08, "qty": 30800},
    {"code": "Jul", "name": "Juli", "large_format": 43.5, "print_a3": 44.7, "doc_fc": 10.3, "finishing": 5.3, "merchandise": 15.0, "sales": 118.79, "qty": 33500},
    {"code": "Agu", "name": "Agustus", "large_format": 68.8, "print_a3": 36.8, "doc_fc": 11.8, "finishing": 5.5, "merchandise": 17.2, "sales": 140.0921152, "qty": 41200},
    {"code": "Sep", "name": "September", "large_format": 46.4, "print_a3": 46.5, "doc_fc": 9.8, "finishing": 5.3, "merchandise": 13.6, "sales": 121.62, "qty": 33900},
    {"code": "Okt", "name": "Oktober", "large_format": 52.2, "print_a3": 53.6, "doc_fc": 17.5, "finishing": 5.1, "merchandise": 9.9, "sales": 138.26, "qty": 38100},
    {"code": "Nov", "name": "November", "large_format": 47.8, "print_a3": 54.9, "doc_fc": 4.9, "finishing": 5.4, "merchandise": 15.1, "sales": 128.12, "qty": 34800},
    {"code": "Des", "name": "Desember", "large_format": 40.8, "print_a3": 52.4, "doc_fc": 18.2, "finishing": 4.9, "merchandise": 14.9, "sales": 131.21, "qty": 34984}
]

total_sales_2025 = 1378951928.7
total_qty_2025 = 387284.0
avg_monthly_sales_2025 = total_sales_2025 / 12.0

monthly_2025 = []
prev_s = None
for m in monthly_2025_raw:
    s_val = m['sales'] * 1e6 if m['sales'] < 1000 else m['sales']
    # Specific Agustus exact: 140092115.2
    if m['code'] == 'Agu':
        s_val = 140092115.2
    pct = (s_val / total_sales_2025) * 100
    mom = 0.0
    if prev_s is not None and prev_s > 0:
        mom = ((s_val - prev_s) / prev_s) * 100
    prev_s = s_val
    
    monthly_2025.append({
        "code": m['code'],
        "name": m['name'],
        "sales": s_val,
        "qty": float(m['qty']),
        "pct": pct,
        "mom_growth": mom,
        "vs_avg": ((s_val - avg_monthly_sales_2025) / avg_monthly_sales_2025) * 100,
        "matrix": {
            "large_format": m['large_format'] * 1e6,
            "print_a3": m['print_a3'] * 1e6,
            "doc_fc": m['doc_fc'] * 1e6,
            "finishing": m['finishing'] * 1e6,
            "merchandise": m['merchandise'] * 1e6
        }
    })

categories_2025 = [
    {
        "name": "Large Format (Banner, Flexy, Stiker Outdoor)",
        "short_name": "Large Format & Outdoor",
        "total_sales": 509000000.0,
        "pct": 36.91,
        "items": "Flexy 280 (110, 220, 320), Flexy 340, Flexy 440, Albatros, Ritrama, Standing Roll Banner, Paket X-Banner"
    },
    {
        "name": "Digital Print A3+ & Kertas Lembaran",
        "short_name": "Digital Print A3+ & Sheet",
        "total_sales": 488600000.0,
        "pct": 35.43,
        "items": "Print Color A3 (1 Muka & Bb), Print Color Kn A3, Brosur Digital, Kertas AC 260, AP 150, Linen, BW A3, Stiker Cromo, Quantac, Vinyl"
    },
    {
        "name": "Dokumen Kantor, Fotocopy & Cetak IR",
        "short_name": "Dokumen & Fotocopy",
        "total_sales": 118900000.0,
        "pct": 8.62,
        "items": "Print Color A4, Print BW A4/A3 Mesin IR, Fotocopy A4/Q/Folio/A3, Kertas HVS (A4, Folio, A3), Kertas Pindo, Amplop F4, Scan A4"
    },
    {
        "name": "Jasa Desain, Nota & Percetakan Khusus",
        "short_name": "Jasa Desain & Cetak Umum",
        "total_sales": 118800000.0,
        "pct": 8.61,
        "items": "Paket Desain Grafis, Setting/Design, Cetak Nota Karbon/NCR, Cetak Khusus, Materai Resmi, Stop Map Folio, Press, Ongkir"
    },
    {
        "name": "Merchandise, Kalender, Custom & Promosi",
        "short_name": "Merchandise & Custom",
        "total_sales": 78850000.0,
        "pct": 5.72,
        "items": "Mug Custom, Pin Peniti 5.8cm, Ganci Akrilik, Mousepad, Kaos Dewasa, Print DTF, ID Card, Lanyard, Stempel Flash, Kalender Meja A5"
    },
    {
        "name": "Finishing, Jilid & Pasca Cetak",
        "short_name": "Finishing, Jilid & Laminasi",
        "total_sales": 64801928.7,
        "pct": 4.71,
        "items": "Cutting A3, Cutting Mate, Circle Cutter, Potong K/N, Laminating Tebal & Roll (Doff/Glossy), Jilid Lakban, Ring Kawat, Mika, Mata Ayam"
    }
]

data_2025 = {
    "year": 2025,
    "kpi": {
        "total_sales": total_sales_2025,
        "total_qty": total_qty_2025,
        "avg_monthly_sales": avg_monthly_sales_2025,
        "peak_month": "Agustus",
        "peak_month_sales": 140092115.2,
        "lowest_month": "April",
        "lowest_month_sales": 84330000.0,
        "h1_sales": sum(m['sales'] for m in monthly_2025[0:6]),
        "h2_sales": sum(m['sales'] for m in monthly_2025[6:12]),
        "q1_sales": sum(m['sales'] for m in monthly_2025[0:3]),
        "q2_sales": sum(m['sales'] for m in monthly_2025[3:6]),
        "q3_sales": sum(m['sales'] for m in monthly_2025[6:9]),
        "q4_sales": sum(m['sales'] for m in monthly_2025[9:12]),
    },
    "monthly": monthly_2025,
    "categories": categories_2025
}

# Save 2025 json inside 2025/
os.makedirs("2025", exist_ok=True)
with open("2025/sales_data_2025.json", "w", encoding="utf-8") as f:
    json.dump(data_2025, f, ensure_ascii=False, indent=2)

# Load 2024 data
with open("sales_data_2024.json", "r", encoding="utf-8") as f:
    data_2024 = json.load(f)

# Compute YoY Comparison
sales_2024 = data_2024['kpi']['total_sales']
qty_2024 = data_2024['kpi']['total_qty']

growth_sales = total_sales_2025 - sales_2024
growth_sales_pct = (growth_sales / sales_2024) * 100

growth_qty = total_qty_2025 - qty_2024
growth_qty_pct = (growth_qty / qty_2024) * 100

monthly_yoy = []
for i in range(12):
    m24 = data_2024['monthly'][i]
    m25 = monthly_2025[i]
    diff = m25['sales'] - m24['sales']
    diff_pct = (diff / m24['sales']) * 100 if m24['sales'] > 0 else 0
    monthly_yoy.append({
        "code": m24['code'],
        "name": m24['name'],
        "sales_2024": m24['sales'],
        "sales_2025": m25['sales'],
        "growth_rp": diff,
        "growth_pct": diff_pct,
        "qty_2024": m24['qty'],
        "qty_2025": m25['qty']
    })

yoy_summary = {
    "growth_sales": growth_sales,
    "growth_sales_pct": growth_sales_pct,
    "growth_qty": growth_qty,
    "growth_qty_pct": growth_qty_pct,
    "monthly_comparison": monthly_yoy
}

master_payload = {
    "data_2024": data_2024,
    "data_2025": data_2025,
    "yoy": yoy_summary
}

with open("sales_data_master.json", "w", encoding="utf-8") as f:
    json.dump(master_payload, f, ensure_ascii=False, indent=2)

with open("data_multiyear.js", "w", encoding="utf-8") as f:
    f.write("const SNAPRINT_MASTER_DATA = " + json.dumps(master_payload, ensure_ascii=False) + ";\n")

print("Created 2025/sales_data_2025.json, sales_data_master.json, and data_multiyear.js successfully!")
print(f"2024 Sales: Rp {sales_2024:,.2f}")
print(f"2025 Sales: Rp {total_sales_2025:,.2f}")
print(f"YoY Growth: +{growth_sales_pct:.2f}% (+Rp {growth_sales:,.2f})")
