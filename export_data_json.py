import json
import pandas as pd
import generate_laporan_2024 as gen

df_all = gen.load_data()
months_order = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']
months_full_name = {
    'Jan': 'Januari',
    'Feb': 'Februari',
    'Mar': 'Maret',
    'Apr': 'April',
    'Mei': 'Mei',
    'Jun': 'Juni',
    'Jul': 'Juli',
    'Agu': 'Agustus',
    'Sep': 'September',
    'Okt': 'Oktober',
    'Nov': 'November',
    'Des': 'Desember'
}

tot_sales_all = float(df_all['sales'].sum())
tot_qty_all = float(df_all['qty'].sum())
avg_monthly_sales = tot_sales_all / 12.0

# 1. Monthly summary
monthly_summary = []
prev_sales = None
for m in months_order:
    sub = df_all[df_all['month'] == m]
    m_qty = float(sub['qty'].sum())
    m_sales = float(sub['sales'].sum())
    pct = (m_sales / tot_sales_all) * 100
    mom_growth = 0.0
    if prev_sales is not None and prev_sales > 0:
        mom_growth = ((m_sales - prev_sales) / prev_sales) * 100
    prev_sales = m_sales
    
    monthly_summary.append({
        "code": m,
        "name": months_full_name[m],
        "sales": m_sales,
        "qty": m_qty,
        "pct": pct,
        "mom_growth": mom_growth,
        "vs_avg": ((m_sales - avg_monthly_sales) / avg_monthly_sales) * 100
    })

# 2. Quarterly & Semester summary
q1_sales = sum(x['sales'] for x in monthly_summary[0:3])
q2_sales = sum(x['sales'] for x in monthly_summary[3:6])
q3_sales = sum(x['sales'] for x in monthly_summary[6:9])
q4_sales = sum(x['sales'] for x in monthly_summary[9:12])

h1_sales = q1_sales + q2_sales
h2_sales = q3_sales + q4_sales

# 3. Product summary with monthly breakdown
prod_group = df_all.groupby('item').agg(
    total_qty=('qty', 'sum'),
    total_sales=('sales', 'sum'),
    unit=('unit', lambda s: [x for x in s if x != ''][0] if any(s != '') else '')
).reset_index().sort_values(by='total_sales', ascending=False)

prod_group['category'] = prod_group['item'].apply(gen.categorize)
prod_group['pct'] = (prod_group['total_sales'] / tot_sales_all) * 100
prod_group['cum_sales'] = prod_group['total_sales'].cumsum()
prod_group['cum_pct'] = (prod_group['cum_sales'] / tot_sales_all) * 100

# Pivot item by month
pivot_sales = df_all.pivot_table(index='item', columns='month', values='sales', aggfunc='sum', fill_value=0).reindex(columns=months_order)
pivot_qty = df_all.pivot_table(index='item', columns='month', values='qty', aggfunc='sum', fill_value=0).reindex(columns=months_order)

products_list = []
rank = 1
for idx, r in prod_group.iterrows():
    item_name = r['item']
    m_sales_dict = {m: float(pivot_sales.loc[item_name, m]) for m in months_order}
    m_qty_dict = {m: float(pivot_qty.loc[item_name, m]) for m in months_order}
    
    # Peak month for this item
    peak_m = max(m_sales_dict, key=m_sales_dict.get)
    peak_m_sales = m_sales_dict[peak_m]
    
    is_pareto_80 = bool(r['cum_pct'] <= 80.5)
    
    products_list.append({
        "rank": rank,
        "name": item_name,
        "category": r['category'],
        "unit": r['unit'] if r['unit'] else "-",
        "total_sales": float(r['total_sales']),
        "total_qty": float(r['total_qty']),
        "pct": float(r['pct']),
        "cum_pct": float(r['cum_pct']),
        "is_pareto_80": is_pareto_80,
        "peak_month": months_full_name[peak_m],
        "peak_month_sales": peak_m_sales,
        "monthly_sales": m_sales_dict,
        "monthly_qty": m_qty_dict
    })
    rank += 1

# 4. Category summary
cat_group = prod_group.groupby('category').agg(
    total_sales=('total_sales', 'sum'),
    total_qty=('total_qty', 'sum'),
    item_count=('item', 'count')
).reset_index().sort_values(by='total_sales', ascending=False)

cat_group['pct'] = (cat_group['total_sales'] / tot_sales_all) * 100

categories_list = []
for idx, r in cat_group.iterrows():
    cat_name = r['category']
    top_items_in_cat = [p['name'] for p in products_list if p['category'] == cat_name][:3]
    categories_list.append({
        "name": cat_name,
        "total_sales": float(r['total_sales']),
        "total_qty": float(r['total_qty']),
        "item_count": int(r['item_count']),
        "pct": float(r['pct']),
        "avg_per_item": float(r['total_sales'] / r['item_count']),
        "top_items": top_items_in_cat
    })

data_payload = {
    "company": "Snaprint Digital Printing",
    "year": 2024,
    "source": "ACCURATE Accounting System",
    "kpi": {
        "total_sales": tot_sales_all,
        "total_qty": tot_qty_all,
        "avg_monthly_sales": avg_monthly_sales,
        "item_count": len(products_list),
        "peak_month": "Agustus",
        "peak_month_sales": 124067870.0,
        "lowest_month": "Januari",
        "lowest_month_sales": 13929250.0,
        "q1_sales": q1_sales,
        "q2_sales": q2_sales,
        "q3_sales": q3_sales,
        "q4_sales": q4_sales,
        "h1_sales": h1_sales,
        "h2_sales": h2_sales,
        "h2_vs_h1_growth": ((h2_sales - h1_sales) / h1_sales) * 100,
        "pareto_80_count": sum(1 for p in products_list if p['is_pareto_80']),
        "pareto_80_pct": (sum(1 for p in products_list if p['is_pareto_80']) / len(products_list)) * 100
    },
    "monthly": monthly_summary,
    "categories": categories_list,
    "products": products_list
}

# Export JSON
with open("sales_data_2024.json", "w", encoding="utf-8") as f:
    json.dump(data_payload, f, ensure_ascii=False, indent=2)

# Export JS (for standalone offline execution without CORS issues)
with open("data_2024.js", "w", encoding="utf-8") as f:
    f.write("const SNAPRINT_DATA_2024 = " + json.dumps(data_payload, ensure_ascii=False) + ";\n")

print("Successfully exported sales_data_2024.json and data_2024.js!")
print(f"Total Sales: Rp {tot_sales_all:,.2f}")
print(f"Products: {len(products_list)}")
print(f"Categories: {len(categories_list)}")
