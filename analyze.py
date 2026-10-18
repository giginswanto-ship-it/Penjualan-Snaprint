import pandas as pd
import generate_laporan_2024 as gen

df = gen.load_data()
tot_sales = df['sales'].sum()
tot_qty = df['qty'].sum()

# Monthly & Quarterly
m_order = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']
monthly = df.groupby('month').agg(sales=('sales', 'sum'), qty=('qty', 'sum')).reindex(m_order)
monthly['pct'] = monthly['sales'] / tot_sales * 100

q1 = monthly.loc[['Jan', 'Feb', 'Mar'], 'sales'].sum()
q2 = monthly.loc[['Apr', 'Mei', 'Jun'], 'sales'].sum()
q3 = monthly.loc[['Jul', 'Agu', 'Sep'], 'sales'].sum()
q4 = monthly.loc[['Okt', 'Nov', 'Des'], 'sales'].sum()

h1 = q1 + q2
h2 = q3 + q4

# Product Pareto
prod = df.groupby('item').agg(
    sales=('sales', 'sum'),
    qty=('qty', 'sum'),
    unit=('unit', lambda s: [x for x in s if x != ''][0] if any(s != '') else '')
).reset_index().sort_values(by='sales', ascending=False)
prod['pct'] = prod['sales'] / tot_sales * 100
prod['cum_pct'] = prod['sales'].cumsum() / tot_sales * 100

top10_sales = prod.head(10)['sales'].sum()
top20_sales = prod.head(20)['sales'].sum()
p80_items = prod[prod['cum_pct'] <= 80]

print(f"Total Sales: Rp {tot_sales:,.2f}")
print(f"Total Qty: {tot_qty:,.2f}")
print(f"Q1: Rp {q1:,.2f} ({q1/tot_sales*100:.2f}%)")
print(f"Q2: Rp {q2:,.2f} ({q2/tot_sales*100:.2f}%)")
print(f"Q3: Rp {q3:,.2f} ({q3/tot_sales*100:.2f}%)")
print(f"Q4: Rp {q4:,.2f} ({q4/tot_sales*100:.2f}%)")
print(f"H1: Rp {h1:,.2f} ({h1/tot_sales*100:.2f}%)")
print(f"H2: Rp {h2:,.2f} ({h2/tot_sales*100:.2f}%) [Growth H2 vs H1: {(h2-h1)/h1*100:+.2f}%]")
print(f"Top 10 items sales: Rp {top10_sales:,.2f} ({top10_sales/tot_sales*100:.2f}%)")
print(f"Top 20 items sales: Rp {top20_sales:,.2f} ({top20_sales/tot_sales*100:.2f}%)")
print(f"Items contributing to 80% sales: {len(p80_items)} of {len(prod)} items ({len(p80_items)/len(prod)*100:.1f}%)")

print("\n--- TOP 10 ITEMS ---")
for i, r in prod.head(10).iterrows():
    print(f"- {r['item']} ({r['unit']}): Rp {r['sales']:,.2f} ({r['pct']:.2f}%) | Qty: {r['qty']:,.2f}")

print("\n--- CATEGORY BREAKDOWN ---")
prod['cat'] = prod['item'].apply(gen.categorize)
cats = prod.groupby('cat').agg(sales=('sales', 'sum'), qty=('qty', 'sum'), items=('item', 'count')).reset_index().sort_values(by='sales', ascending=False)
cats['pct'] = cats['sales'] / tot_sales * 100
for i, r in cats.iterrows():
    print(f"- {r['cat']}: Rp {r['sales']:,.2f} ({r['pct']:.2f}%) | {r['items']} items | Qty: {r['qty']:,.2f}")
