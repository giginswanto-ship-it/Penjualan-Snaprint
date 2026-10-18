import glob, re, pypdf
import pandas as pd

files = sorted(glob.glob("penjualan_per_barang_snaprintdigital_*.pdf"))

def parse_num(s):
    s = s.strip().replace(".", "").replace(",", ".")
    return float(s)

monthly_data = {}
all_files_info = []

for f in files:
    reader = pypdf.PdfReader(f)
    lines = []
    for p in reader.pages:
        lines.extend(p.extract_text().splitlines())
    
    m_range = ""
    for l in lines:
        if "Dari " in l and "s/d" in l:
            m_range = l.strip()
            break
            
    month_match = re.search(r"Dari 01 (\w+) 2024", m_range)
    if not month_match:
        all_files_info.append((f, "NON-2024", m_range, 0, 0))
        continue
        
    m = month_match.group(1)
    
    clean_lines = []
    for l in lines:
        l = l.strip()
        if not l or "ACCURATE Accounting System Report" in l or "Tercetak pada" in l or "Halaman " in l:
            continue
        if "Snaprint Digital" in l or "Penjualan per Barang" in l or "Cabang :" in l or "Nama Barang Satuan" in l:
            continue
        clean_lines.append(l)
    
    rows = []
    i = 0
    while i < len(clean_lines):
        line = clean_lines[i]
        if line.startswith("Total Nama Barang"):
            i += 1
            continue
        if i + 1 < len(clean_lines) and clean_lines[i+1].startswith("/319"):
            line = line + " " + clean_lines[i+1]
            i += 1
        match = re.search(r"^(.*?)\s+([\d\.,]+)\s+([\d\.,]+)$", line)
        if match:
            item_raw = match.group(1).strip()
            qty = parse_num(match.group(2))
            sales = parse_num(match.group(3))
            unit = ""
            for u in ["lembar", "PCS", "METER"]:
                if item_raw.endswith(" " + u):
                    unit = u
                    item_raw = item_raw[:-len(u)].strip()
                    break
            rows.append({"month": m, "item": item_raw, "unit": unit, "qty": qty, "sales": sales})
        i += 1
    
    df_m = pd.DataFrame(rows)
    tot_s = df_m['sales'].sum()
    tot_q = df_m['qty'].sum()
    all_files_info.append((f, m, m_range, tot_s, tot_q))

for info in all_files_info:
    print(info)
