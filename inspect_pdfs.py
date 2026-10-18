import glob, re, pypdf

files = sorted(glob.glob("penjualan_per_barang_snaprintdigital_*.pdf"))
print(f"Total files found: {len(files)}")

for f in files:
    try:
        reader = pypdf.PdfReader(f)
        text = ""
        for p in reader.pages:
            text += p.extract_text() + "\n"
        
        m_range = ""
        for l in text.splitlines():
            if "Dari " in l and "s/d" in l:
                m_range = l.strip()
                break
        print(f"{f}: {m_range} (pages: {len(reader.pages)})")
    except Exception as e:
        print(f"{f}: Error {e}")
