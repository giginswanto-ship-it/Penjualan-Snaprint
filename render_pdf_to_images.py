import pymupdf # PyMuPDF
import os

doc = pymupdf.open("Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf")
os.makedirs("pdf_images", exist_ok=True)
print(f"Total pages: {len(doc)}")

for i, page in enumerate(doc):
    pix = page.get_pixmap(dpi=150)
    out_path = f"pdf_images/page_{i+1}.png"
    pix.save(out_path)
    print(f"Saved {out_path} ({pix.width}x{pix.height})")
