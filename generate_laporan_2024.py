import glob
import re
import os
import pypdf
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas

def parse_num(s):
    s = s.strip().replace(".", "").replace(",", ".")
    return float(s)

def load_data():
    files = sorted(glob.glob("penjualan_per_barang_snaprintdigital_*.pdf"))
    all_rows = []
    seen_months = set()
    
    # Preferred order of months
    months_order = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']
    
    # Map files to month
    parsed_files = []
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
        if month_match:
            m_name = month_match.group(1)
            parsed_files.append((m_name, f, lines))

    for m_target in months_order:
        # Pick the first matching file for each month
        found = False
        for m_name, f, lines in parsed_files:
            if m_name == m_target and m_name not in seen_months:
                seen_months.add(m_name)
                found = True
                
                clean_lines = []
                for l in lines:
                    l = l.strip()
                    if not l or "ACCURATE Accounting System Report" in l or "Tercetak pada" in l or "Halaman " in l:
                        continue
                    if "Snaprint Digital" in l or "Penjualan per Barang" in l or "Cabang :" in l or "Nama Barang Satuan" in l:
                        continue
                    clean_lines.append(l)

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
                        all_rows.append({
                            "month": m_name,
                            "item": item_raw,
                            "unit": unit,
                            "qty": qty,
                            "sales": sales
                        })
                    i += 1
                break
                
    df = pd.DataFrame(all_rows)
    return df

def categorize(item):
    item_l = item.lower()
    if 'flexy' in item_l or 'banner' in item_l or 'ritrama' in item_l or 'albatros' in item_l:
        return 'Large Format & Outdoor Banner'
    elif 'print' in item_l or 'fotocopy' in item_l or 'scaning' in item_l:
        return 'Digital Print & Fotocopy Dokumen'
    elif any(k in item_l for k in ['quantac', 'cromo', 'ac 260', 'ap 150', 'linen', 'hvs', 'pindo', 'kertas bw', 'kertas ncr']):
        return 'Kertas, Sticker Sheet & Bahan Cetak'
    elif any(k in item_l for k in ['laminating', 'jilid', 'ring kawat', 'cutting', 'potong', 'finishing', 'press', 'streples']):
        return 'Finishing, Jilid & Laminasi'
    elif any(k in item_l for k in ['kaos', 'mug', 'ganci', 'stampel', 'notebook', 'tumbler', 'totebag', 'pin ', 'kalender', 'name tag', 'kipas', 'mouse pad', 'jam dinding', 'ballpoint']):
        return 'Merchandise & Produk Custom'
    else:
        return 'Jasa Desain, Percetakan Lain & Perlengkapan'

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_footer(num_pages)
            super().showPage()
        super().save()

    def draw_footer(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 7.5)
        self.setFillColor(colors.HexColor('#64748B'))
        
        # Subtle horizontal rule above footer
        self.setStrokeColor(colors.HexColor('#E2E8F0'))
        self.setLineWidth(0.5)
        self.line(30, 22, 595.27 - 30, 22)
        
        # Footer text
        left_text = "Snaprint Digital Printing — Laporan Penjualan Tahunan 2024"
        right_text = f"Halaman {self._pageNumber} dari {page_count}"
        self.drawString(30, 12, left_text)
        self.drawRightString(595.27 - 30, 12, right_text)
        self.restoreState()

def generate_pdf():
    df_all = load_data()
    tot_sales_all = df_all['sales'].sum()
    tot_qty_all = df_all['qty'].sum()
    avg_sales_monthly = tot_sales_all / 12.0
    
    prod_summary = df_all.groupby('item').agg(
        total_qty=('qty', 'sum'),
        total_sales=('sales', 'sum'),
        unit=('unit', lambda s: [x for x in s if x != ''][0] if any(s != '') else '')
    ).reset_index()
    prod_summary['pct'] = (prod_summary['total_sales'] / tot_sales_all) * 100
    prod_summary = prod_summary.sort_values(by='total_sales', ascending=False)
    prod_summary['category'] = prod_summary['item'].apply(categorize)

    cat_summary = prod_summary.groupby('category').agg(
        total_sales=('total_sales', 'sum'),
        total_qty=('total_qty', 'sum'),
        item_count=('item', 'count')
    ).reset_index().sort_values(by='total_sales', ascending=False)
    cat_summary['pct'] = (cat_summary['total_sales'] / tot_sales_all) * 100

    pdf_filename = "Laporan_Penjualan_Tahunan_2024_Snaprint_Digital.pdf"
    
    # Margins: 30pt (left, right), 28pt (top), 30pt (bottom)
    # Printable width: 595.27 - 60 = 535.27 pt
    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=A4,
        leftMargin=30,
        rightMargin=30,
        topMargin=26,
        bottomMargin=30
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'DocTitle',
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=18,
        textColor=colors.HexColor('#0F172A'),
        alignment=1
    )
    subtitle_style = ParagraphStyle(
        'DocSubtitle',
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=14,
        textColor=colors.HexColor('#0284C7'),
        alignment=1
    )
    meta_style = ParagraphStyle(
        'DocMeta',
        fontName='Helvetica',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#64748B'),
        alignment=1
    )
    h1_style = ParagraphStyle(
        'SectionH1',
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#0F172A'),
        spaceBefore=7,
        spaceAfter=3
    )

    kpi_card_title = ParagraphStyle(
        'KPICardTitle',
        fontName='Helvetica-Bold',
        fontSize=7.2,
        leading=9.5,
        textColor=colors.HexColor('#475569'),
        alignment=1
    )
    kpi_card_val = ParagraphStyle(
        'KPICardVal',
        fontName='Helvetica-Bold',
        fontSize=11.5,
        leading=14,
        textColor=colors.HexColor('#0F172A'),
        alignment=1
    )
    kpi_card_sub = ParagraphStyle(
        'KPICardSub',
        fontName='Helvetica',
        fontSize=6.8,
        leading=8.5,
        textColor=colors.HexColor('#0284C7'),
        alignment=1
    )

    table_hdr = ParagraphStyle(
        'TH',
        fontName='Helvetica-Bold',
        fontSize=7.2,
        leading=9,
        textColor=colors.white,
        alignment=1
    )
    table_hdr_left = ParagraphStyle(
        'THL',
        parent=table_hdr,
        alignment=0
    )
    table_txt = ParagraphStyle(
        'TD',
        fontName='Helvetica',
        fontSize=6.5,
        leading=8.0,
        textColor=colors.HexColor('#1E293B')
    )
    table_txt_center = ParagraphStyle(
        'TDC',
        parent=table_txt,
        alignment=1
    )
    table_txt_right = ParagraphStyle(
        'TDR',
        parent=table_txt,
        alignment=2
    )
    table_bold = ParagraphStyle(
        'TDB',
        fontName='Helvetica-Bold',
        fontSize=6.5,
        leading=8.0,
        textColor=colors.HexColor('#0F172A')
    )
    table_bold_right = ParagraphStyle(
        'TDBR',
        parent=table_bold,
        alignment=2
    )
    table_bold_center = ParagraphStyle(
        'TDBC',
        parent=table_bold,
        alignment=1
    )

    def fmt_idr(val):
        return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def fmt_qty(val):
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    story = []

    # Header
    story.append(Paragraph("SNAPRINT DIGITAL PRINTING", title_style))
    story.append(Paragraph("LAPORAN PENJUALAN TAHUNAN — TAHUN BUKU 2024", subtitle_style))
    story.append(Paragraph("Periode: 01 Januari 2024 s/d 31 Desember 2024 &nbsp;|&nbsp; Sumber Data: ACCURATE Accounting System", meta_style))
    story.append(Spacer(1, 6))

    # KPI Cards (Clean, neatly padded, distinct cells)
    kpi_data = [
        [
            [
                Paragraph("TOTAL PENJUALAN 2024", kpi_card_title),
                Spacer(1, 2),
                Paragraph(fmt_idr(tot_sales_all), kpi_card_val),
                Spacer(1, 1),
                Paragraph("12 Bulan Buku", kpi_card_sub)
            ],
            [
                Paragraph("TOTAL KUANTITAS TERJUAL", kpi_card_title),
                Spacer(1, 2),
                Paragraph(f"{fmt_qty(tot_qty_all)}", kpi_card_val),
                Spacer(1, 1),
                Paragraph("Unit / Meter / Lembar", kpi_card_sub)
            ],
            [
                Paragraph("RATA-RATA PENJUALAN BULANAN", kpi_card_title),
                Spacer(1, 2),
                Paragraph(fmt_idr(avg_sales_monthly), kpi_card_val),
                Spacer(1, 1),
                Paragraph("Rp 77,57 Juta / Bulan", kpi_card_sub)
            ]
        ]
    ]
    t_kpi = Table(kpi_data, colWidths=[175, 175, 175])
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 6))

    # Section 1: Ringkasan Penjualan Bulanan
    story.append(Paragraph("1. Ringkasan Kinerja Penjualan Bulanan (Januari - Desember 2024)", h1_style))
    months_order = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']
    months_full_name = {
        'Jan': 'Januari 2024',
        'Feb': 'Februari 2024',
        'Mar': 'Maret 2024',
        'Apr': 'April 2024',
        'Mei': 'Mei 2024',
        'Jun': 'Juni 2024',
        'Jul': 'Juli 2024',
        'Agu': 'Agustus 2024',
        'Sep': 'September 2024',
        'Okt': 'Oktober 2024',
        'Nov': 'November 2024',
        'Des': 'Desember 2024'
    }

    monthly_table_data = [
        [
            Paragraph("Bulan", table_hdr_left),
            Paragraph("Kuantitas (Volume)", table_hdr),
            Paragraph("Total Omzet Penjualan (Rp)", table_hdr),
            Paragraph("Kontribusi (%)", table_hdr)
        ]
    ]

    for m in months_order:
        sub = df_all[df_all['month'] == m]
        m_qty = sub['qty'].sum()
        m_sales = sub['sales'].sum()
        pct = (m_sales / tot_sales_all) * 100
        monthly_table_data.append([
            Paragraph(months_full_name[m], table_txt),
            Paragraph(fmt_qty(m_qty), table_txt_right),
            Paragraph(fmt_idr(m_sales), table_txt_right),
            Paragraph(f"{pct:.2f}%", table_txt_right)
        ])

    monthly_table_data.append([
        Paragraph("<b>TOTAL TAHUNAN 2024</b>", table_bold),
        Paragraph(fmt_qty(tot_qty_all), table_bold_right),
        Paragraph(fmt_idr(tot_sales_all), table_bold_right),
        Paragraph("100.00%", table_bold_right)
    ])

    t_month = Table(monthly_table_data, colWidths=[150, 115, 160, 100])
    t_month.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F8FAFC')]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 2.2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.2),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_month)
    story.append(Spacer(1, 6))

    # Section 2: Komposisi Penjualan Berdasarkan Kategori
    story.append(Paragraph("2. Komposisi Penjualan Berdasarkan Kategori Produk & Layanan", h1_style))
    cat_table_data = [
        [
            Paragraph("Kategori Produk / Layanan", table_hdr_left),
            Paragraph("Jml Item", table_hdr),
            Paragraph("Kuantitas (Volume)", table_hdr),
            Paragraph("Total Omzet Penjualan (Rp)", table_hdr),
            Paragraph("Porsi (%)", table_hdr)
        ]
    ]

    for idx, r in cat_summary.iterrows():
        cat_table_data.append([
            Paragraph(r['category'], table_txt),
            Paragraph(str(r['item_count']), table_txt_center),
            Paragraph(fmt_qty(r['total_qty']), table_txt_right),
            Paragraph(fmt_idr(r['total_sales']), table_txt_right),
            Paragraph(f"{r['pct']:.2f}%", table_txt_right)
        ])

    cat_table_data.append([
        Paragraph("<b>TOTAL</b>", table_bold),
        Paragraph(f"<b>{len(prod_summary)}</b>", table_bold_center),
        Paragraph(fmt_qty(tot_qty_all), table_bold_right),
        Paragraph(fmt_idr(tot_sales_all), table_bold_right),
        Paragraph("100.00%", table_bold_right)
    ])

    t_cat = Table(cat_table_data, colWidths=[185, 50, 95, 115, 80])
    t_cat.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E293B')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F8FAFC')]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 2.2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.2),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_cat)

    story.append(PageBreak())

    # Section 3: Rincian Produk Lengkap (129 Items)
    story.append(Paragraph("3. Daftar Rincian Penjualan Masing-Masing Produk (Urut Omzet Terbesar)", h1_style))
    desc_p = Paragraph(
        f"Akumulasi kuantitas penjualan dan nilai perolehan seluruh <b>{len(prod_summary)} produk dan layanan</b> sepanjang tahun buku 2024:",
        ParagraphStyle('SubDesc', parent=table_txt, textColor=colors.HexColor('#64748B'), fontSize=7.2, leading=9)
    )
    story.append(desc_p)
    story.append(Spacer(1, 4))

    # Column widths: 525 pt total (leaves 10pt buffer on 535pt printable width)
    # No: 30 pt, Item: 205 pt, Satuan: 45 pt, Qty: 75 pt, Sales: 125 pt, %: 45 pt
    prod_table_data = [
        [
            Paragraph("No", table_hdr),
            Paragraph("Nama Barang / Produk Layanan", table_hdr_left),
            Paragraph("Satuan", table_hdr),
            Paragraph("Kuantitas", table_hdr),
            Paragraph("Total Penjualan (Rp)", table_hdr),
            Paragraph("Porsi (%)", table_hdr)
        ]
    ]

    num = 1
    for idx, r in prod_summary.iterrows():
        u_str = r['unit'] if r['unit'] else '-'
        prod_table_data.append([
            Paragraph(str(num), table_txt_center),
            Paragraph(r['item'], table_txt),
            Paragraph(u_str, table_txt_center),
            Paragraph(fmt_qty(r['total_qty']), table_txt_right),
            Paragraph(fmt_idr(r['total_sales']), table_txt_right),
            Paragraph(f"{r['pct']:.2f}%", table_txt_right)
        ])
        num += 1

    prod_table_data.append([
        Paragraph("", table_txt),
        Paragraph("<b>TOTAL KESELURUHAN (129 ITEM)</b>", table_bold),
        Paragraph("", table_txt),
        Paragraph(fmt_qty(tot_qty_all), table_bold_right),
        Paragraph(fmt_idr(tot_sales_all), table_bold_right),
        Paragraph("100.00%", table_bold_right)
    ])

    t_prod = Table(
        prod_table_data,
        colWidths=[30, 205, 45, 75, 125, 45],
        repeatRows=1
    )
    t_prod.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('GRID', (0,0), (-1,-1), 0.35, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F8FAFC')]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 1.0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.0),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_prod)

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Generated clean PDF: {pdf_filename}")
    print(f"Total Sales: {fmt_idr(tot_sales_all)}")
    print(f"Total Qty: {fmt_qty(tot_qty_all)}")
    print(f"Item Count: {len(prod_summary)}")

if __name__ == "__main__":
    generate_pdf()
