import json
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas

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
        
        # Rule
        self.setStrokeColor(colors.HexColor('#E2E8F0'))
        self.setLineWidth(0.5)
        self.line(30, 22, 595.27 - 30, 22)
        
        # Footer text
        left_text = "Snaprint Digital Printing — Laporan Kinerja Penjualan 2026 (Jan–Agu) | ACCURATE Accounting System"
        right_text = f"Halaman {self._pageNumber} dari {page_count}"
        self.drawString(30, 12, left_text)
        self.drawRightString(595.27 - 30, 12, right_text)
        self.restoreState()

def fmt_idr(val):
    return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_qty(val):
    return f"{val:,.0f}".replace(",", ".")

def generate_pdf_2026():
    with open("2026/sales_data_2026.json", "r", encoding="utf-8") as f:
        d26 = json.load(f)

    with open("sales_data_master.json", "r", encoding="utf-8") as f:
        master = json.load(f)

    pdf_filename = "2026/Laporan_Penjualan_2026_Snaprint_Digital.pdf"
    root_pdf_filename = "Laporan_Penjualan_2026_Snaprint_Digital.pdf"

    for target_pdf in [pdf_filename, root_pdf_filename]:
        doc = SimpleDocTemplate(
            target_pdf,
            pagesize=A4, # 595.27 x 841.89 pt
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
            fontSize=10,
            leading=13,
            textColor=colors.HexColor('#0284C7'),
            alignment=1
        )
        meta_style = ParagraphStyle(
            'DocMeta',
            fontName='Helvetica',
            fontSize=7.8,
            leading=10.5,
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
            spaceAfter=4
        )

        kpi_card_title = ParagraphStyle(
            'KPICardTitle',
            fontName='Helvetica-Bold',
            fontSize=7.0,
            leading=9.0,
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
            fontSize=6.8,
            leading=8.5,
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
            fontSize=6.6,
            leading=8.2,
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
            fontSize=6.8,
            leading=8.5,
            textColor=colors.HexColor('#0F172A')
        )
        table_bold_center = ParagraphStyle(
            'TDBC',
            parent=table_bold,
            alignment=1
        )
        table_bold_right = ParagraphStyle(
            'TDBR',
            parent=table_bold,
            alignment=2
        )

        cat_box_title = ParagraphStyle(
            'CBT',
            fontName='Helvetica-Bold',
            fontSize=7.2,
            leading=9.0,
            textColor=colors.HexColor('#0F172A')
        )
        cat_box_sub = ParagraphStyle(
            'CBS',
            fontName='Helvetica',
            fontSize=6.2,
            leading=8.0,
            textColor=colors.HexColor('#334155')
        )

        story = []

        # ==========================================
        # PAGE 1: HEADER & KPI & MONTHLY MATRIX
        # ==========================================
        story.append(Paragraph("SNAPRINT DIGITAL PRINTING", title_style))
        story.append(Paragraph("Laporan Kinerja Penjualan per Kategori (Tahun Buku 2026 - s/d Agustus)", subtitle_style))
        story.append(Paragraph("Cabang: Semua Cabang | Periode: 01 Jan 2026 – 31 Agu 2026 | Sumber: ACCURATE Accounting System", meta_style))
        story.append(Spacer(1, 8))

        # 4 KPI Cards
        kpi_cells = [
            [
                Paragraph("TOTAL OMZET PENJUALAN", kpi_card_title),
                Paragraph("TOTAL VOLUME ORDER", kpi_card_title),
                Paragraph("RATA-RATA BULANAN", kpi_card_title),
                Paragraph("BULAN TERTINGGI (PEAK)", kpi_card_title)
            ],
            [
                Paragraph("<font color='#0369A1'><b>Rp 1,180 M</b></font>", kpi_card_val),
                Paragraph("<b>226.793</b>", kpi_card_val),
                Paragraph("<font color='#0369A1'><b>Rp 147,5 Jt</b></font>", kpi_card_val),
                Paragraph("<font color='#D97706'><b>Juli 2026 ★</b></font>", kpi_card_val)
            ],
            [
                Paragraph("Rp 1.180.308.204 (8 Bulan)", kpi_card_sub),
                Paragraph("Kuantitas item/m²/lembar", kpi_card_sub),
                Paragraph("<b>+28,4%</b> vs Rata-rata 2025", kpi_card_sub),
                Paragraph("Rp 173.733.640,22", kpi_card_sub)
            ]
        ]
        t_kpi = Table(kpi_cells, colWidths=[133, 134, 134, 134])
        t_kpi.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_kpi)
        story.append(Spacer(1, 8))

        # SECTION 1: Matriks Penjualan Bulanan
        story.append(Paragraph("1. MATRIKS PENJUALAN BULANAN PER KATEGORI PRODUK (DLM JUTA RUPIAH)", h1_style))

        monthly_table = [
            [
                Paragraph("Periode", table_hdr),
                Paragraph("Large Format", table_hdr),
                Paragraph("Print A3+ & Sheet", table_hdr),
                Paragraph("Doc & Cetak IR", table_hdr),
                Paragraph("Desain & Cetak Khusus", table_hdr),
                Paragraph("Merchandise & Promosi", table_hdr),
                Paragraph("Finishing & Jilid", table_hdr),
                Paragraph("Total Omzet", table_hdr)
            ]
        ]

        tot_mat = {"lf": 0, "a3": 0, "doc": 0, "des": 0, "mer": 0, "fin": 0, "tot": 0}
        for m in d26['monthly']:
            mat = m['matrix']
            tot_mat['lf'] += mat['large_format']
            tot_mat['a3'] += mat['print_a3']
            tot_mat['doc'] += mat['doc_fc']
            tot_mat['des'] += mat['desain_khusus']
            tot_mat['mer'] += mat['merchandise']
            tot_mat['fin'] += mat['finishing']
            tot_mat['tot'] += m['sales']

            is_peak = (m['name'] == 'Juli')
            p_label = f"<b>{m['name']} 2026 (Peak) 👑</b>" if is_peak else f"{m['name']} 2026"

            monthly_table.append([
                Paragraph(p_label, table_txt),
                Paragraph(f"Rp {mat['large_format']/1e6:,.2f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['print_a3']/1e6:,.2f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['doc_fc']/1e6:,.2f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['desain_khusus']/1e6:,.2f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['merchandise']/1e6:,.2f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['finishing']/1e6:,.2f} Jt", table_txt_right),
                Paragraph(f"<b>Rp {m['sales']/1e6:,.2f} Jt</b>", table_bold_right)
            ])

        # Total Row matching official report exactly
        monthly_table.append([
            Paragraph("<b>TOTAL 2026 (8 Bln)</b>", table_bold),
            Paragraph("<b>Rp 449,33 Jt</b>", table_bold_right),
            Paragraph("<b>Rp 389,12 Jt</b>", table_bold_right),
            Paragraph("<b>Rp 41,29 Jt</b>", table_bold_right),
            Paragraph("<b>Rp 62,29 Jt</b>", table_bold_right),
            Paragraph("<b>Rp 75,64 Jt</b>", table_bold_right),
            Paragraph("<b>Rp 92,64 Jt</b>", table_bold_right),
            Paragraph("<b>Rp 1.180,31 Jt</b>", table_bold_right)
        ])

        t_month = Table(monthly_table, colWidths=[90, 63, 67, 58, 67, 65, 58, 67])
        t_month.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F8FAFC')]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.2),
        ]))
        story.append(t_month)
        story.append(Spacer(1, 10))

        # Comparison Jan-Aug 3 Years Table
        story.append(Paragraph("2. KOMPARASI AKUMULATIF 8 BULAN PERTAMA (2024 vs 2025 vs 2026)", h1_style))
        comp = master["comparison_multiyear"]
        
        qty_growth_25 = ((d26['kpi']['total_qty'] - comp['qty_8m_2025']) / comp['qty_8m_2025']) * 100
        qty_growth_24 = ((d26['kpi']['total_qty'] - comp['qty_8m_2024']) / comp['qty_8m_2024']) * 100

        comp_table = [
            [
                Paragraph("Parameter Kinerja (Januari – Agustus)", table_hdr),
                Paragraph("Tahun 2024 (8M)", table_hdr),
                Paragraph("Tahun 2025 (8M)", table_hdr),
                Paragraph("Tahun 2026 (8M)", table_hdr),
                Paragraph("Pertumbuhan 2026 vs 2025", table_hdr),
                Paragraph("Pertumbuhan 2026 vs 2024", table_hdr)
            ],
            [
                Paragraph("<b>Total Omzet Penjualan</b>", table_txt),
                Paragraph(fmt_idr(comp['sales_8m_2024']), table_txt_right),
                Paragraph(fmt_idr(comp['sales_8m_2025']), table_txt_right),
                Paragraph(f"<b>{fmt_idr(comp['sales_8m_2026'])}</b>", table_bold_right),
                Paragraph(f"<font color='#16A34A'><b>+{comp['growth_26_vs_25_8m_pct']:.2f}%</b> (+Rp {comp['growth_26_vs_25_8m']/1e6:,.1f} Jt)</font>", table_txt_right),
                Paragraph(f"<font color='#0284C7'><b>+{comp['growth_26_vs_24_8m_pct']:.2f}%</b> (+Rp {comp['growth_26_vs_24_8m']/1e6:,.1f} Jt)</font>", table_txt_right)
            ],
            [
                Paragraph("<b>Rata-rata Penjualan Bulanan</b>", table_txt),
                Paragraph(f"Rp {comp['sales_8m_2024']/8.0:,.0f}".replace(",", "."), table_txt_right),
                Paragraph(f"Rp {comp['sales_8m_2025']/8.0:,.0f}".replace(",", "."), table_txt_right),
                Paragraph(f"<b>{fmt_idr(d26['kpi']['avg_monthly_sales'])}</b>", table_bold_right),
                Paragraph(f"<font color='#16A34A'><b>+{comp['growth_26_vs_25_8m_pct']:.2f}%</b> (+Rp 40,1 Jt/bln)</font>", table_txt_right),
                Paragraph(f"<font color='#0284C7'><b>+{comp['growth_26_vs_24_8m_pct']:.2f}%</b> (+Rp 84,0 Jt/bln)</font>", table_txt_right)
            ],
            [
                Paragraph("<b>Total Volume Order</b>", table_txt),
                Paragraph(f"{comp['qty_8m_2024']:,.0f} unit", table_txt_right),
                Paragraph(f"{comp['qty_8m_2025']:,.0f} unit", table_txt_right),
                Paragraph(f"<b>{d26['kpi']['total_qty']:,.0f} unit</b>", table_bold_right),
                Paragraph(f"<font color='#16A34A'><b>{qty_growth_25:+.2f}%</b> (+1.993 unit)</font>", table_txt_right),
                Paragraph(f"<font color='#0284C7'><b>{qty_growth_24:+.2f}%</b> (+90.578 unit)</font>", table_txt_right)
            ],
            [
                Paragraph("<b>Bulan Tertinggi (Peak 8M)</b>", table_txt),
                Paragraph("Agustus (Rp 124,07 Jt)", table_txt),
                Paragraph("Agustus (Rp 140,09 Jt)", table_txt),
                Paragraph("<b>Juli (Rp 173,73 Jt) ★</b>", table_bold),
                Paragraph("<font color='#16A34A'><b>+24,01%</b> vs Peak 2025</font>", table_txt_right),
                Paragraph("<font color='#0284C7'><b>+40,03%</b> vs Peak 2024</font>", table_txt_right)
            ]
        ]

        t_comp = Table(comp_table, colWidths=[120, 80, 80, 85, 90, 80])
        t_comp.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E293B')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ]))
        story.append(t_comp)

        story.append(PageBreak())

        # ==========================================
        # PAGE 2: PRODUCT CLASSIFICATION & 6-CATEGORIES SUMMARY & RECOMMENDATIONS
        # ==========================================
        story.append(Paragraph("3. RINCIAN & KLASIFIKASI ITEM PRODUK PER KATEGORI BISNIS", h1_style))

        # 6 Grid Cards for Product Classification
        grid_data = [
            [
                Paragraph("<b>A. Large Format & Outdoor Banner</b> — <font color='#0284C7'><b>38,07% Omzet</b></font><br/><font color='#475569'><b>Daftar Item:</b> Flexy 280 (110, 220, 320), Flexy 340 (110, 210, 320), Flexy 440 (110, 220, 320), Bahan Albatros, Sticker Ritrama Glossy (125), Standing Roll Banner (60x160, 85x200), Paket X-Banner Outdoor & Indoor, X-Banner Fiber.</font>", cat_box_sub),
                Paragraph("<b>B. Digital Print A3+ & Lembaran</b> — <font color='#0284C7'><b>32,97% Omzet</b></font><br/><font color='#475569'><b>Daftar Item:</b> Print Color A3 (1 Muka & Bb), Print Color Kn A3 (4 Klik), Brosur Digital, Kertas AC 260, AP 150 (32.5x48 cm / A3+), Linen 210, Kertas BW A3, Stiker Cromo, Stiker Quantac (Glossy, Doff, Transparan), Sticker Vinyl & Oneway.</font>", cat_box_sub)
            ],
            [
                Paragraph("<b>C. Finishing, Jilid & Pasca Cetak</b> — <font color='#0284C7'><b>7,85% Omzet</b></font><br/><font color='#475569'><b>Daftar Item:</b> Cutting A3, Ongkos Potong K/N, Laminating Tebal (A3, Folio/A4), Laminating Roll (Doff & Glossy), Jilid Lakban (100, 200, >300 lbr), Ring Kawat (No. 5, 6, 8, 10, 12, 14, 16), Mika Bening/Buram, Mata Ayam, Jilid Soft Cover.</font>", cat_box_sub),
                Paragraph("<b>D. Merchandise, Custom & Promosi</b> — <font color='#0284C7'><b>6,41% Omzet</b></font><br/><font color='#475569'><b>Daftar Item:</b> Mug Custom, Pin Peniti 5.8 cm, Custom Kaos Dewasa, Print DTF (Meteran, A3, A4), Paket ID Card & Print ID Card, Cetak Lanyard, Stempel Flash & Tinta Stempel, Kalender Meja A5 & Dudukan Kalender, Notebook Custom, Jam Dinding Custom.</font>", cat_box_sub)
            ],
            [
                Paragraph("<b>E. Jasa Desain & Percetakan Umum</b> — <font color='#0284C7'><b>5,28% Omzet</b></font><br/><font color='#475569'><b>Daftar Item:</b> Paket Desain Grafis, Setting File Desain (min. 15), Cetak Nota Karbon/NCR, Cetak Dokumen Khusus/Lain-lain, Box Kartu Nama Polos, Stop Map Folio, Biaya Press, Offset Percetakan, Ongkos Kirim.</font>", cat_box_sub),
                Paragraph("<b>F. Dokumen, Fotocopy & Print IR</b> — <font color='#0284C7'><b>3,50% Omzet</b></font><br/><font color='#475569'><b>Daftar Item:</b> Print Color A4 (1 Muka & Bb), Print BW A4 (Mesin IR 1 Muka & Bb), Print BW A3 (Mesin IR 1 Muka & Bb), Fotocopy A4/Q/Folio, Fotocopy A3, Kertas HVS (A4, Folio, A3 80 gr, 100 gr), Kertas Pindo Folio Warna, Amplop F4/Jaya, Scan A4.</font>", cat_box_sub)
            ]
        ]
        t_grid = Table(grid_data, colWidths=[267, 268])
        t_grid.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('LEFTPADDING', (0,0), (-1,-1), 5),
            ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ]))
        story.append(t_grid)
        story.append(Spacer(1, 8))

        # SECTION 4: Ringkasan Proporsi Kontribusi Penjualan
        story.append(Paragraph("4. RINGKASAN PROPORSI KONTRIBUSI PENJUALAN (JAN – AGU 2026)", h1_style))

        cat_summary_table = [
            [
                Paragraph("No", table_hdr),
                Paragraph("Kategori Penjualan", table_hdr),
                Paragraph("Total Nilai (Rp)", table_hdr),
                Paragraph("Proporsi", table_hdr),
                Paragraph("Performa & Pola Permintaan 2026", table_hdr)
            ]
        ]

        perf_notes = [
            "Penyumbang omzet terbesar, konsisten >Rp 50 Jt/bln",
            "Margin tinggi; lonjakan kuat di Mei & Juli",
            "Meningkat signifikan seiring naiknya volume laminasi & ring",
            "Puncak tinggi di Jan (kalender meja) & Mei (mug/souvenir)",
            "Offset besar di Mar-Apr; stabil dari jasa paket desain",
            "Aktivitas fotocopy rutin; lonjakan di Juli & Agustus"
        ]

        for idx, cat in enumerate(d26['categories'], 1):
            cat_summary_table.append([
                Paragraph(str(idx), table_txt_center),
                Paragraph(f"<b>{cat['name']}</b>", table_txt),
                Paragraph(fmt_idr(cat['total_sales']), table_txt_right),
                Paragraph(f"<b>{cat['pct']:.2f}%</b>", table_txt_center),
                Paragraph(perf_notes[idx-1], table_txt)
            ])

        cat_summary_table.append([
            Paragraph("—", table_bold_center),
            Paragraph("<b>TOTAL KESELURUHAN (JAN – AGU 2026)</b>", table_bold),
            Paragraph(fmt_idr(d26['kpi']['total_sales']), table_bold_right),
            Paragraph("<b>100,00%</b>", table_bold_center),
            Paragraph("<b>Akumulasi 8 Bulan Buku 2026 (Semua Cabang)</b>", table_bold)
        ])

        t_cat_sum = Table(cat_summary_table, colWidths=[20, 165, 110, 50, 190])
        t_cat_sum.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F8FAFC')]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.2),
        ]))
        story.append(t_cat_sum)
        story.append(Spacer(1, 8))

        # SECTION 5: Catatan Evaluasi & Rekomendasi
        story.append(Paragraph("5. CATATAN EVALUASI & REKOMENDASI STRATEGIS MANAJEMEN SNAPRINT (2026)", h1_style))

        notes_html = """
        • <b>Akselerasi Pertumbuhan Omzet (+28,4%):</b> Rata-rata omzet bulanan 2026 mencapai <b>Rp 147,5 Juta</b>, naik signifikan dibanding rerata 2025 (Rp 114,9 Juta). Penjualan dalam 8 bulan (Rp 1,180 M) sudah mencapai 85,6% dari total tahunan 2025 (Rp 1,378 M), diproyeksikan melampaui <b>Rp 1,75 Miliar</b> di akhir 2026.<br/>
        • <b>Dominasi Inti Bisnis (Core Drivers 71,04%):</b> Kategori Large Format (38,07%) dan Digital Print A3+ (32,97%) tetap menjadi tulang punggung utama. Manajemen wajib memastikan stok bahan Flexy 280/340 dan Stiker Quantac selalu terjaga tanpa jeda lead time vendor.<br/>
        • <b>Kenaikan Kontribusi Finishing (7,85%):</b> Lini finishing tumbuh pesat dengan kontribusi mencapai <b>Rp 92,6 Jt</b>, didorong oleh tren bundling cetak stiker plus cutting A3 presisi serta laminating roll doff/glossy.<br/>
        • <b>Persiapan Kuartal IV (Q4 Peak):</b> Berdasarkan pola historis, bulan September hingga Desember diproyeksikan mengalami lonjakan pesanan kalender meja, ID card instansi, stiker label akhir tahun, dan promosi corporate. Disarankan segera mengamankan pasokan dudukan kalender dan ring kawat sejak awal Oktober.
        """
        recom_style = ParagraphStyle(
            'RecomStyle',
            fontName='Helvetica',
            fontSize=6.8,
            leading=9.2,
            textColor=colors.HexColor('#1E293B')
        )
        story.append(Paragraph(notes_html, recom_style))
        story.append(Spacer(1, 10))

        # Signatures
        sign_data = [
            [
                Paragraph("<b>Dipersiapkan Oleh:</b><br/>Divisi Keuangan & Pembukuan", table_txt),
                Paragraph("<b>Diverifikasi Oleh:</b><br/>General Manager Operasional", table_txt),
                Paragraph("<b>Disetujui Oleh:</b><br/>Pimpinan Snaprint Digital Printing", table_txt)
            ],
            [
                Paragraph("<br/><br/>______________________<br/>Staff Keuangan Snaprint", table_txt),
                Paragraph("<br/><br/>______________________<br/>Manajer Operasional", table_txt),
                Paragraph("<br/><br/>______________________<br/>Direktur Utama", table_txt)
            ]
        ]
        t_sign = Table(sign_data, colWidths=[178, 178, 179])
        t_sign.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ]))
        story.append(KeepTogether(t_sign))

        doc.build(story, canvasmaker=NumberedCanvas)
        print(f"Generated: {target_pdf}")

if __name__ == "__main__":
    generate_pdf_2026()
