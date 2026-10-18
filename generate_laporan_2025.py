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
        
        # Line rule
        self.setStrokeColor(colors.HexColor('#E2E8F0'))
        self.setLineWidth(0.5)
        self.line(30, 22, 595.27 - 30, 22)
        
        # Footer text
        left_text = "Snaprint Digital Printing — Laporan Kinerja Penjualan Tahunan 2025 (Konsolidasi ACCURATE)"
        right_text = f"Halaman {self._pageNumber} dari {page_count}"
        self.drawString(30, 12, left_text)
        self.drawRightString(595.27 - 30, 12, right_text)
        self.restoreState()

def fmt_idr(val):
    return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_qty(val):
    return f"{val:,.0f}".replace(",", ".")

def generate_pdf_2025():
    with open("2025/sales_data_2025.json", "r", encoding="utf-8") as f:
        d25 = json.load(f)

    with open("sales_data_master.json", "r", encoding="utf-8") as f:
        master = json.load(f)

    pdf_filename = "2025/Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"
    
    # Also save to root for easy one-click access
    root_pdf_filename = "Laporan_Penjualan_Tahunan_2025_Snaprint_Digital.pdf"

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
            spaceBefore=8,
            spaceAfter=4
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
            fontSize=6.8,
            leading=8.5,
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
            fontSize=7.0,
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
        notes_txt = ParagraphStyle(
            'Notes',
            fontName='Helvetica',
            fontSize=7.0,
            leading=9.5,
            textColor=colors.HexColor('#334155')
        )

        story = []

        # ==========================================
        # PAGE 1: HEADER & KPI & MONTHLY SUMMARY
        # ==========================================
        story.append(Paragraph("SNAPRINT DIGITAL PRINTING", title_style))
        story.append(Paragraph("LAPORAN KINERJA PENJUALAN TAHUNAN — TAHUN BUKU 2025", subtitle_style))
        story.append(Paragraph("Periode: 01 Januari 2025 s/d 31 Desember 2025 | Sumber: ACCURATE Accounting System | Cabang: Semua Cabang", meta_style))
        story.append(Spacer(1, 8))

        # KPI 3-Cards
        kpi_cells = [
            [
                Paragraph("TOTAL PENJUALAN 2025", kpi_card_title),
                Paragraph("TOTAL KUANTITAS ORDER", kpi_card_title),
                Paragraph("RATA-RATA PENJUALAN BULANAN", kpi_card_title)
            ],
            [
                Paragraph(f"<font color='#0369A1'><b>{fmt_idr(d25['kpi']['total_sales'])}</b></font>", kpi_card_val),
                Paragraph(f"<font color='#0F172A'><b>{fmt_qty(d25['kpi']['total_qty'])}</b></font>", kpi_card_val),
                Paragraph(f"<font color='#0369A1'><b>{fmt_idr(d25['kpi']['avg_monthly_sales'])}</b></font>", kpi_card_val)
            ],
            [
                Paragraph("<b>+48,14% YoY</b> (vs 2024: Rp 930,8 Jt)", kpi_card_sub),
                Paragraph("unit/m²/lbr (<b>+40,58% YoY</b>)", kpi_card_sub),
                Paragraph(f"Bulan Puncak: <b>Agustus ({fmt_idr(d25['kpi']['peak_month_sales'])})</b>", kpi_card_sub)
            ]
        ]
        t_kpi = Table(kpi_cells, colWidths=[178, 178, 179])
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

        # SECTION 1: Ringkasan Bulanan 2025
        story.append(Paragraph("1. Rekapitulasi Kinerja Penjualan Bulanan (Januari – Desember 2025)", h1_style))

        monthly_table = [
            [
                Paragraph("Bulan", table_hdr),
                Paragraph("Kuantitas (Unit/m²)", table_hdr),
                Paragraph("Omzet Penjualan (Rp)", table_hdr),
                Paragraph("Porsi (%)", table_hdr),
                Paragraph("MoM Growth", table_hdr),
                Paragraph("Deviasi Rata-rata", table_hdr)
            ]
        ]

        for m in d25['monthly']:
            mom_str = f"{m['mom_growth']:+.1f}%" if m['mom_growth'] != 0 else "-"
            mom_color = "#16A34A" if m['mom_growth'] > 0 else ("#DC2626" if m['mom_growth'] < 0 else "#64748B")
            dev_str = f"{m['vs_avg']:+.1f}%"
            dev_color = "#0284C7" if m['vs_avg'] >= 0 else "#DC2626"
            
            is_peak = (m['name'] == d25['kpi']['peak_month'])
            name_label = f"<b>{m['name']} 2025 👑 (Peak)</b>" if is_peak else f"{m['name']} 2025"

            monthly_table.append([
                Paragraph(name_label, table_txt),
                Paragraph(fmt_qty(m['qty']), table_txt_right),
                Paragraph(fmt_idr(m['sales']), table_txt_right),
                Paragraph(f"{m['pct']:.2f}%", table_txt_right),
                Paragraph(f"<font color='{mom_color}'>{mom_str}</font>", table_txt_right),
                Paragraph(f"<font color='{dev_color}'>{dev_str}</font>", table_txt_right)
            ])

        # Total Row
        monthly_table.append([
            Paragraph("<b>TOTAL TAHUNAN 2025</b>", table_bold),
            Paragraph(fmt_qty(d25['kpi']['total_qty']), table_bold_right),
            Paragraph(fmt_idr(d25['kpi']['total_sales']), table_bold_right),
            Paragraph("100.00%", table_bold_right),
            Paragraph("—", table_bold_center),
            Paragraph("—", table_bold_center)
        ])

        t_month = Table(monthly_table, colWidths=[105, 80, 125, 65, 80, 80])
        t_month.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E293B')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F8FAFC')]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.2),
        ]))
        story.append(t_month)
        story.append(Spacer(1, 8))

        # SECTION 2: Matriks Bulanan Kategori 2025
        story.append(Paragraph("2. Matriks Penerimaan Penjualan per Kategori Produk (Januari – Desember 2025)", h1_style))

        matrix_table = [
            [
                Paragraph("Bulan", table_hdr),
                Paragraph("Large Format", table_hdr),
                Paragraph("Print A3+ & Sheet", table_hdr),
                Paragraph("Doc & Fotocopy", table_hdr),
                Paragraph("Finishing", table_hdr),
                Paragraph("Merchandise", table_hdr),
                Paragraph("Total Bulanan", table_hdr)
            ]
        ]

        tot_mat = {"lf": 0, "a3": 0, "doc": 0, "fin": 0, "mer": 0}
        for m in d25['monthly']:
            mat = m['matrix']
            tot_mat['lf'] += mat['large_format']
            tot_mat['a3'] += mat['print_a3']
            tot_mat['doc'] += mat['doc_fc']
            tot_mat['fin'] += mat['finishing']
            tot_mat['mer'] += mat['merchandise']
            
            matrix_table.append([
                Paragraph(m['name'], table_txt),
                Paragraph(f"Rp {mat['large_format']/1e6:,.1f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['print_a3']/1e6:,.1f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['doc_fc']/1e6:,.1f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['finishing']/1e6:,.1f} Jt", table_txt_right),
                Paragraph(f"Rp {mat['merchandise']/1e6:,.1f} Jt", table_txt_right),
                Paragraph(f"<b>Rp {m['sales']/1e6:,.2f} Jt</b>", table_bold_right)
            ])

        # Matrix Total
        matrix_table.append([
            Paragraph("<b>TOTAL 2025</b>", table_bold),
            Paragraph(f"<b>Rp {tot_mat['lf']/1e6:,.1f} Jt</b>", table_bold_right),
            Paragraph(f"<b>Rp {tot_mat['a3']/1e6:,.1f} Jt</b>", table_bold_right),
            Paragraph(f"<b>Rp {tot_mat['doc']/1e6:,.1f} Jt</b>", table_bold_right),
            Paragraph(f"<b>Rp {tot_mat['fin']/1e6:,.1f} Jt</b>", table_bold_right),
            Paragraph(f"<b>Rp {tot_mat['mer']/1e6:,.1f} Jt</b>", table_bold_right),
            Paragraph(f"<b>Rp {d25['kpi']['total_sales']/1e6:,.2f} Jt</b>", table_bold_right)
        ])

        t_matrix = Table(matrix_table, colWidths=[75, 75, 80, 75, 70, 75, 85])
        t_matrix.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F8FAFC')]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ]))
        story.append(t_matrix)

        story.append(PageBreak())

        # ==========================================
        # PAGE 2: 6 CATEGORIES SUMMARY & YOY COMPARISON & STRATEGY
        # ==========================================
        story.append(Paragraph("3. Proporsi Kontribusi Penjualan per Kategori Bisnis Utama (Tahun 2025)", h1_style))

        cat_table = [
            [
                Paragraph("No", table_hdr),
                Paragraph("Kategori Produk / Layanan", table_hdr),
                Paragraph("Total Omzet (Rp)", table_hdr),
                Paragraph("Porsi (%)", table_hdr),
                Paragraph("Contoh Produk Unggulan & Karakter Pasar", table_hdr)
            ]
        ]

        for idx, cat in enumerate(d25['categories'], 1):
            cat_table.append([
                Paragraph(str(idx), table_txt_center),
                Paragraph(f"<b>{cat['name']}</b>", table_txt),
                Paragraph(fmt_idr(cat['total_sales']), table_txt_right),
                Paragraph(f"{cat['pct']:.2f}%", table_txt_right),
                Paragraph(cat['items'], table_txt)
            ])

        cat_table.append([
            Paragraph("—", table_bold_center),
            Paragraph("<b>TOTAL KESELURUHAN 2025</b>", table_bold),
            Paragraph(fmt_idr(d25['kpi']['total_sales']), table_bold_right),
            Paragraph("100.00%", table_bold_right),
            Paragraph("<b>Konsolidasi 12 Bulan Buku ACCURATE</b>", table_bold)
        ])

        t_cat = Table(cat_table, colWidths=[24, 160, 110, 56, 185])
        t_cat.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E293B')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F8FAFC')]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ]))
        story.append(t_cat)
        story.append(Spacer(1, 8))

        # SECTION 4: Komparasi Year-over-Year (2024 vs 2025)
        story.append(Paragraph("4. Analisis Komparasi Pertumbuhan Year-Over-Year (YoY: 2024 vs 2025)", h1_style))

        yoy = master['yoy']
        kpi24 = master['data_2024']['kpi']
        kpi25 = master['data_2025']['kpi']

        yoy_table = [
            [
                Paragraph("Parameter Kinerja Finansial", table_hdr),
                Paragraph("Tahun 2024", table_hdr),
                Paragraph("Tahun 2025", table_hdr),
                Paragraph("Pertumbuhan (+/-)", table_hdr),
                Paragraph("Status Evaluasi Strategis", table_hdr)
            ],
            [
                Paragraph("<b>Total Omzet Penjualan</b>", table_txt),
                Paragraph(fmt_idr(kpi24['total_sales']), table_txt_right),
                Paragraph(fmt_idr(kpi25['total_sales']), table_txt_right),
                Paragraph(f"<font color='#16A34A'><b>+{yoy['growth_sales_pct']:.2f}%</b> (+Rp {yoy['growth_sales']/1e6:,.1f} Jt)</font>", table_txt_right),
                Paragraph("Lompatan kinerja menembus target Rp 1 Miliar", table_txt)
            ],
            [
                Paragraph("<b>Total Kuantitas Produksi</b>", table_txt),
                Paragraph(f"{kpi24['total_qty']:,.2f} unit", table_txt_right),
                Paragraph(f"{kpi25['total_qty']:,.0f} unit", table_txt_right),
                Paragraph(f"<font color='#16A34A'><b>+{yoy['growth_qty_pct']:.2f}%</b> (+{yoy['growth_qty']:,.0f})</font>", table_txt_right),
                Paragraph("Utilisasi kapasitas cetak mesin meningkat pesat", table_txt)
            ],
            [
                Paragraph("<b>Rata-rata Penjualan Bulanan</b>", table_txt),
                Paragraph(fmt_idr(kpi24['avg_monthly_sales']), table_txt_right),
                Paragraph(fmt_idr(kpi25['avg_monthly_sales']), table_txt_right),
                Paragraph(f"<font color='#16A34A'><b>+{yoy['growth_sales_pct']:.2f}%</b> (+Rp {(kpi25['avg_monthly_sales'] - kpi24['avg_monthly_sales'])/1e6:,.1f} Jt)</font>", table_txt_right),
                Paragraph("Stabilitas cashflow terjaga merata di tiap bulan", table_txt)
            ],
            [
                Paragraph("<b>Bulan Tertinggi (Peak)</b>", table_txt),
                Paragraph(f"Agustus ({fmt_idr(kpi24['peak_month_sales'])})", table_txt_right),
                Paragraph(f"Agustus ({fmt_idr(kpi25['peak_month_sales'])})", table_txt_right),
                Paragraph(f"<font color='#16A34A'><b>{((kpi25['peak_month_sales']-kpi24['peak_month_sales'])/kpi24['peak_month_sales'])*100:+.2f}%</b></font>", table_txt_right),
                Paragraph("Agustus menjadi musim panen tahunan tetap (HUT RI)", table_txt)
            ],
            [
                Paragraph("<b>Penjualan Kuartal 1 (Q1)</b>", table_txt),
                Paragraph(fmt_idr(kpi24['q1_sales']), table_txt_right),
                Paragraph(fmt_idr(kpi25['q1_sales']), table_txt_right),
                Paragraph(f"<font color='#16A34A'><b>{((kpi25['q1_sales']-kpi24['q1_sales'])/kpi24['q1_sales'])*100:+.2f}%</b></font>", table_txt_right),
                Paragraph("Eliminasi 'low season' Januari melalui kontrak korporat", table_txt)
            ]
        ]

        t_yoy = Table(yoy_table, colWidths=[120, 100, 105, 105, 105])
        t_yoy.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ]))
        story.append(t_yoy)
        story.append(Spacer(1, 8))

        # SECTION 5: Rekomendasi Strategis 2026
        story.append(Paragraph("5. Rekomendasi Strategis Manajemen (Roadmap 2026)", h1_style))

        recom_text = """
        <b>1. Pengamanan Bahan Baku Inti:</b> Large Format dan Digital Print A3+ menyumbang <b>72,34% omzet</b>. Manajemen disarankan melakukan kontrak pasokan bahan baku (Flexy 280/440 dan Art Carton/Stiker Quantac) sebelum memasuki kuartal ramai untuk mengunci harga dan mencegah kehabisan stok.<br/>
        <b>2. Antisipasi Musiman Kuartal 3 & 4:</b> Puncak produksi konsisten terjadi di Agustus (HUT RI), Oktober (kampanye/event), serta Desember (kalender/agenda). Persiapan kapasitas mesin dan lembur staf perlu dijadwalkan secara terencana sejak bulan Juli.<br/>
        <b>3. Peningkatan Margin Melalui Finishing & Custom:</b> Jasa finishing (cutting A3, laminasi) dan merchandise memiliki margin laba bersih yang sangat tinggi. Perluas penawaran paket bundling terintegrasi pada saat pemesanan cetak brosur, stiker, dan poster.
        """
        story.append(Paragraph(recom_text, notes_txt))

        # Sign-off box
        story.append(Spacer(1, 10))
        sign_data = [
            [
                Paragraph("<b>Dipersiapkan Oleh:</b><br/>Divisi Keuangan & Administrasi", table_txt),
                Paragraph("<b>Diverifikasi Oleh:</b><br/>General Manager Operasional", table_txt),
                Paragraph("<b>Disetujui Oleh:</b><br/>Direktur Utama / Pemilik Usaha", table_txt)
            ],
            [
                Paragraph("<br/><br/>______________________<br/>Staff Keuangan Snaprint", table_txt),
                Paragraph("<br/><br/>______________________<br/>Manajer Operasional", table_txt),
                Paragraph("<br/><br/>______________________<br/>Pimpinan Snaprint Digital", table_txt)
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
    generate_pdf_2025()
