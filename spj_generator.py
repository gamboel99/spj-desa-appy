import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, sumber_dana, rab_df, bukti_upload,
             nama_kades, nama_ketua_bpd, nama_ketua_lembaga, nama_bendahara, kode_register):

    doc = Document()

    # Tambahkan logo jika ada
    logo_path = os.path.join(os.path.dirname(__file__), "logo_desa.png")
    if os.path.exists(logo_path):
        doc.add_picture(logo_path, width=Inches(1.2))

    # KOP
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("PEMERINTAH DESA KELING\n").bold = True
    p.add_run("KECAMATAN KEPUNG, KABUPATEN KEDIRI\n").bold = True
    p.add_run("Alamat: Jl. Raya Keling, Bukaan, Keling, Kediri, Jawa Timur 64293\n").italic = True
    doc.add_paragraph("_______________________________________________________________")

    doc.add_paragraph(f"\nNomor: {kode_register}/{tgl.month:02d}/{tgl.year}\n", style="Normal")

    # Judul
    judul = doc.add_paragraph("SURAT PERTANGGUNGJAWABAN KEGIATAN")
    judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    judul.runs[0].bold = True

    # Info kegiatan
    doc.add_paragraph(f"\nNama Kegiatan        : {nama_kegiatan}")
    doc.add_paragraph(f"Tanggal Pelaksanaan  : {tgl.strftime('%d-%m-%Y')}")
    doc.add_paragraph(f"Lembaga Pelaksana    : {lembaga}")
    doc.add_paragraph(f"Lokasi               : {lokasi}")
    doc.add_paragraph(f"Sumber Dana          : {sumber_dana}\n")

    # Tabel RAB & Realisasi
    doc.add_paragraph("Tabel RAB dan Realisasi:")

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "No"
    hdr_cells[1].text = "Uraian"
    hdr_cells[2].text = "Volume"
    hdr_cells[3].text = "Satuan"
    hdr_cells[4].text = "Harga Satuan"
    hdr_cells[5].text = "Realisasi"

    total_anggaran = 0
    total_realisasi = 0

    for idx, row in rab_df.iterrows():
        try:
            vol = float(row["Volume"])
            harga = float(row["Harga Satuan"])
            real = float(row["Realisasi"])
        except:
            vol = harga = real = 0

        jumlah = vol * harga
        total_anggaran += jumlah
        total_realisasi += real

        row_cells = table.add_row().cells
        row_cells[0].text = str(idx + 1)
        row_cells[1].text = str(row["Uraian"])
        row_cells[2].text = str(row["Volume"])
        row_cells[3].text = str(row["Satuan"])
        row_cells[4].text = f"Rp {harga:,.0f}"
        row_cells[5].text = f"Rp {real:,.0f}"

    # Total
    doc.add_paragraph(f"\nTotal Anggaran : Rp {total_anggaran:,.0f}")
    doc.add_paragraph(f"Total Realisasi: Rp {total_realisasi:,.0f}")
    doc.add_paragraph(f"Selisih        : Rp {total_anggaran - total_realisasi:,.0f}")

    # Penutup & TTD
    doc.add_paragraph(f"\nDesa Keling, {tgl.strftime('%d-%m-%Y')}")
    table_ttd = doc.add_table(rows=4, cols=2)
    table_ttd.style = 'Table Grid'
    table_ttd.cell(0, 0).text = "Mengetahui:\nKepala Desa"
    table_ttd.cell(0, 1).text = f"Lembaga Pelaksana\nKetua {lembaga}"
    table_ttd.cell(1, 0).text = nama_kades
    table_ttd.cell(1, 1).text = nama_ketua_lembaga
    table_ttd.cell(2, 1).text = "Bendahara"
    table_ttd.cell(3, 1).text = nama_bendahara

    doc.add_paragraph("\nMengesahkan,\nKetua BPD\n" + nama_ketua_bpd)

    # Simpan dokumen
    output_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(output_path)
    return output_path
