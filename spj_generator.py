import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, sumber_dana, rab_df, bukti_upload,
             nama_kades, nama_ketua_bpd, nama_ketua_lembaga, nama_bendahara, kode_register):

    doc = Document()

    # Logo (jika ada)
    logo_path = os.path.join(os.path.dirname(__file__), "logo_desa.png")
    if os.path.exists(logo_path):
        doc.add_picture(logo_path, width=Inches(1.2))

    # Kop Surat
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("PEMERINTAH DESA KELING\n").bold = True
    p.add_run("KECAMATAN KEPUNG, KABUPATEN KEDIRI\n").bold = True
    p.add_run("Alamat: Jl. Raya Keling, Bukaan, Keling, Kediri, Jawa Timur 64293\n").italic = True
    doc.add_paragraph("_______________________________________________________________")

    doc.add_paragraph(f"\nNomor: {kode_register}/{tgl.month:02d}/{tgl.year}\n", style="Normal")

    # Judul SPJ
    judul = doc.add_paragraph("SURAT PERTANGGUNGJAWABAN KEGIATAN")
    judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    judul.runs[0].bold = True

    # Informasi kegiatan
    doc.add_paragraph(f"\nNama Kegiatan        : {nama_kegiatan}")
    doc.add_paragraph(f"Tanggal Pelaksanaan  : {tgl.strftime('%d-%m-%Y')}")
    doc.add_paragraph(f"Lembaga Pelaksana    : {lembaga}")
    doc.add_paragraph(f"Lokasi               : {lokasi}")
    doc.add_paragraph(f"Sumber Dana          : {sumber_dana}\n")

    # Tabel RAB & Realisasi
    doc.add_paragraph("Tabel RAB dan Realisasi:")

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    hdr[0].text = "No"
    hdr[1].text = "Uraian"
    hdr[2].text = "Volume"
    hdr[3].text = "Satuan"
    hdr[4].text = "Harga Satuan"
    hdr[5].text = "Realisasi"

    total_anggaran = 0
    total_realisasi = 0

    for i, row in rab_df.iterrows():
        try:
            volume = float(row["Volume"])
            harga = float(row["Harga Satuan"])
            realisasi = float(row["Realisasi"])
        except:
            volume = harga = realisasi = 0

        jumlah = volume * harga
        total_anggaran += jumlah
        total_realisasi += realisasi

        r = table.add_row().cells
        r[0].text = str(i + 1)
        r[1].text = str(row["Uraian"])
        r[2].text = str(row["Volume"])
        r[3].text = str(row["Satuan"])
        r[4].text = f"Rp {harga:,.0f}"
        r[5].text = f"Rp {realisasi:,.0f}"

    doc.add_paragraph(f"\nTotal Anggaran : Rp {total_anggaran:,.0f}")
    doc.add_paragraph(f"Total Realisasi: Rp {total_realisasi:,.0f}")
    doc.add_paragraph(f"Selisih        : Rp {total_anggaran - total_realisasi:,.0f}")

    # Paragraf TTD Atas
    doc.add_paragraph(f"\nDesa Keling, {tgl.strftime('%d-%m-%Y')}\n")

    # Tanda tangan rapi tanpa border menggunakan table invisible
    ttd_table = doc.add_table(rows=4, cols=2)
    ttd_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Hapus border tabel
    tbl = ttd_table._tbl
    for row in tbl.xpath(".//w:tr"):
        for cell in row.xpath(".//w:tc"):
            tcPr = cell.find(".//w:tcPr", namespaces=cell.nsmap)
            if tcPr is None:
                tcPr = cell.makeelement("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr")
                cell.insert(0, tcPr)
            tcBorders = cell.makeelement("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders")
            tcPr.append(tcBorders)

    ttd_table.cell(0, 0).text = "Mengetahui:\nKepala Desa"
    ttd_table.cell(0, 1).text = f"Lembaga Pelaksana\nKetua {lembaga}"
    ttd_table.cell(1, 0).text = nama_kades
    ttd_table.cell(1, 1).text = nama_ketua_lembaga
    ttd_table.cell(2, 0).text = ""
    ttd_table.cell(2, 1).text = "Bendahara"
    ttd_table.cell(3, 0).text = ""
    ttd_table.cell(3, 1).text = nama_bendahara

    # TTD BPD di bawah
    doc.add_paragraph("\nMengesahkan,\nKetua BPD\n\n" + nama_ketua_bpd)

    # Simpan file
    output_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(output_path)
    return output_path
