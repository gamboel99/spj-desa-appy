import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload,
             nama_kades="Sutrisno, S.E.", nama_ketua_bpd="Misdi", nama_ketua_lembaga="", nama_bendahara="",
             kode_register="470/SPJ-DESA"):

    doc = Document()

    # Tambahkan logo jika ada
    logo_path = os.path.join(os.path.dirname(__file__), "logo_desa.png")
    if os.path.exists(logo_path):
        doc.add_picture(logo_path, width=Inches(1.2))

    # KOP Surat
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("PEMERINTAH DESA KELING\n").bold = True
    p.add_run("KECAMATAN KEPUNG, KABUPATEN KEDIRI\n").bold = True
    p.add_run("Alamat: Jl. Raya Keling, Bukaan, Keling, Kediri, Jawa Timur 64293\n").italic = True
    doc.add_paragraph("_______________________________________________________________")

    # Nomor surat
    doc.add_paragraph(f"Nomor: {kode_register}/{tgl.month:02d}/{tgl.year}").runs[0].bold = True

    # Judul
    judul = doc.add_paragraph("SURAT PERTANGGUNGJAWABAN KEGIATAN")
    judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    judul.runs[0].bold = True
    doc.add_paragraph(" ")

    # Data kegiatan
    doc.add_paragraph(f"Nama Kegiatan        : {nama_kegiatan}")
    doc.add_paragraph(f"Tanggal Pelaksanaan  : {tgl.strftime('%d-%m-%Y')}")
    doc.add_paragraph(f"Lembaga Pelaksana    : {lembaga}")
    doc.add_paragraph(f"Lokasi               : {lokasi}")
    doc.add_paragraph(f"RAB (Anggaran)       : Rp {anggaran:,.0f}")
    doc.add_paragraph(f"Realisasi            : Rp {realisasi:,.0f}")
    doc.add_paragraph(f"Selisih              : Rp {anggaran - realisasi:,.0f}")
    doc.add_paragraph(f"Sumber Dana          : {sumber_dana}")
    doc.add_paragraph(" ")

    # Tanda tangan
    doc.add_paragraph(f"Desa Keling, {tgl.strftime('%d-%m-%Y')}", style='Normal').alignment = WD_ALIGN_PARAGRAPH.RIGHT

    table = doc.add_table(rows=4, cols=2)
    table.style = 'Table Grid'

    table.cell(0, 0).text = "Mengetahui:\nKepala Desa"
    table.cell(0, 1).text = f"Lembaga Pelaksana\nKetua {lembaga}"
    table.cell(1, 0).text = nama_kades
    table.cell(1, 1).text = nama_ketua_lembaga or ".................."
    table.cell(2, 0).text = ""
    table.cell(2, 1).text = "Bendahara"
    table.cell(3, 0).text = ""
    table.cell(3, 1).text = nama_bendahara or ".................."

    doc.add_paragraph(" ")

    bpd = doc.add_paragraph("Mengesahkan,\nKetua BPD\n\n" + nama_ketua_bpd)
    bpd.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Simpan
    output_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(output_path)
    return output_path
