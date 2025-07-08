# =========================
# File: spj_generator.py
# =========================

import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    doc = Document()

    # KOP DESA
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = header.add_run("PEMERINTAH DESA KELING\n")
    header_run.bold = True
    header_run.font.size = Pt(14)
    header.add_run("KECAMATAN KEPUNG, KABUPATEN KEDIRI\n").bold = True
    header.add_run("Alamat: Jl. Raya Keling, Bukaan, Keling, Kediri, Jawa Timur 64293\n").italic = True
    doc.add_paragraph("_______________________________________________________________")

    # JUDUL
    judul = doc.add_paragraph("SURAT PERTANGGUNGJAWABAN KEGIATAN")
    judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    judul.runs[0].bold = True
    doc.add_paragraph("\n")

    # ISI DATA KEGIATAN
    doc.add_paragraph(f"Nama Kegiatan        : {nama_kegiatan}")
    doc.add_paragraph(f"Tanggal Pelaksanaan  : {tgl.strftime('%d-%m-%Y')}")
    doc.add_paragraph(f"Lembaga Pelaksana    : {lembaga}")
    doc.add_paragraph(f"Lokasi               : {lokasi}")
    doc.add_paragraph(f"Anggaran             : Rp {anggaran:,.0f}")
    doc.add_paragraph(f"Realisasi            : Rp {realisasi:,.0f}")
    doc.add_paragraph(f"Sumber Dana          : {sumber_dana}")

    doc.add_paragraph("\n")

    # TANGGAL DAN TEMPAT
    ttd_paragraf = doc.add_paragraph()
    ttd_paragraf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ttd_paragraf.add_run(f"Desa Keling, {tgl.strftime('%d-%m-%Y')}\n")

    # TABEL TANDA TANGAN
    table = doc.add_table(rows=4, cols=2)
    table.style = 'Table Grid'

    # Baris 1
    table.cell(0, 0).text = "Mengetahui:\nKepala Desa"
    table.cell(0, 1).text = f"Lembaga Pelaksana\nKetua {lembaga}"

    # Baris 2
    table.cell(1, 0).text = "Sutrisno, S.E."
    table.cell(1, 1).text = {
        "TPK": "Budi Santoso",
        "PPKAD": "Sri Wahyuni",
        "PPS": "Luluk Maulida",
        "Karang Taruna": "Heri Setiawan",
        "LPMD": "Sukardi"
    }.get(lembaga, "Nama Ketua")

    # Baris 3
    table.cell(2, 0).text = ""
    table.cell(2, 1).text = "Bendahara"

    # Baris 4
    table.cell(3, 0).text = ""
    table.cell(3, 1).text = {
        "TPK": "Rina Puspitasari",
        "PPKAD": "Dwi Lestari",
        "PPS": "Yuli Andriani",
        "Karang Taruna": "Riski Amalia",
        "LPMD": "Dian Sari"
    }.get(lembaga, "Nama Bendahara")

    doc.add_paragraph("\n")

    # TANDA TANGAN KETUA BPD
    bpd_paragraf = doc.add_paragraph()
    bpd_paragraf.add_run("Mengesahkan,\nKetua BPD\n\n\nMisdi")
    bpd_paragraf.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Simpan dokumen
    out_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(out_path)
    return out_path
