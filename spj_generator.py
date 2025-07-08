# =========================
# File: spj_generator.py
# =========================

import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    doc = Document()

    # KOP DESA
    kop = doc.add_paragraph()
    kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
    kop.add_run("PEMERINTAH DESA KELING\nKECAMATAN KEPUNG, KABUPATEN KEDIRI\n").bold = True
    kop.add_run("Alamat: Jl. Raya Keling, Bukaan, Keling, Kediri, Jawa Timur 64293\n").italic = True
    doc.add_paragraph().add_run("_" * 60)

    doc.add_paragraph("\n\nSURAT PERTANGGUNGJAWABAN KEGIATAN", style='Title')

    # ISI KEGIATAN
    doc.add_paragraph(f"Nama Kegiatan       : {nama_kegiatan}")
    doc.add_paragraph(f"Tanggal Pelaksanaan : {tgl.strftime('%d-%m-%Y')}")
    doc.add_paragraph(f"Lembaga Pelaksana   : {lembaga}")
    doc.add_paragraph(f"Lokasi              : {lokasi}")
    doc.add_paragraph(f"Anggaran            : Rp {anggaran:,.0f}")
    doc.add_paragraph(f"Realisasi           : Rp {realisasi:,.0f}")
    doc.add_paragraph(f"Sumber Dana         : {sumber_dana}")

    # TANDA TANGAN
    doc.add_paragraph("\n\nDesa Keling, " + tgl.strftime('%d-%m-%Y'))

    table = doc.add_table(rows=4, cols=2)
    table.style = 'Table Grid'
    table.cell(0, 0).text = "Mengetahui:\nKepala Desa"
    table.cell(0, 1).text = f"Lembaga Pelaksana\nKetua {lembaga}"

    table.cell(1, 0).text = "Sutrisno, S.E."
    table.cell(1, 1).text = {
        "TPK": "Budi Santoso",
        "PPKAD": "Sri Wahyuni",
        "PPS": "Luluk Maulida",
        "Karang Taruna": "Heri Setiawan",
        "LPMD": "Sukardi"
    }.get(lembaga, "Nama Ketua")

    table.cell(2, 1).text = "Bendahara"
    table.cell(3, 1).text = {
        "TPK": "Rina Puspitasari",
        "PPKAD": "Dwi Lestari",
        "PPS": "Yuli Andriani",
        "Karang Taruna": "Riski Amalia",
        "LPMD": "Dian Sari"
    }.get(lembaga, "Nama Bendahara")

    doc.add_paragraph("\n\nMengesahkan,\nKetua BPD\n\nMisdi")

    # Simpan dokumen
    out_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(out_path)
    return out_path
