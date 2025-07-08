import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload,
             nama_kades="Sutrisno, S.E.", nama_ketua_bpd="Misdi", nama_ketua_lembaga="", nama_bendahara="",
             kode_register="470/SPJ-DESA"):

    doc = Document()

    # Logo Desa
    logo_path = os.path.join(os.path.dirname(__file__), "logo_desa.png")
    if os.path.exists(logo_path):
        doc.add_picture(logo_path, width=Inches(1.2))

    # KOP DESA
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header.add_run("PEMERINTAH DESA KELING\n").bold = True
    header.add_run("KECAMATAN KEPUNG, KABUPATEN KEDIRI\n").bold = True
    header.add_run("Alamat: Jl. Raya Keling, Bukaan, Keling, Kediri, Jawa Timur 64293\n").italic = True
    doc.add_paragraph("_______________________________________________________________")

    # Nomor Surat
    nomor = f"Nomor: {kode_register}/{tgl.month:02d}/{tgl.year}"
    nomor_paragraf = doc.add_paragraph(nomor)
    nomor_paragraf.runs[0].bold = True

    # Judul
    judul = doc.add_paragraph("SURAT PERTANGGUNGJAWABAN KEGIATAN")
    judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    judul.runs[0].bold = True
    doc.add_paragraph(" ")

    # Data Kegiatan
    doc.add_paragraph(f"Nama Kegiatan        : {nama_kegiatan}")
    doc.add_paragraph(f"Tanggal Pelaksanaan  : {tgl.strftime('%d-%m-%Y')}")
    doc.add_paragraph(f"Lembaga Pelaksana    : {lembaga}")
    doc.add_paragraph(f"Lo_
