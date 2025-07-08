import os
import pkgutil
from docx import Document
import io

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    # Pastikan nama file template tanpa spasi/kurung agar stabil di Cloud
    template_bytes = pkgutil.get_data(__name__, "templates/template_spj.docx")
    
    if not template_bytes:
        raise FileNotFoundError("template_spj.docx tidak ditemukan")

    doc = Document(io.BytesIO(template_bytes))

    # Tambahkan konten otomatis ke dokumen
    doc.add_heading("Surat Pertanggungjawaban Kegiatan Desa", 0)

    doc.add_paragraph(f"Lembaga Pelaksana: {lembaga}")
    doc.add_paragraph(f"Nama Kegiatan    : {nama_kegiatan}")
    doc.add_paragraph(f"Tanggal Pelaksanaan: {tgl.strftime('%d-%m-%Y')}")
    doc.add_paragraph(f"Lokasi Kegiatan  : {lokasi}")
    doc.add_paragraph(f"Anggaran         : Rp {anggaran:,.0f}")
    doc.add_paragraph(f"Realisasi        : Rp {realisasi:,.0f}")
    doc.add_paragraph(f"Sumber Dana      : {sumber_dana}")

    doc.add_paragraph("\n\nDemikian Surat Pertanggungjawaban ini dibuat untuk digunakan sebagaimana mestinya.")

    out_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(out_path)
    return out_path
