import pypandoc
import os
import pkgutil
from docx import Document
import io
import re

def replace_placeholder(doc, placeholder, value):
    for para in doc.paragraphs:
        if placeholder in para.text:
            inline = para.runs
            for i in range(len(inline)):
                if placeholder in inline[i].text:
                    inline[i].text = inline[i].text.replace(placeholder, value)

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    template_bytes = pkgutil.get_data(__name__, "templates/template_spj.docx")
    
    if not template_bytes:
        raise FileNotFoundError("template_spj.docx tidak ditemukan")

    doc = Document(io.BytesIO(template_bytes))

    # Replace placeholder dalam dokumen
    replace_placeholder(doc, "{{lembaga}}", lembaga)
    replace_placeholder(doc, "{{nama_kegiatan}}", nama_kegiatan)
    replace_placeholder(doc, "{{tgl}}", tgl.strftime("%d-%m-%Y"))
    replace_placeholder(doc, "{{lokasi}}", lokasi)
    replace_placeholder(doc, "{{anggaran}}", f"Rp {anggaran:,.0f}")
    replace_placeholder(doc, "{{realisasi}}", f"Rp {realisasi:,.0f}")
    replace_placeholder(doc, "{{sumber_dana}}", sumber_dana)

    # Simpan hasil
    pdf_path = doc_path.replace(".docx", ".pdf")
    try:
        pypandoc.convert_file(doc_path, 'pdf', outputfile=pdf_path)
    except Exception as e:
        raise RuntimeError(f"Gagal konversi PDF: {e}")

    return pdf_path
