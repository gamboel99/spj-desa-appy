from docx import Document
from docx2pdf import convert
import os
from docx import Document

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    template_path = os.path.join(os.path.dirname(__file__), "templates", "template_spj.docx")
    doc = Document(template_path)

    for p in doc.paragraphs:
        if "<<LEMBAGA>>" in p.text:
            p.text = p.text.replace("<<LEMBAGA>>", lembaga)
        if "<<NAMA_KEGIATAN>>" in p.text:
            p.text = p.text.replace("<<NAMA_KEGIATAN>>", nama_kegiatan)
        if "<<TGL>>" in p.text:
            p.text = p.text.replace("<<TGL>>", str(tgl))
        if "<<LOKASI>>" in p.text:
            p.text = p.text.replace("<<LOKASI>>", lokasi)
        if "<<ANGGARAN>>" in p.text:
            p.text = p.text.replace("<<ANGGARAN>>", f"Rp {anggaran:,.0f}")
        if "<<REALISASI>>" in p.text:
            p.text = p.text.replace("<<REALISASI>>", f"Rp {realisasi:,.0f}")
        if "<<SUMBER_DANA>>" in p.text:
            p.text = p.text.replace("<<SUMBER_DANA>>", sumber_dana)

    filename_docx = f"SPJ_{lembaga}_{nama_kegiatan}.docx".replace(" ", "_")
    doc.save(filename_docx)

    convert(filename_docx)
    filename_pdf = filename_docx.replace(".docx", ".pdf")
    return filename_pdf
