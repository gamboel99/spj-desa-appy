import os
from docx import Document

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    template_path = os.path.join(os.path.dirname(__file__), "templates", "template_spj.docx")
    doc = Document(template_path)
    out_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(out_path)
    return out_path
