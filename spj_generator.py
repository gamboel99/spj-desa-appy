import os
import pkgutil
from docx import Document
import io

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    # Baca file template dari folder templates/ sebagai binary
    template_bytes = pkgutil.get_data(__name__, "templates/template_spj (valid).docx")
    
    if not template_bytes:
        raise FileNotFoundError("template_spj (valid).docx tidak ditemukan")

    # Load dokumen dari bytes
    doc = Document(io.BytesIO(template_bytes))

    # Simpan hasilnya
    out_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(out_path)
    return out_path
