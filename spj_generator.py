import os
import pkgutil
from docx import Document
import io

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    template_bytes = pkgutil.get_data(__name__, "templates/template_spj (valid).docx")
    
    if not template_bytes:
        raise FileNotFoundError("template_spj (valid).docx tidak ditemukan")

    doc = Document(io.BytesIO(template_bytes))
    out_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(out_path)
    return out_path

if submit:
    with st.spinner("📄 Membuat dokumen SPJ..."):
        try:
            file_path = buat_spj(lembaga, nama_kegiatan, tgl_pelaksanaan, lokasi,
                                 anggaran, realisasi, sumber_dana, bukti_upload)
            st.success("✅ SPJ berhasil dibuat!")
            with open(file_path, "rb") as f:
                st.download_button("📥 Unduh SPJ (DOCX)", data=f.read(), file_name="SPJ_Kegiatan.docx")
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat membuat dokumen SPJ:\n\n{e}")
