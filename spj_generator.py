import os
import pkgutil
from docx import Document
import io

def replace_placeholder(doc, placeholder, value):
    # Ganti di paragraf biasa
    for para in doc.paragraphs:
        if placeholder in para.text:
            for run in para.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)

    # Ganti juga di dalam tabel
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, value)

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, anggaran, realisasi, sumber_dana, bukti_upload):
    template_bytes = pkgutil.get_data(__name__, "templates/template_spj.docx")
    if not template_bytes:
        raise FileNotFoundError("template_spj.docx tidak ditemukan")

    doc = Document(io.BytesIO(template_bytes))

    # Data tetap
    nama_kades = "Sutrisno, S.E."
    nama_ketua_bpd = "Misdi"
    daftar_ketua = {
        "TPK": "Budi Santoso",
        "PPKAD": "Sri Wahyuni",
        "PPS": "Luluk Maulida",
        "Karang Taruna": "Heri Setiawan",
        "LPMD": "Sukardi"
    }
    daftar_bendahara = {
        "TPK": "Rina Puspitasari",
        "PPKAD": "Dwi Lestari",
        "PPS": "Yuli Andriani",
        "Karang Taruna": "Riski Amalia",
        "LPMD": "Dian Sari"
    }

    # Replace
    replace_placeholder(doc, "{{lembaga}}", lembaga)
    replace_placeholder(doc, "{{nama_kegiatan}}", nama_kegiatan)
    replace_placeholder(doc, "{{tgl}}", tgl.strftime("%d-%m-%Y"))
    replace_placeholder(doc, "{{lokasi}}", lokasi)
    replace_placeholder(doc, "{{anggaran}}", f"Rp {anggaran:,.0f}")
    replace_placeholder(doc, "{{realisasi}}", f"Rp {realisasi:,.0f}")
    replace_placeholder(doc, "{{sumber_dana}}", sumber_dana)

    replace_placeholder(doc, "{{nama_kades}}", nama_kades)
    replace_placeholder(doc, "{{nama_ketua_bpd}}", nama_ketua_bpd)
    replace_placeholder(doc, "{{nama_ketua}}", daftar_ketua.get(lembaga, "-"))
    replace_placeholder(doc, "{{nama_bendahara}}", daftar_bendahara.get(lembaga, "-"))

    out_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(out_path)
    return out_path
