import streamlit as st
from datetime import date
import pandas as pd
from spj_generator import buat_spj

st.set_page_config(page_title="SPJ Desa", layout="wide")
st.title("📄 SPJ Desa | Input RAB & Realisasi")

with st.form("form_spj"):
    st.subheader("📌 Informasi Kegiatan")
    lembaga = st.selectbox("Lembaga", ["TPK", "PKK", "LPMD", "Karang Taruna", "PPS"])
    nama_kegiatan = st.text_input("Nama Kegiatan")
    tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan", date.today())
    lokasi = st.text_input("Lokasi Pelaksanaan")
    sumber_dana = st.selectbox("Sumber Dana", ["DD", "ADD", "Swadaya", "Lainnya"])
    kode_register = st.text_input("Kode Register SPJ", "470/SPJ-DESA")

    st.subheader("✍️ Identitas Pejabat")
    nama_kades = st.text_input("Nama Kepala Desa", "Sutrisno, S.E.")
    nama_ketua_bpd = st.text_input("Nama Ketua BPD", "Misdi")
    nama_ketua_lembaga = st.text_input(f"Nama Ketua {lembaga}")
    nama_bendahara = st.text_input("Nama Bendahara")

    st.subheader("📊 RAB & Realisasi")
    st.markdown("Isi RAB dan Realisasi kegiatan secara rinci di tabel berikut:")

    # Tabel Input RAB + Realisasi
    rab_df_default = pd.DataFrame({
        "Uraian": ["Contoh: Beli Semen", "Contoh: Upah Tukang"],
        "Volume": [10, 5],
        "Satuan": ["Sak", "Hari"],
        "Harga Satuan": [50000, 100000],
        "Realisasi": [480000, 500000]
    })

    rab_df = st.data_editor(rab_df_default, num_rows="dynamic", use_container_width=True)

    st.subheader("📎 Upload Bukti Transaksi")
    bukti_upload = st.file_uploader("Unggah bukti transaksi (PDF/JPG/ZIP)", type=["pdf", "jpg", "jpeg", "zip"])

    submit = st.form_submit_button("✅ Buat SPJ")

if submit:
    st.info("📄 Sedang memproses dokumen SPJ...")
    try:
        file_path = buat_spj(
            lembaga, nama_kegiatan, tgl_pelaksanaan, lokasi,
            sumber_dana, rab_df, bukti_upload,
            nama_kades, nama_ketua_bpd, nama_ketua_lembaga, nama_bendahara,
            kode_register
        )
        st.success("✅ SPJ berhasil dibuat!")
        with open(file_path, "rb") as f:
            st.download_button("📥 Unduh SPJ", f.read(), file_name="SPJ_Kegiatan.docx")
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan: {e}")
