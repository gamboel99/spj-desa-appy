import streamlit as st
from datetime import date
import pandas as pd
from spj_generator import buat_spj

st.title("📄 SPJ Desa | Input RAB dan Realisasi")

# Form input utama
with st.form("form_spj"):
    st.subheader("📌 Informasi Kegiatan")
    lembaga = st.selectbox("Lembaga", ["TPK", "PKK", "LPMD"])
    nama_kegiatan = st.text_input("Nama Kegiatan")
    tgl_pelaksanaan = st.date_input("Tanggal", date.today())
    lokasi = st.text_input("Lokasi")
    sumber_dana = st.selectbox("Sumber Dana", ["DD", "ADD", "Swadaya", "Lainnya"])
    kode_register = st.text_input("Kode Register SPJ", "470/SPJ-DESA")
    
    st.subheader("✍️ Pejabat")
    nama_kades = st.text_input("Nama Kepala Desa", "Sutrisno, S.E.")
    nama_ketua_bpd = st.text_input("Nama Ketua BPD", "Misdi")
    nama_ketua_lembaga = st.text_input("Nama Ketua Lembaga")
    nama_bendahara = st.text_input("Nama Bendahara")

    st.subheader("📊 Input RAB dan Realisasi")
    st.markdown("Silakan unggah file Excel berisi RAB dan Realisasi atau input manual di bawah.")

    # Input RAB & Realisasi Manual
    rab_df = st.experimental_data_editor(
        pd.DataFrame({
            "Uraian": [""],
            "Volume": [0],
            "Satuan": [""],
            "Harga Satuan": [0],
            "Realisasi": [0],
        }),
        num_rows="dynamic",
        use_container_width=True
    )

    bukti_upload = st.file_uploader("📎 Upload Bukti Transaksi (PDF/ZIP/JPG)", accept_multiple_files=False)

    submit = st.form_submit_button("✅ Buat SPJ")

if submit:
    try:
        st.write("📄 Membuat SPJ...")
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
