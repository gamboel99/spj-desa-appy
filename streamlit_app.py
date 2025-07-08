import streamlit as st
import pandas as pd
from datetime import date
from spj_generator import buat_spj

st.set_page_config(page_title="SPJ Desa", layout="wide")
st.title("📄 Sistem Penerbitan Surat Pertanggungjawaban (SPJ) Desa")

with st.form("form_kegiatan"):
    st.subheader("📌 Data Kegiatan")
    lembaga = st.selectbox("Lembaga Pelaksana", ["TPK", "PPKAD", "PPS", "Karang Taruna", "LPMD"])
    nama_kegiatan = st.text_input("Nama Kegiatan")
    tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan", date.today())
    lokasi = st.text_input("Lokasi")
    anggaran = st.number_input("Anggaran (Rp)", 0)
    realisasi = st.number_input("Realisasi (Rp)", 0)
    sumber_dana = st.selectbox("Sumber Dana", ["DD", "ADD", "BKK", "Swadaya", "Lainnya"])
    bukti_upload = st.file_uploader("Upload Bukti Transaksi (opsional)", accept_multiple_files=True)

    submit = st.form_submit_button("✅ Buat SPJ")

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
